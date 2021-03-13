import pandas as pd
import xlsxwriter 


# Function to convert string to integer...
def conv_int(string):
	return int(string.replace(".", ""))


# Reading data...
bom_data = pd.read_excel("./bom_data.xlsx")

bom_dic = {}
material_dic = {}


# Making all the datas as dictionary...
for i in range(len(bom_data)):
	level = conv_int(bom_data["Level"][i])
	material = ""
	quantity = ""
	unit = ""
	raw_material = ""
	material_quan = ""
	material_unit = ""
	
	if(level > 1):
		raw_material = bom_data["Raw material"][i]
		quantity = bom_data["Quantity"][i]
		unit = bom_data["Unit"][i]
		
		for j in range(i, -1, -1):
			if(conv_int(bom_data["Level"][j]) == level-1 ):
				material = bom_data["Raw material"][j]
				material_quan = bom_data["Quantity"][j]
				material_unit = bom_data["Unit"][j]
				break
	
	else:
		material = bom_data["Item Name"][i]
		raw_material = bom_data["Raw material"][i]
		quantity = bom_data["Quantity"][i]
		unit = bom_data["Unit"][i]
		material_quan = "1"
		material_unit = bom_data["Unit"][i]
		
	if material not in bom_dic:
			bom_dic[material] = []
			
	bom_dic[material].append({"Raw material": raw_material, "Quantity": quantity, "Unit": unit})
	material_dic[material] = {"Quantity": material_quan, "Unit": material_unit}


# Writing all the datas into new excel sheet...
workbook = xlsxwriter.Workbook("BOM.xlsx")
format_default = workbook.add_format()
format_default.set_border(style=1)
format_material = workbook.add_format({"bg_color": "yellow"})
format_header = workbook.add_format({"bg_color":"blue", "bold": True })
format_material.set_border(style=1)
format_header.set_border(style=1)

for key in bom_dic:
	worksheet = workbook.add_worksheet(key)
	worksheet.write(0, 0, "Finished Good List", format_default)
	worksheet.write(1, 0, "#", format_header)
	worksheet.write(1, 1, "Item Description", format_header)
	worksheet.write(1, 2, "Quantity", format_header)
	worksheet.write(1, 3, "Unit", format_header)
	worksheet.write(2, 0, "1", format_default)
	worksheet.write(2, 1, key, format_material)
	worksheet.write(2, 2, material_dic[key]["Quantity"], format_default)
	worksheet.write(2, 3, material_dic[key]["Unit"], format_default)
	worksheet.write(3, 0 ,"End of FG", format_default)
	worksheet.write(4, 0 ,"Raw Material List", format_default)
	worksheet.write(5, 0, "#", format_header)
	worksheet.write(5, 1, "Item Description", format_header)
	worksheet.write(5, 2, "Quantity", format_header)
	worksheet.write(5, 3, "Unit", format_header)
	row = 6
	index = 1
	
	for data in bom_dic[key]:
		worksheet.write(row, 0, index, format_default)
		worksheet.write(row, 1, data["Raw material"], format_material)
		worksheet.write(row, 2, data["Quantity"], format_material)
		worksheet.write(row, 3, data["Unit"], format_material)
		row+=1
		index+=1
	
	worksheet.write(row, 0, "End of RM")
	
	
workbook.close()

