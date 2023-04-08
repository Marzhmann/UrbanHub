import pandas as pd

# creating a dataframe to fill with desired structure
output_df = pd.DataFrame()

# column labels for output dataframe
# labels_list = ["ID",
#                "Typology",
#                "Green space ratio",
#                "X",
#                "Y",
#                "Rotation",
#                "Main street",
#                "Sub street",
#                "Bldg Footprint",
#                "Density",
#                "Type (Bldg:1,Park:0)",
#                "Bldg Centroids x",
#                "Bldg Centroids y",
#                "Lengths",
#                "Widths",
#                "Stories",
#                "Visibility",
#                "Cooling - Cold",
#                "Heating - Cold",
#                "Lighting - Cold",
#                "Hot water - Cold",
#                "Gas - Cold",
#                "Cooling - Hot",
#                "Heating - Hot",
#                "Lighting - Hot",
#                "Hot water - Hot",
#                "Gas - Hot",
#                "Compactness 1",
#                "Shape Factor",
#                "Aspect Ratio",
#                "Annual Solar Hours",
#                "Roof radiation- Cold",
#                "Roof radiation- Hot",
#                "Walk-score",
#                "SVF",
#                "Ave. UTCI - Cold",
#                "Ave. UTCI - Hot",
#                "Ave. Percent of Shaded area",
#                "Total EUI - Cold",
#                "Total EUI - Hot"]

# each label is added to the left side of columns and shift the columns to the right
# so labels_list has to be reversed
# labels_list.reverse()

# inserting labels_list items as columns names
# for labels in labels_list:
#     output_df.insert(loc=0, column=labels, value=1)

# opening each sheet in the input Excel file
# workbook = openpyxl.load_workbook(r"E:\Projects\UrbanHub\files\Output 02 - no energy fix.xlsx")
# for sheet in workbook.worksheets:
for sheet in range(0, 1):
    input_df = pd.read_excel(r"E:\Projects\UrbanHub\files\Output 04 - no energy fix.xlsx", sheet_name=sheet)
    # Transposing dataframe and set index
    input_df = input_df.T
    input_df = input_df.reset_index(drop=True)

    # returning number of rows in the input dataframe
    file_length = len(input_df)

    # reading data from input dataframe (existing Excel file) and assign them to variables
    sheet_typology = pd.DataFrame({"Typology": [input_df.iloc[0, 0]] * file_length})
    sheet_green_space = pd.DataFrame({"Green space ratio": [input_df.iloc[1, 0]] * file_length})
    sheet_x = pd.DataFrame({"X": [input_df.iloc[2, 0]] * file_length})
    sheet_y = pd.DataFrame({"Y": [input_df.iloc[3, 0]] * file_length})
    sheet_rotation = pd.DataFrame({"Rotation": [input_df.iloc[4, 0]] * file_length})
    sheet_main_street = pd.DataFrame({"Main street": [input_df.iloc[5, 0]] * file_length})
    sheet_sub_street = pd.DataFrame({"Sub street": [input_df.iloc[6, 0]] * file_length})
    sheet_bldg_fprint = pd.DataFrame({"Bldg Footprint": [input_df.iloc[0, 2]] * file_length})
    sheet_density = pd.DataFrame({"Density": [input_df.iloc[4, 2]] * file_length})
    sheet_utci_c = pd.DataFrame({"Ave. UTCI- Cold": [input_df.iloc[1, 4]] * file_length})
    sheet_utci_h = pd.DataFrame({"Ave. UTCI- Hot": [input_df.iloc[2, 4]] * file_length})
    sheet_shaded = pd.DataFrame({"Ave. Percent of Shaded area": [input_df.iloc[3, 4]] * file_length})
    sheet_eui_c = pd.DataFrame({"Total EUI - Cold": [input_df.iloc[4, 4]] * file_length})
    sheet_eui_h = pd.DataFrame({"Total EUI - Hot": [input_df.iloc[5, 4]] * file_length})
    sheet_cent = str(input_df.iloc[1:, 6])
    sheet_id = input_df.iloc[1:, 5]
    sheet_type = sheet_id.apply(lambda x: 0 if x.startswith("P") else 1)
    sheet_length = input_df.iloc[1:, 7]
    sheet_width = input_df.iloc[1:, 8]
    sheet_story = input_df.iloc[1:, 9]
    sheet_visibility = input_df.iloc[1:, 11]
    sheet_cooling_c = input_df.iloc[1:, 12]
    sheet_heating_c = input_df.iloc[1:, 13]
    sheet_lighting_c = input_df.iloc[1:, 14]
    sheet_hotwater_c = input_df.iloc[1:, 15]
    sheet_gas_c = input_df.iloc[1:, 16]
    sheet_cooling_h = input_df.iloc[1:, 17]
    sheet_heating_h = input_df.iloc[1:, 18]
    sheet_lighting_h = input_df.iloc[1:, 19]
    sheet_hotwater_h = input_df.iloc[1:, 20]
    sheet_gas_h = input_df.iloc[1:, 21]
    sheet_comp = input_df.iloc[1:, 22]
    sheet_shape_factor = input_df.iloc[1:, 23]
    sheet_aspect = input_df.iloc[1:, 24]
    sheet_annual_h = input_df.iloc[1:, 25]
    sheet_roof_rad_c = input_df.iloc[1:, 26]
    sheet_roof_rad_h = input_df.iloc[1:, 27]
    sheet_walkscore = input_df.iloc[1:, 28]
    sheet_svf = input_df.iloc[1:, 29]

    # making a copy of output dataframe and write new structure to it
    # this dataframe will be overwritten on each iteration an added to the output dataframe
    output_df_dup = output_df.copy()

    # writing variables to the copy of output dataframe
    output_df["ID"] = sheet_id
    output_df["Typology"] = sheet_typology
    output_df["Green space ratio"] = sheet_green_space
    output_df["X"] = sheet_x
    output_df["Y"] = sheet_y
    output_df["Rotation"] = sheet_rotation
    output_df["Main street"] = sheet_main_street
    output_df["Sub street"] = sheet_sub_street
    output_df["Bldg Footprint"] = sheet_bldg_fprint
    output_df["Density"] = sheet_density
    output_df["Type (Bldg:1,Park:0)"] = sheet_type
    # output_df["Bldg Centroids x"] = sheet_cent
    # output_df["Bldg Centroids y"] = sheet_cent
    output_df["Lengths"] = sheet_length
    output_df["Widths"] = sheet_width
    output_df["Stories"] = sheet_story
    output_df["Visibility"] = sheet_visibility
    output_df["Cooling - Cold"] = sheet_cooling_c
    output_df["Heating - Cold"] = sheet_heating_c
    output_df["Lighting - Cold"] = sheet_lighting_c
    output_df["Hot water - Cold"] = sheet_hotwater_c
    output_df["Gas - Cold"] = sheet_gas_c
    output_df["Cooling - Hot"] = sheet_cooling_h
    output_df["Heating - Hot"] = sheet_heating_h
    output_df["Lighting - Hot"] = sheet_lighting_h
    output_df["Hot water - Hot"] = sheet_hotwater_h
    output_df["Gas - Hot"] = sheet_gas_h
    output_df["Compactness 1"] = sheet_comp
    output_df["Shape Factor"] = sheet_shape_factor
    output_df["Aspect Ratio"] = sheet_aspect
    output_df["Annual Solar Hours"] = sheet_annual_h
    output_df["Roof radiation - Cold"] = sheet_roof_rad_c
    output_df["Roof radiation - Hot"] = sheet_roof_rad_h
    output_df["Walk-score"] = sheet_walkscore
    output_df["SVF"] = sheet_svf
    output_df["Ave. UTCI - Cold"] = sheet_utci_c
    output_df["Ave. UTCI - Hot"] = sheet_utci_h
    output_df["Ave. Percent of Shaded area"] = sheet_shaded
    output_df["Total EUI - Cold"] = sheet_eui_c
    output_df["Total EUI - Hot"] = sheet_eui_h

    # appending the copy of output dataframe to final output dataframe
    output_df = pd.concat([output_df, output_df_dup])
    # removing duplicated rows
    output_df = output_df.drop_duplicates()

# reset index numbers
output_df = output_df.reset_index(drop=True)

# converting output dataframe to an excel file
writer = pd.ExcelWriter('output.xlsx')
output_df.to_excel(writer)
writer.save()
