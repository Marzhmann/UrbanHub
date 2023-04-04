import pandas as pd
import openpyxl



output_df = pd.DataFrame()
labels_list = ["ID",
                   "Typology",
                   "Green space ratio",
                   "X",
                   "Y",
                   "Rotation",
                   "Main street",
                   "Sub street",
                   "Bldg Footprint",
                   "Density",
                   "Type (Bldg:1,Park:0)",
                   "Bldg Centroids x",
                   "Bldg Centroids y",
                   "Lengths",
                   "Widths",
                   "Stories",
                   "Visibility",
                   "Cooling - Cold",
                   "Heating - Cold",
                   "Lighting - Cold",
                   "Hot water - Cold",
                   "Gas - Cold",
                   "Cooling - Hot",
                   "Heating - Hot",
                   "Lighting - Hot",
                   "Hot water - Hot",
                   "Gas - Hot",
                   "Compactness 1",
                   "Shape Factor",
                   "Aspect Ratio",
                   "Annual Solar Hours",
                   "Roof radiation- Cold",
                   "Roof radiation- Hot",
                   "Walk-score",
                   "SVF",
                   "Ave. UTCI - Cold",
                   "Ave. UTCI - Hot",
                   "Ave. Percenet of Shaded area",
                   "Total EUI - Cold",
                   "Total EUI - Hot"]

labels_list.reverse()

for labels in labels_list:
    output_df.insert(loc=0, column=labels, value=1)

# workbook = openpyxl.load_workbook(r"E:\Projects\UrbanHub\files\Output 02 - no energy fix.xlsx")
# for sheet in workbook.worksheets:
for sheet in range(1):
    input_df = pd.read_excel(r"E:\Projects\UrbanHub\files\Output 02 - no energy fix.xlsx", sheet_name=sheet)

    input_df = input_df.T
    input_df.columns = input_df.iloc[0]
    input_df = input_df.iloc[1:]
    input_df = input_df.loc[input_df.index[0]:].reset_index(drop=True)

    sheet_id = input_df.iloc[:,5]
    sheet_cent = input_df.iloc[:,6]
    sheet_length = input_df.iloc[:,7]
    sheet_width = input_df.iloc[:,8]
    sheet_story = input_df.iloc[:,9]
    sheet_visibility = input_df.iloc[:,11]
    sheet_cooling_c = input_df.iloc[:,12]
    sheet_heating_c = input_df.iloc[:,13]
    sheet_lighting_c = input_df.iloc[:,14]
    sheet_hotwater_c = input_df.iloc[:,15]
    sheet_gas_c = input_df.iloc[:,16]
    sheet_cooling_h = input_df.iloc[:,17]
    sheet_heating_h = input_df.iloc[:,18]
    sheet_lighting_h = input_df.iloc[:,19]
    sheet_hotwater_h = input_df.iloc[:,20]
    sheet_gas_h = input_df.iloc[:,21]
    sheet_comp = input_df.iloc[:,22]
    sheet_shapefactor = input_df.iloc[:,23]
    sheet_aspect = input_df.iloc[:,24]
    sheet_annula_h = input_df.iloc[:,25]
    sheet_roofrad_c = input_df.iloc[:,26]
    sheet_roofrad_h = input_df.iloc[:,27]
    sheet_walkscore = input_df.iloc[:,28]
    sheet_svf = input_df.iloc[:,29]

    output_df_dup = output_df.copy()

    output_df_dup.loc[:,'ID'] = sheet_id
    output_df_dup.loc[:,'Bldg Centroids x'] = sheet_cent
    output_df_dup.loc[:,'Bldg Centroids y'] = sheet_cent
    output_df_dup.loc[:,'Lengths'] = sheet_length
    output_df_dup.loc[:,'Widths'] = sheet_width
    output_df_dup.loc[:,'Stories'] = sheet_story
    output_df_dup.loc[:,'Visibility'] = sheet_visibility
    output_df_dup.loc[:,'Cooling - Cold'] = sheet_cooling_c
    output_df_dup.loc[:,'Heating - Cold'] = sheet_heating_c
    output_df_dup.loc[:,'Lighting - Cold'] = sheet_lighting_c
    output_df_dup.loc[:,'Hot water - Cold'] = sheet_hotwater_c
    output_df_dup.loc[:,'Gas - Cold'] = sheet_gas_c
    output_df_dup.loc[:,'Cooling - Hot'] = sheet_cooling_h
    output_df_dup.loc[:,'Heating - Hot'] = sheet_heating_h
    output_df_dup.loc[:,'Lighting - Hot'] = sheet_lighting_h
    output_df_dup.loc[:,'Hot water - Hot'] = sheet_hotwater_h
    output_df_dup.loc[:,'Gas - Hot'] = sheet_gas_h
    output_df_dup.loc[:,'Compactness 1'] = sheet_comp
    output_df_dup.loc[:,'Shape Factor'] = sheet_shapefactor
    output_df_dup.loc[:,'Aspect Ratio'] = sheet_aspect
    output_df_dup.loc[:,'Annual Solar Hours'] = sheet_annula_h
    output_df_dup.loc[:,'Roof radiation- Cold'] = sheet_roofrad_c
    output_df_dup.loc[:,'Roof radiation- Hot'] = sheet_roofrad_h
    output_df_dup.loc[:,'Walk-score'] = sheet_walkscore
    output_df_dup.loc[:,'SVF'] = sheet_svf

    output_df = pd.concat([output_df, output_df_dup])

output_df = output_df.reset_index(drop=True)

print(output_df)

writer = pd.ExcelWriter('output.xlsx')
output_df.to_excel(writer)
writer.save()