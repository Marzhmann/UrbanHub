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

workbook = openpyxl.load_workbook(r"E:\Projects\UrbanHub\files\Output 02 - no energy fix.xlsx")
for sheet in workbook.worksheets:
# for sheet in range(2):
#     input_df = pd.read_excel(r"E:\Projects\UrbanHub\files\Output 02 - no energy fix.xlsx", sheet_name=sheet)

    input_df = input_df.T
    input_df.columns = input_df.iloc[0]
    input_df = input_df.iloc[1:]
    input_df = input_df.loc[input_df.index[0]:].reset_index(drop=True)

    sheet_svf = input_df.iloc[:,29]
    sheet_ID = input_df.iloc[:,5]

    output_df_dup = output_df.copy()
    output_df_dup.loc[:,'SVF'] = sheet_svf
    output_df_dup.loc[:,'ID'] = sheet_ID
    output_df = pd.concat([output_df, output_df_dup])

output_df = output_df.reset_index(drop=True)

# print(output_df)

writer = pd.ExcelWriter('output.xlsx')
output_df.to_excel(writer)
writer.save()