import pandas as pd

# Load the data from the Excel file
df = pd.read_excel('add-file-name.xlsx')

# Create the SKU column
df['SKU'] = df['article no.'] + '-' + df['color name'] + '-' + df['size run'].astype(str)
df['function'] = df['function'].fillna('')
df['heel height'] = df['heel height'].fillna('')
df['sole tech'] = df['sole tech'].fillna('')
df['Description'] = "Ιδιότητα: " + df['function'] + df['sole tech'] +  "\n" + "Ύψος Τακουνιού: " + df['heel height']

# Map the old columns to the new ones
column_mapping = {
    'CAN GTIN': 'barcode',
    'Parent CODE': 'Parent CODE',
    'SKU': 'SKU',
    'size run': 'Size',
    'pairs': 'Quantity',
    'color name': 'Color',
    'UVP Griechenland (EUR)': 'Price',
    'title': 'title',
    'Description': 'Description',
    'Brand': 'Brand',
    'type': 'type',
    'article no.': 'Code'
}

# Add a placeholder for 'title' and 'Description'
df['title'] = '-'
# df['Description'] = '-'
df['Brand'] = 'Tamaris'
df["type"] = 'variable'
df['Parent CODE'] = df['article no.']

# Select only the columns we need and rename them
df = df[list(column_mapping.keys())]
df.rename(columns=column_mapping, inplace=True)

# printing all columns of the dataframe 
# print(df.columns.tolist()) 

df = df.sort_values('SKU')

# Write the transformed data to a new Excel file
with pd.ExcelWriter('output_file.xlsx') as writer:
    df.to_excel(writer, index=False)
    for _, row in df.iterrows():
        row.to_frame().T.to_excel(writer, index=False)
        writer.sheets['Sheet1'].append([''] * len(df.columns))
        writer.sheets['Sheet1'].append([''] * len(df.columns))

print("Data transformation complete. The output has been written to 'output_file.xlsx'.")

