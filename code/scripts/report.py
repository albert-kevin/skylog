# load modules
import pandas as pd
from datetime import datetime
import zipfile
import os

# load file
csv_file = '../../data/input.csv'
df = pd.read_csv(csv_file, sep=';', encoding='utf-8')
df = df.convert_dtypes()
# make the names snake_case (remove spaces)
df['Persoon'] = df['Persoon'].apply(lambda x: x.replace(' ', '_'))

# Create backup copy
backup_date = datetime.now().strftime("%d%b%Y")
backup_file = "../../data/bronze/input_backup_" + backup_date + ".csv"
df.to_csv(backup_file, index=False, sep=';', encoding='utf-8')

# Create a zip file to store the Excel files and the backup csv
path = f"../../data/silver/"
zip_filename = f"{path}{datetime.now().strftime('%d%b%Y')}_Vluchtlogboeken.zip"
grouped = df.groupby("Persoon")
with zipfile.ZipFile(zip_filename, 'w') as zip_file:

    # Loop over each person and create a separate Excel file for their data
    for name, group in grouped:
        xslx_filename = f"{path}{name}__{datetime.now().strftime('%d%b%Y')}__Vluchtlogboek.xlsx"
        # Remove the "Persoon" column from the DataFrame
        group = group.drop(columns=["Persoon"])
        # Write the person's data to a sheet in the Excel file
        with pd.ExcelWriter(xslx_filename, engine='xlsxwriter') as writer:
            group.to_excel(writer, index=False, na_rep='NaN', sheet_name="Vluchtlogboek")
            # Auto-adjust columns' width
            worksheet = writer.sheets["Vluchtlogboek"]
            for i, col in enumerate(group.columns):
                column_width = max(group[col].astype(str).map(len).max(), len(col)) + 2
                worksheet.set_column(i, i, column_width)
        # Add the Excel file to the zip file
        zip_file.write(xslx_filename, arcname=os.path.basename(xslx_filename))

    # Add the backup CSV to the zip file
    zip_file.write(backup_file, arcname=os.path.basename(backup_file))