import openpyxl

# Load an existing Excel workbook
workbook = openpyxl.load_workbook('example.xlsx')

# Select a worksheet (replace 'Sheet1' with the name of your sheet)
sheet = workbook['Sheet1']

# Merge cells (specify the range of cells you want to merge)
sheet.merge_cells('A1:B2')  # Merging cells A1 to B2

# Save the changes to a new file
workbook.save('example_merged.xlsx')
