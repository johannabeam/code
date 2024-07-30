import pandas as pd

from docx import Document

# Function to format and combine column values
def format_combined_cell(row):
    # Combine the data into a single formatted string
    return (f"{row['Species']} ({row['Sample Type']})\n"
            f"{row['Band #']}\n"
            f"Ext: {row['Month']}/{row['Day']}/{row['Year']}\n"
            f"{row['Near Town']}, {row['State']}, {row['Country']}")

# Load Excel file
excel_file = 'data.xlsx'
df = pd.read_excel(excel_file)


# Load the existing Word document
doc_path = 'existing_document.docx'
doc = Document(doc_path)

# Assuming the table is the first table in the document
table = doc.tables[0]

# Clear existing rows (excluding header)
for row in table.rows[1:]:
    tbl = table._tbl
    tbl.remove(row._tr)

# Add combined data to the table
for index, row in df.iterrows():
    # Add a new row to the table
    row_cells = table.add_row().cells
    combined_text = format_combined_cell(row)
    row_cells[0].text = combined_text
