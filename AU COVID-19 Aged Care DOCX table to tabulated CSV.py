from docx import Document  # Import the Document class from the docx module to work with Word documents
import pandas as pd  # Import pandas for data manipulation and analysis

# Load the Word document
document = Document('covid-19-outbreaks-in-australian-residential-aged-care-facilities-30-august-2024.docx')

# Initialize an empty list to store tables
tables = []

# Iterate through each table in the document
for table in document.tables:
    # Create a DataFrame structure with empty strings, sized by the number of rows and columns in the table
    df = [['' for _ in range(len(table.columns))] for _ in range(len(table.rows))]
    
    # Iterate through each row in the current table
    for i, row in enumerate(table.rows):
        # Iterate through each cell in the current row
        for j, cell in enumerate(row.cells):
            # If the cell has text, store it in the corresponding DataFrame position
            if cell.text:
                df[i][j] = cell.text
    
    # Convert the list of lists (df) to a pandas DataFrame and add it to the tables list
    tables.append(pd.DataFrame(df))

# Print the list of DataFrames representing the tables
for x in range(len(tables)):
    print(tables[x])
    
# Create CSV with specific data required for COVID-19 Reporting
# Bla Bla Bla
# Bla Bla Bla
