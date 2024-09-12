import csv
from docx import Document

def extract_tables_from_docx(docx_path, output_dir):
    # Load the .docx file
    doc = Document(docx_path)
    
    # Iterate over the tables in the document
    for i, table in enumerate(doc.tables):
        # Create a separate CSV file for each table
        csv_path = f"{output_dir}/table_{i + 1}.csv"
        
        with open(csv_path, mode='w', newline='', encoding='utf-8') as csv_file:
            writer = csv.writer(csv_file)

            # Write each row from the table to the CSV file
            for row in table.rows:
                row_data = [cell.text.strip() for cell in row.cells]
                writer.writerow(row_data)
        
        print(f"Table {i + 1} saved to {csv_path}")

if __name__ == "__main__":
    docx_file = "covid-19-outbreaks-in-australian-residential-aged-care-facilities-30-august-2024.docx"  # Replace with the path to your .docx file
    output_dir = "."  # Replace with the path to your output directory
    extract_tables_from_docx(docx_file, output_dir)

