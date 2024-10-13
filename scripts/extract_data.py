from docx import Document
import pandas as pd
import os

print("Current working directory:", os.getcwd())

def extract_data_from_word(file_path):
    """
    Extract data from the Word document and return it as a list of dictionaries.
    """
    data = []
    document = Document(file_path)
    
    # Loop through each table in the document to extract data
    for table in document.tables:
        headers = [cell.text.strip() for cell in table.rows[0].cells]
        for row in table.rows[1:]:
            row_data = {headers[i]: cell.text.strip() for i, cell in enumerate(row.cells)}
            data.append(row_data)
    
    return data

def save_to_excel(data, output_file_path):
    """
    Save the extracted data to an Excel file using pandas.
    """
    # Ensure the output directory exists
    output_dir = os.path.dirname(output_file_path)
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    df = pd.DataFrame(data)
    df.to_excel(output_file_path, index=False)
    print(f"Data successfully saved to {output_file_path}")

if __name__ == "__main__":
    # Define file paths
    input_file = '/Users/mutai/Desktop/data extraction/data/input/sample.docx'
    output_file = '/Users/mutai/Desktop/data extraction/data/output/structured_data.xlsx'
    
    # Extract data from the Word document
    extracted_data = extract_data_from_word(input_file)
    
    # Save the structured data to an Excel file
    save_to_excel(extracted_data, output_file)
