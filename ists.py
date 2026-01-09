import os
import json
import pdfplumber

def scrape_ists_loss():
    # Define paths
    # Base dir is where the script is located: .../ADANI
    base_dir = os.path.dirname(os.path.abspath(__file__))
    
    input_dir = os.path.join(base_dir, "ists_pdf")
    output_dir = os.path.join(base_dir, "ists_extracted")
    
    # Create output directory if it doesn't exist
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
        
    output_file = os.path.join(output_dir, "ists_loss.json")
    
    extracted_data = {}

    # Iterate over files in input directory
    if not os.path.exists(input_dir):
        print(f"Input directory not found: {input_dir}")
        return

    for filename in os.listdir(input_dir):
        if filename.lower().endswith(".pdf"):
            pdf_path = os.path.join(input_dir, filename)
            try:
                print(f"Processing {filename}...")
                with pdfplumber.open(pdf_path) as pdf:
                    # Scrape only the 1st page
                    if len(pdf.pages) > 0:
                        page = pdf.pages[0]
                        # Extract tables
                        tables = page.extract_tables()
                        
                        for table in tables:
                            if not table:
                                continue
                            
                            # Treat each row as a key-value pair
                            for row in table:
                                # Clean cells
                                clean_row = [cell.strip() if cell else "" for cell in row]
                                
                                # Ensure reasonable length for key-value
                                if len(clean_row) >= 2:
                                    key = clean_row[0]
                                    # Join remaining columns as value if multiple, or just take second
                                    # Assuming standard key-value table: Col 0 is Key, Col 1 is Value
                                    value = clean_row[1]
                                    
                                    if key:
                                        extracted_data[key] = value

            except Exception as e:
                print(f"Error processing {filename}: {e}")

    # Save to JSON file (Key-Value pairs)
    with open(output_file, 'w', encoding='utf-8') as f:
        json.dump(extracted_data, f, indent=4)
        
    print(f"Extraction complete. Saved to {output_file}")

if __name__ == "__main__":
    scrape_ists_loss()
