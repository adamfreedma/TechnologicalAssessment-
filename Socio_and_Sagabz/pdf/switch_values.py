import os
from docx import Document
import re
import glob

def extract_table_numbers(doc):
    """Extract all numbers from each table, preserving table structure."""
    tables_data = []
    for table in doc.tables:
        table_numbers = []
        for row in table.rows:
            row_numbers = []
            for cell in row.cells:
                cell_numbers = []
                for para in cell.paragraphs:
                    for run in para.runs:
                        matches = re.findall(r'\b\d+(?:\.\d+)?\b', run.text)
                        run_color = run.font.color.rgb  # Save the font color
                        cell_numbers.extend(matches)
                        for i, match in enumerate(matches):
                            if i < len(cell_numbers):
                                cell_numbers[i] = (cell_numbers[i], run_color)  # Pair number with its color
                        cell_numbers.extend(matches)
                row_numbers.append(cell_numbers)
            table_numbers.append(row_numbers)
        tables_data.append(table_numbers)
    return tables_data

def replace_table_numbers(doc, source_tables_data):
    """Replace numbers in doc's tables using the corresponding values from source (doc1), preserving images."""
    for table_index, table in enumerate(doc.tables):
        if table_index >= len(source_tables_data):
            break
        source_table = source_tables_data[table_index]
        for row_index, row in enumerate(table.rows):
            if row_index >= len(source_table):
                continue
            for cell_index, cell in enumerate(row.cells):
                if cell_index >= len(source_table[row_index]):
                    continue
                replacement_numbers = iter(source_table[row_index][cell_index])
                for para in cell.paragraphs:
                    for run in para.runs:
                        original_text = run.text
                        if not original_text:
                            continue
                        def replacer(match):
                            try:
                                val = next(replacement_numbers)[0]
                                if isinstance(val, tuple):
                                    val = val[0]
                                return val
                            except StopIteration:
                                return match.group(0)
                        new_text = re.sub(r'\b\d+(?:\.\d+)?\b', replacer, original_text)
                        if new_text != original_text:
                            run.text = new_text  # updates text only, preserving the run and any non-text XML (like images)
                            run.font.color.rgb = source_table[row_index][cell_index][0][1]  # Retain the original color
    return doc

def main(doc1_path, doc2_path, output_path):
    
    if not os.path.exists(doc2_path):
        print(f"❌ Error: no match for {doc1_path} in {doc2_path}.")
        return
    
    doc1 = Document(doc1_path)
    doc2 = Document(doc2_path)

    source_tables_data = extract_table_numbers(doc1)
    replace_table_numbers(doc2, source_tables_data)

    doc2.save(output_path)
    print(f"✅ Done. Updated tables saved to '{output_path}' without removing images.")



if __name__ == "__main__":
    for file in glob.glob("input/*.docx"):
        print(f"Processing {file}...")
        input_file = file
        overwrite_file = os.path.join("word_input", file.split("\\")[-1])
        output_file = os.path.join("output_word", file.split("\\")[-1])
        
        main(input_file, overwrite_file, output_file)