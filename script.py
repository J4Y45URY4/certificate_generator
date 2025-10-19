import pandas as pd
import fitz  # PyMuPDF
import os

def add_names_to_certificates(pdf_template_path, excel_path, output_folder, placeholder_text):
    """
    Adds names from an Excel sheet to a PDF certificate template, centering them.

    Args:
        pdf_template_path (str): The path to the PDF certificate template.
        excel_path (str): The path to the Excel file with names.
        output_folder (str): The folder to save the new certificates.
        placeholder_text (str): The unique text string in the PDF to be replaced by the name.
    """
    try:
        df = pd.read_excel(excel_path)
    except FileNotFoundError:
        print(f"Error: Excel file not found at {excel_path}")
        return

    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    try:
        pdf_template = fitz.open(pdf_template_path)
    except FileNotFoundError:
        print(f"Error: PDF template not found at {pdf_template_path}")
        return

    page = pdf_template[0]
    rects = page.search_for(placeholder_text)

    if not rects:
        print(f"Error: The placeholder text '{placeholder_text}' was not found in the PDF.")
        return

    placeholder_rect = rects[0]
    # Calculate the center of the placeholder rectangle
    center_x = (placeholder_rect.x0 + placeholder_rect.x1) / 2
    center_y = (placeholder_rect.y0 + placeholder_rect.y1) / 2
    
    # Define font and size for the name
    font_name = "Times-italic"
    text_size = 37

    for index, row in df.iterrows():
        name = str(row['Name']).strip()
        
        # Create a copy of the template
        new_pdf = fitz.open()
        new_pdf.insert_pdf(pdf_template)
        new_page = new_pdf[0]
        
        # Calculate the width of the text to be inserted
        text_width = fitz.get_text_length(name, fontname=font_name, fontsize=text_size)
        
        # Calculate the new x-coordinate to center the text
        start_x = center_x - (text_width / 2)
        
        # Insert the name at the calculated center point
        new_page.insert_text((start_x, center_y), name, fontsize=text_size, fontname=font_name)
        
        

        safe_name = name.replace(" ", "_").replace(".", "").replace("/", "").replace("\\", "")
        output_path = os.path.join(output_folder, f"certificate_{safe_name}.pdf")
        
        new_pdf.save(output_path)
        new_pdf.close()
        print(f"Created certificate for {name} at {output_path}")
# --- Example Usage ---
if __name__ == '__main__':
    # Define file paths and settings
    pdf_template = 'template.pdf'
    excel_file = 'names.xlsx'
    output_dir = 'certificates_output'
    placeholder = 'NAME_PLACEHOLDER'  # Make sure this text exists in your PDF template

    add_names_to_certificates(pdf_template, excel_file, output_dir, placeholder)