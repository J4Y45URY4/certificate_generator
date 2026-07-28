import pandas as pd
import fitz  # PyMuPDF
import os
import uuid
import qrcode
import io
import dotenv
from supabase import create_client

def add_names_to_certificates(
    pdf_template_path, 
    excel_path, 
    output_folder, 
    placeholder_text, 
    font_name="Times-italic", 
    text_size=37, 
    include_qr=True,
    qr_placeholder="QR_PLACEHOLDER",
    id_placeholder="ID_PLACEHOLDER",
    event_name="General Event",
    issue_date="2026-07-28",
    supabase_url=None,
    supabase_key=None,
    verification_base_url=None,
    progress_callback=None
):
    """
    Adds names from an Excel sheet to a PDF certificate template, centering them.
    Generates a unique ID and QR code, inserts them into the certificate,
    saves the IDs to the Excel sheet, and uploads metadata to Supabase DB.
    """
    # Load dotenv file
    dotenv.load_dotenv()
    
    # Resolve Supabase variables
    db_url = supabase_url or os.getenv("SUPABASE_URL")
    db_key = supabase_key or os.getenv("SUPABASE_KEY")
    base_url = verification_base_url or os.getenv("VERIFICATION_BASE_URL")
    
    supabase_client = None
    if db_url and db_key and str(db_url).strip() and str(db_key).strip():
        if not str(db_url).startswith("https://your-"):
            try:
                supabase_client = create_client(db_url, db_key)
            except Exception as e:
                print(f"Warning: Failed to initialize Supabase client: {e}")
    
    try:
        # Support both CSV and Excel files
        if str(excel_path).endswith('.csv'):
            df = pd.read_csv(excel_path)
        else:
            df = pd.read_excel(excel_path)
    except Exception as e:
        raise ValueError(f"Error reading names file: {e}")

    if df.empty:
        raise ValueError("The uploaded names file is empty.")

    # Find the name column case-insensitively
    name_col = None
    for col in df.columns:
        if str(col).strip().lower() == 'name':
            name_col = col
            break
    if name_col is None:
        name_col = df.columns[0]  # Fallback to first column

    # Find or create Certificate ID column case-insensitively
    id_col = None
    for col in df.columns:
        if str(col).strip().lower() in ['certificate id', 'certificate_id', 'id']:
            id_col = col
            break
            
    if id_col is None:
        id_col = 'Certificate ID'
        df[id_col] = None

    # Fill missing certificate IDs
    for idx, row in df.iterrows():
        val = str(row[id_col]).strip()
        if not val or val == 'nan' or val == 'None' or pd.isna(row[id_col]):
            unique_id = f"CERT-{uuid.uuid4().hex[:8].upper()}"
            df.at[idx, id_col] = unique_id

    # Save the updated dataframe back to the source names file
    try:
        if str(excel_path).endswith('.csv'):
            df.to_csv(excel_path, index=False)
        else:
            df.to_excel(excel_path, index=False)
    except Exception as e:
        print(f"Warning: Could not write back to names file: {e}")

    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    try:
        pdf_template = fitz.open(pdf_template_path)
    except Exception as e:
        raise ValueError(f"Error opening PDF template: {e}")

    try:
        page = pdf_template[0]
        rects = page.search_for(placeholder_text)

        if not rects:
            raise ValueError(f"The name placeholder text '{placeholder_text}' was not found in the PDF.")

        placeholder_rect = rects[0]
        center_x = (placeholder_rect.x0 + placeholder_rect.x1) / 2
        center_y = (placeholder_rect.y0 + placeholder_rect.y1) / 2
        
        total_names = len(df)
        generated_files = []

        # Helper to generate QR code bytes
        def make_qr_bytes(data):
            qr = qrcode.QRCode(version=1, box_size=10, border=1)
            qr.add_data(data)
            qr.make(fit=True)
            img = qr.make_image(fill_color="black", back_color="white")
            buf = io.BytesIO()
            img.save(buf, format="PNG")
            return buf.getvalue()

        for index, row in df.iterrows():
            name = str(row[name_col]).strip()
            cert_id = str(row[id_col]).strip()
            
            if not name or name.lower() == 'nan':
                continue
            
            # 1. Upload to Supabase if credentials are setup
            if supabase_client:
                try:
                    payload = {
                        "id": cert_id,
                        "name": name,
                        "event_name": event_name,
                        "issue_date": issue_date
                    }
                    supabase_client.table("certificates").upsert(payload).execute()
                except Exception as e:
                    raise ValueError(f"Failed to synchronize certificate for '{name}' with Supabase: {e}")
            
            # Create a copy of the template
            new_pdf = fitz.open()
            new_pdf.insert_pdf(pdf_template)
            new_page = new_pdf[0]
            
            page_w = new_page.rect.width
            page_h = new_page.rect.height
            
            # 2. Insert centered name
            text_width = fitz.get_text_length(name, fontname=font_name, fontsize=text_size)
            start_x = center_x - (text_width / 2)
            new_page.insert_text((start_x, center_y), name, fontsize=text_size, fontname=font_name)
            
            # 3. Insert QR Code and Certificate ID
            if include_qr:
                # Format URL
                if base_url:
                    sep = "&" if "?" in base_url else "?"
                    qr_data = f"{base_url}{sep}id={cert_id}"
                else:
                    qr_data = cert_id
                
                qr_bytes = make_qr_bytes(qr_data)
                
                # Check for QR Placeholder in the PDF
                qr_rects = new_page.search_for(qr_placeholder)
                if qr_rects:
                    qr_rect = qr_rects[0]
                    new_page.draw_rect(qr_rect, color=(1, 1, 1), fill=(1, 1, 1), width=0)
                    new_page.insert_image(qr_rect, stream=qr_bytes)
                else:
                    # Fallback to bottom-left corner
                    margin = 40
                    qr_size = 70
                    qr_rect = fitz.Rect(margin, page_h - margin - qr_size - 15, margin + qr_size, page_h - margin - 15)
                    new_page.insert_image(qr_rect, stream=qr_bytes)
                
                # Check for ID Placeholder in the PDF
                id_rects = new_page.search_for(id_placeholder)
                if id_rects:
                    id_rect = id_rects[0]
                    new_page.draw_rect(id_rect, color=(1, 1, 1), fill=(1, 1, 1), width=0)
                    id_width = fitz.get_text_length(cert_id, fontname="Helvetica", fontsize=10)
                    id_center_x = (id_rect.x0 + id_rect.x1) / 2 - (id_width / 2)
                    id_center_y = (id_rect.y0 + id_rect.y1) / 2
                    new_page.insert_text((id_center_x, id_center_y), cert_id, fontsize=10, fontname="Helvetica")
                else:
                    # Fallback directly under the QR code
                    margin = 40
                    new_page.insert_text((margin, page_h - margin), cert_id, fontsize=9, fontname="Helvetica")
            
            # Save the PDF certificate
            safe_name = name.replace(" ", "_").replace(".", "").replace("/", "").replace("\\", "")
            output_path = os.path.join(output_folder, f"certificate_{safe_name}.pdf")
            
            new_pdf.save(output_path)
            new_pdf.close()
            generated_files.append(output_path)
            
            if progress_callback:
                progress_callback(index + 1, total_names, name)
                
        pdf_template.close()
        return generated_files
    except Exception as e:
        if 'pdf_template' in locals() and pdf_template:
            try:
                pdf_template.close()
            except:
                pass
        raise e

# --- Example Usage ---
if __name__ == '__main__':
    pdf_template = 'template.pdf'
    excel_file = 'names.xlsx'
    output_dir = 'certificates_output'
    placeholder = 'NAME_PLACEHOLDER'

    print("Running in CLI mode...")
    if not os.path.exists(pdf_template) or not os.path.exists(excel_file):
        print("Please ensure template.pdf and names.xlsx are present to run in CLI mode.")
        print("Otherwise, launch the GUI with python gui.py or streamlit run app.py")
    else:
        try:
            def cli_callback(idx, total, name):
                print(f"[{idx}/{total}] Generated certificate for {name}")
                
            files = add_names_to_certificates(
                pdf_template, excel_file, output_dir, placeholder, 
                event_name="Testing Workshop", issue_date="2026-07-28",
                progress_callback=cli_callback
            )
            print(f"\nSuccessfully generated {len(files)} certificates in '{output_dir}'.")
            print("The names list file has been updated with Certificate IDs.")
        except Exception as e:
            print(f"Error: {e}")