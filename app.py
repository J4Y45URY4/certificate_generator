import streamlit as st
import pandas as pd
import fitz  # PyMuPDF
import os
import io
import zipfile
import shutil
import uuid
import datetime
import qrcode
import dotenv
from script import add_names_to_certificates, get_available_fonts, resolve_font

# Load environment variables
dotenv.load_dotenv()

# Page configuration
st.set_page_config(
    page_title="Certificate Generator",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Minimal clean styling
st.markdown("""
<style>
    /* Clean, professional styling */
    .block-container {
        padding-top: 2rem;
        padding-bottom: 2rem;
        max-width: 1400px;
    }
    
    /* Subtle border for preview image */
    img {
        border-radius: 4px;
        box-shadow: 0 1px 3px rgba(0, 0, 0, 0.3);
    }
</style>
""", unsafe_allow_html=True)


def render_live_preview(
    pdf_bytes_or_path, 
    sample_name="Jane Doe", 
    placeholder_text="NAME_PLACEHOLDER",
    font_name="Calligraffitti", 
    font_size=37, 
    custom_font_path=None,
    include_qr=True, 
    qr_placeholder="QR_PLACEHOLDER", 
    id_placeholder="ID_PLACEHOLDER",
    qr_size=45, 
    id_size=8, 
    verification_url="", 
    sample_id="CERT-PREVIEW1",
    highlight_boxes=False, 
    dpi=130
):
    """Renders a single-page certificate preview to PNG bytes in-memory."""
    try:
        if isinstance(pdf_bytes_or_path, bytes):
            doc = fitz.open(stream=pdf_bytes_or_path, filetype="pdf")
        else:
            doc = fitz.open(pdf_bytes_or_path)
            
        page = doc[0]
        w, h = page.rect.width, page.rect.height
        
        font_alias, font_path, font_obj = resolve_font(font_name=font_name, font_file=custom_font_path)
        
        rects = page.search_for(placeholder_text)
        name_found = len(rects) > 0
        if name_found:
            r = rects[0]
            cx = (r.x0 + r.x1) / 2
            cy = (r.y0 + r.y1) / 2
            page.draw_rect(r, color=(1, 1, 1), fill=(1, 1, 1), width=0)
            
            if font_obj and font_path:
                page.insert_font(fontname=font_alias, fontfile=font_path)
                tw = font_obj.text_length(sample_name, fontsize=font_size)
                page.insert_text((cx - (tw / 2), cy), sample_name, fontsize=font_size, fontname=font_alias)
            else:
                tw = fitz.get_text_length(sample_name, fontname=font_alias, fontsize=font_size)
                page.insert_text((cx - (tw / 2), cy), sample_name, fontsize=font_size, fontname=font_alias)
                
            if highlight_boxes:
                page.draw_rect(
                    fitz.Rect(cx - (tw / 2) - 4, cy - font_size, cx + (tw / 2) + 4, cy + 4), 
                    color=(0.2, 0.5, 0.9), 
                    width=1.5
                )
                
        qr_found = False
        id_found = False
        if include_qr:
            qr = qrcode.QRCode(version=1, box_size=10, border=1)
            qr_target = f"{verification_url}?id={sample_id}" if verification_url else sample_id
            qr.add_data(qr_target)
            qr.make(fit=True)
            qr_img = qr.make_image(fill_color="black", back_color="white")
            buf = io.BytesIO()
            qr_img.save(buf)
            qr_bytes = buf.getvalue()
            
            qr_rects = page.search_for(qr_placeholder)
            if qr_rects:
                qr_found = True
                qr_rect = qr_rects[0]
                page.draw_rect(qr_rect, color=(1, 1, 1), fill=(1, 1, 1), width=0)
                page.insert_image(qr_rect, stream=qr_bytes)
                if highlight_boxes:
                    page.draw_rect(qr_rect, color=(0.1, 0.7, 0.3), width=1.5)
            else:
                margin = 40
                qr_rect = fitz.Rect(margin, h - margin - qr_size - 10, margin + qr_size, h - margin - 10)
                page.insert_image(qr_rect, stream=qr_bytes)
                if highlight_boxes:
                    page.draw_rect(qr_rect, color=(0.1, 0.7, 0.3), width=1.5)
                    
            id_rects = page.search_for(id_placeholder)
            if id_rects:
                id_found = True
                id_r = id_rects[0]
                page.draw_rect(id_r, color=(1, 1, 1), fill=(1, 1, 1), width=0)
                iw = fitz.get_text_length(sample_id, fontname="Helvetica", fontsize=id_size)
                icx = (id_r.x0 + id_r.x1) / 2 - (iw / 2)
                icy = (id_r.y0 + id_r.y1) / 2
                page.insert_text((icx, icy), sample_id, fontsize=id_size, fontname="Helvetica")
                if highlight_boxes:
                    page.draw_rect(fitz.Rect(icx - 2, icy - id_size, icx + iw + 2, icy + 2), color=(0.9, 0.6, 0.1), width=1.5)
            else:
                margin = 40
                page.insert_text((margin, h - margin), sample_id, fontsize=id_size, fontname="Helvetica")

        pix = page.get_pixmap(dpi=dpi)
        img_bytes = pix.tobytes("png")
        doc.close()
        
        return img_bytes, {
            "name_found": name_found,
            "qr_found": qr_found,
            "id_found": id_found,
            "width": w,
            "height": h,
            "error": None
        }
    except Exception as e:
        return None, {
            "name_found": False,
            "qr_found": False,
            "id_found": False,
            "width": 0,
            "height": 0,
            "error": str(e)
        }


# Session State Management
if 'generated_files' not in st.session_state:
    st.session_state.generated_files = []
if 'zip_data' not in st.session_state:
    st.session_state.zip_data = None
if 'use_sample_template' not in st.session_state:
    st.session_state.use_sample_template = False

# Application Header
st.title("Certificate Generator")
st.caption("Generate personalized certificates from a PDF template and a spreadsheet with live preview.")

# Sidebar Settings
with st.sidebar:
    st.header("Settings")
    
    show_live_preview = st.checkbox(
        "Show live preview on side", 
        value=True, 
        help="Toggle side-by-side view to preview certificate changes in real-time."
    )
    
    st.subheader("Typography")
    placeholder = st.text_input(
        "Name Placeholder",
        value="NAME_PLACEHOLDER",
        help="Text in the PDF to replace with recipient names."
    )
    
    # Available fonts
    available_fonts = get_available_fonts()
    seen_custom = set()
    custom_options = []
    for f_name in ["Calligraffitti"] + list(available_fonts.keys()):
        if available_fonts.get(f_name) is not None:
            if not f_name.endswith(" Regular") and not f_name.endswith("-Regular"):
                if f_name not in seen_custom:
                    seen_custom.add(f_name)
                    custom_options.append(f_name)
                    
    standard_options = [
        "Times-italic", "Times-Roman", "Times-bold", 
        "Helvetica", "Helvetica-oblique", "Helvetica-bold", 
        "Courier", "Courier-oblique", "Courier-bold"
    ]
    
    upload_choice = "Upload custom font (.ttf, .otf)"
    all_font_choices = custom_options + standard_options + [upload_choice]
    
    font_choice = st.selectbox(
        "Font",
        options=all_font_choices,
        index=0
    )
    
    custom_font_path = None
    selected_font_name = font_choice
    
    if font_choice == upload_choice:
        uploaded_font = st.file_uploader(
            "Upload font file", 
            type=["ttf", "otf"]
        )
        if uploaded_font is not None:
            fonts_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "fonts")
            os.makedirs(fonts_dir, exist_ok=True)
            saved_font_path = os.path.join(fonts_dir, uploaded_font.name)
            with open(saved_font_path, "wb") as f:
                f.write(uploaded_font.getbuffer())
            custom_font_path = saved_font_path
            selected_font_name = os.path.splitext(uploaded_font.name)[0]
            st.caption(f"Loaded font: {selected_font_name}")
        else:
            selected_font_name = "Times-italic"
    else:
        custom_font_path = available_fonts.get(font_choice)
        
    font_size = st.slider(
        "Font size",
        min_value=12,
        max_value=90,
        value=37,
        step=1
    )
    
    st.subheader("Verification & QR Code")
    include_qr = st.checkbox("Include QR code and ID", value=True)
    
    qr_placeholder = "QR_PLACEHOLDER"
    id_placeholder = "ID_PLACEHOLDER"
    verification_url = ""
    qr_size = 45
    id_size = 8
    
    if include_qr:
        verification_url = st.text_input(
            "Verification URL", 
            value=os.getenv("VERIFICATION_BASE_URL", ""),
            help="Base URL for QR code links."
        )
        event_name = st.text_input("Event name", value="Certificate Workshop")
        issue_date = st.text_input("Issue date", value=datetime.date.today().strftime("%Y-%m-%d"))
        
        with st.expander("Placeholder details"):
            qr_placeholder = st.text_input("QR placeholder", value="QR_PLACEHOLDER")
            id_placeholder = st.text_input("ID placeholder", value="ID_PLACEHOLDER")
            qr_size = st.slider("Fallback QR size (pt)", min_value=20, max_value=120, value=45)
            id_size = st.slider("ID text size", min_value=5, max_value=24, value=8)
    else:
        event_name = "General Event"
        issue_date = datetime.date.today().strftime("%Y-%m-%d")

    st.subheader("Output Options")
    local_output = st.checkbox("Save copy to certificates_output folder", value=True)
    
    db_url = os.getenv("SUPABASE_URL")
    db_key = os.getenv("SUPABASE_KEY")
    if db_url and db_key and not str(db_url).startswith("https://your-"):
        st.caption("Database sync: Connected")
    else:
        st.caption("Database sync: Local mode only")


# Column Layout
if show_live_preview:
    col_input, col_preview = st.columns([1, 1], gap="large")
else:
    col_input = st.container()
    col_preview = None

effective_pdf_bytes = None
effective_pdf_name = None

with col_input:
    st.subheader("Template & Data")
    
    # 1. Template PDF
    uploaded_pdf = st.file_uploader(
        "PDF Template", 
        type=["pdf"],
        help="Certificate template with placeholder text."
    )
    
    has_default_template = os.path.isfile("template.pdf")
    if uploaded_pdf is not None:
        effective_pdf_bytes = uploaded_pdf.getvalue()
        effective_pdf_name = uploaded_pdf.name
        st.session_state.use_sample_template = False
    elif has_default_template:
        use_sample = st.checkbox("Use local template.pdf", value=st.session_state.use_sample_template)
        st.session_state.use_sample_template = use_sample
        if use_sample:
            with open("template.pdf", "rb") as f:
                effective_pdf_bytes = f.read()
            effective_pdf_name = "template.pdf"
    
    # 2. Recipient Names File
    uploaded_names = st.file_uploader(
        "Recipient Names (Excel / CSV)", 
        type=["xlsx", "xls", "csv"],
        help="Spreadsheet containing a column named 'Name'."
    )
    
    has_default_names = os.path.isfile("names.xlsx")
    names_df = None
    names_filename = None
    
    if uploaded_names is not None:
        names_filename = uploaded_names.name
        try:
            if uploaded_names.name.endswith('.csv'):
                names_df = pd.read_csv(uploaded_names)
            else:
                names_df = pd.read_excel(uploaded_names)
        except Exception as e:
            st.error(f"Error reading file: {e}")
    elif st.session_state.use_sample_template and has_default_names:
        try:
            names_df = pd.read_excel("names.xlsx")
            names_filename = "names.xlsx"
            st.caption("Loaded names from local 'names.xlsx'")
        except Exception:
            pass

    first_recipient_name = "Jane Doe"
    if names_df is not None:
        name_col = None
        for col in names_df.columns:
            if str(col).strip().lower() == 'name':
                name_col = col
                break
        if name_col is None and len(names_df.columns) > 0:
            name_col = names_df.columns[0]
            
        id_col = None
        for col in names_df.columns:
            if str(col).strip().lower() in ['certificate id', 'certificate_id', 'id']:
                id_col = col
                break
                
        preview_df = names_df.copy()
        if id_col is None:
            id_col = 'Certificate ID'
            preview_df[id_col] = [f"CERT-{uuid.uuid4().hex[:8].upper()}" for _ in range(len(preview_df))]
        else:
            for idx, val in enumerate(preview_df[id_col]):
                if pd.isna(val) or str(val).strip() in ['', 'None', 'nan']:
                    preview_df.at[idx, id_col] = f"CERT-{uuid.uuid4().hex[:8].upper()}"
                    
        for n in preview_df[name_col]:
            if pd.notna(n) and str(n).strip():
                first_recipient_name = str(n).strip()
                break

        st.caption(f"{len(names_df)} recipients found in column '{name_col}'")
        st.dataframe(
            preview_df[[name_col, id_col]].rename(columns={name_col: "Name", id_col: "Certificate ID"}),
            use_container_width=True,
            height=150
        )

    st.subheader("Generate")
    ready_to_generate = (effective_pdf_bytes is not None) and (names_df is not None)
    
    if not ready_to_generate:
        missing = []
        if effective_pdf_bytes is None:
            missing.append("PDF template")
        if names_df is None:
            missing.append("recipient list")
        st.info(f"Upload a {' and '.join(missing)} to generate certificates.")
        
    generate_btn = st.button(
        "Generate Certificates", 
        disabled=not ready_to_generate, 
        use_container_width=True,
        type="primary"
    )

    if generate_btn:
        progress_bar = st.progress(0.0)
        status_text = st.empty()
        
        temp_dir = os.path.join("temp_uploads")
        os.makedirs(temp_dir, exist_ok=True)
        
        temp_pdf_path = os.path.join(temp_dir, "temp_template.pdf")
        with open(temp_pdf_path, "wb") as f:
            f.write(effective_pdf_bytes)
            
        ext = ".csv" if (names_filename and names_filename.endswith(".csv")) else ".xlsx"
        temp_names_path = os.path.join(temp_dir, f"temp_names{ext}")
        if uploaded_names is not None:
            with open(temp_names_path, "wb") as f:
                f.write(uploaded_names.getbuffer())
        else:
            names_df.to_excel(temp_names_path, index=False)
            
        output_folder = "certificates_output" if local_output else "temp_output"
        
        def progress_callback(index, total, name):
            progress_bar.progress(index / total)
            status_text.text(f"Processing ({index}/{total}): {name}")
            
        try:
            files = add_names_to_certificates(
                pdf_template_path=temp_pdf_path,
                excel_path=temp_names_path,
                output_folder=output_folder,
                placeholder_text=placeholder,
                font_name=selected_font_name,
                text_size=font_size,
                include_qr=include_qr,
                qr_placeholder=qr_placeholder,
                id_placeholder=id_placeholder,
                qr_size=qr_size,
                id_size=id_size,
                event_name=event_name,
                issue_date=issue_date,
                verification_base_url=verification_url,
                progress_callback=progress_callback,
                font_file=custom_font_path
            )
            
            status_text.text("Creating ZIP archive...")
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                for file_path in files:
                    zip_file.write(file_path, os.path.basename(file_path))
                zip_file.write(temp_names_path, f"names_with_ids{ext}")
                
            zip_buffer.seek(0)
            st.session_state.generated_files = files
            st.session_state.zip_data = zip_buffer.getvalue()
            
            if not local_output and os.path.exists(output_folder):
                shutil.rmtree(output_folder)
            shutil.rmtree(temp_dir)
            
            status_text.empty()
            progress_bar.empty()
            st.success(f"Generated {len(files)} certificates successfully.")
            
        except Exception as e:
            st.error(f"Generation error: {e}")
            if os.path.exists(temp_dir):
                shutil.rmtree(temp_dir)


# Live Preview Panel
preview_container = col_preview if show_live_preview else st.container()

if show_live_preview or effective_pdf_bytes is not None:
    with preview_container:
        st.subheader("Live Preview")
        
        if effective_pdf_bytes is not None:
            prev_row1, prev_row2 = st.columns([3, 2])
            with prev_row1:
                sample_name_input = st.text_input(
                    "Sample recipient name",
                    value=first_recipient_name
                )
            with prev_row2:
                highlight_zones = st.checkbox("Highlight placeholders", value=False)
                preview_dpi = 130
            
            img_bytes, preview_status = render_live_preview(
                pdf_bytes_or_path=effective_pdf_bytes,
                sample_name=sample_name_input or "Recipient Name",
                placeholder_text=placeholder,
                font_name=selected_font_name,
                font_size=font_size,
                custom_font_path=custom_font_path,
                include_qr=include_qr,
                qr_placeholder=qr_placeholder,
                id_placeholder=id_placeholder,
                qr_size=qr_size,
                id_size=id_size,
                verification_url=verification_url,
                sample_id="CERT-PREVIEW1",
                highlight_boxes=highlight_zones,
                dpi=preview_dpi
            )
            
            # Simple status messages without emojis
            status_items = []
            if preview_status.get("name_found"):
                status_items.append(f"Name placeholder '{placeholder}' detected")
            else:
                status_items.append(f"Notice: Placeholder '{placeholder}' not found in template")
                
            if include_qr:
                if preview_status.get("qr_found"):
                    status_items.append(f"QR placeholder '{qr_placeholder}' detected")
                else:
                    status_items.append("QR position: using bottom-left corner")
            
            st.caption(" | ".join(status_items))
            
            if img_bytes is not None:
                st.image(img_bytes, use_container_width=True)
            elif preview_status.get("error"):
                st.error(f"Preview rendering error: {preview_status.get('error')}")
        else:
            st.info("Upload a PDF template to preview it live here.")
            if has_default_template:
                if st.button("Load default template.pdf"):
                    st.session_state.use_sample_template = True
                    st.rerun()

# Download Section
if st.session_state.zip_data is not None:
    st.divider()
    st.subheader("Download")
    d_col1, d_col2 = st.columns([1, 1])
    with d_col1:
        st.download_button(
            label="Download All Certificates (ZIP)",
            data=st.session_state.zip_data,
            file_name="certificates.zip",
            mime="application/zip",
            use_container_width=True
        )
    with d_col2:
        if local_output:
            st.write(f"Files saved locally in: `{os.path.abspath('certificates_output')}`")
