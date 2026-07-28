import streamlit as st
import pandas as pd
import fitz
import os
import io
import zipfile
import shutil
import uuid
import datetime
import dotenv
from script import add_names_to_certificates

# Load environment variables from .env
dotenv.load_dotenv()

# Configure Page
st.set_page_config(
    page_title="Certificate Generator Dashboard",
    page_icon="🎓",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for modern design and premium look
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Outfit:wght@300;400;600;800&display=swap');
    
    html, body, [class*="css"] {
        font-family: 'Outfit', sans-serif;
    }
    
    .main-title {
        background: linear-gradient(135deg, #60A5FA 0%, #3B82F6 50%, #1D4ED8 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        font-size: 3rem;
        font-weight: 800;
        margin-bottom: 5px;
    }
    
    .subtitle {
        color: #9CA3AF;
        font-size: 1.1rem;
        margin-bottom: 30px;
    }
    
    .card {
        border-radius: 12px;
        background-color: #1F2937;
        padding: 24px;
        margin-bottom: 20px;
        border: 1px solid #374151;
    }
    
    .sidebar-header {
        font-weight: 600;
        font-size: 1.2rem;
        color: #60A5FA;
        margin-bottom: 15px;
    }
</style>
""", unsafe_allow_html=True)

# Application Header
st.markdown('<div class="main-title">Certificate Generator Dashboard</div>', unsafe_allow_html=True)
st.markdown('<div class="subtitle">Bulk personalize certificates with a modern web interface.</div>', unsafe_allow_html=True)

# Initial State setup
if 'generated_files' not in st.session_state:
    st.session_state.generated_files = []
if 'zip_data' not in st.session_state:
    st.session_state.zip_data = None

# Sidebar layout for configuration
with st.sidebar:
    st.markdown('<div class="sidebar-header">🛠️ Configurations</div>', unsafe_allow_html=True)
    
    placeholder = st.text_input(
        "Name Placeholder Key",
        value="NAME_PLACEHOLDER",
        help="The text inside the template PDF that will be replaced with names."
    )
    
    font_name = st.selectbox(
        "Font Style",
        options=[
            "Times-italic", "Times-Roman", "Times-bold", 
            "Helvetica", "Helvetica-oblique", "Helvetica-bold", 
            "Courier", "Courier-oblique", "Courier-bold"
        ],
        index=0,
        help="Font family for printing names."
    )
    
    font_size = st.slider(
        "Font Size",
        min_value=10,
        max_value=100,
        value=37,
        step=1,
        help="Adjust the size of the name text."
    )
    
    st.markdown("---")
    st.markdown('<div class="sidebar-header">🏅 Event & Database Sync</div>', unsafe_allow_html=True)
    
    event_name = st.text_input(
        "Event Name",
        value="Certificate Workshop",
        help="The name of the workshop or event to store in the verification DB."
    )
    
    issue_date = st.text_input(
        "Issue Date",
        value=datetime.date.today().strftime("%Y-%m-%d"),
        help="Date displayed in verification checks (YYYY-MM-DD)."
    )

    include_qr = st.checkbox("Include QR Code & Certificate ID", value=True)
    
    qr_placeholder = "QR_PLACEHOLDER"
    id_placeholder = "ID_PLACEHOLDER"
    verification_url = ""
    qr_size = 45
    id_size = 8
    
    if include_qr:
        verification_url = st.text_input(
            "Verification Web URL", 
            value=os.getenv("VERIFICATION_BASE_URL", ""), 
            help="Base URL of verification portal (e.g. https://yourdomain.vercel.app/)"
        )
        qr_placeholder = st.text_input("QR Key", value="QR_PLACEHOLDER", help="Placeholder text in template PDF for the QR code position.")
        id_placeholder = st.text_input("ID Key", value="ID_PLACEHOLDER", help="Placeholder text in template PDF for the ID text position.")
        qr_size = st.slider("QR Code Size (Fallback)", min_value=15, max_value=150, value=45, step=1, help="Size of the fallback QR code image (in pt).")
        id_size = st.slider("ID Font Size", min_value=5, max_value=30, value=8, step=1, help="Font size of the Certificate ID text.")
    
    # DB Status indicator
    st.markdown("---")
    db_url = os.getenv("SUPABASE_URL")
    db_key = os.getenv("SUPABASE_KEY")
    if db_url and db_key and not str(db_url).startswith("https://your-"):
        st.sidebar.success("⚡ Supabase Database: Connected")
    else:
        st.sidebar.warning("⚠️ Supabase Database: Offline (Local Only)")
        
    st.markdown("### Output Settings")
    local_output = st.checkbox("Save directly to local folder", value=True, help="Saves files to 'certificates_output' directory in the project.")

# Main area split into columns
col1, col2 = st.columns([1, 1], gap="large")

with col1:
    st.markdown("### 📥 1. Upload Files")
    
    # 1. Template PDF Uploader
    uploaded_pdf = st.file_uploader(
        "Upload PDF Template", 
        type=["pdf"],
        help="Drag and drop your PDF certificate template here."
    )
    
    # 2. Names File Uploader
    uploaded_names = st.file_uploader(
        "Upload Names List (Excel / CSV)", 
        type=["xlsx", "xls", "csv"],
        help="Upload an Excel or CSV file containing a column named 'Name'."
    )

with col2:
    st.markdown("### 📊 2. Preview & Generate")
    
    names_df = None
    if uploaded_names is not None:
        try:
            if uploaded_names.name.endswith('.csv'):
                names_df = pd.read_csv(uploaded_names)
            else:
                names_df = pd.read_excel(uploaded_names)
            
            # Find candidate column
            name_col = None
            for col in names_df.columns:
                if str(col).strip().lower() == 'name':
                    name_col = col
                    break
            
            if name_col is None and len(names_df.columns) > 0:
                name_col = names_df.columns[0]
            
            # Locate or generate preview IDs
            id_col = None
            for col in names_df.columns:
                if str(col).strip().lower() in ['certificate id', 'certificate_id', 'id']:
                    id_col = col
                    break
            
            # Generate preview IDs to show in screen
            preview_df = names_df.copy()
            if id_col is None:
                id_col = 'Certificate ID'
                preview_df[id_col] = [f"CERT-{uuid.uuid4().hex[:8].upper()}" for _ in range(len(preview_df))]
            else:
                # Fill missing
                for idx, val in enumerate(preview_df[id_col]):
                    if pd.isna(val) or str(val).strip() in ['', 'None', 'nan']:
                        preview_df.at[idx, id_col] = f"CERT-{uuid.uuid4().hex[:8].upper()}"
            
            st.success(f"Successfully loaded list! Found {len(names_df)} names using column: **'{name_col}'**")
            
            # Display preview
            st.dataframe(
                preview_df[[name_col, id_col]].rename(columns={name_col: "Names to Print", id_col: "Certificate ID Preview"}), 
                use_container_width=True, 
                height=180
            )
            
        except Exception as e:
            st.error(f"Error reading names file: {e}")
    else:
        st.info("Upload a names file to preview names list.")

    # Actions trigger
    ready_to_generate = uploaded_pdf is not None and names_df is not None
    
    generate_btn = st.button(
        "⚡ Generate & Sync Certificates", 
        disabled=not ready_to_generate, 
        use_container_width=True,
        type="primary"
    )

    if generate_btn:
        st.markdown("---")
        progress_bar = st.progress(0.0)
        status_text = st.empty()
        
        # Save uploaded files temporarily to pass to the script function
        temp_dir = os.path.join("temp_uploads")
        os.makedirs(temp_dir, exist_ok=True)
        
        temp_pdf_path = os.path.join(temp_dir, "temp_template.pdf")
        temp_names_path = os.path.join(temp_dir, f"temp_names.{'csv' if uploaded_names.name.endswith('.csv') else 'xlsx'}")
        
        with open(temp_pdf_path, "wb") as f:
            f.write(uploaded_pdf.getbuffer())
        
        with open(temp_names_path, "wb") as f:
            f.write(uploaded_names.getbuffer())
            
        output_folder = "certificates_output" if local_output else "temp_output"
        
        def progress_callback(index, total, name):
            progress_fraction = index / total
            progress_bar.progress(progress_fraction)
            status_text.text(f"Generating ({index}/{total}): {name}")
            
        try:
            # Generate certificates
            files = add_names_to_certificates(
                pdf_template_path=temp_pdf_path,
                excel_path=temp_names_path,
                output_folder=output_folder,
                placeholder_text=placeholder,
                font_name=font_name,
                text_size=font_size,
                include_qr=include_qr,
                qr_placeholder=qr_placeholder,
                id_placeholder=id_placeholder,
                qr_size=qr_size,
                id_size=id_size,
                event_name=event_name,
                issue_date=issue_date,
                verification_base_url=verification_url,
                progress_callback=progress_callback
            )
            
            status_text.text("Packaging files...")
            
            # Package ZIP in memory
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                # Add certificates
                for file_path in files:
                    zip_file.write(file_path, os.path.basename(file_path))
                # Add the updated Excel sheet containing final IDs
                orig_extension = os.path.splitext(uploaded_names.name)[1]
                zip_file.write(temp_names_path, f"names_with_ids{orig_extension}")
                
            zip_buffer.seek(0)
            
            st.session_state.generated_files = files
            st.session_state.zip_data = zip_buffer.getvalue()
            
            # Clean up temp uploads and output if not local output
            if not local_output and os.path.exists(output_folder):
                shutil.rmtree(output_folder)
            shutil.rmtree(temp_dir)
            
            st.balloons()
            st.success(f"Successfully generated and synced {len(files)} certificates!")
            
        except Exception as e:
            st.error(f"Failed to generate certificates: {e}")
            if os.path.exists(temp_dir):
                shutil.rmtree(temp_dir)

# Download Section
if st.session_state.zip_data is not None:
    st.markdown("### 📥 3. Download Results")
    
    col_dl1, col_dl2 = st.columns([1, 1])
    
    with col_dl1:
        st.download_button(
            label="💾 Download Certificates & Excel Sheet (ZIP)",
            data=st.session_state.zip_data,
            file_name="certificates_package.zip",
            mime="application/zip",
            use_container_width=True
        )
    with col_dl2:
        if local_output:
            st.info(f"Certificates and updated spreadsheet are saved locally in `{os.path.abspath('certificates_output')}`")
        else:
            st.info("Files are compiled in-memory for download.")
