# Bulk Certificate Generator with Supabase Authentication 🎓

A modern, fast, and secure system for generating bulk personalized certificates and verifying their authenticity online. 

This tool reads names from an Excel sheet or CSV file, generates a unique Certificate ID, builds a verification QR code, embeds both elements directly into a PDF template, and synchronizes the metadata with a **Supabase Database** for instant online verification.

---

## 🏗️ Project Architecture

* **Core Generator (`script.py`)**: The python engine that handles name calculations, QR code creation, and database uploads.
* **Desktop App (`gui.py`)**: A native desktop application built with `CustomTkinter` featuring dark mode.
* **Web Dashboard (`app.py`)**: A browser-based web application built with `Streamlit` that supports drag-and-drop file uploads, names sheet previews, and ZIP packaging downloads.
* **Verification Portal (`verify/`)**: A client-side static HTML/CSS/JS page (`verify/index.html`) using Supabase JS SDK that verifies scanned credentials in real-time.

---

## 🛠️ Step-by-Step Setup

### Step 1: Install Dependencies
Ensure you have Python 3.8+ installed on your computer. Run the following command in your terminal to install the necessary libraries:
```bash
pip install -r requirements.txt
```

### Step 2: Set up your Supabase Database
1. Create a free account at [Supabase](https://supabase.com/).
2. Create a new project and open the **SQL Editor** from the left-hand menu.
3. Run the following SQL script to set up your `certificates` table and write-protect it using **Row Level Security (RLS)**:

```sql
-- 1. Create the certificates table
create table public.certificates (
  id text not null primary key,
  name text not null,
  event_name text not null,
  issue_date text not null,
  created_at timestamp with time zone default now()
);

-- 2. Turn on Row Level Security (RLS)
alter table public.certificates enable row level security;

-- 3. Create a policy to allow anyone to read records (required for QR code validation)
create policy "Allow public read access" 
  on public.certificates 
  for select 
  using (true);
```

### Step 3: Configure Environment Variables
1. Copy [.env.example](file:///.env.example) and rename it to `.env` in the project root:
2. Fill in your project API keys and database details:
   ```env
   SUPABASE_URL=https://your-project-id.supabase.co
   SUPABASE_KEY=your_supabase_secret_service_role_key
   VERIFICATION_BASE_URL=https://your-domain.vercel.app/
   ```
> [!IMPORTANT]
> * **`SUPABASE_KEY`**: To keep your database secure, use the **`service_role` (secret)** key from your Supabase Dashboard (Settings -> API) for your local generator. This key is private and allows your Python script to insert new rows while the public is restricted to read-only queries.
> * **`.env` Security**: The `.env` file is already listed in `.gitignore` so your secret keys will never be pushed to public repositories.

---

## 🚀 How to Run the Generators

### Option 1: Streamlit Web Dashboard (Recommended)
This launches a premium browser-based interface where you can configure the event, upload assets, preview datasets, adjust sizes on-the-fly, and download results as a single ZIP file:
```bash
streamlit run app.py
```

### Option 2: CustomTkinter Desktop App
This launches a native, dark-themed GUI window on your desktop with folder browsers and sliders:
```bash
python gui.py
```

### Option 3: Command Line (CLI)
Place your `template.pdf` and `names.xlsx` in the root folder, configure `script.py` example defaults, and run:
```bash
python script.py
```

---

## 🌐 How to Deploy the Verification Portal on Vercel

You can host the public verification site for free on Vercel:

1. Import your project repository into Vercel.
2. Under **Project Settings**, set the **Root Directory** to `verify`.
3. Click **Deploy**.
4. Once Vercel generates your live URL (e.g. `https://ieee-verify.vercel.app/`), copy it and paste it as the `VERIFICATION_BASE_URL` in your local `.env` file.

*Note: Since the public Anon Key is already coded into `verify/index.html`, no environment variables need to be set in the Vercel dashboard settings!*

---

## 📝 Document Formatting Guidelines

### 1. The Names Spreadsheet (`names.xlsx` or `.csv`)
* Must contain a column header labeled `"Name"` (case-insensitive).
* If a `"Certificate ID"` column is found, the script will preserve existing IDs. Otherwise, it will automatically generate unique IDs (e.g., `CERT-A1B2C3D4`) and write them back into your source file.

### 2. PDF Certificate Template
For automated alignment, place the following invisible/matching-colored keywords inside your certificate template:
* `NAME_PLACEHOLDER`: Centers the recipient's name at this location.
* `QR_PLACEHOLDER`: Draws the verification QR code over this rectangle.
* `ID_PLACEHOLDER`: Centers the Certificate ID text over this rectangle.

*Fallback*: If `QR_PLACEHOLDER` or `ID_PLACEHOLDER` are missing in your PDF, they will automatically be drawn in a compact layout in the **bottom-left corner** of the page by default.

---

## 🎨 Custom Fonts & Typography

The system supports **any TrueType (`.ttf`) or OpenType (`.otf`) custom font** alongside standard PDF fonts.

* **Calligraffitti Font**: Included out-of-the-box in `fonts/Calligraffitti-Regular.ttf` for elegant calligraphy certificates.
* **Auto-Discovery Folder**: Drop any `.ttf` or `.otf` font file into the [`fonts/`](file:///fonts) directory. It will automatically be detected and listed across the Streamlit Web Dashboard, Desktop GUI, and Python script.
* **Streamlit Web Dashboard**: Select `"➕ Upload Custom Font (.ttf / .otf)..."` from the sidebar dropdown to upload any font directly through your browser.
* **Desktop GUI**: Click the `"📁 Browse Font (.ttf / .otf)..."` button to select any font file from anywhere on your computer.

