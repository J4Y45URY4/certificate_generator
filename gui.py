import os
import threading
import datetime
import tkinter as tk
from tkinter import filedialog, messagebox
import customtkinter as ctk
import dotenv
from script import add_names_to_certificates, get_available_fonts

# Load environment variables from .env
dotenv.load_dotenv()

# Set up CustomTkinter appearance and theme
ctk.set_appearance_mode("Dark")  # Options: "System", "Dark", "Light"
ctk.set_default_color_theme("blue")  # Options: "blue", "green", "dark-blue"

class CertificateGeneratorGUI(ctk.CTk):
    def __init__(self):
        super().__init__()

        # Configure window
        self.title("Modern Certificate Generator")
        self.geometry("750x780")
        self.resizable(False, False)

        # Main layout
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # Title Label
        self.title_label = ctk.CTkLabel(
            self, 
            text="Certificate Generator", 
            font=ctk.CTkFont(family="Outfit", size=26, weight="bold"),
            text_color="#60A5FA"
        )
        self.title_label.grid(row=0, column=0, padx=20, pady=(20, 5), sticky="w")

        self.subtitle_label = ctk.CTkLabel(
            self,
            text="Generate bulk certificates with unique IDs, QR codes, and Supabase DB verification.",
            font=ctk.CTkFont(size=13),
            text_color="#9CA3AF"
        )
        self.subtitle_label.grid(row=1, column=0, padx=20, pady=(0, 15), sticky="w")

        # Create main frame
        self.main_frame = ctk.CTkScrollableFrame(self, height=520)
        self.main_frame.grid(row=2, column=0, padx=20, pady=0, sticky="nsew")
        self.main_frame.grid_columnconfigure(0, weight=1)

        # --- SECTION 1: File Selection ---
        self.file_frame = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        self.file_frame.grid(row=0, column=0, padx=15, pady=5, sticky="ew")
        self.file_frame.grid_columnconfigure(1, weight=1)

        # Template PDF
        self.template_label = ctk.CTkLabel(self.file_frame, text="PDF Template:", font=ctk.CTkFont(weight="bold"))
        self.template_label.grid(row=0, column=0, padx=(0, 10), pady=8, sticky="w")
        
        self.template_entry = ctk.CTkEntry(self.file_frame, placeholder_text="Path to template.pdf...")
        self.template_entry.grid(row=0, column=1, padx=(0, 10), pady=8, sticky="ew")
        
        self.template_btn = ctk.CTkButton(self.file_frame, text="Browse", width=80, command=self.browse_template)
        self.template_btn.grid(row=0, column=2, pady=8)

        # Names list
        self.names_label = ctk.CTkLabel(self.file_frame, text="Names List:", font=ctk.CTkFont(weight="bold"))
        self.names_label.grid(row=1, column=0, padx=(0, 10), pady=8, sticky="w")
        
        self.names_entry = ctk.CTkEntry(self.file_frame, placeholder_text="Path to names.xlsx or names.csv...")
        self.names_entry.grid(row=1, column=1, padx=(0, 10), pady=8, sticky="ew")
        
        self.names_btn = ctk.CTkButton(self.file_frame, text="Browse", width=80, command=self.browse_names)
        self.names_btn.grid(row=1, column=2, pady=8)

        # Output folder
        self.output_label = ctk.CTkLabel(self.file_frame, text="Output Folder:", font=ctk.CTkFont(weight="bold"))
        self.output_label.grid(row=2, column=0, padx=(0, 10), pady=8, sticky="w")
        
        self.output_entry = ctk.CTkEntry(self.file_frame, placeholder_text="Path to certificates_output...")
        self.output_entry.insert(0, os.path.abspath("certificates_output"))
        self.output_entry.grid(row=2, column=1, padx=(0, 10), pady=8, sticky="ew")
        
        self.output_btn = ctk.CTkButton(self.file_frame, text="Browse", width=80, command=self.browse_output)
        self.output_btn.grid(row=2, column=2, pady=8)

        # --- SECTION 2: General Formatting Settings ---
        self.settings_frame = ctk.CTkFrame(self.main_frame)
        self.settings_frame.grid(row=1, column=0, padx=15, pady=10, sticky="ew")
        self.settings_frame.grid_columnconfigure((1, 3), weight=1)

        # Placeholder text
        self.placeholder_label = ctk.CTkLabel(self.settings_frame, text="Name Key:", font=ctk.CTkFont(weight="bold"))
        self.placeholder_label.grid(row=0, column=0, padx=15, pady=(15, 10), sticky="w")
        
        self.placeholder_entry = ctk.CTkEntry(self.settings_frame, placeholder_text="e.g. NAME_PLACEHOLDER")
        self.placeholder_entry.insert(0, "NAME_PLACEHOLDER")
        self.placeholder_entry.grid(row=0, column=1, padx=(0, 15), pady=(15, 10), sticky="ew")

        # Font Family & Custom Font Selection
        self.available_fonts = get_available_fonts()
        self.custom_font_path = None
        
        custom_names = []
        for k in ["Calligraffitti"] + list(self.available_fonts.keys()):
            if self.available_fonts.get(k) is not None and not k.endswith(" Regular") and not k.endswith("-Regular"):
                if k not in custom_names:
                    custom_names.append(k)

        standard_names = [
            "Times-italic", "Times-Roman", "Times-bold", 
            "Helvetica", "Helvetica-oblique", "Helvetica-bold", 
            "Courier", "Courier-oblique", "Courier-bold"
        ]
        font_options = custom_names + standard_names

        self.font_label = ctk.CTkLabel(self.settings_frame, text="Font Style:", font=ctk.CTkFont(weight="bold"))
        self.font_label.grid(row=0, column=2, padx=15, pady=(15, 10), sticky="w")
        
        self.font_combobox = ctk.CTkOptionMenu(self.settings_frame, values=font_options, command=self.on_font_selected)
        self.font_combobox.set("Calligraffitti" if "Calligraffitti" in font_options else "Times-italic")
        self.font_combobox.grid(row=0, column=3, padx=(0, 15), pady=(15, 10), sticky="ew")

        # Font Size
        self.size_label = ctk.CTkLabel(self.settings_frame, text="Font Size:", font=ctk.CTkFont(weight="bold"))
        self.size_label.grid(row=1, column=0, padx=15, pady=10, sticky="w")
        
        self.size_slider = ctk.CTkSlider(self.settings_frame, from_=10, to=100, number_of_steps=90, command=self.update_slider_label)
        self.size_slider.set(37)
        self.size_slider.grid(row=1, column=1, columnspan=2, padx=(0, 15), pady=10, sticky="ew")
        
        self.size_value_label = ctk.CTkLabel(self.settings_frame, text="37 pt", font=ctk.CTkFont(weight="bold"))
        self.size_value_label.grid(row=1, column=3, padx=(0, 15), pady=10, sticky="w")

        # Row 2: Browse Any Custom Font
        self.custom_font_label = ctk.CTkLabel(self.settings_frame, text="Custom Font File:", font=ctk.CTkFont(weight="bold"))
        self.custom_font_label.grid(row=2, column=0, padx=15, pady=(5, 15), sticky="w")

        self.custom_font_btn = ctk.CTkButton(
            self.settings_frame, 
            text="📁 Browse Font (.ttf / .otf)...", 
            fg_color="#374151", 
            hover_color="#4B5563",
            command=self.browse_custom_font
        )
        self.custom_font_btn.grid(row=2, column=1, columnspan=3, padx=(0, 15), pady=(5, 15), sticky="ew")

        # --- SECTION 3: Event & Supabase DB configurations ---
        self.db_frame = ctk.CTkFrame(self.main_frame)
        self.db_frame.grid(row=2, column=0, padx=15, pady=10, sticky="ew")
        self.db_frame.grid_columnconfigure((1, 3), weight=1)

        # Title for DB Section
        self.db_title_label = ctk.CTkLabel(self.db_frame, text="Event Details & Verification Sync", font=ctk.CTkFont(size=14, weight="bold"), text_color="#10B981")
        self.db_title_label.grid(row=0, column=0, columnspan=4, padx=15, pady=10, sticky="w")

        # Event Name
        self.event_label = ctk.CTkLabel(self.db_frame, text="Event Name:", font=ctk.CTkFont(weight="bold"))
        self.event_label.grid(row=1, column=0, padx=15, pady=8, sticky="w")
        
        self.event_entry = ctk.CTkEntry(self.db_frame, placeholder_text="e.g. Certificate Workshop")
        self.event_entry.insert(0, "Certificate Workshop")
        self.event_entry.grid(row=1, column=1, padx=(0, 15), pady=8, sticky="ew")

        # Issue Date (prefilled with today)
        self.date_label = ctk.CTkLabel(self.db_frame, text="Issue Date:", font=ctk.CTkFont(weight="bold"))
        self.date_label.grid(row=1, column=2, padx=15, pady=8, sticky="w")
        
        self.date_entry = ctk.CTkEntry(self.db_frame, placeholder_text="YYYY-MM-DD")
        today_str = datetime.date.today().strftime("%Y-%m-%d")
        self.date_entry.insert(0, today_str)
        self.date_entry.grid(row=1, column=3, padx=(0, 15), pady=8, sticky="ew")

        # QR and Verification details
        self.qr_checkbox = ctk.CTkCheckBox(self.db_frame, text="Enable QR Code / Online Verification", font=ctk.CTkFont(weight="bold"), command=self.toggle_qr_options)
        self.qr_checkbox.select()
        self.qr_checkbox.grid(row=2, column=0, columnspan=4, padx=15, pady=10, sticky="w")

        # Base URL
        self.verification_label = ctk.CTkLabel(self.db_frame, text="Verify URL:", font=ctk.CTkFont(weight="bold"))
        self.verification_label.grid(row=3, column=0, padx=15, pady=8, sticky="w")
        
        self.verification_entry = ctk.CTkEntry(self.db_frame, placeholder_text="e.g. https://domain.vercel.app/")
        dotenv_base_url = os.getenv("VERIFICATION_BASE_URL", "")
        self.verification_entry.insert(0, dotenv_base_url)
        self.verification_entry.grid(row=3, column=1, columnspan=3, padx=(0, 15), pady=8, sticky="ew")

        # Placeholders
        self.qr_placeholder_label = ctk.CTkLabel(self.db_frame, text="QR Key:", font=ctk.CTkFont(weight="bold"))
        self.qr_placeholder_label.grid(row=4, column=0, padx=15, pady=(5, 15), sticky="w")
        
        self.qr_placeholder_entry = ctk.CTkEntry(self.db_frame, placeholder_text="QR_PLACEHOLDER")
        self.qr_placeholder_entry.insert(0, "QR_PLACEHOLDER")
        self.qr_placeholder_entry.grid(row=4, column=1, padx=(0, 15), pady=(5, 15), sticky="ew")

        self.id_placeholder_label = ctk.CTkLabel(self.db_frame, text="ID Key:", font=ctk.CTkFont(weight="bold"))
        self.id_placeholder_label.grid(row=4, column=2, padx=15, pady=(5, 15), sticky="w")
        
        self.id_placeholder_entry = ctk.CTkEntry(self.db_frame, placeholder_text="ID_PLACEHOLDER")
        self.id_placeholder_entry.insert(0, "ID_PLACEHOLDER")
        self.id_placeholder_entry.grid(row=4, column=3, padx=(0, 15), pady=(5, 15), sticky="ew")

        # --- SECTION 4: Progress Panel ---
        self.progress_frame = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        self.progress_frame.grid(row=3, column=0, padx=15, pady=15, sticky="ew")
        self.progress_frame.grid_columnconfigure(0, weight=1)

        # Progress bar
        self.progress_bar = ctk.CTkProgressBar(self.progress_frame)
        self.progress_bar.set(0.0)
        self.progress_bar.grid(row=0, column=0, pady=(0, 10), sticky="ew")

        # Status text
        self.status_label = ctk.CTkLabel(self.progress_frame, text="Status: Ready", font=ctk.CTkFont(size=12))
        self.status_label.grid(row=1, column=0, sticky="w")

        # Generate Button
        self.generate_btn = ctk.CTkButton(
            self,
            text="Generate & Sync Certificates",
            height=45,
            font=ctk.CTkFont(size=16, weight="bold"),
            command=self.start_generation
        )
        self.generate_btn.grid(row=3, column=0, padx=20, pady=15, sticky="ew")

        # DB connection indicator label on startup
        db_url = os.getenv("SUPABASE_URL")
        db_key = os.getenv("SUPABASE_KEY")
        if db_url and db_key and not str(db_url).startswith("https://your-"):
            self.db_title_label.configure(text="Event Details & Verification Sync (Supabase Active)")
        else:
            self.db_title_label.configure(text="Event Details & Verification Sync (Local Only / No DB Configured)", text_color="#F59E0B")

    # Helper methods
    def toggle_qr_options(self):
        state = "normal" if self.qr_checkbox.get() else "disabled"
        self.verification_entry.configure(state=state)
        self.qr_placeholder_entry.configure(state=state)
        self.id_placeholder_entry.configure(state=state)

    def browse_template(self):
        filename = filedialog.askopenfilename(
            title="Select PDF Template",
            filetypes=[("PDF files", "*.pdf")]
        )
        if filename:
            self.template_entry.delete(0, tk.END)
            self.template_entry.insert(0, os.path.abspath(filename))

    def browse_names(self):
        filename = filedialog.askopenfilename(
            title="Select Names File",
            filetypes=[("Excel or CSV files", "*.xlsx *.xls *.csv")]
        )
        if filename:
            self.names_entry.delete(0, tk.END)
            self.names_entry.insert(0, os.path.abspath(filename))

    def browse_output(self):
        directory = filedialog.askdirectory(title="Select Output Folder")
        if directory:
            self.output_entry.delete(0, tk.END)
            self.output_entry.insert(0, os.path.abspath(directory))

    def update_slider_label(self, value):
        self.size_value_label.configure(text=f"{int(value)} pt")

    def on_font_selected(self, choice):
        if choice in self.available_fonts and self.available_fonts[choice]:
            self.custom_font_path = self.available_fonts[choice]
        else:
            self.custom_font_path = None

    def browse_custom_font(self):
        filename = filedialog.askopenfilename(
            title="Select Custom Font File",
            filetypes=[("Font files", "*.ttf *.otf"), ("All files", "*.*")]
        )
        if filename:
            self.custom_font_path = os.path.abspath(filename)
            base_name = os.path.splitext(os.path.basename(filename))[0]
            current_values = list(self.font_combobox.cget("values"))
            if base_name not in current_values:
                current_values.insert(0, base_name)
                self.font_combobox.configure(values=current_values)
            self.font_combobox.set(base_name)
            self.custom_font_btn.configure(text=f"Font: {os.path.basename(filename)} (Click to change)")

    # Background threading for non-blocking generation
    def start_generation(self):
        template = self.template_entry.get().strip()
        names_file = self.names_entry.get().strip()
        output_dir = self.output_entry.get().strip()
        placeholder = self.placeholder_entry.get().strip()
        font_name = self.font_combobox.get()
        font_size = int(self.size_slider.get())
        
        include_qr = bool(self.qr_checkbox.get())
        event_name = self.event_entry.get().strip()
        issue_date = self.date_entry.get().strip()
        verification_url = self.verification_entry.get().strip()
        qr_placeholder = self.qr_placeholder_entry.get().strip()
        id_placeholder = self.id_placeholder_entry.get().strip()

        if not template or not names_file or not output_dir or not placeholder:
            messagebox.showerror("Error", "All file paths and name configurations must be filled out.")
            return

        if not event_name or not issue_date:
            messagebox.showerror("Error", "Event Name and Issue Date must be filled out.")
            return

        if not os.path.exists(template):
            messagebox.showerror("Error", f"PDF Template file not found: {template}")
            return

        if not os.path.exists(names_file):
            messagebox.showerror("Error", f"Names file not found: {names_file}")
            return

        font_file = self.custom_font_path or self.available_fonts.get(font_name)

        # Disable button to prevent double-clicks
        self.generate_btn.configure(state="disabled", text="Processing...")
        self.progress_bar.set(0.0)
        self.status_label.configure(text="Initializing generation...")

        # Run process in a separate thread
        thread = threading.Thread(
            target=self.run_generation,
            args=(template, names_file, output_dir, placeholder, font_name, font_size, 
                  include_qr, qr_placeholder, id_placeholder, event_name, issue_date, verification_url, font_file)
        )
        thread.daemon = True
        thread.start()

    def run_generation(self, template, names_file, output_dir, placeholder, font_name, font_size, 
                       include_qr, qr_placeholder, id_placeholder, event_name, issue_date, verification_url, font_file=None):
        def progress_callback(index, total, name):
            progress_fraction = index / total
            self.after(0, self.update_progress, progress_fraction, f"Generating certificate {index}/{total}: {name}")

        try:
            generated_files = add_names_to_certificates(
                pdf_template_path=template,
                excel_path=names_file,
                output_folder=output_dir,
                placeholder_text=placeholder,
                font_name=font_name,
                text_size=font_size,
                include_qr=include_qr,
                qr_placeholder=qr_placeholder,
                id_placeholder=id_placeholder,
                event_name=event_name,
                issue_date=issue_date,
                verification_base_url=verification_url,
                progress_callback=progress_callback,
                font_file=font_file
            )
            self.after(0, self.generation_finished, True, f"Successfully created {len(generated_files)} certificates!\nThe input file has been updated with Certificate IDs.\nOutput folder:\n{output_dir}")
        except Exception as e:
            self.after(0, self.generation_finished, False, str(e))

    def update_progress(self, fraction, text):
        self.progress_bar.set(fraction)
        self.status_label.configure(text=text)

    def generation_finished(self, success, message):
        self.generate_btn.configure(state="normal", text="Generate & Sync Certificates")
        if success:
            self.progress_bar.set(1.0)
            self.status_label.configure(text="Finished successfully!")
            messagebox.showinfo("Success", message)
        else:
            self.progress_bar.set(0.0)
            self.status_label.configure(text="Error occurred.")
            messagebox.showerror("Error", message)

if __name__ == "__main__":
    app = CertificateGeneratorGUI()
    app.mainloop()
