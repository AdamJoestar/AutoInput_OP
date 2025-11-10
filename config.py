import os

# --- Konfigurasi File ---
TEMPLATES_DIR = os.path.join(os.path.dirname(__file__), "templates")
# Menggunakan nama file yang diunggah pengguna: New_Template.docx
TEMPLATE_FILENAME = "New_Template.docx"
TEMPLATE_PATH = os.path.join(TEMPLATES_DIR, TEMPLATE_FILENAME)
