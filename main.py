import sys
import os
from PyQt5.QtWidgets import QApplication
from config import TEMPLATES_DIR, TEMPLATE_FILENAME
from ui import DocumentGeneratorApp

if __name__ == '__main__':
    # Pastikan struktur folder template ada sebelum aplikasi berjalan
    if not os.path.exists(TEMPLATES_DIR):
        os.makedirs(TEMPLATES_DIR)
        # Jangan memanggil QMessageBox sebelum QApplication dibuat (akan memicu error)
        # Tampilkan pesan di console agar script tetap aman dijalankan dari terminal
        print(f"La carpeta 'templates' acaba de ser creada. Por favor, coloque el archivo '{TEMPLATE_FILENAME}' Dentro de él, luego vuelve a ejecutar la aplicación.")
        sys.exit() # Keluar agar user dapat menaruh template

    app = QApplication(sys.argv)
    window = DocumentGeneratorApp()
    window.show()
    sys.exit(app.exec_())
