from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QLineEdit, QPushButton, QMessageBox,
    QScrollArea, QGridLayout, QTextEdit, QGroupBox, QFileDialog, QDateEdit
)
from PyQt5.QtCore import Qt
from datetime import date
from fields import FIELD_DEFINITIONS
from document_handler import DocumentHandler

class DocumentGeneratorApp(QWidget):
    """Aplikasi untuk menginput data dan menghasilkan dokumen Word dari template."""
    def __init__(self):
        super().__init__()
        # Dictionary untuk menyimpan referensi QLineEdit/QTextEdit
        self.input_widgets = {}
        self.document_handler = DocumentHandler()
        self.setWindowTitle("Generador de Informes de Pruebas Térmicas")
        self.setStyleSheet("""
            QWidget {
                background-color: #f5f5f5;
                font-size: 14px;
                font-family: 'Segoe UI', Arial, sans-serif;
                color: #333;
            }
            QLabel {
                color: #555;
            }
            QGroupBox {
                font-weight: bold;
                border: 2px solid #ddd;
                border-radius: 8px;
                margin-top: 1ex;
                background-color: #ffffff;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 10px 0 10px;
                color: #2c3e50;
            }
            QLineEdit, QTextEdit, QDateEdit {
                border: 1px solid #ccc;
                border-radius: 4px;
                padding: 8px;
                background-color: #fff;
            }
            QLineEdit:focus, QTextEdit:focus, QDateEdit:focus {
                border-color: #3498db;
            }
            QPushButton {
                background-color: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #3498db, stop:1 #2980b9);
                color: white;
                border: none;
                border-radius: 6px;
                padding: 12px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #5dade2, stop:1 #3498db);
            }
            QPushButton:pressed {
                background-color: #21618c;
            }
            QScrollArea {
                border: none;
            }
        """)
        self.init_ui()

    def init_ui(self):
        """Membangun antarmuka pengguna."""
        main_layout = QVBoxLayout(self)

        # --- Judul ---
        title = QLabel("Ingresar Datos Para el Informe")
        title.setAlignment(Qt.AlignCenter)
        title.setStyleSheet("font-size: 20px; font-weight: bold; margin-bottom: 10px; color: #2C3E50;")
        main_layout.addWidget(title)

        # --- Scroll Area untuk banyak input ---
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        content_widget = QWidget()
        form_layout = QVBoxLayout(content_widget)
        form_layout.setSpacing(15)

        # --- Membuat Group Box untuk Input agar terstruktur ---

        # Group 0: Header information (Test Plan number, Revision, Issue Date)
        self.create_input_group(form_layout, "Encabezado - Información del documento", [
            "TEST_PLAN_NUMBER", "REVISION", "ISSUE_DATE"
        ])

        # Group 1: Product and Sample Information (1-6)
        self.create_input_group(form_layout, "Información del producto y muestras (Campos 1-6)", [
            "SAMPLE_DESCRIPTION", "DATE_OF_RECEPTION",
            "COMMERCIAL_BRAND", "MODEL_REFERENCE", "FAMILY", "INSULATION_CLASS"
        ])

        # Group 2: Electrical and Application Details (7-12)
        self.create_input_group(form_layout, "Detalles eléctricos y de aplicación (Campos 7-12)", [
            "LIGHT_SOURCE", "NOMINAL_VOLTAGE", "POWER",
            "FREQUENCY", "LS_CURRENT_VOLTAGE", "APPLICATION"
        ])

        # Group 3: Test and Conclusion Details (13-17)
        self.create_input_group(form_layout, "Detalles de Pruebas y Conclusiones (Campos 13-17)", [
            "EXTENSION_MODELS", "TESTS_PERFORMED", "DATE_OF_TEST",
            "TEST_STANDARDS", "CONCLUSIONS"
        ])


        scroll.setWidget(content_widget)
        main_layout.addWidget(scroll)

        # --- Tombol Generate ---
        self.generate_button = QPushButton("GENERAR DOCUMENTO DE WORD (.docx)")
        self.generate_button.clicked.connect(self.generate_document)
        main_layout.addWidget(self.generate_button)

        # --- Informasi Template ---
        from config import TEMPLATE_FILENAME, TEMPLATES_DIR
        info = QLabel(
            f"**Plantilla utilizada:** '{TEMPLATE_FILENAME}'\n"
            f"Asegúrate de que este archivo esté en la carpeta: '{TEMPLATES_DIR}'"
        )
        info.setStyleSheet("font-size: 10px; color: gray; margin-top: 5px;")
        main_layout.addWidget(info)

        self.setLayout(main_layout)
        self.resize(600, 700) # Ukuran awal yang lebih besar untuk banyak input

    def create_input_group(self, parent_layout, title, keys):
        """Membuat group box untuk input yang terorganisir."""
        group_box = QGroupBox(title)
        grid_layout = QGridLayout()
        grid_layout.setSpacing(10)

        row = 0
        col = 0

        for key in keys:
            definition = FIELD_DEFINITIONS[key]

            label = QLabel(f"{definition['label']}:")
            label.setStyleSheet("font-weight: normal;")

            if key in ["SAMPLE_DESCRIPTION", "EXTENSION_MODELS", "TESTS_PERFORMED", "CONCLUSIONS"]:
                # Gunakan QTextEdit untuk input multi-baris
                input_field = QTextEdit()
                # Tekan Tab akan berpindah fokus ke field berikutnya (default QTextEdit memasukkan tab ke teks)
                try:
                    input_field.setTabChangesFocus(True)
                except Exception:
                    # Jika method tidak ada di versi Qt tertentu, pasangkan event filter fallback later if needed
                    pass
                input_field.setMinimumHeight(60)
                grid_layout.addWidget(label, row, 0, 1, 2)
                grid_layout.addWidget(input_field, row + 1, 0, 1, 2)
                row += 2
                col = 0
            elif key in ["DATE_OF_RECEPTION", "DATE_OF_TEST", "ISSUE_DATE"]:
                # Gunakan QDateEdit untuk input tanggal dengan kalender
                input_field = QDateEdit()
                input_field.setCalendarPopup(True)
                input_field.setDisplayFormat("dd/MM/yyyy")
                input_field.setDate(date.today())

                # Tata letak 2 kolom
                grid_layout.addWidget(label, row, col)
                grid_layout.addWidget(input_field, row + 1, col)

                col = 1 - col # Pindah ke kolom 1 (atau kembali ke 0)
                if col == 0:
                    row += 2 # Pindah ke baris baru setelah 2 kolom terisi
            else:
                # Gunakan QLineEdit untuk input satu baris
                input_field = QLineEdit()
                input_field.setMinimumHeight(30)

                # Tata letak 2 kolom
                grid_layout.addWidget(label, row, col)
                grid_layout.addWidget(input_field, row + 1, col)

                col = 1 - col # Pindah ke kolom 1 (atau kembali ke 0)
                if col == 0:
                    row += 2 # Pindah ke baris baru setelah 2 kolom terisi

            self.input_widgets[key] = input_field

        group_box.setLayout(grid_layout)
        parent_layout.addWidget(group_box)

    def generate_document(self):
        """Logika utama untuk membaca input, memuat template, mengganti placeholder, dan menyimpan file."""

        # 1. Kumpulkan semua data input
        replacement_data = {}
        all_required_filled = True

        for key, definition in FIELD_DEFINITIONS.items():
            input_widget = self.input_widgets.get(key)
            if isinstance(input_widget, QLineEdit):
                value = input_widget.text().strip()
            elif isinstance(input_widget, QTextEdit):
                value = input_widget.toPlainText().strip()
            elif isinstance(input_widget, QDateEdit):
                value = input_widget.date().toString("dd/MM/yyyy")
            else:
                continue

            # Check jika field penting kosong
            if key in ["MODEL_REFERENCE", "COMMERCIAL_BRAND", "TEST_STANDARDS", "CONCLUSIONS"] and not value:
                all_required_filled = False
                QMessageBox.warning(self, "Input Kosong", f"columna obligatoria ('{definition['label']}') no puede estar vacío.")
                return

            replacement_data[definition['placeholder']] = value

        if not all_required_filled:
            return

        # Delegate to DocumentHandler
        self.document_handler.generate_document(replacement_data, self)
