from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QLineEdit, QPushButton, QMessageBox,
    QScrollArea, QGridLayout, QTextEdit, QGroupBox, QFileDialog, QDateEdit
)
from PyQt5.QtCore import Qt, QEvent
from PyQt5.QtGui import QPixmap
from datetime import date
import json
import os
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
                font-size: 14px;
                font-family: 'Segoe UI', Arial, sans-serif;
                background-color: #f5f5f5;
                color: #333;
            }
            QGroupBox {
                font-weight: bold;
                border: 2px solid #ddd;
                border-radius: 8px;
                margin-top: 10px;
                background-color: #ffffff;
                padding: 10px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 5px 0 5px;
                color: #2c3e50;
                font-size: 16px;
            }
            QLabel {
                color: #555;
            }
            QLineEdit, QTextEdit, QDateEdit {
                border: 1px solid #ccc;
                border-radius: 4px;
                padding: 5px;
                background-color: #fff;
            }
            QLineEdit:focus, QTextEdit:focus, QDateEdit:focus {
                border-color: #3498db;
            }
            QPushButton {
                border-radius: 6px;
                padding: 8px 16px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #2980b9;
            }
            QMessageBox QPushButton {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #95a5a6, stop:1 #7f8c8d);
                color: white;
                padding: 8px 16px;
                border-radius: 4px;
                font-weight: bold;
                border: none;
            }
            QMessageBox QPushButton:hover {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #7f8c8d, stop:1 #6c7b7d);
            }
            QMessageBox QPushButton:pressed {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #6c7b7d, stop:1 #566573);
            }
        """)
        self.init_ui()

    def init_ui(self):
        """Membangun antarmuka pengguna."""
        main_layout = QVBoxLayout(self)

        # --- Logo ---
        logo_label = QLabel()
        logo_path = os.path.join(os.path.dirname(__file__), "logo vibia.png")
        if os.path.exists(logo_path):
            pixmap = QPixmap(logo_path).scaledToWidth(200, Qt.SmoothTransformation)
            logo_label.setPixmap(pixmap)
        else:
            logo_label.setText("Logo not found")
            logo_label.setStyleSheet("color: red; font-weight: bold;")
        logo_label.setAlignment(Qt.AlignCenter)
        logo_label.setStyleSheet("margin-bottom: 10px;")
        main_layout.addWidget(logo_label)

        # --- Judul ---
        title = QLabel("Ingresar Datos Para el Informe")
        title.setAlignment(Qt.AlignCenter)
        title.setStyleSheet("""
            font-size: 24px;
            font-weight: bold;
            margin-bottom: 15px;
            color: #808080;
            font-family: 'Gotham', sans-serif;
        """)
        main_layout.addWidget(title)

        # --- Scroll Area untuk banyak input ---
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setStyleSheet("""
            QScrollArea {
                border: 1px solid #ddd;
                border-radius: 8px;
                background-color: #fafafa;
            }
            QScrollArea QWidget {
                background-color: #fafafa;
            }
        """)
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
        self.generate_button.setStyleSheet("""
            QPushButton {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #95a5a6, stop:1 #7f8c8d);
                color: white;
                padding: 12px;
                border-radius: 8px;
                font-weight: bold;
                border: none;
            }
            QPushButton:hover {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #7f8c8d, stop:1 #6c7b7d);
            }
            QPushButton:pressed {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #6c7b7d, stop:1 #566573);
            }
        """)
        self.generate_button.clicked.connect(self.generate_document)
        main_layout.addWidget(self.generate_button)

        # --- Tombol Save and Load ---
        button_layout = QHBoxLayout()
        self.save_button = QPushButton("GUARDAR DATOS")
        self.save_button.setStyleSheet("""
            QPushButton {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #95a5a6, stop:1 #7f8c8d);
                color: white;
                padding: 8px 16px;
                border-radius: 6px;
                font-weight: bold;
                border: none;
            }
            QPushButton:hover {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #7f8c8d, stop:1 #6c7b7d);
            }
            QPushButton:pressed {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #6c7b7d, stop:1 #566573);
            }
        """)
        self.save_button.clicked.connect(self.save_data)
        button_layout.addWidget(self.save_button)

        self.load_button = QPushButton("CARGAR DATOS")
        self.load_button.setStyleSheet("""
            QPushButton {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #95a5a6, stop:1 #7f8c8d);
                color: white;
                padding: 8px 16px;
                border-radius: 6px;
                font-weight: bold;
                border: none;
            }
            QPushButton:hover {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #7f8c8d, stop:1 #6c7b7d);
            }
            QPushButton:pressed {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #6c7b7d, stop:1 #566573);
            }
        """)
        self.load_button.clicked.connect(self.load_data)
        button_layout.addWidget(self.load_button)
        main_layout.addLayout(button_layout)

        # --- Informasi Template ---
        from config import TEMPLATE_FILENAME, TEMPLATES_DIR
        info = QLabel(
            f"**Plantilla utilizada:** '{TEMPLATE_FILENAME}'\n"
            f"Asegúrate de que este archivo esté en la carpeta: '{TEMPLATES_DIR}'"
        )
        info.setStyleSheet("font-size: 12px; color: #7f8c8d; margin-top: 10px; font-style: italic; background-color: #ecf0f1; padding: 5px; border-radius: 4px;")
        main_layout.addWidget(info)

        self.setLayout(main_layout)
        self.resize(600, 700) # Ukuran awal yang lebih besar untuk banyak input

    def closeEvent(self, event):
        """Handle close event with confirmation dialog."""
        reply = QMessageBox.question(
            self,
            "Confirmar salida",
            "¿Estás seguro de que quieres salir de la aplicación?",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        if reply == QMessageBox.Yes:
            event.accept()
        else:
            event.ignore()

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
                QMessageBox.warning(self, "entrada vacía", f"columna obligatoria ('{definition['label']}') no puede estar vacío.")
                return

            replacement_data[definition['placeholder']] = value

        if not all_required_filled:
            return

        # Delegate to DocumentHandler
        self.document_handler.generate_document(replacement_data, self)

    def save_data(self):
        """Save input data to a JSON file."""
        data = {}
        for key, widget in self.input_widgets.items():
            if isinstance(widget, QLineEdit):
                data[key] = widget.text().strip()
            elif isinstance(widget, QTextEdit):
                data[key] = widget.toPlainText().strip()
            elif isinstance(widget, QDateEdit):
                data[key] = widget.date().toString("dd/MM/yyyy")

        file_path, _ = QFileDialog.getSaveFileName(self, "Guardar Datos", "", "JSON Files (*.json)")
        if file_path:
            try:
                with open(file_path, 'w', encoding='utf-8') as f:
                    json.dump(data, f, ensure_ascii=False, indent=4)
                QMessageBox.information(self, "Guardado", "Datos guardados exitosamente.")
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Error al guardar: {str(e)}")

    def load_data(self):
        """Load input data from a JSON file."""
        file_path, _ = QFileDialog.getOpenFileName(self, "Cargar Datos", "", "JSON Files (*.json)")
        if file_path:
            try:
                with open(file_path, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                for key, value in data.items():
                    widget = self.input_widgets.get(key)
                    if widget:
                        if isinstance(widget, QLineEdit):
                            widget.setText(value)
                        elif isinstance(widget, QTextEdit):
                            widget.setPlainText(value)
                        elif isinstance(widget, QDateEdit):
                            from PyQt5.QtCore import QDate
                            date = QDate.fromString(value, "dd/MM/yyyy")
                            widget.setDate(date)
                QMessageBox.information(self, "Cargado", "Datos cargados exitosamente.")
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Error al cargar: {str(e)}")
