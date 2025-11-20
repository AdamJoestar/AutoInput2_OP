from PyQt5.QtWidgets import (
    QLabel, QLineEdit, QPushButton, QScrollArea, QGridLayout, QGroupBox, QFileDialog, QSpinBox, QTextEdit, QHBoxLayout, QWidget, QVBoxLayout, QDateEdit, QComboBox, QDialog, QDateTimeEdit
)
from PyQt5.QtCore import Qt, QDate, QDateTime
from PyQt5.QtGui import QPixmap
from fields import FIELD_DEFINITIONS
from screenshot import ScreenshotSelector
import tempfile
import os
import sys

class NoWheelComboBox(QComboBox):
    def wheelEvent(self, event):
        event.ignore()

class NoWheelSpinBox(QSpinBox):
    def wheelEvent(self, event):
        event.ignore()

class NoWheelDateEdit(QDateEdit):
    def wheelEvent(self, event):
        event.ignore()

def resource_path(relative_path):
    """ Mendapatkan path absolut ke resource, berfungsi untuk dev dan PyInstaller """
    try:
        # PyInstaller membuat folder sementara dan menyimpan path di _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)


class UIBuilder:
    def __init__(self, parent_app):
        """
        Initializes the UIBuilder.

        Args:
            parent_app (DocumentGeneratorApp): A reference to the main application instance.
        """
        self.parent_app = parent_app
        self.input_widgets = {}
        self.equipment_groups = []
        self.spin_boxes = {}
        self.temp_files = []  # List untuk melacak file sementara
        self.rebuilding = False

    def init_ui(self):
        """
        Initializes and builds the entire user interface (UI) for the application.

        Creates the title, row selection spin boxes, scroll area, and main buttons.
        """
        main_layout = self.parent_app.main_layout

        # --- Logo ---
        logo = QLabel()
        logo_path = resource_path("logo vibia.png")
        logo.setPixmap(QPixmap(logo_path).scaledToWidth(200, Qt.SmoothTransformation))
        logo.setAlignment(Qt.AlignCenter)
        logo.setStyleSheet("margin-bottom: 10px;")
        main_layout.addWidget(logo)

        # --- Judul ---
        title = QLabel("Ingresar Datos Para el Anexo")
        title.setAlignment(Qt.AlignCenter)
        title.setStyleSheet("""
            font-size: 24px;
            font-weight: bold;
            margin-bottom: 15px;
            color: #808080;
            font-family: 'Gotham', sans-serif;
            italic;
            font-style: italic;
        """)
        main_layout.addWidget(title)

        # --- Spin Boxes for Row Selection ---
        spin_layout = QHBoxLayout()
        label1 = QLabel("EQUIPOS Y MÉTODOS UTILIZADOS (max 12):")
        label1.setStyleSheet("font-weight: bold; color: #34495e; padding: 5px;")
        spin_layout.addWidget(label1)
        self.spin_equipment = NoWheelSpinBox()
        self.spin_equipment.setRange(1, 12)
        self.spin_equipment.setValue(12)
        self.spin_equipment.setStyleSheet("QSpinBox { border: 1px solid #bdc3c7; border-radius: 4px; padding: 5px; background-color: #ecf0f1; }")
        self.spin_equipment.valueChanged.connect(self.rebuild_form)
        spin_layout.addWidget(self.spin_equipment)

        label2 = QLabel("SONDA TOTAL (max 10):")
        label2.setStyleSheet("font-weight: bold; color: #34495e; padding: 5px;")
        spin_layout.addWidget(label2)
        self.spin_sonda = NoWheelSpinBox()
        self.spin_sonda.setRange(1, 10)
        self.spin_sonda.setValue(10)
        self.spin_sonda.setStyleSheet("QSpinBox { border: 1px solid #bdc3c7; border-radius: 4px; padding: 5px; background-color: #ecf0f1; }")
        self.spin_sonda.valueChanged.connect(self.rebuild_form)
        spin_layout.addWidget(self.spin_sonda)

        # Add a spin box for number of additional photos
        self.additional_photos_layout = QHBoxLayout()
        additional_photos_label = QLabel("Número de fotos adicionales (max 10):")
        additional_photos_label.setStyleSheet("font-weight: bold; color: #34495e; padding: 5px;")
        self.additional_photos_layout.addWidget(additional_photos_label)
        self.spin_additional_photos = NoWheelSpinBox()
        self.spin_additional_photos.setRange(0, 10)
        self.spin_additional_photos.setValue(0)
        self.spin_additional_photos.setStyleSheet("QSpinBox { border: 1px solid #bdc3c7; border-radius: 4px; padding: 5px; background-color: #ecf0f1; }")
        self.spin_additional_photos.valueChanged.connect(self.rebuild_form)
        self.additional_photos_layout.addWidget(self.spin_additional_photos)
        # Initially hide it, will be added in rebuild_form
        self.additional_photos_widget = QWidget()
        self.additional_photos_widget.setLayout(self.additional_photos_layout)

        main_layout.addLayout(spin_layout)

        # --- DateTime Selectors for Excel Data Filtering ---
        datetime_layout = QHBoxLayout()
        start_label = QLabel("Fecha y Hora Inicio para Datos Excel:")
        start_label.setStyleSheet("font-weight: bold; color: #34495e; padding: 5px;")
        datetime_layout.addWidget(start_label)
        self.start_datetime = QDateTimeEdit()
        self.start_datetime.setCalendarPopup(True)
        self.start_datetime.setDateTime(QDateTime.currentDateTime().addDays(-1))  # Default to yesterday
        self.start_datetime.setDisplayFormat("dd/MM/yyyy HH:mm:ss")  # Include seconds for precision
        self.start_datetime.setStyleSheet("QDateTimeEdit { border: 1px solid #bdc3c7; border-radius: 4px; padding: 5px; background-color: #ecf0f1; }")
        datetime_layout.addWidget(self.start_datetime)

        end_label = QLabel("Fecha y Hora Fin para Datos Excel:")
        end_label.setStyleSheet("font-weight: bold; color: #34495e; padding: 5px;")
        datetime_layout.addWidget(end_label)
        self.end_datetime = QDateTimeEdit()
        self.end_datetime.setCalendarPopup(True)
        self.end_datetime.setDateTime(QDateTime.currentDateTime())  # Default to now
        self.end_datetime.setDisplayFormat("dd/MM/yyyy HH:mm:ss")  # Include seconds for precision
        self.end_datetime.setStyleSheet("QDateTimeEdit { border: 1px solid #bdc3c7; border-radius: 4px; padding: 5px; background-color: #ecf0f1; }")
        datetime_layout.addWidget(self.end_datetime)

        main_layout.addLayout(datetime_layout)

        # --- Scroll Area untuk banyak input ---
        self.scroll = QScrollArea()
        self.scroll.setWidgetResizable(True)
        self.scroll.setStyleSheet("""
            QScrollArea {
                border: 1px solid #ddd;
                border-radius: 8px;
                background-color: #fafafa;
            }
            QScrollArea QWidget {
                background-color: #fafafa;
            }
        """)
        self.content_widget = QWidget()
        self.form_layout = QVBoxLayout(self.content_widget)
        self.form_layout.setSpacing(15)
        self.scroll.setWidget(self.content_widget)
        main_layout.addWidget(self.scroll)

        # --- Tombol Load Excel ---
        self.load_excel_button = QPushButton("CARGAR DATOS DE EXCEL")
        self.load_excel_button.setStyleSheet("""
            QPushButton {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #4CAF50, stop:1 #45a049);
                color: white;
                padding: 12px;
                border-radius: 8px;
                font-weight: bold;
                border: none;
            }
            QPushButton:hover {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #45a049, stop:1 #4CAF50);
            }
            QPushButton:pressed {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #4CAF50, stop:1 #3e8e41);
            }
        """)
        self.load_excel_button.clicked.connect(self.parent_app.load_excel_data)
        main_layout.addWidget(self.load_excel_button)

        # --- Tombol Generate ---
        self.generate_button = QPushButton("GENERAR DOCUMENTO DE WORD (.docx)")
        self.generate_button.setStyleSheet("""
            QPushButton {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #C7C7C7, stop:1 #9E9E9E);
                color: white;
                padding: 12px;
                border-radius: 8px;
                font-weight: bold;
                border: none;
            }
            QPushButton:hover {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #9E9E9E, stop:1 #D9D9D9);
            }
            QPushButton:pressed {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #D9D9D9, stop:1 #B0B0B0);
            }
        """)
        self.generate_button.clicked.connect(self.parent_app.generate_document)
        main_layout.addWidget(self.generate_button)

        # --- Informasi Template ---
        info = QLabel(
            f"**Plantilla utilizada:** '{self.parent_app.template_filename}'\n"
            f"Asegúrate de que este archivo esté en la carpeta: '{self.parent_app.templates_dir}'"
        )
        info.setStyleSheet("font-size: 12px; color: #7f8c8d; margin-top: 10px; font-style: italic; background-color: #ecf0f1; padding: 5px; border-radius: 4px;")
        main_layout.addWidget(info)

        self.rebuild_form()

    def rebuild_form(self):
        """
        Rebuilds the dynamic input form based on the values from the spin boxes.

        This function saves existing input values, clears all widgets from the form,
        then rebuilds the input groups according to the selected number of rows for
        "EQUIPOS" and "SONDA". Afterwards, it attempts to restore the saved values
        to the corresponding widgets.
        """
        if self.rebuilding:
            return
        self.rebuilding = True

        # Disconnect signals to prevent recursive calls during rebuild
        if hasattr(self, 'spin_equipment'):
            try:
                self.spin_equipment.valueChanged.disconnect(self.rebuild_form)
            except TypeError:
                pass
        if hasattr(self, 'spin_sonda'):
            try:
                self.spin_sonda.valueChanged.disconnect(self.rebuild_form)
            except TypeError:
                pass
        if hasattr(self, 'spin_additional_photos'):
            try:
                self.spin_additional_photos.valueChanged.disconnect(self.rebuild_form)
            except TypeError:
                pass

        # Save current input values before clearing
        current_values = {}
        for key, widget in self.input_widgets.items():
            if isinstance(widget, QLineEdit):
                current_values[key] = widget.text()
            elif isinstance(widget, QTextEdit):
                current_values[key] = widget.toPlainText()
            elif isinstance(widget, QDateEdit):
                current_values[key] = widget.date().toString("dd/MM/yyyy")
            elif isinstance(widget, QComboBox):
                current_values[key] = widget.currentText()

        # Save spin box values
        saved_equipment = self.spin_equipment.value()
        saved_sonda = self.spin_sonda.value()
        saved_additional = 0
        if hasattr(self, 'spin_additional_photos') and self.spin_additional_photos:
            saved_additional = self.spin_additional_photos.value()

        # Clear existing widgets from form_layout
        for i in reversed(range(self.form_layout.count())):
            item = self.form_layout.itemAt(i)
            if item.widget():
                item.widget().setParent(None)
            elif item.layout():
                # If it's a layout, remove it
                sub_layout = item.layout()
                while sub_layout.count():
                    sub_item = sub_layout.takeAt(0)
                    if sub_item.widget():
                        sub_item.widget().setParent(None)
                self.form_layout.removeItem(item)

        self.input_widgets = {}

        # Header
        title_label = QLabel("Encabezado - Información del documento")
        title_label.setStyleSheet("font-size: 18px; font-weight: bold; margin-bottom: 10px;")
        self.form_layout.addWidget(title_label)
        self.create_input_group(self.form_layout, "Encabezado - Información del documento", [
            "NO_TEST", "REV", "DATE"
        ])

        # 0. INFORMACIÓN DEL SOLICITANTE DEL ENSAYO
        title_label = QLabel("0. INFORMACIÓN DEL SOLICITANTE DEL ENSAYO")
        title_label.setStyleSheet("font-size: 18px; font-weight: bold; margin-bottom: 10px;")
        self.form_layout.addWidget(title_label)
        self.create_input_group(self.form_layout, "0. INFORMACIÓN DEL SOLICITANTE DEL ENSAYO", [
            "TEXT1", "TEXT4", "TEXT2", "TEXT5", "TEXT3"
        ])

        # 1. INFORMACIÓN GENERAL DEL PRODUCTO
        title_label = QLabel("1. INFORMACIÓN GENERAL DEL PRODUCTO")
        title_label.setStyleSheet("font-size: 18px; font-weight: bold; margin-bottom: 10px;")
        self.form_layout.addWidget(title_label)
        self.create_input_group(self.form_layout, "1. INFORMACIÓN GENERAL DEL PRODUCTO", [
            "TEXT6", "TEXT12", "TEXT7", "TEXT8", "TEXT13"
        ])

        # 1.1. CONDICIONES DEL ENSAYO
        self.create_input_group(self.form_layout, "1.1. CONDICIONES DEL ENSAYO", [
            "TEXT9", "TEXT10", "TEXT11"
        ])

        # 2. EQUIPOS Y MÉTODOS UTILIZADOS
        title_label = QLabel("2. EQUIPOS Y MÉTODOS UTILIZADOS")
        title_label.setStyleSheet("font-size: 18px; font-weight: bold; margin-bottom: 10px;")
        self.form_layout.addWidget(title_label)
        num_equip = self.spin_equipment.value()
        for i in range(1, num_equip + 1):
            self.create_input_group(self.form_layout, f"Row {i}", [
                f"EQUIPO{i}", f"MARCA{i}", f"TIPO{i}", f"FECHA{i}", f"OBSER{i}"
            ])

        # Auto-fill marca and tipo for all rows based on equipment selection
        # Also sync FECHA for SONDA TIPO T
        for i in range(1, num_equip + 1):
            equipo_key = f"EQUIPO{i}"
            marca_key = f"MARCA{i}"
            tipo_key = f"TIPO{i}"
            fecha_key = f"FECHA{i}"
            if equipo_key in self.input_widgets:
                equipo_widget = self.input_widgets[equipo_key]
                marca_widget = self.input_widgets.get(marca_key)
                tipo_widget = self.input_widgets.get(tipo_key)
                fecha_widget = self.input_widgets.get(fecha_key)
                if marca_widget and tipo_widget:
                    equipo_widget.currentTextChanged.connect(lambda text, mw=marca_widget, tw=tipo_widget, fw=fecha_widget, num=num_equip: self.auto_fill_marca_tipo(text, mw, tw, fw, num))

        # 2.1. MÉTODO DE ENSAYO
        # Removed as per user request

        # 3. TEMPERATURAS REGISTRADAS
        title_label = QLabel("3. TEMPERATURAS REGISTRADAS")
        title_label.setStyleSheet("font-size: 18px; font-weight: bold; margin-bottom: 10px;")
        self.form_layout.addWidget(title_label)
        num_sonda = self.spin_sonda.value()
        for i in range(1, num_sonda + 1):
            self.create_input_group(self.form_layout, f"Row {i} ", [
                f"PUNTO{i}", f"UNIDAD{i}", f"LIMITE{i}", f"TEMP{i}"
            ])

        # 3.1. GRÁFICA GENERADA
        self.create_input_group(self.form_layout, "3.1. GRÁFICA GENERADA", [
            "IMAGE1", "TITLE1", "IMAGE2"
        ])

        # 4. BILIZACIÓN TÉRMICA
        title_label = QLabel("4. ESTABILIZACIÓN TÉRMICA")
        title_label.setStyleSheet("font-size: 18px; font-weight: bold; margin-bottom: 10px;")
        self.form_layout.addWidget(title_label)
        # Removed TEXT14 as per user request
        for i in range(1, num_sonda + 1):
            self.create_input_group(self.form_layout, f"Row {i}", [
                f"MEDICI{i}", f"UNI{i}", f"VALMIN{i}", f"VALMAX{i}", f"DESVI{i}"
            ])

        # 5. RESULTADOS
        title_label = QLabel("5. RESULTADOS")
        title_label.setStyleSheet("font-size: 18px; font-weight: bold; margin-bottom: 10px;")
        self.form_layout.addWidget(title_label)
        for i in range(1, num_sonda + 1):
            self.create_input_group(self.form_layout, f"Row {i}", [
                f"PUNTODE{i}", f"UNIC{i}", f"TEMPE{i}", f"RESULT{i}"
            ])

        # 6. FOTOGRAFIAS
        title_label = QLabel("6. FOTOGRAFIAS")
        title_label.setStyleSheet("font-size: 18px; font-weight: bold; margin-bottom: 10px;")
        self.form_layout.addWidget(title_label)
        for i in range(3, 3 + num_sonda):  # IMAGE3 to IMAGE{2+num_sonda}, TITLE3 to TITLE{2+num_sonda}
            self.create_input_group(self.form_layout, f"Fotografía {i-2}", [
                f"IMAGE{i}"
            ])
            self.create_input_group(self.form_layout, f"Titulo {i-2}", [
                f"TITLE{i}"
            ])

        # 7. FOTO MONTAJE FINAL
        title_label = QLabel("7. FOTO MONTAJE FINAL")
        title_label.setStyleSheet("font-size: 18px; font-weight: bold; margin-bottom: 10px;")
        self.form_layout.addWidget(title_label)

        # Add the spinbox for additional photos
        self.form_layout.addWidget(self.additional_photos_widget)
        self.spin_additional_photos.blockSignals(True)
        self.spin_additional_photos.setValue(saved_additional)
        self.spin_additional_photos.blockSignals(False)

        num_additional = self.spin_additional_photos.value()
        for i in range(13, 13 + num_additional):  # IMAGE13 to IMAGE{12+num_additional}, TITLE13 to TITLE{12+num_additional}
            self.create_input_group(self.form_layout, f"Montaje {i-12}", [
                f"IMAGE{i}", f"TITLE{i}"
            ])

        # Restore saved values
        for key, value in current_values.items():
            if key in self.input_widgets:
                widget = self.input_widgets[key]
                if isinstance(widget, QLineEdit):
                    widget.setText(value)
                    widget.setCursorPosition(0)  # Reset cursor to avoid invalid selection
                    # Set default for OBSER fields if empty
                    if key.startswith("OBSER") and not value.strip():
                        widget.setText("-")
                        widget.setCursorPosition(0)
                elif isinstance(widget, QTextEdit):
                    widget.setPlainText(value)
                    # Set default template for TEXT_EST if empty
                    if key == "TEXT_EST" and not value.strip():
                        widget.setPlainText(self.parent_app.stabilization_template)
                elif isinstance(widget, QDateEdit):
                    from PyQt5.QtCore import QDate
                    date = QDate.fromString(value, "dd/MM/yyyy")
                    widget.setDate(date)
                elif isinstance(widget, QComboBox):
                    index = widget.findText(value)
                    if index >= 0:
                        widget.setCurrentIndex(index)

        # Sync related fields
        self.sync_related_fields(num_sonda)

        # Reconnect signals after rebuild
        if hasattr(self, 'spin_equipment'):
            self.spin_equipment.valueChanged.connect(self.rebuild_form)
        if hasattr(self, 'spin_sonda'):
            self.spin_sonda.valueChanged.connect(self.rebuild_form)
        if hasattr(self, 'spin_additional_photos'):
            self.spin_additional_photos.valueChanged.connect(self.rebuild_form)

        self.rebuilding = False

    def create_input_group(self, parent_layout, title, keys):
        """
        Creates a QGroupBox containing several input fields.

        Args:
            parent_layout (QLayout): The parent layout to which the group box will be added.
            title (str): The title for the QGroupBox.
            keys (list): A list of keys (from FIELD_DEFINITIONS) for the fields
                         to be created within this group.
        """
        group_box = QGroupBox(title)
        group_box.setStyleSheet("""
            QGroupBox {
                font-weight: bold;
                margin-top: 10px;
                border: 2px solid #bdc3c7;
                border-radius: 8px;
                background-color: #ffffff;
                padding: 10px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 5px 0 5px;
                color: #34495e;
                font-size: 14px;
            }
        """)
        grid_layout = QGridLayout()
        grid_layout.setSpacing(10)
        
        row = 0
        col = 0
        
        for key in keys:
            definition = FIELD_DEFINITIONS[key]

            label = QLabel(f"{definition['label']}:")
            label.setStyleSheet("color: #34495e; font-weight: bold; font-size: 12px;")

            if definition['type'] == "text":
                if key in ["TEXT1", "TEXT4", "TEXT2", "TEXT5", "TEXT3", "TEXT6", "TEXT7", "TEXT8", "TEXT9", "TEXT10", "TEXT11", "TEXT13", "TEXT15"]:
                    input_field = QTextEdit()
                    input_field.setMinimumHeight(60)
                    input_field.setStyleSheet("""
                        QTextEdit {
                            border: 1px solid #bdc3c7;
                            border-radius: 4px;
                            padding: 5px;
                            background-color: #ffffff;
                            font-size: 12px;
                        }
                        QTextEdit:focus {
                            border-color: #3498db;
                        }
                    """)
                    grid_layout.addWidget(label, row, 0, 1, 2)
                    grid_layout.addWidget(input_field, row + 1, 0, 1, 2)
                else:
                    input_field = QLineEdit()
                    input_field.setMinimumHeight(30)
                    input_field.setStyleSheet("""
                        QLineEdit {
                            border: 1px solid #bdc3c7;
                            border-radius: 4px;
                            padding: 5px;
                            background-color: #ffffff;
                            font-size: 12px;
                        }
                        QLineEdit:focus {
                            border-color: #3498db;
                        }
                    """)
                    grid_layout.addWidget(label, row, 0, 1, 2)
                    grid_layout.addWidget(input_field, row + 1, 0, 1, 2)
                    if key.startswith("TEMP") or key.startswith("VALMIN") or key.startswith("VALMAX") or key.startswith("DESVI") or key.startswith("TEMPE"):
                        from PyQt5.QtGui import QDoubleValidator
                        input_field.setValidator(QDoubleValidator(0.0, 9999.99, 2))
            elif definition['type'] == "date":
                input_field = QDateEdit()
                input_field.setCalendarPopup(True)
                input_field.setMinimumHeight(30)
                input_field.setDate(QDate.currentDate())  # Set default to today's date
                input_field.setStyleSheet("""
                    QDateEdit {
                        border: 1px solid #bdc3c7;
                        border-radius: 4px;
                        padding: 5px;
                        background-color: #ffffff;
                        font-size: 12px;
                    }
                    QDateEdit:focus {
                        border-color: #3498db;
                    }
                """)
                grid_layout.addWidget(label, row, 0, 1, 2)
                grid_layout.addWidget(input_field, row + 1, 0, 1, 2)
            elif definition['type'] == "dropdown":
                input_field = NoWheelComboBox()
                input_field.setMinimumHeight(30)
                options = definition.get('options', [])
                input_field.addItems(options)
                input_field.setStyleSheet("""
                    QComboBox {
                        border: 1px solid #bdc3c7;
                        border-radius: 4px;
                        padding: 5px;
                        background-color: #ffffff;
                        font-size: 12px;
                    }
                    QComboBox:focus {
                        border-color: #3498db;
                    }
                """)
                grid_layout.addWidget(label, row, 0, 1, 2)
                grid_layout.addWidget(input_field, row + 1, 0, 1, 2)
            elif definition['type'] == "file":
                input_field = QLineEdit()
                input_field.setMinimumHeight(30)
                input_field.setStyleSheet("""
                    QLineEdit {
                        border: 1px solid #bdc3c7;
                        border-radius: 4px;
                        padding: 5px;
                        background-color: #ffffff;
                        font-size: 12px;
                    }
                    QLineEdit:focus {
                        border-color: #3498db;
                    }
                """)
                browse_button = QPushButton("Browse")
                browse_button.setStyleSheet("""
                    QPushButton {
                        background-color: #95a5a6;
                        color: white;
                        border: none;
                        border-radius: 4px;
                        padding: 5px 10px;
                        font-size: 10px;
                    }
                    QPushButton:hover {
                        background-color: #7f8c8d;
                    }
                """)
                browse_button.clicked.connect(lambda _, field=input_field: self.browse_file(field))
                screenshot_button = QPushButton("Screenshot")
                screenshot_button.setStyleSheet("""
                    QPushButton {
                        background-color: #D9D9D9;
                        color: white;
                        border: none;
                        border-radius: 4px;
                        padding: 5px 10px;
                        font-size: 10px;
                    }
                    QPushButton:hover {
                        background-color: #B0B0B0;
                    }
                """)
                screenshot_button.clicked.connect(lambda _, field=input_field: self.take_screenshot(field))
                grid_layout.addWidget(label, row, 0, 1, 4)
                grid_layout.addWidget(input_field, row + 1, 0, 1, 2)
                grid_layout.addWidget(browse_button, row + 1, 2)
                grid_layout.addWidget(screenshot_button, row + 1, 3)

            self.input_widgets[key] = input_field
            row += 2
            col = 0

        group_box.setLayout(grid_layout)
        parent_layout.addWidget(group_box)

    def browse_file(self, field):
        """
        Opens a file dialog to select an image file.

        Args:
            field (QLineEdit): The QLineEdit widget that will be populated with the
                               selected file path.
        """
        file_path, _ = QFileDialog.getOpenFileName(self.parent_app, "Select Image", "", "Image Files (*.png *.jpg *.jpeg *.gif *.bmp *.tiff *.tif *.webp *.jfif)")
        if file_path:
            field.setText(file_path)

    def take_screenshot(self, field):
        """
        Takes a screenshot of a specific area of the screen.

        Opens the ScreenshotSelector dialog, allows the user to select an area,
        and saves the selected image to a temporary file. The path to this
        temporary file is then inserted into the input field.

        Args:
            field (QLineEdit): The QLineEdit widget that will be populated with the
                               path to the temporary screenshot file.
        """
        dialog = ScreenshotSelector(parent=None)
        if dialog.exec_() == QDialog.Accepted:
            selected_image = dialog.get_selected_image()
            # Tetap gunakan delete=False, tapi kita akan mengelola penghapusannya secara manual
            temp_file = tempfile.NamedTemporaryFile(suffix='.png', delete=False)
            selected_image.save(temp_file.name, 'PNG')
            self.temp_files.append(temp_file.name)  # Tambahkan path ke daftar pelacakan
            field.setText(temp_file.name)

    def sync_related_fields(self, num_sonda):
        """
        Syncs related fields between TEMPERATURAS REGISTRADAS, ESTABILIZACIÓN TÉRMICA, and RESULTADOS.

        - Punto de Medición (PUNTO, MEDICI, PUNTODE) should be the same for each row.
        - Temperatura Medida (TEMP) and Temperatura final (TEMPE) should be the same.
        - Automatic calculation of Desviación (DESVI) from VALMIN and VALMAX.
        - Automatic filling of Resultado (RESULT) based on TEMP and LIMITE.
        """
        for i in range(1, num_sonda + 1):
            # Sync Punto de Medición
            punto_key = f"PUNTO{i}"
            medici_key = f"MEDICI{i}"
            puntode_key = f"PUNTODE{i}"

            if punto_key in self.input_widgets and medici_key in self.input_widgets and puntode_key in self.input_widgets:
                punto_widget = self.input_widgets[punto_key]
                medici_widget = self.input_widgets[medici_key]
                puntode_widget = self.input_widgets[puntode_key]

                # Connect signals to sync changes
                punto_widget.currentTextChanged.connect(lambda text, mw=medici_widget, pw=puntode_widget: self.sync_punto(text, mw, pw))
                medici_widget.currentTextChanged.connect(lambda text, pw=punto_widget, pw2=puntode_widget: self.sync_punto(text, pw, pw2))
                puntode_widget.currentTextChanged.connect(lambda text, pw=punto_widget, mw=medici_widget: self.sync_punto(text, pw, mw))

            # Sync Temperatura Medida and Temperatura final
            temp_key = f"TEMP{i}"
            tempe_key = f"TEMPE{i}"

            if temp_key in self.input_widgets and tempe_key in self.input_widgets:
                temp_widget = self.input_widgets[temp_key]
                tempe_widget = self.input_widgets[tempe_key]

                # Connect signals to sync changes
                temp_widget.textChanged.connect(lambda text, tw=tempe_widget: tw.setText(text))
                tempe_widget.textChanged.connect(lambda text, tw=temp_widget: tw.setText(text))

            # Automatic calculation of Desviación
            valmin_key = f"VALMIN{i}"
            valmax_key = f"VALMAX{i}"
            desvi_key = f"DESVI{i}"

            if valmin_key in self.input_widgets and valmax_key in self.input_widgets and desvi_key in self.input_widgets:
                valmin_widget = self.input_widgets[valmin_key]
                valmax_widget = self.input_widgets[valmax_key]
                desvi_widget = self.input_widgets[desvi_key]

                # Connect signals to calculate Desviación
                valmin_widget.textChanged.connect(lambda text, vw=valmin_widget, vw2=valmax_widget, dw=desvi_widget: self.calculate_desviacion(vw, vw2, dw))
                valmax_widget.textChanged.connect(lambda text, vw=valmin_widget, vw2=valmax_widget, dw=desvi_widget: self.calculate_desviacion(vw, vw2, dw))

            # Automatic filling of Resultado
            limite_key = f"LIMITE{i}"
            result_key = f"RESULT{i}"

            if temp_key in self.input_widgets and limite_key in self.input_widgets and result_key in self.input_widgets:
                temp_widget = self.input_widgets[temp_key]
                limite_widget = self.input_widgets[limite_key]
                result_widget = self.input_widgets[result_key]

                # Connect signals to fill Resultado
                temp_widget.textChanged.connect(lambda text, tw=temp_widget, lw=limite_widget, rw=result_widget: self.fill_resultado(tw, lw, rw))
                limite_widget.currentTextChanged.connect(lambda text, tw=temp_widget, lw=limite_widget, rw=result_widget: self.fill_resultado(tw, lw, rw))

    def sync_punto(self, text, widget1, widget2):
        """Sync the Punto de Medición dropdowns."""
        widget1.blockSignals(True)
        widget2.blockSignals(True)
        widget1.setCurrentText(text)
        widget2.setCurrentText(text)
        widget1.blockSignals(False)
        widget2.blockSignals(False)

    def calculate_desviacion(self, valmin_widget, valmax_widget, desvi_widget):
        """Calculate Desviación as VALMAX - VALMIN."""
        try:
            valmin_text = valmin_widget.text().strip()
            valmax_text = valmax_widget.text().strip()
            if valmin_text and valmax_text:
                valmin = float(valmin_text)
                valmax = float(valmax_text)
                desvi = valmax - valmin
                desvi_widget.setText(f"{desvi:.2f}")
            else:
                desvi_widget.setText("")
        except ValueError:
            desvi_widget.setText("")

    def fill_resultado(self, temp_widget, limite_widget, result_widget):
        """Fill Resultado as 'Pass' if TEMP <= LIMITE, 'Fail' if TEMP > LIMITE, else 'N/A'."""
        try:
            temp_text = temp_widget.text().strip()
            limite_text = limite_widget.currentText().strip()
            if temp_text and limite_text and limite_text != "N/A":
                temp = float(temp_text)
                limite = float(limite_text)
                if temp <= limite:
                    result_widget.setCurrentText("Pass")
                else:
                    result_widget.setCurrentText("Fail")
            else:
                result_widget.setCurrentText("N/A")
        except ValueError:
            result_widget.setCurrentText("N/A")

    def auto_fill_marca_tipo(self, equipo_text, marca_widget, tipo_widget, fecha_widget, num_equip):
        """Auto-fill 'Marca/Modelo' and 'Tipo/Aplicación' fields based on 'Equipo' selection. Also sync FECHA for SONDA TIPO T."""
        if equipo_text == "ALMEMO":
            marca_widget.setCurrentText("MA710")
            tipo_widget.setCurrentText("Registrador de Temperatura")
        elif equipo_text == "TERMOHIGRÓMETRO":
            marca_widget.setCurrentText("MA24702S")
            tipo_widget.setCurrentText("Medición Temperatura Ambiente")
        elif equipo_text == "CAMARA ENDURANCIA":
            marca_widget.setCurrentText("CET10/15312")
            tipo_widget.setCurrentText("Dycometal")
        elif equipo_text == "SONDA TIPO T":
            # Sync FECHA for all SONDA TIPO T rows
            if fecha_widget:
                fecha_value = fecha_widget.date().toString("dd/MM/yyyy")
                for j in range(1, num_equip + 1):
                    if f"EQUIPO{j}" in self.input_widgets and self.input_widgets[f"EQUIPO{j}"].currentText() == "SONDA TIPO T":
                        fecha_key = f"FECHA{j}"
                        if fecha_key in self.input_widgets:
                            self.input_widgets[fecha_key].setDate(fecha_widget.date())
        else:
            # Clear if not one of the auto-fill options
            marca_widget.setCurrentIndex(0)
            tipo_widget.setCurrentIndex(0)
