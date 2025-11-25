from PyQt5.QtWidgets import (
    QLabel, QLineEdit, QPushButton, QScrollArea, QGridLayout, QGroupBox, QFileDialog, QSpinBox, QTextEdit, QHBoxLayout, QWidget, QVBoxLayout, QDateEdit, QComboBox, QDialog, QDateTimeEdit
, QTabWidget)
from PyQt5.QtCore import Qt, QDate, QDateTime
from PyQt5.QtGui import QPixmap
from fields import FIELD_DEFINITIONS, EQUIPO_AUTOFILL_DATA, CODIGO_OPTIONS_BY_EQUIPO
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
        self.num_sonda = 10  # Default number of sensors

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
        # Widget untuk spinbox equipment, akan ditambahkan di dalam scroll area nanti
        spin_layout = QHBoxLayout()
        label1 = QLabel("EQUIPOS Y MÉTODOS UTILIZADOS (max 12):")
        label1.setStyleSheet("font-weight: bold; color: #34495e; padding: 5px;")
        spin_layout.addWidget(label1)
        self.spin_equipment = NoWheelSpinBox()
        self.spin_equipment.setRange(1, 12)
        self.spin_equipment.setValue(12)
        self.spin_equipment.setValue(12) # Nilai default
        self.spin_equipment.setStyleSheet("QSpinBox { border: 1px solid #bdc3c7; border-radius: 4px; padding: 5px; background-color: #ecf0f1; }")
        self.spin_equipment.valueChanged.connect(self.rebuild_form)
        spin_layout.addWidget(self.spin_equipment)



        self.equipment_spin_widget = QWidget()
        self.equipment_spin_widget.setLayout(spin_layout)
        
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
        
        # --- Tab Widget untuk Navigasi Formulir ---
        self.tab_widget = QTabWidget()
        self.tab_widget.setStyleSheet("QTabBar::tab { padding: 10px 15px; }") # Beri padding pada judul tab
        main_layout.addWidget(self.tab_widget)

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
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #0078D4, stop:1 #005A9E);
                color: white;
                padding: 12px;
                border-radius: 8px;
                font-weight: bold;
                border: none;
            }
            QPushButton:hover {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #0084E6, stop:1 #006AB1);
            }
            QPushButton:pressed {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #006AB1, stop:1 #004C87);
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
        print(f"Rebuilding form with num_sonda = {self.num_sonda}")
        self.rebuilding = True

        # Simpan indeks tab yang sedang aktif sebelum membangun ulang
        saved_tab_index = self.tab_widget.currentIndex()

        # Disconnect signals to prevent recursive calls during rebuild
        if hasattr(self, 'spin_equipment'):
            try:
                self.spin_equipment.valueChanged.disconnect(self.rebuild_form)
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
        saved_additional = 0
        if hasattr(self, 'spin_additional_photos') and self.spin_additional_photos:
            saved_additional = self.spin_additional_photos.value()

        # Hapus semua tab yang ada
        self.tab_widget.clear()
        self.input_widgets = {}

        # --- Tab 1: Header & Solicitante ---
        header_solicitante_layout = self._create_tab_and_get_layout("0. Encabezado y Solicitante")
        self.create_input_group(header_solicitante_layout, "Encabezado - Información del documento", [
            "NO_TEST", "REV", "DATE"
        ])
        self.create_input_group(header_solicitante_layout, "0. INFORMACIÓN DEL SOLICITANTE DEL ENSAYO", [
            "TEXT1", "TEXT4", "TEXT2", "TEXT5", "TEXT3"
        ])

        # --- Tab 2: Producto & Condiciones ---
        producto_layout = self._create_tab_and_get_layout("1. Producto y Condiciones")
        self.create_input_group(producto_layout, "1. INFORMACIÓN GENERAL DEL PRODUCTO", [
            "TEXT6", "TEXT12", "TEXT7", "TEXT8", "TEXT13"
        ])
        self.create_input_group(producto_layout, "1.1. CONDICIONES DEL ENSAYO", [
            "TEXT9", "TEXT10", "TEXT11"
        ])

        # --- Tab 3: Equipos ---
        equipos_layout = self._create_tab_and_get_layout("2. Equipos Utilizados")
        equipos_layout.addWidget(self.equipment_spin_widget) # Tambahkan spinbox di sini
        num_equip = self.spin_equipment.value()
        for i in range(1, num_equip + 1):
            self.create_input_group(equipos_layout, f"Equipo {i}", [
                f"EQUIPO{i}", f"MARCA{i}", f"TIPO{i}", f"FECHA{i}", f"OBSER{i}"
            ])

        # Auto-fill marca and tipo for all rows based on equipment selection
        # Also sync FECHA for SONDA TIPO T
        for i in range(1, num_equip + 1):
            equipo_key = f"EQUIPO{i}"
            marca_key = f"MARCA{i}"
            tipo_key = f"TIPO{i}"
            fecha_key = f"FECHA{i}"
            codigo_key = f"OBSER{i}"
            if equipo_key in self.input_widgets:
                equipo_widget = self.input_widgets[equipo_key]
                marca_widget = self.input_widgets.get(marca_key)
                tipo_widget = self.input_widgets.get(tipo_key)
                fecha_widget = self.input_widgets.get(fecha_key)
                codigo_widget = self.input_widgets.get(codigo_key)
                if marca_widget and tipo_widget:
                    equipo_widget.currentTextChanged.connect(lambda text, mw=marca_widget, tw=tipo_widget, fw=fecha_widget, cw=codigo_widget, num=num_equip: self.auto_fill_marca_tipo(text, mw, tw, fw, cw, num))

        # --- Tab 4: Temperaturas & Gráfica ---
        temperaturas_layout = self._create_tab_and_get_layout("3. Temperaturas y Gráfica")
        for i in range(1, self.num_sonda + 1):
            self.create_input_group(temperaturas_layout, f"Temperatura Punto {i}", [
                f"PUNTO{i}", f"LIMITE{i}", f"TEMP{i}"
            ])
        self.create_input_group(temperaturas_layout, "3.1. GRÁFICA GENERADA", [
            "IMAGE1", "TITLE1", "IMAGE2"
        ])

        # --- Tab 5: Estabilización Térmica ---
        estabilizacion_layout = self._create_tab_and_get_layout("4. Estabilización Térmica")
        for i in range(1, self.num_sonda + 1):
            self.create_input_group(estabilizacion_layout, f"Estabilización Punto {i}", [
                f"MEDICI{i}", f"VALMIN{i}", f"VALMAX{i}", f"DESVI{i}"
            ])

        # --- Tab 6: Resultados ---
        resultados_layout = self._create_tab_and_get_layout("5. Resultados")
        for i in range(1, self.num_sonda + 1):
            self.create_input_group(resultados_layout, f"Resultado Punto {i}", [
                f"PUNTODE{i}", f"TEMPE{i}", f"RESULT{i}"
            ])

        # --- Tab 7: Fotografías ---
        fotografias_layout = self._create_tab_and_get_layout("6. Fotografías")
        for i in range(3, 3 + self.num_sonda):  # IMAGE3 to IMAGE{2+num_sonda}, TITLE3 to TITLE{2+num_sonda}
            self.create_input_group(fotografias_layout, f"Fotografía {i-2}", [
                f"IMAGE{i}"
            ])
            self.create_input_group(fotografias_layout, f"Titulo {i-2}", [
                f"TITLE{i}"
            ])

        # --- Tab 8: Montaje Final ---
        montaje_layout = self._create_tab_and_get_layout("7. Montaje Final")

        # Add the spinbox for additional photos
        montaje_layout.addWidget(self.additional_photos_widget)
        self.spin_additional_photos.blockSignals(True)
        self.spin_additional_photos.setValue(saved_additional)
        self.spin_additional_photos.blockSignals(False)

        num_additional = self.spin_additional_photos.value()
        for i in range(13, 13 + num_additional):  # IMAGE13 to IMAGE{12+num_additional}, TITLE13 to TITLE{12+num_additional}
            self.create_input_group(montaje_layout, f"Montaje {i-12}", [
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
        self.sync_related_fields(self.num_sonda)

        # Re-apply highlights from Excel load
        self._apply_highlights()

        # Pulihkan tab yang sebelumnya aktif
        if saved_tab_index >= 0 and saved_tab_index < self.tab_widget.count():
            self.tab_widget.setCurrentIndex(saved_tab_index)

        # Reconnect signals after rebuild
        if hasattr(self, 'spin_equipment'):
            self.spin_equipment.valueChanged.connect(self.rebuild_form)
        if hasattr(self, 'spin_additional_photos'):
            self.spin_additional_photos.valueChanged.connect(self.rebuild_form)

        self.rebuilding = False

    def _apply_highlights(self):
        """
        Menerapkan kembali sorotan kuning ke field yang diisi dari Excel
        setelah form dibangun ulang.
        """
        highlight_style = "background-color: #fff9c4;"
        for key in self.parent_app.highlighted_fields:
            if key in self.input_widgets:
                self.input_widgets[key].setStyleSheet(highlight_style)

    def _create_tab_and_get_layout(self, title):
        """
        Membuat tab baru di QTabWidget, menempatkan QScrollArea di dalamnya,
        dan mengembalikan layout vertikal untuk konten tab tersebut.
        """
        tab = QWidget()
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setStyleSheet("""
            QScrollArea { border: none; background-color: #fafafa; }
            QScrollArea > QWidget > QWidget { background-color: #fafafa; }
        """)

        content_widget = QWidget()
        form_layout = QVBoxLayout(content_widget)
        form_layout.setSpacing(15)
        form_layout.setContentsMargins(10, 10, 10, 10)

        scroll_area.setWidget(content_widget)
        tab_layout = QVBoxLayout(tab)
        tab_layout.addWidget(scroll_area)
        self.tab_widget.addTab(tab, title)
        return form_layout

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
                input_field = NoWheelDateEdit()
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

    def auto_fill_marca_tipo(self, equipo_text, marca_widget, tipo_widget, fecha_widget, codigo_widget, num_equip):
        """Auto-fill fields based on 'Equipo' selection. Also sync FECHA for SONDA TIPO T."""
        # --- Filter Código Dropdown ---
        if codigo_widget:
            current_code = codigo_widget.currentText()
            codigo_widget.blockSignals(True)
            codigo_widget.clear()
            options = CODIGO_OPTIONS_BY_EQUIPO.get(equipo_text, [])
            codigo_widget.addItems(options)
            # Coba atur kembali ke nilai sebelumnya jika masih ada di opsi baru
            if current_code in options:
                codigo_widget.setCurrentText(current_code)
            codigo_widget.blockSignals(False)

        # --- Auto-fill other fields ---
        if equipo_text == "SONDA TIPO T": # Kasus khusus untuk SONDA TIPO T
            # Untuk SONDA TIPO T, hanya sinkronkan tanggal dan jangan sentuh field lain.
            if fecha_widget:
                # Pertama, atur tanggal baris saat ini dari data auto-fill jika ada
                if equipo_text in EQUIPO_AUTOFILL_DATA and "fecha" in EQUIPO_AUTOFILL_DATA[equipo_text]:
                    date = QDate.fromString(EQUIPO_AUTOFILL_DATA[equipo_text]["fecha"], "dd/MM/yyyy")
                    if date.isValid():
                        fecha_widget.setDate(date)
                
                # Kemudian, sinkronkan tanggal ini ke semua baris SONDA TIPO T lainnya
                for j in range(1, num_equip + 1):
                    if f"EQUIPO{j}" in self.input_widgets and self.input_widgets[f"EQUIPO{j}"].currentText() == "SONDA TIPO T":
                        fecha_key = f"FECHA{j}"
                        if fecha_key in self.input_widgets:
                            self.input_widgets[fecha_key].setDate(fecha_widget.date())
        elif equipo_text in EQUIPO_AUTOFILL_DATA: # Kasus untuk equipo lain di auto-fill
            # Untuk equipo lain, isi semua field seperti biasa.
            data = EQUIPO_AUTOFILL_DATA[equipo_text]
            if "marca" in data:
                marca_widget.setCurrentText(data["marca"])
            if "tipo" in data:
                tipo_widget.setCurrentText(data["tipo"])
            if fecha_widget and "fecha" in data:
                date = QDate.fromString(data["fecha"], "dd/MM/yyyy")
                if date.isValid():
                    fecha_widget.setDate(date)
        else: # Jika equipo tidak ada di daftar auto-fill
            # Hapus isian jika bukan salah satu opsi auto-fill
            marca_widget.setCurrentIndex(0)
            tipo_widget.setCurrentIndex(0)
