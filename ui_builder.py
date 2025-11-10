from PyQt5.QtWidgets import (
    QLabel, QLineEdit, QPushButton, QScrollArea, QGridLayout, QGroupBox, QFileDialog, QSpinBox, QTextEdit, QHBoxLayout, QWidget, QVBoxLayout, QDateEdit, QComboBox
)
from PyQt5.QtCore import Qt, QDate
from fields import FIELD_DEFINITIONS
from screenshot import ScreenshotSelector
import tempfile


class UIBuilder:
    def __init__(self, parent_app):
        self.parent_app = parent_app
        self.input_widgets = {}
        self.equipment_groups = []
        self.spin_boxes = {}

    def init_ui(self):
        """Membangun antarmuka pengguna."""
        main_layout = self.parent_app.layout()

        # --- Judul ---
        title = QLabel("Ingresar Datos Para el Anexo II")
        title.setAlignment(Qt.AlignCenter)
        title.setStyleSheet("font-size: 20px; font-weight: bold; margin-bottom: 10px; color: #2C3E50;")
        main_layout.addWidget(title)

        # --- Spin Boxes for Row Selection ---
        spin_layout = QHBoxLayout()
        spin_layout.addWidget(QLabel("EQUIPOS Y MÉTODOS UTILIZADOS (max 12):"))
        self.spin_equipment = QSpinBox()
        self.spin_equipment.setRange(1, 12)
        self.spin_equipment.setValue(12)
        self.spin_equipment.valueChanged.connect(self.rebuild_form)
        spin_layout.addWidget(self.spin_equipment)

        spin_layout.addWidget(QLabel("SONDA TOTAL (max 10):"))
        self.spin_sonda = QSpinBox()
        self.spin_sonda.setRange(1, 10)
        self.spin_sonda.setValue(10)
        self.spin_sonda.valueChanged.connect(self.rebuild_form)
        spin_layout.addWidget(self.spin_sonda)

        main_layout.addLayout(spin_layout)

        # --- Scroll Area untuk banyak input ---
        self.scroll = QScrollArea()
        self.scroll.setWidgetResizable(True)
        self.content_widget = QWidget()
        self.form_layout = QVBoxLayout(self.content_widget)
        self.form_layout.setSpacing(15)
        self.scroll.setWidget(self.content_widget)
        main_layout.addWidget(self.scroll)

        # --- Tombol Generate ---
        self.generate_button = QPushButton("GENERAR DOCUMENTO DE WORD (.docx)")
        self.generate_button.setStyleSheet(
            "background-color: #3498DB; color: white; padding: 12px; border-radius: 8px; font-weight: bold;"
        )
        self.generate_button.clicked.connect(self.parent_app.generate_document)
        main_layout.addWidget(self.generate_button)

        # --- Informasi Template ---
        info = QLabel(
            f"**Plantilla utilizada:** '{self.parent_app.template_filename}'\n"
            f"Asegúrate de que este archivo esté en la carpeta: '{self.parent_app.templates_dir}'"
        )
        info.setStyleSheet("font-size: 10px; color: gray; margin-top: 5px;")
        main_layout.addWidget(info)

        self.rebuild_form()

    def rebuild_form(self):
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
            "TEXT6", "TEXT7", "TEXT8"
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
        self.create_input_group(self.form_layout, "NOTA", [
            "TEXT13"
        ])

        # 3.1. GRÁFICA GENERADA
        self.create_input_group(self.form_layout, "3.1. GRÁFICA GENERADA", [
            "IMAGE1", "TITLE1", "DESC1", "IMAGE2", "DESC2"
        ])

        # 4. ESTABILIZACIÓN TÉRMICA
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

        # 6. CONCLUSIONES DEL LABORATORIO
        title_label = QLabel("6. CONCLUSIONES DEL LABORATORIO")
        title_label.setStyleSheet("font-size: 18px; font-weight: bold; margin-bottom: 10px;")
        self.form_layout.addWidget(title_label)
        self.create_input_group(self.form_layout, "CONCLUSIONES", [
            "TEXT15"
        ])

        # 7. FOTOGRAFIAS
        title_label = QLabel("7. FOTOGRAFIAS")
        title_label.setStyleSheet("font-size: 18px; font-weight: bold; margin-bottom: 10px;")
        self.form_layout.addWidget(title_label)
        for i in range(3, 3 + num_sonda):  # IMAGE3 to IMAGE{2+num_sonda}, TITLE3 to TITLE{2+num_sonda}
            self.create_input_group(self.form_layout, f"Fotografía {i-2}", [
                f"IMAGE{i}"
            ])
            self.create_input_group(self.form_layout, f"Titulo {i-2}", [
                f"TITLE{i}"
            ])

        # Restore saved values
        for key, value in current_values.items():
            if key in self.input_widgets:
                widget = self.input_widgets[key]
                if isinstance(widget, QLineEdit):
                    widget.setText(value)
                    # Set default for OBSER fields if empty
                    if key.startswith("OBSER") and not value.strip():
                        widget.setText("-")
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

    def create_input_group(self, parent_layout, title, keys):
        """Membuat group box untuk input yang terorganisir."""
        group_box = QGroupBox(title)
        group_box.setStyleSheet("font-weight: bold; margin-top: 10px;")
        grid_layout = QGridLayout()
        grid_layout.setSpacing(10)
        
        row = 0
        col = 0
        
        for key in keys:
            definition = FIELD_DEFINITIONS[key]
            
            label = QLabel(f"{definition['label']}:")
            
            if definition['type'] == "text":
                if key in ["TEXT1", "TEXT4", "TEXT2", "TEXT5", "TEXT3", "TEXT6", "TEXT7", "TEXT8", "TEXT9", "TEXT10", "TEXT11", "TEXT13", "TEXT15"]:
                    input_field = QTextEdit()
                    input_field.setMinimumHeight(60)
                    grid_layout.addWidget(label, row, 0, 1, 2)
                    grid_layout.addWidget(input_field, row + 1, 0, 1, 2)
                    row += 2
                    col = 0
                else:
                    input_field = QLineEdit()
                    input_field.setMinimumHeight(30)
                    grid_layout.addWidget(label, row, col)
                    grid_layout.addWidget(input_field, row + 1, col)
                    col = 1 - col
                    if col == 0:
                        row += 2
            elif definition['type'] == "date":
                input_field = QDateEdit()
                input_field.setCalendarPopup(True)
                input_field.setMinimumHeight(30)
                input_field.setDate(QDate.currentDate())  # Set default to today's date
                grid_layout.addWidget(label, row, col)
                grid_layout.addWidget(input_field, row + 1, col)
                col = 1 - col
                if col == 0:
                    row += 2
            elif definition['type'] == "dropdown":
                input_field = QComboBox()
                input_field.setMinimumHeight(30)
                options = definition.get('options', [])
                input_field.addItems(options)
                grid_layout.addWidget(label, row, col)
                grid_layout.addWidget(input_field, row + 1, col)
                col = 1 - col
                if col == 0:
                    row += 2
            elif definition['type'] == "file":
                input_field = QLineEdit()
                input_field.setMinimumHeight(30)
                browse_button = QPushButton("Browse")
                browse_button.clicked.connect(lambda _, field=input_field: self.browse_file(field))
                screenshot_button = QPushButton("Screenshot")
                screenshot_button.clicked.connect(lambda _, field=input_field: self.take_screenshot(field))
                grid_layout.addWidget(label, row, 0, 1, 4)
                grid_layout.addWidget(input_field, row + 1, 0, 1, 2)
                grid_layout.addWidget(browse_button, row + 1, 2)
                grid_layout.addWidget(screenshot_button, row + 1, 3)
                row += 2
                col = 0
            
            self.input_widgets[key] = input_field

        group_box.setLayout(grid_layout)
        parent_layout.addWidget(group_box)

    def browse_file(self, field):
        """Browse file untuk input gambar."""
        file_path, _ = QFileDialog.getOpenFileName(self.parent_app, "Select Image", "", "Image Files (*.png *.jpg *.jpeg *.gif *.bmp *.tiff *.tif *.webp *.jfif)")
        if file_path:
            field.setText(file_path)

    def take_screenshot(self, field):
        """Take screenshot and allow area selection like snipping tool."""
        dialog = ScreenshotSelector(parent=self.parent_app)
        if dialog.exec_() == QDialog.Accepted:
            selected_image = dialog.get_selected_image()
            temp_file = tempfile.NamedTemporaryFile(suffix='.png', delete=False)
            selected_image.save(temp_file.name, 'PNG')
            field.setText(temp_file.name)
