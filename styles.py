LIGHT_THEME = """
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
    QLineEdit, QTextEdit, QDateEdit, QComboBox, QSpinBox {
        border: 1px solid #ccc;
        border-radius: 4px;
        padding: 5px;
        background-color: #fff;
    }
    QLineEdit:focus, QTextEdit:focus, QDateEdit:focus, QComboBox:focus, QSpinBox:focus {
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
"""

DARK_THEME = """
    QWidget {
        font-size: 14px;
        font-family: 'Segoe UI', Arial, sans-serif;
        background-color: #2b2b2b;
        color: #f0f0f0;
    }
    QGroupBox {
        font-weight: bold;
        border: 2px solid #555;
        border-radius: 8px;
        margin-top: 10px;
        background-color: #3c3c3c;
        padding: 10px;
    }
    QGroupBox::title {
        subcontrol-origin: margin;
        left: 10px;
        padding: 0 5px 0 5px;
        color: #f0f0f0;
        font-size: 16px;
    }
    QLabel {
        color: #f0f0f0;
    }
    QLineEdit, QTextEdit, QDateEdit, QComboBox, QSpinBox {
        border: 1px solid #555;
        border-radius: 4px;
        padding: 5px;
        background-color: #3c3c3c;
        color: #f0f0f0;
    }
    QLineEdit:focus, QTextEdit:focus, QDateEdit:focus, QComboBox:focus, QSpinBox:focus {
        border-color: #0078d7;
    }
    QPushButton {
        border-radius: 6px;
        padding: 8px 16px;
        font-weight: bold;
        background-color: #555;
        color: #f0f0f0;
        border: 1px solid #666;
    }
    QPushButton:hover {
        background-color: #666;
    }
"""