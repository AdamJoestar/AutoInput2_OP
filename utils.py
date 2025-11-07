from PyQt5.QtWidgets import QLineEdit, QTextEdit

# Utility functions for the application

def validate_required_fields(input_widgets, field_definitions):
    """Validate that all required fields are filled."""
    for key in input_widgets:
        definition = field_definitions[key]
        input_widget = input_widgets[key]
        if isinstance(input_widget, QLineEdit):
            value = input_widget.text().strip()
        elif isinstance(input_widget, QTextEdit):
            value = input_widget.toPlainText().strip()
        else:
            continue

        if key in ["TEXT1", "TEXT4", "TEXT2", "TEXT5", "TEXT3", "TEXT6", "TEXT7", "TEXT8", "TEXT9", "TEXT10", "TEXT11", "TEXT12", "TEXT13", "TEXT14", "TEXT15"] and not value:
            return False, definition['label']
    return True, None

# Note: Import QLineEdit and QTextEdit if needed, but since this is a utility, we can assume they are imported where used.
