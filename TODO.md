# TODO: Refactoring main2.py into Modular Structure

## Overview
Refactor the monolithic `main2.py` file into smaller, modular files for better maintainability, readability, and testability.

## Steps
- [ ] Create `config.py`: Extract configuration constants (TEMPLATES_DIR, TEMPLATE_FILENAME, etc.)
- [ ] Create `fields.py`: Extract FIELD_DEFINITIONS dictionary and related field logic
- [ ] Create `screenshot.py`: Extract ScreenshotSelector class
- [ ] Create `utils.py`: Extract utility functions (if any, e.g., file helpers)
- [ ] Create `ui_builder.py`: Extract UI building methods (init_ui, rebuild_form, create_input_group)
- [ ] Create `document_processor.py`: Extract document processing logic (generate_document, replace methods, etc.)
- [ ] Create `main_app.py`: Create main application entry point that imports and combines all modules
- [ ] Update `main2.py`: Rename or replace with `main_app.py` and ensure all imports work
- [ ] Test the refactored application: Run the app and generate a document to verify functionality
- [ ] Debug and fix any errors: Address import issues, path problems, or runtime errors
- [ ] Final cleanup: Remove old code, ensure no dependencies are broken

## Notes
- Ensure all imports are updated correctly in each new file.
- Maintain the same functionality; no feature changes.
- After each step, run a quick test if possible to catch issues early.
