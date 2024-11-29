# Gmail Draft Generator for Google Sheets

## Overview
This Google Apps Script project allows you to generate Gmail drafts directly from a Google Sheets spreadsheet with a custom HTML interface.

## Features
- Create Gmail drafts with dynamic content from spreadsheet ranges
- Custom HTML interface for selecting data ranges
- Supports multiple sheets for configuring email components
- Automatic date stamping
- Styled HTML email generation

## Files
- `Código.gs`: Main Google Apps Script containing core functionality
- `gui.html`: User interface for draft generation

## Functions

### `abrirFormulario()`
Opens a modal dialog with the HTML interface for generating Gmail drafts.

### `generarBorradorGmail(rangoNombre)`
Generates a Gmail draft with the following steps:
- Reads data from specified named range
- Retrieves email components from different sheets
- Creates an HTML-formatted email with:
  - Pre-message text
  - Tabular data
  - Post-message text
- Supports custom styling
- Handles errors gracefully

## Required Sheets
Your Google Sheets document should have the following named sheets:
- `Destinatarios`: Email recipients (Column A)
- `cc`: Carbon copy recipients (Column A)
- `Asunto`: Email subject lines (Column A)
- `cuerpo_del_mensaje_pre_tabla`: Text before the data table
- `cuerpo_del_mensaje_post_tabla`: Text after the data table

## Usage
1. Open your Google Sheets document
2. Go to Tools > Script editor
3. Paste the `Código.gs` script
4. Create a new HTML file named `gui` and paste the HTML content
5. Save and reload the spreadsheet
6. Use the "Crear Borrador Gmail" menu item or custom function

## Error Handling
- Provides detailed error messages
- Logs errors to console
- Returns user-friendly error notifications

## Requirements
- Google Workspace account
- Google Sheets
- Google Apps Script enabled

## Customization
- Modify HTML/CSS in `gui.html` to change interface styling
- Adjust email formatting in `generarBorradorGmail()` function

## License
Open-source project. Feel free to modify and distribute.

## Author
[Your Name]

## Version
1.0.0
