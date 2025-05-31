# Report Generator

**ReportGen** is a Flask-based web application designed to automatically generate dynamic reports from user-uploaded Word templates and data sources. This tool is ideal for automating periodic report generation, especially in business, academic, or administrative environments.

## Features

* Upload `.docx` templates containing placeholder variables (e.g., `{{name}}`, `{{total_sales}}`)
* Upload CSV or JSON data to replace those placeholders
* Render live HTML previews before downloading
* Supports inline formula calculation using JavaScript Math engine
* Preview UI includes tooltips for variable descriptions
* Integrates Google Picker API for selecting documents from Google Drive
* Auto-save mechanism for modified HTML previews
* Designed for deployment on Heroku (Procfile included)

## Directory Structure

```
reportGen/
├── app.py                  # Flask application entry point
├── requirements.txt        # Python dependencies
├── Procfile                # Heroku process file
├── README.md               # This documentation
├── templates/              # HTML templates (Jinja2)
│   ├── home.html
│   ├── index.html
│   └── preview.html
├── static/                 # Static assets
│   ├── app-logo-no-text.svg
│   ├── autocomplete.min.js
│   ├── favicon.ico
│   ├── main.js
│   ├── math.min.js
│   ├── picker.js
│   └── style.css
├── fonts/                  # Custom font used in rendering
│   └── cwTeXQYuan-Medium.ttf
├── uploads/                # Temporary storage for uploaded Word/CSV/JSON files
├── generated/              # Stores generated output documents (e.g., rendered .docx)
├── report/                 # LaTeX Beamer presentation source and assets
│   ├── bm_reportGen.tex          # Main TeX source for slides
│   ├── bm_reportGen.pdf          # Compiled Beamer PDF presentation
│   ├── beamer*.sty               # Custom theme files for styling slides
│   ├── convert.pptx              # Converted version of the slides in PowerPoint format
│   ├── figures/                  # Diagrams and visual illustrations
│   └── svg-inkscape/            # Inkscape-exported LaTeX-compatible SVGs
└── .git/                   # Git repo (optional; can be ignored)
```

## Slide Presentation (in `report/`)

The `report/` folder contains a professional slide deck (`bm_reportGen.pdf`) created with LaTeX Beamer. It includes:

* **Custom themes** (`beamerthemefocus.sty`, etc.) for modern visuals
* **Figures and SVG diagrams** (e.g., report logic, UI, workflows)
* **Inkscape support**: `.svg_tex` exports ensure LaTeX-compatible vector graphics
* **PowerPoint version** (`convert.pptx`) for broader accessibility

This presentation is intended for demos or academic/professional talks explaining how ReportGen works.

## Installation

### 1. Clone the Repository

```bash
git clone https://github.com/xiaolong70701/reportGen.git
cd reportGen
```

### 2. Create a Virtual Environment

```bash
python3 -m venv venv
source venv/bin/activate   # On Windows: venv\Scripts\activate
```

### 3. Install Dependencies

```bash
pip install -r requirements.txt
```

### 4. Run the Application Locally

```bash
python app.py
```

Then visit `http://127.0.0.1:5000/` in your browser.

## Deployment

This project is ready to be deployed on **Heroku**.

```bash
heroku create your-app-name
git push heroku main
heroku buildpacks:add heroku/python
heroku open
```

## How It Works

* Word templates are parsed for `{{variable_name}}` style placeholders.
* Users upload data files (CSV or JSON).
* A dictionary is generated from the uploaded data to match each variable.
* Formulas inside double curly braces like `{{price * quantity}}` are evaluated using a JavaScript math engine.
* HTML preview is rendered and saved with JavaScript enhancements for interactivity.
* Modified previews can be saved and exported as `.docx`.

## Notes

* Requires Python 3.7 or above
* JavaScript libraries used: `math.min.js`, `autocomplete.min.js`, `picker.js`
* Preview rendering uses custom CSS for print-friendly formatting
* SVG and custom fonts are embedded for cross-platform styling consistency

## Future Improvements

* Export directly to PDF
* Add user authentication
* History tracking for uploaded reports
* Support for more complex templating logic (loops, conditionals)

## License

This project is open-source and available under the MIT License.