
# Report Generator

**ReportGen** is a Flask-based web application designed to automatically generate dynamic reports from user-uploaded Word templates and data sources. This tool is ideal for automating periodic report generation, especially in business, academic, or administrative environments.

## Features

- Upload `.docx` templates containing placeholder variables (e.g., `{{name}}`, `{{total_sales}}`)
- Upload CSV or JSON data to replace those placeholders
- Render live HTML previews before downloading
- Supports inline formula calculation using JavaScript Math engine
- Preview UI includes tooltips for variable descriptions
- Integrates Google Picker API for selecting documents from Google Drive
- Auto-save mechanism for modified HTML previews
- Designed for deployment on Heroku (Procfile included)

## Directory Structure

```
reportGen/
├── app.py                  # Flask application entry point
├── requirements.txt        # Python dependencies
├── Procfile                # Heroku process file
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
└── .git/                   # Embedded git repo (can be ignored or removed)
```

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

1. Create a Heroku app:
    ```bash
    heroku create your-app-name
    ```

2. Push code to Heroku:
    ```bash
    git push heroku main
    ```

3. Ensure buildpacks are set:
    ```bash
    heroku buildpacks:add heroku/python
    ```

4. Open in browser:
    ```bash
    heroku open
    ```

## How It Works

- Word templates are parsed for `{{variable_name}}` style placeholders.
- Users upload data files (CSV or JSON).
- A dictionary is generated from the uploaded data to match each variable.
- Formulas inside double curly braces like `{{price * quantity}}` are evaluated using a JavaScript math engine.
- HTML preview is rendered and saved with JavaScript enhancements for interactivity.
- Modified previews can be saved and exported as `.docx`.

## Notes

- Requires Python 3.7 or above
- JavaScript libraries used: `math.min.js`, `autocomplete.min.js`, `picker.js`
- Preview rendering uses custom CSS for print-friendly formatting
- SVG and custom fonts are embedded for cross-platform styling consistency

## Future Improvements

- Export directly to PDF
- Add user authentication
- History tracking for uploaded reports
- Support for more complex templating logic (loops, conditionals)

## License

This project is open-source and available under the MIT License.
