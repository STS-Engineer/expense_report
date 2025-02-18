from flask import Flask, request, send_file, render_template_string, jsonify
import os
from PyPDF2 import PdfReader
import pandas as pd
import requests
import fitz  # PyMuPDF for extracting images
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as ExcelImage
from openpyxl.styles import Font
app = Flask(__name__)

UPLOAD_FOLDER = 'uploads'
OUTPUT_FOLDER = 'outputs'

# Live currency API
FIXER_API_URL = "http://data.fixer.io/api/latest"
FIXER_API_KEY = "653aca7bac0ce92affcdcb0116ecbc0a"


@app.route('/')
def home():
    return '''
    <!doctype html>
    <html lang="en">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>PDF to Excel Converter</title>
        <style>
            body {
                font-family: Arial, sans-serif;
                background-color: #f4f4f9;
                margin: 0;
                padding: 20px;
                display: flex;
                justify-content: center;
                align-items: center;
                height: 100vh;
            }
            .container {
                text-align: center;
                background: white;
                padding: 20px 40px;
                border-radius: 8px;
                box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
                width: 100%;
                max-width: 500px;
            }
            form {
                margin-top: 20px;
            }
            input[type="file"] {
                padding: 10px;
                margin: 10px 0;
                width: 100%;
            }
            button {
                background-color: #007bff;
                color: white;
                border: none;
                padding: 10px 20px;
                cursor: pointer;
                border-radius: 4px;
                font-size: 16px;
            }
            button:hover {
                background-color: #0056b3;
            }
        </style>
    </head>
    <body>
        <div class="container">
            <h1>PDF to Excel Converter</h1>
            <form action="/upload" method="post" enctype="multipart/form-data">
                <input type="file" name="file" accept="application/pdf" required>
                <button type="submit">Convert to Excel</button>
            </form>
        </div>
    </body>
    </html>
    '''


@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return jsonify({"error": "No file uploaded"}), 400

    file = request.files['file']
    if file.filename == '':
        return jsonify({"error": "No selected file"}), 400

    file_path = os.path.join(UPLOAD_FOLDER, file.filename)
    file.save(file_path)

    try:
        # Extract data and convert to Excel (your existing functionality)
        extracted_data, excel_file_path = convert_pdf_to_excel(file_path, file.filename.replace('.pdf', '.xlsx'))

        # Extract images and add them to the Excel file
        image_files = extract_images_from_pdf(file_path)
        if image_files:
            add_images_to_excel(image_files, excel_file_path)
            print(f"Extracted and added {len(image_files)} images to Excel.")

        return render_template_string(generate_response_html(extracted_data, excel_file_path))
    except Exception as e:
        return jsonify({"error": str(e)}), 500

IMAGE_FOLDER = os.path.join(OUTPUT_FOLDER, 'images')
os.makedirs(IMAGE_FOLDER, exist_ok=True)
def extract_images_from_pdf(pdf_path):
    """Extract images from a PDF and save them as files."""
    doc = fitz.open(pdf_path)
    image_paths = []

    for page_number in range(len(doc)):
        page = doc[page_number]
        images = page.get_images(full=True)

        for img_index, img in enumerate(images):
            xref = img[0]
            base_image = doc.extract_image(xref)
            image_bytes = base_image["image"]
            image_ext = base_image["ext"]

            output_path = os.path.join(IMAGE_FOLDER, f'page-{page_number + 1}_image-{img_index + 1}.{image_ext}')
            with open(output_path, "wb") as img_file:
                img_file.write(image_bytes)
            image_paths.append(output_path)

    return image_paths

def add_images_to_excel(image_files, excel_file_path):
    """Embed images into separate sheets in an existing Excel file without duplication."""
    workbook = load_workbook(excel_file_path)

    added_images = set()  # Track added images to prevent duplicates

    for index, image_path in enumerate(image_files, start=1):
        if image_path in added_images:
            continue  # Skip duplicate images

        img = ExcelImage(image_path)
        img.width, img.height = 300, 300  # Resize images for better visibility

        sheet_name = f"Image {index}"
        if sheet_name not in workbook.sheetnames:  # Ensure unique sheet names
            image_sheet = workbook.create_sheet(title=sheet_name)

            # Set header
            image_sheet.append(["Extracted Image"])
            image_sheet["A1"].font = Font(bold=True)

            # Insert image
            cell = "A3"  # Leave space for the header
            image_sheet.add_image(img, cell)

            # Adjust column width and row height
            image_sheet.column_dimensions["A"].width = 50
            image_sheet.row_dimensions[3].height = 200

            added_images.add(image_path)  # Mark image as added

    workbook.save(excel_file_path)

def convert_pdf_to_excel(pdf_path, output_filename):
    """Extracts structured data from a PDF and converts it to an Excel file."""
    import re
    reader = PdfReader(pdf_path)
    extracted_data = {"Libelle": None, "Department": None, "Object": None, "Table": []}
    
    # Comptes comptables mapping
    compte_comptable_mapping = {
        "train": 625100,
        "avion": 625100,
        "parking": 625100,
        "taxi": 625100,
        "carburant": 625110,
        "peages": 625130,
        "entretien vehicule": 625140,
        "hotel": 625200,
        "repas restaurant": 625300,
        "reception": 625700,
        "affranchissement": 626000,
        "telephonie": 626100,
        "achats divers": 606300
    }

    # Fetch conversion rates from Fixer.io
    try:
        response = requests.get(FIXER_API_URL, params={"access_key": FIXER_API_KEY})
        response.raise_for_status()
        conversion_data = response.json()
        rates = conversion_data.get("rates", {})
    except Exception as e:
        print(f"Failed to fetch conversion rates: {e}")
        rates = {}

    for page in reader.pages:
        text = page.extract_text()
        if not text:
            continue
        lines = text.split('\n')

        headers_found = False
        for line in lines:
            # Extract Name and Department
            if "NAME" in line and "DEPARTMENT" in line:
                match = re.search(r"NAME\s*(.*?)\s*DEPARTMENT\s*(.*)", line, re.IGNORECASE)
                if match:
                    extracted_data["Libelle"] = match.group(1).strip()
                    extracted_data["Department"] = match.group(2).strip()

            # Extract Object
            if "OBJECT" in line:
                match = re.search(r"OBJECT\s*(.*)", line, re.IGNORECASE)
                if match:
                    extracted_data["Object"] = match.group(1).strip()
            elif "RESPONSIBLE" in line:
                match = re.search(r"(\w+)\s+RESPONSIBLE", line)
                if match:
                    word_before_responsible = match.group(1).strip()
                    extracted_data["Object"] = f"{extracted_data.get('Object', '')} {word_before_responsible}".strip()

        # Process Table Data
        for line in lines:
            if "Type DateFraisDevi" in line:  # Detect table header
                headers_found = True
                continue

            if headers_found:
                if "Validation" in line or "Click here" in line or line.strip() == "":
                    continue  # Skip invalid rows

                # Regex to match table rows
                match = re.match(r"(\w+)\s+(\d+\s+\w+\s+\d{4})(\d+)([a-zA-Z]{3})([a-zA-Z]+)", line)
                if match:
                    labelle, date, frais, devis, card = match.groups()
                    frais = int(frais)
                    devis = devis.upper()

                    # Convert Frais to EUR
                    converted_value = frais / rates.get(devis, 1.0)  # Default to original if no rate found
                    
                    # Get Compte Comptable
                    compte_comptable = compte_comptable_mapping.get(labelle.lower(), "Non défini")

                    extracted_data["Table"].append([compte_comptable, labelle, date, frais, devis, round(converted_value, 2), card])
                else:
                    print(f"Skipping unrecognized row: {line}")

    # Convert extracted data to DataFrame
    if extracted_data["Table"]:
        df = pd.DataFrame(extracted_data["Table"], columns=["Compte Comptable", "Libelle", "Date", "Frais", "Devis", "EUR", "Card"])
        df = df.sort_values(by="Libelle")
    else:
        df = pd.DataFrame(columns=["Compte Comptable", "Libelle", "Date", "Frais", "Devis", "EUR", "Card"])  # Empty fallback DataFrame

    # Ensure output directory exists
    os.makedirs(OUTPUT_FOLDER, exist_ok=True)

    # Save DataFrame to Excel
    output_path = os.path.join(OUTPUT_FOLDER, output_filename)
    df.to_excel(output_path, index=False)

    return extracted_data, output_path


def generate_response_html(extracted_data, excel_file_path):
    # Generate HTML to display extracted data and converted Excel
    table_rows = "".join(
        f"<tr><td>{row[0]}</td><td>{row[1]}</td><td>{row[2]}</td><td>{row[3]}</td><td>{row[4]}</td><td>{row[5]}</td></tr>"
        for row in extracted_data["Table"]
    )
    return f'''
    <!doctype html>
    <html lang="en">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Extracted Data</title>
        <style>
            body {{
                font-family: Arial, sans-serif;
                background-color: #f4f4f9;
                margin: 0;
                padding: 20px;
            }}
            .container {{
                max-width: 800px;
                margin: auto;
                background: white;
                padding: 20px;
                border-radius: 8px;
                box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
            }}
            table {{
                width: 100%;
                border-collapse: collapse;
                margin-top: 20px;
            }}
            table, th, td {{
                border: 1px solid #ddd;
                text-align: left;
                padding: 8px;
            }}
            th {{
                background-color: #007bff;
                color: white;
            }}
            a {{
                display: block;
                margin-top: 20px;
                text-align: center;
                color: white;
                background-color: #007bff;
                padding: 10px;
                text-decoration: none;
                border-radius: 4px;
            }}
            a:hover {{
                background-color: #0056b3;
            }}
        </style>
    </head>
    <body>
        <div class="container">
            <h1>Extracted Data</h1>
            <p><strong>Name:</strong> {extracted_data["Libelle"]}</p>
            <p><strong>Department:</strong> {extracted_data["Department"]}</p>
            <p><strong>Object:</strong> {extracted_data["Object"]}</p>
            <h2>Table Data</h2>
            <table>
                <thead>
                      <tr>
                        <th>Compte Comptable</th>
                        <th>Libelle</th>
                        <th>Date</th>
                        <th>Montant en Devise</th>
                        <th>Devise</th>
                        <th>EUR</th>
                        <th>Card</th>
                    </tr>
                </thead>
                <tbody>
                    {table_rows}
                </tbody>
            </table>
            <a href="/download?file={os.path.basename(excel_file_path)}">Download Excel File</a>
        </div>
    </body>
    </html>
    '''


@app.route('/download')
def download_file():
    file_name = request.args.get('file')
    file_path = os.path.join(OUTPUT_FOLDER, file_name)
    if os.path.exists(file_path):
        return send_file(file_path, as_attachment=True)
    else:
        return "File not found", 404


if __name__ == '__main__':
    app.run(debug=True)
