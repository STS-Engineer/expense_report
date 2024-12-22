from flask import Flask, request, send_file, render_template_string, jsonify
import os
from PyPDF2 import PdfReader
import pandas as pd
import requests
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
        # Extract data and convert to Excel
        extracted_data, excel_file_path = convert_pdf_to_excel(file_path, file.filename.replace('.pdf', '.xlsx'))
        return render_template_string(generate_response_html(extracted_data, excel_file_path))
    except Exception as e:
        return jsonify({"error": str(e)}), 500


def convert_pdf_to_excel(pdf_path, output_filename):
    reader = PdfReader(pdf_path)
    extracted_data = {"Labelle": None, "Department": None, "Object": None, "Table": []}
    import re
    # Fetch conversion rates from Fixer.io
    try:
        response = requests.get(FIXER_API_URL, params={"access_key": FIXER_API_KEY})
        response.raise_for_status()
        conversion_data = response.json()
        print("Fetched conversion rates:", conversion_data)  # Debugging
        rates = conversion_data.get("rates", {})
    except Exception as e:
        print(f"Failed to fetch conversion rates: {e}")
        rates = {}

    for page in reader.pages:
        text = page.extract_text()
        lines = text.split('\n')
        print("Extracted Lines:", lines)  # Debugging: Print all lines extracted from the PDF

        # Flag to detect when table starts
        headers_found = False
        for line in lines:
            # Extract Name and Department
            if "NAME" in line and "DEPARTMENT" in line:
                match = re.search(r"NAME(.*?)DEPARTMENT(.*)", line)
                if match:
                    extracted_data["Labelle"] = match.group(1).strip()
                    extracted_data["Department"] = match.group(2).strip()
                    print(f"Extracted Labelle: {extracted_data['Labelle']}, Department: {extracted_data['Department']}")

            # Extract Object
            if "OBJECT" in line:
                match = re.search(r"OBJECT(.*)", line)
                if match:
                    extracted_data["Object"] = match.group(1).strip()

                # Append word before RESPONSIBLE
            elif "RESPONSIBLE" in line:
                match = re.search(r"(\w+)\s+RESPONSIBLE", line)
                if match:
                    word_before_responsible = match.group(1).strip()
                    if extracted_data["Object"]:
                        extracted_data["Object"] += f" {word_before_responsible}"
                    else:
                        extracted_data["Object"] = word_before_responsible
        for line in lines:
            # Detect the table header
            if "Name DateFraisDevi" in line:
                headers_found = True
                print(f"Table header detected: {line}")
                continue

            # Process table rows
            if headers_found:
                print(f"Processing potential table row: {line}")  # Debugging

                # Skip invalid rows
                if "Validation" in line or "Click here" in line or line.strip() == "":
                    print(f"Skipping non-table row: {line}")
                    continue

                # Handle rows like "Plane 20 Dec 202430EUR"
                match = re.match(r"(\w+)\s+(\d+\s+\w+\s+\d{4})(\d+)([a-zA-Z]{3})([a-zA-Z]+)", line)
                if match:
                    labelle = match.group(1)  # First column: Labelle
                    date = match.group(2)  # Second column: Date
                    frais = float(match.group(3))  # Third column: Frais
                    devis = match.group(4).upper()  # Fourth column: Currency (e.g., "EUR", "USD", "TND")
                    card = match.group(5)
                    # Convert Frais to EUR
                    rate_to_eur = rates.get(devis, None)
                    if rate_to_eur:
                        converted_value = frais / rate_to_eur
                    else:
                        print(f"No rate found for currency {devis}, using original Frais")
                        converted_value = frais  # Default to original Frais value

                    # Add a placeholder for Card column
                    extracted_data["Table"].append([labelle, date, frais, devis, round(converted_value, 2), card])  # Empty card
                    print(f"Extracted Table Row: {[labelle, date, frais, devis, round(converted_value, 2), card]}")
                else:
                    print(f"Failed to parse row: {line}")

    # Convert the table to a DataFrame
    if extracted_data["Table"]:
        df = pd.DataFrame(
            extracted_data["Table"],
            columns=["Labelle", "Date", "Frais", "Devis", "EUR", "Card"]
        )

        # Summing EUR values for the same Labelle
        summary = df.groupby("Labelle")["EUR"].sum().reset_index()
        summary.rename(columns={"EUR": "Total EUR"}, inplace=True)

        # Add summary rows under each group
        df = df.sort_values(by="Labelle")
        summary_rows = []
        for labelle in df["Labelle"].unique():
            group_rows = df[df["Labelle"] == labelle]
            total_eur = group_rows["EUR"].sum()
            summary_rows.append(group_rows)
            summary_rows.append(pd.DataFrame({
                "Labelle": [labelle],
                "Date": ["Total"],
                "Frais": [""],
                "Devis": [""],
                "EUR": [total_eur],
                "Card": [""]
            }))

        df = pd.concat(summary_rows, ignore_index=True)

    else:
        df = pd.DataFrame(columns=["Labelle", "Date", "Frais", "Devis", "EUR", "Card"])  # Empty DataFrame fallback

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
            <p><strong>Name:</strong> {extracted_data["Labelle"]}</p>
            <p><strong>Department:</strong> {extracted_data["Department"]}</p>
            <p><strong>Object:</strong> {extracted_data["Object"]}</p>
            <h2>Table Data</h2>
            <table>
                <thead>
                    <tr>
                        <th>Labelle</th>
                        <th>Date</th>
                        <th>Frais</th>
                        <th>Devis</th>
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
