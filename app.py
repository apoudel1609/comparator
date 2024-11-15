import logging
from flask import Flask, request, render_template, send_file
import os
import shutil
import pandas as pd
from comparator import (
    read_names_from_excel,
    highlight_custom_words_in_pdf,
    create_matching_string_excel_file,  # Verify this line
    highlight_names_in_excel_in_pdf
)


# Configure logging
logging.basicConfig(level=logging.DEBUG, filename="app_debug.log", filemode="w",
                    format="%(asctime)s - %(levelname)s - %(message)s")

app = Flask(__name__)

# Ensure 'uploads' directory exists
if not os.path.exists('uploads'):
    os.makedirs('uploads')
    logging.info("Created 'uploads' directory.")

@app.route('/')
def home():
    logging.info("Rendering home page.")
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    logging.info("Received a file upload request.")
    if request.method == 'POST':
        try:
            pdf_file = request.files.get('pdf')
            excel_file = request.files.get('excel')
            custom_string = request.form.get('match_string')

            if pdf_file and excel_file:
                # Define paths
                pdf_path = os.path.join('uploads', pdf_file.filename)
                interim_pdf_path = os.path.join('uploads', 'interim_' + pdf_file.filename)
                final_pdf_path = os.path.join('uploads', 'final_' + pdf_file.filename)
                original_excel_path = os.path.join('uploads', excel_file.filename)
                comp_excel_path = os.path.join('uploads', 'comp.xlsx')
                comp_copy_path = os.path.join('uploads', 'comp_copy.xlsx')
                
                # Save the PDF and Excel files
                pdf_file.save(pdf_path)
                excel_file.save(original_excel_path)
                logging.info("Saved PDF and Excel files.")

                # Check if comp.xlsx exists, if not create a placeholder
                if not os.path.exists(comp_excel_path):
                    pd.DataFrame().to_excel(comp_excel_path, index=False)
                    logging.info("Created a placeholder comp.xlsx as it was missing.")
                
                # Copy comp.xlsx to comp_copy.xlsx
                shutil.copyfile(comp_excel_path, comp_copy_path)
                logging.info("Copied comp.xlsx to comp_copy.xlsx.")

                # First pass: Highlight custom words in PDF using the custom string
                logging.info("Starting first pass to highlight custom words.")
                highlight_custom_words_in_pdf(pdf_path, interim_pdf_path, original_excel_path, custom_string)
                logging.info("Completed first pass for custom words.")

                # Create a new Excel file with matching words
                matching_excel_path = os.path.join('uploads', 'matching_words.xlsx')
                create_matching_string_excel_file(original_excel_path, matching_excel_path, custom_string)
                logging.info("Created matching_words.xlsx.")

                # Second pass: Highlight names in the interim PDF
                logging.info("Starting second pass to highlight names from comp_copy.xlsx.")
                names_to_match = read_names_from_excel(comp_copy_path)
                highlight_names_in_excel_in_pdf(interim_pdf_path, final_pdf_path, names_to_match)
                logging.info("Completed second pass for names.")

                # Send the final PDF back to the user
                logging.info("Sending final PDF to user.")
                return send_file(final_pdf_path, as_attachment=True)

            else:
                logging.error("PDF or Excel file missing in the request.")
                return "PDF or Excel file missing in the request.", 400

        except Exception as e:
            logging.error("Error processing upload request: %s", str(e))
            return "An error occurred while processing the files.", 500

    return render_template('index.html')

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)), debug=True)

