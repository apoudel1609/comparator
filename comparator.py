import logging
import fitz  # PyMuPDF
import pandas as pd
import re
import shutil

# Set up logging
logging.basicConfig(level=logging.DEBUG, filename="debug.log", filemode="w",
                    format="%(asctime)s - %(levelname)s - %(message)s")

def read_names_from_excel(excel_path):
    try:
        # Load the entire file first to check what columns are available
        df = pd.read_excel(excel_path)
        
        # Log the columns for debugging
        logging.info(f"Columns in Excel file: {df.columns.tolist()}")
        
        # Check if column A exists by name or index
        if "A" not in df.columns and df.shape[1] > 0:
            # Attempt to read the first column by index if "A" isn't recognized as a header
            df.columns = ["A"]
        
        # Extract names from column A, dropping any NaN values
        names = [(name, False) for name in df["A"].dropna()]
        logging.info("Read names from Excel successfully.")
        return names
    except Exception as e:
        logging.error(f"Error reading names from Excel: {e}")
        raise


def create_matching_string_excel_file(original_excel_path, matching_excel_path, custom_string):
    try:
        df = pd.read_excel(original_excel_path)
        matching_df = df[df.iloc[:, 0].str.contains(custom_string, case=False, na=False)]
        matching_df.to_excel(matching_excel_path, index=False)
        logging.info("Created matching string Excel file successfully.")
    except Exception as e:
        logging.error(f"Error creating matching string Excel file: {e}")
        raise

def highlight_custom_words_in_pdf(pdf_path, interim_pdf_path, excel_path, custom_string):
    # Check if custom_string is empty
    if not custom_string:
        logging.warning("Custom string is empty. Copying original PDF to interim file and skipping highlighting.")
        # Copy the original PDF to the interim path if custom_string is empty
        shutil.copy(pdf_path, interim_pdf_path)
        return

    try:
        words_found = set()
        doc = fitz.open(pdf_path)
        pattern = re.compile(r'\b\w*' + re.escape(custom_string) + r'\w*\b', re.IGNORECASE)

        for page_num, page in enumerate(doc, start=1):
            text = page.get_text("text")
            words = pattern.finditer(text)
            for word in words:
                word_text = word.group()

                # Skip empty words
                if not word_text.strip():
                    logging.warning(f"Skipping empty word on page {page_num}")
                    continue

                words_found.add(word_text)
                word_bbox = page.search_for(word_text)
                if word_bbox is None:
                    logging.warning(f"No bounding box found for '{word_text}' on page {page_num}")
                    continue
                for rect in word_bbox:
                    annot = page.add_highlight_annot(rect)
                    annot.set_colors(stroke=(0, 0, 1))  # Blue color for highlighting
                    annot.update()
                    annot.set_opacity(0.3)

        # Save words to an Excel file if any words were found
        if words_found:
            df = pd.DataFrame(list(words_found), columns=[f"Words Containing '{custom_string}'"])
            df.to_excel(excel_path, index=False)
            logging.info("Saved highlighted words to Excel.")
        else:
            logging.info("No matching words found to save to Excel.")

        # Save the interim PDF
        doc.save(interim_pdf_path)
        doc.close()
        logging.info("Completed highlighting custom words in PDF.")
    except Exception as e:
        logging.error(f"Error in highlight_custom_words_in_pdf: {e}")
        raise

def highlight_names_in_excel_in_pdf(interim_pdf_path, final_pdf_path, names):
    try:
        doc = fitz.open(interim_pdf_path)
        for page_num, page in enumerate(doc, start=1):
            text = page.get_text("text").lower()

            # Iterate over the list of names from Excel
            for idx, (name, found) in enumerate(names):
                if not found:
                    # Convert name to lowercase for case-insensitive search
                    name_lower = name.lower()

                    # Search only for exact matches of the name in the PDF text
                    if name_lower in text:
                        names[idx] = (name, True)  # Mark name as found
                        text_instances = page.search_for(name)  # Find the exact text instances

                        if text_instances is None:
                            logging.warning(f"No bounding box found for name '{name}' on page {page_num}")
                            continue

                        # Highlight each found instance of the name
                        for inst in text_instances:
                            annot = page.add_highlight_annot(inst)
                            annot.set_colors(stroke=(0, 1, 0))  # Green color for highlighting
                            annot.update()
                            annot.set_opacity(0.3)

        doc.save(final_pdf_path)
        doc.close()
        logging.info("Completed highlighting names in Excel in PDF.")
        return names
    except Exception as e:
        logging.error(f"Error in highlight_names_in_excel_in_pdf: {e}")
        raise
