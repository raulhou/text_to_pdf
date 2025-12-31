import os
import comtypes.client
import argparse
import re

def natural_sort_key(filename):
    """
    Splits a string into a list of integers and strings. 
    'unit10.txt' -> ['unit', 10, '.txt']
    This allows sorting to respect numeric value (2 < 10) rather than text (1 < 2).
    """
    return [int(c) if c.isdigit() else c.lower() for c in re.split(r'(\d+)', filename)]

def batch_convert_to_pdf(input_folder, output_folder=None):
    """
    Scans for .txt, .doc, and .docx files and converts them to PDF individually.
    """
    input_folder = os.path.abspath(input_folder)
    output_folder = os.path.abspath(output_folder) if output_folder else input_folder

    if not os.path.exists(input_folder):
        print(f"Error: Folder {input_folder} not found.")
        return

    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    wdFormatPDF = 17 

    word_app = None
    try:
        word_app = comtypes.client.CreateObject("Word.Application")
        word_app.Visible = False

        valid_extensions = (".doc", ".docx", ".txt")
        files = [f for f in os.listdir(input_folder) if f.lower().endswith(valid_extensions)]

        # Apply natural sort for processing order, just in case
        files.sort(key=natural_sort_key)

        if not files:
            print("No compatible documents found.")
            return

        print(f"Found {len(files)} files. Starting individual conversion...")

        for filename in files:
            input_path = os.path.join(input_folder, filename)
            name_without_ext = os.path.splitext(filename)[0]
            output_path = os.path.join(output_folder, f"{name_without_ext}.pdf")

            try:
                doc = word_app.Documents.Open(input_path)
                doc.SaveAs(output_path, FileFormat=wdFormatPDF)
                print(f"Converted: {filename}")
                doc.Close(SaveChanges=0)
            except Exception as e:
                print(f"Failed to convert {filename}: {e}")

    except Exception as e:
        print(f"An error occurred initializing Word: {e}")
    finally:
        if word_app:
            word_app.Quit()

def merge_to_single_pdf(file_paths, output_pdf_path):
    """
    Merges a list of text/doc files into a single PDF with page breaks.
    """
    output_pdf_path = os.path.abspath(output_pdf_path)
    
    # Word Constants
    wdFormatPDF = 17
    wdPage
