import os
import comtypes.client

def batch_convert_to_pdf(input_folder, output_folder=None):
    """
    Scans for .txt, .doc, and .docx files and converts them to PDF using MS Word.
    """
    # Normalize paths for COM automation
    input_folder = os.path.abspath(input_folder)
    output_folder = os.path.abspath(output_folder) if output_folder else input_folder

    if not os.path.exists(input_folder):
        print(f"Error: Folder {input_folder} not found.")
        return

    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    # Word constant for PDF format export
    wdFormatPDF = 17 

    word_app = None
    try:
        # Create Microsoft Word application instance
        word_app = comtypes.client.CreateObject("Word.Application")
        word_app.Visible = False # Run in background

        # Filter for Word and Text documents
        valid_extensions = (".doc", ".docx", ".txt")
        files = [f for f in os.listdir(input_folder) if f.lower().endswith(valid_extensions)]

        if not files:
            print("No compatible documents found.")
            return

        print(f"Found {len(files)} files. Starting conversion...")

        for filename in files:
            input_path = os.path.join(input_folder, filename)
            
            # Generate output file name (replacing old extension with .pdf)
            name_without_ext = os.path.splitext(filename)[0]
            output_path = os.path.join(output_folder, f"{name_without_ext}.pdf")

            try:
                # Open document (Word automatically handles .txt encoding)
                doc = word_app.Documents.Open(input_path)
                
                # Export to PDF
                doc.SaveAs(output_path, FileFormat=wdFormatPDF)
                print(f"Converted: {filename}")
                
                # Close the document without saving changes
                doc.Close(SaveChanges=0)
            except Exception as e:
                print(f"Failed to convert {filename}: {e}")

    except Exception as e:
        print(f"An error occurred initializing Word: {e}")
    finally:
        if word_app:
            word_app.Quit()

if __name__ == "__main__":
    # Define target folder (defaults to script's directory)
    target = os.path.dirname(os.path.abspath(__file__))
    batch_convert_to_pdf(target)