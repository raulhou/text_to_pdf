import os
import comtypes.client
import argparse

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

    wdFormatPDF = 17 

    word_app = None
    try:
        word_app = comtypes.client.CreateObject("Word.Application")
        word_app.Visible = False

        valid_extensions = (".doc", ".docx", ".txt")
        files = [f for f in os.listdir(input_folder) if f.lower().endswith(valid_extensions)]

        if not files:
            print("No compatible documents found.")
            return

        print(f"Found {len(files)} files. Starting conversion...")

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
    wdPageBreak = 7

    word_app = None
    try:
        word_app = comtypes.client.CreateObject("Word.Application")
        word_app.Visible = False
        
        # Create a new blank document to act as the container
        merged_doc = word_app.Documents.Add()
        selection = word_app.Selection

        print(f"Merging {len(file_paths)} files into a single PDF...")

        for index, file_path in enumerate(file_paths):
            abs_path = os.path.abspath(file_path)
            if not os.path.exists(abs_path):
                print(f"Warning: File {abs_path} not found. Skipping.")
                continue
            
            # Insert the content of the file
            selection.InsertFile(FileName=abs_path)
            
            # Add a page break after every file except the last one
            if index < len(file_paths) - 1:
                selection.InsertBreak(Type=wdPageBreak)

        # Save the combined document as PDF
        merged_doc.SaveAs(output_pdf_path, FileFormat=wdFormatPDF)
        merged_doc.Close(SaveChanges=0)
        print(f"Success! Merged PDF saved at: {output_pdf_path}")

    except Exception as e:
        print(f"An error occurred during merging: {e}")
    finally:
        if word_app:
            word_app.Quit()

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Convert text/doc files to PDF or merge them.")
    
    # Argument for input folder (optional, defaults to script dir)
    parser.add_argument("--folder", type=str, default=os.path.dirname(os.path.abspath(__file__)), 
                        help="The folder containing the files.")
    
    # Argument to trigger merge mode
    parser.add_argument("--merge", action="store_true", 
                        help="Merge files into a single PDF instead of individual conversion.")
    
    # Argument for specific order
    parser.add_argument("--order", nargs="+", 
                        help="List of filenames to merge in a specific order. Example: --order file1.txt file2.doc")
    
    # Argument for output filename (only used in merge mode)
    parser.add_argument("--output", type=str, default="merged_output.pdf", 
                        help="The output filename for the merged PDF.")

    args = parser.parse_args()

    if args.merge:
        # Determine which files to process
        files_to_process = []
        
        if args.order:
            # Use the specific user-defined order
            # Note: We assume these files are inside the 'folder' or are full paths
            for fname in args.order:
                # Check if user provided full path or just filename
                if os.path.isabs(fname):
                    files_to_process.append(fname)
                else:
                    files_to_process.append(os.path.join(args.folder, fname))
        else:
            # Default: Merge all supported files in the folder alphabetically
            valid_extensions = (".doc", ".docx", ".txt")
            all_files = sorted([f for f in os.listdir(args.folder) if f.lower().endswith(valid_extensions)])
            files_to_process = [os.path.join(args.folder, f) for f in all_files]

        if not files_to_process:
            print("No files found to merge.")
        else:
            output_full_path = args.output
            if not os.path.isabs(output_full_path):
                output_full_path = os.path.join(args.folder, args.output)
                
            merge_to_single_pdf(files_to_process, output_full_path)
            
    else:
        # Default behavior: Batch convert individually
        batch_convert_to_pdf(args.folder)
