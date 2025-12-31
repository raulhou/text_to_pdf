import os
import sys
import subprocess
import argparse
import re

# --- Dependency Check & Auto-Install ---
def ensure_dependencies():
    """
    Checks for required external libraries and installs them if missing.
    """
    try:
        import comtypes.client
    except ImportError:
        print("Dependency 'comtypes' not found. Attempting to install...")
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install", "comtypes"])
            print("Successfully installed 'comtypes'.")
            global comtypes
            import comtypes.client
        except subprocess.CalledProcessError as e:
            print(f"Error: Failed to auto-install 'comtypes'. Please run: pip install comtypes")
            print(f"Details: {e}")
            sys.exit(1)

ensure_dependencies()
import comtypes.client

# --- Helper Functions ---

def natural_sort_key(filename):
    """
    Splits a string into a list of integers and strings. 
    'unit10.txt' -> ['unit', 10, '.txt']
    """
    return [int(c) if c.isdigit() else c.lower() for c in re.split(r'(\d+)', filename)]

def sanitize_bookmark_name(filename, index):
    """
    Creates a valid Word bookmark name from a filename.
    Rules: Must start with a letter, no spaces, alphanumeric only, max 40 chars.
    """
    # Remove extension
    name = os.path.splitext(filename)[0]
    # Replace invalid characters (non-alphanumeric) with underscores
    clean_name = re.sub(r'\W', '_', name)
    # Ensure it starts with a letter and is unique using the index prefix
    # Example: "01_Unit_1_Intro"
    bookmark_name = f"B{index:02d}_{clean_name}"
    # Truncate to 40 chars to satisfy Word's limit
    return bookmark_name[:40]

# --- Core Logic ---

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
    Merges files into a single PDF with page breaks and bookmarks.
    """
    output_pdf_path = os.path.abspath(output_pdf_path)
    
    # Word Constants
    wdFormatPDF = 17
    wdPageBreak = 7
    wdExportCreateWordBookmarks = 2  # Essential for generating the PDF index

    word_app = None
    try:
        word_app = comtypes.client.CreateObject("Word.Application")
        word_app.Visible = False
        
        merged_doc = word_app.Documents.Add()
        selection = word_app.Selection

        print(f"Merging {len(file_paths)} files into a single PDF with bookmarks...")

        for index, file_path in enumerate(file_paths):
            abs_path = os.path.abspath(file_path)
            if not os.path.exists(abs_path):
                print(f"Warning: File {abs_path} not found. Skipping.")
                continue
            
            # 1. Create a bookmark for navigation
            filename = os.path.basename(file_path)
            b_name = sanitize_bookmark_name(filename, index)
            
            # Add bookmark at current insertion point (Start of the new file)
            merged_doc.Bookmarks.Add(Name=b_name, Range=selection.Range)
            
            # 2. Insert the file content
            selection.InsertFile(FileName=abs_path)
            
            # 3. Add page break if not the last file
            if index < len(file_paths) - 1:
                selection.InsertBreak(Type=wdPageBreak)

        # 4. Export using ExportAsFixedFormat to include Bookmarks
        # Note: 'CreateBookmarks=wdExportCreateWordBookmarks' is the key here
        merged_doc.ExportAsFixedFormat(
            OutputFileName=output_pdf_path, 
            ExportFormat=wdFormatPDF, 
            CreateBookmarks=wdExportCreateWordBookmarks
        )
        
        merged_doc.Close(SaveChanges=0)
        print(f"Success! Merged PDF saved at: {output_pdf_path}")

    except Exception as e:
        print(f"An error occurred during merging: {e}")
    finally:
        if word_app:
            word_app.Quit()

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Convert text/doc files to PDF or merge them.")
    
    parser.add_argument("--folder", type=str, default=os.path.dirname(os.path.abspath(__file__)), 
                        help="The folder containing the files.")
    parser.add_argument("--merge", action="store_true", 
                        help="Merge files into a single PDF instead of individual conversion.")
    parser.add_argument("--order", nargs="+", 
                        help="List of filenames to merge in a specific order.")
    parser.add_argument("--output", type=str, default="merged_output.pdf", 
                        help="The output filename for the merged PDF.")

    args = parser.parse_args()

    if args.merge:
        files_to_process = []
        
        if args.order:
            for fname in args.order:
                if os.path.isabs(fname):
                    files_to_process.append(fname)
                else:
                    files_to_process.append(os.path.join(args.folder, fname))
        else:
            valid_extensions = (".doc", ".docx", ".txt")
            all_files = [f for f in os.listdir(args.folder) if f.lower().endswith(valid_extensions)]
            all_files.sort(key=natural_sort_key)
            files_to_process = [os.path.join(args.folder, f) for f in all_files]

        if not files_to_process:
            print("No files found to merge.")
        else:
            output_full_path = args.output
            if not os.path.isabs(output_full_path):
                output_full_path = os.path.join(args.folder, args.output)
            
            print("Processing files in this order:")
            for f in files_to_process:
                print(f" -> {os.path.basename(f)}")
                
            merge_to_single_pdf(files_to_process, output_full_path)
            
    else:
        batch_convert_to_pdf(args.folder)
