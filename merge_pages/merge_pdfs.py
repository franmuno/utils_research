import os
import re
from PyPDF2 import PdfMerger
from pathlib import Path

def merge_pdfs_in_folder(folder_path, output_filename):
    """
    Merge all PDF files in the specified folder into a single PDF.

    Args:
        folder_path (str): The folder containing the PDF files.
        output_filename (str): The output file name for the merged PDF.
    """
    # List all the PDF files in the folder and sort them
    ordered_files = sorted(os.listdir(folder_path), key=lambda x: (int(re.sub('\D', '', x)), x))
    
    # Filter only PDF files
    pdf_list = [os.path.join(folder_path, f) for f in ordered_files if f.endswith(".pdf")]
    
    # Create a PdfMerger object
    merger = PdfMerger()
    
    # Merge the PDFs
    for pdf in pdf_list:
        print(f"Adding {pdf} to the merger.")
        merger.append(pdf)
    
    # Save the merged PDF
    output_path = os.path.join(folder_path, output_filename)
    with open(output_path, "wb") as fout:
        merger.write(fout)
    merger.close()
    
    print(f"Merged PDF saved as {output_path}")


def main():
    # Set the folder path and output file name
    folder_path = "./Anales de la Universidad de Chile  Appendix (1873)"  # Replace with the path to your folder containing the PDFs
    output_filename = "../Anales de la Universidad de Chile  Appendix (1873).pdf"
    
    # Ensure the folder exists
    Path(folder_path).mkdir(parents=True, exist_ok=True)
    
    # Merge PDFs in the folder
    merge_pdfs_in_folder(folder_path, output_filename)

if __name__ == "__main__":
    main()