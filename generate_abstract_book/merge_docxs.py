import os
import argparse
from docxcompose.composer import Composer
from docx import Document

def merge_docs(output_path, input_paths):
    """
    Merge multiple DOCX files into a single document.

    Args:
        output_path (str): The output file path for the merged DOCX.
        input_paths (list): List of file paths to DOCX files to merge.
    """
    # Load the first document as the base document
    base_doc = Document(input_paths[0])
    composer = Composer(base_doc)

    # Append all other documents
    for file_path in input_paths[1:]:
        doc = Document(file_path)
        composer.append(doc)

    # Save the merged document
    composer.save(output_path)
    print(f"Documents merged successfully into {output_path}")

def main():
    # Argument parser to receive input folder and output filename
    parser = argparse.ArgumentParser(description="Merge DOCX files from a folder into a single DOCX file.")
    parser.add_argument('-i', '--input-folder', type=str, required=True, help="Path to the folder containing DOCX files.")
    parser.add_argument('-o', '--output-file', type=str, required=True, help="Output file path for the merged DOCX file.")

    args = parser.parse_args()

    # Get the input folder and output file from the arguments
    input_folder = args.input_folder
    output_file = args.output_file

    # Get a list of all .docx files in the input folder
    docx_files = sorted([os.path.join(input_folder, f) for f in os.listdir(input_folder) if f.endswith('.docx')])

    # Check if there are any .docx files in the folder
    if not docx_files:
        print(f"No DOCX files found in folder: {input_folder}")
        return

    # Merge the DOCX files into the output file
    merge_docs(output_file, docx_files)

if __name__ == "__main__":
    main()
