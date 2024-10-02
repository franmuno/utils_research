import os
import argparse
import json
import pandas as pd
from docxtpl import DocxTemplate, RichText

# Sample config to show if not present
SAMPLE_CONFIG = {
    "excel_folder": "xlsx/",
    "template_folder": "template/",
    "result_folder": "result/",
    "master_input_filename": "Master_list_IARC2024_30May_BOA.xlsx",
    "master_output_filename": "master_output_v3.xlsx",
    "template_input_filename": "tpl_resumenes_IARC2024_v6.docx",
    "template_separator_filename": "tpl_separadores_IARC2024_v3.docx",
    "sheet_name": 0,
    "abstract_id_column": "AbstractId",
    "line_dict": {
        "ARs as a component of compound events": "AR1",
        "ARs in past, present, and future climates": "AR2",
        "Environmental and socioeconomic impacts  ARs": "AR3",
        "Forecasting of ARs": "AR4",
        "Observing, identification, and monitoring of ARs": "AR5",
        "Physical, dynamic, & microphysic aspects of ARs": "AR6",
        "Role of ARs in the changing Cryosphere": "AR7",
        "Rejected": "R"
    },
    "area_column_name": "Theme"
}

# Function to print sample config if no config file is provided
def show_sample_config():
    print("\nSample Configuration (save this as a config.json file):")
    print(json.dumps(SAMPLE_CONFIG, indent=4))

# Function to extract and display the template fields
def extract_template_fields(template_path):
    doc = DocxTemplate(template_path)
    fields = doc.get_undeclared_template_variables()
    return fields

def main():
    # Argument parser to accept config file path
    parser = argparse.ArgumentParser(description='Generate documents from a template using configuration.')
    parser.add_argument('-c', '--config', type=str, help='Path to the configuration JSON file', default='config.json')
    args = parser.parse_args()

    config_path = args.config

    # Check if the config file exists
    if not os.path.exists(config_path):
        print(f"\nError: Configuration file '{config_path}' not found.")
        show_sample_config()
        return

    # Load the config file
    with open(config_path, 'r') as file:
        config = json.load(file)

    # Extract config values
    excel_folder = config['excel_folder']
    template_folder = config['template_folder']
    tmp_result_folder = config['tmp_result_folder']
    master_input_filename = config['master_input_filename']
    master_output_filename = config['master_output_filename']
    template_input_filename = config['template_input_filename']
    template_separator_filename = config['template_separator_filename']
    sheet_name = config['sheet_name']
    AbstractId = config['abstract_id_column']
    LineDict = config['line_dict']
    
    area_column_name = config['area_column_name']
    abstract_column_name = config['abstract_column_name']
    title_column_name = config['title_column_name']

    # READ MASTER EXCEL WITH ALL DATA
    df = pd.read_excel(os.path.join(excel_folder, master_input_filename), sheet_name=sheet_name, engine='openpyxl', dtype=str)
    df.dropna(subset=[AbstractId], inplace=True)
    df = df.fillna('')

    df['AbstractIdFill'] = df[AbstractId].apply(lambda x: str(x).zfill(3))
    df['WordCountSumm'] = df[abstract_column_name].apply(lambda n: len(str(n).split()) if not isinstance(n, float) else 0)

    df['ThemeCode'] = df[area_column_name].map(LineDict).fillna('Unknown')

    df.sort_values(['ThemeCode', 'AbstractIdFill'], ascending=[True, True], inplace=True)
    df.to_excel(os.path.join(excel_folder, master_output_filename), sheet_name='Summary')

    # Ensure the result folder exists
    os.makedirs(tmp_result_folder, exist_ok=True)
    # Extract fields from the main template
    template_fields = extract_template_fields(template_path)
    print(f"Fields found in template '{template_input_filename}': {template_fields}")

    # Extract fields from the separator template
    separator_template_path = os.path.join(template_folder, template_separator_filename)
    separator_template_fields = extract_template_fields(separator_template_path)
    print(f"Fields found in separator template '{template_separator_filename}': {separator_template_fields}")


    # Iterate through the rows OF ALL THE ABSTRACTS
    for _, row in df.iterrows():
        doc = DocxTemplate(os.path.join(template_folder, template_input_filename))

        # Check if there is an image URL
        #if row['AbstractImage'] and pd.notna(row['AbstractImage']) and row['AbstractImage'] != '':
        #    rt = RichText()
        #    rt.add('View Image', url_id=doc.build_url_id(row['AbstractImage']), color='0000FF', underline=True)
        #else:
        #    rt = ""

        context = {
            'Name': row['Name'],
            'Affiliation': row['Affiliation'],
            'City': row['City'],
            'Country': row['Country'],
            
            #'CoauthorsAffiliations': row['CoauthorsAffiliations'],
            'AbsID': row[AbstractId],
            'Theme': row[area_column_name],
            'ThemeCode': row['ThemeCode'],
            'Title': row[title_column_name],
            'Abstract': row[abstract_column_name],
            'Preference': row['Preference'],
            #'rt': rt  # RichText object or None
        }

        doc.render(context)
        doc.add_page_break()

        final_filename = os.path.join(tmp_result_folder, f"{row['ThemeCode']}_{row['AbstractIdFill']}.docx")
        doc.save(final_filename)
        print(f"Single Avstract Document saved: {final_filename}")

    # Iterate through the AREAS - THEMES
    for area_description, theme_code in LineDict.items():
        context = {
            'ThemeCode': theme_code,
            'Theme': area_description
        }

        doc = DocxTemplate(os.path.join(template_folder, template_separator_filename))
        doc.render(context)
        doc.add_page_break()

        final_filename = os.path.join(tmp_result_folder, f"{theme_code}_000.docx")
        doc.save(final_filename)
        print(f"Document saved: {final_filename}")

if __name__ == "__main__":
    main()