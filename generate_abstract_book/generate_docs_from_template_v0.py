#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on 2024-01-01

@author: fmunoz
Genera docx basado en template
"""

import pandas as pd
#from docxtpl import DocxTemplate
from docxtpl import DocxTemplate, RichText
#from datetime import date

excel_folder= 'xlsx/'
template_folder='template/'

master_input_filename='Master_list_IARC2024_30May_BOA.xlsx'
master_output_filename = 'master_output_v3.xlsx'
template_input_filename='tpl_resumenes_IARC2024_v6.docx'
template_separator_filename='tpl_separadores_IARC2024_v3.docx'
sheet_name=0

###AbstractId = 'Number' # now its AbstractId
AbstractId = 'AbstractId'

# SE PUEDE SACAR DEL XLSX DE SEPARADORES
LineDict = {'ARs as a component of compound events': 'AR1', 
            'ARs in past, present, and future climates': 'AR2', 
            'Environmental and socioeconomic impacts  ARs': 'AR3', 
            'Forecasting of ARs': 'AR4',
            'Observing, identification, and monitoring of ARs': 'AR5', 
            'Physical, dynamic, & microphysic aspects of ARs': 'AR6',
            'Role of ARs in the changing Cryosphere': 'AR7',
            'Rejected': 'R'}

# Define a function to concatenate non-empty strings with a separator
def concatenate_non_empty(row, separator=", "):
    return separator.join(filter(None, row))

# READ MASTER ESXEL WITH ALL DATA
df = pd.read_excel(excel_folder+master_input_filename, sheet_name=sheet_name, engine='openpyxl', dtype = str)
df.dropna(subset=[AbstractId], inplace=True)
df = df.fillna('')

print(df)
# KEEP first m columns  (not necessary)
#m=20
#df = df.iloc[: , :m]
df.dropna(subset=[AbstractId], inplace=True)

###area_column_name='Final Theme'
area_column_name='Theme'
###no es necesario df[area_column_name] = df[area_column_name].replace('AC.3','AC3.', regex=True)
### Coauthor position: OBS comienza con 0
##df.iloc[:,13] = df.iloc[:,13].fillna('-')
### Rest of columns: fill empty with ''
df = df.fillna('') 

##df['AllCoauthors'] = df.iloc[:, 13:20].apply(concatenate_non_empty, axis=1)

###df['Prefix'] = df[area_column_name].str.split('.').str[0]
###prefix_counts = df.groupby('Prefix').size().reset_index(name='CountMaster')
####print(prefix_counts)

df['AbstractIdFill']= df[AbstractId].apply(lambda x: str(x).zfill(3))

df['WordCountSumm'] = df['AbstractSummary'].apply(lambda n: len(str(n).split()) if not isinstance(n, float) else 0)

df['WordCountCoauth'] = df['CoauthorsAffiliations'].apply(lambda n: len(str(n).split()) if not isinstance(n, float) else 0)

df['ThemeCode'] = df[area_column_name].map(LineDict).fillna('Unknown')

df.sort_values(['ThemeCode', 'AbstractIdFill'], ascending=[True, True], inplace=True)

df.to_excel(excel_folder+master_output_filename, sheet_name='Summary')

# Iterate through the rows OF ALL THE ABSTRACTS!!!!!
for _, row in df.iterrows():
    doc = DocxTemplate(template_folder+template_input_filename)
    
    # Check if there is an image URL
    if pd.notna(row['AbstractImage']) and row['AbstractImage'] != '':
        rt = RichText()
        rt.add('View Image', url_id=doc.build_url_id(row['AbstractImage']), color='0000FF', underline=True)
        #image_url = row['AbstractImage']
    else:
        rt = ""

#rt = RichText()
#rt.add('eHYD Link', url_id=doc.build_url_id('https://ehyd.gv.at/'))
#context_to_load['rt'] = rt
    
    context = {'Name': row['Name'],
               'Affiliation': row['Affiliation'],
               'Country': row['Country'],
               'CoauthorsAffiliations': row['CoauthorsAffiliations'],
               'AbstractId': row[AbstractId],
               'Theme': row[area_column_name],
               'ThemeCode': row['ThemeCode'],
               'AbstractTitle': row['AbstractTitle'],
               'AbstractSummary': row['AbstractSummary'],
               'rt': rt,  # RichText object or None
               }

    # Using the value in 'area_column_name' as a key to find the corresponding value in LineDict



    doc.render(context)
    doc.add_page_break()
    final_filename = 'result/'+row['ThemeCode']+'_'+row['AbstractIdFill']+'.docx'
    doc.save(final_filename)
    print(final_filename)
    
# Iterate through the AREAS - THEMES (!)
for area_description, theme_code in LineDict.items():
    # Define the context for the DOCX template using the current area and theme code
    context = {
        'ThemeCode': theme_code,
        'Theme': area_description,
    }
    # Load the DOCX template
    doc = DocxTemplate(template_folder+ template_separator_filename)
    # Render the context into the template
    doc.render(context)
    doc.add_page_break()
    final_filename = 'result/'+theme_code+'_000'+'.docx'
    # Save the rendered document
    doc.save(final_filename)
    print(f"Document saved: {final_filename}")