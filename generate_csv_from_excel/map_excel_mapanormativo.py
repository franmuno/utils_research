#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on 2023-06-13

@author: fmunoz
Este programa lee la tabla excel del Mapa Normativo de los Elementos, enviada por correo por Antoine Mailet.
Cambia el orden de las columnas para su despliegue en el plugin del wordpress, 
yllena algunos campos para claridad del lector.

Luego lo pasa a CSV para cargarlo directamente a la tabla 42 
en el plugin Tablepress

Se usan dataframes de pandas para cargar excel y generar csv

Updated on 2024-05-210
"""
import os
import pandas as pd

# Constants for input/output paths and file names
FOLDER_IN = 'xlsx/'
FOLDER_OUT = 'csv/'
TIMESTAMP_MASTER = '20240812'
MASTER_FILENAME = os.path.join(FOLDER_IN, f'MapaNormativoElementos_{TIMESTAMP_MASTER}.xlsx')
SHEET_NAME1 = '4. Base de Datos'
SHEET_NAME2 = '5. Leyes Transversales'
OUT_CSV_FILENAME1 = os.path.join(FOLDER_OUT, f'MapaNormativo_BaseDatos_{TIMESTAMP_MASTER}.csv')
OUT_CSV_FILENAME2 = os.path.join(FOLDER_OUT, f'MapaNormativo_LeyesTransv_{TIMESTAMP_MASTER}.csv')


# Dictionaries for conversion
TIPO_ACCION_TEXTO = {
    "RA": "Regular acceso",
    "CO": "Controlar",
    "PP": "Proteger preventivamente",
    "RE": "Restaurar",
    "SA": "Sancionar"
}

TIPO_ACTOR_TEXTO = {
    "a": "Estado",
    "b": "Empresa",
    "c": "ONGs",
    "d": "Asociaciones comunitarias",
    "e": "Academia",
    "f": "Individuos"
}

def convertir_tipo_accion(texto):
    #print (texto)
    # Ensure the input is a string to handle cases where it might be NaN or another type
    if not isinstance(texto, str):
        return ""  # or return texto if you want to keep the original NaN values

    # Separa las siglas y elimina espacios
    partes = texto.split('-')
    partes = [p.strip() for p in partes]
    # Convierte cada sigla usando el diccionario
    descripciones = [tipo_accion_texto.get(parte, parte) for parte in partes]
    # Une las descripciones con coma
    return ', '.join(descripciones)

def convertir_tipo_actor(texto):
    #print (texto)
    if not isinstance(texto, str):
        return ""  # or return texto if you want to keep the original NaN values
    # Extrae solo la parte antes del paréntesis si existe
    texto = texto.split('(')[0]
    # Separa las categorías por comas
    categorias = texto.split(',')
    # Elimina espacios extra y convierte cada categoría
    descripciones = [tipo_actor_texto.get(cat.strip(), cat.strip()) for cat in categorias]
    # Une las descripciones con coma
    return ', '.join(descripciones)

def append_url(row):
    """Append URL as a hyperlink to legal/regulatory texts."""
    if row['URL Normativa']:
        if row['Textos legales']:
            row['Textos legales'] += f", <a href='{row['URL Normativa']}'> (enlace)</a>"
        elif row['Textos reglamentarios']:
            row['Textos reglamentarios'] += f", <a href='{row['URL Normativa']}'> (enlace)</a>"
    return row

def concat_non_empty(row):
    """Concatenate non-empty elements from the row with ' - '."""
    return ' - '.join([val for val in row if val])

df1 = pd.read_excel(master_filename, sheet_name=sheet_name1, engine='openpyxl', dtype = str)
# Replace 'nan' strings with an actual empty string or any other placeholder
#df1.replace('nan', '', inplace=True)

df1.fillna('', inplace=True)
df1 = df1.replace('\n',' ', regex=True)
#COLUMNS
df1.dropna(axis=1, how='all', inplace=True)
#ROWS
df1.dropna(how='all', inplace=True)

df2 = pd.read_excel(master_filename, sheet_name=sheet_name2, engine='openpyxl', dtype = str)
df2.replace('nan', '', inplace=True)
df2.dropna(axis=1, how='all', inplace=True)
df2.dropna(how='all', inplace=True)

print(df1.columns)
print("Columnas "+master_filename+" "+sheet_name1)
print(len(df1.columns))

column_names1 = [
    "Identificador",  # 1.
    "Elemento Natural",  # 2.
    "Ámbito del elemento", # 3.
    "Área de Impacto", #4. "Categorización espacial-cualitativa del elemento",
    "Desafío socioambiental", #5. "Desafíos y problemas socioambientales a investigar",
    "Nivel de Gobierno", #6.
    "Tipo de Normativa", #7.
    "Textos legales", #8.
    "Textos reglamentarios", #9.
    "URL Normativa", #10.
    "Artículo/s", #11.
    "Acción específica", #12.
    "Instrumento o programa de políticas públicas", #13. "Instrumentos o programas de políticas públicas",    
    "Acción específica del instrumento", #15.
    "Tipo de acción", #16.
    "Actor que da origen a la acción", #17.
    "Tipo de actor que da origen a la acción", #18. "Tipo de actor que da origen a la acción a) Estado b) Empresa c) ONGs d)Asociaciones comunitarias e) Academia f) Individuos",
    "Actor que implementa la acción", #19.
    "Tipo de actor que implementa la acción", #20. "Tipo de actor que implementa la acción a) Estado b) Empresa c) ONGs d)Asociaciones comunitarias e) Academia f) Individuos"
]

### "URL Instrumento", #14.  

print(column_names1)
print(len(column_names1))
df1.columns=column_names1

# Checkeo-Mapeo contenido para legibilidad online:
# Suponiendo que df1 es tu DataFrame y "Tipo de acción" es la columna a transformar
df1['Tipo de acción'] = df1['Tipo de acción'].apply(convertir_tipo_accion)
df1['Tipo de actor que da origen a la acción'] = df1['Tipo de actor que da origen a la acción'].apply(convertir_tipo_actor)
df1['Tipo de actor que implementa la acción'] = df1['Tipo de actor que implementa la acción'].apply(convertir_tipo_actor)

df1['Actor que da origen a la acción'] = df1['Actor que da origen a la acción']+". <emph>Tipo de actor</emph>: "+df1['Tipo de actor que da origen a la acción']
df1['Actor que implementa la acción'] = df1['Actor que implementa la acción']+". <emph>Tipo de actor</emph>: "+df1['Tipo de actor que implementa la acción']

def concat_non_empty(row):
    return ' -  '.join([val for val in row if val])
# Define a custom function to append the URL as a hyperlink
def append_url(row):
    if row['URL Normativa']:
        if row['Textos legales']:
            row['Textos legales'] = row['Textos legales'] + ", <a href='" + row['URL Normativa'] + "'>(enlace)</a>"
        elif row['Textos reglamentarios']:
            row['Textos reglamentarios'] = row['Textos reglamentarios'] + ", <a href='" + row['URL Normativa'] + "'>(enlace)</a>"
    return row

df1['Resumen Normativas']= df1[['Textos legales', 'Textos reglamentarios', 'Instrumento o programa de políticas públicas']].apply(concat_non_empty, axis=1)
####df1["Textos legales"]+" | "+df1["Textos reglamentarios"]+" | "+df1["Instrumento o programa de políticas públicas"]
###df1['Actor que origina']= df1['Actor que da origen a la acción']+" ("+df1['Tipo de actor que da origen a la acción']+")"
df1['Actor que origina']= df1['Actor que da origen a la acción']

# Agregar enlaces si hay. Ojo, despues de la concatenacion de los textos:
df1 = df1.apply(append_url, axis=1)

##if df1['URL Instrumento']:
    ##df1['Instrumento o programa de políticas públicas'] = df1['Instrumento o programa de políticas públicas']+"<a href='"+df1['URL Instrumento']+"'>(enlace)</a>" 



print(df1)
### CAMBIAR EL ORDEN DE SALIDA AL GUARDAR COMO PDF
### M:Elemento ,   D:Desafíos y problemas socioambientales a investigar ..   
### F: Textos legales    .. 
### G: Textos reglamentarios   .. 
### M: Actor que da origen a la acción 
desired_order = ["Elemento Natural", 
                 "Desafío socioambiental",
                 "Tipo de Normativa",
                 "Resumen Normativas",
                 "Actor que origina",
    "Identificador", 
    "Desafío socioambiental",     
    "Ámbito del elemento",
    "Área de Impacto", #"Categorización espacial-cualitativa del elemento",
    "Nivel de Gobierno",
    "Tipo de Normativa",
    "Textos legales",
    "Textos reglamentarios",
    "Artículo/s",
    "Acción específica",    
    "Instrumento o programa de políticas públicas", #"Instrumentos o programas de políticas públicas",:w
    "Acción específica del instrumento",
    "Tipo de acción", #"Tipo de acción Regular acceso: RA Controlar: CO Proteger preventivamente: PP Restaurar: RE Sancionar: SA",
    "Actor que da origen a la acción",
    ###"Tipo de actor que da origen a la acción", #"Tipo de actor que da origen a la acción a) Estado b) Empresa c) ONGs d)Asociaciones comunitarias e) Academia f) Individuos",
    "Actor que implementa la acción",
    ###"Tipo de actor que implementa la acción", #"Tipo de actor que implementa la acción a) Estado b) Empresa c) ONGs d)Asociaciones comunitarias e) Academia f) Individuos"
]
df1 = df1[desired_order]

###df1.to_csv('reordered_df.csv', index=False)  # Set index=False to avoid saving the row index

df1.to_csv(out_csv_filename1, index=False)
df2.to_csv(out_csv_filename2, index=False)

print(df2.columns)



#df2 = df2.add_prefix("file_")
#df = df.add_prefix("master_")
#print(df.columns)
#print(df2.columns)

#merged_df = pd.merge(df, df2, left_on=['master_Abstract', 'master_Prefix'], right_on=['file_ID', 'file_Prefix'], how='outer')


