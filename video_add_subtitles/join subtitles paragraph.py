#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on 

@author: fmunoz

"""
import pandas as pd



def convert_dmc_time(time_str):
    # Ensure time_str is a string
    time_str = str(time_str)
    # Split the time string into its components
    minutes, seconds_tenths = time_str.split(':')
    seconds, tenths = seconds_tenths.split('.')
    # Return a Timedelta object
    return pd.to_timedelta(f'{minutes} minutes {seconds} seconds {tenths}00 milliseconds')

def convert_dmcclasificacion_time(time_str):
    # Ensure time_str is a string
    time_str = str(time_str)
    # Return a Timedelta object
    return pd.to_timedelta(time_str)


# Load the data with all columns as strings
##df_subtitles = pd.read_csv('taller1DMC.csv', encoding='latin1', dtype=str)
##df_classification = pd.read_csv('taller1DMCCLASIFICACION.csv', encoding='latin1', dtype=str)

df_subtitles = pd.read_csv('taller1SENAMHI.csv', encoding='latin1', dtype=str)
df_classification = pd.read_csv('taller1SENAMHICLASIFICACION.csv', encoding='latin1', dtype=str)


# Drop rows with empty start_time in both dataframes
df_subtitles.dropna(subset=['start_time'], inplace=True)
df_classification.dropna(subset=['start_time'], inplace=True)

# Convert time columns to timedelta for comparison
df_subtitles['start_time'] = df_subtitles['start_time'].apply(convert_dmc_time)
df_classification['start_time'] = df_classification['start_time'].apply(convert_dmcclasificacion_time)

# Calculate end_time for each paragraph by taking the start_time of the next paragraph
df_classification['end_time'] = df_classification['start_time'].shift(-1)
# Handle the last row if needed by setting to the end_time of the last subtitle or an arbitrary large time
df_classification.iloc[-1, df_classification.columns.get_loc('end_time')] = pd.to_timedelta('99:59:59')

# Ensure end_time in df_subtitles is also a timedelta for comparison
df_subtitles['end_time'] = df_subtitles['end_time'].apply(convert_dmc_time)

# Initialize a list to store the new rows
combined_rows = []

# Iterate over each row in the subtitles dataframe
for index, row in df_subtitles.iterrows():
    start_time = row['start_time']
    end_time = row['end_time']

    # Find matching row in classification dataframe
    match = df_classification[
        (df_classification['start_time'] <= start_time) & 
        (df_classification['end_time'] > start_time)
    ]

    if not match.empty:
        # Get speaker, classification, summary from the closest start_time before the subtitle start_time
        speaker = match.iloc[0]['speaker']
        classification = match.iloc[0]['classification']
        summary = match.iloc[0]['summary']

        # Add these to the current row
        combined_row = row.to_dict()
        combined_row.update({'speaker': speaker, 'classification': classification, 'summary': summary})
        combined_rows.append(combined_row)
    else:
        # Handle cases where no match is found
        combined_rows.append(row.to_dict())

# Create a new dataframe from the combined rows
df_combined = pd.DataFrame(combined_rows)

# Save the combined dataframe to a new Excel file
df_combined.to_excel('combined_data_senamhi.xlsx', index=False)

# Calculate end_time for each paragraph by taking the start_time of the next paragraph
df_classification['end_time'] = df_classification['start_time'].shift(-1)

# Handle the last row if needed by setting to an arbitrary large time
df_classification.iloc[-1, df_classification.columns.get_loc('end_time')] = pd.to_timedelta('99:59:59')

# Calculate the duration in seconds and in minutes:seconds format
df_classification['duration_seconds'] = df_classification.apply(
    lambda row: (row['end_time'] - row['start_time']).total_seconds(), axis=1)

df_classification['duration_min_sec'] = df_classification['duration_seconds'].apply(
    lambda x: f"{int(x // 60)}:{int(x % 60):02d}")

# FORMAT WITHOUT DAYS
df_classification['end_time'] = df_classification['end_time'].dt.components.apply(
    lambda x: f"{x.hours:02d}:{x.minutes:02d}:{x.seconds:02d}", axis=1)
df_classification['start_time'] = df_classification['start_time'].dt.components.apply(
    lambda x: f"{x.hours:02d}:{x.minutes:02d}:{x.seconds:02d}", axis=1)

# Save the DataFrame with the additional end_time and duration columns to a new CSV file
#df_classification.to_csv('taller1SENAMHICLASIFICACION_with_end_time.csv', index=False)
df_classification.to_csv('taller1SENAMHICLASIFICACION_with_end_time.csv', index=False)