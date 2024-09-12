import pandas as pd
import webvtt
from io import StringIO

def vtt_to_csv(vtt_path, csv_path):
    # Try reading the file with different encodings
    encodings = ['utf-8', 'iso-8859-1', 'windows-1252']
    for encoding in encodings:
        try:
            # Read the file content with the specified encoding
            with open(vtt_path, 'r', encoding=encoding) as file:
                content = file.read()

            # Parse the VTT content from string
            captions = webvtt.read_buffer(StringIO(content))

            # Extract timestamp and text
            rows = []
            for caption in captions:
                start = caption.start
                end = caption.end
                text = caption.text.replace('\n', ' ').strip()
                rows.append([start, end, text])

            # Create a DataFrame and save to CSV
            df = pd.DataFrame(rows, columns=['Start', 'End', 'Text'])
            df.to_csv(csv_path, index=False)
            print(f"File converted successfully with encoding: {encoding}")
            break
        except UnicodeDecodeError:
            print(f"Failed to decode with {encoding}, trying next.")
        except Exception as e:
            print(f"An error occurred: {e}")
            break


# Usage
vtt_file = 'VTT - video1997169842.mp4.vtt'  # Replace with your VTT file path
csv_file = 'taller3DMC_subtitulos.csv'  # Replace with your desired CSV file path
vtt_to_csv(vtt_file, csv_file)