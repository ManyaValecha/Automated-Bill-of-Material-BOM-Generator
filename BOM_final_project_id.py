import ssl
import certifi
from PIL import Image, ImageEnhance, ImageFilter
import pandas as pd
import re
import requests
from io import BytesIO
import os
import pytesseract
import joblib

pytesseract.pytesseract.tesseract_cmd = '/opt/homebrew/bin/tesseract'

def fetch_image(image_location):
    try:
        if image_location.startswith(('http://', 'https://')):
            context = ssl.create_default_context(cafile=certifi.where())
            response = requests.get(image_location, verify=context)
            image = Image.open(BytesIO(response.content))
        elif os.path.isfile(image_location):
            image = Image.open(image_location)
        else:
            raise ValueError("Invalid image location provided")
        return image
    except Exception as e:
        print(f"Error fetching image: {e}")
        return None

def preprocess_image(image):
    image = image.convert('L')
    image = image.filter(ImageFilter.MedianFilter())
    enhancer = ImageEnhance.Contrast(image)
    image = enhancer.enhance(2)
    return image

def extract_text(image):
    text = pytesseract.image_to_string(image)
    return text

def parse_ids(text):
    return re.findall(r'\b[A-Za-z0-9-]+\b', text)

def create_dataframe(ids):
    df = pd.DataFrame(ids, columns=['Component ID'])
    df['Count'] = 1
    df = df.groupby('Component ID').sum().reset_index()
    df['Component ID'] = df['Component ID'].astype(str)
    return df

def fetch_excel(file_path):
    try:
        xls = pd.ExcelFile(file_path)
        return xls
    except Exception as e:
        print(f"Error fetching Excel file: {e}")
        return None

def process_delta_sheet(xls):
    try:
        delta_sheet = pd.read_excel(xls, sheet_name='Delta ')
        delta_cleaned = delta_sheet.iloc[3:, :]
        delta_cleaned.columns = ['Item', 'Tally Part No.', 'Make', 'Unit', 'Quantity', 'Unnamed: 5', 'Total', 'Labour', 'Unnamed: 8', 'Unnamed: 9', 'Unnamed: 10', 'Unnamed: 11']
        delta_cleaned = delta_cleaned[['Item', 'Tally Part No.', 'Make', 'Quantity']]
        delta_cleaned = delta_cleaned.dropna(subset=['Item', 'Tally Part No.', 'Make', 'Quantity'])
        delta_cleaned['Quantity'] = pd.to_numeric(delta_cleaned['Quantity'], errors='coerce')
        component_summary = delta_cleaned.groupby('Tally Part No.')['Quantity'].sum().reset_index()
        component_summary.columns = ['Component ID', 'Count']
        return component_summary
    except Exception as e:
        print(f"Error processing sheet: {e}")
        return None

def save_to_excel(df, filename):
    try:
        df.to_excel(filename, index=False)
        print(f"DataFrame written to {filename} successfully.")
    except Exception as e:
        print(f"Error saving to Excel: {e}")

def main():
    print('Enhanced BOM Component ID Extractor with OCR and Excel Integration')
    
    image_location = '/Users/manyavalecha/Documents/bom project/BOM_image.jpeg'
    image = fetch_image(image_location)

    if image is not None:
        print("Image fetched successfully.")
        processed_image = preprocess_image(image)
        text = extract_text(processed_image)
        print("Text extracted from image:")
        print(text)

        component_ids = parse_ids(text)

        if component_ids:
            df_image = create_dataframe(component_ids)
            print("\nExtracted Component IDs from Image:")
            print(df_image)
            excel_filename_image = '/Users/manyavalecha/Downloads/component_ids_from_image.xlsx'
            save_to_excel(df_image, excel_filename_image)
        else:
            print("No component IDs found in image.")
    else:
        print("Failed to fetch image. Continuing with Excel processing...")

    file_path = '/Users/manyavalecha/Downloads/BOM- project data .xlsx'
    xls = fetch_excel(file_path)

    if xls is not None:
        print("Excel file fetched successfully.")
        component_summary = process_delta_sheet(xls)

        if component_summary is not None:
            print("\nExtracted Component IDs from Excel:")
            print(component_summary)
            excel_filename_excel = '/Users/manyavalecha/Downloads/component_ids_summary.xlsx'
            save_to_excel(component_summary, excel_filename_excel)
        else:
            print("Failed to process sheet or no component IDs found in Excel.")
    else:
        print("Failed to fetch Excel file. Exiting program.")

if __name__ == "__main__":
    main()
