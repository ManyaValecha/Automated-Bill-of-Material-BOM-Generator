import ssl
import certifi
from PIL import Image, ImageEnhance, ImageFilter
import pandas as pd
import re
import requests
from io import BytesIO
import os
import pytesseract
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

def main():
    print('Enhanced BOM Component ID Extractor with ML')
    print("Extracting and analyzing component IDs from an example BOM image...")

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
            df = create_dataframe(component_ids)
            print("\nExtracted Component IDs:")
            print(df)
            excel_filename = 'component_ids.xlsx'
            df.to_excel(excel_filename, index=False)
            print(f"\nDataFrame written to {excel_filename} successfully.")
        else:
            print("No component IDs found.")
    else:
        print("Failed to fetch image. Exiting program.")

if __name__ == "__main__":
    main()
