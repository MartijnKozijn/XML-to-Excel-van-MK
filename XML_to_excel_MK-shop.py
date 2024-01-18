import csv
import requests
import pandas as pd
import html
import xml.etree.ElementTree as ET  # Voeg deze regel toe

# Functie om de XML van de opgegeven URL op te halen en te converteren naar een ElementTree-object
def get_xml_data(url):
    response = requests.get(url)
    root = ET.fromstring(response.content)
    return root

# Rest van je script...

# Functie om de gegevens naar een Excel-bestand te schrijven
def write_to_excel(root):
    # Maak een lege lijst aan om de gegevens op te slaan
    data = []

    # Loop door elk Advertisement-element in de XML
    for advertisement in root.findall('.//Advertisement'):
        # Verkrijg de benodigde informatie uit het Advertisement-element
        title = advertisement.find('Title').text

        # Decodeer HTML-entiteiten in de "Description"-tekst
        description = advertisement.find('Description').text if advertisement.find('Description') is not None else ''
        description_decoded = html.unescape(description)

        price = advertisement.find('Price').text
        afmetingen = advertisement.find('Afmetingen').text
        sku = advertisement.find('SKU').text
        website_link = advertisement.find('WebsiteLink').text

        # Combineer Image-elementen tot een enkele string (gescheiden door komma's)
        images = ';'.join([image.text for image in advertisement.findall('.//Images/Image')])

        # Voeg de naam van de categorie toe aan het werkblad
        category_name = advertisement.find('.//Category/Name').text

        # Voeg de gegevens toe aan de lijst
        data.append([title, description_decoded, price, afmetingen, sku, website_link, category_name, images])

    # Maak een DataFrame aan met de gegevens
    df = pd.DataFrame(data, columns=['Title', 'Description', 'Price', 'Afmetingen', 'SKU', 'WebsiteLink', 'CategoryName', 'Images'])

    # Sla het DataFrame op naar een Excel-bestand
    excel_file_path = 'C:\\Users\\Rik\\Desktop\\Prestashop_uitvoer\\mijn_uitvoer.xlsx'
    df.to_excel(excel_file_path, index=False)

    # Print de bestandsnaam om te controleren of het script wordt uitgevoerd
    print(f"Excel-bestand opgeslagen op: {excel_file_path}")

# URL van de XML-file
xml_url = 'https://easyads.itchconsultancy.nl/Feeds/martijnkozijn.xml'

# Haal XML-data op en schrijf naar Excel
xml_data = get_xml_data(xml_url)
write_to_excel(xml_data)
