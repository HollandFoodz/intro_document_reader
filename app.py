import pandas as pd
import os
from lxml import etree


class Article():

    def __init__(self):
        self.d = {}

def fetch_data():
    for root, dir, files in os.walk(INTRO_DIR):
        for f in files:
            filename, ext = os.path.splitext(f)
            if ext == '.xlsx':
                df = pd.read_excel(os.path.join(root, f), header=None)
                article = Article()
                article.article_number = df[1][7]
                article.article_desc = df[7][7]
                article.d['Legal decleration'] = df[7][8]
                article.d['D.C. Location'] = df[7][9]
                article.d['Brand'] = df[1][8]

                article.d['CP EAN Code'] = df[2][11]
                article.d['CP Net weight (gr)'] = df[6][11]
                article.d['CP Width (mm)'] = df[8][11]
                article.d['CP Height (mm)'] = df[10][11]
                article.d['CP Depth (mm)'] = df[12][11]

                article.d['DU EAN Code'] = df[2][12]
                article.d['DU Gross weight (gr)'] = df[6][12]
                article.d['DU Height (mm)'] = df[8][12]
                article.d['DU Width (mm)'] = df[10][12]
                article.d['DU Depth (mm)'] = df[12][12]

                article.d['Quantity CE per case'] = df[4][13]
                article.d['Quantity DU per OP'] = df[10][13]

                article.d['OP EAN Code'] = df[2][14]
                article.d['OP Gross weight (kg)'] = df[6][14]
                article.d['OP Height (mm)'] = df[8][14]
                article.d['OP Width (mm)'] = df[10][14]
                article.d['OP Depth (mm)'] = df[12][14]
                article.d['OP Minimum order quantity'] = df[4][15]

                article.d['Pallet Format'] = df[2][16]
                article.d['Pallet Amount per layer'] = df[6][16]
                article.d['Pallyet Amount layers'] = df[9][16]
                article.d['Pallet Gross weight (kg)'] = df[6][17]
                article.d['Pallet Total cases'] = df[12][16]
                article.d['Pallet Height (cm)'] = df[12][17]

                article.d['Ingredients (NL)'] = df[2][19]
                article.d['Allergens (Yes/No)'] = df[11][19]
                article.d['GMO-free (Yes/No)'] = df[11][20]
                article.d['Health claim (NL)'] = df[1][23]
                article.d['Other claims'] = df[1][26]
                article.d['Shelf life (days)'] = df[1][27]
                article.d['Minimum shelf life (days)'] = df[1][30]
                article.d['Storage conditions'] = df[1][31]

                article.d['Energy (kcal)'] = df[1][34]
                article.d['Energy (kJ)'] = df[1][35]
                article.d['Fat (gr)'] = df[1][36]
                article.d['Fat saturated (gr)'] = df[1][37]
                article.d['Carbohydrates (gr)'] = df[1][38]
                article.d['Carbohydrates sugars (gr)'] = df[1][39]
                article.d['Fibres (gr)'] = df[1][40]
                article.d['Protein (gr)'] = df[1][41]
                article.d['Salt (gr)'] = df[1][42]

                article.d['Tax percentage'] = df[6][39]
                article.d['Country of origin'] = df[6][40]

                artikel = etree.SubElement(artikelen, 'ARTIKEL')
                etree.SubElement(artikel, "ART_NUMMER").text = str(article.article_number)
                etree.SubElement(artikel, "ART_OMSCHRIJVING").text = str(article.article_desc)

                vrije_rubrieken = etree.SubElement(artikel, 'ART_VRIJERUBRIEKEN')
                for key, value in article.d.items():
                    vrije_rubriek = etree.SubElement(artikel, 'ART_VRIJERUBRIEK')
                    etree.SubElement(vrije_rubriek, "ART_VRIJERUBRIEK_NAAM").text = str(key).strip()
                    etree.SubElement(vrije_rubriek, "ART_VRIJERUBRIEK_WAARDE").text = str(value).strip()

if __name__ == '__main__':
    KING_ARTIKELEN = etree.Element('KING_ARTIKELEN')
    doc = etree.ElementTree(KING_ARTIKELEN)
    artikelen = etree.SubElement(KING_ARTIKELEN, 'ARTIKELEN')
    INTRO_DIR = './intro_documents'

    if not os.path.exists(INTRO_DIR):
        os.makedirs(INTRO_DIR)

    fetch_data()
    doc.write('output.xml', xml_declaration=True, encoding='utf-8', pretty_print=True)
