import os
import sys
import xml.etree.ElementTree as ET

import pandas as pd

namespaces = {
    'nfe': 'http://www.portalfiscal.inf.br/nfe'
}

folder_path = sys.argv[1]

with pd.ExcelWriter(f'{folder_path}/output.xlsx', 'xlsxwriter') as writer:
    for filename in os.listdir(folder_path):
        if filename.endswith('.xml'):
            file_path = os.path.join(folder_path, filename)
            print("\n")
            print(f'Processing file: {file_path}')

            # Parsing the file
            tree = ET.parse(f'{folder_path}/{filename}')
            root = tree.getroot()
            data = []

            # Get product information
            for det in root.findall('.//nfe:det', namespaces):
                prod = det.find('nfe:prod', namespaces)
                if prod is not None:
                    row = {
                        'Product': prod.find('nfe:xProd', namespaces).text,
                        'Quantity': prod.find('nfe:qCom', namespaces).text,
                        'Unit Value': prod.find('nfe:vUnCom', namespaces).text,
                        'Total Value': prod.find('nfe:vProd', namespaces).text
                    }
                    data.append(row)

            sheet_name = filename
            df = pd.DataFrame(data)
            df.to_excel(writer, sheet_name=sheet_name)
