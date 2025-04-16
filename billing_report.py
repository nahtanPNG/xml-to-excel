import os
import sys
import xml.etree.ElementTree as ET

import pandas as pd

NAMESPACES = {'nfe': 'http://www.portalfiscal.inf.br/nfe'}

TYPE_MAP = {'0': 'Entrada', '1': 'Saída'}

FINALITY_MAP = {
    '1': 'Normal',
    '2': 'Complementar',
    '3': 'Ajuste',
    '4': 'Devolução'
}


def parse_text(element, path):
    found = element.find(path, NAMESPACES)
    return found.text if found is not None else None


def extract_cnpj_or_cpf(parent):
    cpf = parse_text(parent, 'nfe:CPF')
    cnpj = parse_text(parent, 'nfe:CNPJ')
    return cpf or cnpj


def process_xml_file(file_path):
    tree = ET.parse(file_path)
    root = tree.getroot()

    return {
        'Número': parse_text(root, './/nfe:nNF'),
        'Serie': parse_text(root, './/nfe:serie'),
        'Tipo': TYPE_MAP.get(parse_text(root, './/nfe:tpNF'), 'Desconhecido'),
        'Finalidade': FINALITY_MAP.get(parse_text(root, './/nfe:finNFe'), 'Desconhecida'),
        'Natureza da Operação': parse_text(root, './/nfe:natOp'),
        'Dt_Emissão': parse_text(root, './/nfe:dhEmi'),
        'CPF_CNPJ_Emit': extract_cnpj_or_cpf(root.find('.//nfe:emit', NAMESPACES)),
        'Rz_Emit': parse_text(root, './/nfe:emit/nfe:xNome'),
        'IE_Emit': parse_text(root, './/nfe:emit/nfe:IE'),
        'CPF_CNPJ_Dest': extract_cnpj_or_cpf(root.find('.//nfe:dest', NAMESPACES)),
        'Rz_Dest': parse_text(root, './/nfe:dest/nfe:xNome'),
        'IE_Dest': parse_text(root, './/nfe:dest/nfe:indIEDest'),
        'UF_Dest': parse_text(root, './/nfe:dest/nfe:enderDest/nfe:UF'),
        'Valor': float(parse_text(root, './/nfe:vNF') or 0),
        'Chave': parse_text(root, './/nfe:chNFe'),
        'Status': parse_text(root, './/nfe:xMotivo'),
    }


def process_folder(folder_path):
    all_data = []

    for filename in os.listdir(folder_path):
        if filename.endswith('.xml'):
            file_path = os.path.join(folder_path, filename)
            print(f'Processing: {file_path}')
            try:
                row = process_xml_file(file_path)
                all_data.append(row)
            except Exception as e:
                print(f'Error processing {filename}: {e}')

    return all_data


def save_to_excel(data, output_path):
    df = pd.DataFrame(data)

    with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
        df.to_excel(writer, sheet_name='Notas Fiscais', index=False)

        summary = df.groupby('Tipo')['Valor'].sum().reset_index()
        summary.to_excel(writer, sheet_name='Resumo', index=False)


def main():
    if len(sys.argv) < 2:
        print("Uso: python billing_report.py <caminho_da_pasta>")
        return

    folder_path = sys.argv[1]
    data = process_folder(folder_path)

    if data:
        output_file = os.path.join(folder_path, 'notas_fiscais.xlsx')
        save_to_excel(data, output_file)
        print(f'\nArquivo salvo em: {output_file}')
    else:
        print("Nenhum dado encontrado.")


if __name__ == '__main__':
    main()
