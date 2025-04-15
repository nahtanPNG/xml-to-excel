# TODO:
# Ler todas as notas fiscais de uma pasta
# Identificar notas de entrada/devolução e saída -> (1 ou 2 - entrada) (5 ou 6 - saida)
# Retornar o valor da nota de saida - valores das notas de devolução

import os
import sys
import xml.etree.ElementTree as ET

import pandas as pd

namespaces = {
    'nfe': 'http://www.portalfiscal.inf.br/nfe'
}


def main():
    folder_path = sys.argv[1]
    all_data = []

    for filename in os.listdir(folder_path):
        if filename.endswith('.xml'):
            file_path = os.path.join(folder_path, filename)
            print("\n")
            print(f'Processing file: {file_path}')

            # Parsing the file
            tree = ET.parse(f'{folder_path}/{filename}')
            root = tree.getroot()
            data = []

            # Get the number and serie
            nNF = root.find('.//nfe:nNF', namespaces).text
            serie = root.find('.//nfe:serie', namespaces).text

            tpNF = root.find('.//nfe:tpNF', namespaces).text

            if tpNF == '0':
                tpNF = "Entrada"
            elif tpNF == '1':
                tpNF = "Saida"

            # Get the finalidade e natureza da operação
            finNFe = root.find('.//nfe:finNFe', namespaces)
            if finNFe.text == "1":
                finNFe = "Normal"
            elif finNFe.text == "2":
                finNFe = "Complementar"
            elif finNFe.text == "3":
                finNFe = "Ajuste"
            else:
                finNFe = "Devolução"

            natOp = root.find('.//nfe:natOp', namespaces).text

            # Get the date
            dhEmi = root.find('.//nfe:dhEmi', namespaces).text

            # Get the emits
            emit = root.find('.//nfe:emit', namespaces)
            cpf = emit.find('nfe:CPF', namespaces)
            cnpj = emit.find('nfe:CNPJ', namespaces)

            if cpf is not None:
                emit_CNPJ_or_CPF = cpf.text
            elif cnpj is not None:
                emit_CNPJ_or_CPF = cnpj.text
            else:
                emit_CNPJ_or_CPF = None

            emit_xNome = root.find('.//nfe:emit/nfe:xNome', namespaces).text
            emit_IE = root.find('.//nfe:emit/nfe:IE', namespaces).text

            # Get the dest
            dest = root.find('.//nfe:dest', namespaces)
            cpf = dest.find('nfe:CPF', namespaces)
            cnpj = dest.find('nfe:CNPJ', namespaces)

            if cpf is not None:
                dest_CNPJ_or_CPF = cpf.text
            elif cnpj is not None:
                dest_CNPJ_or_CPF = cnpj.text
            else:
                dest_CNPJ_or_CPF = None

            dest_xNome = root.find('.//nfe:dest/nfe:xNome', namespaces).text
            dest_IE = root.find('.//nfe:dest/nfe:indIEDest', namespaces).text
            dest_UF = root.find('.//nfe:dest/nfe:enderDest/nfe:UF', namespaces).text

            # Get values
            vNF = root.find('.//nfe:vNF', namespaces).text

            # Get key
            chNFe = root.find('.//nfe:chNFe', namespaces).text

            # Status
            xMotivo = root.find('.//nfe:xMotivo', namespaces).text

            row = {
                'Número': nNF,
                'Serie': serie,
                'Tipo': tpNF,
                'Finalidade': finNFe,
                'Natureza da Operação': natOp,
                'Dt_Emissão': dhEmi,
                'CPF_CNPJ_Emit': emit_CNPJ_or_CPF,
                'Rz_Emit': emit_xNome,
                'IE_Emit': emit_IE,
                'CPF_CNPJ_Dest': dest_CNPJ_or_CPF,
                'Rz_Dest': dest_xNome,
                'IE_Dest': dest_IE,
                'UF_Dest': dest_UF,
                'Valor': vNF,
                'Chave': chNFe,
                'Status': xMotivo
            }
            all_data.append(row)

    if all_data:
        with pd.ExcelWriter(f'{folder_path}/teste.xlsx', engine='xlsxwriter') as writer:
            pd.DataFrame(all_data).to_excel(writer, sheet_name='Notas Fiscais', index=False)

            # Add summary sheet for your business logic
            df = pd.DataFrame(all_data)
            summary = df.groupby('Tipo')['Valor'].sum().reset_index()
            summary.to_excel(writer, sheet_name='Resumo', index=False)

if __name__ == '__main__':
    main()

main()
