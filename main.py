import xmltodict
import os
import json
import pandas as pd


def get_file(path_in, files, table_values):
    with open(f"{path_in}{files}", "rb") as xml_file:
        dic_files = xmltodict.parse(xml_file)
        try:
            if "NFe" in dic_files:
                info_nf = dic_files["NFe"]["infNFe"]
            else:
                info_nf = dic_files["nfeProc"]["NFe"]["infNFe"]
            empresa_emissora = info_nf["emit"]["xNome"]
            numero_nota = info_nf["@Id"]
            nome_cliente = info_nf["dest"]["xNome"]
            endereco = info_nf["dest"]["enderDest"]
            logradouro = endereco["xLgr"]
            numero_log = endereco["nro"]
            if "xCpl" in endereco:
                complemento = endereco["xCpl"]
            else:
                complemento = "Não informado"
            bairro = endereco["xBairro"]
            municipio = endereco["xMun"]
            uf = endereco["UF"]
            cep = endereco["CEP"]
            pais = endereco["xPais"]
            if "vol" in info_nf["transp"]:
                peso = info_nf["transp"]["vol"]["pesoB"]
            else:
                peso = "Não informado"
            table_values.append(
                [
                    numero_nota,
                    empresa_emissora,
                    nome_cliente,
                    peso,
                    logradouro,
                    numero_log,
                    complemento,
                    bairro,
                    municipio,
                    uf,
                    cep,
                    pais,
                ]
            )
            print(table_values)
        except Exception as e:
            print(e)


table_columns = [
    "numero_nota",
    "empresa_emissora",
    "nome_cliente",
    "peso",
    "logradouro",
    "numero_log",
    "complemento",
    "bairro",
    "municipio",
    "uf",
    "cep",
    "pais",
]
table_values = []

path_in = "xml/"
path_out = "xlsx/"
file_out = "notasfiscais.xlsx"

list_files = os.listdir(path_in)

for files in list_files:
    get_file(path_in, files, table_values)

table = pd.DataFrame(columns=table_columns, data=table_values)
table.to_excel(f"{path_out}{file_out}", index=False)
