from docxtpl import DocxTemplate
import pandas as pd
import os

tabela = pd.read_excel("Termos Responsabilidade/Informações-Recebimento.xlsx")

for linha in tabela.index:
    documento = DocxTemplate("Termos Responsabilidade/TermoResponsabilidadeTemplate.docx")

    nome = tabela.loc[linha, "Nome"]
    nomeDocumento = tabela.loc[linha, "NomeDocumento"]
    equipamento = tabela.loc[linha, "Equipamento"]
    patrimonio = tabela.loc[linha, "Patrimonio"]
    serial = tabela.loc[linha, "Serial"]
    modelo = tabela.loc[linha, "Modelo"]

    referencias = {
        "nome": nome,
        "equipamento": equipamento,
        "patrimonio": patrimonio,
        "serial": serial,
        "modelo": modelo,
    }

    nameDocument = (f"Termos Responsabilidade/Termos-Docx/{equipamento}-{patrimonio}-{nomeDocumento}.docx")

    documento.render(referencias)

    documento.save(nameDocument)

    os.system (f'soffice --headless --convert-to pdf "{nameDocument}" --outdir "./Termos Responsabilidade/Termos-PDF"')
    