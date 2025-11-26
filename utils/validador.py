import pandas as pd
from validate_docbr import CPF, CNPJ

def col_excel(idx):
    result = ""
    while idx >= 0:
        result = chr(idx % 26 + ord('A')) + result
        idx = idx // 26 - 1
        if idx < 0:
            break
    return result

def normaliza_doc(valor):
    """Preenche zeros à esquerda em CPF/CNPJ para forçar 11/14 dígitos"""
    valor_str = str(valor).strip().replace(".", "").replace("-", "").replace("/", "")
    if valor_str.isdigit():
        if len(valor_str) <= 11:
            return valor_str.zfill(11)
        elif len(valor_str) <= 14:
            return valor_str.zfill(14)
    return valor_str  # Caso esteja muito errado, só retorna como veio

def validar_planilha(df):
    erros = []

    colunas_obrigatorias = [
        "NUMERO DO CONTRATO",
        "CPF_CNPJ",
        "DATA DO REGISTRO",
        "CPF_LOGIN",
        "SENHA_LOGIN",
        "ARTCREA",
        "FAZENDA"
    ]

    for col in colunas_obrigatorias:
        if col not in df.columns:
            erros.append(f"Coluna obrigatória ausente: '{col}'")

    checar_todas = ["NUMERO DO CONTRATO", "CPF_CNPJ", "DATA DO REGISTRO", "FAZENDA"]
    for col in checar_todas:
        if col in df.columns:
            for i, valor in enumerate(df[col]):
                if pd.isna(valor) or str(valor).strip() == "":
                    col_letra = col_excel(df.columns.get_loc(col))
                    erros.append(f"Célula vazia na coluna '{col}' ({col_letra}{i+2})")

    for col in ["CPF_LOGIN", "SENHA_LOGIN", "ARTCREA"]:
        if col in df.columns:
            if pd.isna(df[col].iloc[0]) or str(df[col].iloc[0]).strip() == "":
                col_letra = col_excel(df.columns.get_loc(col))
                erros.append(f"Célula vazia na coluna '{col}' ({col_letra}2)")

    # Validação básica de CPF/CNPJ com zeros à esquerda corrigidos!
    if "CPF_CNPJ" in df.columns:
        for i, valor in enumerate(df["CPF_CNPJ"]):
            doc_corrigido = normaliza_doc(valor)
            col_letra = col_excel(df.columns.get_loc("CPF_CNPJ"))
            if len(doc_corrigido) == 11:
                if not CPF().validate(doc_corrigido):
                    erros.append(f"CPF inválido em {col_letra}{i+2}: {valor}")
            elif len(doc_corrigido) == 14:
                if not CNPJ().validate(doc_corrigido):
                    erros.append(f"CNPJ inválido em {col_letra}{i+2}: {valor}")
            else:
                erros.append(f"Documento inválido em {col_letra}{i+2}: {valor}")

    return erros
