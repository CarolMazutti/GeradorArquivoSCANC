import pandas as pd
import tkinter as tk
from tkinter import filedialog
from tkinter.simpledialog import askstring

# Função para selecionar arquivo
def selecionar_arquivo(nome_arquivo):
    root = tk.Tk()
    root.withdraw()
    arquivo = filedialog.askopenfilename(title=f"Selecione o arquivo {nome_arquivo}")
    return arquivo

# Função para pedir o mês e o ano
def pedir_mes_ano():
    root = tk.Tk()
    root.withdraw()
    mes_ano = askstring("Mês e Ano", "Informe o mês e ano (MM/YYYY):")
    return mes_ano

# Função para salvar o arquivo
def salvar_arquivo():
    root = tk.Tk()
    root.withdraw()
    arquivo_salvar = filedialog.asksaveasfilename(defaultextension=".ods", filetypes=[("ODS files", "*.ods")])
    return arquivo_salvar

# Função para carregar os arquivos
def carregar_arquivos():
    arquivo1 = selecionar_arquivo("Saída com Chave de Acesso")
    arquivo2 = selecionar_arquivo("ICMS Monofásico")

    # Carregar os arquivos corretamente
    df1 = pd.read_excel(arquivo1, engine="xlrd")  # Arquivo Excel (.xls)
    df2 = pd.read_csv(arquivo2, delimiter=";", encoding="ISO-8859-1")  # Arquivo CSV delimitado por ponto e vírgula

    return df1, df2

def processar_dados(df1, df2, mes_ano):
    # Ordenar os DataFrames
    if "Número" in df1.columns:
        df1 = df1.sort_values(by=["Número"])
    if "Número Lançamento" in df2.columns:
        df2 = df2.sort_values(by=["Número Lançamento"])

    # Garantir que as colunas sejam strings antes de usar .str.replace
    if "Quantidade" in df1.columns:
        df1["Quantidade"] = (
            df1["Quantidade"]
            .astype(str)  # Converter para string
            .str.replace(".", "", regex=False)
            .str.replace(",", ".", regex=False)
            .astype(float)  # Converter de volta para float
        )

    if "ALIQADREMICMSRETIDAANT" in df2.columns:
        df2["ALIQADREMICMSRETIDAANT"] = (
            df2["ALIQADREMICMSRETIDAANT"]
            .astype(str)  # Converter para string
            .str.replace(".", "", regex=False)
            .str.replace(",", ".", regex=False)
            .astype(float)  # Converter de volta para float
        )

    # Tratar valores inconsistentes na coluna de datas
    if "Data Emissão" in df1.columns:
        df1["Data Emissão"] = pd.to_datetime(df1["Data Emissão"], format='%d/%m/%Y', errors='coerce')
    if "DATALCTOFIS" in df2.columns:
        df2["DATALCTOFIS"] = pd.to_datetime(df2["DATALCTOFIS"], format='%d/%m/%Y', errors='coerce')

    # Remover linhas com valores inválidos de data
    df1 = df1.dropna(subset=["Data Emissão"])
    df2 = df2.dropna(subset=["DATALCTOFIS"])

    # Filtrar por mês e ano
    mes, ano = map(int, mes_ano.split("/"))
    df1_filtrado = df1[(df1["Data Emissão"].dt.month == mes) & (df1["Data Emissão"].dt.year == ano)]
    df2_filtrado = df2[(df2["DATALCTOFIS"].dt.month == mes) & (df2["DATALCTOFIS"].dt.year == ano)]

    # Criar o DataFrame final
    resultado = pd.DataFrame()
    resultado["Apuração"] = mes_ano
    resultado["Data Emissão"] = df1_filtrado["Data Emissão"].dt.strftime('%d/%m/%Y')
    resultado["Número"] = df1_filtrado["Número"]
    resultado["Natureza"] = df1_filtrado["Natureza"]
    resultado["Razão Social"] = df1_filtrado["Razão Social"]
    resultado["Produto"] = df1_filtrado["Produto"]
    resultado["Quant. Total"] = df1_filtrado["Quantidade"]

    # Adicionar a coluna Aliq. com base no segundo arquivo
    if not df2_filtrado.empty and "ALIQADREMICMSRETIDAANT" in df2_filtrado.columns:
        resultado["Aliq."] = df2_filtrado["ALIQADREMICMSRETIDAANT"].values[:len(resultado)]
    else:
        resultado["Aliq."] = 0  # Valor padrão caso df2_filtrado esteja vazio

    # Calcular as colunas Aliq. Quant. 86% e Aliq. Quant. 14%
    resultado["Aliq. Quant. 86%"] = resultado["Quant. Total"] * 0.86 * resultado["Aliq."]
    resultado["Aliq. Quant. 14%"] = resultado["Quant. Total"] * 0.14 * resultado["Aliq."]

    return resultado

# Função principal
def main():
    df1, df2 = carregar_arquivos()
    mes_ano = pedir_mes_ano()
    resultado = processar_dados(df1, df2, mes_ano)
    arquivo_salvar = salvar_arquivo()

    # Salvar o resultado em um arquivo ODS
    resultado.to_excel(arquivo_salvar, index=False, engine="odf")
    print(f"Arquivo gerado com sucesso: {arquivo_salvar}")

# Executar o script
if __name__ == "__main__":
    main()
