import pandas as pd
from tkinter import Tk, filedialog
from odf.opendocument import OpenDocumentSpreadsheet
from odf.table import Table, TableRow, TableCell
from odf.text import P

# Função para criar o arquivo ODS
def create_ods(df, output_path):
    ods = OpenDocumentSpreadsheet()
    table = Table(name="Relatório")
    
    # Adicionar cabeçalhos
    headers = list(df.columns)
    header_row = TableRow()
    for header in headers:
        cell = TableCell()
        cell.addElement(P(text=header))
        header_row.addElement(cell)
    table.addElement(header_row)
    
    # Adicionar dados
    for _, row in df.iterrows():
        data_row = TableRow()
        for value in row:
            cell = TableCell()
            cell.addElement(P(text=str(value)))
            data_row.addElement(cell)
        table.addElement(data_row)
    
    ods.spreadsheet.addElement(table)
    ods.save(output_path)

# Função principal
def main():
    root = Tk()
    root.withdraw()
    
    # Selecionar arquivos
    file1 = filedialog.askopenfilename(title="Selecione o arquivo Saída com Chave de Acesso", filetypes=[("CSV files", "*.csv")])
    file2 = filedialog.askopenfilename(title="Selecione o arquivo ICMS Monofásico", filetypes=[("CSV files", "*.csv")])
    
    # Informar mês e ano
    mes_ano = input("Informe o mês e ano (MM/AAAA): ")
    mes, ano = mes_ano.split("/")
    
    # Carregar os arquivos CSV com seus respectivos delimitadores
    df1 = pd.read_csv(file1, sep=',', dtype=str, encoding='ISO-8859-1')  # Saída com Chave de Acesso
    df2 = pd.read_csv(file2, sep=';', dtype=str, encoding='ISO-8859-1')  # ICMS Monofásico

    
    # Filtrar os dados do primeiro arquivo
    df1['Data Emissão'] = pd.to_datetime(df1.iloc[:, 5], errors='coerce')
    df1 = df1[df1['Data Emissão'].dt.month == int(mes)]
    df1 = df1[df1['Data Emissão'].dt.year == int(ano)]
    
    # Preencher a coluna Inscrição Produtor
    df1['Inscrição Produtor'] = df1.iloc[:, 18].fillna(df1.iloc[:, 19])
    
    # Ordenar o segundo arquivo
    df2 = df2.sort_values(by=df2.columns[0])

    # Certifique-se de redefinir os índices antes de combinar as tabelas
    df1 = df1.reset_index(drop=True)
    df2 = df2.reset_index(drop=True)

    print("Tamanho de df1:", len(df1))
    print("Tamanho de df2:", len(df2))
    
    # Mesclar e calcular as colunas
    result = pd.DataFrame({
        "Apuração": [mes_ano] * len(df1),
        "Data Emissão": df1.iloc[:, 5].dt.strftime('%d/%m/%Y'),
        "Número": df1.iloc[:, 1],
        "Natureza": df1.iloc[:, 6],
        "Razão Social": df1.iloc[:, 16],
        "Inscrição Produtor": df1['Inscrição Produtor'],
        "Produto": df1.iloc[:, 7],
        "Quant. Total": df1.iloc[:, 9].str.replace('.', '', regex=False).str.replace(',', '.', regex=False).astype(float),
        "Aliq.": df2.iloc[:, 21].str.replace('.', '', regex=False).str.replace(',', '.', regex=False).astype(float),
    })
    result["Quant. 86%"] = result["Quant. Total"] * 0.86
    result["Aliq. Quant. 86%"] = result["Quant. 86%"] * result["Aliq."]
    result["Quant. 14%"] = result["Quant. Total"] * 0.14
    result["Aliq. Quant. 14%"] = result["Quant. 14%"] * result["Aliq."]
    
    # Selecionar local para salvar
    output_path = filedialog.asksaveasfilename(defaultextension=".ods", filetypes=[("ODS files", "*.ods")])
    create_ods(result, output_path)
    print("Arquivo gerado com sucesso!")

if __name__ == "__main__":
    main()
