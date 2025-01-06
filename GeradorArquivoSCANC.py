import pandas as pd
from tkinter import Tk, filedialog, simpledialog, messagebox
import os

def main():
    # Ocultar a janela principal do Tkinter
    root = Tk()
    root.withdraw()

    # Solicitar ao usuário os dois arquivos
    messagebox.showinfo("Seleção de Arquivo", "Selecione o arquivo Saída com Chave de Acesso (planilha Excel).")
    saida_chave_path = filedialog.askopenfilename(
        title="Selecione o arquivo Saída com Chave de Acesso",
        filetypes=[("Arquivos Excel", "*.xls *.xlsx")]
    )
    
    messagebox.showinfo("Seleção de Arquivo", "Selecione o arquivo ICMS Monofásico (arquivo CSV).")
    icms_mono_path = filedialog.askopenfilename(
        title="Selecione o arquivo ICMS Monofásico",
        filetypes=[("Arquivos CSV", "*.csv")]
    )

    if not saida_chave_path or not icms_mono_path:
        messagebox.showerror("Erro", "Você deve selecionar ambos os arquivos.")
        return

    # Solicitar ao usuário o mês e ano
    mes = simpledialog.askinteger("Entrada", "Informe o mês (1-12):")
    ano = simpledialog.askinteger("Entrada", "Informe o ano (ex: 2023):")
    
    if not mes or not ano:
        messagebox.showerror("Erro", "Você deve informar o mês e o ano.")
        return

    # Carregar os arquivos
    saida_chave = pd.read_excel(saida_chave_path)
    icms_mono = pd.read_csv(icms_mono_path, delimiter=';', encoding='latin-1')

    print(icms_mono.head())  # Exibe as primeiras linhas do DataFrame
    print(icms_mono.columns)  # Exibe os nomes das colunas para verificar posições e nomes

    # Ordenar os dados
    saida_chave.sort_values(by=saida_chave.columns[1], inplace=True)  # Ordenar pela 2ª coluna (Número)
    icms_mono.sort_values(by=icms_mono.columns[0], inplace=True)  # Ordenar pela 1ª coluna (Número Lançamento)

    # Filtrar dados pelo mês e ano
    saida_chave['Data'] = pd.to_datetime(saida_chave.iloc[:, 5], errors='coerce')  # Coluna 6 para datetime
    icms_mono['Data'] = pd.to_datetime(icms_mono.iloc[:, 2], errors='coerce')  # Coluna 3 para datetime

    saida_chave = saida_chave[(saida_chave['Data'].dt.month == mes) & (saida_chave['Data'].dt.year == ano)]
    icms_mono = icms_mono[(icms_mono['Data'].dt.month == mes) & (icms_mono['Data'].dt.year == ano)]

    # Montar o DataFrame final
    resultado = pd.DataFrame()
    resultado['Apuração'] = f"{mes:02d}/{ano}"  # Mês e ano
    resultado['Data Emissão'] = saida_chave['Data'].dt.strftime('%d/%m/%Y')  # Data formatada
    resultado['Número'] = saida_chave.iloc[:, 1]  # Número
    resultado['Natureza'] = saida_chave.iloc[:, 6]  # Natureza
    resultado['Razão Social'] = saida_chave.iloc[:, 9]  # Razão Social

    # Inscrição Produtor (coluna 19 ou 20)
    inscricao_produtor = saida_chave.iloc[:, 18].fillna(saida_chave.iloc[:, 19])
    resultado['Inscrição Produtor'] = inscricao_produtor

    resultado['Produto'] = saida_chave.iloc[:, 7]  # Produto
    resultado['Quant. Total'] = saida_chave.iloc[:, 9]  # Quant. Total

    # Adicionar Aliq (coluna 22 do ICMS Monofásico)
    resultado['Aliq.'] = icms_mono.iloc[:, 21].values  # Aliq.

    # Colunas calculadas
    resultado['Quant. 86%'] = resultado['Quant. Total'] * 0.86
    resultado['Aliq. Quant. 86%'] = resultado['Quant. 86%'] * resultado['Aliq.']
    resultado['Quant. 14%'] = resultado['Quant. Total'] * 0.14
    resultado['Aliq. Quant. 14%'] = resultado['Quant. 14%'] * resultado['Aliq.']

    # Solicitar o local e nome para salvar o arquivo
    save_path = filedialog.asksaveasfilename(
        title="Salvar arquivo",
        defaultextension=".xlsx",
        filetypes=[("Arquivo Excel", "*.xlsx")]
    )

    if save_path:
        resultado.to_excel(save_path, index=False, engine='openpyxl')
        messagebox.showinfo("Sucesso", f"Arquivo salvo com sucesso em: {save_path}")
    else:
        messagebox.showwarning("Cancelado", "Operação cancelada pelo usuário.")

if __name__ == "__main__":
    main()
