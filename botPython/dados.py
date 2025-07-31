import pandas as pd
from pathlib import Path
# from tkinter import filedialog

# arquivo_planilha = filedialog.askopenfilename()
arquivo_planilha = Path(r"C:\Users\novoa\OneDrive\Área de Trabalho\notas_MB\planilhas\zona_sul\escola_canadenseZS_jul25\Numeração de Boletos_Zona Sul_2025_07 2025 (1).xlsx")
arquivo_progresso = arquivo_planilha.parent / 'progresso.log'
dados = pd.read_excel(arquivo_planilha, 'dados', header=1, skipfooter=1)

dados['Aluno'] = dados['Aluno'].apply(lambda i: i.split()[0])

def encurta_nome(nome):
    nome_completo = nome.split()
    return ''.join([nome_completo[0], ' ', nome_completo[-1]])
dados['ResponsávelFinanceiro'] = dados['ResponsávelFinanceiro'].apply(encurta_nome)

dados.loc[dados['Turma'].str.contains('Y1|Y2|Year'), 'Acumulador'] = '2'
dados['Acumulador'] = dados['Acumulador'].fillna('1')

def formatar_valores(valor):
    valor_2casas = '{:0.2f}'.format(valor)
    return valor_2casas.replace('.',',')

dados['Mensalidade'] = dados['Mensalidade'].apply(formatar_valores)
dados['ValorTotal'] = dados['ValorTotal'].apply(formatar_valores)
dados['Alimentação'] = dados['Alimentação'].apply(formatar_valores)
try:
    dados['Extra'] = dados['Extra'].apply(formatar_valores)
except KeyError:
    ...
dados['Notas'] = dados['Notas'].astype(str)


if __name__ == '__main__':
    print(dados)
    # for linha in dados.itertuples():
    #     if (float(linha.Extra.replace(',','.')) != 0.0) and (linha.Extra != 'nan'):
    #         print(linha.Extra)
