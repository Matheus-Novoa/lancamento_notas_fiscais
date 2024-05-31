import pandas as pd
from pathlib import Path
from tkinter import filedialog

# arquivo_planilha = filedialog.askopenfilename()
arquivo_planilha = Path(r"C:\Users\novoa\OneDrive\Área de Trabalho\notas_MB\planilhas\zona_sul\escola_canadenseZS_mai24\Notas Maio 2024 zona sul.xlsx")
arquivo_progresso = arquivo_planilha.parent / 'progresso.log'
dados = pd.read_excel(arquivo_planilha, 'dados', header=0)

dados['Aluno'] = dados['Aluno'].apply(lambda i: i.split()[0])

def encurta_nome(nome):
    nome_completo = nome.split()
    return ''.join([nome_completo[0], ' ', nome_completo[-1]])
# dados['ResponsavelFinanceiro'] = dados['ResponsavelFinanceiro'].apply(encurta_nome)

dados.loc[dados['Turma'].isin(['Year', 'Y1']), 'Acumulador'] = '2'
dados['Acumulador'] = dados['Acumulador'].fillna('1')

dados['Mensalidade'] = dados['Mensalidade'].apply(lambda x: '{:0.2f}'.format(x))
dados['ValorTotal'] = dados['ValorTotal'].apply(lambda x: '{:0.2f}'.format(x))
dados['Alimentação'] = dados['Alimentação'].apply(lambda x: '{:0.2f}'.format(x))
dados['Extra'] = dados['Extra'].apply(lambda x: '{:0.2f}'.format(x))
dados['Nota'] = dados['Nota'].astype(str)


if __name__ == '__main__':
    print(dados.columns)
