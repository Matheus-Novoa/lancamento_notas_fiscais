import pandas as pd
from pathlib import Path
import pyautogui as gui


arquivo_planilha = Path(r'C:\Users\novoa\OneDrive\Área de Trabalho\notas_MB\planilhas\zona_sul\escola_canadenseZS_abr24\abril2024.xlsx')
arquivo_progresso = arquivo_planilha.parent / 'progresso.log'
planilha = pd.read_excel(arquivo_planilha, 'teste', header=0)

if arquivo_progresso.exists():
    with open(arquivo_progresso) as f:
        linha = int(f.read().split()[-1])
    planilha = planilha.iloc[linha:]

for linha in planilha.itertuples():
    if linha:
        try:
            aluno = linha.Aluno.split()[0]
            responsavel_completo = linha.Responsavel_Financeiro.split()
            responsavel = ''.join([responsavel_completo[0], ' ', responsavel_completo[-1]])
            turma = linha.Turma
            if 'Year' in turma or 'Y1' in turma:
                acumulador = '2'
            else: 
                acumulador = '1'
            valor = str('{:.2f}'.format(linha.Mensalidade)).replace('.',',')
            refeicao = str(linha.Alimentação).replace('.',',')
            valor_total = str('{:.2f}'.format(linha.Valor_Total)).replace('.',',')
            extra = str('{:.2f}'.format(linha.Extra)).replace('.',',') if linha.Extra else None
            numero_nota = str(linha.Nota)
            print('aluno',aluno)
            print('responsavel',responsavel)
            print('turma', turma)
            print('acumulador', acumulador)
            print('valor',valor)
            print('refeição',refeicao)
            print('valor total', valor_total)
            if extra:
                print('valor extra', extra)
            print('Nota', numero_nota)
            print()

            resposta = gui.confirm(title='Os dados foram preenchidos corretamente?', buttons=['Continuar', 'Pausa']) 
            if resposta == 'Pausa':
                with open(arquivo_progresso, 'w') as f:
                    f.write(f'Pausa linha {linha.Index + 1}')
                break
        except:
            with open(arquivo_progresso, 'w') as f:
                f.write(f'Erro linha {linha.Index}')
                raise
    else:
        print('linha vazia')
        break