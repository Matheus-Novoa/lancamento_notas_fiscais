# -*- coding: utf-8 -*-
import pyautogui as gui
import pyperclip as clip
import openpyxl
import pickle
from pathlib import Path
import time


arquivo_planilha = r'escola_canadenseZS_mar24\Numeração de Boletos_Zona Sul_03 2024.xlsx'
planilha = openpyxl.load_workbook(arquivo_planilha)['Abril2024_valores']
arquivo_salvo = Path('escola_canadenseZS_mar24/dados_salvos.pkl')
data_emissao = '29022024'

def percorrer_planilha(dados):
    yield from dados

if not arquivo_salvo.exists():
    linhas = [linha for linha in planilha.iter_rows(min_row=2, values_only=True)]
else:
    with open(arquivo_salvo, 'rb') as f:
        linhas = pickle.load(f)

dados_restantes = percorrer_planilha(linhas)

gui.PAUSE = 0.3

posicao_campo_acumulador = [765, 264]
posicao_campo_emissao = [770, 243]
posicao_aba_itens = [345, 171]
posicao_campo_codigo = [262, 240]

posicao_aba_contabilidade = [445, 169]
posicao_campo_valor_historico = [794, 233]
posicao_botao_novo = [1003, 619]

posicao_botao_gravar = [378, 674]
posicao_botao_proximo = [689, 169]
posicao_aba_servicos = [232, 170]
posicao_botao_yes = [551, 647]

# aluno = linha_planilha.split('\t')[0].split()[0]
# responsavel_completo = linha_planilha.split('\t')[1].split()
# responsavel = ''.join([responsavel_completo[0], ' ', responsavel_completo[-1]])
# valor = linha_planilha.split('\t')[4]
# refeicao = linha_planilha.split('\t')[6]

gui.hotkey('win', 't')
gui.press('right', presses=5)
gui.press(['up', 'right', 'enter'])

for linha in dados_restantes:
    if linha[0] is not None:
        aluno = linha[0].split()[0]
        responsavel_completo = linha[6].split()
        responsavel = ''.join([responsavel_completo[0], ' ', responsavel_completo[-1]])
        turma = linha[1]
        acumulador = '2' if 'Year' in turma else '1'
        valor = str('{:.2f}'.format(linha[3])).replace('.',',')
        refeicao = str(linha[4]).replace('.',',')
        valor_total = str('{:.2f}'.format(linha[2])).replace('.',',')
        numero_nota = str(linha[-1])
        # print('aluno',aluno)
        # print('responsavel',responsavel)
        # print('turma', turma)
        # print('acumulador', acumulador)
        # print('valor',valor)
        # print('refeição',refeicao)
        # print('valor total', valor_total)
        # print('Nota', numero_nota)
        # print()
    else:
        break
    
    gui.click(x=714, y=255) # ativar edição
    # gui.click(*posicao_campo_emissao)
    # gui.write(data_emissao)
    # # gui.press('enter', presses=2)
    # gui.click(*posicao_campo_acumulador)
    # gui.write(acumulador)
    # gui.press('enter')
    gui.click(*posicao_aba_itens)
    gui.press('enter')
    gui.write('1')
    gui.press('enter')
    gui.write('1')
    gui.press('enter')
    gui.write(valor_total)
    gui.press('enter')
    
    # resposta = gui.confirm(title='Os dados foram preenchidos corretamente?', buttons=['Continuar', 'Interromper'])

    # if resposta == 'Interromper':
    #     break
    
    gui.click(*posicao_aba_contabilidade)
    gui.click(*posicao_campo_valor_historico)
    gui.write(valor)
    gui.press('tab', presses=2)
    gui.press('end')
    vlr_provisao = f'nf 2024/{numero_nota} {responsavel} aluno {aluno}'
    clip.copy(vlr_provisao)
    # gui.write(vlr_provisao)
    gui.hotkey('ctrl', 'v')
    gui.press('pagedown')
    gui.press('end')
    vlr_devido = f'issqn cf {vlr_provisao}'
    clip.copy(vlr_devido)
    # gui.write(vlr_devido)
    gui.hotkey('ctrl', 'v')

    if float(refeicao.replace(',','.')) != 0.0:
        gui.click(*posicao_botao_novo)
        gui.write('100')
        gui.press('enter')
        gui.write('3103')
        # gui.write('3004')
        gui.press('enter')
        gui.write(refeicao)
        gui.press('tab')
        gui.write('8')
        gui.press(['tab', 'end', ' '])
        vlr_provisao_refeicao = f'refeição cf {vlr_provisao}'
        clip.copy(vlr_provisao_refeicao)
        # gui.write(vlr_provisao_refeicao)
        gui.hotkey('ctrl', 'v')
        gui.press('enter')

    resposta = gui.confirm(title='Os dados foram preenchidos corretamente?', buttons=['Continuar', 'Interromper'])

    if resposta == 'Interromper':
        break

    gui.click(*posicao_botao_gravar)
    time.sleep(3)
    # gui.click(*posicao_botao_gravar)
    # time.sleep(4.5)
    # # gui.click(*posicao_aba_servicos)
    gui.click(*posicao_botao_proximo)
    time.sleep(1.5)

with open(arquivo_salvo, 'wb') as f:
    print('Salvando progresso')
    pickle.dump(list(dados_restantes), f)
