import pandas as pd
from botcity.core import DesktopBot
from botcity.maestro import *
from pathlib import Path
import pygetwindow as gw
import pyautogui as gui
from dados import *
BotMaestroSDK.RAISE_NOT_CONNECTED = False

def main():
    bot = DesktopBot()

    data_emissao = '31052024'
    codigo_refeicao = '3103' # zona sul
    # codigo_refeicao = '3004' # zona norte
    codigo_extra = '3609'

    if arquivo_progresso.exists():
        with open(arquivo_progresso) as f:
            linha = int(f.read().split()[-1])
        dados = dados.iloc[linha:]

    titulo_janela = "Lista de Programas"
    # Obtém todas as janelas com o título especificado
    janelas = gw.getWindowsWithTitle(titulo_janela)
    # Verifica se foi encontrada alguma janela com o título especificado
    if janelas:
        # Seleciona a primeira janela encontrada (você pode iterar sobre a lista para selecionar a janela desejada)
        janela = janelas[0] 
        # Foca na janela (torna-a ativa)
        janela.activate()
    else:
        print("Nenhuma janela encontrada com o título especificado.")
    bot.wait(1000)
    gui.hotkey('alt','esc')

    try:
        for linha in dados.itertuples():
            if not bot.find("ativar_edicao", matching=0.97, waiting_time=10000):
                not_found("ativar_edicao")
            bot.click()

            if not bot.find("campo_emissao", matching=0.97, waiting_time=10000):
                not_found("campo_emissao")
            bot.click_relative(104, 7, wait_after=1500, clicks=3)
            # bot.type_keys(["ctrl", "type_left"])
            # bot.type_keys(["ctrl", "shift", "type_right"])
            bot.type_key(data_emissao)

            if not bot.find("campo_acumulador", matching=0.97, waiting_time=10000):
                not_found("campo_acumulador")
            bot.click_relative(106, 5)
            bot.type_key(linha.Acumulador)
            bot.enter()
            
            if not bot.find("aba_itens", matching=0.97, waiting_time=10000):
                not_found("aba_itens")
            bot.click()
            bot.enter()
            bot.type_key('1')
            bot.enter()
            bot.type_key('1')
            bot.enter()
            bot.type_key(linha.ValorTotal)
            bot.enter()

            if not bot.find("aba_contabilidade", matching=0.97, waiting_time=10000):
                not_found("aba_contabilidade")
            bot.click()
            if not bot.find("campo_valor_historico", matching=0.97, waiting_time=10000):
                not_found("campo_valor_historico")
            bot.click_relative(145, 38)
            bot.type_key(linha.Mensalidade)
            bot.tab(presses=2)
            bot.key_end()
            vlr_provisao = f'nf 2024/{linha.Nota} {linha.ResponsávelFinanceiro} aluno {linha.Aluno}'
            bot.paste(vlr_provisao)
            bot.page_down()
            bot.key_end()
            vlr_devido = f'issqn cf {vlr_provisao}'
            bot.paste(vlr_devido)
            bot.enter()

            if float(linha.Alimentação.replace(',','.')) != 0.0:
                # if not bot.find("botao_novo", matching=0.97, waiting_time=10000):
                #     not_found("botao_novo")
                # bot.click()
                bot.enter()
                bot.type_key('100')
                bot.enter()
                bot.type_key(codigo_refeicao)
                bot.enter()
                bot.type_key(linha.Alimentação)
                bot.tab()
                bot.type_key('8')
                bot.tab()
                bot.key_end()
                bot.space()
                vlr_provisao_refeicao = f'refeição cf {vlr_provisao}'
                bot.paste(vlr_provisao_refeicao)
                bot.enter()

            if linha.Extra:
                bot.enter()
                bot.type_key('100')
                bot.enter()
                bot.type_key(codigo_extra)
                bot.enter()
                bot.type_key(linha.Extra)
                bot.tab()
                bot.type_key('8')
                bot.tab()
                bot.key_end()
                bot.space()
                vlr_provisao_extra = f'extra cf {vlr_provisao}'
                bot.paste(vlr_provisao_extra)
                bot.enter()

            resposta = gui.confirm(title='Os dados foram preenchidos corretamente?', buttons=['Continuar', 'Pausa']) 
            if resposta == 'Pausa':
                with open(arquivo_progresso, 'w') as f:
                    f.write(f'Pausa linha {linha.Index + 1}')
                break

            if not bot.find("botao_gravar", matching=0.97, waiting_time=10000):
                not_found("botao_gravar")
            bot.click()
            ############ PRA ZONA NORTE TEM MAIS ESSE AQUI ############
            # if not bot.find("botao_gravar", matching=0.97, waiting_time=10000):
            #     not_found("botao_gravar")
            # bot.click()

            if not bot.find("botao_proximo", matching=0.97, waiting_time=10000):
                not_found("botao_proximo")
            bot.click()
    except:
        with open(arquivo_progresso, 'w') as f:
            f.write(f'Erro linha {linha.Index}')
            raise

def not_found(label):
    print(f"Element not found: {label}")


if __name__ == '__main__':
    main()