import re
from botcity.core import DesktopBot
from botcity.maestro import *
# import pygetwindow as gw
from dados import *
BotMaestroSDK.RAISE_NOT_CONNECTED = False



def ultimos_digitos_nao_zero(sequencia):
    # Expressão regular para encontrar a sequência de dígitos não-zeros no final
    match = re.search(r'(\d{4})(0*)([1-9]\d*)', sequencia) 
    return match.group(3) if match else sequencia
  

if arquivo_progresso.exists():
    with open(arquivo_progresso) as f:
        linha = int(f.read().split()[-1])
    dados = dados.iloc[linha:]


def main(empresa, data_lancto):
    bot = DesktopBot()
    
    codigo_refeicao = '3004' if empresa == 'MB_ZN' else '3103'
    codigo_extra = '3609'

    bot.wait(10000) # Espera para o usuário mudar para a tela do programa

    '''
    Colocar aqui uma janela dizendo para mudar para a janela do programa
    A janela deve conter uma contagem regressiva e uma botão para cancelar
    Após a contagem terminar a janela deve ser fechada
    '''

    try:
        for linha in dados.itertuples():
            if not bot.find("ativar_edicao", matching=0.97, waiting_time=10000):
                not_found("ativar_edicao")
            bot.click()

            bot.copy_to_clipboard('') # esvazia a área de transferência

            if not bot.find("numero_nota", matching=0.97, waiting_time=10000):
                not_found("numero_nota")
            bot.click_relative(143, 5)
            bot.control_c()
            
            # Garantia de que a variável consiga obter o valor da área de transferência
            textoCtrlC = ''
            while (len(textoCtrlC) == 0):
                textoCtrlC = bot.get_clipboard()

            numero_nota_dominio = ultimos_digitos_nao_zero(textoCtrlC)
            if numero_nota_dominio != linha.Notas:
                print('Número da nota não bate com a retornada pelo sistema')
                print(f'{numero_nota_dominio} | {linha.Notas}')
                raise

            if not bot.find("campo_emissao", matching=0.97, waiting_time=10000):
                not_found("campo_emissao")
            bot.click_relative(104, 7, wait_after=1500, clicks=3)
            bot.type_key(data_lancto)

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
            vlr_provisao = f'nf 2025/{linha.Notas} {linha.ResponsávelFinanceiro} aluno {linha.Aluno}'
            bot.paste(vlr_provisao)
            bot.page_down()
            bot.key_end()
            vlr_devido = f'issqn cf {vlr_provisao}'
            bot.paste(vlr_devido)
            bot.enter()

            if (float(linha.Alimentação.replace(',','.')) != 0.0) and (linha.Alimentação != 'nan'):
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
    
            try:
                if (float(linha.Extra.replace(',','.')) != 0.0) and (linha.Extra != 'nan'):
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

            except AttributeError:
                ...
            

            if not bot.find("botao_gravar", matching=0.97, waiting_time=10000):
                not_found("botao_gravar")
            bot.click()
            ############ PRA ZONA NORTE TEM MAIS ESSE AQUI ############
            if empresa == 'MB_ZN':
                if not bot.find("botao_gravar", matching=0.97, waiting_time=10000):
                    not_found("botao_gravar")
                bot.click()

            if not bot.find("botao_proximo", matching=0.97, waiting_time=10000):
                not_found("botao_proximo")
            bot.click()#clicks=2, interval_between_clicks=1500)
            bot.wait(2000)
    except:
        with open(arquivo_progresso, 'w') as f:
            f.write(f'Erro {linha.ResponsávelFinanceiro} linha {linha.Index}')
            raise

def not_found(label):
    print(f"Element not found: {label}")


if __name__ == '__main__':
    main(empresa='MB_ZS', data_lancto='30072025')
    