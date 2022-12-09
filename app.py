from leitura_ficha_financeira import ler_planilha_verbas
from login_unico import logar_unico
from selecionar_mod_DP import acessar_modulo_folha   
from modulo_rendimentos_variaveis import *
from formatar_valores import formatar_valores_reais
from de_para import de_para_unico

if __name__ == '__main__':
    # logar_unico()
    # acessar_modulo_folha()
    sleep(10)
    acessar_rendimentos_variariveis()
    nova_empresa = True
    novo_empregado = True
    funcionario_existente = True
    dic_empresa = ler_planilha_verbas()
    trocar_empresa = False
    
    num_lancamento = 3527

    for cod_empresa in dic_empresa.keys():
        nova_empresa = True
        if trocar_empresa:
            pyautogui.click(653,128, duration=1)
            sleep(5)
            pyautogui.click(1344,97, duration=1)
            sleep(5)
            pyautogui.click(238,59, duration=1)
            sleep(5)
            pyautogui.click(88,250, duration=1)
            sleep(5)
            acessar_rendimentos_variariveis()
        
        for cod_empregado in dic_empresa.get(cod_empresa).keys():
            novo_empregado = True
            funcionario_existente = True
            
            for rendimentos_variaveis in dic_empresa.get(cod_empresa).get(cod_empregado).keys():
                if not funcionario_existente:
                    break
                
                for competencias, valores in dic_empresa.get(cod_empresa).get(cod_empregado).get(rendimentos_variaveis).items():
                    num_lancamento = num_lancamento + 1
                    print(cod_empresa, cod_empregado, rendimentos_variaveis, competencias, formatar_valores_reais(valores), num_lancamento)
    
                    funcionario_existente = cadastro_ficha_rendimentos_variaveis(cod_empresa, cod_empregado.strip(), competencias, de_para_unico(rendimentos_variaveis), formatar_valores_reais(valores), nova_empresa, novo_empregado)
                    nova_empresa = False
                    novo_empregado = False
                    trocar_empresa = True
                    if not funcionario_existente:
                        break
                

    
    

