from leitura_ficha_financeira import ler_planilha_verbas
from login_unico import logar_unico
from selecionar_mod_DP import acessar_modulo_folha   
from modulo_rendimentos_variaveis import *
from formatar_valores import formatar_valores_reais
from de_para import de_para_unico

if __name__ == '__main__':
    # logar_unico()
    # acessar_modulo_folha()
    # acessar_rendimentos_variariveis()
    nova_empresa = True
    novo_empregado = True
    dic_empresa = ler_planilha_verbas()
    # print(dic_empresa)
    # print(dic_empresa.get('2371').get('52 '))
    for cod_empresa in dic_empresa.keys():
        nova_empresa = True
        # print(cod_empresa)
        for cod_empregado in dic_empresa.get(cod_empresa).keys():
            novo_empregado = True
            # print(cod_empregado)
            for rendimentos_variaveis in dic_empresa.get(cod_empresa).get(cod_empregado).keys():
                # print(rendimentos_variaveis)
                for competencias, valores in dic_empresa.get(cod_empresa).get(cod_empregado).get(rendimentos_variaveis).items():
                    print(cod_empresa, cod_empregado, rendimentos_variaveis, competencias, formatar_valores_reais(valores))
                    print(type(valores))
                    cadastro_ficha_rendimentos_variaveis(cod_empresa, cod_empregado.strip(), competencias, de_para_unico(rendimentos_variaveis), formatar_valores_reais(valores), nova_empresa, novo_empregado)
                    nova_empresa = False
                    novo_empregado = False
            exit(1)

    