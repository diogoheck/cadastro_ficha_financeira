from openpyxl import load_workbook
import os
class Indice:
    EMPRESA = 0
    NOME_EMPRESA = 5
    POSICAO_COMISSAO = 1
    POS_EMPREGADO = 1
    NOV = 12
    DEZ = 15
    JAN = 19
    FEV = 21
    MAR = 24
    ABR = 26
    MAI = 29
    JUN = 31
    JUL = 34
    AGO = 37
    SET = 41
    OUT = 46



def ler_planilha_verbas():
    pasta_verbas = load_workbook('planilhas/Ficha Financeira Garra Santo Angelo.xlsx')
    planilha_verbas = pasta_verbas['Ficha Financeira']
    
    dic_empregado = {}
    dic_empresa = {}
    for linha in planilha_verbas.values:
        if linha[Indice.EMPRESA] == 'Empresa:':
            # codigo_empresa = linha[Indice.NOME_EMPRESA].split('-')[0]
            codigo_empresa = '2371'
            # print(codigo_empresa)
            dic_empresa[codigo_empresa] = {}
        
        if linha[Indice.POS_EMPREGADO] == 'Empregado:':
            codigo_empregado = linha[8].split('-')[0]
            # print(f'empregado => {codigo_empregado}')
            dic_empregado[codigo_empregado] = {}
            dic_empregado[codigo_empregado][37] = {}
            dic_empregado[codigo_empregado][250] = {}
            dic_empresa[codigo_empresa].update(dic_empregado)

    

        if linha[Indice.POSICAO_COMISSAO] == 37:

            if linha[Indice.NOV]:
                dic_empregado[codigo_empregado][37].update({'112021' : linha[Indice.NOV]})
                
            if linha[Indice.DEZ]:
                dic_empregado[codigo_empregado][37].update({'122021' : linha[Indice.DEZ]})
                
            if linha[Indice.JAN]:
                dic_empregado[codigo_empregado][37].update({'012022' : linha[Indice.JAN]})
                
            if linha[Indice.FEV]:
                
                dic_empregado[codigo_empregado][37].update({'022022' : linha[Indice.FEV]})
            if linha[Indice.MAR]:
                
                dic_empregado[codigo_empregado][37].update({'032022' : linha[Indice.MAR]})
            if linha[Indice.ABR]:
                dic_empregado[codigo_empregado][37].update({'042022' : linha[Indice.ABR]})
                
            if linha[Indice.MAI]:
                dic_empregado[codigo_empregado][37].update({'052022' : linha[Indice.MAI]})
                
            if linha[Indice.JUN]:
                dic_empregado[codigo_empregado][37].update({'062022' : linha[Indice.JUN]})
                
            if linha[Indice.JUL]:
                
                dic_empregado[codigo_empregado][37].update({'072022' : linha[Indice.JUL]})
            if linha[Indice.AGO]:
                dic_empregado[codigo_empregado][37].update({'082022' : linha[Indice.AGO]})
                
            if linha[Indice.SET]:
                dic_empregado[codigo_empregado][37].update({'092022' : linha[Indice.SET]})
                
            if linha[Indice.OUT]:
                dic_empregado[codigo_empregado][37].update({'102022' : linha[Indice.OUT]})

        if linha[Indice.POSICAO_COMISSAO] == 250 or linha[Indice.POSICAO_COMISSAO] == 853 or linha[Indice.POSICAO_COMISSAO] == 854:

            if linha[Indice.NOV]:
                if not dic_empregado[codigo_empregado][250].get('112021'):
                    dic_empregado[codigo_empregado][250].update({'112021' : linha[Indice.NOV]})
                else:
                    dic_empregado[codigo_empregado][250]['112021'] += linha[Indice.NOV]
                
            if linha[Indice.DEZ]:

                if not dic_empregado[codigo_empregado][250].get('122021'):
                    dic_empregado[codigo_empregado][250].update({'122021' : linha[Indice.DEZ]})
                else:
                    dic_empregado[codigo_empregado][250]['122021'] += linha[Indice.DEZ]
                
                
            if linha[Indice.JAN]:
                
                if not dic_empregado[codigo_empregado][250].get('012022'):
                    dic_empregado[codigo_empregado][250].update({'012022' : linha[Indice.JAN]})
                else:
                    dic_empregado[codigo_empregado][250]['012022'] += linha[Indice.JAN]

                
            if linha[Indice.FEV]:
                
                if not dic_empregado[codigo_empregado][250].get('022022'):
                    dic_empregado[codigo_empregado][250].update({'022022' : linha[Indice.FEV]})
                else:
                    dic_empregado[codigo_empregado][250]['022022'] += linha[Indice.FEV]
                

            if linha[Indice.MAR]:

                if not dic_empregado[codigo_empregado][250].get('032022'):
                    dic_empregado[codigo_empregado][250].update({'032022' : linha[Indice.MAR]})
                else:
                    dic_empregado[codigo_empregado][250]['032022'] += linha[Indice.MAR]
                


            if linha[Indice.ABR]:

                if not dic_empregado[codigo_empregado][250].get('042022'):
                    dic_empregado[codigo_empregado][250].update({'042022' : linha[Indice.ABR]})
                else:
                    dic_empregado[codigo_empregado][250]['042022'] += linha[Indice.ABR]


                
            if linha[Indice.MAI]:

                if not dic_empregado[codigo_empregado][250].get('052022'):
                    dic_empregado[codigo_empregado][250].update({'052022' : linha[Indice.MAI]})
                else:
                    dic_empregado[codigo_empregado][250]['052022'] += linha[Indice.MAI]
                
                
            if linha[Indice.JUN]:

                if not dic_empregado[codigo_empregado][250].get('062022'):
                    dic_empregado[codigo_empregado][250].update({'062022' : linha[Indice.JUN]})
                else:
                    dic_empregado[codigo_empregado][250]['062022'] += linha[Indice.JUN]
                
                
            if linha[Indice.JUL]:

                if not dic_empregado[codigo_empregado][250].get('072022'):
                    dic_empregado[codigo_empregado][250].update({'072022' : linha[Indice.JUL]})
                else:
                    dic_empregado[codigo_empregado][250]['072022'] += linha[Indice.JUL]
                
                
            if linha[Indice.AGO]:

                if not dic_empregado[codigo_empregado][250].get('082022'):
                    dic_empregado[codigo_empregado][250].update({'082022' : linha[Indice.AGO]})
                else:
                    dic_empregado[codigo_empregado][250]['082022'] += linha[Indice.AGO]
                
                
            if linha[Indice.SET]:

                if not dic_empregado[codigo_empregado][250].get('092022'):
                    dic_empregado[codigo_empregado][250].update({'092022' : linha[Indice.SET]})
                else:
                    dic_empregado[codigo_empregado][250]['092022'] += linha[Indice.SET]
                
                
            if linha[Indice.OUT]:

                if not dic_empregado[codigo_empregado][250].get('102022'):
                    dic_empregado[codigo_empregado][250].update({'102022' : linha[Indice.OUT]})
                else:
                    dic_empregado[codigo_empregado][250]['102022'] += linha[Indice.OUT]
                
                 
    return dic_empresa    