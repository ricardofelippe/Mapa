import pyautogui as p
from pyautogui import click, doubleClick
import os
import schedule
import schedule
import time
#from random import *
import random 
from typing import TextIO
import pyperclip
#from sklearn.covariance import oas
#from sklearn.discriminant_analysis import QuadraticDiscriminantAnalysis



# PREMISSAScf
# 1) PCA 2023-Em Execução aberta
# 2) Criar no DFD clicando no ícone respectivo Point(x=1679, y=872)
# 3) Aceitar copia do DFD Point(x=1142, y=314)
# 4) Alterar data 01/01/2024 Point(x=546, y=564)
# 5) Clicar aba responsaveis Point(x=136, y=730)
# 6) Adicionar responsaveis Point(x=1770, y=407)
# 5) Preencher CPF 40504948172 Point(x=624, y=302)
# 6) Preencher email hudson.galvani@agro.gov.br Point(x=629, y=567)
# 7) Finalizar com botao confirmar Point(x=1266, y=930)

# 2.1) Excel aberto e maximizado e deve estar em segundo plano
# 3) Preencher 
#   3.1- Informações Gerais
        # * Data Conclusão
        # * Área requisitante
        # * Descrição Sucinta
        # * Prioridade
        # * Justificativa somente para Prioridade Alta
#   3.2- Justificativa de Necessidade
        # Copiar e Colar justificativa
#   3.3- Materiais/Serviços
        # Clicar botão-Adicionar   
#   3.4- Responsáveis
        # Clicar botão-Adicionar   


#global temp,tempo_mouse,info_excel
    

def job_responsa():
        p.confirm("Iniciar DFD's ?") 
                
        tempo_mouse=0.01
        dfd_invalidas=[43,65,101,102,103,105,106,111,122,126,131,136,139,141,144,184,185,186,187,188,190,193,194,195,196,197,198,199,200,201,204,205,206,207,208,209,210,211,212,213,214,215,216,217,218,219,220,221,222,223,224,225,226,227,228,229,230,231,249,250,251,252,253,254,255,256,257,258,259,260,261,262,263,264,265,266,267,271,274,279,283,287,288,289,291]
        dfd_invalidas_2022=[]

        for i in dfd_invalidas:
                j=str(i)+"/2022"
                dfd_invalidas_2022.append(j)
        
        



    # 1) PCA 2023-Em Execução aberta
    # 2) Clicar na caixa de procura Point(x=88, y=18) usar 
    # 2) Criar no DFD clicando no ícone respectivo Point(x=1679, y=872)
    # 3) Aceitar copia do DFD Point(x=1142, y=314)
    # 4) Alterar data 01/01/2024 Point(x=546, y=564)
    # 5) Clicar aba responsaveis Point(x=136, y=730)
    # 6) Adicionar responsaveis Point(x=1770, y=407)
    # 7) Preencher CPF 40504948172 Point(x=624, y=302)
    # 8) Preencher email hudson.galvani@agro.gov.br Point(x=629, y=567)
    # 9) Finalizar com botao confirmar Point(x=1266, y=930)


    # 2) Criar no DFD clicando no ícone respectivo Point(x=1679, y=872)
        #dfd_temp=list(range(18,293))
        #dfd_temp=list(range(28,293))
        #dfd_temp=list(range(41,293))
        #dfd_temp=list(range(51,293))
        #dfd_temp=list(range(91,293))
        dfd_temp=list(range(151,293))
        
        
        dfd=[]
        for i in dfd_temp:
                j=str(i)+"/2022"
                dfd.append(j)
        

        for i in dfd:
                if i in dfd_invalidas_2022:
                        continue
                else:        
                        # 2) Clicar na caixa de procura Point(x=88, y=18) usar 
                        p.hotkey("ctrl","home")
                        p.moveTo(714, 585,tempo_mouse)
                        p.click(714, 585)
                        p.hotkey("enter")
                        time.sleep(0.1)
                        p.write(i)
                        time.sleep(1)


                        #p.confirm('Buscando DFD ' + i +' Continuar?')
                        p.moveTo(1742, 591,tempo_mouse)
                        p.click(x=1742, y=591)
                        time.sleep(1)


        
                        # 2) Criar no DFD clicando no ícone respectivo Point(x=1679, y=872)
                        p.confirm('Criar DFD: ' + i + 'Continuar?')
                        p.moveTo(1679, 872,tempo_mouse)
                        p.click(1679, 872)
                        time.sleep(1)

                        # 3) Aceitar copia do DFD Point(x=1142, y=314)
                        p.moveTo(1142, 314,tempo_mouse)
                        p.confirm('Aceitar copia? ')
                        
                        p.click(x=1142, y=314)
                        #time.sleep(1)

                        # 4) Alterar data 01/01/2024 Point(x=546, y=564)
                        #x=546, y=571)
                        p.moveTo(546, 571,tempo_mouse)
                        p.confirm('Confirma data ? ')
                        p.click(x=546, y=571)
                        time.sleep(0.5)
                        p.write("01/01/2024")
                        time.sleep(1)
                        



                        # 5) Clicar aba responsaveis Point(x=136, y=730)
                        p.moveTo(136, 730,tempo_mouse)
                        #p.confirm('Confirma responsaveis ? ')
                        
                        p.click(x=136, y=730)
                        time.sleep(0.5)

                        # 6) Adicionar responsaveis Point(x=1770, y=407)
                        p.moveTo(1770, 407,tempo_mouse)
                        p.confirm('Confirma adicionar responsaveis ? ')
                                
                        p.click(x=1770, y=407)
                        time.sleep(0.5)

                        # 7) Preencher CPF 40504948172 Point(x=624, y=302)
                        p.click(x=624, y=302)
                        time.sleep(0.5)
                        p.write("40504948172")

                        # 8) Preencher email hudson.galvani@agro.gov.br Point(x=629, y=567)
                        p.moveTo(631, 530,tempo_mouse)
                        #p.confirm('Confirma adicionar email? ')
                        p.click(631, 530)
                        p.write("hudson.galvani@agro.gov.br")
                        time.sleep(0.5)
                        

                        #
                        #  9) Finalizar com botao confirmar Point(x=1266, y=930)
                        p.confirm('Confirma adicionar responsaveis ? ')
                        p.moveTo(1264, 871,tempo_mouse)
                        
                        p.click(x=1264, y=871)
                        time.sleep(1)

                        #  Voltar tela inicial
                        p.moveTo(587, 216,tempo_mouse)
                        p.confirm('Confirma volta tela inicial ? ')
                        p.click(x=587, y=216)
                        time.sleep(0.5)





                        # Limpar campo de busca
                        p.hotkey("ctrl","home")
                        time.sleep(1)
                        p.click(x=1035, y=586) # seleciona flecha
                        time.sleep(0.5)
                        p.press('backspace',10)
                        time.sleep(1)
                        
                        



                


        return

def job():
        
    job_responsa()

       





    return

job()

