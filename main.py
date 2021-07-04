from playsound import playsound
from menu import Menu
#from tratamento_planilha import Planilha
import sys

#p = Planilha()
m = Menu()

while True:
    while True:
        print('Carregando cores')
        break
    m.LPrompt()
    m.menu_cabecalho()
    m.menu_opcoes()

    m.LPrompt()

    m.opcao_1()
    m.opcao_2()
    m.opcao_3()
    m.opcao_4()
    m.opcao_5()
    
    final = input('Execução finalizada. Deseja sair do programa? [S/N]\n').upper()
    if final == "S":
        sys.exit()
    
    elif final == "N":
        m.LPrompt()