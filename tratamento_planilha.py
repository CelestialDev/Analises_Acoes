import pandas as pd
from IPython.display import display
from pandas import read_excel
from openpyxl import Workbook
from openpyxl import load_workbook
import csv
import os
import sys
from playsound import playsound

class Planilha():
    def Verificar(self,nome):
        try:
            self.acao_df = pd.read_excel(f'pastaplanilhas\{nome}.xlsx')
        except:
            return False
    def Iniciar(self,nome):
        self.acao_df_antigo = pd.read_excel(f'pastaplanilhas\{nome}.xlsx')
        self.acao_df = self.acao_df_antigo.apply(lambda x: x.replace('-','0'))
        try:
            self.acao_df[['ROE','ROA','ROIC','DIV/PL','DIV/EBITDA','DIV/EBIT','LIQ.CORRENTE','VALORIZAÇÃO ATUAL', 'VALORIZAÇÃO MES', 'VALORIZAÇÃO ANO']] = self.acao_df[['ROE','ROA','ROIC','DIV/PL','DIV/EBITDA','DIV/EBIT','LIQ.CORRENTE','VALORIZAÇÃO ATUAL', 'VALORIZAÇÃO MES', 'VALORIZAÇÃO ANO']].apply(pd.to_numeric)
        except ValueError:
            playsound('sound_effects\error.wav')
            erro_destruir = input(f"\033[91mErro 256 ocorreu, tente reiniciar o programa.\033[m")
            sys.exit()

        self.em_alta = self.acao_df.loc[((self.acao_df['VALORIZAÇÃO ATUAL'] > 0) & (self.acao_df['VALORIZAÇÃO MES'] > 0) & (self.acao_df['VALORIZAÇÃO ANO'] > 0)),['EMPRESA','AÇÃO','PREÇO ATUAL']]
        self.positivos = self.acao_df.loc[((self.acao_df['ROE'] > 0) & (self.acao_df['ROA'] > 0) & (self.acao_df['ROIC'] > 0)),['EMPRESA','AÇÃO','PREÇO ATUAL']]
        self.dividas = self.acao_df.loc[((self.acao_df['DIV/PL'] <= 0.30) & (self.acao_df['DIV/EBIT'] < 0.30) & (self.acao_df['DIV/EBIT'] < 0.30)),['EMPRESA','AÇÃO','PREÇO ATUAL']]
        self.melhor = self.acao_df.loc[((self.acao_df['VALORIZAÇÃO ATUAL'] > 0) & (self.acao_df['VALORIZAÇÃO MES'] > 0) & (self.acao_df['VALORIZAÇÃO ANO'] > 0) & (self.acao_df['ROE'] > 0) & (self.acao_df['ROA'] > 0) & (self.acao_df['ROIC'] > 0) & (self.acao_df['DIV/PL'] <= 0.30) & (self.acao_df['DIV/EBIT'] < 0.30) & (self.acao_df['DIV/EBIT'] < 0.30)),['EMPRESA','AÇÃO','PREÇO ATUAL']]
        #self.em_alta = self.acao_df.query("VALORIZAÇÃO MES >0 & VALORIZAÇÃO ATUAL >0")

        os.system('cls' if os.name == 'nt' else 'clear')
        while True:
            self.aba = int(input('Qual informação deseja?\n'
            '[1] Melhores ações\n'
            '[2] Ações em altas\n'
            '[3] Ações com ROE/ROA/ROIC positivos\n'
            '[4] Ações com dívidas baixas\n'
            '\033[1;31m[5] Voltar\033[m\n'
            '----------------------------\n'
            'Sua opção: '))
            if self.aba >= 1 and self.aba <= 5:
                break
            os.system('cls' if os.name == 'nt' else 'clear')
            print('\033[91mOcorreu um erro. Digite o número: 1, 2, 3, 4 ou 5\033[m')
            playsound('error.wav')  
        os.system('cls' if os.name == 'nt' else 'clear')
        if self.aba == 1:
            print('Melhores ações para investir:')
            display(self.melhor)
        elif self.aba == 2:
            print('Investimentos em alta:')
            display(self.em_alta)
        elif self.aba == 3:
            print('Investimentos com ROE/ ROA/ ROIC positivos:')
            display(self.positivos)
        elif self.aba == 4:
            print('Investimentos com dívidas baixas:')
            display(self.dividas)
        
        lista_nomes = ['melhores_acoes','em_alta','positivos','dividas']
        aba = input('Deseja salvar estes dados [S | N] ?\n').lower()
        for x in range(0,5):
            if aba == 's':
                if self.aba == x:
                    self.dividas.to_excel(f'pastaplanilhas\{lista_nomes[x]}_{nome}.xlsx', index = False)