from playsound import playsound
import os
import string

from tratamento_planilha import Planilha
from webscraping import WebScraping

ws = WebScraping()
p = Planilha()

class Menu():
    def menu_cabecalho(self):
        print("-" * 28)
        print(' ' * 5,'\033[1;32mAnálise de dados\033[m')
        print("-" * 28)
        print("")

    def LPrompt(self):
        os.system('cls' if os.name == 'nt' else 'clear')

    def voltar(self):
        self.LPrompt()
        self.menu_cabecalho()
        self.menu_opcoes()

    def menu_opcoes(self):
        while True:
            self.aba1 = int(input('O que deseja fazer?\n'
                       '[1] Buscar ações\n' # Feito
                       '[2] Analisar ações\n'
                       '[3] Informações sobre ações\n' # Feito
                       '[4] Informações sobre o programa\n' # Feito
                       '[5] Ações obtidas\n'
                       '----------------------------\n'
                       'Sua opção é: '))
            if self.aba1 == 1 or self.aba1 == 2 or self.aba1 == 3 or self.aba1 == 4 or self.aba1 == 5:
                break
            self.LPrompt()
            print('\033[91mOcorreu um erro. Digite o número: 1, 2, 3, 4 ou 5\033[m')
            playsound('error.wav')
    
    def opcao_1(self):
        if self.aba1 == 1:
            while True:
                self.aba2 = int(input('Como deseja buscar as ações?\n'
                             '[1] Buscar ação específica\n' # Feito
                             '[2] Altas e baixas do dia\n' # Feito
                             '\033[1;31m[3] Voltar\033[m\n'
                             '----------------------------\n'
                             'Sua opção é: '))
                if self.aba2 == 1 or self.aba2 == 2:
                    break
                elif self.aba2 == 3:
                    break
                self.LPrompt()
                print('\033[91mOcorreu um erro. Digite o número: 1 ou 2\033[m')
                playsound('error.wav')
            self.LPrompt()
            self.opcao_1_1()
            self.opcao_1_2()
    
    def opcao_1_1(self):
        if self.aba2 == 1:
            ws.Iniciar(self.aba2, 0, 0)
    
    def opcao_1_2(self):
        if self.aba2 == 2:
            aux_loop = True
            while True:
                self.indice = int(input( 'Deseja buscar as ações com base no:\n' # Feito
                                '[1] Ibovespa\n' 
                                '[2] Indíce Financeiro\n'
                                '[3] Indíce de Consumo\n'
                                '[4] Geral\n'
                                '[5] Indíce Imobiliário\n'
                                '[6] Indíce Industrial\n'
                                '[7] Indíce Valor\n'
                                '\033[1;31m[8] Voltar\033[m\n'
                                '----------------------------\n'
                                'Sua opção é: '))
                if self.indice >= 1 and self.indice <= 8:
                    if self.indice != 8:
                        aux_loop = False
                    break
                self.LPrompt()
                print('\033[91mOcorreu um erro. Digite o número 1 até 7\033[m')
                playsound('error.wav')
            if aux_loop == False:
                self.LPrompt()
                ws.Iniciar(self.aba2, self.indice, 0)
    
    def opcao_2(self):
        if self.aba1 == 2:
            while True:
                self.lista_nome_investimento = ["investimentos_ibovespa","investimentos_financeiro",
                "investimentos_consumo","investimentos_geral","investimentos_imobiliario","investimentos_industrial","investimento_valor"]

                aba = int(input('Que planilha deseja analisar?\n'
                '[1] investimentos_ibovespa\n'
                '[2] investimentos_financeiro\n'
                '[3] investimentos_consumo\n'
                '[4] investimentos_geral\n'
                '[5] investimentos_imobiliario\n'
                '[6] investimentos_industrial\n'
                '[7] investimento_valor\n'
                '\033[1;31m[8] Voltar\033[m\n'
                '----------------------------\n'
                'Sua opção é: '))
                if aba >= 1 and aba <= 9:
                    break
                self.LPrompt()
                print('\033[91mOcorreu um erro. Digite o número: 1, 2, 3, 4 ou 5\033[m')
                playsound('error.wav')  
            for x in range(0,7):
                if aba == x:
                    if p.Verificar(self.lista_nome_investimento[x - 1]) == False:
                        ws.destruir_programa('128')
                    else:
                        p.Iniciar(self.lista_nome_investimento[x - 1])

    def opcao_3(self):
        if self.aba1 == 3:
            print("\033[1;97mA planilha desenvolvida pelo programa entregará os seguintes dados:\033[m")
            print("")
            print("\033[1;93mEMPRESA\033[m = Nome da empresa\n"
            "\033[1;93mAÇÃO\033[m = Cognome da ação"
            "\n\033[1;93mPREÇO ATUAL\033[m = Valor da ação no dia em que o programa foi realizado, ou seja, atual\n"
            "\033[1;93mROE\033[m = Retorno sobre Patrimônio Líquido (ROE) é um indicador que analisa a capacidade que uma empresa tem para gerar valor para o negócio e para investidores.\n"
            "\033[1;93mROA\033[m = Retorno sobre o Ativo (ROA) é um índice de rentabilidade que tem como objetivo medir os lucros da empresa com base nos ativos..\n" 
            "\033[1;93mROIC\033[m = Retorno Sobre Capital Investido (ROIC) é um indicador que avalia a rentabilidade do investimento aplicado pelos acionistas e credores na empresa.\n"
            "\033[1;93mDIV/PL\033[m = Dívida líquida se refere à soma do volume de empréstimos e financiamentos realizados para alcançar um determinado fim, subtraindo o caixa do negócio.\n"
            "\033[1;93mDIV/EBITDA\033[m = Indica quanto tempo seria necessário para pagar a dívida líquida da empresa considerando o EBITDA atual.\n"
            "\033[1;93mDIV/EBIT\033[m = Indica quanto tempo seria necessário para pagar a dívida liquida da empresa considerando o EBIT atual.\n"
            "\033[1;93mLIQ.CORRENTE\033[m = Indica a capacidade de pagamento da empresa a curto prazo.\n"
            "\033[1;93mVALORIZAÇÃO ATUAL\033[m = Indica quanto a ação subiu ou caiu no dia\n"
            "\033[1;93mVALORIZAÇÃO MES\033[m = Indica quanto a ação subiu ou caiu no mês\n"
            "\033[1;93mVALORIZAÇÃO ANO\033[m = Indica quanto a ação subiu ou caiu no ano\n")
            print("")
            print('Caso tenha um "-" em algum dado na planilha, significa que não foi encontrado esta informação.')
            print("Quanto maior o dado de ROE, ROA, ROIC, LIQ.CORRENTE, melhor.\nQuanto menor o dado de DIV/PL, DIV/EBITDA, DIV/EBIT, melhor")
        
    def opcao_4(self):
        if self.aba1 == 4:
            print("No decorrer do programa, podem ocorrer alguns erros, o que será necessário apenas uma reinicialização do programa no máximo.")
            print("")
            print("\033[91mErro 002: \033[mNão foi possível localizar uma propaganda por causa do tempo de demora ao acesso do site.\n"
            "\033[91mErro 004: \033[mNão foi possível localizar uma propaganda porque não foi localizado nenhum elemento.\n"
            "\033[91mErro 008: \033[mNão foi possível localizar a ação inserida por ser uma ação inválida.\n"
            '\033[91mErro 016: \033[mNão foi possível localizar elemento "TODOS".\n'
            "\033[91mErro 032: \033[mO navegador não estava aberto.\n"
            "\033[91mErro 064: \033[mNão foi possível localizar o elemento de filtro por indíce.\n"
            "\033[91mErro 128: \033[mNão foi possível localizar a planilha\n"
            "\033[91mErro 256: \033[mNão foi possível ler a planilha, por ter um item com um ponto ou mais.")
    
    def opcao_5(self):
        if self.aba1 == 5:
            while True:
                aba = int(input('O que você deseja fazer?\n'
                '[1] Registrar ações obtidas\n'
                '[2] Ver registro\n'
                '[3] Atualizar ações\n'
                '[4] Deletar registro\n'
                '\033[1;31m[5] Voltar\033[m\n'
                '----------------------------\n'
                'Sua opção é: '))
                if aba >= 1 and aba <= 5:
                    break
                self.LPrompt()
                print('\033[91mOcorreu um erro. Digite o número: 1, 2, 3, 4 ou 5\033[m')
                playsound('error.wav')

            if aba == 1:
                self.registrar_acao()
            elif aba == 2:
                self.LPrompt()
                print('Procurando dados de ações obtidas...')
                self.verificar_dado()

                if os.stat("dados").st_size == 0:
                    regist = input('Parece que você não registrou nenhuma ação obtida. Deseja registrar suas ações? [S/N]\n').upper()
                    if regist == "S":
                        self.registrar_acao()
                else:
                    with open('dados','r') as arquivo:
                        linhas = arquivo.readlines()
                        if len(linhas) == 1:
                            print(f'Foi encontrado {len(linhas)} dado no registro.')
                        else:
                            print(f'Foram encontrados {len(linhas)} dados no registro.')
                        print('')
                        for linha in linhas:
                            dado = linha.split(";")
                            print(f'Empresa: {dado[0]} | Ação: {dado[1]} | Quantidade: {dado[2]} | Valor inicial: {dado[3]}')

            elif aba == 3:
                while True:
                    self.lista_acoes = []
                    self.dados = []
                    self.verificar_dado()
                    if os.stat("dados").st_size == 0:
                        break
                    with open('dados','r') as arquivo:
                        for linha in arquivo:
                            dado = linha.split(";")
                            self.dados.append(dado)
                            self.lista_acoes.append(dado[1])
                    ws.acao_obtida(self.dados)
                    break
                if os.stat("dados").st_size == 0:
                    print('Você não tem nenhum dado no registro!')

            elif aba == 4:
                try:
                    os.remove('dados')
                except:
                    print('Não foi encontrado nenhum arquivo.')
                finally:
                    print('Dados excluidos')
    
    def verificar_dado(self):
        try:
            arquivo = open("dados", "rt")
        except FileNotFoundError:
            arquivo = open ("dados", "wt")
        arquivo.close()
            
    def registrar_acao(self):
        qtd = int(input('Quantas ações deseja registrar?'))
        
        for x in range(0,qtd):
            print(f'Registro {x + 1}/{qtd}')
            empresa_nome = input('Qual o nome da empresa?\n')
            bdr = input('Ação ou BDR: ').lower()
            acao_nome = input('Qual o nome da ação?\n').lower()
            qtd_acao = input('Quantas ações você obteu?\n')
            valor_acao_obtida = input('Preço da ação: ')
            if bdr == 'ação' or bdr == 'acao':
                b = 'acoes'
            else:
                b = 'bdrs'
            with open('dados','a') as arquivo: 
                arquivo.write(empresa_nome+";"+acao_nome+";"+qtd_acao+";"+valor_acao_obtida+";"+b+"\n")
            if qtd > 1:
                continuar = input(f'Registro finalizado {x}, deseja continuar os registros? [S/N]\n').upper()
                if continuar == "N":
                    break
            self.LPrompt()
        print('Registro(s) finalizado(s).')