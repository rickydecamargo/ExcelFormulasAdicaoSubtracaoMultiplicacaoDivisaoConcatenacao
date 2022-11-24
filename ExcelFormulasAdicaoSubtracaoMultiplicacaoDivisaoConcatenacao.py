#Script para criar um arquivo excel, inserir dados nele e formatar células com allign, cor da fonte, negrito e cor de fundo.

import xlsxwriter as opcoesDoXlsxWriter
import os

#1 - indicando onde será criado o arquivo, seu nome e sua. Importante a questão das barras duplas (testar).
nomeCaminhoArquivo = 'C:\\Users\\Windows\\Desktop\\Python Projetos\\xlsxwriter\PintaFundoEFonte.xlsx'
minhaPlanilha = opcoesDoXlsxWriter.Workbook(nomeCaminhoArquivo)
sheetDados = minhaPlanilha.add_worksheet("Dados")

#Para adicionar cor de fundo a célula
corFundo = minhaPlanilha.add_format({'fg_color':'yellow'})

#Para colorir a fonte
corFonte = minhaPlanilha.add_format()
corFonte.set_font_color('blue')

#cria a variável para alinhar, mudar a cor da fonte, negrito e cor de fundo.
corFonteFundo = minhaPlanilha.add_format({'align': 'center',
                                          'font_color': 'white',
                                          'bold': True,
                                          'bg_color': 'gray'})

#Preto = black
#Branco = white
#Amarelo = yellow
#Laranja = orange
#Vermelho = red
#Azul = blue
#Verde = green
#Cinza = gray
#Rosa = pink
#Roxo = purple
#Marinho = navy
#Prata = silver

#2 - Criando as Colunas e Linhas com algumas informações
sheetDados.write("A1", "Nome", corFonteFundo) #chama a variável que pinta a célula
sheetDados.write("B1", "Idade", corFonteFundo) #chama a variável que pinta a célula
sheetDados.write("A2", "Amanda", corFonte)
sheetDados.write("B2", 21, corFonte)
sheetDados.write("A3", "Allan", corFonte)
sheetDados.write("B3", 28, corFonte)

#3 - Para fechar e salvar as informações
minhaPlanilha.close()

#4 - Abrir o arquivo para verificar o resultado
os.startfile(nomeCaminhoArquivo)