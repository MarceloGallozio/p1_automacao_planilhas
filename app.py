import openpyxl
import pyperclip
import pyautogui
from webbrowser import open
from time import sleep
from tkinter import *

def iniciar_cadastro():
# Entrar no site
  open('https://cadastro-produtos-devaprender.netlify.app/')
  sleep(10)
# Entrar na planilha
  workbook = openpyxl.load_workbook('produtos_ficticios.xlsx')
  sheet_produtos = workbook['Produtos']
  # Copiar informação de um campo e colar no seu campo correspondente
  for linha in sheet_produtos.iter_rows(min_row=2):
    #Nome produto
    nome_produto = linha[0].value
    pyperclip.copy(nome_produto)
    pyautogui.click(158,350, duration=1)
    pyautogui.hotkey('ctrl','v')

    #Descrição
    descricao = linha[1].value
    pyperclip.copy(descricao)
    pyautogui.click(150,454, duration=1)
    pyautogui.hotkey('ctrl','v')

    #Categoria
    categoria = linha[2].value
    pyperclip.copy(categoria)
    pyautogui.click(149,572, duration=1)
    pyautogui.hotkey('ctrl','v')

    #Código Produto
    codigo_produto = linha[3].value
    pyperclip.copy(codigo_produto)
    pyautogui.click(158,657, duration=1)
    pyautogui.hotkey('ctrl','v')

    #Peso
    peso = linha[4].value
    pyperclip.copy(peso)
    pyautogui.click(158,745, duration=1)
    pyautogui.hotkey('ctrl','v')

    #Dimensôes
    dimensoes = linha[5].value
    pyperclip.copy(dimensoes)
    pyautogui.click(160,830, duration=1)
    pyautogui.hotkey('ctrl','v')

    #Botão Próximo
    pyautogui.click(171,887, duration=1)
    sleep(2)

    #Preço
    preco = linha[6].value
    pyperclip.copy(preco)
    pyautogui.click(158,350, duration=1)
    pyautogui.hotkey('ctrl','v')

    #Quantidade em estoque
    quantidade_em_estoque = linha[7].value
    pyperclip.copy(quantidade_em_estoque)
    pyautogui.click(161,483, duration=1)
    pyautogui.hotkey('ctrl','v')

    #Data de validade
    data_de_validade = linha[8].value
    pyperclip.copy(data_de_validade)
    pyautogui.click(149,572, duration=1)
    pyautogui.hotkey('ctrl','v')

    #Cor
    cor = linha[9].value
    pyperclip.copy(cor)
    pyautogui.click(158,657, duration=1)
    pyautogui.hotkey('ctrl','v')

    #Tamanho
    tamanho = linha[10].value  
    pyautogui.click(199,737, duration=1)
    if tamanho == 'Pequeno':
      pyautogui.click(174,768, duration=1)
    elif tamanho == 'Médio':
      pyautogui.click(194,790, duration=1)
    else:
      pyautogui.click(185,812, duration=1)

    #Material
    material = linha[11].value
    pyperclip.copy(material)
    pyautogui.click(140,834, duration=1)
    pyautogui.hotkey('ctrl','v')

    #Botão próximo
    pyautogui.click(171,887, duration=1)
    sleep(2)

    #Fabricante
    fabricante = linha[12].value
    pyperclip.copy(fabricante)
    pyautogui.click(149,410, duration=1)
    pyautogui.hotkey('ctrl','v')

    #País Origem
    pais_origem = linha[13].value
    pyperclip.copy(pais_origem)
    pyautogui.click(151,496, duration=1)
    pyautogui.hotkey('ctrl','v')

    #Observações
    observacoes = linha[14].value
    pyperclip.copy(observacoes)
    pyautogui.click(149,587, duration=1)
    pyautogui.hotkey('ctrl','v')

    #Código de Barras
    codigo_de_barras = linha[15].value
    pyperclip.copy(codigo_de_barras)
    pyautogui.click(159,718, duration=1)
    pyautogui.hotkey('ctrl','v')

    #Localização Armazem
    localizacao_armazem = linha[16].value
    pyperclip.copy(localizacao_armazem)
    pyautogui.click(164,804, duration=1)
    pyautogui.hotkey('ctrl','v')

    #Botão Continuar
    pyautogui.click(169,861, duration=1)

    #Botão Ok
    pyautogui.click(647,189, duration=1)
    
    #Botão Adicionar mais um item
    pyautogui.click(486,622, duration=1)
    sleep(2)
  
# Repitir esses passos para outros campos até preencher campos daquela página
# Clicar em próxima
# Repetir os mesmos passos e ir para a próxima página(página 2)
# Repetir os mesmos passos e finalizar o cadastro daquele produto e clicar em concluir
# clicar em ok, para finalizar o processo
# clicar no ok mais uma vez na mensagem de confirmação de salvamento no banco de dados
# clicar em "adicionar mais um" e repetir o processo até finalizar a planilha
def centralizar_janela(janela, largura, altura):
    largura_tela = janela.winfo_screenwidth()
    altura_tela = janela.winfo_screenheight()
    x = (largura_tela - largura) // 2
    y = (altura_tela - altura) // 2
    janela.geometry(f"{largura}x{altura}+{x}+{y}")


janela = Tk()
janela.title('Cadastro de Produtos')

largura_janela = 400
altura_janela = 300
janela.geometry(f"{largura_janela}x{altura_janela}")
centralizar_janela(janela, largura_janela, altura_janela)

texto_orientacao = Label(janela, text='Clique no botão para iniciar o Cadastro dos Produtos')
texto_orientacao.grid(column=0, row=0, padx=50, pady=70)

botao = Button(janela, text='Iniciar Cadastro', command=iniciar_cadastro)
botao.grid(column=0, row=1, padx=50, pady=0)

janela.mainloop()