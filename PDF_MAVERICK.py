import tkinter as tk
from tkinter import messagebox, filedialog, ttk
import os
import shutil
from funcoes_utilizadas import baixa_postal, backup_excel, fichas_nome, escreve_arquivo_não_encontrado, escreve_arquivo
import pandas as pd
from tkinter import ttk


#função para o usuario selecionar o arquivo de excel a ser usado 
def selecionar_arquivo():
    global caminho_arquivo
    caminho_arquivo = filedialog.askopenfilename()
    caminho_arquivo = rf"{caminho_arquivo}"
    print(f"Arquivo selecionado: {caminho_arquivo}")  
    nome = os.path.split(caminho_arquivo)
    if caminho_arquivo == '':
        messagebox.showwarning("Arquivo não selecionado.", "Por favor, selecione um arquivo para continuar.")
    else:
        label_selecionar_arq.config(text=f"Arquivo: {nome[1]}")
        backup_excel(caminho_arquivo)




#Lê o arquivo da variavel caminho_arquivo e seleciona a lista de fichas dentro do excel a serem usadas para copiar o pdf
def lista_fichas(caminho_arquivo):
    try:
        
        
        # Criar uma lista para guardar os números das linhas como global
        global numeros_linhas
        numeros_linhas = set()

        # Ler o arquivo Excel
        df = pd.read_excel(caminho_arquivo, sheet_name="DADOS", engine="openpyxl")
        
        # Acessar as últimas 2000 linhas do DataFrame
        df_ultimas_linhas = df.tail(2000)
        
        # Acessar a coluna R nas últimas 200 linhas
        coluna_r = df_ultimas_linhas["E-mail enviado?"]

        # Percorrer a coluna R nas últimas 200 linhas
        for indice, valor in coluna_r.items():
            if coluna_r.isna().any():
                pass
            if valor == "OK":
                # Se tiver um "OK" na linha, acessar a coluna S
                valor_coluna_s = df_ultimas_linhas.at[indice, "Postal?"]
                if valor_coluna_s == "POSTAL":
                    # Se a coluna T (Postal enviado?) estiver vazia, adiciona o número da ficha à lista
                    valor_coluna_t = df_ultimas_linhas.at[indice, "Postal enviado?"]
                    if pd.isna(valor_coluna_t):
                        numero_linha = df_ultimas_linhas.at[indice, "Ficha"]
                          # Converter a coluna C para o tipo inteiro
                        numeros_linhas.add(int(numero_linha))
        if len(numeros_linhas) == 0:
            messagebox.showinfo("Postal Atualizado", "Não há fichas no momento")
        else:
            print(numeros_linhas)
            label_lista_fichas.config(text=" Fichas para postal sepadaras.")
            return numeros_linhas
    except Exception as erro:
        messagebox.showerror('ERRO!', f"{erro}")




#Selecionar pasta para copiar os pdfs do set numeros_linhas
def selecionar_pasta_origem():
    try:
        global pasta_origem
        pasta_origem = filedialog.askdirectory()
        pasta_origem = rf"{pasta_origem}"
        print(f"Pasta de origem selecionada: {pasta_origem}")
        nome_caminho_origem = os.path.split(pasta_origem)
        if pasta_origem == '':
            messagebox.showwarning("Caminho não selecionado.", "Por favor, selecione um caminho para continuar.")
        else:
            label_pasta_origem.config(text=f"Pasta de origem: {nome_caminho_origem[1]}")
    except Exception as erro:
        messagebox.showerror('ERRO!', f"{erro}")




#Selecionar pasta para salvar os pdfs copiados do set numeros_linhas      
def selecionar_pasta_destino():
    try:
        global pasta_destino
        pasta_destino = filedialog.askdirectory()
        pasta_destino = rf"{pasta_destino}"
        
        print(f"Pasta de destino selecionada: {pasta_destino}")
        nome_caminho_destino = os.path.split(pasta_destino)
        if pasta_destino == '':
            messagebox.showwarning("Caminho não selecionado.", "Por favor, selecione um caminho para continuar.")
        else:
            label_pasta_destino.config(text=f"Pasta de destino: {nome_caminho_destino[1]}")
            
    except Exception as erro:
        messagebox.showerror('ERRO!', f"{erro}")
    



#Função que copia os pdfs na pasta selecionada em pasta_destino
def copia_pdfs(pasta_origem, pasta_destino, numeros_linhas):
    pdfs_nao_encontrados = set()
    pdfs_encontrados = set()
    arquivos_pdf = set(map(str, numeros_linhas))
    try:
       
        pasta = os.listdir(pasta_origem)

        for numero_ficha in arquivos_pdf:
           
            if numero_ficha + ".pdf" in pasta:
                caminho_ficha = os.path.join(pasta_origem, numero_ficha + ".pdf")
                shutil.copy(caminho_ficha, pasta_destino)
                print(f'Arquivo {numero_ficha + ".pdf"} copiado com sucesso para {pasta_destino}')
                pdfs_encontrados.add(numero_ficha)
                print(pdfs_encontrados)

            else:
                pdfs_nao_encontrados.add(numero_ficha)
    
        if len(pdfs_nao_encontrados) != 0:
            messagebox.showwarning("Aviso", f"PDFs não encontrados:{pdfs_nao_encontrados}")
            print(pdfs_nao_encontrados)
            escreve_arquivo_não_encontrado(pasta_destino, pdfs_nao_encontrados)

        if len(pdfs_encontrados) != 0:
            messagebox.showinfo("Pdfs Postais",f"PDFs prontos para impressão: {len(pdfs_encontrados)} ficha(s)")
            baixa_postal(caminho_arquivo, pdfs_encontrados)
            ficha = fichas_nome(caminho_arquivo, pdfs_encontrados)
            escreve_arquivo(pasta_destino, ficha, caminho_arquivo)
        

        label_executar.config(text=f"Programa encerrado")
       
    except Exception as erro:
        messagebox.showerror('ERRO!', f"{erro}")
    



# Cria uma janela 
janela = tk.Tk()
# Carrega a imagem de fundo do postal
imagem_path = "postal-image.gif"


imagem = tk.PhotoImage(file=imagem_path)

label_lista_fichas = tk.Label(janela, text='')
label_pasta_destino = tk.Label(janela, text='')
label_pasta_origem = tk.Label(janela, text='')
label_selecionar_arq = tk.Label(janela, text='')
label_executar = tk.Label(janela, text='')

# carrega a imagem de fundo do postal
imagem = tk.PhotoImage(file=imagem_path)

# define o caminho para o ícone da janela
icone_path = 'ww.ico'
janela.iconbitmap(icone_path)

# Define o título da janela com o icone do Heisenberg
janela.title("PDF Maverick")

# Redimensiona a imagem para que ela se ajuste à janela
imagem = imagem.subsample(7)  # ajusta a imagem, quanto maior o numero, menor ela será

# Cria um widget ttk.Label e adiciona a imagem como seu conteúdo
label_imagem = ttk.Label(janela, image=imagem)
label_imagem.place(x=0, y=0, relwidth=1, relheight=1) # posicionar a imagem na janela


label_selecionar_arq = tk.Label(janela, text="")
label_selecionar_arq.place(x=40, y=150)

label_lista_fichas = tk.Label(janela, text="")
label_lista_fichas.place(x=40, y=240)

label_pasta_origem = tk.Label(janela, text="")
label_pasta_origem.place(x=50, y=330)    

label_pasta_destino = tk.Label(janela, text="")
label_pasta_destino.place(x=50, y=430)

label_executar = tk.Label(janela, text="")
label_executar.place(x=440, y=420)


#novo comando de botão do chaGPT
botao_selecionar_arq = ttk.Button(janela, text="1.Selecionar arquivo", style='W.TButton', command=selecionar_arquivo)
botao_selecionar_arq.place(x=30, y=90, width=200, height=50)

estilo = ttk.Style()
estilo.configure('W.TButton', font=("Arial", 12, "italic"), foreground="#DC143C", anchor="center")


botao_salvar_fichas = ttk.Button(janela, text="2.Selecionar Fichas", style='W.TButton', command=lambda: lista_fichas(caminho_arquivo))
botao_salvar_fichas.place(x=30, y=180, width=200, height=50)

botao_pasta_origem = ttk.Button(janela, text="3.Pasta resultados", style='W.TButton', command=selecionar_pasta_origem)
botao_pasta_origem.place(x=30, y=270, width=200, height=50)

botao_pasta_destino = ttk.Button(janela, text="4.Pasta postal pdfs", style='W.TButton', command=selecionar_pasta_destino)
botao_pasta_destino.place(relx= 0.18, rely=0.8, anchor="center", width=205, height=50)


botao_executar = ttk.Button(janela, text="5.Executar", style='W.TButton', command=lambda:copia_pdfs(pasta_origem, pasta_destino, numeros_linhas))
botao_executar.place(relx= 0.55, rely=0.7, width=200, height=50)

# Define o tamanho da janela
janela.geometry("720x490")

# Desabilita a opção de maximizar a janela
janela.resizable(False, False)


# Executa a janela principal
janela.mainloop()
