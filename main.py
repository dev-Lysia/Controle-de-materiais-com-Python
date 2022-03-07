# passo 1 --> importar bibliotecas
import tkinter as tk
from tkinter import ttk
import datetime as dt
import pandas as pd

materiais = pd.read_excel("Materiais.xlsx", engine= 'openpyxl')

lista_tipos = ["Unidade","Caixa"]
lista_codigos = []

# passo 2 --> criar a interface grafica
janela = tk.Tk()

# criação da função

def inserir_cadastro():

    descricao = entry_descricao.get()

    tipo = combobox_selecionar_tipo.get()

    quantidade = entry_quantidade.get()

    data_criacao = dt.datetime.now()
    data_criacao = data_criacao.strftime("%d/%m/%Y %H:%M")

    codigo = materiais.shape[0] + len(lista_codigos)+1
    codigo_str = f"COD-{codigo}"

    cliente = entry_cliente.get()

    cnpj = entry_cnpj.get()

    lista_codigos.append((codigo_str,descricao,tipo,quantidade,data_criacao,cliente,cnpj))


janela.title("Cadastros de Materias")

label_descricao = tk.Label(text="Descrição do Material")
label_descricao.grid(row= 1, column= 0, padx= 10,pady= 10, sticky='nswe', columnspan= 4)

entry_descricao = tk.Entry()
entry_descricao.grid(row=2, column=0, padx=10, pady=10, sticky='nswe', columnspan= 4)

label_tipo_unidade = tk.Label(text="Tipo da unidade do material")
label_tipo_unidade.grid(row= 3, column= 0, padx= 10, pady= 10, sticky='nswe',columnspan= 2)

combobox_selecionar_tipo = ttk.Combobox(values=lista_tipos)
combobox_selecionar_tipo.grid(row= 3, column= 2, padx= 10, pady= 10, sticky='nswe',columnspan= 2)

label_quantidade = tk.Label(text="Quantidade na unidade de material")
label_quantidade.grid(row= 4, column= 0, padx= 10, pady= 10, sticky='nswe',columnspan= 2)

entry_quantidade = tk.Entry()
entry_quantidade.grid(row= 4, column= 2, padx= 10, pady= 10, sticky='nswe',columnspan= 2)

label_cliente = tk.Label(text="Nome do cliente: ")
label_cliente.grid(row=5, column=0, padx=10, pady=10, sticky='nswe', columnspan= 4)

entry_cliente = tk.Entry()
entry_cliente.grid(row=6, column=0, padx=10, pady=10, sticky='nswe', columnspan= 4)

label_cnpj = tk.Label(text="CNPJ do cliente:")
label_cnpj.grid(row=7, column=0, padx=10, pady=10, sticky='nswe', columnspan= 4)

entry_cnpj=tk.Entry()
entry_cnpj.grid(row=8, column=0, padx=10, pady=10, sticky='nswe', columnspan= 4)

botao_criar_codigo = tk.Button(text='Criar cadastro', command=inserir_cadastro)
botao_criar_codigo.grid(row=9,column= 0, padx=10, pady=10, sticky='nswe', columnspan= 4 )

janela.mainloop()


# adicionar um novo material para a planilha

novo_material = pd.DataFrame(lista_codigos, columns=['Código','Descrição','Tipo','Quantidade','Data Criação','Cliente','CNPJ'])
materiais = materiais.append(novo_material, ignore_index=True)

# atualizar o excel
materiais.to_excel('Materiais.xlsx',index=False)
