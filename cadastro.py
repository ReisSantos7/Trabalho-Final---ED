import tkinter as tk
from tkinter import ttk
import openpyxl

# Função para carregar os dados do Excel
def load_data():
    path = r"C:\Users\Administrador\Desktop\formulario\funcionarios.xlsx"
    workbook = openpyxl.load_workbook(path)
    sheet = workbook.active

    list_values = list(sheet.values)
    print(list_values)

    for col_name in cols:
        treeview.heading(col_name, text=col_name)

    if list_values:
        for col_name in cols:
            treeview.heading(col_name, text=col_name)

        for value_tuple in list_values[1:]:
            treeview.insert('', tk.END, values=value_tuple)

# Função para inserir dados
def insert_data():
    nome = name_entry.get()
    endereço = endereço_entry.get()
    idade = int(idade_spinbox.get())
    genero = status_genero.get()
    nação_value = "Estrangeiro" if nação.get() else "Nativo"


    # Função salvar dados no banco
    path = r"C:\Users\Administrador\Desktop\formulario\funcionarios.xlsx"
    workbook = openpyxl.load_workbook(path)
    sheet = workbook.active
    row_values = [nome, endereço, idade, genero, nação_value]
    sheet.append(row_values)
    workbook.save(path)

    # Função inserir dado salvo na treeview
    treeview.insert('', tk.END, values=row_values)

    #Limpar e colocos os dados iniciais dos campos
    name_entry.delete(0, 'end')
    name_entry.insert(0, "Nome")

    idade_spinbox.delete(0, 'end')
    idade_spinbox.insert(0, "Idade")

    status_genero.set(genero_list[0])

    checkbutton.state(["!selected"])
# Função para alterar o tema
def modo_tema():
    if modo_switch.instate(["selected"]):
        style.theme_use('forest-light')
    else:
        style.theme_use('forest-dark')

# Iniciando a janela
root = tk.Tk()
root.title('Cadastro de Parceiros')

# Configuração do tema
style = ttk.Style(root)
root.tk.call("source", "forest-light.tcl")
root.tk.call("source", "forest-dark.tcl")
style.theme_use("forest-dark")

# Principal
frame = ttk.Frame(root)
frame.pack()

widgets_frame = ttk.LabelFrame(frame, text="Insira os dados")
widgets_frame.grid(row=0, column=0, padx=20, pady=10)

# Criação de campos do formulário de cadastro
name_entry = ttk.Entry(widgets_frame)
name_entry.insert(0, "Nome")
name_entry.bind("<FocusIn>", lambda e: name_entry.delete('0', 'end'))
name_entry.grid(row=0, column=0, padx=5, pady=(30, 5), sticky='ew')

endereço_entry = ttk.Entry(widgets_frame)
endereço_entry.insert(0, "Endereço")
endereço_entry.bind("<FocusIn>", lambda e: endereço_entry.delete('0', 'end'))
endereço_entry.grid(row=1, column=0, padx=5, pady=(10, 5), sticky='ew')

idade_spinbox = ttk.Spinbox(widgets_frame, from_=1, to=100)
idade_spinbox.insert(0, "Idade")
idade_spinbox.grid(row=2, column=0, padx=5, pady=5, sticky='ew')

genero_list = ["Masculino", "Feminino", "Outro"]
status_genero = ttk.Combobox(widgets_frame, values=genero_list)
status_genero.current(0)
status_genero.grid(row=3, column=0, padx=5, pady=5, sticky='ew')

nação = tk.BooleanVar()
checkbutton = ttk.Checkbutton(widgets_frame, text="Estrangeiro?", variable=nação)
checkbutton.grid(row=4, column=0, padx=5, pady=5, sticky='nsew')

botão = ttk.Button(widgets_frame, text='Gravar os Dados'.upper(), command=insert_data)
botão.grid(row=5, column=0, padx=5, pady=5, sticky='nesw')

separador = ttk.Separator(widgets_frame)
separador.grid(row=6, column=0, padx=5, pady=10, sticky='ew')

modo_switch = ttk.Checkbutton(widgets_frame, text='Modo', style='Switch', command=modo_tema)
modo_switch.grid(row=7, column=0, padx=5, pady=10, sticky='nsew')

# Criando preview para dados do banco
treeFrame = ttk.Frame(frame)
treeFrame.grid(row=0, column=1, padx=(0, 20), pady=10)
treeScroll = ttk.Scrollbar(treeFrame)
treeScroll.pack(side="right", fill="y")

cols = ("Nome", "Endereço", "Idade", "Genero", "Nação")
treeview = ttk.Treeview(treeFrame, show="headings", yscrollcommand=treeScroll.set, columns=cols, height=13)

for col in cols:
    treeview.column(col, width=100)
    treeview.heading(col, text=col)

treeview.pack()
treeScroll.config(command=treeview.yview)

load_data()

# Rodando a janela
root.mainloop()