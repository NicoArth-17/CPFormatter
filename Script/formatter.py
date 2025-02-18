import tkinter as tk
from tkinter.filedialog import askopenfilename, askdirectory
import pandas as pd
from itertools import zip_longest
import os

# Cria janela
window = tk.Tk()

window.configure(bg='gray')

window.rowconfigure(0, weight=1)
window.columnconfigure(0, weight=1)

window.title('CPF Formatter')

texto1 = tk.Label(text='Insira um arquivo excel com a listagem de CPFs na primeira coluna:', bg='gray', fg='white')
texto1.grid(row=0, column=0, sticky='nswe', padx=10, pady=10)

arquivo = tk.StringVar()

def anexar_arquivo():

    caminho_input = askopenfilename(title='Encontre o arquivo')
    caminho_input_name = os.path.basename(caminho_input)

    arquivo.set(caminho_input)

    if caminho_input:

        result1.config(text='')
        result1.config(text=caminho_input_name, fg='black')

    
def formatar_cpf():

    try:
        df = pd.read_excel(arquivo.get())

        cpfs = df.iloc[:,0]

        cpfs_formatados = []

        cpfinvalido = []


        for cpf in cpfs:

            cpf = str(cpf)

            if len(cpf) == 11: 
                cpf = f'{cpf[:3]}.{cpf[3:6]}.{cpf[6:9]}-{cpf[9:]}'
                cpfs_formatados.append(cpf)

            elif len(cpf) == 10:
                cpf = f'0{cpf[:2]}.{cpf[2:5]}.{cpf[5:8]}-{cpf[8:]}'
                cpfs_formatados.append(cpf)

            elif len(cpf) == 9:
                cpf = f'00{cpf[:1]}.{cpf[1:4]}.{cpf[4:7]}-{cpf[7:]}'
                cpfs_formatados.append(cpf)

            elif len(cpf) == 8:
                cpf = f'000.{cpf[:3]}.{cpf[3:6]}-{cpf[6:]}'
                cpfs_formatados.append(cpf)

            else:
                cpfinvalido.append(cpf)
                cpfs_formatados.append(cpf)

        print(cpfs_formatados)
        print(cpfinvalido)

        caminho_output = askdirectory(title='Escolha a pasta para armazernar os CPFs formatados')

        new_df = pd.DataFrame(list(zip_longest(cpfs_formatados, cpfinvalido)), columns=['CPF', 'CPF Inválido'])
        new_df.to_excel(f'{caminho_output}/Formatted CPFs.xlsx', index=None)

        result2.config(text='')
        result2.config(text='Processo concluído', fg='black', bg='green')

        result1.config(text='Arquivo selecionado')

    except:

        result2.config(text='')
        result2.config(text='Anexe um arquivo válido', fg='black', bg='red')


anexo = tk.Button(text='Anexar', command=anexar_arquivo)
anexo.grid(row=1, column=0, padx=10, pady=10)

result1 = tk.Label(text='Arquivo selecionado', fg='gray', borderwidth=2, relief='solid')
result1.grid(row=2, column=0, sticky='nswe', padx=10, pady=10)

formatar = tk.Button(text='Executar', command=formatar_cpf)
formatar.grid(row=3, column=0)

result2 = tk.Label(text='log', fg='gray', borderwidth=2, relief='solid')
result2.grid(row=4, column=0, sticky='nswe', padx=10, pady=10)

# Mantem a janela aberta
window.mainloop()