import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from database import inserir_dados_banco_brasil, inserir_dados_protheus, analise_financeira, inserir_dados_itau
from utils import formatar_valor, formatar_valor_protheus

# Variáveis globais
banco_path = None

def selecionar_banco():
    global banco_path
    banco_path = filedialog.askopenfilename(title="Selecionar arquivo Banco do Brasil", 
                                             filetypes=[("Excel files", ".xlsx;.xls")])
    if banco_path:
        try:
            df = pd.read_excel(banco_path, header=None)

            # Pegar a agência e conta da linha 2
            agencia = df.iloc[1, 1]  # Coluna B linha 2
            conta = df.iloc[1, 3]    # Coluna D linha 2

            # Filtrar os dados apenas onde a coluna J (índice 9) é "D"
            dados = df.iloc[3:, [0, 7, 8, 9]].copy()  # Colunas A, H, I, J
            dados.columns = ['Data', 'Historico', 'Valor R$', 'Tipo']
            dados['Valor R$'] = dados['Valor R$'].apply(formatar_valor)

            # Filtrar apenas as linhas onde Tipo é "D"
            dados = dados[dados['Tipo'] == 'D']
            dados = dados[dados['Valor R$'].notna()]  # Remover linhas com valores não válidos

            # Inserir dados na tabela do Banco do Brasil
            total_inseridos = inserir_dados_banco_brasil(dados, agencia, conta)

            if total_inseridos > 0:
                messagebox.showinfo("Sucesso", f"{total_inseridos} dados inseridos com sucesso na tabela do Banco do Brasil.")
            else:
                messagebox.showinfo("Aviso", "Nenhum dado foi inserido, pois todos já estão no banco de dados.")

        except Exception as e:
            messagebox.showerror("Erro", str(e))

def selecionar_itau():
    global itau_path
    itau_path = filedialog.askopenfilename(title="Selecionar arquivo Itaú", 
                                            filetypes=[("Excel files", ".xlsx;.xls")])
    if itau_path:
        try:
            df = pd.read_excel(itau_path, header=None)

            # Pegar a agência e conta da célula B4 e B5
            agencia = df.iloc[3, 1]  # Coluna B linha 4
            conta = df.iloc[4, 1]    # Coluna B linha 5
            
            # Filtrar os dados a partir da linha 10
            dados = df.iloc[9:, [0, 3, 4]].copy()  # Colunas A, D, E
            dados.columns = ['Data', 'Razão Social', 'Valor (R$)']
            
            # Filtra apenas os valores negativos e formata
            dados['Valor (R$)'] = dados['Valor (R$)'].apply(lambda x: formatar_valor_protheus(x))
            dados = dados[dados['Valor (R$)'].notna() & (dados['Valor (R$)'] < 0)]  # Remove valores não válidos e positivos

            # Remover o sinal negativo
            dados['Valor (R$)'] = dados['Valor (R$)'].abs()

            # Verifica se há dados a serem inseridos
            if not dados.empty:
                total_inseridos = inserir_dados_itau(dados, agencia, conta)  # Passa o DataFrame completo

                if total_inseridos > 0:
                    messagebox.showinfo("Sucesso", f"{total_inseridos} dados inseridos com sucesso na tabela do Itaú.")
                else:
                    messagebox.showinfo("Aviso", "Nenhum dado foi inserido, pois todos já estão no banco de dados.")
            else:
                print("Nenhum dado válido encontrado.")

        except Exception as e:
            messagebox.showerror("Erro", str(e))

def selecionar_protheus():
    global protheus_path
    protheus_path = filedialog.askopenfilename(title="Selecionar arquivo Protheus", 
                                                filetypes=[("Excel files", ".xlsx;.xls")])
    if protheus_path:
        try:
            df = pd.read_excel(protheus_path, sheet_name=1)  # Lê a segunda aba

            # Filtrar e formatar os dados
            dados = df[['DATA', 'PREFIXO/TITULO', 'SAIDAS']].copy()
            dados['SAIDAS'] = dados['SAIDAS'].apply(formatar_valor_protheus)

            # Remover linhas com valores não válidos ou que são zero
            dados = dados[dados['SAIDAS'].notna() & (dados['SAIDAS'] > 0)]

            # Inserir dados na tabela do Protheus
            total_inseridos = inserir_dados_protheus(dados)

            if total_inseridos > 0:
                messagebox.showinfo("Sucesso", f"{total_inseridos} dados inseridos com sucesso na tabela do Protheus.")
            else:
                messagebox.showinfo("Aviso", "Nenhum dado foi inserido, pois todos já estão no banco de dados.")

        except Exception as e:
            messagebox.showerror("Erro", str(e))

def criar_interface():
    root = tk.Tk()
    root.title("Análise Financeira - Sagg Consultoria")
    root.geometry("650x400")  
    root.configure(bg="#f0f4ff")  

    # Definindo a lista de meses
    meses = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro", "Geral"]

    # Título da aplicação
    title_label = tk.Label(root, text="Análise Financeira - Sagg Consultoria", font=("Arial", 18, "bold"), bg="#f0f4ff", fg="#003366")
    title_label.pack(pady=15)

    # Criação do Notebook para as abas
    notebook = ttk.Notebook(root)
    notebook.pack(pady=10, expand=True, fill='both')

    # Estilizando o Notebook
    style = ttk.Style()
    style.configure("TFrame", background="#f0f4ff")
    style.configure("TNotebook", background="#f0f4ff", borderwidth=0)
    style.configure("TButton", background="#66b3ff", foreground="white", font=("Arial", 12, "bold"), borderwidth=1, relief="flat")
    style.map("TButton", background=[("active", "#3399ff")])  

    # Aba 1: Seleção de arquivos
    aba_selecao = ttk.Frame(notebook)
    notebook.add(aba_selecao, text="Inserir Extratos")

    # Criando um Frame para centralizar os botões
    frame_botoes = tk.Frame(aba_selecao, bg="#f0f4ff")
    frame_botoes.pack(expand=True)

    # Botões para selecionar os arquivos
    tk.Button(frame_botoes, text="Inserir Extrato BB", command=selecionar_banco, width=40, height=2).pack(pady=10)
    tk.Button(frame_botoes, text="Inserir Extrato Itaú", command=selecionar_itau, width=40, height=2).pack(pady=10)
    tk.Button(frame_botoes, text="Inserir Extrato Protheus", command=selecionar_protheus, width=40, height=2).pack(pady=10)

    # Aba 2: Gerar Análise Financeira
    aba_analise = ttk.Frame(notebook)
    notebook.add(aba_analise, text="Gerar Análise Financeira")

    # Labels e Dropdowns lado a lado
    frame_dropdowns = tk.Frame(aba_analise, bg="#f0f4ff")
    frame_dropdowns.pack(pady=10)

    tk.Label(frame_dropdowns, text="Selecione o mês:", bg="#f0f4ff", font=("Arial", 12)).grid(row=0, column=0, padx=10)
    dropdown_mes = ttk.Combobox(frame_dropdowns, values=meses, state="readonly", width=20)
    dropdown_mes.grid(row=0, column=1, padx=10)
    dropdown_mes.current(0)

    tk.Label(frame_dropdowns, text="Selecione o ano:", bg="#f0f4ff", font=("Arial", 12)).grid(row=0, column=2, padx=10)
    anos = [2024, 2025, 2026]
    dropdown_ano = ttk.Combobox(frame_dropdowns, values=anos, state="readonly", width=20)
    dropdown_ano.grid(row=0, column=3, padx=10)

    # Botão para gerar análise
    tk.Button(aba_analise, text="Gerar Análise Financeira", command=lambda: analise_financeira(dropdown_mes.get(), dropdown_ano.get(), meses), width=40, height=2).pack(pady=20)

    # Rodapé
    footer_label = tk.Label(root, text="Inovação Serviços © 2024", font=("Arial", 10), bg="#f0f4ff", fg="#003366")
    footer_label.pack(side="bottom", pady=15)

    root.mainloop()

if __name__ == '__main__':
    criar_interface()