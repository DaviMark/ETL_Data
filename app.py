import customtkinter as ctk  # Biblioteca para interface gráfica personalizada
from tkinter import filedialog, messagebox  # Módulos para manipulação de arquivos e mensagens
import pandas as pd  # Biblioteca para manipulação de arquivos Excel

# Opções de GR disponíveis no dropdown
GR_OPCOES = ["Skymark_Usiminas", "Opentech", "Klios_Cadastros","Klios_Consulta", "J&C","Skymark"]

# Caminho do arquivo Excel onde os dados serão inseridos
CAMINHO_DESTINO = r"\\192.168.0.254\datapar\TI\GR-Boletos.xlsx"


def selecionar_arquivo():
    """Abre uma janela para selecionar um arquivo Excel e exibe o caminho no campo de entrada."""
    caminho_arquivo = filedialog.askopenfilename(
        title="Selecionar arquivo Excel",
        filetypes=[("Arquivos Excel", "*.xlsx")]  # Permite apenas arquivos com extensão .xlsx
    )
    
    # Se um arquivo for selecionado, insere o caminho no campo de entrada
    if caminho_arquivo:
        entry_arquivo.delete(0, ctk.END)  # Limpa o campo antes de inserir um novo caminho
        entry_arquivo.insert(0, caminho_arquivo)  # Insere o caminho do arquivo


def inserir_dados():
    """Lê os dados do arquivo Excel selecionado e insere na planilha correspondente à GR escolhida."""
    
    gr_selecionada = dropdown.get()  # Obtém a GR selecionada pelo usuário
    caminho_arquivo = entry_arquivo.get()  # Obtém o caminho do arquivo selecionado

    # Verifica se o usuário selecionou uma GR e um arquivo antes de continuar
    if not gr_selecionada or not caminho_arquivo:
        messagebox.showerror("Erro", "Selecione uma GR e um arquivo antes de continuar!")
        return

    try:
        # Carrega o arquivo Excel
        xls = pd.ExcelFile(caminho_arquivo)
        
        # Lê o arquivo uma vez
        df_origem = pd.read_excel(xls)

        # Filtro específico para "J&C" (segunda coluna não nula)
        if gr_selecionada == "J&C":
            df_origem = df_origem[df_origem.iloc[:, 1].notnull()]  # Remove linhas onde a segunda coluna é nula
            df_origem = df_origem.iloc[1:].reset_index(drop=True)  # Remove a primeira linha e reajusta os índices

        # Filtro específico para "Opentech" (remove a primeira linha e reajusta os índices)
        elif gr_selecionada == "Opentech":
            df_origem = df_origem.iloc[1:].reset_index(drop=True)  # Remove a primeira linha e reajusta os índices

        # Filtro específico para "Klios" (remove a primeira linha e reajusta os índices)
        elif gr_selecionada == "Klios_Cadastros":
            # Verifica se a planilha se chama "CADASTRO"
            if 'CADASTRO' in xls.sheet_names:
                # Carrega a planilha "CADASTRO"
                df_origem = pd.read_excel(xls, sheet_name="CADASTRO")
                
                # Realiza o filtro para remover linhas onde a 18ª coluna (índice 17) é nula
                df_origem = df_origem[df_origem.iloc[:, 17].notnull()]  # Remove linhas onde a 18ª coluna (índice 17) é nula
                df_origem = df_origem[df_origem.iloc[:, 1].notnull()]  # Remove linhas onde a 2ª coluna (índice 1) é nula
                
                # Caso seja necessário, remova a primeira linha e reajuste os índices
                df_origem = df_origem.iloc[1:].reset_index(drop=True)  # Remove a primeira linha e reajusta os índices
            else:
                messagebox.showerror("Erro", "A planilha 'CADASTRO' não foi encontrada no arquivo Excel selecionado!")
                return
        
        # Filtro específico para "Klios_Consulta" (remove a primeira linha e reajusta os índices)
        elif gr_selecionada == "Klios_Consulta":
            # Verifica se a planilha se chama "CONSULTA"
            if 'CONSULTA' in xls.sheet_names:
                # Carrega a planilha "CONSULTA"
                df_origem = pd.read_excel(xls, sheet_name="CONSULTA")
                
                # Realiza o filtro para remover linhas onde a 18ª coluna (índice 17) é nula
                df_origem = df_origem[df_origem.iloc[:, 2].notnull()]  # Remove linhas onde a 18ª coluna (índice 17) é nula
                
                # Caso seja necessário, remova a primeira linha e reajuste os índices
                df_origem = df_origem.iloc[1:].reset_index(drop=True)  # Remove a primeira linha e reajusta os índices
            else:
                messagebox.showerror("Erro", "A planilha 'CONSULTA' não foi encontrada no arquivo Excel selecionado!")
                return


        elif gr_selecionada == "Skymark_Usiminas":
            df_origem = df_origem[df_origem.iloc[:, 0].notnull()]  # Remove linhas onde a primeira coluna é nula
            df_origem = df_origem[df_origem.iloc[:, 0] != "Total:"]  # Remove linhas que contenham "Total:"
            df_origem = df_origem.iloc[1:].reset_index(drop=True)  # Remove a primeira linha e reajusta os índices
        
        elif gr_selecionada == "Skymark":
            df_origem = df_origem[df_origem.iloc[:, 0].notnull()]  # Remove linhas onde a primeira coluna é nula
            df_origem = df_origem[df_origem.iloc[:, 0] != "Total:"]  # Remove linhas que contenham "Total:"
            df_origem = df_origem.iloc[1:].reset_index(drop=True)  # Remove a primeira linha e reajusta os índices

        # Adiciona uma coluna indicando a GR utilizada
        df_origem['GR_Utilizada'] = gr_selecionada

        try:
            # Tenta carregar os dados existentes na planilha de destino
            xls_destino = pd.ExcelFile(CAMINHO_DESTINO)
            df_destino = pd.read_excel(xls_destino, sheet_name=gr_selecionada)
        except (FileNotFoundError, ValueError):
            df_destino = pd.DataFrame()  # Se o arquivo não existir, cria um DataFrame vazio

        # Concatena os novos dados com os dados existentes
        df_destino = pd.concat([df_destino, df_origem], ignore_index=True)

        # Salva os dados no arquivo Excel de destino
        with pd.ExcelWriter(CAMINHO_DESTINO, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            df_destino.to_excel(writer, sheet_name=gr_selecionada, index=False)

        messagebox.showinfo("Sucesso", f"Dados inseridos com sucesso na planilha {gr_selecionada}!")
    except Exception as e:
        # Exibe uma mensagem de erro caso algo dê errado
        messagebox.showerror("Erro", f"Ocorreu um erro ao inserir os dados: {str(e)}")


# Configuração da interface gráfica
ctk.set_appearance_mode("dark")  # Define o modo escuro para a interface
ctk.set_default_color_theme("blue")  # Define o tema azul para a interface

app = ctk.CTk()  # Cria a janela principal
app.title("Inserir Dados no Excel")  # Define o título da janela
app.geometry("400x300")  # Define o tamanho da janela
app.resizable(False, False)  # Impede o redimensionamento da janela

# Rótulo para seleção da GR
label = ctk.CTkLabel(app, text="Selecione a GR utilizada:", font=("Arial", 16))
label.pack(pady=10)

# Dropdown para selecionar a GR
dropdown = ctk.CTkComboBox(app, values=GR_OPCOES, font=("Arial", 14))
dropdown.pack(pady=10)

# Rótulo para seleção do arquivo
label_arquivo = ctk.CTkLabel(app, text="Selecione um arquivo Excel:", font=("Arial", 14))
label_arquivo.pack(pady=5)

# Campo de entrada para exibir o caminho do arquivo selecionado
entry_arquivo = ctk.CTkEntry(app, font=("Arial", 14))
entry_arquivo.pack(pady=5)

# Botão para abrir a janela de seleção de arquivos
btn_selecionar = ctk.CTkButton(app, text="Selecionar Arquivo", command=selecionar_arquivo, font=("Arial", 14))
btn_selecionar.pack(pady=10)

# Botão para inserir os dados no Excel
btn_inserir = ctk.CTkButton(app, text="Inserir Dados", command=inserir_dados, font=("Arial", 14))
btn_inserir.pack(pady=10)

# Inicia o loop da interface gráfica
app.mainloop()
