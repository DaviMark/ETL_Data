import customtkinter as ctk
from tkinter import filedialog, messagebox
import pandas as pd
import logging

# Configuração de logging
logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")

# Definição das opções de GR
GR_OPCOES = ["Skymark_Usiminas", "Opentech", "Klios_Cadastros", "Klios_Consulta", "J&C", "Skymark"]

# Caminho do arquivo de destino
CAMINHO_DESTINO = r"Local_excel_para_enviar.xlsx"

class Aplicacao(ctk.CTk):
    def __init__(self):
        super().__init__()
        
        # Configuração da janela principal
        self.title("Gerenciador de Dados Excel")
        self.geometry("480x320")
        self.resizable(False, False)
        
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")
        
        self.criar_interface()

    def criar_interface(self):
        """Cria a interface gráfica do aplicativo."""
        ctk.CTkLabel(self, text="Selecione a GR:", font=("Poppins", 16)).pack(pady=10)
        self.dropdown = ctk.CTkComboBox(self, values=GR_OPCOES, font=("Poppins", 14))
        self.dropdown.pack(pady=10)
        
        ctk.CTkLabel(self, text="Arquivo Excel:", font=("Poppins", 14)).pack(pady=5)
        self.entry_arquivo = ctk.CTkEntry(self, font=("Poppins", 14), width=320)
        self.entry_arquivo.pack(pady=5)
        
        ctk.CTkButton(self, text="Selecionar Arquivo", command=self.selecionar_arquivo, font=("Poppins", 14)).pack(pady=10)
        ctk.CTkButton(self, text="Inserir Dados", command=self.inserir_dados, font=("Poppins", 14)).pack(pady=10)

    def selecionar_arquivo(self):
        """Abre um seletor de arquivos para escolher um arquivo Excel."""
        caminho_arquivo = filedialog.askopenfilename(title="Selecionar arquivo Excel", filetypes=[("Arquivos Excel", "*.xlsx")])
        if caminho_arquivo:
            self.entry_arquivo.delete(0, ctk.END)
            self.entry_arquivo.insert(0, caminho_arquivo)

    def inserir_dados(self):
        """Processa e insere os dados do arquivo selecionado na planilha correspondente."""
        gr_selecionada = self.dropdown.get()
        caminho_arquivo = self.entry_arquivo.get()

        if not gr_selecionada or not caminho_arquivo:
            messagebox.showerror("Erro", "Selecione uma GR e um arquivo antes de continuar!")
            return
        
        try:
            df_origem = self.carregar_dados_origem(caminho_arquivo, gr_selecionada)
            df_origem["GR_Utilizada"] = gr_selecionada
            
            df_destino = self.carregar_dados_existentes(gr_selecionada)
            df_destino = pd.concat([df_destino, df_origem], ignore_index=True)
            
            self.salvar_dados(df_destino, gr_selecionada)
            messagebox.showinfo("Sucesso", f"Dados inseridos com sucesso na planilha {gr_selecionada}!")

        except Exception as e:
            logging.error(f"Erro ao inserir dados: {e}")
            messagebox.showerror("Erro", f"Ocorreu um erro: {str(e)}")

    def carregar_dados_origem(self, caminho_arquivo, gr):
        """Carrega e filtra os dados do arquivo Excel de origem conforme a GR selecionada."""
        try:
            xls = pd.ExcelFile(caminho_arquivo)
            df = pd.read_excel(xls)

            filtros = {
                "J&C": lambda df: df[df.iloc[:, 1].notnull()].iloc[1:].reset_index(drop=True),
                "Opentech": lambda df: df.iloc[1:].reset_index(drop=True),
                "Klios_Cadastros": lambda df: self.carregar_planilha(xls, "CADASTRO"),
                "Klios_Consulta": lambda df: self.carregar_planilha(xls, "CONSULTA"),
                "Skymark_Usiminas": lambda df: df[df.iloc[:, 0].notnull() & (df.iloc[:, 0] != "Total:")].iloc[1:].reset_index(drop=True),
                "Skymark": lambda df: df[df.iloc[:, 0].notnull() & (df.iloc[:, 0] != "Total:")].iloc[1:].reset_index(drop=True)
            }

            return filtros.get(gr, lambda df: df)(df)

        except Exception as e:
            raise ValueError(f"Erro ao carregar dados do arquivo: {str(e)}")

    def carregar_planilha(self, xls, sheet_name):
        """Carrega uma planilha específica de um arquivo Excel."""
        if sheet_name not in xls.sheet_names:
            raise ValueError(f"A planilha '{sheet_name}' não foi encontrada no arquivo!")
        df = pd.read_excel(xls, sheet_name=sheet_name)
        return df[df.iloc[:, 1].notnull()].iloc[1:].reset_index(drop=True)

    def carregar_dados_existentes(self, gr):
        """Carrega os dados existentes no arquivo de destino, se houver."""
        try:
            xls_destino = pd.ExcelFile(CAMINHO_DESTINO)
            return pd.read_excel(xls_destino, sheet_name=gr)
        except (FileNotFoundError, ValueError):
            return pd.DataFrame()

    def salvar_dados(self, df, gr):
        """Salva os dados processados no arquivo de destino."""
        try:
            with pd.ExcelWriter(CAMINHO_DESTINO, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
                df.to_excel(writer, sheet_name=gr, index=False)
        except Exception as e:
            raise ValueError(f"Erro ao salvar os dados: {str(e)}")

if __name__ == "__main__":
    app = Aplicacao()
    app.mainloop()
