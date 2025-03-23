# **📊 Automacao de Insercao de Dados em Excel**

Este projeto consiste em uma aplicacao desenvolvida com **Python** e **CustomTkinter** para facilitar a insercao de dados em planilhas Excel, aplicando filtros especificos para diferentes tipos de GRs (**Gerenciamento de Riscos**).

A ferramenta permite selecionar arquivos Excel, processa-los automaticamente e armazenar os dados na planilha de destino correta, garantindo uma estruturacao organizada e confiavel.

---

## **📌 Recursos Principais**

✅ Interface grafica moderna e responsiva utilizando **CustomTkinter**  
✅ Selecao de arquivos Excel atraves de um **explorador de arquivos integrado**  
✅ Aplicacao de **filtros inteligentes** para diferentes tipos de GRs  
✅ **Armazenamento automatico** de dados organizados no Excel  
✅ **Validacoes robustas** para evitar erros e inconsistencias  
✅ **Mensagens interativas** para informar o usuario sobre o status da operacao  

---

## **📂 Estrutura do Projeto**

```
📁 Automacao_Excel
│── 📄 main.py              # Codigo principal da aplicacao
│── 📄 requirements.txt     # Dependencias do projeto
│── 📄 README.md            # Documentacao do projeto
```

---

## **⚙️ Tecnologias Utilizadas**

📌 **Python** – Linguagem principal do projeto  
📌 **CustomTkinter** – Framework para interfaces graficas modernas  
📌 **pandas** – Manipulacao e tratamento de dados Excel  
📌 **openpyxl** – Leitura e escrita em arquivos `.xlsx`  
📌 **tkinter** – Para janelas de dialogo e mensagens interativas  

---

## **📥 Instalacao e Configuracao**

### **1️⃣ Clone este repositorio**  
```bash
git clone https://github.com/seuusuario/Automacao_Excel.git
cd Automacao_Excel
```

### **2️⃣ Instale as dependencias**  
```bash
pip install -r requirements.txt
```

### **3️⃣ Execute a aplicacao**  
```bash
python main.py
```

---

## **📌 Como Funciona?**

### **1️⃣ Selecione a GR**  
Escolha a categoria correta no menu suspenso (**Skymark_Usiminas, Opentech, Klios_Cadastros, Klios_Consulta, J&C, Skymark**).  

### **2️⃣ Escolha um Arquivo Excel**  
Clique no botao **"Selecionar Arquivo"** e escolha o arquivo `.xlsx` com os dados a serem processados.  

### **3️⃣ Insira os Dados**  
Clique no botao **"Inserir Dados"** para que o sistema processe o arquivo e insira as informacoes na planilha de destino correta.  

### **4️⃣ Sucesso!** 🎉  
Uma mensagem de confirmacao aparecera informando que os dados foram inseridos com sucesso!  

---

## **🔧 Funcionalidades Detalhadas**

### **🎯 Filtros Especificos para Cada GR**
🔹 **J&C:** Remove linhas onde a segunda coluna esta vazia e ajusta indices  
🔹 **Opentech:** Remove a primeira linha do arquivo  
🔹 **Klios_Cadastros:** Processa apenas a aba "CADASTRO", removendo linhas invalidas  
🔹 **Klios_Consulta:** Processa apenas a aba "CONSULTA", removendo valores inconsistentes  
🔹 **Skymark_Usiminas:** Remove linhas vazias e exclui totalizacoes  
🔹 **Skymark:** Remove linhas vazias e ajusta indices  

### **🛠 Tratamento de Erros e Validacoes**
✅ Verificacao se o usuario **selecionou um arquivo** antes de processar  
✅ **Mensagens de erro amigaveis** caso algo esteja incorreto  
✅ **Filtragem automatica** para manter apenas os dados relevantes  
✅ **Validacao de planilhas** antes da importacao  

---

## **🚀 Melhorias Futuras**

📌 Suporte para mais formatos de entrada alem de `.xlsx`  
📌 Melhor interface grafica com exibicao dos dados processados  
📌 Opcao de exportacao para diferentes formatos  
📌 Suporte para multiplos arquivos ao mesmo tempo  
