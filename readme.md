# 📌 Inserção de Dados no Excel

## 📋 Descrição do Projeto

Este projeto consiste em um script Python que automatiza a inserção de dados em uma planilha Excel compartilhada. O usuário pode selecionar um arquivo Excel contendo informações, escolher algumas informações específica e inserir os dados filtrados no arquivo de destino.

## 🚀 Funcionalidades

- Seleção de um arquivo Excel para processamento.
- Escolha de uma opção específica através de um dropdown.
- Processamento e filtragem de dados conforme a GR selecionada.
- Inserção dos dados na planilha correspondente dentro do arquivo de destino.
- Interface gráfica amigável utilizando a biblioteca `customtkinter`.
- Mensagens de erro e sucesso para melhor experiência do usuário.

## 🛠️ Tecnologias Utilizadas

- Python 3
- `customtkinter` – Interface gráfica personalizada
- `pandas` – Manipulação de dados
- `openpyxl` – Manipulação de arquivos Excel
- `tkinter` – Para manipulação de arquivos e exibição de mensagens

## 📂 Estrutura do Código

1. **Definição de constantes:**
   - Lista de opções disponíveis.
   - Caminho do arquivo Excel onde os dados serão armazenados.

2. **Funções principais:**
   - `selecionar_arquivo()`: Abre um seletor de arquivos e insere o caminho do arquivo selecionado na interface.
   - `inserir_dados()`: Processa o arquivo Excel escolhido, aplica filtros específicos dependendo da GR e insere os dados no arquivo de destino.

3. **Interface gráfica:**
   - Criada com `customtkinter`, inclui:
     - Dropdown para selecionar a GR.
     - Campo de entrada para exibição do caminho do arquivo selecionado.
     - Botões para selecionar o arquivo e inserir os dados.

## 🏗️ Como Executar

1. Instale as dependências necessárias:
   ```sh
   pip install customtkinter pandas openpyxl
   ```
2. Execute o script Python:
   ```sh
   python nome_do_script.py
   ```
3. Selecione a GR desejada, escolha um arquivo Excel e clique em "Inserir Dados".

## ⚠️ Considerações

- O script deve ter acesso à unidade de rede especificada ou local onde será enviado.
- Certifique-se de que o arquivo Excel de destino não esteja aberto ao executar o script.


## 🚀 Como Executar o Script

### 1️⃣ Verifique se o Python está instalado
Execute o seguinte comando para verificar a versão instalada:
```bash
python --version
```

### 2️⃣ Crie um ambiente virtual
Para isolar as dependências do projeto, crie um ambiente virtual:
```bash
python -m venv venv
```

### 3️⃣ Ative o ambiente virtual
- **Windows (CMD/Powershell)**:
  ```bash
  venv\Scripts\activate
  ```
- **Linux/macOS (Terminal)**:
  ```bash
  source venv/bin/activate
  ```

### 4️⃣ Instale as dependências
Caso o projeto tenha dependências, instale-as utilizando:
```bash
pip install -r requirements.txt
```

### 5️⃣ Execute o script
```bash
python nome_do_script.py
```

## 🔧 Como Criar um Executável

Para transformar o script em um executável `.exe`, siga os passos abaixo:

### 1️⃣ Instalar o PyInstaller
Se ainda não tem o PyInstaller instalado, execute:
```bash
pip install pyinstaller
```

### 2️⃣ Gerar o Executável
Navegue até a pasta onde o script está localizado e execute um dos seguintes comandos:

#### Criar um único arquivo executável:
```bash
pyinstaller --onefile --noconsole nome_do_script.py
```
- `--onefile`: Gera um único arquivo `.exe`
- `--noconsole`: Remove a janela do terminal ao executar o programa

Este comando cria uma pasta `dist/` contendo o executável.

### 3️⃣ Acessar o Executável
Após a finalização do processo, vá até a pasta `dist/` e encontre o arquivo `.exe` gerado.

### 4️⃣ (Opcional) Adicionar Ícone Personalizado
Se quiser definir um ícone para o executável, use o parâmetro `--icon`:
```bash
pyinstaller --onefile --noconsole --icon=icone.ico nome_do_script.py
```
(Nota: O arquivo `.ico` deve estar no mesmo diretório que o script.)

## 🛠 Possíveis Problemas e Soluções
- **Erro de falta de módulo**: Se o executável não rodar, verifique se todas as bibliotecas necessárias estão instaladas.
- **Antivírus bloqueando o `.exe`**: O PyInstaller pode gerar executáveis que alguns antivírus identificam como suspeitos. Tente adicionar o `.exe` à lista de exceções do antivírus.

---
📌 **Desenvolvido por Davi Marques**