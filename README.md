# transporExcel

O **transporExcel** é um utilitário desktop em Python desenvolvido para realizar a transposição inteligente e automatizada de dados entre planilhas do Excel através de uma **Interface Gráfica (GUI)** intuitiva.

---

## 🖥️ Como Executar o Programa

### 1. Executando o Código Fonte (Desenvolvimento)
Você pode executar o programa localmente usando o atalho:
👉 **`transporExcel.bat`**

### 2. Gerando o Executável Autônomo (Para testes em outras máquinas)
Os arquivos binários compilados são ignorados pelo Git (`.gitignore`). Para gerar o arquivo executável autônomo (`transporExcel.exe`) para testes em outras máquinas Windows:

1. Ative o ambiente virtual e instale as dependências:
   ```bash
   .venv\Scripts\activate.bat
   pip install -r requirements.txt
   ```
2. Execute o comando de compilação do PyInstaller:
   ```bash
   pyinstaller --noconsole --onefile transporExcel.py
   ```
3. Mova o executável gerado da pasta `dist/` para a pasta `compilado/`.
4. Copie o arquivo `transporExcel.exe` da pasta `compilado/` para qualquer outra máquina Windows e execute-o com um clique duplo (sem necessidade de ter Python instalado).

---

## 🛠️ Recursos e Funcionalidades

* **Seleção Livre de Arquivos e Abas:** Permite selecionar qualquer arquivo de origem e destino diretamente pelo computador e escolher as abas de trabalho correspondentes.
* **Mapeamento Semântico de Colunas:** O motor analisa os cabeçalhos das planilhas em tempo real e identifica colunas equivalentes de forma inteligente (exato ou por proximidade semântica de sinônimos).
* **Mapeamento Flexível:** Suporta tanto planilhas com colunas agrupadas por categorias repetitivas quanto tabelas lineares simples.
* **Ajuste de Tabelas Nativas (ListObjects):** Identifica tabelas estruturadas nativas do Excel e as redimensiona dinamicamente de acordo com o volume de dados inserido, garantindo que o layout, filtros e formatações de cor permaneçam uniformes.
* **Opções de Gravação:**
  * **Sobrescrever:** Limpa o conteúdo existente abaixo do cabeçalho da tabela de destino e grava os novos registros.
  * **Adicionar abaixo (Append):** Localiza a última linha ativa na tabela de destino e insere os registros na sequência, preservando a formatação.
* **Limpar Tabela:** Botão dedicado que apaga os lançamentos da tabela de destino mantendo sua estrutura e cabeçalhos intactos.

---

## 🔒 Requisitos do Sistema (Tratamento de Arquivos Protegidos)

Para viabilizar a leitura e escrita em planilhas que possuem proteção corporativa ou de direitos digitais (DRM/IRM), o programa utiliza automação COM (`win32com.client`) integrada diretamente ao Microsoft Excel local. 

* **Requisito:** A máquina de execução precisa ter o **Microsoft Excel instalado** localmente para que a descriptografia ocorra de forma transparente pelo sistema operacional.

---

## 📂 Estrutura de Arquivos
* `transporExcel.py`: Código fonte principal da interface gráfica.
* `transporExcel.bat`: Inicializador local para a interface gráfica.
* `requirements.txt`: Dependências necessárias do Python (inclui `pywin32` e `pyinstaller`).
* `gemini.md`: Contexto técnico para desenvolvedores.
