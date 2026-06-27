# transporExcel

O **transporExcel** é um utilitário desktop em Python desenvolvido para realizar a transposição inteligente e automatizada de dados entre planilhas do Excel. O programa oferece duas formas de uso: uma **Interface Gráfica (GUI)** intuitiva e uma **Interface de Linha de Comando (CLI)** rápida.

---

## 🖥️ Como Executar o Programa

### 1. Interface Gráfica (GUI)
Você pode rodar a aplicação através do arquivo de inicialização:
👉 **`transporExcel.bat`**

Para utilizar em outras máquinas Windows que não possuam Python instalado:
1. Acesse a pasta **`compilado/`**.
2. Copie o arquivo executável **`transporExcel.exe`** para a máquina de destino.
3. Clique duas vezes no executável para iniciar.

### 2. Interface de Linha de Comando (CLI)
Você pode rodar a versão de terminal através do atalho:
👉 **`executar.bat`**

---

## 🛠️ Recursos e Funcionalidades

* **Seleção Dinâmica de Arquivos e Abas:** Permite selecionar qualquer arquivo de origem e destino e escolher quais abas serão utilizadas a partir de um menu de seleção.
* **Mapeamento Semântico de Colunas:** O motor do programa analisa os cabeçalhos das planilhas em tempo real e identifica colunas equivalentes de forma inteligente (exato ou por proximidade semântica de sinônimos).
* **Detecção de Layouts Múltiplos:** Suporta a leitura de planilhas com colunas agrupadas repetitivas ou tabelas lineares simples.
* **Ajuste de Tabelas Nativas (ListObjects):** Identifica tabelas estruturadas nativas do Excel e as redimensiona dinamicamente de acordo com o volume de dados inserido, garantindo que o layout, filtros e formatações de cor permaneçam uniformes.
* **Opções de Gravação:**
  * **Sobrescrever:** Limpa o conteúdo existente abaixo do cabeçalho da tabela de destino e grava os novos registros.
  * **Adicionar abaixo (Append):** Localiza a última linha ativa na tabela de destino e insere os registros na sequência, preservando a formatação.
* **Limpeza de Dados:** Permite esvaziar os registros da tabela de destino mantendo sua estrutura e cabeçalhos intactos.

---

## 🔒 Requisitos do Sistema (Tratamento de Arquivos Protegidos)

Para viabilizar a leitura e escrita em planilhas que possuem proteção corporativa ou de direitos digitais (DRM/IRM), o programa utiliza automação COM (`win32com.client`) integrada diretamente ao Microsoft Excel local. 

* **Requisito:** A máquina de execução precisa ter o **Microsoft Excel instalado** localmente para que a descriptografia ocorra de forma transparente pelo sistema operacional.

---

## 📂 Estrutura de Arquivos
* `transporExcel.py`: Código fonte principal da interface gráfica.
* `transporExcel.bat`: Inicializador local para a interface gráfica.
* `main.py`: Código fonte principal da interface CLI.
* `executar.bat`: Inicializador local para a interface CLI.
* `compilado/transporExcel.exe`: Executável autônomo.
* `requirements.txt`: Dependências necessárias do Python.
* `gemini.md`: Contexto técnico para desenvolvedores.
