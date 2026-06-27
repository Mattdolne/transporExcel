# transporExcel

O **transporExcel** é um utilitário desktop em Python desenvolvido para realizar a transposição inteligente e automatizada de dados entre planilhas do Excel através de uma **Interface Gráfica (GUI)** intuitiva.

---

## 🖥️ Como Executar o Programa

### 1. Na sua máquina local (Desenvolvimento)
Você pode executar o programa usando o atalho de lote:
👉 **`transporExcel.bat`**

### 2. Em outras máquinas Windows (Sem Python instalado)
Para utilizar o programa em outras máquinas de forma autônoma:
1. Acesse a pasta **`compilado/`**.
2. Copie o arquivo executável **`transporExcel.exe`** para a máquina de destino.
3. Clique duas vezes no executável para iniciar.

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
* `compilado/transporExcel.exe`: Executável autônomo.
* `requirements.txt`: Dependências necessárias do Python.
* `gemini.md`: Contexto técnico para desenvolvedores.
