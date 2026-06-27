# transporExcel - AI Context Handoff

Este arquivo serve como contexto técnico para guiar modelos de IA que venham a trabalhar neste projeto no futuro. Ele descreve a arquitetura, regras de negócio e decisões técnicas de implementação.

---

## 🎯 Visão Geral do Projeto
O **transporExcel** é um utilitário desktop em Python com interface gráfica (GUI) feito para transpor dados de compras (faturas) de planilhas de **Controle de Cartão Corporativo** (geralmente estruturadas por colunas de categorias repetidas) para planilhas de **Relatório de Prestação de Contas** (estruturadas em uma tabela linear).

---

## 🛠️ Stack Tecnológico
1. **Linguagem:** Python 3.12+
2. **Interface Gráfica:** Tkinter / `ttk` (nativa do Python, estilizada com tema moderno e fontes Segoe UI).
3. **Automação do Excel:** `win32com.client` (automação COM nativa do Windows).
4. **Distribuição:** PyInstaller (para compilar em `.exe` autônomo).

---

## 🏗️ Arquitetura e Decisões Técnicas

### 1. Tratamento de DRM/Criptografia Corporativa
As planilhas de controle utilizam segurança corporativa do Office (DRM/IRM). Bibliotecas Python puras (como pandas ou openpyxl) falham ao abrir esses arquivos (`zipfile.BadZipFile` ou `XLRDError`).
* **Solução:** O script utiliza o Microsoft Excel instalado localmente via automação COM. O Excel abre o arquivo em segundo plano (invisível/headless) e realiza a descriptografia transparente usando as credenciais do usuário local.

### 2. Execução Multi-threaded
Operações de leitura e escrita via COM no Excel podem demorar de 1 a 5 segundos. Para evitar que a interface gráfica trave ou mostre "Não Respondendo":
* **Solução:** Todas as rotinas pesadas (carregar abas, mapear colunas, transpor, limpar) rodam em threads separadas (`threading.Thread`).
* **Regra Importante do COM em Threads:** Sempre inicialize e encerre o ecossistema COM na thread de execução:
  ```python
  import pythoncom
  pythoncom.CoInitialize()
  # Executa ações com win32com...
  pythoncom.CoUninitialize()
  ```
* **Atualizações de UI seguras:** As updates na interface gráfica vindas das threads são enfileiradas de forma segura na thread principal usando `root.after()`.

### 3. Mapeamento Semântico de Colunas
O programa possui um motor que analisa os cabeçalhos das planilhas para mapear equivalências:
* **Sinônimos Semânticos:** Busca palavras-chave em português e inglês para associar as colunas (ex: "Data"/"Dia", "Descrição"/"Item", "Nota"/"NF", "Valor"/"Preço").
* **Layouts Suportados:**
  * **Categorias Múltiplas (Agrupado):** Se houver mais de uma coluna de descrição na planilha fonte (ex: Colunas B, E, H...), o motor interpreta que a planilha é agrupada em colunas de categorias repetidas compostas pelo trio `(Preço, Descrição, NF)`. Ele extrai os registros de todas elas e os unifica.
  * **Tabela Linear (Flat):** Se houver apenas uma coluna de descrição, ele faz o mapeamento um-para-um clássico de colunas.

### 4. Integração com Tabelas Nativas do Excel (`ListObjects`)
A planilha de destino possui uma tabela nativa formatada ("Formatar como Tabela" do Excel, com filtros e cores zebradas). Automação COM direta de células não aciona a expansão automática da tabela no Excel.
* **Solução:** O script pesquisa por `ListObjects` na planilha destino e altera dinamicamente o intervalo da tabela usando `table.Resize()` para que o novo intervalo englobe todas as novas linhas transpostas ou seja reduzido a 1 linha vazia na limpeza.

### 5. Replicação de Formatações (`PasteSpecial`)
Para garantir que as novas linhas fiquem perfeitamente idênticas às anteriores:
* O script copia a primeira linha de dados (`dst_header_row + 1`) como modelo de estilo e aplica `PasteSpecial(Paste=-4122)` (apenas formatos) em todas as novas linhas gravadas, copiando bordas, fontes e formatos numéricos sem alterar os valores.

### 6. Limpeza Total de Linhas
Ao limpar ou sobrescrever, o programa apaga linhas inteiras utilizando o intervalo dinâmico `.Rows(f"{start_row}:{last_used_row}").ClearContents()`, garantindo compatibilidade com planilhas com qualquer número de colunas.

---

## 📂 Estrutura de Arquivos
* `transporExcel.py`: Arquivo principal da aplicação Tkinter.
* `transporExcel.bat`: Inicializador local que ativa a `.venv` e roda o script em UTF-8.
* `requirements.txt`: Contém a lista de dependências (`pywin32` e `pyinstaller`).
* `compilado/transporExcel.exe`: Arquivo binário empacotado para execução direta sem necessidade de instalar dependências.
* `main.py`: Arquivo da versão CLI (linha de comando).
* `executar.bat`: Inicializador local da versão CLI.
* `gemini.md`: Este arquivo de contexto de IA.

---

## 💻 Instruções de Desenvolvimento e Build

### Instalar dependências
```bash
python -m venv .venv
.venv\Scripts\activate.bat
pip install -r requirements.txt
```

### Compilar o executável
Para gerar um novo executável após fazer alterações no código:
```bash
.venv\Scripts\pyinstaller --noconsole --onefile transporExcel.py
```
*(O binário final será colocado na pasta `dist/` e deve ser movido para a pasta `compilado/`)*.
