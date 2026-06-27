# transporExcel

Este projeto fornece soluções automatizadas para realizar a transposição de dados de compras de uma planilha de **Controle de Cartão** para uma planilha de **Relatório de Prestação de Contas**, além de possibilitar a limpeza dos dados da planilha de relatório.

---

## 🖥️ Como Executar o Programa (Interface Gráfica)

### 1. Na sua máquina local (Desenvolvimento)
Você pode executar o programa usando o atalho de lote:
👉 **[transporExcel.bat](file:///C:/Users/mattheus.pereira/Desktop/script_gestao_de_cartao/transporExcel.bat)**

### 2. Em outras máquinas Windows (Sem Python instalado)
Para testar e utilizar o programa em outras máquinas de forma autônoma:
1. Abra a pasta **[compilado](file:///C:/Users/mattheus.pereira/Desktop/script_gestao_de_cartao/compilado)**.
2. Copie o arquivo executável **`transporExcel.exe`** para a outra máquina.
3. Clique duas vezes no executável para iniciar. Não é necessário ter o Python ou qualquer dependência instalada na outra máquina!

---

## 🛠️ Recursos da Interface Gráfica

* **Seleção Livre de Arquivos:** Escolha qualquer planilha fonte e destino no seu computador usando as caixas com botões "Procurar...".
* **Seleção de Abas:** O programa lê e exibe as abas disponíveis em um menu suspenso para você escolher.
* **Mapeamento Automático Inteligente (Preview):** O programa analisa semanticamente os cabeçalhos das planilhas em segundo plano e exibe uma caixa detalhada informando quais colunas equivalentes foram encontradas.
* **Tabela Única Uniforme (ListObject.Resize):** O programa detecta Tabelas nativas no Excel (ListObjects) e as redimensiona para o tamanho exato dos novos dados, garantindo que o layout visual (zebra, bordas e filtros) permaneça uniforme após cada inserção ou limpeza.
* **Modos de Transposição:**
  * **Sobrescrever:** Limpa a tabela existente (mantendo o cabeçalho) e insere os novos dados.
  * **Adicionar abaixo:** Localiza a última linha preenchida na tabela e adiciona as novas compras logo abaixo, estendendo a tabela nativa.
* **Limpar Tabela:** Botão dedicado que apaga os lançamentos da tabela, deixando-a vazia e pronta para uso.

---

## 🔒 Tecnologia de Leitura (Tratamento de DRM)

As planilhas de cartão corporativo utilizam criptografia de direitos digitais (DRM/IRM). Para possibilitar a decodificação de forma segura e transparente, o programa utiliza o Microsoft Excel instalado na máquina em segundo plano (COM Automation). A descriptografia ocorre automaticamente usando suas credenciais locais do Windows e da conta Office. **Portanto, a máquina de destino onde o executável for rodar precisa ter o Microsoft Excel instalado.**

---

## 📂 Arquivos do Projeto
* [compilado/transporExcel.exe](file:///C:/Users/mattheus.pereira/Desktop/script_gestao_de_cartao/compilado/transporExcel.exe): Executável autônomo compilado para distribuição.
* [transporExcel.py](file:///C:/Users/mattheus.pereira/Desktop/script_gestao_de_cartao/transporExcel.py): Código fonte da Interface Gráfica.
* [transporExcel.bat](file:///C:/Users/mattheus.pereira/Desktop/script_gestao_de_cartao/transporExcel.bat): Atalho de inicialização para Windows local.
* [main.py](file:///C:/Users/mattheus.pereira/Desktop/script_gestao_de_cartao/main.py): Código fonte da versão terminal CLI.
* [executar.bat](file:///C:/Users/mattheus.pereira/Desktop/script_gestao_de_cartao/executar.bat): Atalho de inicialização para a versão terminal.
* [requirements.txt](file:///C:/Users/mattheus.pereira/Desktop/script_gestao_de_cartao/requirements.txt): Dependências do projeto.
