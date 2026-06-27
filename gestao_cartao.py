import os
import sys
import glob
import win32com.client

def get_spreadsheets(folder_path):
    """Retorna uma lista de caminhos absolutos das planilhas na pasta."""
    if not os.path.exists(folder_path):
        return []
    # Busca por extensões comuns de planilhas
    extensions = ["*.xlsx", "*.xls", "*.xlsm", "*.csv"]
    files = []
    for ext in extensions:
        files.extend(glob.glob(os.path.join(folder_path, ext)))
    # Garante caminhos únicos e absolutos
    return list(set(os.path.abspath(f) for f in files))

def check_folders_and_files():
    """Verifica as pastas e arquivos de planilha de acordo com os requisitos."""
    cwd = os.getcwd()
    subdirs = [d for d in os.listdir(cwd) if os.path.isdir(os.path.join(cwd, d))]
    
    source_dir = None
    dest_dir = None
    
    # Identifica as pastas ignorando diferenças pequenas na escrita (de/do, acentos)
    for d in subdirs:
        d_lower = d.lower()
        if "controle do cart" in d_lower or "controle de cart" in d_lower:
            source_dir = os.path.join(cwd, d)
        elif "planilha relatorio" in d_lower or "planilha relat" in d_lower:
            dest_dir = os.path.join(cwd, d)
            
    if not source_dir:
        print("Erro: Pasta de controle ('controle do cartão') não encontrada no diretório atual.")
        return None, None
    if not dest_dir:
        print("Erro: Pasta de relatório ('planilha relatorio') não encontrada no diretório atual.")
        return None, None

    source_files = get_spreadsheets(source_dir)
    dest_files = get_spreadsheets(dest_dir)
    
    # Requisito: O script não pode rodar se houver mais de uma planilha nestas pastas.
    if len(source_files) > 1:
        print(f"Erro: Mais de uma planilha encontrada na pasta '{os.path.basename(source_dir)}':")
        for f in source_files:
            print(f" - {os.path.basename(f)}")
        print("O script não pode ser executado.")
        return None, None
        
    if len(dest_files) > 1:
        print(f"Erro: Mais de uma planilha encontrada na pasta '{os.path.basename(dest_dir)}':")
        for f in dest_files:
            print(f" - {os.path.basename(f)}")
        print("O script não pode ser executado.")
        return None, None
        
    if len(source_files) == 0:
        print(f"Erro: Nenhuma planilha encontrada na pasta '{os.path.basename(source_dir)}'.")
        return None, None
        
    if len(dest_files) == 0:
        print(f"Erro: Nenhuma planilha encontrada na pasta '{os.path.basename(dest_dir)}'.")
        return None, None
        
    return source_files[0], dest_files[0]

def extract_transactions_from_sheet(sheet):
    """Lê a aba da planilha de controle e extrai os dados das transações."""
    # 1. Encontrar as colunas que contém a descrição da compra no cabeçalho (linha 3)
    desc_cols = []
    last_col = sheet.UsedRange.Columns.Count
    for c in range(1, last_col + 1):
        cell_val = sheet.Cells(3, c).Value
        if cell_val and isinstance(cell_val, str) and "descri" in cell_val.lower():
            desc_cols.append(c)
            
    # 2. Construir mapeamento de colunas para cada categoria
    category_mappings = []
    for c_desc in desc_cols:
        c_price = c_desc - 1
        
        # Encontra a coluna de Nota Fiscal (NF) associada
        # Tenta verificar a célula imediatamente à direita (+1) ou na próxima (+2)
        c_nf = c_desc + 1
        val_r3_next = sheet.Cells(3, c_desc + 1).Value
        if val_r3_next and isinstance(val_r3_next, str) and ("nf" in val_r3_next.lower() or "nota" in val_r3_next.lower()):
            c_nf = c_desc + 1
        else:
            val_r3_next2 = sheet.Cells(3, c_desc + 2).Value
            if val_r3_next2 and isinstance(val_r3_next2, str) and ("nf" in val_r3_next2.lower() or "nota" in val_r3_next2.lower()):
                c_nf = c_desc + 2
                
        category_mappings.append({
            "desc_col": c_desc,
            "price_col": c_price,
            "nf_col": c_nf,
        })
        
    # 3. Varre as linhas a partir da linha 5
    transactions = []
    last_row = sheet.UsedRange.Rows.Count
    for r in range(5, last_row + 1):
        for mapping in category_mappings:
            desc_val = sheet.Cells(r, mapping["desc_col"]).Value
            price_val = sheet.Cells(r, mapping["price_col"]).Value
            nf_val = sheet.Cells(r, mapping["nf_col"]).Value
            
            # Se descrição e valor estão preenchidos
            if desc_val and price_val is not None:
                desc_str = str(desc_val).strip()
                # Ignora células que contêm o texto de instrução/cabeçalho
                if desc_str == "" or "descri" in desc_str.lower():
                    continue
                    
                try:
                    price_float = float(price_val)
                except (ValueError, TypeError):
                    continue
                    
                # Trata o formato do número da nota fiscal
                nf_str = ""
                if nf_val is not None:
                    if isinstance(nf_val, float) and nf_val.is_integer():
                        nf_str = str(int(nf_val))
                    else:
                        nf_str = str(nf_val).strip()
                        if nf_str.endswith(".0"):
                            nf_str = nf_str[:-2]
                            
                transactions.append({
                    "desc": desc_str,
                    "price": price_float,
                    "nf": nf_str
                })
                
    return transactions

def clear_report_sheet(dest_sheet):
    """Limpa os dados da planilha de relatório, preservando o cabeçalho."""
    last_row = dest_sheet.UsedRange.Rows.Count
    if last_row >= 4:
        # Limpa o conteúdo do intervalo de A4 até D e a última linha usada
        dest_sheet.Range(f"A4:D{last_row}").ClearContents()

def execute_transpose(source_path, dest_path):
    """Executa a transposição dos dados selecionados da planilha controle para a relatorio."""
    excel = None
    try:
        print("\n[1/3] Abrindo planilha de controle...")
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        
        wb_source = excel.Workbooks.Open(source_path)
        
        # Obter abas
        sheets_count = wb_source.Sheets.Count
        sheets = [wb_source.Sheets(i).Name for i in range(1, sheets_count + 1)]
        
        # Mostra as últimas 15 abas disponíveis
        print("\nAbas de Controle disponíveis (últimas 15):")
        start_idx = max(0, len(sheets) - 15)
        available_options = {}
        opt_num = 1
        for idx in range(start_idx, len(sheets)):
            sheet_name = sheets[idx]
            suffix = " (Mais recente)" if idx == len(sheets) - 1 else ""
            print(f"[{opt_num}] {sheet_name}{suffix}")
            available_options[str(opt_num)] = sheet_name
            opt_num += 1
            
        print("[O] Outra aba (digitar nome manualmente)")
        print("[V] Voltar ao menu principal")
        
        escolha = input("\nSelecione a aba desejada: ").strip()
        if escolha.upper() == 'V':
            wb_source.Close(SaveChanges=False)
            return
            
        aba_selecionada = None
        if escolha.upper() == 'O':
            aba_selecionada = input("Digite o nome exato da aba desejada: ").strip()
        elif escolha in available_options:
            aba_selecionada = available_options[escolha]
        else:
            print("Opção inválida.")
            wb_source.Close(SaveChanges=False)
            return
            
        # Tenta carregar a aba selecionada
        try:
            sheet_source = wb_source.Sheets(aba_selecionada)
        except Exception:
            print(f"Erro: Aba '{aba_selecionada}' não encontrada na planilha de controle.")
            wb_source.Close(SaveChanges=False)
            return
            
        # Requisito: Lembrar o usuário que os dados serão sobrescritos e pedir confirmação.
        print("\n" + "!" * 80)
        print("AVISO IMPORTANTE:")
        print("Toda vez que a opção de transpor for acionada, os dados na planilha de")
        print("relatório serão SOBRESCRITOS e perdidos permanentemente.")
        print("!" * 80)
        
        confirmar = input(f"Confirma a transposição da aba '{aba_selecionada}' para o relatório? (S/N): ").strip().upper()
        if confirmar not in ["S", "SIM"]:
            print("Operação cancelada pelo usuário.")
            wb_source.Close(SaveChanges=False)
            return
            
        print(f"\n[2/3] Extraindo transações da aba '{aba_selecionada}'...")
        transactions = extract_transactions_from_sheet(sheet_source)
        wb_source.Close(SaveChanges=False)
        
        if not transactions:
            print("Aviso: Nenhuma transação válida encontrada para transpor nesta aba.")
            return
            
        print(f"[3/3] Gravando {len(transactions)} transações na planilha de relatório...")
        wb_dest = excel.Workbooks.Open(dest_path)
        try:
            sheet_dest = wb_dest.Sheets(1)
            
            # Limpa dados anteriores
            clear_report_sheet(sheet_dest)
            
            # Escreve novas transações
            for idx, t in enumerate(transactions):
                r_target = 4 + idx
                sheet_dest.Cells(r_target, 1).Value = None  # Coluna Data vazia
                sheet_dest.Cells(r_target, 2).Value = t["desc"]
                sheet_dest.Cells(r_target, 3).Value = t["nf"]
                sheet_dest.Cells(r_target, 4).Value = t["price"]
                
            wb_dest.Close(SaveChanges=True)
            print("\n>>> Sucesso! Dados transpostos com sucesso para a planilha de relatório. <<<")
        except Exception as e:
            wb_dest.Close(SaveChanges=False)
            raise e
            
    except Exception as e:
        print(f"\nErro inesperado durante a transposição: {e}")
    finally:
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass

def execute_clear(dest_path):
    """Apaga os dados da planilha relatório."""
    confirmar = input("\nTem certeza que deseja APAGAR todos os dados da planilha de relatório? (S/N): ").strip().upper()
    if confirmar not in ["S", "SIM"]:
        print("Operação cancelada pelo usuário.")
        return
        
    excel = None
    try:
        print("\nAbrindo planilha de relatório...")
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        
        wb_dest = excel.Workbooks.Open(dest_path)
        try:
            sheet_dest = wb_dest.Sheets(1)
            clear_report_sheet(sheet_dest)
            wb_dest.Close(SaveChanges=True)
            print("\n>>> Sucesso! Todos os dados da planilha de relatório foram apagados. <<<")
        except Exception as e:
            wb_dest.Close(SaveChanges=False)
            raise e
            
    except Exception as e:
        print(f"\nErro inesperado ao apagar os dados: {e}")
    finally:
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass

def main():
    # 1. Valida arquivos e pastas
    source_file, dest_file = check_folders_and_files()
    if not source_file or not dest_file:
        print("\nNão foi possível iniciar o script devido aos erros acima.")
        input("\nPressione Enter para sair...")
        sys.exit(1)
        
    while True:
        print("\n" + "=" * 60)
        print("              SISTEMA DE GESTÃO DE CARTÕES")
        print("=" * 60)
        print(f"Origem (Controle): {os.path.basename(source_file)}")
        print(f"Destino (Relatório): {os.path.basename(dest_file)}")
        print("-" * 60)
        print("Escolha uma das opções abaixo:")
        print("1 - Transpor dados da planilha de Controle para Relatório")
        print("2 - Apagar dados da planilha de Relatório")
        print("0 - Sair")
        print("=" * 60)
        
        opcao = input("Opção: ").strip()
        
        if opcao == "1":
            execute_transpose(source_file, dest_file)
        elif opcao == "2":
            execute_clear(dest_file)
        elif opcao == "0":
            print("\nSaindo do sistema. Até logo!")
            break
        else:
            print("Opção inválida! Tente novamente.")

if __name__ == "__main__":
    main()
