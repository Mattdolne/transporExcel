import os
import sys
import glob
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import win32com.client
import pythoncom

# Constante do Excel para encontrar a última linha (VBA: xlUp)
xlUp = -4162

class CardGestorGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Transpositor de Planilhas - transporExcel")
        self.root.geometry("850x700")
        self.root.minsize(800, 600)
        
        # Configuração de Estilo
        self.style = ttk.Style()
        self.style.theme_use('vista' if 'vista' in self.style.theme_names() else 'clam')
        
        # Cores e Fontes
        self.bg_color = "#f4f6f9"
        self.header_bg = "#1e3d59"
        self.btn_green = "#17b978"
        self.btn_red = "#ff5722"
        self.text_color = "#333333"
        self.font_family = "Segoe UI"
        
        self.root.configure(bg=self.bg_color)
        
        # Variáveis
        self.source_file_var = tk.StringVar()
        self.dest_file_var = tk.StringVar()
        self.source_sheet_var = tk.StringVar()
        self.dest_sheet_var = tk.StringVar()
        self.trans_mode_var = tk.StringVar(value="overwrite") # options: "overwrite", "append"
        self.current_mappings = None
        
        # Criação dos Componentes
        self.create_widgets()
        
    def create_widgets(self):
        # 1. Banner de Cabeçalho
        header_frame = tk.Frame(self.root, bg=self.header_bg, height=70)
        header_frame.pack(fill="x", side="top")
        header_frame.pack_propagate(False)
        
        header_title = tk.Label(
            header_frame, 
            text="transporExcel - Transpositor de Planilhas", 
            font=(self.font_family, 16, "bold"), 
            fg="white", 
            bg=self.header_bg
        )
        header_title.pack(pady=18, padx=20, side="left")
        
        # Container Principal com padding
        main_container = tk.Frame(self.root, bg=self.bg_color, padx=20, pady=15)
        main_container.pack(fill="both", expand=True)
        
        # 2. Frame de Seleção de Arquivos
        files_frame = tk.LabelFrame(
            main_container, 
            text=" Seleção de Planilhas ", 
            font=(self.font_family, 10, "bold"), 
            bg=self.bg_color, 
            fg=self.header_bg,
            padx=10, 
            pady=10
        )
        files_frame.pack(fill="x", pady=(0, 10))
        
        # Planilha Fonte
        tk.Label(files_frame, text="Planilha Fonte:", font=(self.font_family, 9, "bold"), bg=self.bg_color).grid(row=0, column=0, sticky="w", pady=5)
        source_entry = tk.Entry(files_frame, textvariable=self.source_file_var, font=(self.font_family, 9), width=65)
        source_entry.grid(row=0, column=1, padx=5, pady=5, sticky="we")
        btn_browse_src = tk.Button(
            files_frame, text="Procurar...", font=(self.font_family, 9), 
            command=lambda: self.browse_file(self.source_file_var, self.cb_source_sheet, self.lbl_status_src)
        )
        btn_browse_src.grid(row=0, column=2, padx=5, pady=5)
        
        tk.Label(files_frame, text="Aba Fonte:", font=(self.font_family, 9), bg=self.bg_color).grid(row=1, column=0, sticky="w", pady=5)
        self.cb_source_sheet = ttk.Combobox(files_frame, textvariable=self.source_sheet_var, font=(self.font_family, 9), state="readonly")
        self.cb_source_sheet.grid(row=1, column=1, padx=5, pady=5, sticky="w")
        self.cb_source_sheet.bind("<<ComboboxSelected>>", self.on_sheet_selected)
        self.lbl_status_src = tk.Label(files_frame, text="", font=(self.font_family, 8), bg=self.bg_color, fg="gray")
        self.lbl_status_src.grid(row=1, column=1, columnspan=2, padx=(180, 5), pady=5, sticky="w")
        
        # Divisor simples
        ttk.Separator(files_frame, orient="horizontal").grid(row=2, column=0, columnspan=3, pady=10, sticky="we")
        
        # Planilha Destino
        tk.Label(files_frame, text="Planilha Destino:", font=(self.font_family, 9, "bold"), bg=self.bg_color).grid(row=3, column=0, sticky="w", pady=5)
        dest_entry = tk.Entry(files_frame, textvariable=self.dest_file_var, font=(self.font_family, 9), width=65)
        dest_entry.grid(row=3, column=1, padx=5, pady=5, sticky="we")
        btn_browse_dst = tk.Button(
            files_frame, text="Procurar...", font=(self.font_family, 9), 
            command=lambda: self.browse_file(self.dest_file_var, self.cb_dest_sheet, self.lbl_status_dst)
        )
        btn_browse_dst.grid(row=3, column=2, padx=5, pady=5)
        
        tk.Label(files_frame, text="Aba Destino:", font=(self.font_family, 9), bg=self.bg_color).grid(row=4, column=0, sticky="w", pady=5)
        self.cb_dest_sheet = ttk.Combobox(files_frame, textvariable=self.dest_sheet_var, font=(self.font_family, 9), state="readonly")
        self.cb_dest_sheet.grid(row=4, column=1, padx=5, pady=5, sticky="w")
        self.cb_dest_sheet.bind("<<ComboboxSelected>>", self.on_sheet_selected)
        self.lbl_status_dst = tk.Label(files_frame, text="", font=(self.font_family, 8), bg=self.bg_color, fg="gray")
        self.lbl_status_dst.grid(row=4, column=1, columnspan=2, padx=(180, 5), pady=5, sticky="w")
        
        files_frame.columnconfigure(1, weight=1)
        
        # 3. Opções de Transposição
        options_frame = tk.LabelFrame(
            main_container, 
            text=" Opções de Transposição ", 
            font=(self.font_family, 10, "bold"), 
            bg=self.bg_color, 
            fg=self.header_bg,
            padx=10, 
            pady=5
        )
        options_frame.pack(fill="x", pady=5)
        
        tk.Radiobutton(
            options_frame, 
            text="Sobrescrever todos os dados da planilha destino (Limpa a tabela abaixo do cabeçalho e insere novos dados)", 
            variable=self.trans_mode_var, 
            value="overwrite", 
            font=(self.font_family, 9),
            bg=self.bg_color,
            activebackground=self.bg_color
        ).pack(anchor="w", pady=2)
        
        tk.Radiobutton(
            options_frame, 
            text="Adicionar abaixo dos dados existentes (Mantém o histórico, insere na próxima linha vazia e preserva a formatação)", 
            variable=self.trans_mode_var, 
            value="append", 
            font=(self.font_family, 9),
            bg=self.bg_color,
            activebackground=self.bg_color
        ).pack(anchor="w", pady=2)
        
        # 5. Botões de Ação
        actions_frame = tk.Frame(main_container, bg=self.bg_color, pady=10)
        actions_frame.pack(fill="x")
        
        self.btn_transpose = tk.Button(
            actions_frame, 
            text="Confirmar e Transpor Dados", 
            font=(self.font_family, 10, "bold"), 
            fg="white", 
            bg=self.btn_green,
            activeforeground="white",
            activebackground="#149f67",
            padx=15, 
            pady=8,
            relief="groove",
            command=self.run_transposition
        )
        self.btn_transpose.pack(side="left", padx=(0, 10))
        
        self.btn_clear = tk.Button(
            actions_frame, 
            text="Limpar Dados da Planilha Destino", 
            font=(self.font_family, 10, "bold"), 
            fg="white", 
            bg=self.btn_red,
            activeforeground="white",
            activebackground="#e04c1b",
            padx=15, 
            pady=8,
            relief="groove",
            command=self.run_clear_destination
        )
        self.btn_clear.pack(side="left")
        
        self.lbl_action_progress = tk.Label(
            actions_frame, 
            text="", 
            font=(self.font_family, 9, "bold"), 
            bg=self.bg_color
        )
        self.lbl_action_progress.pack(side="right", padx=10)
        
        # 4. Preview do Mapeamento de Colunas Equivalentes
        preview_frame = tk.LabelFrame(
            main_container, 
            text=" Colunas Equivalentes Identificadas (Verificação) ", 
            font=(self.font_family, 10, "bold"), 
            bg=self.bg_color, 
            fg=self.header_bg,
            padx=10, 
            pady=10
        )
        preview_frame.pack(fill="both", expand=True, pady=5)
        
        self.lbl_mapping_status = tk.Label(
            preview_frame, 
            text="Selecione as planilhas e as abas correspondentes para analisar as colunas.", 
            font=(self.font_family, 9, "italic"), 
            bg=self.bg_color, 
            fg="gray",
            anchor="w"
        )
        self.lbl_mapping_status.pack(fill="x", pady=(0, 5))
        
        # Caixa de Texto com Scrollbar para exibir o mapeamento
        text_container = tk.Frame(preview_frame, bg="white", bd=1, relief="sunken")
        text_container.pack(fill="both", expand=True)
        
        scrollbar = tk.Scrollbar(text_container)
        scrollbar.pack(side="right", fill="y")
        
        self.txt_preview = tk.Text(
            text_container, 
            font=("Consolas", 10), 
            yscrollcommand=scrollbar.set, 
            wrap="word", 
            bg="#fdfdfd", 
            fg=self.text_color,
            state="disabled"
        )
        self.txt_preview.pack(fill="both", expand=True, side="left")
        scrollbar.config(command=self.txt_preview.yview)
        
        # 6. Rodapé de Logs/Console
        footer_frame = tk.Frame(self.root, bg="#eaecef", height=25)
        footer_frame.pack(fill="x", side="bottom")
        self.lbl_console = tk.Label(
            footer_frame, 
            text="Pronto.", 
            font=(self.font_family, 8), 
            bg="#eaecef", 
            fg="#555555",
            anchor="w",
            padx=10
        )
        self.lbl_console.pack(fill="x", pady=2)
        
    def log(self, message):
        """Atualiza a barra de logs no rodapé."""
        self.lbl_console.config(text=message)
        
    def browse_file(self, var, combobox, label_status):
        """Abre janela para selecionar o arquivo e inicia a carga de abas."""
        filepath = filedialog.askopenfilename(
            title="Selecionar Planilha do Excel",
            filetypes=[("Arquivos do Excel", "*.xlsx;*.xls;*.xlsm"), ("Todos os Arquivos", "*.*")]
        )
        if filepath:
            filepath = os.path.abspath(filepath)
            var.set(filepath)
            combobox['values'] = []
            combobox.set('')
            label_status.config(text="Carregando abas...", fg="blue")
            self.log(f"Carregando abas do arquivo: {os.path.basename(filepath)}")
            
            # Executa a carga das abas em segundo plano
            threading.Thread(target=self.load_sheets_thread, args=(filepath, combobox, label_status), daemon=True).start()

    def load_sheets_thread(self, filepath, combobox, label_status):
        """Thread para ler as abas de uma planilha sem travar a interface."""
        pythoncom.CoInitialize()
        excel = None
        try:
            excel = win32com.client.Dispatch("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            
            wb = excel.Workbooks.Open(filepath, ReadOnly=True)
            sheets = [wb.Sheets(i).Name for i in range(1, wb.Sheets.Count + 1)]
            wb.Close(SaveChanges=False)
            
            # Atualiza UI no main thread
            self.root.after(0, lambda: self.update_sheets_cb(combobox, sheets, label_status))
        except Exception as e:
            err_msg = str(e)
            self.root.after(0, lambda: label_status.config(text="Erro ao ler abas.", fg="red"))
            self.root.after(0, lambda: messagebox.showerror("Erro de Leitura", f"Não foi possível abrir o arquivo:\n{err_msg}"))
        finally:
            if excel is not None:
                try:
                    excel.Quit()
                except:
                    pass
            pythoncom.CoUninitialize()

    def update_sheets_cb(self, combobox, sheets, label_status):
        """Atualiza o Combobox de abas na UI."""
        combobox['values'] = sheets
        if sheets:
            combobox.current(0)
            label_status.config(text=f"{len(sheets)} abas carregadas.", fg="green")
            self.log(f"Abas carregadas com sucesso.")
            self.on_sheet_selected(None)
        else:
            label_status.config(text="Nenhuma aba encontrada.", fg="orange")
            
    def on_sheet_selected(self, event):
        """Disparado quando uma aba é selecionada em qualquer combobox."""
        src_file = self.source_file_var.get()
        src_sheet = self.source_sheet_var.get()
        dst_file = self.dest_file_var.get()
        dst_sheet = self.dest_sheet_var.get()
        
        if src_file and src_sheet and dst_file and dst_sheet:
            self.lbl_mapping_status.config(text="Analisando equivalência de colunas...", fg="blue")
            self.txt_preview.config(state="normal")
            self.txt_preview.delete("1.0", "end")
            self.txt_preview.insert("1.0", "Carregando colunas...")
            self.txt_preview.config(state="disabled")
            
            # Roda detecção de colunas em background thread
            threading.Thread(
                target=self.detect_columns_thread, 
                args=(src_file, src_sheet, dst_file, dst_sheet), 
                daemon=True
            ).start()

    def detect_columns_thread(self, src_file, src_sheet, dst_file, dst_sheet):
        """Thread para ler e analisar cabeçalhos das duas planilhas."""
        pythoncom.CoInitialize()
        excel = None
        try:
            excel = win32com.client.Dispatch("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            
            wb_src = excel.Workbooks.Open(src_file, ReadOnly=True)
            wb_dst = excel.Workbooks.Open(dst_file, ReadOnly=True)
            
            sheet_src = wb_src.Sheets(src_sheet)
            sheet_dst = wb_dst.Sheets(dst_sheet)
            
            # Detecta linhas de cabeçalho
            src_header_row = self.find_header_row(sheet_src)
            dst_header_row = self.find_header_row(sheet_dst)
            
            # Analisa colunas
            mappings, src_headers, dst_headers = self.analyze_columns(
                sheet_src, src_header_row, 
                sheet_dst, dst_header_row
            )
            
            wb_src.Close(SaveChanges=False)
            wb_dst.Close(SaveChanges=False)
            
            # Prepara string de visualização
            preview_str = self.format_preview_string(mappings, src_headers, dst_headers, src_header_row, dst_header_row)
            
            # Atualiza UI no main thread
            self.root.after(0, lambda: self.update_mapping_preview(mappings, preview_str))
            
        except Exception as e:
            err_msg = str(e)
            self.root.after(0, lambda: self.lbl_mapping_status.config(text="Erro ao analisar colunas.", fg="red"))
            self.root.after(0, lambda: self.set_preview_text(f"Erro na análise de colunas:\n{err_msg}"))
        finally:
            if excel is not None:
                try:
                    excel.Quit()
                except:
                    pass
            pythoncom.CoUninitialize()

    def find_header_row(self, sheet):
        """Detecta automaticamente qual linha representa o cabeçalho."""
        desc_synonyms = ["descri", "item", "produto", "serviço", "servico", "detalhe"]
        # Verifica as 5 primeiras linhas
        for r in range(1, 6):
            for c in range(1, 20):  # primeiras 20 colunas
                val = sheet.Cells(r, c).Value
                if val and isinstance(val, str):
                    if any(syn in val.lower() for syn in desc_synonyms):
                        return r
        # Caso contrário, pega a linha com mais células de texto preenchidas
        best_row = 3
        max_text_cells = 0
        for r in range(1, 6):
            count = 0
            for c in range(1, 20):
                val = sheet.Cells(r, c).Value
                if val and isinstance(val, str) and len(val.strip()) > 1:
                    count += 1
            if count > max_text_cells:
                max_text_cells = count
                best_row = r
        return best_row

    def analyze_columns(self, src_sheet, src_header_row, dst_sheet, dst_header_row):
        """Identifica equivalência semântica entre as colunas fonte e destino."""
        src_cols_count = src_sheet.UsedRange.Columns.Count
        dst_cols_count = dst_sheet.UsedRange.Columns.Count
        
        src_headers = {}
        for c in range(1, src_cols_count + 1):
            val = src_sheet.Cells(src_header_row, c).Value
            if val is not None:
                src_headers[c] = str(val).strip()
                
        dst_headers = {}
        for c in range(1, dst_cols_count + 1):
            val = dst_sheet.Cells(dst_header_row, c).Value
            if val is not None:
                dst_headers[c] = str(val).strip()
                
        # Sinônimos de colunas comuns
        date_syns = ["data", "date", "dia", "vencimento", "emissão"]
        desc_syns = ["descri", "item", "produto", "serviço", "servico", "detalhe", "compra"]
        nf_syns = ["nf", "nota", "nº nota", "nº nf", "n nf", "comprovante", "cupom"]
        val_syns = ["valor", "preço", "preco", "total", "quantia", "pago"]
        
        # Encontra as colunas alvo na planilha destino
        dst_date_col = None
        dst_desc_col = None
        dst_nf_col = None
        dst_val_col = None
        
        for c, val in dst_headers.items():
            val_lower = val.lower()
            if any(s in val_lower for s in date_syns) and dst_date_col is None:
                dst_date_col = c
            elif any(s in val_lower for s in desc_syns) and dst_desc_col is None:
                dst_desc_col = c
            elif any(s in val_lower for s in nf_syns) and dst_nf_col is None:
                dst_nf_col = c
            elif any(s in val_lower for s in val_syns) and dst_val_col is None:
                dst_val_col = c
                
        # Procura colunas de descrição na planilha fonte
        src_desc_cols = []
        for c, val in src_headers.items():
            val_lower = val.lower()
            if any(s in val_lower for s in desc_syns):
                src_desc_cols.append(c)
                
        mappings = []
        is_grouped = len(src_desc_cols) > 1
        
        if is_grouped:
            # Caso A: Layout controle (múltiplas categorias em grupos)
            for c_desc in src_desc_cols:
                c_price = c_desc - 1
                c_nf = c_desc + 1
                
                # valida se a coluna de NF está correta verificando sua legenda no cabeçalho
                if c_desc + 1 in src_headers:
                    val_next = src_headers[c_desc + 1].lower()
                    if any(s in val_next for s in nf_syns):
                        c_nf = c_desc + 1
                    elif c_desc + 2 in src_headers:
                        val_next2 = src_headers[c_desc + 2].lower()
                        if any(s in val_next2 for s in nf_syns):
                            c_nf = c_desc + 2
                
                # Nome da categoria
                cat_name = ""
                if src_header_row > 1:
                    cat_val = src_sheet.Cells(src_header_row - 1, c_price).Value
                    if cat_val:
                        cat_name = str(cat_val).strip()
                if not cat_name:
                    cat_name = src_headers.get(c_price, f"Coluna {c_price}")
                    
                mappings.append({
                    "type": "grouped",
                    "category": cat_name,
                    "src_desc_col": c_desc,
                    "src_desc_name": src_headers.get(c_desc),
                    "src_price_col": c_price,
                    "src_price_name": src_headers.get(c_price),
                    "src_nf_col": c_nf,
                    "src_nf_name": src_headers.get(c_nf, ""),
                    "dst_date_col": dst_date_col,
                    "dst_desc_col": dst_desc_col,
                    "dst_nf_col": dst_nf_col,
                    "dst_val_col": dst_val_col
                })
        else:
            # Caso B: Tabela simples linear (uma coluna de descrição apenas)
            c_desc = src_desc_cols[0] if src_desc_cols else None
            c_price = None
            c_nf = None
            c_date = None
            
            for c, val in src_headers.items():
                val_lower = val.lower()
                if any(s in val_lower for s in val_syns) and c_price is None:
                    c_price = c
                elif any(s in val_lower for s in nf_syns) and c_nf is None and c != c_desc:
                    c_nf = c
                elif any(s in val_lower for s in date_syns) and c_date is None:
                    c_date = c
                    
            if c_desc:
                if c_price is None and c_desc - 1 in src_headers:
                    c_price = c_desc - 1
                if c_nf is None and c_desc + 1 in src_headers:
                    c_nf = c_desc + 1
                    
            mappings.append({
                "type": "flat",
                "src_date_col": c_date,
                "src_date_name": src_headers.get(c_date, "") if c_date else "",
                "src_desc_col": c_desc,
                "src_desc_name": src_headers.get(c_desc, "") if c_desc else "",
                "src_price_col": c_price,
                "src_price_name": src_headers.get(c_price, "") if c_price else "",
                "src_nf_col": c_nf,
                "src_nf_name": src_headers.get(c_nf, "") if c_nf else "",
                "dst_date_col": dst_date_col,
                "dst_desc_col": dst_desc_col,
                "dst_nf_col": dst_nf_col,
                "dst_val_col": dst_val_col
            })
            
        return mappings, src_headers, dst_headers

    def format_preview_string(self, mappings, src_headers, dst_headers, src_h_row, dst_h_row):
        """Retorna uma string formatada explicando de forma clara as equivalências encontradas."""
        if not mappings:
            return "Nenhum mapeamento de colunas pôde ser estabelecido automaticamente."
            
        res = []
        res.append("=========================================================================")
        res.append("             EQUIVALÊNCIA DE COLUNAS DETECTADA COM SUCESSO               ")
        res.append("=========================================================================\n")
        
        # Mapeamento do Destino
        first = mappings[0]
        res.append("ESTRUTURA DE DESTINO IDENTIFICADA:")
        res.append(f" • Linha do Cabeçalho -> Linha {dst_h_row}")
        res.append(f" • Data               -> Coluna {self.col_letter(first['dst_date_col'])}: '{dst_headers.get(first['dst_date_col'], 'NÃO DETECTADA')}'")
        res.append(f" • Descrição          -> Coluna {self.col_letter(first['dst_desc_col'])}: '{dst_headers.get(first['dst_desc_col'], 'NÃO DETECTADA')}'")
        res.append(f" • Nº Nota            -> Coluna {self.col_letter(first['dst_nf_col'])}: '{dst_headers.get(first['dst_nf_col'], 'NÃO DETECTADA')}'")
        res.append(f" • Valor              -> Coluna {self.col_letter(first['dst_val_col'])}: '{dst_headers.get(first['dst_val_col'], 'NÃO DETECTADA')}'\n")
        
        if first["type"] == "grouped":
            res.append("ESTRUTURA FONTE DETECTADA: [Layout de Categorias Múltiplas]")
            res.append("O programa irá extrair e unificar as compras das seguintes categorias:\n")
            for m in mappings:
                cat = m['category'].replace('\n', ' ')
                res.append(f" Categoria: '{cat}'")
                res.append(f"   └─ Valor: Coluna {self.col_letter(m['src_price_col'])} ('{m['src_price_name']}')")
                res.append(f"   └─ Descrição: Coluna {self.col_letter(m['src_desc_col'])} ('{m['src_desc_name']}')")
                res.append(f"   └─ Nº Nota: Coluna {self.col_letter(m['src_nf_col'])} ('{m['src_nf_name']}')")
                res.append("")
            res.append("Nota: As datas ficarão em branco pois a planilha fonte não possui coluna de data.")
        else:
            res.append("ESTRUTURA FONTE DETECTADA: [Layout de Tabela Simples/Linear]")
            m = first
            res.append(f" • Data:      Coluna {self.col_letter(m['src_date_col'])} ('{m['src_date_name']}') => {self.col_letter(m['dst_date_col'])}")
            res.append(f" • Descrição: Coluna {self.col_letter(m['src_desc_col'])} ('{m['src_desc_name']}') => {self.col_letter(m['dst_desc_col'])}")
            res.append(f" • Nº Nota:   Coluna {self.col_letter(m['src_nf_col'])} ('{m['src_nf_name']}') => {self.col_letter(m['dst_nf_col'])}")
            res.append(f" • Valor:     Coluna {self.col_letter(m['src_price_col'])} ('{m['src_price_name']}') => {self.col_letter(m['dst_val_col'])}")
            
        res.append("\nSe as equivalências acima estiverem corretas, clique em 'Confirmar e Transpor Dados'.")
        return "\n".join(res)

    def col_letter(self, col_index):
        """Converte índice numérico (1, 2, ...) para letra da coluna Excel (A, B, ...)."""
        if col_index is None:
            return "N/A"
        result = ""
        while col_index > 0:
            col_index, remainder = divmod(col_index - 1, 26)
            result = chr(65 + remainder) + result
        return result

    def set_preview_text(self, text):
        """Escreve o texto na caixa de preview na UI."""
        self.txt_preview.config(state="normal")
        self.txt_preview.delete("1.0", "end")
        self.txt_preview.insert("1.0", text)
        self.txt_preview.config(state="disabled")

    def update_mapping_preview(self, mappings, preview_str):
        """Callback do main thread para atualizar o mapeamento."""
        self.current_mappings = mappings
        self.set_preview_text(preview_str)
        
        # Valida se os campos mínimos foram encontrados
        first = mappings[0]
        missing = []
        if not first["dst_desc_col"]: missing.append("Descrição no Destino")
        if not first["dst_val_col"]: missing.append("Valor no Destino")
        
        if first["type"] == "flat":
            if not first["src_desc_col"]: missing.append("Descrição na Fonte")
            if not first["src_price_col"]: missing.append("Valor/Preço na Fonte")
        else:
            if not any(m["src_desc_col"] for m in mappings):
                missing.append("Nenhuma coluna de Descrição na Fonte")
                
        if missing:
            self.lbl_mapping_status.config(
                text=f"Aviso: Algumas colunas não puderam ser associadas automaticamente: {', '.join(missing)}", 
                fg="orange"
            )
        else:
            self.lbl_mapping_status.config(text="Equivalência mapeada com sucesso!", fg="green")

    def run_transposition(self):
        """Valida e inicia a thread de transposição de dados."""
        src_file = self.source_file_var.get()
        src_sheet = self.source_sheet_var.get()
        dst_file = self.dest_file_var.get()
        dst_sheet = self.dest_sheet_var.get()
        
        if not src_file or not src_sheet or not dst_file or not dst_sheet:
            messagebox.showwarning("Campos Incompletos", "Por favor, selecione os arquivos fonte e destino e suas respectivas abas.")
            return
            
        if not self.current_mappings:
            messagebox.showwarning("Mapeamento Inválido", "Aguarde a análise das colunas ou verifique se as planilhas contêm os dados necessários.")
            return
            
        mode = self.trans_mode_var.get()
        mode_text = "SOBRESCREVER os dados atuais" if mode == "overwrite" else "ADICIONAR abaixo dos dados existentes"
        
        # Alerta e Confirmação
        msg = f"Aviso de Transposição:\n\nO programa irá {mode_text} na planilha de relatório.\n\nDeseja continuar?"
        if not messagebox.askyesno("Confirmar Transposição", msg):
            return
            
        self.btn_transpose.config(state="disabled")
        self.btn_clear.config(state="disabled")
        self.lbl_action_progress.config(text="Executando transposição...", fg="blue")
        self.log("Executando transposição de dados em segundo plano...")
        
        # Inicia a thread de execução da transposição
        threading.Thread(
            target=self.transpose_thread,
            args=(src_file, src_sheet, dst_file, dst_sheet, mode),
            daemon=True
        ).start()

    def resize_excel_tables(self, sheet, header_row, last_row):
        """Redimensiona objetos de Tabela nativos (ListObjects) do Excel para manter a formatação uniforme."""
        if sheet.ListObjects.Count > 0:
            for i in range(1, sheet.ListObjects.Count + 1):
                table = sheet.ListObjects(i)
                # Célula superior esquerda da tabela
                start_cell = table.Range.Cells(1, 1)
                col_count = table.Range.Columns.Count
                start_col = table.Range.Column
                
                # Garante que a tabela tenha no mínimo 1 linha de dados além do cabeçalho
                target_last_row = max(header_row + 1, last_row)
                end_cell = sheet.Cells(target_last_row, start_col + col_count - 1)
                
                new_range = sheet.Range(start_cell, end_cell)
                table.Resize(new_range)

    def transpose_thread(self, src_file, src_sheet_name, dst_file, dst_sheet_name, mode):
        """Thread executora da transposição de dados."""
        pythoncom.CoInitialize()
        excel = None
        try:
            excel = win32com.client.Dispatch("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            
            wb_src = excel.Workbooks.Open(src_file, ReadOnly=True)
            sheet_src = wb_src.Sheets(src_sheet_name)
            
            src_header_row = self.find_header_row(sheet_src)
            sheet_src_last_row = sheet_src.UsedRange.Rows.Count
            
            # 1. Extração de dados da planilha fonte usando os mapeamentos correntes
            extracted_data = []
            
            for r in range(src_header_row + 2, sheet_src_last_row + 1):  # Começa abaixo do cabeçalho
                for m in self.current_mappings:
                    # Lê os valores
                    desc_val = sheet_src.Cells(r, m["src_desc_col"]).Value
                    price_val = sheet_src.Cells(r, m["src_price_col"]).Value
                    
                    # Trata Nota Fiscal
                    nf_val = None
                    if m.get("src_nf_col"):
                        nf_val = sheet_src.Cells(r, m["src_nf_col"]).Value
                        
                    # Trata Data (se flat)
                    date_val = None
                    if m.get("src_date_col"):
                        date_val = sheet_src.Cells(r, m["src_date_col"]).Value
                        
                    if desc_val and price_val is not None:
                        desc_str = str(desc_val).strip()
                        if desc_str == "" or "descri" in desc_str.lower():
                            continue
                            
                        try:
                            price_float = float(price_val)
                        except (ValueError, TypeError):
                            continue
                            
                        # Limpeza do número da NF
                        nf_str = ""
                        if nf_val is not None:
                            if isinstance(nf_val, float) and nf_val.is_integer():
                                nf_str = str(int(nf_val))
                            else:
                                nf_str = str(nf_val).strip()
                                if nf_str.endswith(".0"):
                                    nf_str = nf_str[:-2]
                                    
                        extracted_data.append({
                            "date": date_val,
                            "desc": desc_str,
                            "nf": nf_str,
                            "price": price_float
                        })
                        
            wb_src.Close(SaveChanges=False)
            
            if not extracted_data:
                self.root.after(0, lambda: self.finish_action("Nenhuma transação extraída.", "orange"))
                return
                
            # 2. Gravação na planilha destino
            wb_dst = excel.Workbooks.Open(dst_file)
            sheet_dst = wb_dst.Sheets(dst_sheet_name)
            
            dst_header_row = self.find_header_row(sheet_dst)
            
            # Encontra a linha de partida para gravação
            if mode == "overwrite":
                # Limpa registros antigos do destino da linha do cabeçalho+1 em diante
                last_used_row = sheet_dst.Cells(sheet_dst.Rows.Count, 2).End(xlUp).Row
                start_row = dst_header_row + 1
                if last_used_row >= start_row:
                    # Limpa todas as colunas de forma dinâmica limpando as linhas inteiras
                    sheet_dst.Rows(f"{start_row}:{last_used_row}").ClearContents()
            else:
                # Append
                last_used_row = sheet_dst.Cells(sheet_dst.Rows.Count, 2).End(xlUp).Row
                if last_used_row < dst_header_row:
                    start_row = dst_header_row + 1
                else:
                    start_row = last_used_row + 1
                    
            # Grava registros na destino
            m = self.current_mappings[0]  # O mapeamento do destino é o mesmo para todos
            
            curr_row = start_row
            for idx, item in enumerate(extracted_data):
                curr_row = start_row + idx
                
                if m["dst_date_col"]:
                    sheet_dst.Cells(curr_row, m["dst_date_col"]).Value = item["date"]
                if m["dst_desc_col"]:
                    sheet_dst.Cells(curr_row, m["dst_desc_col"]).Value = item["desc"]
                if m["dst_nf_col"]:
                    sheet_dst.Cells(curr_row, m["dst_nf_col"]).Value = item["nf"]
                if m["dst_val_col"]:
                    sheet_dst.Cells(curr_row, m["dst_val_col"]).Value = item["price"]
            
            # Redimensiona o objeto de Tabela (ListObject) do Excel para incluir as novas linhas
            try:
                self.resize_excel_tables(sheet_dst, dst_header_row, curr_row)
            except Exception as table_err:
                print(f"Aviso ao redimensionar tabela: {table_err}")

            # Copia a formatação da primeira linha de dados (template) para as novas linhas inseridas (reforço de layout)
            try:
                template_row = dst_header_row + 1
                if curr_row >= start_row:
                    sheet_dst.Rows(template_row).Copy()
                    # PasteSpecial com Paste=-4122 aplica apenas a formatação (estilos, bordas, cores, formatos numéricos)
                    sheet_dst.Rows(f"{start_row}:{curr_row}").PasteSpecial(Paste=-4122)
                    excel.CutCopyMode = False
            except Exception as format_err:
                print(f"Aviso ao copiar formatacao: {format_err}")
                    
            wb_dst.Close(SaveChanges=True)
            self.root.after(0, lambda: self.finish_action(f"Transposto: {len(extracted_data)} registros copiados!", "green"))
            
        except Exception as e:
            err_msg = str(e)
            self.root.after(0, lambda: self.finish_action("Erro ao transpor dados.", "red"))
            self.root.after(0, lambda: messagebox.showerror("Erro de Execução", f"Falha na transposição:\n{err_msg}"))
        finally:
            if excel is not None:
                try:
                    excel.Quit()
                except:
                    pass
            pythoncom.CoUninitialize()

    def run_clear_destination(self):
        """Inicia a limpeza da planilha de relatório na thread em background."""
        dst_file = self.dest_file_var.get()
        dst_sheet = self.dest_sheet_var.get()
        
        if not dst_file or not dst_sheet:
            messagebox.showwarning("Campo Vazio", "Selecione primeiro a planilha destino e a respectiva aba para limpar.")
            return
            
        if not messagebox.askyesno("Confirmar Limpeza", "Isso irá APAGAR todos os dados da tabela de destino.\nDeseja prosseguir?"):
            return
            
        self.btn_transpose.config(state="disabled")
        self.btn_clear.config(state="disabled")
        self.lbl_action_progress.config(text="Limpando planilha destino...", fg="blue")
        self.log("Executando limpeza de planilha em segundo plano...")
        
        threading.Thread(
            target=self.clear_destination_thread,
            args=(dst_file, dst_sheet),
            daemon=True
        ).start()

    def clear_destination_thread(self, dst_file, dst_sheet_name):
        """Thread executora da limpeza da planilha destino."""
        pythoncom.CoInitialize()
        excel = None
        try:
            excel = win32com.client.Dispatch("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            
            wb_dst = excel.Workbooks.Open(dst_file)
            sheet_dst = wb_dst.Sheets(dst_sheet_name)
            
            dst_header_row = self.find_header_row(sheet_dst)
            last_used_row = sheet_dst.Cells(sheet_dst.Rows.Count, 2).End(xlUp).Row
            
            start_row = dst_header_row + 1
            if last_used_row >= start_row:
                # Limpa todas as colunas de forma dinâmica limpando as linhas inteiras
                sheet_dst.Rows(f"{start_row}:{last_used_row}").ClearContents()
                
                # Redimensiona a tabela de volta para ter apenas 1 linha vazia abaixo do cabeçalho
                try:
                    self.resize_excel_tables(sheet_dst, dst_header_row, dst_header_row + 1)
                except Exception as table_err:
                    print(f"Aviso ao redimensionar tabela: {table_err}")
                
                wb_dst.Close(SaveChanges=True)
                self.root.after(0, lambda: self.finish_action("Tabela limpa com sucesso!", "green"))
            else:
                wb_dst.Close(SaveChanges=False)
                self.root.after(0, lambda: self.finish_action("A tabela já está vazia.", "orange"))
                
        except Exception as e:
            err_msg = str(e)
            self.root.after(0, lambda: self.finish_action("Erro ao limpar planilha.", "red"))
            self.root.after(0, lambda: messagebox.showerror("Erro de Execução", f"Falha ao limpar:\n{err_msg}"))
        finally:
            if excel is not None:
                try:
                    excel.Quit()
                except:
                    pass
            pythoncom.CoUninitialize()

    def finish_action(self, status_msg, color):
        """Habilita botões e atualiza status ao terminar qualquer ação."""
        self.btn_transpose.config(state="normal")
        self.btn_clear.config(state="normal")
        self.lbl_action_progress.config(text=status_msg, fg=color)
        self.log(status_msg)

if __name__ == "__main__":
    root = tk.Tk()
    app = CardGestorGUI(root)
    root.mainloop()
