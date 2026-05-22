import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
import io
import sys
import shutil
from PyPDF2 import PdfMerger
import threading
from PIL import Image, ImageTk
import tempfile

try:
    from pdf2image import convert_from_path
    PDF2IMAGE_AVAILABLE = True
except ImportError:
    PDF2IMAGE_AVAILABLE = False

try:
    import pytesseract
    PYTESSERACT_AVAILABLE = True
except ImportError:
    PYTESSERACT_AVAILABLE = False


class PDFMergerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Ferramentas PDF")
        self.root.geometry("900x700")
        self.root.resizable(True, True)

        # Variáveis - Juntar PDFs
        self.pasta_origem = tk.StringVar()
        self.arquivo_destino = tk.StringVar()
        self.arquivos_pdf = []
        self.preview_image = None

        # Variáveis - OCR
        self.ocr_arquivo_entrada = tk.StringVar()
        self.ocr_arquivo_saida = tk.StringVar()

        # Variáveis - Comprimir PDF
        self.compress_arquivo_entrada = tk.StringVar()
        self.compress_arquivo_saida = tk.StringVar()
        self.compress_nivel = tk.StringVar(value="media")

        self.configurar_estilo()
        self.criar_interface()

        if not PDF2IMAGE_AVAILABLE:
            messagebox.showwarning(
                "Aviso",
                "A biblioteca pdf2image não está instalada.\n"
                "A pré-visualização de PDFs não estará disponível.\n"
                "Para instalar: pip install pdf2image"
            )

    def configurar_estilo(self):
        style = ttk.Style()
        style.theme_use('clam')
        style.configure('Title.TLabel', font=('Arial', 18, 'bold'), foreground='#2c3e50')
        style.configure('Heading.TLabel', font=('Arial', 12, 'bold'), foreground='#34495e')
        style.configure('Custom.TButton', font=('Arial', 10))

    def criar_interface(self):
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(0, weight=1)

        self.notebook = ttk.Notebook(main_frame)
        self.notebook.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        self.criar_aba_juntar()
        self.criar_aba_ocr()
        self.criar_aba_comprimir()

    # ── Aba: Juntar PDFs ────────────────────────────────────────────────────

    def criar_aba_juntar(self):
        aba = ttk.Frame(self.notebook, padding="20")
        self.notebook.add(aba, text="🔗 Juntar PDFs")

        aba.columnconfigure(1, weight=1)
        aba.rowconfigure(3, weight=1)

        ttk.Label(aba, text="🔗 Juntar PDFs", style='Title.TLabel').grid(
            row=0, column=0, columnspan=3, pady=(0, 20))

        self.criar_secao_selecao(aba)

        files_frame = ttk.LabelFrame(aba, text="Gerenciamento de Arquivos", padding="15")
        files_frame.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=20)
        files_frame.columnconfigure(0, weight=1)
        files_frame.columnconfigure(1, weight=1)
        files_frame.rowconfigure(1, weight=1)

        self.criar_secao_arquivos(files_frame)
        self.criar_secao_preview(files_frame)
        self.criar_secao_progresso(aba)

        self.btn_juntar = ttk.Button(aba, text="🔗 Juntar PDFs",
                                     command=self.juntar_pdfs_thread,
                                     state="disabled",
                                     style='Custom.TButton')
        self.btn_juntar.grid(row=6, column=0, columnspan=3, pady=20, ipadx=20, ipady=10)

    def criar_secao_selecao(self, parent):
        ttk.Label(parent, text="📁 Pasta com PDFs:", style='Heading.TLabel').grid(
            row=1, column=0, sticky=tk.W, pady=5)

        entry_pasta = ttk.Entry(parent, textvariable=self.pasta_origem,
                                state="readonly", width=50)
        entry_pasta.grid(row=1, column=1, sticky=(tk.W, tk.E), pady=5, padx=5)

        btn_pasta = ttk.Button(parent, text="Procurar",
                               command=self.selecionar_pasta)
        btn_pasta.grid(row=1, column=2, pady=5)

        ttk.Label(parent, text="💾 Salvar como:", style='Heading.TLabel').grid(
            row=2, column=0, sticky=tk.W, pady=5)

        entry_destino = ttk.Entry(parent, textvariable=self.arquivo_destino,
                                  state="readonly", width=50)
        entry_destino.grid(row=2, column=1, sticky=(tk.W, tk.E), pady=5, padx=5)

        btn_destino = ttk.Button(parent, text="Procurar",
                                 command=self.selecionar_destino)
        btn_destino.grid(row=2, column=2, pady=5)

    def criar_secao_arquivos(self, parent):
        left_frame = ttk.Frame(parent)
        left_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(0, 10))
        left_frame.columnconfigure(0, weight=1)
        left_frame.rowconfigure(1, weight=1)

        control_frame = ttk.Frame(left_frame)
        control_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        control_frame.columnconfigure(0, weight=1)

        ttk.Label(control_frame, text="📄 Arquivos PDF:", style='Heading.TLabel').grid(
            row=0, column=0, sticky=tk.W)

        btn_frame = ttk.Frame(control_frame)
        btn_frame.grid(row=0, column=1, sticky=tk.E)

        self.btn_adicionar = ttk.Button(btn_frame, text="➕ Adicionar",
                                        command=self.adicionar_arquivos)
        self.btn_adicionar.grid(row=0, column=0, padx=2)

        self.btn_remover = ttk.Button(btn_frame, text="➖ Remover",
                                      command=self.remover_selecionados,
                                      state="disabled")
        self.btn_remover.grid(row=0, column=1, padx=2)

        self.btn_subir = ttk.Button(btn_frame, text="⬆️",
                                    command=self.mover_para_cima,
                                    state="disabled", width=3)
        self.btn_subir.grid(row=0, column=2, padx=2)

        self.btn_descer = ttk.Button(btn_frame, text="⬇️",
                                     command=self.mover_para_baixo,
                                     state="disabled", width=3)
        self.btn_descer.grid(row=0, column=3, padx=2)

        list_frame = ttk.Frame(left_frame)
        list_frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        list_frame.columnconfigure(0, weight=1)
        list_frame.rowconfigure(0, weight=1)

        self.listbox = tk.Listbox(list_frame, height=12, selectmode=tk.EXTENDED)
        self.listbox.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        self.listbox.bind('<<ListboxSelect>>', self.on_listbox_select)

        scrollbar_y = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.listbox.yview)
        scrollbar_y.grid(row=0, column=1, sticky=(tk.N, tk.S))
        self.listbox.configure(yscrollcommand=scrollbar_y.set)

        scrollbar_x = ttk.Scrollbar(list_frame, orient=tk.HORIZONTAL, command=self.listbox.xview)
        scrollbar_x.grid(row=1, column=0, sticky=(tk.W, tk.E))
        self.listbox.configure(xscrollcommand=scrollbar_x.set)

    def criar_secao_preview(self, parent):
        right_frame = ttk.LabelFrame(parent, text="🔍 Pré-visualização", padding="10")
        right_frame.grid(row=0, column=1, rowspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(10, 0))
        right_frame.columnconfigure(0, weight=1)
        right_frame.rowconfigure(1, weight=1)

        self.preview_info = ttk.Label(right_frame, text="Selecione um arquivo para ver a pré-visualização")
        self.preview_info.grid(row=0, column=0, pady=(0, 10))

        self.preview_canvas = tk.Canvas(right_frame, bg='white', relief='sunken', bd=2)
        self.preview_canvas.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        preview_scroll_y = ttk.Scrollbar(right_frame, orient=tk.VERTICAL, command=self.preview_canvas.yview)
        preview_scroll_y.grid(row=1, column=1, sticky=(tk.N, tk.S))
        self.preview_canvas.configure(yscrollcommand=preview_scroll_y.set)

        preview_scroll_x = ttk.Scrollbar(right_frame, orient=tk.HORIZONTAL, command=self.preview_canvas.xview)
        preview_scroll_x.grid(row=2, column=0, sticky=(tk.W, tk.E))
        self.preview_canvas.configure(xscrollcommand=preview_scroll_x.set)

    def criar_secao_progresso(self, parent):
        self.progress = ttk.Progressbar(parent, mode='determinate')
        self.progress.grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=10)

        self.status_label = ttk.Label(parent, text="Selecione uma pasta com arquivos PDF ou adicione arquivos individuais")
        self.status_label.grid(row=5, column=0, columnspan=3, pady=5)

    # ── Aba: OCR ────────────────────────────────────────────────────────────

    def criar_aba_ocr(self):
        aba = ttk.Frame(self.notebook, padding="20")
        self.notebook.add(aba, text="🔍 OCR - PDF Pesquisável")

        aba.columnconfigure(1, weight=1)

        ttk.Label(aba, text="🔍 OCR - PDF Pesquisável", style='Title.TLabel').grid(
            row=0, column=0, columnspan=3, pady=(0, 20))

        ttk.Label(aba, text="📄 PDF de entrada:", style='Heading.TLabel').grid(
            row=1, column=0, sticky=tk.W, pady=5)
        ttk.Entry(aba, textvariable=self.ocr_arquivo_entrada, state="readonly", width=50).grid(
            row=1, column=1, sticky=(tk.W, tk.E), pady=5, padx=5)
        ttk.Button(aba, text="Procurar", command=self.ocr_selecionar_entrada).grid(
            row=1, column=2, pady=5)

        ttk.Label(aba, text="💾 Salvar como:", style='Heading.TLabel').grid(
            row=2, column=0, sticky=tk.W, pady=5)
        ttk.Entry(aba, textvariable=self.ocr_arquivo_saida, state="readonly", width=50).grid(
            row=2, column=1, sticky=(tk.W, tk.E), pady=5, padx=5)
        ttk.Button(aba, text="Procurar", command=self.ocr_selecionar_saida).grid(
            row=2, column=2, pady=5)

        ttk.Label(aba,
                  text="ℹ️  Requer Tesseract OCR instalado  —  github.com/tesseract-ocr/tesseract",
                  foreground='#666666').grid(row=3, column=0, columnspan=3, pady=(15, 5))

        self.ocr_progress = ttk.Progressbar(aba, mode='determinate')
        self.ocr_progress.grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=10)

        self.ocr_status_label = ttk.Label(aba, text="Selecione um PDF de imagem para converter em pesquisável")
        self.ocr_status_label.grid(row=5, column=0, columnspan=3, pady=5)

        self.btn_ocr = ttk.Button(aba, text="🔍 Gerar PDF com OCR",
                                   command=self.ocr_realizar_thread,
                                   state="disabled",
                                   style='Custom.TButton')
        self.btn_ocr.grid(row=6, column=0, columnspan=3, pady=30, ipadx=20, ipady=10)

    def ocr_selecionar_entrada(self):
        arquivo = filedialog.askopenfilename(
            title="Selecione o PDF de imagem",
            filetypes=[("Arquivo PDF", "*.pdf")]
        )
        if arquivo:
            self.ocr_arquivo_entrada.set(arquivo)
            self._ocr_atualizar_botao()

    def ocr_selecionar_saida(self):
        arquivo = filedialog.asksaveasfilename(
            title="Salvar PDF pesquisável como",
            defaultextension=".pdf",
            filetypes=[("Arquivo PDF", "*.pdf")]
        )
        if arquivo:
            self.ocr_arquivo_saida.set(arquivo)
            self._ocr_atualizar_botao()

    def _ocr_atualizar_botao(self):
        state = "normal" if (self.ocr_arquivo_entrada.get() and self.ocr_arquivo_saida.get()) else "disabled"
        self.btn_ocr.config(state=state)

    def ocr_realizar_thread(self):
        thread = threading.Thread(target=self.ocr_realizar)
        thread.daemon = True
        thread.start()

    def ocr_realizar(self):
        if not PDF2IMAGE_AVAILABLE:
            messagebox.showerror("Erro",
                "A biblioteca pdf2image não está disponível.\n"
                "Instale com: pip install pdf2image")
            return

        if not PYTESSERACT_AVAILABLE:
            messagebox.showerror("Erro",
                "A biblioteca pytesseract não está disponível.\n"
                "Instale com: pip install pytesseract\n\n"
                "Instale também o Tesseract OCR:\n"
                "github.com/tesseract-ocr/tesseract")
            return

        try:
            self.btn_ocr.config(state="disabled")
            arquivo_entrada = self.ocr_arquivo_entrada.get()
            arquivo_saida = self.ocr_arquivo_saida.get()

            self.ocr_status_label.config(text="Convertendo PDF em imagens...")
            self.root.update_idletasks()

            poppler_path = self._get_poppler_path()
            kwargs = {'dpi': 300}
            if poppler_path:
                kwargs['poppler_path'] = poppler_path

            images = convert_from_path(arquivo_entrada, **kwargs)
            total_paginas = len(images)

            pdf_pages = []
            for i, img in enumerate(images):
                self.ocr_status_label.config(
                    text=f"Processando página {i + 1} de {total_paginas}...")
                self.ocr_progress.config(value=int((i / total_paginas) * 90))
                self.root.update_idletasks()

                pdf_bytes = pytesseract.image_to_pdf_or_hocr(img, extension='pdf', lang='por')
                pdf_pages.append(io.BytesIO(pdf_bytes))

            self.ocr_status_label.config(text="Salvando PDF pesquisável...")
            self.ocr_progress.config(value=90)
            self.root.update_idletasks()

            merger = PdfMerger()
            for page in pdf_pages:
                merger.append(page)
            merger.write(arquivo_saida)
            merger.close()

            self.ocr_progress.config(value=100)
            self.ocr_status_label.config(text="PDF com OCR gerado com sucesso!")
            messagebox.showinfo("Sucesso",
                f"PDF pesquisável gerado com sucesso!\nArquivo salvo em: {arquivo_saida}")

        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao realizar OCR: {str(e)}")
            self.ocr_status_label.config(text="Erro durante o processo de OCR")

        finally:
            self._ocr_atualizar_botao()
            self.ocr_progress.config(value=0)

    # ── Aba: Comprimir PDF ──────────────────────────────────────────────────

    def criar_aba_comprimir(self):
        aba = ttk.Frame(self.notebook, padding="20")
        self.notebook.add(aba, text="🗜️ Comprimir PDF")

        aba.columnconfigure(1, weight=1)

        ttk.Label(aba, text="🗜️ Comprimir PDF", style='Title.TLabel').grid(
            row=0, column=0, columnspan=3, pady=(0, 20))

        ttk.Label(aba, text="📄 PDF de entrada:", style='Heading.TLabel').grid(
            row=1, column=0, sticky=tk.W, pady=5)
        ttk.Entry(aba, textvariable=self.compress_arquivo_entrada, state="readonly", width=50).grid(
            row=1, column=1, sticky=(tk.W, tk.E), pady=5, padx=5)
        ttk.Button(aba, text="Procurar", command=self.compress_selecionar_entrada).grid(
            row=1, column=2, pady=5)

        ttk.Label(aba, text="💾 Salvar como:", style='Heading.TLabel').grid(
            row=2, column=0, sticky=tk.W, pady=5)
        ttk.Entry(aba, textvariable=self.compress_arquivo_saida, state="readonly", width=50).grid(
            row=2, column=1, sticky=(tk.W, tk.E), pady=5, padx=5)
        ttk.Button(aba, text="Procurar", command=self.compress_selecionar_saida).grid(
            row=2, column=2, pady=5)

        nivel_frame = ttk.LabelFrame(aba, text="Nível de compressão", padding="10")
        nivel_frame.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(15, 5))

        ttk.Radiobutton(nivel_frame, text="Baixa  (maior qualidade — DPI 200, qualidade 85%)",
                        variable=self.compress_nivel, value="baixa").grid(
            row=0, column=0, sticky=tk.W, pady=2)
        ttk.Radiobutton(nivel_frame, text="Média  (equilibrado — DPI 150, qualidade 70%)",
                        variable=self.compress_nivel, value="media").grid(
            row=1, column=0, sticky=tk.W, pady=2)
        ttk.Radiobutton(nivel_frame, text="Alta   (menor arquivo — DPI 100, qualidade 50%)",
                        variable=self.compress_nivel, value="alta").grid(
            row=2, column=0, sticky=tk.W, pady=2)

        self.compress_progress = ttk.Progressbar(aba, mode='determinate')
        self.compress_progress.grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=10)

        self.compress_status_label = ttk.Label(aba, text="Selecione um PDF para comprimir")
        self.compress_status_label.grid(row=5, column=0, columnspan=3, pady=5)

        self.btn_comprimir = ttk.Button(aba, text="🗜️ Comprimir",
                                        command=self.comprimir_pdf_thread,
                                        state="disabled",
                                        style='Custom.TButton')
        self.btn_comprimir.grid(row=6, column=0, columnspan=3, pady=30, ipadx=20, ipady=10)

    def compress_selecionar_entrada(self):
        arquivo = filedialog.askopenfilename(
            title="Selecione o PDF para comprimir",
            filetypes=[("Arquivo PDF", "*.pdf")]
        )
        if arquivo:
            self.compress_arquivo_entrada.set(arquivo)
            base, _ = os.path.splitext(arquivo)
            self.compress_arquivo_saida.set(base + "_comprimido.pdf")
            self._compress_atualizar_botao()

    def compress_selecionar_saida(self):
        arquivo = filedialog.asksaveasfilename(
            title="Salvar PDF comprimido como",
            defaultextension=".pdf",
            filetypes=[("Arquivo PDF", "*.pdf")]
        )
        if arquivo:
            self.compress_arquivo_saida.set(arquivo)
            self._compress_atualizar_botao()

    def _compress_atualizar_botao(self):
        ok = bool(self.compress_arquivo_entrada.get() and self.compress_arquivo_saida.get())
        self.btn_comprimir.config(state="normal" if ok else "disabled")

    def comprimir_pdf_thread(self):
        thread = threading.Thread(target=self.comprimir_pdf)
        thread.daemon = True
        thread.start()

    def comprimir_pdf(self):
        if not PDF2IMAGE_AVAILABLE:
            messagebox.showerror("Erro",
                "A biblioteca pdf2image não está disponível.\n"
                "Instale com: pip install pdf2image")
            return

        niveis = {
            "baixa": (200, 85),
            "media": (150, 70),
            "alta":  (100, 50),
        }
        dpi, qualidade = niveis[self.compress_nivel.get()]

        try:
            self.btn_comprimir.config(state="disabled")
            arquivo_entrada = self.compress_arquivo_entrada.get()
            arquivo_saida = self.compress_arquivo_saida.get()

            self.compress_status_label.config(text="Convertendo páginas...")
            self.compress_progress.config(value=0)
            self.root.update_idletasks()

            poppler_path = self._get_poppler_path()
            kwargs = {'dpi': dpi}
            if poppler_path:
                kwargs['poppler_path'] = poppler_path

            imagens_originais = convert_from_path(arquivo_entrada, **kwargs)
            total = len(imagens_originais)

            imagens_comprimidas = []
            for i, img in enumerate(imagens_originais):
                self.compress_status_label.config(
                    text=f"Comprimindo página {i + 1} de {total}...")
                self.compress_progress.config(value=int((i / total) * 90))
                self.root.update_idletasks()

                buf = io.BytesIO()
                img.convert("RGB").save(buf, format="JPEG", quality=qualidade, optimize=True)
                buf.seek(0)
                imagens_comprimidas.append(Image.open(buf).copy())

            self.compress_status_label.config(text="Salvando PDF comprimido...")
            self.compress_progress.config(value=90)
            self.root.update_idletasks()

            imagens_comprimidas[0].save(
                arquivo_saida,
                save_all=True,
                append_images=imagens_comprimidas[1:]
            )

            tamanho_original = os.path.getsize(arquivo_entrada) / 1024
            tamanho_final = os.path.getsize(arquivo_saida) / 1024

            if tamanho_final >= tamanho_original:
                shutil.copy2(arquivo_entrada, arquivo_saida)
                tamanho_final = tamanho_original
                self.compress_progress.config(value=100)
                self.compress_status_label.config(
                    text=f"Arquivo já otimizado — cópia salva sem alteração ({tamanho_original:.0f} KB)")
                messagebox.showinfo("Sem redução",
                    f"O arquivo já está bem comprimido ou contém principalmente texto/vetores.\n"
                    f"A compressão por imagem não reduziria o tamanho ({tamanho_original:.0f} KB).\n"
                    f"O arquivo original foi copiado para: {arquivo_saida}")
                return

            reducao = (1 - tamanho_final / tamanho_original) * 100

            self.compress_progress.config(value=100)
            self.compress_status_label.config(
                text=f"Concluído! {tamanho_original:.0f} KB → {tamanho_final:.0f} KB  ({reducao:.0f}% menor)")
            messagebox.showinfo("Sucesso",
                f"PDF comprimido com sucesso!\n"
                f"Original: {tamanho_original:.0f} KB\n"
                f"Comprimido: {tamanho_final:.0f} KB  ({reducao:.0f}% menor)\n"
                f"Salvo em: {arquivo_saida}")

        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao comprimir PDF: {str(e)}")
            self.compress_status_label.config(text="Erro durante a compressão")

        finally:
            self._compress_atualizar_botao()
            self.compress_progress.config(value=0)

    # ── Helpers ─────────────────────────────────────────────────────────────

    def _get_poppler_path(self):
        if getattr(sys, 'frozen', False):
            base_path = sys._MEIPASS
        else:
            base_path = os.path.dirname(os.path.abspath(__file__))
        poppler_path = os.path.join(base_path, 'poppler_bin')
        return poppler_path if os.path.exists(poppler_path) else None

    # ── Juntar PDFs: lógica existente ───────────────────────────────────────

    def selecionar_pasta(self):
        pasta = filedialog.askdirectory(title="Selecione a pasta com os PDFs")
        if pasta:
            self.pasta_origem.set(pasta)
            self.listar_pdfs_da_pasta()

    def selecionar_destino(self):
        arquivo = filedialog.asksaveasfilename(
            title="Salvar PDF unificado como",
            defaultextension=".pdf",
            filetypes=[("Arquivo PDF", "*.pdf")]
        )
        if arquivo:
            self.arquivo_destino.set(arquivo)
            self.verificar_botao_juntar()

    def adicionar_arquivos(self):
        arquivos = filedialog.askopenfilenames(
            title="Selecione arquivos PDF",
            filetypes=[("Arquivo PDF", "*.pdf")]
        )

        for arquivo in arquivos:
            nome_arquivo = os.path.basename(arquivo)
            caminho_completo = arquivo

            if caminho_completo not in [item['caminho'] for item in self.arquivos_pdf]:
                self.arquivos_pdf.append({
                    'nome': nome_arquivo,
                    'caminho': caminho_completo
                })

        self.atualizar_listbox()
        self.verificar_botao_juntar()

    def listar_pdfs_da_pasta(self):
        pasta = self.pasta_origem.get()
        if not pasta:
            return

        try:
            self.arquivos_pdf.clear()

            arquivos_encontrados = []
            for arquivo in os.listdir(pasta):
                if arquivo.lower().endswith('.pdf'):
                    arquivos_encontrados.append(arquivo)

            arquivos_encontrados.sort()

            for arquivo in arquivos_encontrados:
                self.arquivos_pdf.append({
                    'nome': arquivo,
                    'caminho': os.path.join(pasta, arquivo)
                })

            self.atualizar_listbox()

            if self.arquivos_pdf:
                self.status_label.config(text=f"{len(self.arquivos_pdf)} arquivo(s) PDF encontrado(s)")
            else:
                self.status_label.config(text="Nenhum arquivo PDF encontrado na pasta")

            self.verificar_botao_juntar()

        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao listar arquivos: {str(e)}")
            self.status_label.config(text="Erro ao acessar a pasta")

    def atualizar_listbox(self):
        self.listbox.delete(0, tk.END)
        for i, arquivo in enumerate(self.arquivos_pdf):
            self.listbox.insert(tk.END, f"{i + 1:02d}. {arquivo['nome']}")

    def remover_selecionados(self):
        selecionados = self.listbox.curselection()
        if not selecionados:
            messagebox.showwarning("Aviso", "Selecione pelo menos um arquivo para remover")
            return

        for i in reversed(selecionados):
            del self.arquivos_pdf[i]

        self.atualizar_listbox()
        self.limpar_preview()
        self.verificar_botao_juntar()
        self.atualizar_botoes_controle()

    def mover_para_cima(self):
        selecionados = self.listbox.curselection()
        if len(selecionados) != 1:
            messagebox.showwarning("Aviso", "Selecione apenas um arquivo para mover")
            return

        indice = selecionados[0]
        if indice > 0:
            self.arquivos_pdf[indice], self.arquivos_pdf[indice - 1] = \
                self.arquivos_pdf[indice - 1], self.arquivos_pdf[indice]
            self.atualizar_listbox()
            self.listbox.selection_set(indice - 1)

    def mover_para_baixo(self):
        selecionados = self.listbox.curselection()
        if len(selecionados) != 1:
            messagebox.showwarning("Aviso", "Selecione apenas um arquivo para mover")
            return

        indice = selecionados[0]
        if indice < len(self.arquivos_pdf) - 1:
            self.arquivos_pdf[indice], self.arquivos_pdf[indice + 1] = \
                self.arquivos_pdf[indice + 1], self.arquivos_pdf[indice]
            self.atualizar_listbox()
            self.listbox.selection_set(indice + 1)

    def on_listbox_select(self, event):
        self.atualizar_botoes_controle()
        self.mostrar_preview()

    def atualizar_botoes_controle(self):
        selecionados = self.listbox.curselection()

        if selecionados:
            self.btn_remover.config(state="normal")
        else:
            self.btn_remover.config(state="disabled")

        if len(selecionados) == 1:
            indice = selecionados[0]
            self.btn_subir.config(state="normal" if indice > 0 else "disabled")
            self.btn_descer.config(state="normal" if indice < len(self.arquivos_pdf) - 1 else "disabled")
        else:
            self.btn_subir.config(state="disabled")
            self.btn_descer.config(state="disabled")

    def mostrar_preview(self):
        if not PDF2IMAGE_AVAILABLE:
            self.preview_info.config(text="Pré-visualização não disponível (pdf2image não instalado)")
            return

        selecionados = self.listbox.curselection()
        if len(selecionados) != 1:
            self.limpar_preview()
            return

        indice = selecionados[0]
        arquivo = self.arquivos_pdf[indice]

        try:
            self.preview_info.config(text=f"Carregando pré-visualização de: {arquivo['nome']}")
            self.root.update_idletasks()

            poppler_path = self._get_poppler_path()
            kwargs = {'first_page': 1, 'last_page': 1, 'dpi': 150}
            if poppler_path:
                kwargs['poppler_path'] = poppler_path

            images = convert_from_path(arquivo['caminho'], **kwargs)

            if images:
                img = images[0]
                canvas_width = self.preview_canvas.winfo_width()
                canvas_height = self.preview_canvas.winfo_height()

                if canvas_width > 1 and canvas_height > 1:
                    img_ratio = img.width / img.height
                    canvas_ratio = canvas_width / canvas_height

                    if img_ratio > canvas_ratio:
                        new_width = min(canvas_width - 20, img.width)
                        new_height = int(new_width / img_ratio)
                    else:
                        new_height = min(canvas_height - 20, img.height)
                        new_width = int(new_height * img_ratio)

                    img = img.resize((new_width, new_height), Image.Resampling.LANCZOS)

                self.preview_image = ImageTk.PhotoImage(img)

                self.preview_canvas.delete("all")
                self.preview_canvas.create_image(
                    self.preview_canvas.winfo_width() // 2,
                    self.preview_canvas.winfo_height() // 2,
                    image=self.preview_image,
                    anchor=tk.CENTER
                )
                self.preview_canvas.configure(scrollregion=self.preview_canvas.bbox("all"))
                self.preview_info.config(text=f"Pré-visualização: {arquivo['nome']}")
            else:
                self.preview_info.config(text="Erro ao carregar pré-visualização")

        except Exception as e:
            self.preview_info.config(text=f"Erro na pré-visualização: {str(e)}")
            self.limpar_preview()

    def limpar_preview(self):
        self.preview_canvas.delete("all")
        self.preview_image = None
        self.preview_info.config(text="Selecione um arquivo para ver a pré-visualização")

    def verificar_botao_juntar(self):
        if self.arquivo_destino.get() and len(self.arquivos_pdf) > 0:
            self.btn_juntar.config(state="normal")
        else:
            self.btn_juntar.config(state="disabled")

    def juntar_pdfs_thread(self):
        thread = threading.Thread(target=self.juntar_pdfs)
        thread.daemon = True
        thread.start()

    def juntar_pdfs(self):
        try:
            self.btn_juntar.config(state="disabled")
            self.status_label.config(text="Juntando PDFs...")

            arquivo_destino = self.arquivo_destino.get()
            total_arquivos = len(self.arquivos_pdf)

            merger = PdfMerger()

            for i, arquivo_info in enumerate(self.arquivos_pdf):
                try:
                    merger.append(arquivo_info['caminho'])
                    progresso = int((i + 1) / total_arquivos * 100)
                    self.progress.config(value=progresso)
                    self.status_label.config(text=f"Processando: {arquivo_info['nome']}")
                    self.root.update_idletasks()
                except Exception as e:
                    messagebox.showwarning("Aviso", f"Erro ao processar {arquivo_info['nome']}: {str(e)}")
                    continue

            self.status_label.config(text="Salvando arquivo final...")
            merger.write(arquivo_destino)
            merger.close()

            self.progress.config(value=100)
            self.status_label.config(text="PDFs unidos com sucesso!")
            messagebox.showinfo("Sucesso",
                f"PDFs unidos com sucesso!\nArquivo salvo em: {arquivo_destino}")

        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao juntar PDFs: {str(e)}")
            self.status_label.config(text="Erro durante o processo")

        finally:
            self.btn_juntar.config(state="normal")
            self.progress.config(value=0)


def main():
    root = tk.Tk()
    app = PDFMergerApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
