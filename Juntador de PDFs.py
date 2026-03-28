import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
from PyPDF2 import PdfMerger
import threading
from PIL import Image, ImageTk
import tempfile

try:
    from pdf2image import convert_from_path
    PDF2IMAGE_AVAILABLE = True
except ImportError:
    PDF2IMAGE_AVAILABLE = False

class PDFMergerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Juntador de PDFs - Versão Melhorada")
        self.root.geometry("900x700")
        self.root.resizable(True, True)
        
        # Variáveis
        self.pasta_origem = tk.StringVar()
        self.arquivo_destino = tk.StringVar()
        self.arquivos_pdf = []  # Lista para armazenar os arquivos PDF
        self.preview_image = None
        
        # Configurar estilo
        self.configurar_estilo()
        
        self.criar_interface()
        
        # Verificar se pdf2image está disponível
        if not PDF2IMAGE_AVAILABLE:
            messagebox.showwarning(
                "Aviso", 
                "A biblioteca pdf2image não está instalada.\n"
                "A pré-visualização de PDFs não estará disponível.\n"
                "Para instalar: pip install pdf2image"
            )
        
    def configurar_estilo(self):
        """Configura o estilo da aplicação"""
        style = ttk.Style()
        style.theme_use('clam')
        
        # Configurar cores personalizadas
        style.configure('Title.TLabel', font=('Arial', 18, 'bold'), foreground='#2c3e50')
        style.configure('Heading.TLabel', font=('Arial', 12, 'bold'), foreground='#34495e')
        style.configure('Custom.TButton', font=('Arial', 10))
        
    def criar_interface(self):
        # Frame principal com padding
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configurar grid
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        
        # Título principal
        titulo = ttk.Label(main_frame, text="🔗 Juntador de PDFs", 
                          style='Title.TLabel')
        titulo.grid(row=0, column=0, columnspan=3, pady=(0, 30))
        
        # Frame para seleção de pasta e arquivo
        self.criar_secao_selecao(main_frame)
        
        # Frame principal para arquivos (dividido em duas colunas)
        files_frame = ttk.LabelFrame(main_frame, text="Gerenciamento de Arquivos", padding="15")
        files_frame.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=20)
        files_frame.columnconfigure(0, weight=1)
        files_frame.columnconfigure(1, weight=1)
        files_frame.rowconfigure(1, weight=1)
        
        # Lado esquerdo - Lista de arquivos e controles
        self.criar_secao_arquivos(files_frame)
        
        # Lado direito - Pré-visualização
        self.criar_secao_preview(files_frame)
        
        # Barra de progresso e status
        self.criar_secao_progresso(main_frame)
        
        # Botão principal
        self.btn_juntar = ttk.Button(main_frame, text="🔗 Juntar PDFs", 
                                    command=self.juntar_pdfs_thread, 
                                    state="disabled",
                                    style='Custom.TButton')
        self.btn_juntar.grid(row=6, column=0, columnspan=3, pady=20, ipadx=20, ipady=10)
        
        # Configurar redimensionamento
        main_frame.rowconfigure(3, weight=1)
        
    def criar_secao_selecao(self, parent):
        """Cria a seção de seleção de pasta e arquivo de destino"""
        # Seleção da pasta de origem
        ttk.Label(parent, text="📁 Pasta com PDFs:", style='Heading.TLabel').grid(
            row=1, column=0, sticky=tk.W, pady=5)
        
        entry_pasta = ttk.Entry(parent, textvariable=self.pasta_origem, 
                               state="readonly", width=50)
        entry_pasta.grid(row=1, column=1, sticky=(tk.W, tk.E), pady=5, padx=5)
        
        btn_pasta = ttk.Button(parent, text="Procurar", 
                              command=self.selecionar_pasta)
        btn_pasta.grid(row=1, column=2, pady=5)
        
        # Arquivo de destino
        ttk.Label(parent, text="💾 Salvar como:", style='Heading.TLabel').grid(
            row=2, column=0, sticky=tk.W, pady=5)
        
        entry_destino = ttk.Entry(parent, textvariable=self.arquivo_destino, 
                                 state="readonly", width=50)
        entry_destino.grid(row=2, column=1, sticky=(tk.W, tk.E), pady=5, padx=5)
        
        btn_destino = ttk.Button(parent, text="Procurar", 
                                command=self.selecionar_destino)
        btn_destino.grid(row=2, column=2, pady=5)
        
    def criar_secao_arquivos(self, parent):
        """Cria a seção de gerenciamento de arquivos"""
        # Frame esquerdo
        left_frame = ttk.Frame(parent)
        left_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(0, 10))
        left_frame.columnconfigure(0, weight=1)
        left_frame.rowconfigure(1, weight=1)
        
        # Título e botões de controle
        control_frame = ttk.Frame(left_frame)
        control_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        control_frame.columnconfigure(0, weight=1)
        
        ttk.Label(control_frame, text="📄 Arquivos PDF:", style='Heading.TLabel').grid(
            row=0, column=0, sticky=tk.W)
        
        # Botões de controle
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
        
        # Lista de arquivos com scrollbar
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
        """Cria a seção de pré-visualização"""
        # Frame direito
        right_frame = ttk.LabelFrame(parent, text="🔍 Pré-visualização", padding="10")
        right_frame.grid(row=0, column=1, rowspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(10, 0))
        right_frame.columnconfigure(0, weight=1)
        right_frame.rowconfigure(1, weight=1)
        
        # Label de informações
        self.preview_info = ttk.Label(right_frame, text="Selecione um arquivo para ver a pré-visualização")
        self.preview_info.grid(row=0, column=0, pady=(0, 10))
        
        # Canvas para a imagem de pré-visualização
        self.preview_canvas = tk.Canvas(right_frame, bg='white', relief='sunken', bd=2)
        self.preview_canvas.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Scrollbars para o canvas
        preview_scroll_y = ttk.Scrollbar(right_frame, orient=tk.VERTICAL, command=self.preview_canvas.yview)
        preview_scroll_y.grid(row=1, column=1, sticky=(tk.N, tk.S))
        self.preview_canvas.configure(yscrollcommand=preview_scroll_y.set)
        
        preview_scroll_x = ttk.Scrollbar(right_frame, orient=tk.HORIZONTAL, command=self.preview_canvas.xview)
        preview_scroll_x.grid(row=2, column=0, sticky=(tk.W, tk.E))
        self.preview_canvas.configure(xscrollcommand=preview_scroll_x.set)
        
    def criar_secao_progresso(self, parent):
        """Cria a seção de progresso e status"""
        # Barra de progresso
        self.progress = ttk.Progressbar(parent, mode='determinate')
        self.progress.grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=10)
        
        # Label de status
        self.status_label = ttk.Label(parent, text="Selecione uma pasta com arquivos PDF ou adicione arquivos individuais")
        self.status_label.grid(row=5, column=0, columnspan=3, pady=5)
        
    def selecionar_pasta(self):
        """Seleciona uma pasta e lista os PDFs encontrados"""
        pasta = filedialog.askdirectory(title="Selecione a pasta com os PDFs")
        if pasta:
            self.pasta_origem.set(pasta)
            self.listar_pdfs_da_pasta()
            
    def selecionar_destino(self):
        """Seleciona o arquivo de destino"""
        arquivo = filedialog.asksaveasfilename(
            title="Salvar PDF unificado como",
            defaultextension=".pdf",
            filetypes=[("Arquivo PDF", "*.pdf")]
        )
        if arquivo:
            self.arquivo_destino.set(arquivo)
            self.verificar_botao_juntar()
            
    def adicionar_arquivos(self):
        """Adiciona arquivos PDF individuais"""
        arquivos = filedialog.askopenfilenames(
            title="Selecione arquivos PDF",
            filetypes=[("Arquivo PDF", "*.pdf")]
        )
        
        for arquivo in arquivos:
            nome_arquivo = os.path.basename(arquivo)
            caminho_completo = arquivo
            
            # Verificar se o arquivo já está na lista
            if caminho_completo not in [item['caminho'] for item in self.arquivos_pdf]:
                self.arquivos_pdf.append({
                    'nome': nome_arquivo,
                    'caminho': caminho_completo
                })
        
        self.atualizar_listbox()
        self.verificar_botao_juntar()
        
    def listar_pdfs_da_pasta(self):
        """Lista todos os PDFs da pasta selecionada"""
        pasta = self.pasta_origem.get()
        if not pasta:
            return
            
        try:
            # Limpar lista atual
            self.arquivos_pdf.clear()
            
            # Encontrar todos os arquivos PDF
            arquivos_encontrados = []
            for arquivo in os.listdir(pasta):
                if arquivo.lower().endswith('.pdf'):
                    arquivos_encontrados.append(arquivo)
            
            # Ordenar alfabeticamente
            arquivos_encontrados.sort()
            
            # Adicionar à lista interna
            for arquivo in arquivos_encontrados:
                self.arquivos_pdf.append({
                    'nome': arquivo,
                    'caminho': os.path.join(pasta, arquivo)
                })
                
            self.atualizar_listbox()
            
            # Atualizar status
            if self.arquivos_pdf:
                self.status_label.config(text=f"{len(self.arquivos_pdf)} arquivo(s) PDF encontrado(s)")
            else:
                self.status_label.config(text="Nenhum arquivo PDF encontrado na pasta")
                
            self.verificar_botao_juntar()
            
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao listar arquivos: {str(e)}")
            self.status_label.config(text="Erro ao acessar a pasta")
            
    def atualizar_listbox(self):
        """Atualiza a listbox com os arquivos atuais"""
        self.listbox.delete(0, tk.END)
        for i, arquivo in enumerate(self.arquivos_pdf):
            self.listbox.insert(tk.END, f"{i+1:02d}. {arquivo['nome']}")
            
    def remover_selecionados(self):
        """Remove os arquivos selecionados da lista"""
        selecionados = self.listbox.curselection()
        if not selecionados:
            messagebox.showwarning("Aviso", "Selecione pelo menos um arquivo para remover")
            return
            
        # Remover em ordem reversa para não afetar os índices
        for i in reversed(selecionados):
            del self.arquivos_pdf[i]
            
        self.atualizar_listbox()
        self.limpar_preview()
        self.verificar_botao_juntar()
        self.atualizar_botoes_controle()
        
    def mover_para_cima(self):
        """Move o arquivo selecionado para cima na lista"""
        selecionados = self.listbox.curselection()
        if len(selecionados) != 1:
            messagebox.showwarning("Aviso", "Selecione apenas um arquivo para mover")
            return
            
        indice = selecionados[0]
        if indice > 0:
            # Trocar posições
            self.arquivos_pdf[indice], self.arquivos_pdf[indice-1] = \
                self.arquivos_pdf[indice-1], self.arquivos_pdf[indice]
            
            self.atualizar_listbox()
            self.listbox.selection_set(indice-1)
            
    def mover_para_baixo(self):
        """Move o arquivo selecionado para baixo na lista"""
        selecionados = self.listbox.curselection()
        if len(selecionados) != 1:
            messagebox.showwarning("Aviso", "Selecione apenas um arquivo para mover")
            return
            
        indice = selecionados[0]
        if indice < len(self.arquivos_pdf) - 1:
            # Trocar posições
            self.arquivos_pdf[indice], self.arquivos_pdf[indice+1] = \
                self.arquivos_pdf[indice+1], self.arquivos_pdf[indice]
            
            self.atualizar_listbox()
            self.listbox.selection_set(indice+1)
            
    def on_listbox_select(self, event):
        """Evento chamado quando um item da listbox é selecionado"""
        self.atualizar_botoes_controle()
        self.mostrar_preview()
        
    def atualizar_botoes_controle(self):
        """Atualiza o estado dos botões de controle"""
        selecionados = self.listbox.curselection()
        
        # Botão remover
        if selecionados:
            self.btn_remover.config(state="normal")
        else:
            self.btn_remover.config(state="disabled")
            
        # Botões de movimento
        if len(selecionados) == 1:
            indice = selecionados[0]
            
            # Botão subir
            if indice > 0:
                self.btn_subir.config(state="normal")
            else:
                self.btn_subir.config(state="disabled")
                
            # Botão descer
            if indice < len(self.arquivos_pdf) - 1:
                self.btn_descer.config(state="normal")
            else:
                self.btn_descer.config(state="disabled")
        else:
            self.btn_subir.config(state="disabled")
            self.btn_descer.config(state="disabled")
            
    def mostrar_preview(self):
        """Mostra a pré-visualização do arquivo selecionado"""
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
            
            # Converter primeira página para imagem
            images = convert_from_path(arquivo['caminho'], first_page=1, last_page=1, dpi=150)
            
            if images:
                # Redimensionar imagem para caber no canvas
                img = images[0]
                canvas_width = self.preview_canvas.winfo_width()
                canvas_height = self.preview_canvas.winfo_height()
                
                if canvas_width > 1 and canvas_height > 1:  # Canvas já foi renderizado
                    # Calcular novo tamanho mantendo proporção
                    img_ratio = img.width / img.height
                    canvas_ratio = canvas_width / canvas_height
                    
                    if img_ratio > canvas_ratio:
                        new_width = min(canvas_width - 20, img.width)
                        new_height = int(new_width / img_ratio)
                    else:
                        new_height = min(canvas_height - 20, img.height)
                        new_width = int(new_height * img_ratio)
                        
                    img = img.resize((new_width, new_height), Image.Resampling.LANCZOS)
                
                # Converter para PhotoImage
                self.preview_image = ImageTk.PhotoImage(img)
                
                # Limpar canvas e adicionar imagem
                self.preview_canvas.delete("all")
                self.preview_canvas.create_image(
                    self.preview_canvas.winfo_width()//2, 
                    self.preview_canvas.winfo_height()//2, 
                    image=self.preview_image, 
                    anchor=tk.CENTER
                )
                
                # Configurar região de scroll
                self.preview_canvas.configure(scrollregion=self.preview_canvas.bbox("all"))
                
                self.preview_info.config(text=f"Pré-visualização: {arquivo['nome']}")
            else:
                self.preview_info.config(text="Erro ao carregar pré-visualização")
                
        except Exception as e:
            self.preview_info.config(text=f"Erro na pré-visualização: {str(e)}")
            self.limpar_preview()
            
    def limpar_preview(self):
        """Limpa a área de pré-visualização"""
        self.preview_canvas.delete("all")
        self.preview_image = None
        self.preview_info.config(text="Selecione um arquivo para ver a pré-visualização")
        
    def verificar_botao_juntar(self):
        """Verifica se o botão de juntar deve estar habilitado"""
        if (self.arquivo_destino.get() and len(self.arquivos_pdf) > 0):
            self.btn_juntar.config(state="normal")
        else:
            self.btn_juntar.config(state="disabled")
            
    def juntar_pdfs_thread(self):
        """Executa a junção de PDFs em thread separada"""
        thread = threading.Thread(target=self.juntar_pdfs)
        thread.daemon = True
        thread.start()
        
    def juntar_pdfs(self):
        """Junta os PDFs selecionados"""
        try:
            # Desabilitar botão durante o processo
            self.btn_juntar.config(state="disabled")
            self.status_label.config(text="Juntando PDFs...")
            
            arquivo_destino = self.arquivo_destino.get()
            total_arquivos = len(self.arquivos_pdf)
            
            # Inicializar o merger
            merger = PdfMerger()
            
            # Juntar cada PDF
            for i, arquivo_info in enumerate(self.arquivos_pdf):
                try:
                    merger.append(arquivo_info['caminho'])
                    # Atualizar progresso
                    progresso = int((i + 1) / total_arquivos * 100)
                    self.progress.config(value=progresso)
                    self.status_label.config(text=f"Processando: {arquivo_info['nome']}")
                    self.root.update_idletasks()
                    
                except Exception as e:
                    messagebox.showwarning("Aviso", f"Erro ao processar {arquivo_info['nome']}: {str(e)}")
                    continue
                    
            # Salvar arquivo final
            self.status_label.config(text="Salvando arquivo final...")
            merger.write(arquivo_destino)
            merger.close()
            
            # Finalizar
            self.progress.config(value=100)
            self.status_label.config(text="PDFs unidos com sucesso!")
            
            messagebox.showinfo("Sucesso", 
                              f"PDFs unidos com sucesso!\n"
                              f"Arquivo salvo em: {arquivo_destino}")
            
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao juntar PDFs: {str(e)}")
            self.status_label.config(text="Erro durante o processo")
            
        finally:
            # Reabilitar botão
            self.btn_juntar.config(state="normal")
            self.progress.config(value=0)

def main():
    root = tk.Tk()
    app = PDFMergerApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()

