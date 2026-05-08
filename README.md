# Juntador de PDFs

## Descrição

O **Juntador de PDFs** é uma aplicação gráfica para Windows que permite unir múltiplos arquivos PDF em um único documento, de forma simples, rápida e visual. O programa oferece recursos de pré-visualização, ordenação e seleção dos arquivos, além de uma interface moderna e intuitiva.

## Funcionalidades
- Seleção de múltiplos arquivos PDF ou de uma pasta inteira
- Ordenação manual dos arquivos (subir/descer)
- Remoção de arquivos da lista
- Pré-visualização da primeira página de cada PDF (requer `pdf2image` e `poppler`)
- Barra de progresso e status do processo
- Interface gráfica amigável (Tkinter + ttk)
- Geração de executável standalone para Windows

## Requisitos
- Python 3.8+
- Windows 10/11
- As bibliotecas listadas em `requirements.txt`:
  - PyPDF2
  - Pillow
  - pdf2image
- Pasta `poppler_bin` (já inclusa no projeto) para pré-visualização de PDFs

## Instalação e Uso

### 1. Ambiente de Desenvolvimento
1. Crie o ambiente virtual:
   ```powershell
   python -m venv .venv
   .venv\Scripts\activate
   ```
2. Instale as dependências:
   ```powershell
   pip install -r requirements.txt
   ```

### 2. Executando o Projeto
```powershell
.venv\Scripts\python.exe "Juntador de PDFs.py"
```

### 3. Gerando o Executável
1. Instale o PyInstaller (se necessário):
   ```powershell
   pip install pyinstaller
   ```
2. Gere o executável:
   ```powershell
   .venv\Scripts\python.exe -m PyInstaller --onefile --noconsole --add-data "poppler_bin;poppler_bin" "Juntador de PDFs.py"
   ```
3. O executável estará em `dist\Juntador de PDFs.exe`.

### 4. Uso do Executável
- Basta executar o arquivo `Juntador de PDFs.exe` na pasta `dist`.
- Certifique-se de que a pasta `poppler_bin` está junto ao executável para a pré-visualização funcionar.

## Observações
- A pré-visualização de PDFs depende da biblioteca `pdf2image` e da pasta `poppler_bin`.
- O programa não exibe console ao ser executado como `.exe`.
- Para personalizar o ícone do executável, adicione a opção `--icon=icone.ico` ao comando do PyInstaller.

## Licença
Este projeto foi desenvolvido por Virgilio Pereira dos santos e-mail: virgiliopereira37@gmail.com Fone: (79) 9 9936-9410. Todos direitos reservados.


