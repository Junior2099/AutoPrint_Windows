import os
import subprocess
import tkinter as tk
from tkinter import messagebox, ttk
import tempfile
from PyPDF2 import PdfReader, PdfWriter
import time

# Definir a pasta de Downloads
pasta_downloads = os.path.expanduser(r"~\Downloads")

def listar_arquivos():
    """Lista os arquivos PDF na pasta Downloads"""
    return [f for f in os.listdir(pasta_downloads) if f.endswith(".pdf")]

def atualizar_lista():
    """Atualiza a lista de arquivos disponíveis"""
    # Limpar a treeview
    for item in tree.get_children():
        tree.delete(item)
    
    # Obter os arquivos
    arquivos = listar_arquivos()
    
    # Adicionar à treeview
    for i, arquivo in enumerate(arquivos):
        # Obter o tamanho do arquivo
        tamanho = os.path.getsize(os.path.join(pasta_downloads, arquivo))
        tamanho_formatado = formatar_tamanho(tamanho)
        
        # Obter a data de modificação
        data_mod = os.path.getmtime(os.path.join(pasta_downloads, arquivo))
        data_formatada = formatar_data(data_mod)
        
        tree.insert("", "end", values=(arquivo, tamanho_formatado, data_formatada))
    
    # Exibir contagem de arquivos
    status_label.config(text=f"✅ {len(arquivos)} arquivos PDF encontrados | PyPDF2 instalado")

def formatar_tamanho(tamanho_bytes):
    """Formata o tamanho do arquivo em KB, MB, etc."""
    for unidade in ['B', 'KB', 'MB', 'GB']:
        if tamanho_bytes < 1024.0:
            return f"{tamanho_bytes:.1f} {unidade}"
        tamanho_bytes /= 1024.0
    return f"{tamanho_bytes:.1f} TB"

def formatar_data(timestamp):
    """Formata o timestamp em data legível"""
    from datetime import datetime
    return datetime.fromtimestamp(timestamp).strftime('%d/%m/%Y %H:%M')

def parse_pages(pages_str, max_pages):
    """Converte uma string de páginas ('1,3,5-7') em uma lista de números de páginas"""
    pages = set()
    parts = pages_str.split(',')
    
    for part in parts:
        part = part.strip()
        if '-' in part:
            try:
                start, end = map(int, part.split('-'))
                if 1 <= start <= end <= max_pages:
                    pages.update(range(start, end + 1))
            except ValueError:
                continue
        else:
            try:
                page = int(part)
                if 1 <= page <= max_pages:
                    pages.add(page)
            except ValueError:
                continue
    
    return sorted(list(pages))

def fechar_com_mensagem(mensagem):
    """Exibe mensagem de sucesso e fecha o programa após breve pausa"""
    messagebox.showinfo("✅ Sucesso", mensagem)
    janela.after(500, janela.destroy)  # Fecha o programa após 500ms

def imprimir_arquivo():
    """Pergunta antes de imprimir e executa a impressão"""
    selecionado = tree.focus()
    if not selecionado:
        messagebox.showwarning("⚠️ Aviso", "Por favor, selecione um arquivo para imprimir.")
        return
    
    # Obter o nome do arquivo selecionado da primeira coluna
    arquivo = tree.item(selecionado)['values'][0]
    caminho_arquivo = os.path.join(pasta_downloads, arquivo)

    # Perguntar se deseja imprimir tudo ou páginas específicas
    resposta = messagebox.askquestion("Impressão", "Deseja imprimir todas as páginas?")
    
    if resposta == "yes":
        # Imprimir todas as páginas
        try:
            subprocess.run(["powershell", f'Start-Process -FilePath "{caminho_arquivo}" -Verb Print'], check=True, shell=True)
            fechar_com_mensagem(f"Impressão enviada para {arquivo}!")
        except subprocess.CalledProcessError as e:
            messagebox.showerror("❌ Erro", f"Erro ao imprimir:\n{e}")
    else:
        try:
            # Obter o número total de páginas
            with open(caminho_arquivo, 'rb') as f:
                pdf = PdfReader(f)
                total_pages = len(pdf.pages)
            
            # Criar janela personalizada para seleção de páginas
            selecionar_paginas(caminho_arquivo, arquivo, total_pages)
                
        except Exception as e:
            messagebox.showerror("❌ Erro", f"Erro ao processar o PDF:\n{e}")

def selecionar_paginas(caminho_arquivo, nome_arquivo, total_pages):
    """Cria uma janela para selecionar páginas de forma mais amigável"""
    # Criar janela
    janela_paginas = tk.Toplevel()
    janela_paginas.title(f"Selecionar Páginas - {nome_arquivo}")
    janela_paginas.geometry("500x300")
    janela_paginas.configure(bg="#f0f0f0")
    janela_paginas.transient(janela)  # Torna modal
    janela_paginas.grab_set()
    
    # Informações do arquivo
    tk.Label(janela_paginas, text=f"Arquivo: {nome_arquivo}", font=("Arial", 12), bg="#f0f0f0").pack(pady=(10, 5))
    tk.Label(janela_paginas, text=f"Total de páginas: {total_pages}", font=("Arial", 12), bg="#f0f0f0").pack(pady=(0, 10))
    
    # Frame para entrada
    frame_entrada = tk.Frame(janela_paginas, bg="#f0f0f0")
    frame_entrada.pack(pady=10, fill=tk.X, padx=20)
    
    # Explicação
    tk.Label(frame_entrada, text="Digite as páginas a imprimir:", font=("Arial", 11), bg="#f0f0f0").grid(row=0, column=0, sticky=tk.W, pady=(0, 5))
    tk.Label(frame_entrada, text="Exemplos: 1,3,5-7,10", font=("Arial", 10), bg="#f0f0f0", fg="#555555").grid(row=1, column=0, sticky=tk.W, pady=(0, 10))
    
    # Entrada de texto
    entrada_paginas = tk.Entry(frame_entrada, font=("Arial", 12), width=30)
    entrada_paginas.grid(row=2, column=0, sticky=tk.EW)
    entrada_paginas.focus_set()
    
    # Frame de botões
    frame_botoes = tk.Frame(janela_paginas, bg="#f0f0f0")
    frame_botoes.pack(pady=20)
    
    # Função para processar impressão
    def processar_impressao():
        paginas = entrada_paginas.get().strip()
        if not paginas:
            messagebox.showwarning("⚠️ Aviso", "Por favor, digite os números das páginas.")
            return
        
        # Converter a string de páginas em números
        page_numbers = parse_pages(paginas, total_pages)
        
        if not page_numbers:
            messagebox.showwarning("⚠️ Aviso", "Nenhuma página válida selecionada.")
            return
        
        # Criar um novo PDF temporário com apenas as páginas selecionadas
        pdf_writer = PdfWriter()
        
        try:
            with open(caminho_arquivo, 'rb') as f:
                pdf = PdfReader(f)
                
                # Adicionar as páginas selecionadas (compensando o índice baseado em zero)
                for page_num in page_numbers:
                    pdf_writer.add_page(pdf.pages[page_num - 1])
                
                # Salvar o PDF temporário
                with tempfile.NamedTemporaryFile(suffix='.pdf', delete=False) as temp_file:
                    temp_path = temp_file.name
                
                with open(temp_path, 'wb') as temp_pdf:
                    pdf_writer.write(temp_pdf)
            
            # Imprimir o PDF temporário
            try:
                subprocess.run(["powershell", f'Start-Process -FilePath "{temp_path}" -Verb Print'], check=True, shell=True)
                
                # Fechar janela de páginas
                janela_paginas.destroy()
                
                # Esperar um pouco antes de tentar excluir o arquivo temporário
                time.sleep(2)
                try:
                    os.unlink(temp_path)
                except:
                    pass  # Ignora erros ao excluir o arquivo temporário
                
                # Exibir mensagem e fechar o programa principal
                fechar_com_mensagem(f"Impressão de páginas {', '.join(map(str, page_numbers))} enviada!")
                    
            except subprocess.CalledProcessError as e:
                messagebox.showerror("❌ Erro", f"Erro ao imprimir páginas específicas:\n{e}")
                
        except Exception as e:
            messagebox.showerror("❌ Erro", f"Erro ao processar o PDF:\n{e}")
    
    # Botões
    btn_cancelar = tk.Button(frame_botoes, text="Cancelar", command=janela_paginas.destroy, 
                           font=("Arial", 11), bg="#f44336", fg="white", padx=15, pady=5)
    btn_cancelar.pack(side=tk.LEFT, padx=10)
    
    btn_imprimir = tk.Button(frame_botoes, text="🖨️ Imprimir", command=processar_impressao, 
                          font=("Arial", 11, "bold"), bg="#008CBA", fg="white", padx=15, pady=5)
    btn_imprimir.pack(side=tk.LEFT, padx=10)
    
    # Atalho Enter
    janela_paginas.bind('<Return>', lambda event: processar_impressao())

def abrir_pdf():
    """Abre o PDF selecionado"""
    selecionado = tree.focus()
    if not selecionado:
        messagebox.showwarning("⚠️ Aviso", "Por favor, selecione um arquivo para abrir.")
        return
    
    # Obter o nome do arquivo selecionado da primeira coluna
    arquivo = tree.item(selecionado)['values'][0]
    caminho_arquivo = os.path.join(pasta_downloads, arquivo)
    
    try:
        subprocess.run(["start", "", caminho_arquivo], shell=True)
    except Exception as e:
        messagebox.showerror("❌ Erro", f"Erro ao abrir o arquivo:\n{e}")

# Configurar estilo ttk
def configurar_estilo():
    style = ttk.Style()
    style.theme_use('clam')  # Usar um tema mais moderno
    
    # Estilo para o cabeçalho da Treeview
    style.configure("Treeview.Heading", 
                    font=('Arial', 10, 'bold'), 
                    background="#4b6584", 
                    foreground="white")
    
    # Estilo para o Treeview
    style.configure("Treeview", 
                    font=('Arial', 10), 
                    rowheight=25,
                    background="#f5f6fa",
                    fieldbackground="#f5f6fa")
    
    # Estilo para linhas selecionadas
    style.map('Treeview', background=[('selected', '#54a0ff')])

# Criar janela principal
janela = tk.Tk()
janela.title("🖨️ Impressão de Arquivos PDF")
janela.geometry("800x500")
janela.configure(bg="#f0f0f0")
janela.minsize(600, 400)  # Tamanho mínimo

# Configurar estilo
configurar_estilo()

# Título e descrição
frame_titulo = tk.Frame(janela, bg="#f0f0f0")
frame_titulo.pack(fill=tk.X, pady=(10, 5))

titulo = tk.Label(frame_titulo, text="Impressão de Arquivos PDF", font=("Arial", 16, "bold"), bg="#f0f0f0")
titulo.pack()

descricao = tk.Label(frame_titulo, text="Selecione um arquivo e imprima todas as páginas ou apenas as que você precisa", 
                    font=("Arial", 10), bg="#f0f0f0", fg="#555555")
descricao.pack(pady=(0, 10))

# Frame principal
frame_principal = tk.Frame(janela, bg="#f0f0f0")
frame_principal.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)

# Criar Treeview com colunas
columns = ("arquivo", "tamanho", "data")
tree = ttk.Treeview(frame_principal, columns=columns, show="headings", selectmode="browse")

# Definir cabeçalhos
tree.heading("arquivo", text="Nome do Arquivo")
tree.heading("tamanho", text="Tamanho")
tree.heading("data", text="Data de Modificação")

# Configurar colunas
tree.column("arquivo", width=350, anchor=tk.W)
tree.column("tamanho", width=100, anchor=tk.E)
tree.column("data", width=150, anchor=tk.CENTER)

# Adicionar barra de rolagem
scrollbar_y = ttk.Scrollbar(frame_principal, orient=tk.VERTICAL, command=tree.yview)
scrollbar_x = ttk.Scrollbar(frame_principal, orient=tk.HORIZONTAL, command=tree.xview)
tree.configure(yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)

# Posicionar elementos
scrollbar_y.pack(side=tk.RIGHT, fill=tk.Y)
tree.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
scrollbar_x.pack(side=tk.BOTTOM, fill=tk.X)

# Associar duplo clique para abrir o PDF
tree.bind("<Double-1>", lambda event: abrir_pdf())

# Frame de botões
frame_botoes = tk.Frame(janela, bg="#f0f0f0")
frame_botoes.pack(fill=tk.X, pady=15, padx=20)

# Botões
btn_atualizar = tk.Button(frame_botoes, text="🔄 Atualizar Lista", command=atualizar_lista, 
                        font=("Arial", 11), bg="#4CAF50", fg="white", padx=10, pady=5)
btn_atualizar.pack(side=tk.LEFT, padx=(0, 10))

btn_abrir = tk.Button(frame_botoes, text="📂 Abrir PDF", command=abrir_pdf, 
                     font=("Arial", 11), bg="#FF9800", fg="white", padx=10, pady=5)
btn_abrir.pack(side=tk.LEFT, padx=10)

btn_imprimir = tk.Button(frame_botoes, text="🖨️ IMPRIMIR", command=imprimir_arquivo, 
                       font=("Arial", 12, "bold"), bg="#008CBA", fg="white", padx=15, pady=7)
btn_imprimir.pack(side=tk.RIGHT)

# Status da instalação
status_frame = tk.Frame(janela, bg="#f0f0f0")
status_frame.pack(fill=tk.X, pady=5)

status_label = tk.Label(status_frame, text="", font=("Arial", 10), bg="#f0f0f0", fg="#555555")
status_label.pack(side=tk.LEFT, padx=20)

# Verificar se a biblioteca PyPDF2 está instalada
try:
    import PyPDF2
    status_label.config(text="✅ PyPDF2 instalado")
except ImportError:
    status_label.config(text="❌ PyPDF2 não instalado. Execute 'pip install PyPDF2' no terminal.")

# Atualizar lista com os arquivos da pasta Downloads
atualizar_lista()

# Rodar interface gráfica
janela.mainloop()