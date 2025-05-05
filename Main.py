import pandas as pd
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
import os
import textwrap
import sys
from datetime import datetime
from pathlib import Path
import ttkbootstrap as tb
from ttkbootstrap.constants import *

# --- CONFIGURAÇÃO INICIAL ---
CSV_ORIGINAL = 'Dados.csv'  # Nome do arquivo CSV original (embutido no executável)
BACKUP_DIR = 'backups'  # Diretório para backups

ULTIMA_MODIFICACAO = None
ultimos_resultados = pd.DataFrame()

def caminho_dados():
    base_path = Path(sys.executable).parent if getattr(sys, 'frozen', False) else Path(__file__).parent
    return base_path / "Dados.csv"
    
def caminho_recurso(relativo):
    """Garante o acesso ao arquivo tanto em .py quanto em .exe"""
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relativo)
    return os.path.join(os.path.abspath("."), relativo)

def criar_backup():
    """Cria um backup do arquivo CSV antes de edições"""
    backup_dir = Path(caminho_dados()).parent / BACKUP_DIR
    if not backup_dir.exists():
        backup_dir.mkdir(parents=True, exist_ok=True)
    
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_path = backup_dir / f"backup_{timestamp}.csv"
    
    try:
        df.to_csv(backup_path, sep=';', index=False, encoding='ISO-8859-1')
        print(f"Backup criado: {backup_path}")
    except Exception as e:
        print(f"Erro ao criar backup: {e}")

# Carregar o arquivo CSV
caminho_csv = caminho_dados()
caminho_csv_original = caminho_recurso(CSV_ORIGINAL)

try:
    if caminho_csv.exists():
        # Carrega o CSV editável se existir
        df = pd.read_csv(caminho_csv, encoding='ISO-8859-1', sep=';', on_bad_lines='skip')
    else:
        # Se não existir, carrega o CSV original e cria uma cópia editável
        df = pd.read_csv(caminho_csv_original, encoding='ISO-8859-1', sep=';', on_bad_lines='skip')
        df.to_csv(caminho_csv, sep=';', index=False, encoding='ISO-8859-1')
except Exception as e:
    messagebox.showerror("Erro", f"Não foi possível carregar o arquivo CSV:\n{e}")
    df = pd.DataFrame()

# --- CARREGAR DADOS ---
def carregar_dados():
    global df, ULTIMA_MODIFICACAO
    try:
        mod = os.path.getmtime(caminho_csv) 
        if ULTIMA_MODIFICACAO is None or mod != ULTIMA_MODIFICACAO:
            df = pd.read_csv(caminho_csv, encoding='ISO-8859-1', sep=';', on_bad_lines='skip')
            ULTIMA_MODIFICACAO = mod
            print("📁 Planilha atualizada.")
    except Exception as e:
        print(f"Erro ao carregar CSV: {e}")

# --- FUNÇÕES DE INTERFACE ---
def buscar_texto():
    termo = entrada.get().strip().lower()
    if not termo:
        messagebox.showinfo("Atenção", "Digite algo para buscar.")
        return
    
    try:
        resultado = df[df.apply(lambda row: termo in str(row.values).lower(), axis=1)]
        exibir_resultados(resultado)
    except Exception as e:
        messagebox.showerror("Erro", f"Ocorreu um erro na busca:\n{e}")

def aplicar_filtros():
    tipo = tipo_combo.get()
    grupo = grupo_combo.get()
    local = local_combo.get()

    try:
        filtro = df.copy()
        if tipo:
            filtro = filtro[filtro['Tipo'] == tipo]
        if grupo:
            filtro = filtro[filtro['Grupo encarregado'] == grupo]
        if local:
            filtro = filtro[filtro['Localização'] == local]

        exibir_resultados(filtro)
    except Exception as e:
        messagebox.showerror("Erro", f"Ocorreu um erro ao filtrar:\n{e}")

def limpar_tudo():
    entrada.delete(0, tk.END)
    tipo_combo.set('')
    grupo_combo.set('')
    local_combo.set('')
    resultado_texto.delete('1.0', tk.END)

def formatar_inventario(valor):
    """Formata o número de inventário corretamente, removendo .0 se existir"""
    try:
        if pd.isna(valor):
            return 'Não encontrado'
        if isinstance(valor, (int, float)):
            # Remove .0 de números inteiros
            return str(int(valor)) if valor == int(valor) else str(valor)
        # Se for string, tenta converter para número primeiro
        try:
            num = float(valor)
            return str(int(num)) if num == int(num) else str(num)
        except ValueError:
            return str(valor)
    except:
        return 'Inválido'

def exibir_resultados(resultados):
    global ultimos_resultados
    ultimos_resultados = resultados

    resultado_texto.config(state=tk.NORMAL)
    resultado_texto.delete('1.0', tk.END)
    
    if resultados.empty:
        resultado_texto.insert(tk.END, "❌ Nenhum item encontrado.\n")
    else:
        for idx, (_, row) in enumerate(resultados.iterrows(), 1):
            inventario = formatar_inventario(row.get('Número de inventário', ''))
            nome = row.get('Nome', 'Não encontrado') if pd.notna(row.get('Nome')) else 'Não encontrado'
            tipo = row.get('Tipo', 'Não encontrado') if pd.notna(row.get('Tipo')) else 'Não encontrado'
            grupo = row.get('Grupo encarregado', 'Não encontrado') if pd.notna(row.get('Grupo encarregado')) else 'Não encontrado'
            local = row.get('Localização', 'Não encontrado') if pd.notna(row.get('Localização')) else 'Não encontrado'

            resultado_texto.insert(tk.END,
                f"🔢 Índice: {idx}\n"
                f"🆔 Patrimônio: {inventario}\n"
                f"📄 Nome: {nome}\n"
                f"🛠 Tipo: {tipo}\n"
                f"👥 Grupo: {grupo}\n"
                f"📍 Localização: {local}\n"
                f"{'-' * 50}\n"
            )
    
    resultado_texto.config(state=tk.DISABLED)

def exportar_excel():
    if ultimos_resultados.empty:
        messagebox.showwarning("Nada para exportar", "Realize uma busca primeiro.")
        return
    
    try:
        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx"), ("Todos os arquivos", "*.*")],
            title="Salvar como Excel"
        )
        if path:
            ultimos_resultados.to_excel(path, index=False)
            messagebox.showinfo("Sucesso", f"Exportado para Excel:\n{os.path.basename(path)}")
    except Exception as e:
        messagebox.showerror("Erro", f"Falha ao exportar para Excel:\n{e}")

def exportar_pdf():
    if ultimos_resultados.empty:
        messagebox.showwarning("Nada para exportar", "Realize uma busca primeiro.")
        return

    try:
        path = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf"), ("Todos os arquivos", "*.*")],
            title="Salvar como PDF"
        )
        if not path:
            return

        c = canvas.Canvas(path, pagesize=A4)
        width, height = A4
        y = height - 40
        c.setFont("Helvetica", 9)

        # Cabeçalho
        c.setFont("Helvetica-Bold", 12)
        c.drawString(40, y, "Relatório de Patrimônio - " + datetime.now().strftime("%d/%m/%Y %H:%M"))
        y -= 20
        c.setFont("Helvetica", 9)
        c.line(40, y, width-40, y)
        y -= 20

        for _, row in ultimos_resultados.iterrows():
            inventario = formatar_inventario(row.get('Número de inventário', ''))
            nome = row.get('Nome', 'Não encontrado') if pd.notna(row.get('Nome')) else 'Não encontrado'
            tipo = row.get('Tipo', 'Não encontrado') if pd.notna(row.get('Tipo')) else 'Não encontrado'
            grupo = row.get('Grupo encarregado', 'Não encontrado') if pd.notna(row.get('Grupo encarregado')) else 'Não encontrado'
            local = row.get('Localização', 'Não encontrado') if pd.notna(row.get('Localização')) else 'Não encontrado'

            linha = f"Patrimônio: {inventario} | Nome: {nome} | Tipo: {tipo} | Grupo: {grupo} | Local: {local}"

            for sublinha in textwrap.wrap(linha, width=110):
                c.drawString(40, y, sublinha)
                y -= 15

            y -= 10  # Espaço entre itens

            if y < 60:  # Nova página se necessário
                c.showPage()
                y = height - 40
                c.setFont("Helvetica", 9)

        c.save()
        messagebox.showinfo("Sucesso", f"PDF salvo como:\n{os.path.basename(path)}")
    except Exception as e:
        messagebox.showerror("Erro", f"Falha ao exportar para PDF:\n{e}")

def atualizar_interface():
    try:
        carregar_dados()
        tipo_combo['values'] = sorted(df['Tipo'].dropna().unique().tolist())
        grupo_combo['values'] = sorted(df['Grupo encarregado'].dropna().unique().tolist())
        local_combo['values'] = sorted(df['Localização'].dropna().unique().tolist())
    except Exception as e:
        print(f"Erro ao atualizar interface: {e}")
    
    janela.after(30000, atualizar_interface)  # Atualiza a cada 30 segundos

def selecionar_item_para_edicao():
    """Permite selecionar um item diretamente da lista de resultados para edição"""
    if ultimos_resultados.empty:
        messagebox.showinfo("Atenção", "Nenhum resultado para editar.")
        return
    
    try:
        # Obter a seleção atual no widget de texto
        sel_start = resultado_texto.index(tk.SEL_FIRST)
        sel_end = resultado_texto.index(tk.SEL_LAST)
        selected_text = resultado_texto.get(sel_start, sel_end)
        
        # Extrair o número do índice da linha selecionada
        for line in selected_text.split('\n'):
            if line.startswith("🔢 Índice:"):
                idx = int(line.split(":")[1].strip())
                abrir_janela_edicao(idx)
                return
                
        messagebox.showinfo("Atenção", "Selecione uma linha que comece com 'Índice:' para editar.")
        
    except tk.TclError:
        # Se não houver seleção, perguntar pelo índice
        perguntar_indice_edicao()

def perguntar_indice_edicao():
    """Abre uma janela para perguntar qual item editar pelo índice"""
    janela_indice = tb.Toplevel(janela)
    janela_indice.title("Editar Item por Índice")
    janela_indice.geometry("300x150")
    
    main_frame = tb.Frame(janela_indice)
    main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
    
    tb.Label(
        main_frame,
        text=f"Digite o índice do item (1 a {len(ultimos_resultados)}):",
        bootstyle="primary"
    ).pack(pady=5)
    
    entrada_indice = tb.Entry(main_frame)
    entrada_indice.pack(pady=10, fill=tk.X)
    
    button_frame = tb.Frame(main_frame)
    button_frame.pack(fill=tk.X)
    
    tb.Button(
        button_frame,
        text="Editar",
        command=lambda: validar_indice_edicao(entrada_indice.get(), janela_indice),
        bootstyle="success"
    ).pack(side=tk.LEFT, padx=5, expand=True)
    
    tb.Button(
        button_frame,
        text="Cancelar",
        command=janela_indice.destroy,
        bootstyle="danger"
    ).pack(side=tk.LEFT, padx=5, expand=True)
    
    entrada_indice.focus_set()
    
    # Centralizar janela
    janela_indice.update_idletasks()
    width = janela_indice.winfo_width()
    height = janela_indice.winfo_height()
    x = (janela_indice.winfo_screenwidth() // 2) - (width // 2)
    y = (janela_indice.winfo_screenheight() // 2) - (height // 2)
    janela_indice.geometry(f'+{x}+{y}')

def validar_indice_edicao(indice_str, janela_indice):
    """Valida o índice digitado pelo usuário"""
    try:
        idx = int(indice_str)
        if 1 <= idx <= len(ultimos_resultados):
            janela_indice.destroy()
            abrir_janela_edicao(idx)
        else:
            messagebox.showerror(
                "Erro", 
                f"Índice inválido. Digite um número entre 1 e {len(ultimos_resultados)}."
            )
    except ValueError:
        messagebox.showerror("Erro", "Por favor, digite um número válido.")

def abrir_janela_edicao(idx):
    """Abre a janela de edição para o item selecionado"""
    global df, ultimos_resultados
    
    try:
        # Ajuste para índice baseado em 0
        idx = idx - 1
        if idx < 0 or idx >= len(ultimos_resultados):
            messagebox.showerror("Erro", "Índice inválido.")
            return

        # Recarregar os dados para garantir sincronia
        carregar_dados()
        
        # Obter a linha dos resultados filtrados
        linha_filtrada = ultimos_resultados.iloc[idx]
        inventario = formatar_inventario(linha_filtrada["Número de inventário"])
        
        # Remover .0 se existir no número do patrimônio
        inventario_busca = str(inventario).replace('.0', '') if '.0' in str(inventario) else str(inventario)
        
        # Encontrar a linha correspondente no DataFrame
        # Converter ambos os lados para string e remover .0 para comparação
        mascara = df["Número de inventário"].astype(str).str.replace('.0', '') == inventario_busca
        linha_df = df[mascara]
        
        if linha_df.empty:
            # Debug detalhado para ajudar a identificar o problema
            print("\nDEBUG - Falha ao localizar item:")
            print(f"Patrimônio buscado: '{inventario_busca}'")
            print("Valores de patrimônio no DataFrame (amostra):")
            print(df["Número de inventário"].astype(str).unique()[:10])
            
            messagebox.showerror(
                "Erro", 
                f"Item não encontrado na base de dados.\n\n"
                f"Patrimônio: {inventario_busca}"
            )
            return

        # Obter o índice no DataFrame
        idx_original = linha_df.index[0]
        
        def salvar_edicao():
            """Salva as alterações no arquivo CSV"""
            global df
            try:
                criar_backup()  # Sempre criar backup antes de editar
                
                # Atualizar os valores no DataFrame
                df.at[idx_original, "Nome"] = entrada_nome.get()
                df.at[idx_original, "Tipo"] = entrada_tipo.get()
                df.at[idx_original, "Grupo encarregado"] = entrada_grupo.get()
                df.at[idx_original, "Localização"] = entrada_local.get()

                # Salvar no arquivo CSV editável
                df.to_csv(caminho_csv, sep=';', index=False, encoding='ISO-8859-1')
                
                messagebox.showinfo("Sucesso", "Alterações salvas com sucesso!")
                janela_edicao.destroy()
                
                # Atualizar a interface
                carregar_dados()
                if entrada.get().strip():  # Se havia uma busca ativa
                    buscar_texto()
                else:
                    aplicar_filtros()
                    
            except Exception as e:
                messagebox.showerror("Erro", f"Falha ao salvar alterações:\n{e}")

        # Criar janela de edição
        janela_edicao = tb.Toplevel(janela)
        janela_edicao.title(f"Editar Item - Patrimônio {inventario}")
        janela_edicao.geometry("500x350")
        
        # Frame principal
        main_edit_frame = tb.Frame(janela_edicao)
        main_edit_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Título
        tb.Label(
            main_edit_frame,
            text=f"Editando Patrimônio: {inventario}",
            font=("Helvetica", 12, "bold"),
            bootstyle="primary"
        ).pack(pady=(0, 15))
        
        # Campos de edição em um frame
        edit_fields = tb.Frame(main_edit_frame)
        edit_fields.pack(fill=tk.X, pady=5)
        
        # Criar campos com labels e entradas
        campos = [
            ("Nome:", "Nome", 0),
            ("Tipo:", "Tipo", 1),
            ("Grupo encarregado:", "Grupo encarregado", 2),
            ("Localização:", "Localização", 3)
        ]
        
        widgets = {}
        for label, coluna, row in campos:
            tb.Label(edit_fields, text=label).grid(row=row, column=0, sticky=tk.W, pady=5, padx=5)
            entry = tb.Entry(edit_fields, width=40)
            entry.grid(row=row, column=1, pady=5, padx=5)
            entry.insert(0, linha_df.iloc[0][coluna] if pd.notna(linha_df.iloc[0][coluna]) else "")
            widgets[coluna] = entry
        
        # Referências aos widgets
        entrada_nome = widgets["Nome"]
        entrada_tipo = widgets["Tipo"]
        entrada_grupo = widgets["Grupo encarregado"]
        entrada_local = widgets["Localização"]
        
        # Frame de botões
        button_frame = tb.Frame(main_edit_frame)
        button_frame.pack(pady=10)
        
        tb.Button(
            button_frame,
            text="Salvar Alterações",
            command=salvar_edicao,
            bootstyle="success"
        ).pack(side=tk.LEFT, padx=10)
        
        tb.Button(
            button_frame,
            text="Cancelar",
            command=janela_edicao.destroy,
            bootstyle="danger"
        ).pack(side=tk.LEFT, padx=10)
        
        # Centralizar janela
        janela_edicao.update_idletasks()
        width = janela_edicao.winfo_width()
        height = janela_edicao.winfo_height()
        x = (janela_edicao.winfo_screenwidth() // 2) - (width // 2)
        y = (janela_edicao.winfo_screenheight() // 2) - (height // 2)
        janela_edicao.geometry(f'+{x}+{y}')
        
        entrada_nome.focus_set()
        
    except Exception as e:
        messagebox.showerror("Erro", f"Falha ao editar item:\n{e}")

# --- INTERFACE PRINCIPAL ---
janela = tb.Window(themename="flatly")
janela.title("Consulta de Patrimônio")
janela.geometry("1000x700")
janela.minsize(800, 600)

# Frame principal
main_frame = tb.Frame(janela)
main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

# Título principal
title_label = tb.Label(
    main_frame, 
    text="Consulta de Patrimônio", 
    font=("Helvetica", 20, "bold"),
    bootstyle="primary"
)
title_label.pack(pady=(0, 15))

# Frame de busca
frame_busca = tb.LabelFrame(main_frame, text="Busca Textual", bootstyle="primary")
frame_busca.pack(fill=tk.X, pady=5)

entrada = tb.Entry(
    frame_busca,
    width=50,
    bootstyle="primary"
)
entrada.pack(side=tk.LEFT, padx=5, pady=5, fill=tk.X, expand=True)

search_btn = tb.Button(
    frame_busca,
    text="Buscar",
    command=buscar_texto,
    bootstyle="primary",
    width=10
)
search_btn.pack(side=tk.LEFT, padx=5)

# Separador
tb.Separator(main_frame, bootstyle="primary").pack(fill=tk.X, pady=10)

# Frame de filtros
frame_filtros = tb.LabelFrame(main_frame, text="Filtros Avançados", bootstyle="primary")
frame_filtros.pack(fill=tk.X, pady=5)

# Comboboxes para filtros
tb.Label(frame_filtros, text="Tipo:").grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
tipo_combo = tb.Combobox(frame_filtros, state="readonly", width=25, bootstyle="primary")
tipo_combo.grid(row=0, column=1, padx=5, pady=5)

tb.Label(frame_filtros, text="Grupo:").grid(row=0, column=2, padx=5, pady=5, sticky=tk.W)
grupo_combo = tb.Combobox(frame_filtros, state="readonly", width=25, bootstyle="primary")
grupo_combo.grid(row=0, column=3, padx=5, pady=5)

tb.Label(frame_filtros, text="Localização:").grid(row=0, column=4, padx=5, pady=5, sticky=tk.W)
local_combo = tb.Combobox(frame_filtros, state="readonly", width=25, bootstyle="primary")
local_combo.grid(row=0, column=5, padx=5, pady=5)

# Botões de ação
filter_btn = tb.Button(
    frame_filtros,
    text="Filtrar",
    command=aplicar_filtros,
    bootstyle="primary"
)
filter_btn.grid(row=0, column=6, padx=10)

clear_btn = tb.Button(
    frame_filtros,
    text="Limpar",
    command=limpar_tudo,
    bootstyle="danger"
)
clear_btn.grid(row=0, column=7, padx=5)

# Frame de ações
action_frame = tb.Frame(main_frame)
action_frame.pack(fill=tk.X, pady=10)

# Botões de exportação e edição
export_excel_btn = tb.Button(
    action_frame,
    text="Exportar Excel",
    command=exportar_excel,
    bootstyle="success",
    width=15
)
export_excel_btn.pack(side=tk.LEFT, padx=5)

export_pdf_btn = tb.Button(
    action_frame,
    text="Exportar PDF",
    command=exportar_pdf,
    bootstyle="info",
    width=15
)
export_pdf_btn.pack(side=tk.LEFT, padx=5)

edit_btn = tb.Button(
    action_frame,
    text="Editar Item",
    command=selecionar_item_para_edicao,
    bootstyle="warning",
    width=15
)
edit_btn.pack(side=tk.LEFT, padx=5)

# Frame de resultados
frame_resultado = tb.Frame(main_frame, bootstyle="default")
frame_resultado.pack(fill=tk.BOTH, expand=True, pady=10)

# Área de texto com scrollbar
text_scroll = tb.Scrollbar(frame_resultado)
text_scroll.pack(side=tk.RIGHT, fill=tk.Y)

resultado_texto = tb.Text(
    frame_resultado,
    wrap="word",
    height=10,
    width=50
)
resultado_texto.pack(fill=tk.BOTH, expand=True)
text_scroll.config(command=resultado_texto.yview)

style = ttk.Style()
style.configure("Custom.TButton", font=("Arial", 12), background="black")

# Barra de status
status_bar = tb.Frame(janela, bootstyle="secondary")
status_bar.pack(side=tk.BOTTOM, fill=tk.X)

tb.Label(
    status_bar,
    text="🕒 Atualização automática a cada 30 segundos | Ctrl+E para editar item selecionado",
    bootstyle="inverse-secondary",
    font=("Segoe UI", 9)
).pack(side=tk.LEFT, padx=10, pady=3)

# Configuração de atalhos
entrada.bind("<Return>", lambda event: buscar_texto())
janela.bind('<Control-e>', lambda e: selecionar_item_para_edicao())

# --- INICIALIZAÇÃO ---
if __name__ == "__main__":
    # Verificar se o arquivo CSV existe
    if not os.path.exists(caminho_csv):
        messagebox.showerror("Erro", f"Arquivo CSV não encontrado em:\n{caminho_csv}")
        janela.destroy()
    else:
        # Iniciar interface
        carregar_dados()
        atualizar_interface()
        janela.mainloop()