import sqlite3
import customtkinter as ctk
import xlwings as xw
import os
from tkinter import messagebox

ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

app = ctk.CTk()
app.geometry("420x300")
app.title("Editor de Dimensões Excel")

# Caminho do banco de dados SQLite
db_path = "configuracao.db"

# Função para criar o banco de dados e a tabela de configurações
def criar_banco_de_dados():
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()

    # Cria a tabela para armazenar o caminho do arquivo, caso não exista
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS configuracoes (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            caminho TEXT NOT NULL
        )
    """)
    
    # Commit e fechamento da conexão
    conn.commit()
    conn.close()

# Função para adicionar o caminho no banco de dados
def adicionar_caminho(caminho):
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()

    # Insere o caminho fornecido na tabela de configurações
    cursor.execute("INSERT INTO configuracoes (caminho) VALUES (?)", (caminho,))
    
    # Commit e fechamento da conexão
    conn.commit()
    conn.close()

# Função para ler o caminho do banco de dados
def ler_caminho():
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()

    # Lê o primeiro caminho armazenado no banco de dados
    cursor.execute("SELECT caminho FROM configuracoes LIMIT 1")
    caminho = cursor.fetchone()

    conn.close()

    if caminho:
        return caminho[0]
    else:
        return None

# Função para gerar as cotas e usar o caminho armazenado
def gerar_ou_editar_excel(event=None):
    # Lê o caminho do arquivo do banco de dados
    caminho_excel = ler_caminho()

    if caminho_excel:
        print("Caminho do arquivo:", caminho_excel)
        
        # O restante do código para gerar as cotas vai utilizar o caminho armazenado
        if os.path.exists(caminho_excel):
            wb = xw.Book(caminho_excel)
        else:
            wb = xw.Book()
            wb.save(caminho_excel)  # Salva o novo arquivo como .xlsm

        sht = wb.sheets[0]
        sht.name = "Dimensões"

        # Não apagar as formas VBA (alterar a exclusão para as cotas geradas)
        for shape in sht.api.Shapes:
            if shape.Name.startswith("Cota_"):  # Só remove as formas geradas pela função
                shape.Delete()

        # Criando setas e textos
        arrow_vertical = sht.api.Shapes.AddLine(80, 100, 80, 200)
        arrow_vertical.Name = "Cota_ArrowVertical"  # Identifica a seta vertical
        arrow_vertical.Line.EndArrowheadStyle = 3
        arrow_vertical.Line.BeginArrowheadStyle = 3
        arrow_vertical.Line.ForeColor.RGB = COR_VERMELHO
        arrow_vertical.Line.Weight = 1.5  # Largura mais fina

        text_v = sht.api.Shapes.AddTextbox(1, 85, 140, 50, 20)
        text_v.Name = "Cota_TextVertical"  # Identifica o texto de altura
        text_v.TextFrame2.TextRange.Text = altura
        text_v.TextFrame2.TextRange.Font.Size = 12
        text_v.TextFrame2.TextRange.ParagraphFormat.Alignment = 1
        text_v.TextFrame2.VerticalAnchor = 1
        text_v.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = COR_VERMELHO
        text_v.Line.Visible = False
        text_v.Fill.Visible = False

        # Seta horizontal
        arrow_horizontal = sht.api.Shapes.AddLine(100, 220, 200, 220)
        arrow_horizontal.Name = "Cota_ArrowHorizontal"  # Identifica a seta horizontal
        arrow_horizontal.Line.EndArrowheadStyle = 3
        arrow_horizontal.Line.BeginArrowheadStyle = 3
        arrow_horizontal.Line.ForeColor.RGB = COR_VERMELHO
        arrow_horizontal.Line.Weight = 1.5

        text_h = sht.api.Shapes.AddTextbox(1, 150, 225, 60, 20)
        text_h.Name = "Cota_TextHorizontal"  # Identifica o texto de largura
        text_h.TextFrame2.TextRange.Text = largura
        text_h.TextFrame2.TextRange.Font.Size = 12
        text_h.TextFrame2.TextRange.ParagraphFormat.Alignment = 1
        text_h.TextFrame2.VerticalAnchor = 1
        text_h.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = COR_VERMELHO
        text_h.Line.Visible = False
        text_h.Fill.Visible = False

        # Salvar o arquivo como .xlsm para permitir macros
        wb.save(caminho_excel)
        wb.app.visible = True

        # Garantir que o Excel tenha as macros habilitadas
        wb.app.api.AutomationSecurity = 3  # Habilitar macros automaticamente

        status_label.configure(text="✅ Cotas geradas com sucesso!", text_color="green")

    else:
        status_label.configure(text="Erro: Caminho do arquivo não encontrado no banco de dados.", text_color="red")


# Interface
label_altura = ctk.CTkLabel(app, text="Altura (ex: 2.00):")
label_altura.pack(pady=(20, 5))

entrada_altura = ctk.CTkEntry(app, placeholder_text="Digite a altura")
entrada_altura.pack()

label_largura = ctk.CTkLabel(app, text="Largura (ex: 5.25):")
label_largura.pack(pady=(20, 5))

entrada_largura = ctk.CTkEntry(app, placeholder_text="Digite a largura")
entrada_largura.pack()

btn_gerar = ctk.CTkButton(app, text="Gerar as cotas", command=gerar_ou_editar_excel)
btn_gerar.pack(pady=10)

status_label = ctk.CTkLabel(app, text="", text_color="green")
status_label.pack()

# Criar o banco de dados e a tabela se necessário
criar_banco_de_dados()

# Adicionar o caminho manualmente ao banco de dados, se necessário
# Exemplo: adicionar_caminho(r"G:\Meu Drive\17 - MODELOS\RELATÓRIO AUTOMATIZADO\RF MODELO COTAS\dimensoes_editaveis.xlsm")

# Enter chama o botão de gerar
app.bind("<Return>", gerar_ou_editar_excel)

app.mainloop()
