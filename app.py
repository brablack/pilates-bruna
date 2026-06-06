from flask import Flask, render_template, request, redirect, session
import sqlite3
from flask import send_file
from openpyxl import Workbook
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from datetime import datetime
import os

app = Flask(__name__)
app.secret_key = "chave_secreta_pilates"

def conectar():
    return sqlite3.connect("banco.db")

def criar_banco():
    conexao = conectar()
    cursor = conexao.cursor()

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS usuarios (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            usuario TEXT NOT NULL UNIQUE,
            senha TEXT NOT NULL,
            nivel TEXT NOT NULL DEFAULT 'usuario'
        )
    """)

    cursor.execute("""
        INSERT OR IGNORE INTO usuarios (usuario, senha, nivel)
        VALUES ('admin', '1234', 'admin')
    """)
    
    cursor.execute("""
        INSERT OR IGNORE INTO usuarios (usuario, senha, nivel)
        VALUES ('professor', '1234', 'professor')
    """)
    
    cursor.execute("""
            CREATE TABLE IF NOT EXISTS alunos (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                nome TEXT NOT NULL,
                endereco TEXT, 
                telefone TEXT,
                email TEXT, 
                data_nascimento TEXT,
                observacoes TEXT
            )
        """)
    
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS agenda(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            aluno_id INTEGER NOT NULL,
            data_aula TEXT NOT NULL, 
            horario TEXT NOT NULL,
            tipo_aula TEXT,
            status TEXT DEFAULT 'Agendada',
            observacoes TEXT,
            FOREIGN KEY (aluno_id) REFERENCES alunos(id)
        )
    """)
    
    cursor.execute("""
       CREATE TABLE IF NOT EXISTS mensalidades(
           id INTEGER PRIMARY KEY AUTOINCREMENT, 
           aluno_id INTEGER NOT NULL,
           mes TEXT NOT NULL,
           valor REAL NOT NULL,
           data_vencimento TEXT,
           status TEXT DEFAULT 'Pendente',
           FOREIGN KEY (aluno_id) REFERENCES alunos(id)
        ) 
    """)

    conexao.commit()
    conexao.close()
    
def somente_admin():
    return session.get("nivel") == "admin"

criar_banco()
@app.route("/", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        usuario = request.form["usuario"]
        senha = request.form["senha"]

        conexao = conectar()
        cursor = conexao.cursor()

        cursor.execute(
            "SELECT * FROM usuarios WHERE usuario = ? AND senha = ?",
            (usuario, senha)
        )

        usuario_encontrado = cursor.fetchone()
        conexao.close()

        if usuario_encontrado:
            session["usuario"] = usuario_encontrado[1]
            session["nivel"] = usuario_encontrado[3]
            return redirect("/dashboard")
        else:
            return render_template("login.html", erro="Usuário ou senha incorretos")

    return render_template("login.html")

@app.route("/alunos")
def alunos():
    if "usuario" not in session:
        return redirect("/")
    
    conexao = conectar()
    cursor = conexao.cursor()
    
    cursor.execute("SELECT * FROM alunos ORDER BY nome")
    lista_alunos = cursor.fetchall()
    
    conexao.close()
    
    return render_template("alunos.html", alunos=lista_alunos)

@app.route("/dashboard")
def dashboard():
    if "usuario" not in session:
        return redirect("/")

    hoje = datetime.now().strftime("%Y-%m-%d")
    mes_atual = datetime.now().strftime("%Y-%m")

    conexao = conectar()
    cursor = conexao.cursor()

    cursor.execute("SELECT COUNT(*) FROM alunos")
    total_alunos = cursor.fetchone()[0]

    cursor.execute("""
        SELECT COUNT(*) FROM agenda
        WHERE data_aula = ?
    """, (hoje,))
    aulas_hoje = cursor.fetchone()[0]

    cursor.execute("""
        SELECT COUNT(*) FROM agenda
        WHERE data_aula LIKE ?
    """, (mes_atual + "%",))
    aulas_mes = cursor.fetchone()[0]

    cursor.execute("""
        SELECT COUNT(*) FROM agenda
        WHERE status = 'Presente' AND data_aula LIKE ?
    """, (mes_atual + "%",))
    total_presencas = cursor.fetchone()[0]

    cursor.execute("""
        SELECT COUNT(*) FROM agenda
        WHERE status = 'Faltou' AND data_aula LIKE ?
    """, (mes_atual + "%",))
    total_faltas = cursor.fetchone()[0]

    cursor.execute("""
        SELECT COUNT(*) FROM agenda
        WHERE status = 'Cancelada' AND data_aula LIKE ?
    """, (mes_atual + "%",))
    total_canceladas = cursor.fetchone()[0]

    cursor.execute("""
        SELECT agenda.id, alunos.nome, agenda.data_aula,
               agenda.horario, agenda.tipo_aula, agenda.status
        FROM agenda
        JOIN alunos ON alunos.id = agenda.aluno_id
        WHERE agenda.data_aula LIKE ?
        ORDER BY agenda.data_aula, agenda.horario
    """, (mes_atual + "%",))
    aulas_mes_lista = cursor.fetchall()

    cursor.execute("""
        SELECT alunos.nome, mensalidades.valor,
               mensalidades.data_vencimento, mensalidades.status
        FROM mensalidades
        JOIN alunos ON alunos.id = mensalidades.aluno_id
        WHERE mensalidades.status != 'Pago'
        ORDER BY mensalidades.data_vencimento
        LIMIT 5
    """)
    mensalidades_pendentes = cursor.fetchall()

    conexao.close()

    return render_template(
        "dashboard.html",
        total_alunos=total_alunos,
        aulas_hoje=aulas_hoje,
        aulas_mes=aulas_mes,
        total_presencas=total_presencas,
        total_faltas=total_faltas,
        total_canceladas=total_canceladas,
        aulas_mes_lista=aulas_mes_lista,
        mensalidades_pendentes=mensalidades_pendentes
    )

@app.route("/agenda")
def agenda():
    if "usuario" not in session:
        return redirect("/")
    
    conexao = conectar()
    cursor = conexao.cursor()
    
    cursor.execute("SELECT id, nome FROM alunos ORDER BY nome")
    alunos = cursor.fetchall()
    
    cursor.execute("""
        SELECT agenda.id, alunos.nome, agenda.data_aula, agenda.horario,
               agenda.tipo_aula, agenda.status, agenda.observacoes
        FROM agenda
        JOIN alunos ON alunos.id = agenda.aluno_id
        ORDER BY agenda.data_aula, agenda.horario
    """)
    aulas = cursor.fetchall()
    
    conexao.close()
    
    return render_template("agenda.html", alunos=alunos, aulas=aulas)

@app.route("/cadastrar_aula", methods=["POST"])
def cadastrar_aula():
    if "usuario" not in session:
        return redirect("/")
    
    aluno_id = request.form["aluno_id"]
    data_aula = request.form["data_aula"]
    horario = request.form["horario"]
    tipo_aula = request.form["tipo_aula"]
    status = request.form["status"]
    observacoes = request.form["observacoes"]
    
    conexao = conectar()
    cursor = conexao.cursor()
    
    cursor.execute("""
        INSERT INTO agenda
        (aluno_id, data_aula, horario, tipo_aula, status, observacoes)
        VALUES (?, ?, ?, ?, ?, ?)
    """, (aluno_id, data_aula, horario, tipo_aula, status, observacoes))
    
    conexao.commit()
    conexao.close()
    
    return redirect("/agenda")

@app.route("/atualizar_presenca/<int:id>/<status>")
def atualizar_presenca(id, status):
    if "usuario" not in session:
        return redirect("/")
    
    conexao = conectar()
    cursor = conexao.cursor()
    
    cursor.execute("""
        UPDATE agenda
        SET status = ?
        WHERE id = ?
    """, (status, id))
    
    conexao.commit()
    conexao.close()
    
    return redirect("/agenda")

@app.route("/mensalidades")
def mensalidades():
    if "usuario" not in session:
        return redirect("/")
    
    if not somente_admin():
        return redirect("/dashboard")
    
    conexao = conectar()
    cursor = conexao.cursor()
    
    cursor.execute("SELECT id, nome FROM alunos ORDER BY nome")
    alunos = cursor.fetchall()
    
    cursor.execute("""
        SELECT mensalidades.id, alunos.nome, mensalidades.mes,
               mensalidades.valor, mensalidades.data_vencimento,
               mensalidades.status
        FROM mensalidades
        JOIN alunos ON alunos.id = mensalidades.aluno_id
        ORDER BY mensalidades.id DESC
    """)
    mensalidades = cursor.fetchall()
    
    conexao.close()
    
    return render_template("mensalidades.html", alunos=alunos, mensalidades=mensalidades)

@app.route("/cadastrar_mensalidades", methods=["POST"])
def cadastrar_mensalidade():
    if "usuario" not in session:
        return redirect("/")
    
    aluno_id = request.form["aluno_id"]
    mes = request.form["mes"]
    valor = request.form["valor"]
    data_vencimento = request.form["data_vencimento"]
    status = request.form["status"]
    
    conexao = conectar()
    cursor = conexao.cursor()
    
    cursor.execute("""
       INSERT INTO mensalidades (aluno_id, mes, valor, data_vencimento, status)
       VALUES (?, ?, ?, ?, ?) 
    """, (aluno_id, mes, valor, data_vencimento, status))
    
    conexao.commit()
    conexao.close()
    
    return redirect("/mensalidades")

@app.route("/relatorio_mensalidades_excel")
def relatorio_mensalidades_excel():
    if "usuario" not in session:
        return redirect("/")
    
    conexao = conectar()
    cursor = conexao.cursor()
    
    cursor.execute("""
        SELECT alunos.nome, mensalidades.mes, mensalidades.valor,
               mensalidades.data_vencimento, mensalidades.status
        FROM mensalidades
        JOIN alunos ON alunos.id = mensalidades.aluno_id
        ORDER BY alunos.nome
    """)
    
    dados = cursor.fetchall()
    conexao.close()
    
    wb = Workbook()
    ws = wb.active
    ws.title = "Mensalidades"
    
    ws.append(["Aluno", "Mês", "Valor", "Vencimento", "Status"])
    
    for linha in dados:
        ws.append(linha)
        
    caminho = "relatorio_mensalidades.xlsx"
    wb.save(caminho)
    
    return send_file(caminho, as_attachment=True)

@app.route("/relatorio_mensalidades_pdf")
def relatorio_mensalidades_pdf():
    if "usuario" not in session:
        return redirect("/")
    
    conexao = conectar()
    cursor = conexao.cursor()
    
    cursor.execute("""
        SELECT alunos.nome, mensalidades.mes, mensalidades.valor,
               mensalidades.data_vencimento, mensalidades.status
        FROM mensalidades
        JOIN alunos ON alunos.id = mensalidades.aluno_id
        ORDER BY alunos.nome
    """)
    
    dados = cursor.fetchall()
    conexao.close()
    
    caminho = "relatorio_mensalidades.pdf"
    
    pdf = canvas.Canvas(caminho, pagesize=A4)
    largura, altura = A4
    
    y = altura - 50
    
    pdf.setFont("Helvetica-Bold", 16)
    pdf.drawString(50, y, "Relatório de Mensalidades")
    
    y -= 40
    pdf.setFont("Helvetica-Bold", 10)
    pdf.drawString(50, y, "Aluno")
    pdf.drawString(200, y, "Mês")
    pdf.drawString(300, y, "Valor")
    pdf.drawString(380, y, "Vencimento")
    pdf.drawString(480, y, "Status")
    
    y -=20
    pdf.setFont("Helvetica", 9)
    
    for aluno, mes, valor, vencimento, status in dados:
        if y < 50:
            pdf.showPage()
            y = altura - 50
            pdf.setFont("Helvetica", 9)
            
        pdf.drawString(50, y, str(aluno))
        pdf.drawString(200, y, str(mes))
        pdf.drawString(300, y, f"R$ {valor:.2f}")
        pdf.drawString(380, y, str(vencimento))
        pdf.drawString(480, y, str(status))
        
        y -= 18
        
    pdf.save()
    
    return send_file(caminho, as_attachment=True)

@app.route("/sair")
def sair():
    session.clear()
    return redirect("/")

if __name__ == "__main__":
    criar_banco()
    app.run(host="0.0.0.0", port=5000)