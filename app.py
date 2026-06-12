from flask import Flask, render_template, request, redirect, session
import sqlite3
from flask import send_file
from openpyxl import Workbook
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from datetime import datetime, timedelta
import os
import base64
from io import BytesIO
from reportlab.lib.utils import ImageReader

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
    novos_campos = [
        ("tipo_plano", "TEXT"),
        ("aulas_semana", "TEXT"),
        ("dias_horarios", "TEXT"),
        ("valor_mensal", "REAL"),
        ("dia_vencimento", "INTEGER"),
        ("data_inicio", "TEXT"),
        ("assinatura", "TEXT")
    ]
        
    for campo, tipo in novos_campos:
        try:
            cursor.execute(f"ALTER TABLE alunos ADD COLUMN {campo} {tipo}")
        except sqlite3.OperationalError:
            pass
    
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

from datetime import date
import calendar

def gerar_12_mensalidades(aluno_id, valor, dia_vencimento, data_inicio):
    conexao = conectar()
    cursor = conexao.cursor()
    
    ano, mes, _ = map(int, data_inicio.split("-"))
    
    for i in range(12):
        mes_atual = mes + i
        ano_atual = ano + (mes_atual - 1) // 12
        mes_corrigido = ((mes_atual - 1) % 12) + 1
        
        ultimo_dia = calendar.monthrange(ano_atual, mes_corrigido)[1]
        dia = min(int(dia_vencimento), ultimo_dia)
        
        vencimento = date(ano_atual, mes_corrigido, dia).strftime("%Y-%m-%d")
        mes_referencia = date(ano_atual, mes_corrigido, 1).strftime("%m/%Y")
        
        cursor.execute("""
            INSERT INTO mensalidades
            (aluno_id, mes, valor, data_vencimento, status)
            VALUES (?, ?, ?, ?, ?)
        """, (aluno_id, mes_referencia, valor, vencimento, "Pendente"))
        
    conexao.commit()
    conexao.close()

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

@app.route("/cadastrar_aluno", methods=["POST"])
def cadastrar_aluno():
    if "usuario" not in session:
        return redirect("/")

    nome = request.form["nome"]
    endereco = request.form.get("endereco", "")
    telefone = request.form.get("telefone", "")
    email = request.form.get("email", "")
    data_nascimento = request.form.get("data_nascimento", "")
    observacoes = request.form.get("observacoes", "")

    tipo_plano = request.form.get("tipo_plano", "")
    aulas_semana = request.form.get("aulas_semana", "")
    dias_horarios = request.form.get("dias_horarios", "")
    valor_mensal = float(request.form.get("valor_mensal") or 0)
    dia_vencimento = int(request.form.get("dia_vencimento") or 1)
    data_inicio = request.form.get("data_inicio", "")

    conexao = conectar()
    cursor = conexao.cursor()

    cursor.execute("""
        INSERT INTO alunos (
            nome, endereco, telefone, email, data_nascimento, observacoes,
            tipo_plano, aulas_semana, dias_horarios,
            valor_mensal, dia_vencimento, data_inicio
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, (
        nome, endereco, telefone, email, data_nascimento, observacoes,
        tipo_plano, aulas_semana, dias_horarios,
        valor_mensal, dia_vencimento, data_inicio
    ))

    aluno_id = cursor.lastrowid

    conexao.commit()
    conexao.close()

    if valor_mensal > 0 and data_inicio:
        gerar_12_mensalidades(aluno_id, valor_mensal, dia_vencimento, data_inicio)

    return redirect(f"/contrato/{aluno_id}")

@app.route("/editar_aluno/<int:aluno_id>")
def editar_aluno(aluno_id):
    if "usuario" not in session:
        return redirect("/")

    conexao = conectar()
    cursor = conexao.cursor()

    cursor.execute("""
        SELECT id, nome, endereco, telefone, email, data_nascimento, observacoes,
               tipo_plano, aulas_semana, dias_horarios,
               valor_mensal, dia_vencimento, data_inicio
        FROM alunos
        WHERE id = ?
    """, (aluno_id,))

    aluno = cursor.fetchone()
    conexao.close()

    if not aluno:
        return "Aluno não encontrado"

    return render_template("editar_aluno.html", aluno=aluno)


@app.route("/atualizar_aluno/<int:aluno_id>", methods=["POST"])
def atualizar_aluno(aluno_id):
    if "usuario" not in session:
        return redirect("/")

    nome = request.form["nome"]
    endereco = request.form.get("endereco", "")
    telefone = request.form.get("telefone", "")
    email = request.form.get("email", "")
    data_nascimento = request.form.get("data_nascimento", "")
    observacoes = request.form.get("observacoes", "")

    tipo_plano = request.form.get("tipo_plano", "")
    aulas_semana = request.form.get("aulas_semana", "")
    dias_horarios = request.form.get("dias_horarios", "")
    valor_mensal = float(request.form.get("valor_mensal") or 0)
    dia_vencimento = int(request.form.get("dia_vencimento") or 1)
    data_inicio = request.form.get("data_inicio", "")

    conexao = conectar()
    cursor = conexao.cursor()

    cursor.execute("""
        UPDATE alunos
        SET nome = ?,
            endereco = ?,
            telefone = ?,
            email = ?,
            data_nascimento = ?,
            observacoes = ?,
            tipo_plano = ?,
            aulas_semana = ?,
            dias_horarios = ?,
            valor_mensal = ?,
            dia_vencimento = ?,
            data_inicio = ?
        WHERE id = ?
    """, (
        nome, endereco, telefone, email, data_nascimento, observacoes,
        tipo_plano, aulas_semana, dias_horarios,
        valor_mensal, dia_vencimento, data_inicio,
        aluno_id
    ))

    conexao.commit()
    conexao.close()

    return redirect("/alunos")

@app.route("/contrato/<int:aluno_id>")
def contrato(aluno_id):
    if "usuario" not in session:
        return redirect("/")
    
    conexao = conectar()
    cursor = conexao.cursor()
    
    cursor.execute("""
        SELECT id, nome, endereco, telefone, email, data_nascimento, observacoes,
        tipo_plano, aulas_semana, dias_horarios,
        valor_mensal, dia_vencimento, data_inicio, assinatura
        FROM alunos
        WHERE id = ?
    """, (aluno_id,))
    
    aluno = cursor.fetchone()
    conexao.close()
    
    if not aluno:
        return "Aluno não encontrado"
    
    caminho = f"contrato_aluno_{aluno_id}.pdf"
    
    pdf = canvas.Canvas(caminho, pagesize=A4)
    largura, altura = A4
    
    y = altura - 50
    
    pdf.setFont("Helvetica-Bold", 16)
    pdf.drawCentredString(largura / 2, y, "CONTRATO DE PRESTAÇÃO DE SERVIÇOS")
    y -= 35
    
    pdf.setFont("Helvetica", 10)
    pdf.drawString(50, y, f"CONTRATANTE: {aluno[1]}")
    y -= 18
    pdf.drawString(50, y, f"Endereço: {aluno[2] or ''}")
    y -= 18
    pdf.drawString(50, y, f"Telefone: {aluno[3] or ''}")
    y -= 18
    pdf.drawString(50, y, f"E-mail: {aluno[4] or ''}")
    y -= 18
    pdf.drawString(50, y, f"Data de Nascimento: {aluno[5] or ''}")
    
    y -= 35
    
    pdf.setFont("Helvetica-Bold", 12)
    pdf.drawString(50, y, "DADOS DO PLANO")
    y -= 22

    pdf.setFont("Helvetica", 10)
    pdf.drawString(50, y, f"Tipo de plano: {aluno[7] or ''}")
    y -= 18
    pdf.drawString(50, y, f"Aulas por semana: {aluno[8] or ''}")
    y -= 18
    pdf.drawString(50, y, f"Dias e horários: {aluno[9] or ''}")
    y -= 18
    pdf.drawString(50, y, f"Valor mensal: $ {float(aluno[10] or 0):.2f}")
    y -= 18
    pdf.drawString(50, y, f"Dia do vencimento: {aluno[11] or ''}")
    y -= 18
    pdf.drawString(50, y, f"Data de início: {aluno[12] or ''}")
    y -= 18

    y -= 35
           
    texto = [
        "Pelo presente contrato, o CONTRATANTE declara estar ciente das condições",
        "referentes às aulas de Pilates contratadas, incluindo frequência semanal,",
        "dias e horários estabelecidos, valor mensal e data de vencimento.",
        "",
        "O pagamento deverá ser realizado mensalmente até o dia de vencimento informado.",
        "O não comparecimento às aulas não isenta o CONTRATANTE do pagamento da mensalidade.",
        "Salvo em situações previamente acordadas com a administração.",
        "",
        "O presente contrato terá validade de 12 meses a partir da data de início informada.",
        "As mensalidades referentes ao período contratado serão geradas automaticamente",
        "no sistema financeiro."
    ]
     
    pdf.setFont("Helvetica", 10)

    for linha in texto:
        pdf.drawString(50, y, linha)
        y -= 14

    y -= 80
    
    linha_y = y

    print("ASSINATURA:", aluno[13])

    if aluno[13]:
        try:
            assinatura_base64 = aluno[13].split(",")[1]
            assinatura_bytes = base64.b64decode(assinatura_base64)

            assinatura_img = ImageReader(
                BytesIO(assinatura_bytes)
            )

            pdf.drawImage(
                assinatura_img,
                70,
                linha_y - 5,
                width=140,
                height=40,
                mask="auto"
            )

        except Exception as erro:
            print("ERRO ASSINATURA:", erro)

    pdf.drawString(50, linha_y, "________________________________________")
    pdf.drawString(330, linha_y, "________________________________________")

    y += 20

    pdf.drawString(95, linha_y - 15, "Assinatura do Contratante")
    pdf.drawString(370, linha_y - 15, "Responsável pelo Estúdio")

    y -= 70

    pdf.drawString(50, y, "Data: ____/____/________")

    pdf.save()

    return send_file(
        caminho,
        mimetype="application/pdf",
        as_attachment=False
    )
    
@app.route("/assinar_contrato/<int:aluno_id>")
def assinar_contrato(aluno_id):
    if "usuario" not in session:
        return redirect("/")

    return render_template("assinar_contrato.html", aluno_id=aluno_id)


@app.route("/salvar_assinatura/<int:aluno_id>", methods=["POST"])
def salvar_assinatura(aluno_id):
    if "usuario" not in session:
        return redirect("/")

    assinatura = request.form.get("assinatura")

    conexao = conectar()
    cursor = conexao.cursor()

    cursor.execute("""
        UPDATE alunos
        SET assinatura = ?
        WHERE id = ?
    """, (assinatura, aluno_id))

    conexao.commit()
    conexao.close()

    return redirect(f"/contrato/{aluno_id}")

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
        ORDER BY agenda.data_aula, agenda.horario
    """)
    aulas_mes_lista = cursor.fetchall()
    eventos_calendario = []
    
    for aula in aulas_mes_lista:
        eventos_calendario.append({
            "title":f"{aula[1]} - {aula[3]}",
            "start": aula[2],
            "description": aula[4] or "",
            "status": aula[5]
        })

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
        mensalidades_pendentes=mensalidades_pendentes,
        eventos_calendario=eventos_calendario
    )

@app.route("/agenda")
def agenda():
    if "usuario" not in session:
        return redirect("/")

    aluno_filtro = request.args.get("aluno_id", "")
    mes_filtro = request.args.get("mes", datetime.now().strftime("%Y-%m"))

    conexao = conectar()
    cursor = conexao.cursor()

    cursor.execute("SELECT id, nome FROM alunos ORDER BY nome")
    alunos = cursor.fetchall()

    sql = """
        SELECT agenda.id, alunos.nome, agenda.data_aula, agenda.horario,
               agenda.tipo_aula, agenda.status, agenda.observacoes
        FROM agenda
        JOIN alunos ON alunos.id = agenda.aluno_id
        WHERE agenda.data_aula LIKE ?
    """

    parametros = [mes_filtro + "%"]

    if aluno_filtro:
        sql += " AND agenda.aluno_id = ?"
        parametros.append(aluno_filtro)

    sql += " ORDER BY agenda.data_aula, agenda.horario"

    cursor.execute(sql, parametros)
    aulas = cursor.fetchall()

    eventos_calendario = []

    for aula in aulas:
        eventos_calendario.append({
            "title": f"{aula[1]} - {aula[3]}",
            "start": aula[2],
            "status": aula[5]
        })

    conexao.close()

    return render_template(
        "agenda.html",
        alunos=alunos,
        aulas=aulas,
        eventos_calendario=eventos_calendario,
        aluno_filtro=aluno_filtro,
        mes_filtro=mes_filtro
    )

@app.route("/cadastrar_aula", methods=["POST"])
def cadastrar_aula():
    if "usuario" not in session:
        return redirect("/")

    aluno_id = request.form["aluno_id"]
    data_aula = request.form["data_aula"]
    horario = request.form["horario"]
    tipo_aula = request.form.get("tipo_aula", "")
    status = request.form.get("status", "Agendada")
    observacoes = request.form.get("observacoes", "")

    recorrente = request.form.get("recorrente", "nao")
    dias_semana = request.form.getlist("dias_semana")

    conexao = conectar()
    cursor = conexao.cursor()

    if recorrente == "sim" and dias_semana:
        data_inicio = datetime.strptime(data_aula, "%Y-%m-%d").date()
        data_fim = data_inicio + timedelta(days=365)

        data_atual = data_inicio

        while data_atual <= data_fim:
            if str(data_atual.weekday()) in dias_semana:
                cursor.execute("""
                    SELECT COUNT(*) FROM agenda
                    WHERE aluno_id = ?
                      AND data_aula = ?
                      AND horario = ?
                """, (aluno_id, data_atual.strftime("%Y-%m-%d"), horario))

                existe = cursor.fetchone()[0]

                if existe == 0:
                    cursor.execute("""
                        INSERT INTO agenda
                        (aluno_id, data_aula, horario, tipo_aula, status, observacoes)
                        VALUES (?, ?, ?, ?, ?, ?)
                    """, (
                        aluno_id,
                        data_atual.strftime("%Y-%m-%d"),
                        horario,
                        tipo_aula,
                        status,
                        observacoes
                    ))

            data_atual += timedelta(days=1)

    else:
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
               CASE
                  WHEN mensalidades.status != 'Pago'
                       AND mensalidades.data_vencimento < date('now')
                  THEN 'Atrasado'
                  ELSE mensalidades.status
                END as status
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

@app.route("/receber_mensalidades", methods=["GET", "POST"])
def receber_mensalidades():
    if "usuario" not in session:
        return redirect("/")

    if not somente_admin():
        return redirect("/dashboard")

    conexao = conectar()
    cursor = conexao.cursor()

    cursor.execute("SELECT id, nome FROM alunos ORDER BY nome")
    alunos = cursor.fetchall()

    mensalidades = []
    aluno_selecionado = None

    if request.method == "POST":
        aluno_selecionado = request.form.get("aluno_id")

        cursor.execute("""
            SELECT id, mes, valor, data_vencimento,
                   CASE
                       WHEN status != 'Pago'
                            AND data_vencimento < date('now')
                       THEN 'Atrasado'
                       ELSE status
                   END as status
            FROM mensalidades
            WHERE aluno_id = ?
              AND status != 'Pago'
            ORDER BY data_vencimento
        """, (aluno_selecionado,))

        mensalidades = cursor.fetchall()

    conexao.close()

    return render_template(
        "receber_mensalidades.html",
        alunos=alunos,
        mensalidades=mensalidades,
        aluno_selecionado=aluno_selecionado
    )


@app.route("/confirmar_recebimento/<int:mensalidade_id>")
def confirmar_recebimento(mensalidade_id):
    if "usuario" not in session:
        return redirect("/")

    if not somente_admin():
        return redirect("/dashboard")

    conexao = conectar()
    cursor = conexao.cursor()

    cursor.execute("""
        UPDATE mensalidades
        SET status = 'Pago'
        WHERE id = ?
    """, (mensalidade_id,))

    conexao.commit()
    conexao.close()

    return redirect("/receber_mensalidades")

@app.route("/whatsapp_contrato/<int:aluno_id>")
def whatsapp_contrato(aluno_id):
    if "usuario" not in session:
        return redirect("/")

    conexao = conectar()
    cursor = conexao.cursor()

    cursor.execute("""
        SELECT nome, telefone
        FROM alunos
        WHERE id = ?
    """, (aluno_id,))

    aluno = cursor.fetchone()
    conexao.close()

    if not aluno:
        return "Aluno não encontrado"

    nome = aluno[0]
    telefone = aluno[1] or ""

    telefone = telefone.replace(" ", "").replace("-", "").replace("(", "").replace(")", "")

    link_contrato = request.host_url.rstrip("/") + f"/contrato/{aluno_id}"

    mensagem = f"""Olá {nome}, segue o link do seu contrato de Pilates para visualização e assinatura:

{link_contrato}

Obrigado."""

    import urllib.parse
    mensagem_codificada = urllib.parse.quote(mensagem)

    url = f"https://wa.me/{telefone}?text={mensagem_codificada}"

    return redirect(url)

@app.route("/receber_mensalidade/<int:id>")
def receber_mensalidade(id):
    if "usuario" not in session:
        return redirect("/")

    if not somente_admin():
        return redirect("/dashboard")

    conexao = conectar()
    cursor = conexao.cursor()

    cursor.execute("""
        UPDATE mensalidades
        SET status = 'Pago'
        WHERE id = ?
    """, (id,))

    conexao.commit()
    conexao.close()

    return redirect("/mensalidades")

@app.route("/sair")
def sair():
    session.clear()
    return redirect("/")

if __name__ == "__main__":
    criar_banco()
    app.run(host="0.0.0.0", port=5000)