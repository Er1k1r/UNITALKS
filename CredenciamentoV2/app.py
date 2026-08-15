"""
Sistema de Gestão de Eventos - 100% Gratuito (Python/Flask)
=============================================================
Funcionalidades:
  - Inscrição pública com geração automática de QR Code (credencial)
  - Painel administrativo com lista de participantes e estatísticas
  - Check-in / Check-out via leitura de QR Code (câmera do navegador)
  - Check-in / Check-out manual (busca por nome/CPF), caso não haja QR
  - Registro do horário EXATO (data/hora) de cada entrada e saída
  - Exportação em CSV (abre direto no Excel/Google Sheets)

Como rodar localmente:
  1) pip install -r requirements.txt
  2) python app.py
  3) Acesse http://localhost:5000

Como rodar em produção (Render, etc.):
  gunicorn app:app
  (o banco é inicializado automaticamente na importação do módulo,
  então funciona tanto local quanto atrás do gunicorn)

Banco de dados: SQLite (arquivo evento.db, criado automaticamente).
Nenhum serviço externo pago é necessário.
"""

import csv
import io
import os
import sqlite3
import uuid
from datetime import datetime, timedelta
from functools import wraps

import qrcode
from flask import (Flask, g, jsonify, redirect, render_template, request,
                    send_file, session, url_for, flash)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(BASE_DIR, "evento.db")
QR_DIR = os.path.join(BASE_DIR, "static", "qrcodes")

app = Flask(__name__)
# Em produção (Render), defina a variável de ambiente SECRET_KEY com um
# valor aleatório. Localmente, o valor abaixo já funciona para testes.
app.secret_key = os.environ.get("SECRET_KEY", "troque-esta-chave-em-producao")

# Duas senhas de equipe:
#  - ADMIN: acesso total (check-in/check-out + painel + exportação).
#  - NORMAL: acesso só ao check-in/check-out (scanner). Sem painel.
# Em produção (Render), defina as variáveis de ambiente ADMIN_PASSWORD e
# ACCESS_PASSWORD no lugar dos valores padrão abaixo.
ADMIN_PASSWORD = os.environ.get("ADMIN_PASSWORD", "ADMUNiTALKS2026")
ACCESS_PASSWORD = os.environ.get("ACCESS_PASSWORD", "unitalks2026")

# Cada dispositivo/navegador que fizer login fica com sua própria sessão,
# válida por várias horas — assim, várias pessoas da equipe podem estar
# logadas ao mesmo tempo em celulares diferentes, sem se atrapalharem.
app.permanent_session_lifetime = timedelta(hours=18)

# Base para o número de inscrição: o primeiro inscrito vira 202601, o
# segundo 202602, e assim por diante (202600 + id do participante).
NUMERO_INSCRICAO_BASE = 202600

# Locais de credenciamento do evento. Dia 1: só o Auditório Central (acesso
# obrigatório, com apenas 1 check-in e 1 check-out). Dia 2: 3 salas
# simultâneas, entre as quais o participante pode transitar livremente.
LOCAIS = {
    "auditorio": {"label": "Auditório Central", "dia": 1, "unico": True},
    "sala_1": {"label": "Sala 1", "dia": 2, "unico": False},
    "sala_2": {"label": "Sala 2", "dia": 2, "unico": False},
    "sala_3": {"label": "Sala 3", "dia": 2, "unico": False},
}

os.makedirs(QR_DIR, exist_ok=True)


def login_required(f):
    """Protege rotas que exigem login da equipe (scanner), aceitando tanto
    o acesso ADMIN quanto o acesso normal."""
    @wraps(f)
    def decorada(*args, **kwargs):
        if not session.get("autenticado"):
            if request.path.startswith("/api/"):
                return jsonify({"ok": False, "mensagem": "Sessão expirada. Faça login novamente."}), 401
            return redirect(url_for("login", proximo=request.path))
        return f(*args, **kwargs)
    return decorada


def admin_required(f):
    """Protege rotas restritas ao ADM (painel, exportação): exige login E
    nível admin. Quem logou com a senha normal é redirecionado ao scanner."""
    @wraps(f)
    def decorada(*args, **kwargs):
        if not session.get("autenticado"):
            if request.path.startswith("/api/"):
                return jsonify({"ok": False, "mensagem": "Sessão expirada. Faça login novamente."}), 401
            return redirect(url_for("login", proximo=request.path))
        if session.get("nivel") != "admin":
            flash("Essa área é restrita ao administrador.", "erro")
            return redirect(url_for("scanner"))
        return f(*args, **kwargs)
    return decorada


# ---------------------------------------------------------------------------
# Banco de dados
# ---------------------------------------------------------------------------
def get_db():
    if "db" not in g:
        g.db = sqlite3.connect(DB_PATH)
        g.db.row_factory = sqlite3.Row
        g.db.execute("PRAGMA foreign_keys = ON")
    return g.db


@app.teardown_appcontext
def close_db(exception=None):
    db = g.pop("db", None)
    if db is not None:
        db.close()


def init_db():
    """Cria as tabelas se ainda não existirem. Seguro para rodar toda vez."""
    db = sqlite3.connect(DB_PATH)
    db.executescript(
        """
        CREATE TABLE IF NOT EXISTS participantes (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nome TEXT NOT NULL,
            tipo TEXT,
            email TEXT,
            formacao TEXT,
            instituicao TEXT,
            cpf TEXT,
            empresa TEXT,
            numero_inscricao INTEGER UNIQUE,
            token TEXT UNIQUE NOT NULL,
            criado_em TEXT NOT NULL
        );

        CREATE TABLE IF NOT EXISTS eventos_acesso (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            participante_id INTEGER NOT NULL,
            tipo TEXT NOT NULL CHECK (tipo IN ('entrada', 'saida')),
            horario TEXT NOT NULL,
            local TEXT,
            FOREIGN KEY (participante_id) REFERENCES participantes (id)
        );
        """
    )
    db.commit()

    # Migração leve: adiciona colunas novas em bancos criados por versões
    # anteriores do sistema (não faz nada se a coluna já existir).
    colunas_novas = [
        "ALTER TABLE participantes ADD COLUMN numero_inscricao INTEGER",
        "ALTER TABLE participantes ADD COLUMN tipo TEXT",
        "ALTER TABLE participantes ADD COLUMN formacao TEXT",
        "ALTER TABLE participantes ADD COLUMN instituicao TEXT",
        "ALTER TABLE participantes ADD COLUMN cpf TEXT",
        "ALTER TABLE eventos_acesso ADD COLUMN local TEXT",
    ]
    for comando in colunas_novas:
        try:
            db.execute(comando)
            db.commit()
        except sqlite3.OperationalError:
            pass  # coluna já existe

    db.close()


# Inicializa o banco assim que o módulo é importado — funciona tanto
# rodando "python app.py" localmente quanto atrás do gunicorn no Render.
init_db()


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
def gerar_qrcode(token: str) -> str:
    """Gera a imagem do QR Code para o token e retorna o caminho do arquivo."""
    img = qrcode.make(token)
    path = os.path.join(QR_DIR, f"{token}.png")
    img.save(path)
    return path


def gerar_pdf_credencial(participante) -> bytes:
    """Gera a credencial em PDF: QR Code + nome + número de inscrição."""
    from reportlab.lib.pagesizes import A5
    from reportlab.pdfgen import canvas
    from reportlab.lib.units import mm

    buffer = io.BytesIO()
    largura, altura = A5
    c = canvas.Canvas(buffer, pagesize=A5)

    # Faixa de topo
    c.setFillColorRGB(0.06, 0.16, 0.29)  # navy
    c.rect(0, altura - 28 * mm, largura, 28 * mm, fill=1, stroke=0)
    c.setFillColorRGB(1, 1, 1)
    c.setFont("Helvetica-Bold", 16)
    c.drawCentredString(largura / 2, altura - 12 * mm, "UNI TALKS")
    c.setFont("Helvetica", 9)
    c.drawCentredString(largura / 2, altura - 19 * mm, "Um negócio por trás dos negócios")

    # QR Code
    qr_path = os.path.join(QR_DIR, f"{participante['token']}.png")
    qr_tamanho = 70 * mm
    qr_x = (largura - qr_tamanho) / 2
    qr_y = altura - 45 * mm - qr_tamanho
    c.drawImage(qr_path, qr_x, qr_y, width=qr_tamanho, height=qr_tamanho)

    # Nome
    c.setFillColorRGB(0.06, 0.16, 0.29)
    c.setFont("Helvetica-Bold", 15)
    c.drawCentredString(largura / 2, qr_y - 12 * mm, participante["nome"])

    # Tipo + Instituição (se preenchidos)
    linha_extra = ""
    if participante["tipo"] == "aluno":
        linha_extra = "Aluno(a)"
    elif participante["tipo"] == "participante":
        linha_extra = "Participante"
    if participante["instituicao"]:
        linha_extra = f"{linha_extra} — {participante['instituicao']}" if linha_extra else participante["instituicao"]

    proxima_linha_y = qr_y - 19 * mm
    if linha_extra:
        c.setFont("Helvetica", 10)
        c.setFillColorRGB(0.4, 0.44, 0.49)
        c.drawCentredString(largura / 2, proxima_linha_y, linha_extra)
        proxima_linha_y -= 9 * mm
    else:
        proxima_linha_y -= 3 * mm

    # Número de inscrição (destaque)
    c.setFont("Helvetica-Bold", 13)
    c.setFillColorRGB(0.88, 0.48, 0.10)  # laranja
    c.drawCentredString(
        largura / 2, proxima_linha_y, f"Inscrição nº {participante['numero_inscricao']}"
    )

    c.setFont("Helvetica", 8)
    c.setFillColorRGB(0.4, 0.44, 0.49)
    c.drawCentredString(
        largura / 2, 12 * mm,
        "Apresente este QR Code (ou informe o número acima) no credenciamento."
    )

    c.showPage()
    c.save()
    buffer.seek(0)
    return buffer.read()


def ultimo_status(db, participante_id: int, local: str = None):
    if local:
        row = db.execute(
            "SELECT tipo FROM eventos_acesso WHERE participante_id = ? AND local = ? "
            "ORDER BY horario DESC, id DESC LIMIT 1",
            (participante_id, local),
        ).fetchone()
    else:
        row = db.execute(
            "SELECT tipo FROM eventos_acesso WHERE participante_id = ? "
            "ORDER BY horario DESC, id DESC LIMIT 1",
            (participante_id,),
        ).fetchone()
    return row["tipo"] if row else None


def formatar_duracao(inicio_iso: str, fim_iso: str = None) -> str:
    """Formata o tempo de estadia entre uma entrada e uma saída (ou 'agora',
    se a pessoa ainda estiver dentro do local)."""
    inicio = datetime.fromisoformat(inicio_iso)
    fim = datetime.fromisoformat(fim_iso) if fim_iso else datetime.now()
    minutos_totais = max(0, int((fim - inicio).total_seconds() // 60))
    horas, minutos = divmod(minutos_totais, 60)
    if horas:
        return f"{horas}h{minutos:02d}min"
    return f"{minutos}min"


def montar_visitas(eventos):
    """Recebe os eventos (tipo, horario, local) de um participante, já
    ordenados do mais antigo para o mais novo, e agrupa em 'visitas' por
    local: cada visita é um par entrada→saída (ou uma entrada em aberto,
    se a pessoa ainda estiver lá dentro). Isso permite, por exemplo, ver
    que alguém esteve na Sala 1 das 19h05 às 19h40 e depois na Sala 2."""
    por_local = {}
    for ev in eventos:
        chave = ev["local"] or "outro"
        por_local.setdefault(chave, []).append(ev)

    visitas = []
    for local, evs in por_local.items():
        entrada_pendente = None
        for ev in sorted(evs, key=lambda e: e["horario"]):
            if ev["tipo"] == "entrada":
                entrada_pendente = ev["horario"]
            elif ev["tipo"] == "saida" and entrada_pendente:
                visitas.append({
                    "local": local,
                    "entrada": entrada_pendente,
                    "saida": ev["horario"],
                    "em_andamento": False,
                })
                entrada_pendente = None
        if entrada_pendente:
            visitas.append({
                "local": local,
                "entrada": entrada_pendente,
                "saida": None,
                "em_andamento": True,
            })

    for v in visitas:
        v["local_label"] = LOCAIS.get(v["local"], {}).get("label", "Não informado")
        v["duracao"] = formatar_duracao(v["entrada"], v["saida"])

    visitas.sort(key=lambda v: v["entrada"])
    return visitas


# ---------------------------------------------------------------------------
# Página inicial do evento
# ---------------------------------------------------------------------------
@app.route("/", methods=["GET"])
def home():
    return render_template("home.html")


# ---------------------------------------------------------------------------
# Rotas públicas: inscrição
# ---------------------------------------------------------------------------
@app.route("/inscricao", methods=["GET", "POST"])
def inscricao():
    if request.method == "POST":
        nome = request.form.get("nome", "").strip()
        tipo = request.form.get("tipo", "").strip()
        email = request.form.get("email", "").strip()
        formacao = request.form.get("formacao", "").strip()
        instituicao = request.form.get("instituicao", "").strip()
        cpf = request.form.get("cpf", "").strip()

        if not nome:
            flash("O nome é obrigatório.", "erro")
            return redirect(url_for("inscricao"))

        if tipo not in ("aluno", "participante"):
            flash("Selecione se você é aluno ou participante externo.", "erro")
            return redirect(url_for("inscricao"))

        token = uuid.uuid4().hex[:12]
        db = get_db()
        cursor = db.execute(
            "INSERT INTO participantes (nome, tipo, email, formacao, instituicao, cpf, token, criado_em) "
            "VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
            (nome, tipo, email, formacao, instituicao, cpf, token, datetime.now().isoformat(timespec="seconds")),
        )
        novo_id = cursor.lastrowid
        numero_inscricao = NUMERO_INSCRICAO_BASE + novo_id
        db.execute(
            "UPDATE participantes SET numero_inscricao = ? WHERE id = ?",
            (numero_inscricao, novo_id),
        )
        db.commit()
        gerar_qrcode(token)
        return redirect(url_for("confirmacao", token=token))

    return render_template("inscricao.html")


@app.route("/confirmacao/<token>")
def confirmacao(token):
    db = get_db()
    participante = db.execute(
        "SELECT * FROM participantes WHERE token = ?", (token,)
    ).fetchone()
    if not participante:
        flash("Inscrição não encontrada.", "erro")
        return redirect(url_for("inscricao"))
    return render_template("confirmacao.html", p=participante)


@app.route("/credencial/<token>.pdf")
def credencial_pdf(token):
    db = get_db()
    participante = db.execute(
        "SELECT * FROM participantes WHERE token = ?", (token,)
    ).fetchone()
    if not participante:
        flash("Inscrição não encontrada.", "erro")
        return redirect(url_for("inscricao"))

    pdf_bytes = gerar_pdf_credencial(participante)
    return send_file(
        io.BytesIO(pdf_bytes),
        mimetype="application/pdf",
        as_attachment=False,
        download_name=f"credencial_{participante['numero_inscricao']}.pdf",
    )


# ---------------------------------------------------------------------------
# Login da equipe (protege scanner, painel e exportação)
# ---------------------------------------------------------------------------
@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        senha = request.form.get("senha", "")
        proximo = request.form.get("proximo") or url_for("scanner")

        if senha == ADMIN_PASSWORD:
            session.permanent = True
            session["autenticado"] = True
            session["nivel"] = "admin"
            # Se o admin chegou aqui vindo de um link genérico, manda ele
            # direto pro painel; se veio tentando abrir uma página
            # específica (ex: /exportar), respeita esse destino.
            if proximo in (url_for("scanner"), url_for("login")):
                proximo = url_for("painel")
            return redirect(proximo)

        if senha == ACCESS_PASSWORD:
            session.permanent = True
            session["autenticado"] = True
            session["nivel"] = "equipe"
            return redirect(proximo)

        flash("Senha incorreta.", "erro")
        return redirect(url_for("login", proximo=proximo))

    proximo = request.args.get("proximo", url_for("scanner"))
    return render_template("login.html", proximo=proximo)


@app.route("/logout")
def logout():
    session.pop("autenticado", None)
    session.pop("nivel", None)
    return redirect(url_for("login"))


# ---------------------------------------------------------------------------
# Painel administrativo
# ---------------------------------------------------------------------------
@app.route("/painel")
@admin_required
def painel():
    db = get_db()
    participantes = db.execute(
        "SELECT * FROM participantes ORDER BY criado_em DESC"
    ).fetchall()

    lista = []
    dentro_agora = 0
    presentes_por_local = {chave: 0 for chave in LOCAIS}

    for p in participantes:
        eventos = db.execute(
            "SELECT tipo, horario, local FROM eventos_acesso WHERE participante_id = ? "
            "ORDER BY horario ASC",
            (p["id"],),
        ).fetchall()

        visitas = montar_visitas(eventos)

        visita_aberta = next((v for v in reversed(visitas) if v["em_andamento"]), None)
        if visita_aberta:
            dentro_agora += 1
            if visita_aberta["local"] in presentes_por_local:
                presentes_por_local[visita_aberta["local"]] += 1

        ultima_entrada = visitas[-1]["entrada"] if visitas else None
        ultima_saida = next((v["saida"] for v in reversed(visitas) if v["saida"]), None)

        lista.append(
            {
                "p": p,
                "visitas": visitas,
                "local_atual": visita_aberta["local_label"] if visita_aberta else None,
                "ultima_entrada": ultima_entrada,
                "ultima_saida": ultima_saida,
            }
        )

    total = len(participantes)
    total_alunos = sum(1 for p in participantes if p["tipo"] == "aluno")
    total_participantes = sum(1 for p in participantes if p["tipo"] == "participante")
    return render_template(
        "painel.html",
        lista=lista,
        total=total,
        dentro_agora=dentro_agora,
        presentes_por_local=presentes_por_local,
        locais=LOCAIS,
        total_alunos=total_alunos,
        total_participantes=total_participantes,
    )


@app.route("/scanner")
@login_required
def scanner():
    return render_template("scanner.html", locais=LOCAIS)


# ---------------------------------------------------------------------------
# API: check-in / check-out
# ---------------------------------------------------------------------------
@app.route("/api/checkin", methods=["POST"])
@login_required
def api_checkin():
    """
    Recebe um token (lido do QR Code) OU um número de inscrição (digitado
    manualmente) MAIS o local onde o recepcionista está posicionado
    (Auditório, Sala 1, Sala 2 ou Sala 3) e registra automaticamente ENTRADA
    ou SAÍDA naquele local específico (alternando conforme o último status
    da pessoa NAQUELE local), salvando o horário exato do servidor. Isso
    permite acompanhar quanto tempo cada participante ficou em cada sala.
    """
    data = request.get_json(silent=True) or request.form
    token = (data.get("token") or "").strip()
    numero_inscricao = (data.get("numero_inscricao") or "").strip()
    local = (data.get("local") or "").strip()

    if local not in LOCAIS:
        return jsonify({"ok": False, "mensagem": "Selecione o local (Auditório ou Sala) antes de registrar."}), 400

    if not token and not numero_inscricao:
        return jsonify({"ok": False, "mensagem": "Informe o QR Code ou o número de inscrição."}), 400

    db = get_db()
    if token:
        participante = db.execute(
            "SELECT * FROM participantes WHERE token = ?", (token,)
        ).fetchone()
    else:
        participante = db.execute(
            "SELECT * FROM participantes WHERE numero_inscricao = ?", (numero_inscricao,)
        ).fetchone()

    if not participante:
        return jsonify({"ok": False, "mensagem": "Participante não encontrado."}), 404

    info_local = LOCAIS[local]

    eventos_local = db.execute(
        "SELECT tipo, horario FROM eventos_acesso WHERE participante_id = ? AND local = ? "
        "ORDER BY horario ASC",
        (participante["id"], local),
    ).fetchall()

    # No Auditório (dia 1), só é permitido 1 check-in + 1 check-out no total.
    if info_local.get("unico") and len(eventos_local) >= 2:
        return jsonify({
            "ok": False,
            "mensagem": f"{participante['nome']} já concluiu check-in e check-out no {info_local['label']}.",
        }), 400

    status_atual_local = eventos_local[-1]["tipo"] if eventos_local else None
    novo_tipo = "saida" if status_atual_local == "entrada" else "entrada"
    agora = datetime.now().isoformat(timespec="seconds")

    # Aviso (não bloqueante): a pessoa ainda está com entrada em aberto em
    # OUTRO local — útil para o recepcionista perceber, por exemplo, que ela
    # esqueceu de dar saída da Sala 1 antes de entrar na Sala 2.
    aviso = None
    if novo_tipo == "entrada":
        outros_eventos = db.execute(
            "SELECT tipo, horario, local FROM eventos_acesso WHERE participante_id = ? AND local != ? "
            "ORDER BY horario ASC",
            (participante["id"], local),
        ).fetchall()
        aberto_em = {}
        for ev in outros_eventos:
            if ev["tipo"] == "entrada":
                aberto_em[ev["local"]] = True
            else:
                aberto_em[ev["local"]] = False
        local_aberto = next((loc for loc, aberto in aberto_em.items() if aberto), None)
        if local_aberto:
            nome_local_aberto = LOCAIS.get(local_aberto, {}).get("label", local_aberto)
            aviso = f"⚠️ {participante['nome']} ainda não deu saída em {nome_local_aberto}."

    db.execute(
        "INSERT INTO eventos_acesso (participante_id, tipo, horario, local) VALUES (?, ?, ?, ?)",
        (participante["id"], novo_tipo, agora, local),
    )
    db.commit()

    return jsonify(
        {
            "ok": True,
            "nome": participante["nome"],
            "tipo": novo_tipo,
            "horario": agora,
            "local": local,
            "local_label": info_local["label"],
            "aviso": aviso,
        }
    )


@app.route("/api/busca")
@login_required
def api_busca():
    """Busca participantes por nome/e-mail/número de inscrição (check-in manual, sem QR)."""
    termo = request.args.get("q", "").strip()
    db = get_db()
    if not termo:
        return jsonify([])
    rows = db.execute(
        "SELECT id, nome, email, numero_inscricao, token, cpf FROM participantes "
        "WHERE nome LIKE ? OR email LIKE ? OR CAST(numero_inscricao AS TEXT) LIKE ? OR cpf LIKE ? LIMIT 10",
        (f"%{termo}%", f"%{termo}%", f"%{termo}%", f"%{termo}%"),
    ).fetchall()
    return jsonify([dict(r) for r in rows])


# ---------------------------------------------------------------------------
# Exportação
# ---------------------------------------------------------------------------
@app.route("/exportar")
@admin_required
def exportar():
    db = get_db()
    participantes = db.execute(
        "SELECT * FROM participantes ORDER BY criado_em ASC"
    ).fetchall()

    output = io.StringIO()
    writer = csv.writer(output, delimiter=";")
    writer.writerow(
        [
            "Nº Inscrição", "Nome", "Tipo (aluno/participante)", "E-mail", "CPF",
            "Formação", "Instituição", "Inscrito em",
            "Local", "Data/Hora Entrada", "Data/Hora Saída", "Tempo de Estadia",
        ]
    )

    for p in participantes:
        eventos = db.execute(
            "SELECT tipo, horario, local FROM eventos_acesso WHERE participante_id = ? "
            "ORDER BY horario ASC",
            (p["id"],),
        ).fetchall()
        base = [
            p["numero_inscricao"], p["nome"], p["tipo"], p["email"], p["cpf"],
            p["formacao"], p["instituicao"], p["criado_em"],
        ]

        visitas = montar_visitas(eventos)
        if not visitas:
            writer.writerow(base + ["", "", "", ""])
            continue

        for v in visitas:
            writer.writerow(
                base
                + [
                    v["local_label"],
                    v["entrada"].replace("T", " "),
                    v["saida"].replace("T", " ") if v["saida"] else "Ainda dentro",
                    v["duracao"],
                ]
            )

    mem = io.BytesIO(output.getvalue().encode("utf-8-sig"))
    mem.seek(0)
    return send_file(
        mem,
        mimetype="text/csv",
        as_attachment=True,
        download_name="relatorio_evento.csv",
    )


if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))
