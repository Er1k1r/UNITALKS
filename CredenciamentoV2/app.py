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
import logging
import os
import secrets
import smtplib
import sqlite3
import uuid
from datetime import datetime, timedelta
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from functools import wraps

import qrcode
from flask import (Flask, g, jsonify, redirect, render_template, request,
                    send_file, session, url_for, flash)

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger("unitalks")

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(BASE_DIR, "evento.db")
QR_DIR = os.path.join(BASE_DIR, "static", "qrcodes")

app = Flask(__name__)
# Cache de 7 dias para arquivos estáticos (CSS, imagens) — o navegador do
# recepcionista/participante não precisa baixar tudo de novo a cada visita.
# Como esses arquivos raramente mudam durante o evento, isso reduz bastante
# o tempo de carregamento nas visitas seguintes.
app.config["SEND_FILE_MAX_AGE_DEFAULT"] = 60 * 60 * 24 * 7


@app.context_processor
def util_static_versionado():
    """Acrescenta ?v=<data de modificação> automaticamente em toda URL de
    /static usada nos templates. Assim, o navegador pode cachear os
    arquivos por 7 dias com segurança: se um arquivo for atualizado (ex.:
    trocar uma imagem), o carimbo muda sozinho e o navegador busca a
    versão nova, sem precisar esperar o cache expirar."""
    def static_v(filename):
        caminho = os.path.join(app.static_folder, filename)
        try:
            versao = int(os.path.getmtime(caminho))
        except OSError:
            versao = 0
        return url_for("static", filename=filename, v=versao)
    return {"static_v": static_v}
# Em produção (Render), defina a variável de ambiente SECRET_KEY com um
# valor fixo e aleatório (veja o README). Se não for definida, o sistema
# gera uma chave nova e aleatória a cada reinício do servidor — isso é
# mais seguro que um valor fixo no código, e tem como efeito colateral
# desejado: toda sessão antiga (de testes, de qualquer navegador) é
# automaticamente invalidada sempre que o serviço reinicia, então
# ninguém consegue abrir o site já "logado" sem querer.
app.secret_key = os.environ.get("SECRET_KEY") or secrets.token_hex(32)

# Duas senhas de equipe:
#  - ADMIN: acesso total (check-in/check-out + painel + exportação).
#  - NORMAL: acesso só ao check-in/check-out (scanner). Sem painel.
# Em produção (Render), defina as variáveis de ambiente ADMIN_PASSWORD e
# ACCESS_PASSWORD no lugar dos valores padrão abaixo.
ADMIN_PASSWORD = os.environ.get("ADMIN_PASSWORD", "ADMINNEGOCIOS2026")
ACCESS_PASSWORD = os.environ.get("ACCESS_PASSWORD", "CRED2026")

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
        # timeout=10: se o banco estiver momentaneamente ocupado (duas
        # leituras de QR Code quase simultâneas), a conexão espera até 10s
        # em vez de falhar na hora com "database is locked".
        g.db = sqlite3.connect(DB_PATH, timeout=10)
        g.db.row_factory = sqlite3.Row
        g.db.execute("PRAGMA foreign_keys = ON")
        # WAL permite leituras e escritas simultâneas sem travar o banco —
        # importante porque o painel pode estar sendo lido (admin) ao mesmo
        # tempo em que o scanner está gravando check-ins.
        g.db.execute("PRAGMA journal_mode = WAL")
        g.db.execute("PRAGMA busy_timeout = 10000")
    return g.db


@app.teardown_appcontext
def close_db(exception=None):
    db = g.pop("db", None)
    if db is not None:
        db.close()


# ---------------------------------------------------------------------------
# Backup do banco de dados
# ---------------------------------------------------------------------------
# O SQLite roda em disco temporário nos planos gratuitos de hospedagem: se o
# serviço reiniciar, os dados podem se perder. Duas camadas de proteção:
#  1) Botão "Baixar backup agora" no painel (sob demanda, sempre disponível).
#  2) Envio automático por e-mail em intervalos regulares, se as variáveis
#     de ambiente de SMTP estiverem configuradas (ver README).
def gerar_backup_db() -> bytes:
    """Gera uma cópia segura e consistente do banco, mesmo com o site em uso
    (usa a API oficial de backup do SQLite, que não corrompe o arquivo
    original nem trava as escritas em andamento)."""
    tmp_path = os.path.join(BASE_DIR, f"_backup_tmp_{uuid.uuid4().hex}.db")
    origem = sqlite3.connect(DB_PATH)
    destino = sqlite3.connect(tmp_path)
    with destino:
        origem.backup(destino)
    destino.close()
    origem.close()
    try:
        with open(tmp_path, "rb") as f:
            dados = f.read()
    finally:
        os.remove(tmp_path)
    return dados


def enviar_backup_por_email():
    """Job periódico: gera o backup e envia por e-mail, se configurado via
    variáveis de ambiente. Não configurado = não faz nada (sem erro).
    Retorna (sucesso: bool, mensagem: str) para poder ser usado tanto pelo
    agendador automático quanto pelo botão de teste manual no painel."""
    destino = os.environ.get("BACKUP_EMAIL_TO")
    smtp_host = os.environ.get("BACKUP_SMTP_HOST")
    smtp_user = os.environ.get("BACKUP_SMTP_USER")
    smtp_senha = os.environ.get("BACKUP_SMTP_PASSWORD")

    if not (destino and smtp_host and smtp_user and smtp_senha):
        return False, "Configure BACKUP_EMAIL_TO, BACKUP_SMTP_HOST, BACKUP_SMTP_USER e BACKUP_SMTP_PASSWORD nas variáveis de ambiente."

    try:
        smtp_port = int(os.environ.get("BACKUP_SMTP_PORT", "587"))
        agora = datetime.now().strftime("%d/%m/%Y %H:%M")
        dados = gerar_backup_db()

        msg = MIMEMultipart()
        msg["Subject"] = f"[UniTalks] Backup automático do evento — {agora}"
        msg["From"] = smtp_user
        msg["To"] = destino
        msg.attach(MIMEText(
            f"Backup automático gerado em {agora}.\n\n"
            "Este é um e-mail automático do sistema de credenciamento UniTalks. "
            "Guarde este arquivo .db como cópia de segurança dos inscritos e check-ins.",
            "plain",
        ))
        anexo = MIMEApplication(dados, Name="backup_unitalks.db")
        anexo["Content-Disposition"] = 'attachment; filename="backup_unitalks.db"'
        msg.attach(anexo)

        with smtplib.SMTP(smtp_host, smtp_port, timeout=20) as smtp:
            smtp.starttls()
            smtp.login(smtp_user, smtp_senha)
            smtp.sendmail(smtp_user, [destino], msg.as_string())
        logger.info("Backup automático enviado por e-mail para %s", destino)
        return True, f"Backup enviado com sucesso para {destino}."
    except Exception as e:
        logger.exception("Falha ao enviar backup automático por e-mail")
        return False, f"Falha ao enviar: {e}"


def iniciar_agendador_backup():
    """Liga o agendador de backup em segundo plano, se a biblioteca
    APScheduler estiver instalada e houver e-mail configurado. Roda dentro
    do próprio processo web — por isso o Procfile usa 1 worker (com
    threads), para não disparar o job várias vezes em paralelo."""
    if not os.environ.get("BACKUP_EMAIL_TO"):
        return
    try:
        from apscheduler.schedulers.background import BackgroundScheduler
    except ImportError:
        logger.warning(
            "BACKUP_EMAIL_TO configurado, mas a biblioteca APScheduler não "
            "está instalada. Adicione 'APScheduler' ao requirements.txt."
        )
        return

    intervalo_min = int(os.environ.get("BACKUP_INTERVAL_MINUTES", "120"))
    scheduler = BackgroundScheduler(daemon=True, timezone="America/Sao_Paulo")
    scheduler.add_job(enviar_backup_por_email, "interval", minutes=intervalo_min,
                       id="backup_email", replace_existing=True)
    scheduler.start()
    logger.info("Agendador de backup automático ligado (a cada %s min).", intervalo_min)


# Liga o agendador uma vez, quando o módulo é importado (tanto pelo
# `python app.py` local quanto pelo gunicorn em produção).
iniciar_agendador_backup()


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
            curso TEXT,
            periodo TEXT,
            rgm TEXT,
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
        "ALTER TABLE participantes ADD COLUMN curso TEXT",
        "ALTER TABLE participantes ADD COLUMN periodo TEXT",
        "ALTER TABLE participantes ADD COLUMN rgm TEXT",
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
    """Gera a credencial em PDF: QR Code + nome + número de inscrição.

    Visual alinhado à identidade do site (gradiente roxo-escuro → magenta,
    selo em degradê laranja/rosa para o número) e ao motivo de "ingresso":
    um canhoto perfurado separa a credencial das instruções, como em um
    ticket de verdade.
    """
    from reportlab.lib.pagesizes import A5
    from reportlab.pdfgen import canvas
    from reportlab.lib.units import mm

    buffer = io.BytesIO()
    largura, altura = A5
    c = canvas.Canvas(buffer, pagesize=A5)

    # Fundo levemente creme (mais quente que branco puro), pra combinar
    # com o restante do material do evento e não parecer um PDF genérico.
    c.setFillColorRGB(0.985, 0.98, 0.975)
    c.rect(0, 0, largura, altura, fill=1, stroke=0)

    # --- Cabeçalho em degradê (roxo-escuro → magenta), como no site ---
    faixa_altura = 44 * mm
    faixa_y0 = altura - faixa_altura
    n_faixas = 48
    cor_inicio = (0x19 / 255, 0x1a / 255, 0x32 / 255)   # #191a32 (bg-deep)
    cor_fim = (0x8a / 255, 0x4f / 255, 0x6b / 255)       # #8a4f6b (bg-dawn)
    for i in range(n_faixas):
        t = i / (n_faixas - 1)
        r = cor_inicio[0] + (cor_fim[0] - cor_inicio[0]) * t
        g = cor_inicio[1] + (cor_fim[1] - cor_inicio[1]) * t
        b = cor_inicio[2] + (cor_fim[2] - cor_inicio[2]) * t
        c.setFillColorRGB(r, g, b)
        h = faixa_altura / n_faixas
        c.rect(0, faixa_y0 + i * h, largura, h + 0.5, fill=1, stroke=0)

    # Logo (ícone) + nome do evento, lado a lado, centralizados
    logo_path = os.path.join(BASE_DIR, "static", "img", "icone_unitalks.png")
    titulo = "UNITALKS"
    c.setFont("Helvetica-Bold", 20)
    largura_titulo = c.stringWidth(titulo, "Helvetica-Bold", 20)
    logo_tam = 9 * mm
    espaco = 3 * mm
    bloco_largura = logo_tam + espaco + largura_titulo
    bloco_x = (largura - bloco_largura) / 2
    centro_y = altura - 17 * mm

    if os.path.exists(logo_path):
        c.drawImage(
            logo_path, bloco_x, centro_y - logo_tam / 2.4, width=logo_tam, height=logo_tam,
            mask="auto", preserveAspectRatio=True,
        )
    c.setFillColorRGB(1, 1, 1)
    c.drawString(bloco_x + logo_tam + espaco, centro_y - 6, titulo)

    c.setFont("Helvetica", 9.5)
    c.setFillColorRGB(0.89, 0.867, 0.937)  # tom claro seguro (mesmo do site)
    c.drawCentredString(largura / 2, altura - 26 * mm, "Um negócio por trás dos negócios")
    c.setFont("Helvetica", 7.5)
    c.setFillColorRGB(0.8, 0.77, 0.85)
    c.drawCentredString(largura / 2, altura - 31 * mm, "Evento licenciado pelo Centro Universitário UNIPÊ")

    # --- Corpo: QR Code com moldura sutil ---
    qr_path = os.path.join(QR_DIR, f"{participante['token']}.png")
    qr_tamanho = 64 * mm
    qr_x = (largura - qr_tamanho) / 2
    qr_y = faixa_y0 - 12 * mm - qr_tamanho

    moldura_pad = 4 * mm
    c.setFillColorRGB(1, 1, 1)
    c.setStrokeColorRGB(0.87, 0.85, 0.9)
    c.roundRect(
        qr_x - moldura_pad, qr_y - moldura_pad,
        qr_tamanho + 2 * moldura_pad, qr_tamanho + 2 * moldura_pad,
        6, fill=1, stroke=1,
    )
    c.drawImage(qr_path, qr_x, qr_y, width=qr_tamanho, height=qr_tamanho)

    # --- Nome do participante ---
    y_cursor = qr_y - moldura_pad - 11 * mm
    c.setFillColorRGB(0.09, 0.04, 0.18)
    c.setFont("Helvetica-Bold", 16)
    c.drawCentredString(largura / 2, y_cursor, participante["nome"])

    # --- Tipo + instituição (se preenchidos) ---
    linha_extra = ""
    if participante["tipo"] == "aluno":
        linha_extra = "Aluno(a)"
    elif participante["tipo"] == "participante":
        linha_extra = "Participante"
    if participante["instituicao"]:
        linha_extra = f"{linha_extra} — {participante['instituicao']}" if linha_extra else participante["instituicao"]

    if linha_extra:
        y_cursor -= 7 * mm
        c.setFont("Helvetica", 10)
        c.setFillColorRGB(0.45, 0.42, 0.52)
        c.drawCentredString(largura / 2, y_cursor, linha_extra)

    # --- Selo do número de inscrição, em degradê laranja → rosa ---
    y_cursor -= 13 * mm
    texto_num = f"Inscrição nº {participante['numero_inscricao']}"
    c.setFont("Helvetica-Bold", 13)
    largura_texto = c.stringWidth(texto_num, "Helvetica-Bold", 13)
    selo_largura = largura_texto + 16 * mm
    selo_altura = 10 * mm
    selo_x = (largura - selo_largura) / 2
    selo_y = y_cursor - selo_altura / 2.6

    n_fatias = 40
    cor_a = (0xff / 255, 0xc1 / 255, 0x57 / 255)  # dourado
    cor_b = (0xff / 255, 0x6b / 255, 0x4a / 255)  # coral
    c.saveState()
    caminho_clip = c.beginPath()
    caminho_clip.roundRect(selo_x, selo_y, selo_largura, selo_altura, selo_altura / 2)
    c.clipPath(caminho_clip, stroke=0)
    for i in range(n_fatias):
        t = i / (n_fatias - 1)
        r = cor_a[0] + (cor_b[0] - cor_a[0]) * t
        g = cor_a[1] + (cor_b[1] - cor_a[1]) * t
        b = cor_a[2] + (cor_b[2] - cor_a[2]) * t
        c.setFillColorRGB(r, g, b)
        w = selo_largura / n_fatias
        c.rect(selo_x + i * w, selo_y, w + 0.5, selo_altura, fill=1, stroke=0)
    c.restoreState()

    c.setFillColorRGB(0x1b / 255, 0x15 / 255, 0x33 / 255)  # tinta escura (contraste sobre dourado/coral)
    c.setFont("Helvetica-Bold", 13)
    c.drawCentredString(largura / 2, selo_y + selo_altura / 2 - 4.6, texto_num)

    # --- Canhoto perfurado, como em um ingresso de verdade ---
    linha_perfuracao_y = 26 * mm
    c.setFillColorRGB(0.985, 0.98, 0.975)
    c.circle(0, linha_perfuracao_y, 4 * mm, fill=1, stroke=0)
    c.circle(largura, linha_perfuracao_y, 4 * mm, fill=1, stroke=0)
    c.setDash(3, 3)
    c.setStrokeColorRGB(0.75, 0.72, 0.8)
    c.setLineWidth(1)
    c.line(6 * mm, linha_perfuracao_y, largura - 6 * mm, linha_perfuracao_y)
    c.setDash()

    # --- Instruções no canhoto ---
    c.setFont("Helvetica", 8.5)
    c.setFillColorRGB(0.45, 0.42, 0.52)
    c.drawCentredString(
        largura / 2, linha_perfuracao_y - 8 * mm,
        "Apresente este QR Code (ou informe o número acima) no credenciamento.",
    )
    c.setFont("Helvetica", 7)
    c.setFillColorRGB(0.6, 0.57, 0.68)
    c.drawCentredString(
        largura / 2, linha_perfuracao_y - 13 * mm,
        "16 e 17 de setembro de 2026 · Avenida Diógenes Chianca, João Pessoa - PB",
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


def proximo_numero_inscricao(db) -> int:
    """Calcula o próximo número de inscrição a ser usado, REAPROVEITANDO
    números liberados por exclusões. Ex.: se 202601 está em uso e 202602 foi
    excluído, o próximo inscrito recebe 202602 (não 202603) — a numeração
    nunca fica com "buracos" enquanto houver um número livre mais baixo."""
    usados = {
        row["numero_inscricao"]
        for row in db.execute(
            "SELECT numero_inscricao FROM participantes WHERE numero_inscricao IS NOT NULL"
        ).fetchall()
    }
    candidato = NUMERO_INSCRICAO_BASE + 1
    while candidato in usados:
        candidato += 1
    return candidato


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
        # Curso, período e RGM só fazem sentido pra quem é aluno — mesmo
        # que o campo tenha sido preenchido no navegador (ex.: JS
        # desabilitado), ignoramos esses valores se a pessoa marcou
        # "participante externo", pra manter o banco consistente.
        if tipo == "aluno":
            curso = request.form.get("curso", "").strip()
            periodo = request.form.get("periodo", "").strip()
            rgm = request.form.get("rgm", "").strip()
        else:
            curso = periodo = rgm = ""

        if not nome:
            flash("O nome é obrigatório.", "erro")
            return redirect(url_for("inscricao"))

        if tipo not in ("aluno", "participante"):
            flash("Selecione se você é aluno ou participante externo.", "erro")
            return redirect(url_for("inscricao"))

        token = uuid.uuid4().hex[:12]
        db = get_db()
        # Tenta algumas vezes: se duas pessoas se inscreverem no mesmíssimo
        # instante, as duas podem calcular o mesmo "próximo número
        # disponível" antes de qualquer uma delas salvar. A trava UNIQUE no
        # banco rejeita a segunda tentativa, e o loop recalcula e tenta de
        # novo com o número seguinte.
        for tentativa in range(5):
            numero_inscricao = proximo_numero_inscricao(db)
            try:
                db.execute(
                    "INSERT INTO participantes (nome, tipo, email, formacao, instituicao, cpf, curso, periodo, rgm, token, numero_inscricao, criado_em) "
                    "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
                    (nome, tipo, email, formacao, instituicao, cpf, curso, periodo, rgm, token, numero_inscricao, datetime.now().isoformat(timespec="seconds")),
                )
                db.commit()
                break
            except sqlite3.IntegrityError:
                db.rollback()
                if tentativa == 4:
                    flash("Não foi possível concluir a inscrição, tente novamente.", "erro")
                    return redirect(url_for("inscricao"))
                continue
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


@app.route("/api/stats")
@login_required
def api_stats():
    """Números rápidos para a barra de estatísticas do scanner: total de
    inscritos, total de check-ins já feitos (histórico) e % de presença
    (quantos inscritos estão fisicamente dentro de algum local agora)."""
    db = get_db()
    total_inscritos = db.execute("SELECT COUNT(*) AS n FROM participantes").fetchone()["n"]
    total_checkins = db.execute(
        "SELECT COUNT(*) AS n FROM eventos_acesso WHERE tipo = 'entrada'"
    ).fetchone()["n"]

    participantes = db.execute("SELECT id FROM participantes").fetchall()
    dentro_agora = 0
    for p in participantes:
        eventos = db.execute(
            "SELECT tipo, horario, local FROM eventos_acesso WHERE participante_id = ? ORDER BY horario ASC",
            (p["id"],),
        ).fetchall()
        visitas = montar_visitas(eventos)
        if any(v["em_andamento"] for v in visitas):
            dentro_agora += 1

    presenca_pct = round((dentro_agora / total_inscritos) * 100) if total_inscritos else 0

    return jsonify({
        "inscritos": total_inscritos,
        "checkins": total_checkins,
        "presenca_pct": presenca_pct,
    })


# ---------------------------------------------------------------------------
# Exportação
# ---------------------------------------------------------------------------
@app.route("/admin/backup")
@admin_required
def admin_backup():
    """Baixa uma cópia .db do banco inteiro, na hora, com segurança (mesmo
    com o site em uso). Use para guardar um backup manual quando quiser,
    além do backup automático por e-mail (se configurado)."""
    dados = gerar_backup_db()
    carimbo = datetime.now().strftime("%Y%m%d_%H%M%S")
    return send_file(
        io.BytesIO(dados),
        mimetype="application/octet-stream",
        as_attachment=True,
        download_name=f"backup_unitalks_{carimbo}.db",
    )


@app.route("/admin/participante/<int:participante_id>/excluir", methods=["POST"])
@admin_required
def admin_excluir_participante(participante_id):
    """Exclui um único inscrito e todo o seu histórico de check-in/check-out.
    Usado quando alguém se inscreveu por engano, duplicado, ou desistiu."""
    db = get_db()
    p = db.execute("SELECT nome FROM participantes WHERE id = ?", (participante_id,)).fetchone()
    if not p:
        flash("❌ Participante não encontrado (talvez já tenha sido excluído).", "erro")
        return redirect(url_for("painel"))

    db.execute("DELETE FROM eventos_acesso WHERE participante_id = ?", (participante_id,))
    db.execute("DELETE FROM participantes WHERE id = ?", (participante_id,))
    db.commit()

    logger.info("Participante excluído pelo admin: %s (id=%s)", p["nome"], participante_id)
    flash(f"🗑️ Inscrição de \"{p['nome']}\" excluída com sucesso.", "sucesso")
    return redirect(url_for("painel"))


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
            "Curso", "Período", "RGM", "Formação (externo)", "Instituição", "Inscrito em",
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
            p["curso"], p["periodo"], p["rgm"], p["formacao"], p["instituicao"], p["criado_em"],
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


@app.errorhandler(500)
def erro_500(e):
    logger.exception("Erro interno não tratado")
    if request.path.startswith("/api/"):
        return jsonify({"ok": False, "mensagem": "Erro interno. Tente novamente em instantes."}), 500
    return render_template("erro.html", titulo="Ops, algo deu errado",
                            mensagem="Ocorreu um erro interno. Tente novamente em instantes."), 500


@app.errorhandler(404)
def erro_404(e):
    if request.path.startswith("/api/"):
        return jsonify({"ok": False, "mensagem": "Recurso não encontrado."}), 404
    return render_template("erro.html", titulo="Página não encontrada",
                            mensagem="O endereço acessado não existe."), 404


if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))
