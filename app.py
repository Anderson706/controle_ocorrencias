import os
import zipfile
from io import BytesIO
from datetime import datetime
from functools import wraps
from uuid import uuid4

from flask import (
    Flask, render_template, request, redirect, url_for,
    flash, session, send_from_directory, send_file, current_app
)

from flask_sqlalchemy import SQLAlchemy
from werkzeug.security import generate_password_hash, check_password_hash
from werkzeug.utils import secure_filename

from docx import Document
from docx.shared import Cm, Inches, Pt, RGBColor
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer,
    Image, Table, TableStyle
)

BASE_DIR = os.path.abspath(os.path.dirname(__file__))
DB_FOLDER = os.path.join(BASE_DIR, "database")
UPLOAD_FOLDER = os.path.join(BASE_DIR, "static", "uploads")
DB_PATH = os.path.join(DB_FOLDER, "controle_ocorrencia.db")

os.makedirs(DB_FOLDER, exist_ok=True)
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

app = Flask(__name__)
app.config["SECRET_KEY"] = "controle-ocorrencia-executivo"
app.config["SQLALCHEMY_DATABASE_URI"] = f"sqlite:///{DB_PATH}"
app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False
app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER
app.config["MAX_CONTENT_LENGTH"] = 10 * 1024 * 1024

db = SQLAlchemy(app)

ALLOWED_EXTENSIONS = {"png", "jpg", "jpeg", "webp"}
def allowed_file(filename):
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS

EXTENSOES_PERMITIDAS_IMAGEM = {"png", "jpg", "jpeg", "webp"}
EXTENSOES_PERMITIDAS_POST = {"png", "jpg", "jpeg", "webp", "pdf", "doc", "docx", "xlsx"}


# =========================
# MODELS
# =========================
class Usuario(db.Model):
    __tablename__ = "usuarios"

    id = db.Column(db.Integer, primary_key=True)
    nome = db.Column(db.String(120), nullable=False)
    username = db.Column(db.String(80), unique=True, nullable=False)
    senha_hash = db.Column(db.String(255), nullable=False)
    perfil = db.Column(db.String(30), nullable=False, default="OPERACIONAL")
    ativo = db.Column(db.Boolean, default=True)
    criado_em = db.Column(db.DateTime, default=datetime.utcnow)

class AnaliseInvestigativa(db.Model):
    __tablename__ = "analises_investigativas"

    id = db.Column(db.Integer, primary_key=True)

    id_relatorio = db.Column(db.Integer, nullable=False)
    empresa = db.Column(db.String(120), nullable=False)
    unidade = db.Column(db.String(180), nullable=False)
    endereco = db.Column(db.String(255), nullable=False)
    classificacao = db.Column(db.String(80), nullable=False)
    produtos_segmento = db.Column(db.String(120), nullable=False)
    clientes = db.Column(db.String(120), nullable=False)

    objetivo = db.Column(db.Text, nullable=False)
    responsavel = db.Column(db.String(150), nullable=False)
    nome_funcao_data = db.Column(db.String(255), nullable=False)

    descricao_registro = db.Column(db.Text, nullable=False)
    conclusao = db.Column(db.Text, nullable=False)
    sugestao = db.Column(db.Text, nullable=True)

    criado_em = db.Column(db.DateTime, default=datetime.utcnow)

    imagens = db.relationship(
        "ImagemAnaliseInvestigativa",
        backref="analise",
        cascade="all, delete-orphan",
        lazy=True
    )


class ImagemAnaliseInvestigativa(db.Model):
    __tablename__ = "imagens_analises_investigativas"

    id = db.Column(db.Integer, primary_key=True)
    analise_id = db.Column(db.Integer, db.ForeignKey("analises_investigativas.id"), nullable=False)

    arquivo = db.Column(db.String(255), nullable=False)
    descricao = db.Column(db.Text, nullable=False)

    criado_em = db.Column(db.DateTime, default=datetime.utcnow)

class Ocorrencia(db.Model):
    __tablename__ = "ocorrencias"

    id = db.Column(db.Integer, primary_key=True)
    data_hora = db.Column(db.String(30), nullable=False)
    hora_ocorrencia = db.Column(db.String(10), nullable=False)
    natureza = db.Column(db.String(120), nullable=False)
    descricao = db.Column(db.Text, nullable=False)
    local = db.Column(db.String(120), nullable=False)
    operador = db.Column(db.String(120), nullable=False)
    gc = db.Column(db.String(120), nullable=False)
    envolvido = db.Column(db.String(120), nullable=True)
    foto = db.Column(db.String(255), nullable=True)

    prioridade = db.Column(db.String(20), nullable=False, default="MEDIA")
    status = db.Column(db.String(30), nullable=False, default="PENDENTE")

    situacao_investigacao = db.Column(db.String(30), nullable=True)
    conclusao_investigacao = db.Column(db.Text, nullable=True)
    anexo_post = db.Column(db.String(255), nullable=True)

    criado_por = db.Column(db.String(120), nullable=True)
    criado_em = db.Column(db.DateTime, default=datetime.utcnow)
    atualizado_em = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)
    responsavel_fechamento = db.Column(db.String(120), nullable=True)


# =========================
# HELPERS
# =========================
def arquivo_permitido(nome_arquivo, extensoes):
    return "." in nome_arquivo and nome_arquivo.rsplit(".", 1)[1].lower() in extensoes


def salvar_arquivo(arquivo, extensoes):
    if not arquivo or not arquivo.filename:
        return None

    if not arquivo_permitido(arquivo.filename, extensoes):
        return None

    nome_seguro = secure_filename(arquivo.filename)
    extensao = nome_seguro.rsplit(".", 1)[1].lower()
    nome_final = f"{uuid4().hex}.{extensao}"
    caminho = os.path.join(app.config["UPLOAD_FOLDER"], nome_final)
    arquivo.save(caminho)
    return nome_final


def set_senha(usuario, senha):
    usuario.senha_hash = generate_password_hash(senha)


def check_senha(usuario, senha):
    return check_password_hash(usuario.senha_hash, senha)


def remover_arquivo(nome_arquivo):
    if not nome_arquivo:
        return
    caminho = os.path.join(app.config["UPLOAD_FOLDER"], nome_arquivo)
    if os.path.exists(caminho):
        try:
            os.remove(caminho)
        except OSError:
            pass


def login_required(func):
    @wraps(func)
    def wrapper(*args, **kwargs):
        if not session.get("user_id"):
            flash("Faça login para continuar.", "warning")
            return redirect(url_for("login"))
        return func(*args, **kwargs)
    return wrapper


def perfil_required(*perfis):
    def decorator(func):
        @wraps(func)
        def wrapper(*args, **kwargs):
            if not session.get("user_id"):
                flash("Faça login para continuar.", "warning")
                return redirect(url_for("login"))

            perfil = (session.get("user_perfil") or "").upper()
            if perfil not in perfis:
                flash("Você não tem permissão para acessar esta área.", "danger")
                return redirect(url_for("ocorrencias"))
            return func(*args, **kwargs)
        return wrapper
    return decorator


def normalizar_status(valor):
    return (valor or "").strip().upper()


def normalizar_prioridade(valor):
    return (valor or "").strip().upper()


def aplicar_filtros(query):
    data_inicial = (request.args.get("data_inicial") or "").strip()
    data_final = (request.args.get("data_final") or "").strip()
    local = (request.args.get("local") or "").strip()
    natureza = (request.args.get("natureza") or "").strip()
    status = (request.args.get("status") or "").strip().upper()

    registros = query.all()
    filtrados = []

    for r in registros:
        ok = True
        data_base = (r.data_hora or "")[:10]

        if data_inicial:
            ok = ok and (data_base >= data_inicial)

        if data_final:
            ok = ok and (data_base <= data_final)

        if local:
            ok = ok and (local.lower() in (r.local or "").lower())

        if natureza:
            ok = ok and (natureza.lower() in (r.natureza or "").lower())

        if status:
            ok = ok and (normalizar_status(r.status) == status)

        if ok:
            filtrados.append(r)

    filtros = {
        "data_inicial": data_inicial,
        "data_final": data_final,
        "local": local,
        "natureza": natureza,
        "status": status,
    }
    return filtrados, filtros


# =========================
# AUTH
# =========================
@app.route("/")
def index():
    if session.get("user_id"):
        return redirect(url_for("ocorrencias"))
    return redirect(url_for("login"))



@app.route("/analise-investigativa", methods=["GET"])
@login_required
def analise_investigativa():
    return render_template("analise_investigativa.html")


def classify_line(valor):
    valor = (valor or "").strip()
    if not valor:
        return "-"
    return valor


def style_heading(paragraph, text, size=12, bold=True, align="left"):
    paragraph.text = ""

    if align == "center":
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    else:
        paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT

    run = paragraph.add_run(text)
    run.bold = bold
    run.font.name = "Arial"
    run.font.size = Pt(size)


def set_cell_background(cell, color_hex):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), color_hex)
    tc_pr.append(shd)


def add_grid_table(doc, data, col_widths_cm=None, header_bold_cols=None):
    table = doc.add_table(rows=len(data), cols=len(data[0]))
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    header_bold_cols = header_bold_cols or []

    for r_idx, row in enumerate(data):
        for c_idx, value in enumerate(row):
            cell = table.rows[r_idx].cells[c_idx]
            cell.text = str(value or "")

            if col_widths_cm and c_idx < len(col_widths_cm):
                cell.width = Cm(col_widths_cm[c_idx])

            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    run.font.name = "Arial"
                    run.font.size = Pt(9)

                    if c_idx in header_bold_cols:
                        run.bold = True

            if c_idx in header_bold_cols:
                set_cell_background(cell, "F2F2F2")

    return table


def convert_doc_to_pdf_bytes(doc, base_name="documento"):
    """
    PDF simples com o mesmo conteúdo textual.
    A conversão visual idêntica ao DOCX só acontece usando LibreOffice/Word.
    """
    buffer = BytesIO()

    pdf = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        rightMargin=36,
        leftMargin=36,
        topMargin=36,
        bottomMargin=36
    )

    styles = getSampleStyleSheet()
    story = []

    story.append(Paragraph("ANÁLISE INVESTIGATIVA", styles["Title"]))
    story.append(Spacer(1, 12))
    story.append(Paragraph("DHL - SECURITY", styles["Heading2"]))
    story.append(Spacer(1, 12))

    pdf.build(story)
    buffer.seek(0)

    return buffer

@app.route("/gerar_docx", methods=["POST"])
def gerar_analise_investigativa():
    f = request.form

    # ===== Coleta do form =====
    id_relatorio = (f.get("id_relatorio", "") or "").strip()

    dados_operacao = {
        "Empresa": f.get("empresa", ""),
        "Unidade": f.get("unidade", ""),
        "Endereço": f.get("endereco", ""),
        "Classificação do Site": classify_line(f.get("classificacao", "")),
        "Produtos Segmento (Setor)": f.get("produtos_segmento", ""),
        "Cliente(s)": f.get("clientes", ""),
    }

    dados_levantamento = {
        "Objetivo": f.get("objetivo", ""),
        "Responsável pelo Levantamento": f.get("responsavel", ""),
        "Nome / Função / Data": f.get("nome_funcao_data", ""),
    }

    descricao_registro = f.get("descricao_registro", "")
    conclusao = f.get("conclusao", "")
    sugestao = f.get("sugestao", "")

    files = request.files.getlist("imagens[]")
    descricoes = request.form.getlist("descricoes[]")
    if len(descricoes) < len(files):
        descricoes += [""] * (len(files) - len(descricoes))
    evidencias = [(file, desc) for file, desc in zip(files, descricoes) if file and file.filename]

    # ===== Montagem do DOCX =====
    doc = Document()

    # Margens
    for section in doc.sections:
        section.left_margin = Cm(2.0)
        section.right_margin = Cm(2.0)
        section.top_margin = Cm(2.0)
        section.bottom_margin = Cm(2.0)

    # ===== Cabeçalho com logo =====
    logo_path = os.path.join(app.root_path, "static", "logo.png")

    for section in doc.sections:
        header = section.header

        # limpar conteúdo anterior (sem usar p.clear(), que pode não existir)
        for p in header.paragraphs:
            p.text = ""

        # largura disponível (página - margens)
        available_width = section.page_width - section.left_margin - section.right_margin

        # ✅ AQUI É A CORREÇÃO: add_table no header exige width
        header_table = header.add_table(rows=1, cols=2, width=available_width)
        header_table.alignment = WD_TABLE_ALIGNMENT.LEFT

        # opcional: controlar proporção das colunas
        try:
            header_table.columns[0].width = Cm(5.0)
            header_table.columns[1].width = Cm(12.0)
        except Exception:
            pass

        cell_logo = header_table.rows[0].cells[0]
        p_logo = cell_logo.paragraphs[0]
        p_logo.alignment = WD_ALIGN_PARAGRAPH.LEFT

        if os.path.exists(logo_path):
            run_logo = p_logo.add_run()
            run_logo.add_picture(logo_path, width=Cm(3.5))
        else:
            run_logo = p_logo.add_run("DHL")
            run_logo.bold = True
            run_logo.font.name = "Arial"
            run_logo.font.size = Pt(12)

        header_table.rows[0].cells[1].text = ""

    # ===== Título e faixa =====
    p = doc.add_paragraph()
    style_heading(p, "ANÁLISE INVESTIGATIVA", size=16, bold=True, align="center")

    code_tbl = add_grid_table(doc, [[" DHL - SECURITY "]])
    for cell in code_tbl.rows[0].cells:
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        for run in cell.paragraphs[0].runs:
            run.font.size = Pt(9)

    # ===== DADOS DA OPERAÇÃO =====
    h2 = doc.add_paragraph()
    style_heading(h2, "Dados da Operação", size=14, bold=True, align="left")

    add_grid_table(
        doc,
        [["Nº do Relatório (ID):", id_relatorio]],
        col_widths_cm=[5.0, 15.5],
        header_bold_cols=[0],
    )

    add_grid_table(
        doc,
        [["Empresa:", dados_operacao["Empresa"], "Unidade:", dados_operacao["Unidade"]]],
        col_widths_cm=[3.0, 8.5, 3.0, 6.0],
        header_bold_cols=[0, 2],
    )

    add_grid_table(
        doc,
        [["Endereço:", dados_operacao["Endereço"]]],
        col_widths_cm=[3.0, 17.5],
        header_bold_cols=[0],
    )

    add_grid_table(
        doc,
        [["Classificação do Site:", dados_operacao["Classificação do Site"]]],
        col_widths_cm=[5.0, 15.5],
        header_bold_cols=[0],
    )

    add_grid_table(
        doc,
        [
            ["Produtos Segmento (Setor):", dados_operacao["Produtos Segmento (Setor)"]],
            ["Cliente(s):", dados_operacao["Cliente(s)"]],
        ],
        col_widths_cm=[6.0, 14.5],
        header_bold_cols=[0],
    )

    # ===== DADOS DO LEVANTAMENTO =====
    h2 = doc.add_paragraph()
    style_heading(h2, "Dados do Levantamento", size=14, bold=True, align="left")

    add_grid_table(
        doc,
        [
            ["Objetivo:", dados_levantamento["Objetivo"]],
            ["Responsável pelo Levantamento:", dados_levantamento["Responsável pelo Levantamento"]],
            ["Nome / Função / Data:", dados_levantamento["Nome / Função / Data"]],
        ],
        col_widths_cm=[6.0, 14.5],
        header_bold_cols=[0],
    )

    # ===== DESCRIÇÃO DO REGISTRO =====
    p_sub = doc.add_paragraph()
    style_heading(p_sub, "Descrição do Registro", size=12, bold=True, align="left")
    doc.add_paragraph(descricao_registro or "")

    # ===== EVIDÊNCIAS =====
    if evidencias:
        doc.add_paragraph().add_run("EVIDÊNCIAS ABAIXO").bold = True

        max_width_cm = 8.0
        row_pair = []
        for idx, (fimg, desc) in enumerate(evidencias, start=1):
            bio = BytesIO(fimg.read())
            bio.seek(0)
            row_pair.append((bio, f"Imagem {idx}: {desc or ''}"))

            if len(row_pair) == 2 or idx == len(evidencias):
                tbl = doc.add_table(rows=1, cols=2)
                tbl.alignment = WD_TABLE_ALIGNMENT.CENTER

                for ci, (img_bytes, legend) in enumerate(row_pair):
                    cell = tbl.rows[0].cells[ci]
                    run_img = cell.paragraphs[0].add_run()
                    run_img.add_picture(img_bytes, width=Cm(max_width_cm))

                    p_cap = cell.add_paragraph(legend)
                    p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    for rn in p_cap.runs:
                        rn.font.size = Pt(9)

                if len(row_pair) == 1:
                    tbl.rows[0].cells[1].text = ""

                row_pair = []

    # ===== CONCLUSÃO e SUGESTÃO =====
    h2 = doc.add_paragraph()
    style_heading(h2, "Conclusão", size=14, bold=True, align="left")
    doc.add_paragraph(conclusao or "")

    if (sugestao or "").strip():
        h2 = doc.add_paragraph()
        style_heading(h2, "Sugestão", size=14, bold=True, align="left")
        doc.add_paragraph(sugestao)

    # ===== Rodapé textual =====
    code_para = doc.add_paragraph()
    code_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if id_relatorio:
        run_id = code_para.add_run(f"Nº do Relatório (ID): {id_relatorio}")
        run_id.font.name = "Arial"
        run_id.font.size = Pt(9)

    base_name = "Analise_Investigativa"
    if id_relatorio:
        base_name = f"Analise_Investigativa_ID-{id_relatorio}"

    # ===== Gera DOCX em memória =====
    docx_buf = BytesIO()
    doc.save(docx_buf)
    docx_buf.seek(0)

    # ===== Converter para PDF e gerar ZIP =====
    try:
        pdf_io = convert_doc_to_pdf_bytes(doc, base_name=base_name)
        pdf_io.seek(0)

        zip_buf = BytesIO()
        with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
            zf.writestr(f"{base_name}.docx", docx_buf.getvalue())
            zf.writestr(f"{base_name}.pdf", pdf_io.getvalue())

        zip_buf.seek(0)
        return send_file(
            zip_buf,
            as_attachment=True,
            download_name=f"{base_name}_DOCX_PDF.zip",
            mimetype="application/zip",
        )

    except Exception:
        docx_buf.seek(0)
        return send_file(
            docx_buf,
            as_attachment=True,
            download_name=f"{base_name}.docx",
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )




@app.route("/login", methods=["GET", "POST"])
def login():
    if session.get("user_id"):
        return redirect(url_for("ocorrencias"))

    if request.method == "POST":
        username = (request.form.get("username") or "").strip().lower()
        senha = (request.form.get("senha") or "").strip()

        if not username or not senha:
            flash("Preencha usuário e senha.", "warning")
            return render_template("login.html")

        usuario = Usuario.query.filter_by(username=username, ativo=True).first()

        if not usuario or not check_senha(usuario, senha):
            flash("Usuário ou senha inválidos.", "danger")
            return render_template("login.html")

        session["user_id"] = usuario.id
        session["user_nome"] = usuario.nome
        session["username"] = usuario.username
        session["user_perfil"] = usuario.perfil

        flash("Login realizado com sucesso.", "success")
        return redirect(url_for("ocorrencias"))

    return render_template("login.html")


@app.route("/logout")
def logout():
    session.clear()
    flash("Sessão encerrada com sucesso.", "success")
    return redirect(url_for("login"))


# =========================
# USERS
# =========================
@app.route("/register", methods=["GET", "POST"])
@login_required
@perfil_required("SUPERUSUARIO")
def register():
    if request.method == "POST":
        nome = (request.form.get("nome") or "").strip()
        username = (request.form.get("username") or "").strip().lower()
        senha = (request.form.get("senha") or "").strip()
        confirmar_senha = (request.form.get("confirmar_senha") or "").strip()
        perfil = (request.form.get("perfil") or "").strip().upper()

        perfis_validos = {"SUPERUSUARIO", "USUARIO", "OPERACIONAL"}

        if not nome or not username or not senha or not confirmar_senha or not perfil:
            flash("Preencha todos os campos.", "warning")
            return render_template("register.html")

        if perfil not in perfis_validos:
            flash("Perfil inválido.", "danger")
            return render_template("register.html")

        if senha != confirmar_senha:
            flash("As senhas não conferem.", "danger")
            return render_template("register.html")

        if Usuario.query.filter_by(username=username).first():
            flash("Já existe um usuário com esse login.", "warning")
            return render_template("register.html")

        novo_usuario = Usuario(
            nome=nome,
            username=username,
            perfil=perfil,
            ativo=True
        )
        set_senha(novo_usuario, senha)

        db.session.add(novo_usuario)
        db.session.commit()

        flash("Usuário criado com sucesso.", "success")
        return redirect(url_for("listar_usuarios"))

    return render_template("register.html")


@app.route("/usuarios")
@login_required
@perfil_required("SUPERUSUARIO")
def listar_usuarios():
    usuarios = Usuario.query.order_by(Usuario.nome.asc()).all()
    return render_template("usuarios.html", usuarios=usuarios)


# =========================
# OCORRENCIAS
# =========================
@app.route("/ocorrencias", methods=["GET", "POST"])
@login_required
def ocorrencias():
    registro_edicao = None
    modo_edicao = False

    editar_id = request.args.get("editar", type=int)
    if editar_id:
        registro_edicao = Ocorrencia.query.get_or_404(editar_id)
        modo_edicao = True

    if request.method == "POST":
        ocorrencia_id = request.form.get("ocorrencia_id", type=int)

        data_hora = (request.form.get("data_hora") or "").strip()
        hora_ocorrencia = (request.form.get("hora_ocorrencia") or "").strip()
        natureza = (request.form.get("natureza") or "").strip()
        descricao = (request.form.get("descricao") or "").strip()
        local = (request.form.get("local") or "").strip()
        operador = (request.form.get("operador") or "").strip()
        gc = (request.form.get("gc") or "").strip()
        envolvido = (request.form.get("envolvido") or "").strip()
        prioridade = normalizar_prioridade(request.form.get("prioridade"))
        status = normalizar_status(request.form.get("status") or "PENDENTE")

        if not data_hora or not hora_ocorrencia or not natureza or not descricao or not local or not operador or not gc or not prioridade:
            flash("Preencha todos os campos obrigatórios.", "warning")
            if ocorrencia_id:
                return redirect(url_for("ocorrencias", editar=ocorrencia_id))
            return redirect(url_for("ocorrencias"))

        foto = request.files.get("foto")
        novo_nome_foto = None

        if foto and foto.filename:
            novo_nome_foto = salvar_arquivo(foto, EXTENSOES_PERMITIDAS_IMAGEM)
            if not novo_nome_foto:
                flash("Formato de imagem inválido. Use JPG, JPEG, PNG ou WEBP.", "danger")
                if ocorrencia_id:
                    return redirect(url_for("ocorrencias", editar=ocorrencia_id))
                return redirect(url_for("ocorrencias"))

        if ocorrencia_id:
            registro = Ocorrencia.query.get_or_404(ocorrencia_id)

            registro.data_hora = data_hora
            registro.hora_ocorrencia = hora_ocorrencia
            registro.natureza = natureza
            registro.descricao = descricao
            registro.local = local
            registro.operador = operador
            registro.gc = gc
            registro.envolvido = envolvido
            registro.prioridade = prioridade
            registro.status = status

            if novo_nome_foto:
                remover_arquivo(registro.foto)
                registro.foto = novo_nome_foto

            db.session.commit()
            flash("Ocorrência atualizada com sucesso.", "success")
            return redirect(url_for("ocorrencias"))

        nova = Ocorrencia(
            data_hora=data_hora,
            hora_ocorrencia=hora_ocorrencia,
            natureza=natureza,
            descricao=descricao,
            local=local,
            operador=operador,
            gc=gc,
            envolvido=envolvido,
            foto=novo_nome_foto,
            prioridade=prioridade,
            status=status,
            criado_por=session.get("user_nome")
        )
        db.session.add(nova)
        db.session.commit()

        flash("Ocorrência cadastrada com sucesso.", "success")
        return redirect(url_for("ocorrencias"))

    query = Ocorrencia.query.order_by(Ocorrencia.id.desc())
    registros, filtros = aplicar_filtros(query)

    hoje_str = datetime.now().strftime("%Y-%m-%d")
    total_hoje = len([r for r in registros if (r.data_hora or "").startswith(hoje_str)])

    resumo = {
        "total": len(registros),
        "hoje": total_hoje,
        "com_foto": len([r for r in registros if r.foto]),
        "pendentes": len([r for r in registros if normalizar_status(r.status) == "PENDENTE"]),
    }

    agora = datetime.now().strftime("%Y-%m-%dT%H:%M")
    hora_atual = datetime.now().strftime("%H:%M")

    return render_template(
        "ocorrencias.html",
        registros=registros,
        resumo=resumo,
        modo_edicao=modo_edicao,
        registro_edicao=registro_edicao,
        agora=agora,
        hora_atual=hora_atual,
        filtros=filtros
    )


@app.route("/post/<int:ocorrencia_id>", methods=["GET", "POST"])
@login_required
def post_ocorrencia(ocorrencia_id):
    registro = Ocorrencia.query.get_or_404(ocorrencia_id)

    if request.method == "POST":
        status_post = normalizar_status(request.form.get("status_post"))
        responsavel = (request.form.get("responsavel_fechamento") or "").strip()
        anexo_post = request.files.get("anexo_post")

        if not status_post:
            flash("Selecione a conclusão da publicação.", "warning")
            return redirect(url_for("post_ocorrencia", ocorrencia_id=registro.id))

        if status_post not in {"CONCLUIDO", "INCONCLUSIVA"}:
            flash("Opção inválida para a publicação.", "danger")
            return redirect(url_for("post_ocorrencia", ocorrencia_id=registro.id))

        if not responsavel:
            flash("Informe o responsável pelo fechamento.", "warning")
            return redirect(url_for("post_ocorrencia", ocorrencia_id=registro.id))

        registro.status = status_post
        registro.situacao_investigacao = status_post
        registro.responsavel_fechamento = responsavel

        if anexo_post and anexo_post.filename:
            novo_anexo = salvar_arquivo(anexo_post, EXTENSOES_PERMITIDAS_POST)
            if not novo_anexo:
                flash("Formato de arquivo inválido para o post.", "danger")
                return redirect(url_for("post_ocorrencia", ocorrencia_id=registro.id))

            remover_arquivo(registro.anexo_post)
            registro.anexo_post = novo_anexo

        db.session.commit()
        flash("Publicação da ocorrência atualizada com sucesso.", "success")
        return redirect(url_for("post_ocorrencia", ocorrencia_id=registro.id))

    return render_template("post_ocorrencia.html", registro=registro)
@app.route("/excluir/<int:ocorrencia_id>", methods=["POST"])
@login_required
@perfil_required("SUPERUSUARIO", "USUARIO")
def excluir_ocorrencia(ocorrencia_id):
    registro = Ocorrencia.query.get_or_404(ocorrencia_id)

    remover_arquivo(registro.foto)
    remover_arquivo(registro.anexo_post)

    db.session.delete(registro)
    db.session.commit()
    flash("Ocorrência excluída com sucesso.", "success")
    return redirect(url_for("ocorrencias"))


# =========================
# DASHBOARD
# =========================
@app.route("/dashboard")
@login_required
def dashboard():
    query = Ocorrencia.query.order_by(Ocorrencia.id.desc())
    registros, filtros = aplicar_filtros(query)

    total = len(registros)
    pendentes = len([r for r in registros if normalizar_status(r.status) == "PENDENTE"])
    andamento = len([r for r in registros if normalizar_status(r.status) == "EM ANDAMENTO"])
    concluidas = len([r for r in registros if normalizar_status(r.status) == "CONCLUIDO"])
    altas = len([r for r in registros if normalizar_prioridade(r.prioridade) == "ALTA"])

    natureza_count = {}
    local_count = {}
    status_count = {}

    for r in registros:
        natureza_key = r.natureza or "Não informado"
        local_key = r.local or "Não informado"
        status_key = r.status or "Não informado"

        natureza_count[natureza_key] = natureza_count.get(natureza_key, 0) + 1
        local_count[local_key] = local_count.get(local_key, 0) + 1
        status_count[status_key] = status_count.get(status_key, 0) + 1

    return render_template(
        "dashboard.html",
        registros=registros[:10],
        filtros=filtros,
        resumo={
            "total": total,
            "pendentes": pendentes,
            "andamento": andamento,
            "concluidas": concluidas,
            "altas": altas,
        },
        labels_natureza=list(natureza_count.keys()),
        valores_natureza=list(natureza_count.values()),
        labels_local=list(local_count.keys()),
        valores_local=list(local_count.values()),
        labels_status=list(status_count.keys()),
        valores_status=list(status_count.values()),
    )


# =========================
# EXPORTACOES
# =========================
@app.route("/exportar/excel")
@login_required
def exportar_excel():
    query = Ocorrencia.query.order_by(Ocorrencia.id.desc())
    registros, _ = aplicar_filtros(query)

    wb = Workbook()
    ws = wb.active
    ws.title = "Ocorrencias"

    headers = [
        "ID", "Data/Hora", "Hora Ocorrência", "Natureza", "Descrição",
        "Local", "Operador", "GC", "Envolvido", "Prioridade",
        "Status", "Situação Investigação", "Conclusão Investigação", "Criado por"
    ]
    ws.append(headers)

    fill = PatternFill("solid", fgColor="FFCC00")
    font = Font(bold=True)

    for col in range(1, len(headers) + 1):
        ws.cell(row=1, column=col).fill = fill
        ws.cell(row=1, column=col).font = font

    for r in registros:
        ws.append([
            r.id,
            r.data_hora,
            r.hora_ocorrencia,
            r.natureza,
            r.descricao,
            r.local,
            r.operador,
            r.gc,
            r.envolvido,
            r.prioridade,
            r.status,
            r.situacao_investigacao,
            r.conclusao_investigacao,
            r.criado_por
        ])

    output = BytesIO()
    wb.save(output)
    output.seek(0)

    return send_file(
        output,
        as_attachment=True,
        download_name="controle_ocorrencias.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


@app.route("/exportar/pdf")
@login_required
def exportar_pdf():
    query = Ocorrencia.query.order_by(Ocorrencia.id.desc())
    registros, _ = aplicar_filtros(query)

    buffer = BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        leftMargin=25,
        rightMargin=25,
        topMargin=25,
        bottomMargin=25
    )
    styles = getSampleStyleSheet()
    elements = []

    elements.append(Paragraph("Controle de Ocorrências - Relatório PDF", styles["Title"]))
    elements.append(Spacer(1, 12))

    data = [["ID", "Data/Hora", "Natureza", "Local", "Status", "Prioridade"]]
    for r in registros:
        data.append([
            str(r.id),
            r.data_hora or "-",
            r.natureza or "-",
            r.local or "-",
            r.status or "-",
            r.prioridade or "-"
        ])

    tabela = Table(data, repeatRows=1)
    tabela.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#ffcc00")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.black),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
        ("FONTSIZE", (0, 0), (-1, -1), 8),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.whitesmoke, colors.lightgrey]),
    ]))

    elements.append(tabela)
    doc.build(elements)
    buffer.seek(0)

    return send_file(
        buffer,
        as_attachment=True,
        download_name="controle_ocorrencias.pdf",
        mimetype="application/pdf"
    )


# =========================
# UPLOADS
# =========================
@app.route("/uploads/<filename>")
@login_required
def uploaded_file(filename):
    return send_from_directory(app.config["UPLOAD_FOLDER"], filename)


# =========================
# INIT DB
# =========================
with app.app_context():
    db.create_all()

    existe_super = Usuario.query.filter_by(username="admin").first()
    if not existe_super:
        user = Usuario(
            nome="Administrador Master",
            username="admin",
            perfil="SUPERUSUARIO",
            ativo=True
        )
        set_senha(user, "123456")
        db.session.add(user)
        db.session.commit()


if __name__ == "__main__":
    app.run(debug=True)