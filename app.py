import os
import base64
import zipfile
from io import BytesIO
from datetime import datetime
from functools import wraps
import oracledb
from docx.shared import Cm, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from flask import (
    Flask, render_template, request, redirect, url_for,
    flash, session, send_file, current_app
)

from flask_sqlalchemy import SQLAlchemy
from werkzeug.security import generate_password_hash, check_password_hash

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

app = Flask(__name__)
app.config["SECRET_KEY"] = "controle-ocorrencia-executivo"
app.config["SQLALCHEMY_DATABASE_URI"] = (
    "oracle+oracledb://SECPANEL:SEC003q2w3e4r2026"
    "@usqasap023-scan.phx-dc.dhl.com:1521"
    "/?service_name=SECPANEL"
)
app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False
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
    __tablename__ = "USERS_LIVRO"

    id = db.Column("ID", db.Integer, db.Identity(start=1), primary_key=True)
    nome = db.Column("NOME", db.String(120), nullable=False)
    email = db.Column("EMAIL", db.String(120), unique=True, nullable=False, index=True)
    password_hash = db.Column("PASSWORD_HASH", db.String(255), nullable=False)
    perfil = db.Column("ROLE", db.String(30), nullable=False, default="OPERACIONAL")
    site = db.Column("SITE", db.String(80), nullable=True)
    is_active = db.Column("IS_ACTIVE", db.Boolean, nullable=False, default=True)
    created_at = db.Column("CREATED_AT", db.DateTime, nullable=False, default=datetime.utcnow)

    def set_password(self, senha: str):
        self.password_hash = generate_password_hash(senha)

    def check_password(self, senha: str) -> bool:
        return check_password_hash(self.password_hash, senha)

class AnaliseInvestigativa(db.Model):
    __tablename__ = "analises_investigativas"

    id = db.Column(db.Integer, db.Identity(start=1), primary_key=True)

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

    id = db.Column(db.Integer, db.Identity(start=1), primary_key=True)
    analise_id = db.Column(db.Integer, db.ForeignKey("analises_investigativas.id"), nullable=False)

    arquivo = db.Column(db.String(255), nullable=False)
    descricao = db.Column(db.Text, nullable=False)

    criado_em = db.Column(db.DateTime, default=datetime.utcnow)

class Ocorrencia(db.Model):
    __tablename__ = "ocorrencias"

    id = db.Column(db.Integer, db.Identity(start=1), primary_key=True)
    data_hora = db.Column(db.String(30), nullable=False)
    hora_ocorrencia = db.Column(db.String(10), nullable=False)
    natureza = db.Column(db.String(120), nullable=False)
    descricao = db.Column(db.Text, nullable=False)
    local = db.Column(db.String(120), nullable=False)
    operador = db.Column(db.String(120), nullable=False)
    gc = db.Column(db.String(120), nullable=False)
    envolvido = db.Column(db.String(120), nullable=True)
    foto = db.Column(db.Text, nullable=True)

    prioridade = db.Column(db.String(20), nullable=False, default="MEDIA")
    status = db.Column(db.String(30), nullable=False, default="PENDENTE")

    situacao_investigacao = db.Column(db.String(30), nullable=True)
    conclusao_investigacao = db.Column(db.Text, nullable=True)
    anexo_post = db.Column(db.Text, nullable=True)
    anexo_post_nome = db.Column(db.String(255), nullable=True)

    criado_por = db.Column(db.String(120), nullable=True)
    criado_em = db.Column(db.DateTime, default=datetime.utcnow)
    atualizado_em = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)
    responsavel_fechamento = db.Column(db.String(120), nullable=True)


# =========================
# HELPERS
# =========================
MIME_MAP = {
    "jpg": "image/jpeg", "jpeg": "image/jpeg",
    "png": "image/png", "webp": "image/webp",
    "gif": "image/gif", "pdf": "application/pdf",
    "doc": "application/msword",
    "docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    "xls": "application/vnd.ms-excel",
    "xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
}

def arquivo_para_base64(arquivo, extensoes):
    if not arquivo or not arquivo.filename:
        return None, None
    ext = arquivo.filename.rsplit(".", 1)[-1].lower() if "." in arquivo.filename else ""
    if ext not in extensoes:
        return None, None
    mime = MIME_MAP.get(ext, "application/octet-stream")
    dados = base64.b64encode(arquivo.read()).decode("utf-8")
    return f"data:{mime};base64,{dados}", arquivo.filename


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

def set_cell_bg(cell, color):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), color)
    tcPr.append(shd)


def set_cell_border(cell, color="D9D9D9", size="8"):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    borders = OxmlElement("w:tcBorders")

    for border_name in ["top", "left", "bottom", "right"]:
        border = OxmlElement(f"w:{border_name}")
        border.set(qn("w:val"), "single")
        border.set(qn("w:sz"), size)
        border.set(qn("w:space"), "0")
        border.set(qn("w:color"), color)
        borders.append(border)

    tcPr.append(borders)


def format_cell(cell, bold=False, bg=None, align="left"):
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER

    if bg:
        set_cell_bg(cell, bg)

    set_cell_border(cell)

    for p in cell.paragraphs:
        p.paragraph_format.space_before = Pt(3)
        p.paragraph_format.space_after = Pt(3)

        if align == "center":
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        else:
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT

        for run in p.runs:
            run.font.name = "Arial"
            run.font.size = Pt(9)
            run.bold = bold
            run.font.color.rgb = RGBColor(31, 41, 55)


def add_section_title(doc, title):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(14)
    p.paragraph_format.space_after = Pt(6)

    run = p.add_run(title.upper())
    run.bold = True
    run.font.name = "Arial"
    run.font.size = Pt(11)
    run.font.color.rgb = RGBColor(212, 5, 17)

    return p


def add_professional_table(doc, rows, col_widths):
    table = doc.add_table(rows=0, cols=len(col_widths))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False

    for row_data in rows:
        row = table.add_row()
        for i, value in enumerate(row_data):
            cell = row.cells[i]
            cell.width = Cm(col_widths[i])
            cell.text = str(value or "")

            is_label = i % 2 == 0
            format_cell(
                cell,
                bold=is_label,
                bg="FFF2CC" if is_label else "FFFFFF",
                align="left"
            )

    doc.add_paragraph().paragraph_format.space_after = Pt(3)
    return table

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

    doc = Document()

    # ===== Margens =====
    for section in doc.sections:
        section.left_margin = Cm(1.7)
        section.right_margin = Cm(1.7)
        section.top_margin = Cm(1.5)
        section.bottom_margin = Cm(1.5)

    # ==========================================================
    # FUNÇÕES DE FORMATAÇÃO
    # ==========================================================
    def set_cell_bg(cell, color):
        tcPr = cell._tc.get_or_add_tcPr()
        shd = OxmlElement("w:shd")
        shd.set(qn("w:fill"), color)
        tcPr.append(shd)

    def set_cell_border(cell, color="D9D9D9", size="8"):
        tcPr = cell._tc.get_or_add_tcPr()
        tcBorders = OxmlElement("w:tcBorders")

        for border_name in ["top", "left", "bottom", "right"]:
            border = OxmlElement(f"w:{border_name}")
            border.set(qn("w:val"), "single")
            border.set(qn("w:sz"), size)
            border.set(qn("w:space"), "0")
            border.set(qn("w:color"), color)
            tcBorders.append(border)

        tcPr.append(tcBorders)

    def format_cell(cell, bold=False, bg="FFFFFF", align="left"):
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        set_cell_bg(cell, bg)
        set_cell_border(cell)

        for p in cell.paragraphs:
            p.paragraph_format.space_before = Pt(4)
            p.paragraph_format.space_after = Pt(4)

            if align == "center":
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            else:
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT

            for run in p.runs:
                run.font.name = "Arial"
                run.font.size = Pt(9)
                run.bold = bold
                run.font.color.rgb = RGBColor(31, 41, 55)

    def add_section_title(doc, title):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(14)
        p.paragraph_format.space_after = Pt(6)

        run = p.add_run(title.upper())
        run.bold = True
        run.font.name = "Arial"
        run.font.size = Pt(11)
        run.font.color.rgb = RGBColor(212, 5, 17)

    def add_professional_table(doc, rows, col_widths):
        table = doc.add_table(rows=0, cols=len(col_widths))
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        table.autofit = False

        for row_data in rows:
            row = table.add_row()

            for i, value in enumerate(row_data):
                cell = row.cells[i]
                cell.width = Cm(col_widths[i])
                cell.text = str(value or "")

                is_label = i % 2 == 0

                format_cell(
                    cell,
                    bold=is_label,
                    bg="FFF2CC" if is_label else "FFFFFF",
                    align="left"
                )

        doc.add_paragraph().paragraph_format.space_after = Pt(2)
        return table

    def add_text_box(doc, text="\n\n\n"):
        table = doc.add_table(rows=1, cols=1)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER

        cell = table.rows[0].cells[0]
        cell.text = text or "\n\n\n"
        format_cell(cell, bold=False, bg="FFFFFF")

        return table

    # ==========================================================
    # CABEÇALHO COM LOGO
    # ==========================================================
    logo_path = os.path.join(app.root_path, "static", "logo.png")

    for section in doc.sections:
        header = section.header

        for p in header.paragraphs:
            p.text = ""

        available_width = section.page_width - section.left_margin - section.right_margin

        header_table = header.add_table(rows=1, cols=2, width=available_width)
        header_table.alignment = WD_TABLE_ALIGNMENT.CENTER
        header_table.autofit = False

        header_table.columns[0].width = Cm(5.0)
        header_table.columns[1].width = Cm(14.0)

        cell_logo = header_table.rows[0].cells[0]
        cell_info = header_table.rows[0].cells[1]

        for cell in [cell_logo, cell_info]:
            set_cell_border(cell, color="FFFFFF")
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER

        p_logo = cell_logo.paragraphs[0]
        p_logo.alignment = WD_ALIGN_PARAGRAPH.LEFT

        if os.path.exists(logo_path):
            run_logo = p_logo.add_run()
            run_logo.add_picture(logo_path, width=Cm(3.8))
        else:
            run_logo = p_logo.add_run("DHL")
            run_logo.bold = True
            run_logo.font.name = "Arial"
            run_logo.font.size = Pt(16)
            run_logo.font.color.rgb = RGBColor(212, 5, 17)

        p_info = cell_info.paragraphs[0]
        p_info.alignment = WD_ALIGN_PARAGRAPH.RIGHT

        r1 = p_info.add_run("DHL SECURITY\n")
        r1.bold = True
        r1.font.name = "Arial"
        r1.font.size = Pt(10)
        r1.font.color.rgb = RGBColor(212, 5, 17)

        r2 = p_info.add_run("Relatório Corporativo de Análise Investigativa")
        r2.font.name = "Arial"
        r2.font.size = Pt(8)
        r2.font.color.rgb = RGBColor(90, 90, 90)

    doc.add_paragraph()

    # ==========================================================
    # TÍTULO PRINCIPAL
    # ==========================================================
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(10)

    run = p.add_run("ANÁLISE INVESTIGATIVA")
    run.bold = True
    run.font.name = "Arial"
    run.font.size = Pt(17)
    run.font.color.rgb = RGBColor(31, 41, 55)

    # ==========================================================
    # FAIXA DHL SECURITY
    # ==========================================================
    faixa = doc.add_table(rows=1, cols=1)
    faixa.alignment = WD_TABLE_ALIGNMENT.CENTER

    cell = faixa.rows[0].cells[0]
    cell.text = "DHL SECURITY"
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER

    set_cell_bg(cell, "FFCC00")
    set_cell_border(cell, color="D40511", size="12")

    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(6)

    for run in p.runs:
        run.bold = True
        run.font.name = "Arial"
        run.font.size = Pt(10)
        run.font.color.rgb = RGBColor(120, 0, 0)

    doc.add_paragraph()

    # ==========================================================
    # DADOS DA OPERAÇÃO
    # ==========================================================
    add_section_title(doc, "Dados da Operação")

    add_professional_table(
        doc,
        [
            ["Nº do Relatório (ID):", id_relatorio],
            ["Empresa:", dados_operacao["Empresa"]],
            ["Unidade:", dados_operacao["Unidade"]],
            ["Endereço:", dados_operacao["Endereço"]],
            ["Classificação do Site:", dados_operacao["Classificação do Site"]],
            ["Produtos Segmento (Setor):", dados_operacao["Produtos Segmento (Setor)"]],
            ["Cliente(s):", dados_operacao["Cliente(s)"]],
        ],
        col_widths=[6.0, 13.5]
    )

    # ==========================================================
    # DADOS DO LEVANTAMENTO
    # ==========================================================
    add_section_title(doc, "Dados do Levantamento")

    add_professional_table(
        doc,
        [
            ["Objetivo:", dados_levantamento["Objetivo"]],
            ["Responsável pelo Levantamento:", dados_levantamento["Responsável pelo Levantamento"]],
            ["Nome / Função / Data:", dados_levantamento["Nome / Função / Data"]],
        ],
        col_widths=[6.0, 13.5]
    )

    # ==========================================================
    # DESCRIÇÃO DO REGISTRO
    # ==========================================================
    add_section_title(doc, "Descrição do Registro")
    add_text_box(doc, descricao_registro or "\n\n\n\n")

    # ==========================================================
    # EVIDÊNCIAS
    # ==========================================================
    if evidencias:
        add_section_title(doc, "Evidências")

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
                    format_cell(cell, bg="FFFFFF", align="center")

                    p_img = cell.paragraphs[0]
                    p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER

                    run_img = p_img.add_run()
                    run_img.add_picture(img_bytes, width=Cm(max_width_cm))

                    p_cap = cell.add_paragraph(legend)
                    p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER

                    for rn in p_cap.runs:
                        rn.font.name = "Arial"
                        rn.font.size = Pt(8)
                        rn.font.color.rgb = RGBColor(90, 90, 90)

                if len(row_pair) == 1:
                    tbl.rows[0].cells[1].text = ""

                doc.add_paragraph()
                row_pair = []

    # ==========================================================
    # CONCLUSÃO
    # ==========================================================
    add_section_title(doc, "Conclusão")
    add_text_box(doc, conclusao or "\n\n\n\n")

    # ==========================================================
    # SUGESTÃO
    # ==========================================================
    add_section_title(doc, "Sugestão")
    add_text_box(doc, sugestao or "\n\n\n\n")

    # ==========================================================
    # RODAPÉ TEXTUAL
    # ==========================================================
    doc.add_paragraph()

    code_para = doc.add_paragraph()
    code_para.alignment = WD_ALIGN_PARAGRAPH.CENTER

    texto_rodape = (
        f"Nº do Relatório (ID): {id_relatorio}"
        if id_relatorio
        else "Relatório gerado pelo sistema DHL Security"
    )

    run_id = code_para.add_run(texto_rodape)
    run_id.font.name = "Arial"
    run_id.font.size = Pt(8)
    run_id.font.color.rgb = RGBColor(100, 100, 100)

    # ==========================================================
    # NOME DO ARQUIVO
    # ==========================================================
    base_name = "Analise_Investigativa"

    if id_relatorio:
        base_name = f"Analise_Investigativa_ID-{id_relatorio}"

    # ==========================================================
    # GERA DOCX EM MEMÓRIA
    # ==========================================================
    docx_buf = BytesIO()
    doc.save(docx_buf)
    docx_buf.seek(0)

    # ==========================================================
    # CONVERTE PARA PDF E GERA ZIP
    # ==========================================================
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
        email = (request.form.get("email") or "").strip().lower()
        senha = (request.form.get("password") or "").strip()

        if not email or not senha:
            flash("Preencha e-mail e senha.", "warning")
            return render_template("login.html")

        usuario = Usuario.query.filter_by(email=email, is_active=True).first()

        if not usuario or not usuario.check_password(senha):
            flash("E-mail ou senha inválidos.", "danger")
            return render_template("login.html")

        session["user_id"] = usuario.id
        session["user_nome"] = usuario.nome
        session["username"] = usuario.email
        session["user_perfil"] = usuario.perfil
        session["user_site"] = usuario.site or ""

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
        email = (request.form.get("email") or "").strip().lower()
        senha = (request.form.get("senha") or "").strip()
        confirmar_senha = (request.form.get("confirmar_senha") or "").strip()
        perfil = (request.form.get("perfil") or "").strip().upper()

        perfis_validos = {"SUPERUSUARIO", "USUARIO", "OPERACIONAL"}

        if not nome or not email or not senha or not confirmar_senha or not perfil:
            flash("Preencha todos os campos.", "warning")
            return render_template("register.html")

        if perfil not in perfis_validos:
            flash("Perfil inválido.", "danger")
            return render_template("register.html")

        if senha != confirmar_senha:
            flash("As senhas não conferem.", "danger")
            return render_template("register.html")

        if Usuario.query.filter_by(email=email).first():
            flash("Já existe um usuário com esse e-mail.", "warning")
            return render_template("register.html")

        site = (request.form.get("site") or "").strip()

        novo_usuario = Usuario(
            nome=nome,
            email=email,
            perfil=perfil,
            site=site,
            is_active=True
        )
        novo_usuario.set_password(senha)

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
        nova_foto_b64 = None

        if foto and foto.filename:
            nova_foto_b64, _ = arquivo_para_base64(foto, EXTENSOES_PERMITIDAS_IMAGEM)
            if not nova_foto_b64:
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

            if nova_foto_b64:
                registro.foto = nova_foto_b64

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
            foto=nova_foto_b64,
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
            novo_anexo_b64, nome_original = arquivo_para_base64(anexo_post, EXTENSOES_PERMITIDAS_POST)
            if not novo_anexo_b64:
                flash("Formato de arquivo inválido para o post.", "danger")
                return redirect(url_for("post_ocorrencia", ocorrencia_id=registro.id))

            registro.anexo_post = novo_anexo_b64
            registro.anexo_post_nome = nome_original

        db.session.commit()
        flash("Publicação da ocorrência atualizada com sucesso.", "success")
        return redirect(url_for("post_ocorrencia", ocorrencia_id=registro.id))

    return render_template("post_ocorrencia.html", registro=registro)
@app.route("/excluir/<int:ocorrencia_id>", methods=["POST"])
@login_required
@perfil_required("SUPERUSUARIO", "USUARIO")
def excluir_ocorrencia(ocorrencia_id):
    registro = Ocorrencia.query.get_or_404(ocorrencia_id)

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
# INIT DB
# =========================
with app.app_context():
    # Cria tabelas da aplicação no Oracle (USERS_LIVRO já existe, será ignorada)
    db.create_all()


if __name__ == "__main__":
    app.run(debug=True)