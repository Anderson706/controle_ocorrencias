import os
import re
import base64
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
from reportlab.lib.units import cm as rcm
from reportlab.lib.enums import TA_CENTER, TA_RIGHT
from reportlab.lib.utils import ImageReader
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer,
    Image as RLImage, Table, TableStyle, PageBreak
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

class SiteCompleto(db.Model):
    __tablename__ = "SITES_COMPLETO"

    nome_do_site = db.Column("NOME_DO_SITE", db.String(128), primary_key=True)
    endereco     = db.Column("ENDEREÇO", db.String(255), nullable=True)
    cidade       = db.Column("CIDADE", db.String(50), nullable=True)
    estado       = db.Column("ESTADO", db.String(2), nullable=True)
    pais         = db.Column("PAÍS", db.String(26), nullable=True)
    responsavel_security = db.Column("RESPONSÁVEL_SECURITY", db.String(128), nullable=True)
    coordenador  = db.Column("COORDENADOR", db.String(26), nullable=True)
    latitude     = db.Column("LATIDUDE", db.Numeric(38, 0), nullable=True)
    longitude    = db.Column("LONGITUDE", db.Numeric(38, 0), nullable=True)
    sector       = db.Column("SECTOR", db.String(26), nullable=True)
    security_responsible = db.Column("SECURITY_RESPONSIBLE", db.String(26), nullable=True)


class AnaliseInvestigativa(db.Model):
    __tablename__ = "analises_investigativas"

    id = db.Column(db.Integer, db.Identity(start=1), primary_key=True)
    codigo = db.Column(db.String(30), nullable=True, unique=True)
    numero_site = db.Column(db.Integer, nullable=True)
    site = db.Column(db.String(128), nullable=True)

    id_relatorio = db.Column(db.String(30), nullable=True)
    empresa = db.Column(db.String(120), nullable=True)
    unidade = db.Column(db.String(180), nullable=True)
    endereco = db.Column(db.String(255), nullable=True)
    classificacao = db.Column(db.String(80), nullable=True)
    produtos_segmento = db.Column(db.String(120), nullable=True)
    clientes = db.Column(db.String(120), nullable=True)

    objetivo = db.Column(db.Text, nullable=True)
    responsavel = db.Column(db.String(150), nullable=True)
    nome_funcao_data = db.Column(db.String(255), nullable=True)

    descricao_registro = db.Column(db.Text, nullable=True)
    conclusao = db.Column(db.Text, nullable=True)
    sugestao = db.Column(db.Text, nullable=True)

    criado_por = db.Column(db.String(120), nullable=True)
    criado_em = db.Column(db.DateTime, default=datetime.utcnow)
    docx_arquivo = db.Column(db.Text, nullable=True)

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


class ANC(db.Model):
    __tablename__ = "ancs"

    id = db.Column(db.Integer, db.Identity(start=1), primary_key=True)
    numero_anc = db.Column(db.String(50), nullable=True, unique=True)
    data_nc = db.Column(db.String(10), nullable=True)
    hora_nc = db.Column(db.String(5), nullable=True)
    site = db.Column(db.String(120), nullable=True)
    setor = db.Column(db.String(120), nullable=True)
    tipo_ocorrencia = db.Column(db.String(120), nullable=True)
    gravidade = db.Column(db.String(30), nullable=True)
    natureza = db.Column(db.String(255), nullable=True)
    responsavel = db.Column(db.String(120), nullable=True)
    gestor_responsavel = db.Column(db.String(120), nullable=True)
    local = db.Column(db.String(120), nullable=True)
    envolvido = db.Column(db.String(255), nullable=True)
    tipo = db.Column(db.String(80), nullable=True)
    turno = db.Column(db.String(20), nullable=True)
    status = db.Column(db.String(30), nullable=False, default="ABERTO")
    descricao = db.Column(db.Text, nullable=True)
    inicio_investigacao = db.Column(db.String(16), nullable=True)
    fim_investigacao = db.Column(db.String(16), nullable=True)
    imagem_1 = db.Column(db.Text, nullable=True)
    imagem_2 = db.Column(db.Text, nullable=True)
    imagem_3 = db.Column(db.Text, nullable=True)
    imagem_4 = db.Column(db.Text, nullable=True)
    imagem_5 = db.Column(db.Text, nullable=True)
    imagem_6 = db.Column(db.Text, nullable=True)
    numero_site = db.Column(db.Integer, nullable=True)
    criado_por = db.Column(db.String(80), nullable=True)
    criado_em = db.Column(db.DateTime, default=datetime.utcnow)
    atualizado_em = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)


class Ocorrencia(db.Model):
    __tablename__ = "ocorrencias"

    id = db.Column(db.Integer, db.Identity(start=1), primary_key=True)
    codigo = db.Column(db.String(30), nullable=True, unique=True)
    numero_site = db.Column(db.Integer, nullable=True)
    data_hora = db.Column(db.String(30), nullable=False)
    hora_ocorrencia = db.Column(db.String(10), nullable=False)
    natureza = db.Column(db.String(120), nullable=False)
    descricao = db.Column(db.Text, nullable=False)
    site = db.Column(db.String(128), nullable=True)
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

def gerar_codigo(model_class, site):
    clean = re.sub(r'[^A-Z0-9]', '', (site or "SITE").upper())[:8] or "SITE"
    ano = datetime.now().year
    seq = model_class.query.filter_by(site=site).count() + 1
    return f"{clean}-{ano}-{seq:04d}", seq


def gerar_numero_anc(site):
    clean = re.sub(r'[^A-Z0-9]', '', (site or "SITE").upper())[:8] or "SITE"
    ano = datetime.now().year
    seq = ANC.query.filter_by(site=site).count() + 1
    return f"ANC-{clean}-{ano}-{seq:04d}", seq


def fmt_data_br(data_str):
    """Converte YYYY-MM-DD para DD/MM/YYYY."""
    if data_str and len(data_str) >= 10:
        try:
            return datetime.strptime(data_str[:10], "%Y-%m-%d").strftime("%d/%m/%Y")
        except ValueError:
            pass
    return data_str or "—"


def aplicar_filtros_anc(query):
    data_inicial = (request.args.get("data_inicial") or "").strip()
    data_final   = (request.args.get("data_final")   or "").strip()
    status       = (request.args.get("status")        or "").strip().upper()
    gravidade    = (request.args.get("gravidade")     or "").strip().upper()
    turno        = (request.args.get("turno")         or "").strip().upper()
    setor        = (request.args.get("setor")         or "").strip()
    natureza     = (request.args.get("natureza")      or "").strip()

    registros = query.all()
    filtrados = []
    for r in registros:
        ok = True
        if data_inicial: ok = ok and ((r.data_nc or "") >= data_inicial)
        if data_final:   ok = ok and ((r.data_nc or "") <= data_final)
        if status:       ok = ok and ((r.status or "").upper() == status)
        if gravidade:    ok = ok and ((r.gravidade or "").upper() == gravidade)
        if turno:        ok = ok and ((r.turno or "").upper() == turno)
        if setor:        ok = ok and (setor.lower() in (r.setor or "").lower())
        if natureza:     ok = ok and (natureza.lower() in (r.natureza or "").lower())
        if ok:
            filtrados.append(r)

    filtros = {
        "data_inicial": data_inicial, "data_final": data_final,
        "status": status, "gravidade": gravidade,
        "turno": turno, "setor": setor, "natureza": natureza,
    }
    return filtrados, filtros


def gerar_pdf_anc_bytes(anc):
    """Gera PDF do ANC no formato oficial DHL Security."""
    buffer  = BytesIO()
    BLACK   = colors.black
    YELLOW  = colors.HexColor("#FFCC00")

    pw = A4[0] - 3.0 * rcm
    doc_pdf = SimpleDocTemplate(
        buffer, pagesize=A4,
        leftMargin=1.5*rcm, rightMargin=1.5*rcm,
        topMargin=1.5*rcm, bottomMargin=1.5*rcm,
    )

    s_normal = ParagraphStyle("an", fontName="Helvetica",      fontSize=9,  textColor=BLACK)
    s_th     = ParagraphStyle("ath",fontName="Helvetica-Bold", fontSize=9,  textColor=BLACK, alignment=TA_CENTER)
    s_td     = ParagraphStyle("atd",fontName="Helvetica",      fontSize=9,  textColor=BLACK, alignment=TA_CENTER)
    s_h3     = ParagraphStyle("ah3",fontName="Helvetica-Bold", fontSize=11, textColor=BLACK)
    s_title  = ParagraphStyle("ati",fontName="Helvetica-Bold", fontSize=13, textColor=BLACK, alignment=TA_CENTER)
    s_foot   = ParagraphStyle("afo",fontName="Helvetica",      fontSize=9,  textColor=BLACK)

    def _fit_img(source, max_w, max_h):
        """Retorna RLImage escalada proporcionalmente para caber em max_w × max_h."""
        if isinstance(source, str):
            raw = source.split(",", 1)[1] if "," in source else source
            bio = BytesIO(base64.b64decode(raw))
        else:
            bio = source
        bio.seek(0)
        iw, ih = ImageReader(bio).getSize()
        scale = min(max_w / iw, max_h / ih, 1.0)
        bio.seek(0)
        return RLImage(bio, width=iw * scale, height=ih * scale)

    def yellow_bar(text):
        t = Table([[Paragraph(text, s_h3)]], colWidths=[pw])
        t.setStyle(TableStyle([
            ("BACKGROUND",    (0,0),(-1,-1), YELLOW),
            ("BOX",           (0,0),(-1,-1), 0.5, BLACK),
            ("TOPPADDING",    (0,0),(-1,-1), 5),
            ("BOTTOMPADDING", (0,0),(-1,-1), 5),
            ("LEFTPADDING",   (0,0),(-1,-1), 8),
        ]))
        return t

    story = []

    # ── 1. CABEÇALHO ──────────────────────────────────────────────
    logo_path = os.path.join(app.root_path, "static", "logo.png")
    if os.path.exists(logo_path):
        with open(logo_path, "rb") as _lf:
            logo_cell = _fit_img(BytesIO(_lf.read()), max_w=pw * 0.22, max_h=1.6*rcm)
    else:
        logo_cell = Paragraph("<b>DHL</b>", s_normal)

    hdr = Table(
        [[logo_cell, Paragraph("AVISO DE NÃO CONFORMIDADE   ", s_title)]],
        colWidths=[pw * 0.25, pw * 0.75],
    )
    hdr.setStyle(TableStyle([
        ("BACKGROUND",    (1,0),(1,0), YELLOW),
        ("BOX",           (0,0),(-1,-1), 0.5, BLACK),
        ("INNERGRID",     (0,0),(-1,-1), 0.5, BLACK),
        ("VALIGN",        (0,0),(-1,-1), "MIDDLE"),
        ("TOPPADDING",    (0,0),(-1,-1), 8),
        ("BOTTOMPADDING", (0,0),(-1,-1), 8),
        ("LEFTPADDING",   (0,0),(-1,-1), 8),
        ("RIGHTPADDING",  (0,0),(-1,-1), 8),
    ]))
    story += [hdr, Spacer(1, 0.5*rcm)]

    # ── 2. TABELA DE IDENTIFICAÇÃO ────────────────────────────────
    id_heads = ["DATA", "HORA DA\nOCORRÊNCIA", "NATUREZA", "LOCAL",
                "PESSOAS\nENVOLVIDAS", "STATUS", "Nº ANC"]
    id_vals  = [fmt_data_br(anc.data_nc), anc.hora_nc or "—", anc.natureza or "—",
                anc.local or "—", anc.envolvido or "—",
                anc.status or "—", str(anc.numero_site or anc.id)]
    cw = pw / 7
    id_tbl = Table(
        [[Paragraph(h, s_th) for h in id_heads],
         [Paragraph(v, s_td) for v in id_vals]],
        colWidths=[cw] * 7,
    )
    id_tbl.setStyle(TableStyle([
        ("BACKGROUND",    (0,0),(-1,0), YELLOW),
        ("BOX",           (0,0),(-1,-1), 0.5, BLACK),
        ("INNERGRID",     (0,0),(-1,-1), 0.5, BLACK),
        ("VALIGN",        (0,0),(-1,-1), "MIDDLE"),
        ("TOPPADDING",    (0,0),(-1,-1), 6),
        ("BOTTOMPADDING", (0,0),(-1,-1), 6),
    ]))
    story += [id_tbl, Spacer(1, 0.5*rcm)]

    # ── 3. DESCRIÇÃO ──────────────────────────────────────────────
    story.append(yellow_bar("Descrição da ocorrência:"))
    desc_tbl = Table([[Paragraph(anc.descricao or "—", s_normal)]], colWidths=[pw])
    desc_tbl.setStyle(TableStyle([
        ("BOX",           (0,0),(-1,-1), 0.5, BLACK),
        ("TOPPADDING",    (0,0),(-1,-1), 10),
        ("BOTTOMPADDING", (0,0),(-1,-1), 40),
        ("LEFTPADDING",   (0,0),(-1,-1), 10),
        ("RIGHTPADDING",  (0,0),(-1,-1), 10),
    ]))
    story += [desc_tbl, Spacer(1, 0.5*rcm)]

    # ── 4. GRAVIDADE / RESPONSÁVEL / STATUS ───────────────────────
    grav_tbl = Table(
        [[Paragraph(h, s_th) for h in ["GRAVIDADE", "RESPONSÁVEL", "STATUS"]],
         [Paragraph(v, s_td) for v in [anc.gravidade or "—",
                                        anc.responsavel or "—",
                                        anc.status or "—"]]],
        colWidths=[pw/3, pw/3, pw/3],
    )
    grav_tbl.setStyle(TableStyle([
        ("BACKGROUND",    (0,0),(-1,0), YELLOW),
        ("BOX",           (0,0),(-1,-1), 0.5, BLACK),
        ("INNERGRID",     (0,0),(-1,-1), 0.5, BLACK),
        ("VALIGN",        (0,0),(-1,-1), "MIDDLE"),
        ("TOPPADDING",    (0,0),(-1,-1), 6),
        ("BOTTOMPADDING", (0,0),(-1,-1), 6),
    ]))
    story += [grav_tbl, Spacer(1, 0.5*rcm)]

    # ── 5. REGISTROS FOTOGRÁFICOS (até 6 imagens, 2 por linha) ───
    story.append(yellow_bar("Registros Fotográficos:"))

    imgs_b64 = [x for x in [
        anc.imagem_1, anc.imagem_2, anc.imagem_3,
        anc.imagem_4, anc.imagem_5, anc.imagem_6,
    ] if x]

    img_col_w = pw / 2
    img_h     = 6.5 * rcm

    foto_rows = []
    for i in range(0, max(len(imgs_b64), 2), 2):
        row = []
        for j in range(2):
            idx = i + j
            cell = ""
            if idx < len(imgs_b64):
                try:
                    cell = _fit_img(imgs_b64[idx], img_col_w - 1.0*rcm, img_h)
                except Exception:
                    pass
            row.append(cell)
        foto_rows.append(row)

    foto_tbl = Table(foto_rows, colWidths=[img_col_w, img_col_w])
    foto_tbl.setStyle(TableStyle([
        ("BOX",           (0,0),(-1,-1), 0.5, BLACK),
        ("INNERGRID",     (0,0),(-1,-1), 0.5, BLACK),
        ("ALIGN",         (0,0),(-1,-1), "CENTER"),
        ("VALIGN",        (0,0),(-1,-1), "MIDDLE"),
        ("TOPPADDING",    (0,0),(-1,-1), 8),
        ("BOTTOMPADDING", (0,0),(-1,-1), 8),
        ("MINROWHEIGHT",  (0,0),(-1,-1), img_h),
    ]))
    story += [foto_tbl, Spacer(1, 1.0*rcm)]

    # ── 6. RODAPÉ ─────────────────────────────────────────────────
    story.append(Paragraph(
        f"<b>RESPONSÁVEL PELAS INFORMAÇÕES:</b> {anc.gestor_responsavel or anc.responsavel or '—'}",
        s_foot,
    ))
    story.append(Spacer(1, 0.2*rcm))
    story.append(Paragraph(
        f"<b>CARGO:</b> {anc.tipo or 'Segurança Patrimonial'}",
        s_foot,
    ))

    doc_pdf.build(story)
    buffer.seek(0)
    return buffer


def gerar_docx_de_registro(registro):
    """Gera DOCX a partir dos campos salvos no banco (sem imagens)."""
    doc = Document()
    for sec in doc.sections:
        sec.left_margin  = Cm(1.7)
        sec.right_margin = Cm(1.7)
        sec.top_margin   = Cm(1.5)
        sec.bottom_margin = Cm(1.5)

    logo_path = os.path.join(app.root_path, "static", "logo.png")
    for sec in doc.sections:
        header = sec.header
        for p in header.paragraphs:
            p.text = ""
        avail = sec.page_width - sec.left_margin - sec.right_margin
        ht = header.add_table(rows=1, cols=2, width=avail)
        ht.alignment = WD_TABLE_ALIGNMENT.CENTER
        ht.autofit = False
        ht.columns[0].width = Cm(4.5)
        ht.columns[1].width = Cm(13.1)
        c_logo, c_info = ht.rows[0].cells[0], ht.rows[0].cells[1]
        for c in [c_logo, c_info]:
            set_cell_border(c, color="FFFFFF")
            c.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        p_logo = c_logo.paragraphs[0]
        p_logo.alignment = WD_ALIGN_PARAGRAPH.LEFT
        if os.path.exists(logo_path):
            p_logo.add_run().add_picture(logo_path, width=Cm(3.8))
        else:
            r = p_logo.add_run("DHL")
            r.bold = True; r.font.name = "Arial"
            r.font.size = Pt(16); r.font.color.rgb = RGBColor(212, 5, 17)
        p_info = c_info.paragraphs[0]
        p_info.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        r1 = p_info.add_run("DHL SECURITY\n")
        r1.bold = True; r1.font.name = "Arial"
        r1.font.size = Pt(10); r1.font.color.rgb = RGBColor(212, 5, 17)
        r2 = p_info.add_run("Relatório de Análise Investigativa")
        r2.font.name = "Arial"; r2.font.size = Pt(8)
        r2.font.color.rgb = RGBColor(90, 90, 90)

    doc.add_paragraph()

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after  = Pt(4)
    run = p.add_run("ANÁLISE INVESTIGATIVA")
    run.bold = True; run.font.name = "Arial"
    run.font.size = Pt(17); run.font.color.rgb = RGBColor(31, 41, 55)

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
    p.paragraph_format.space_after  = Pt(6)
    for run in p.runs:
        run.bold = True; run.font.name = "Arial"
        run.font.size = Pt(10); run.font.color.rgb = RGBColor(120, 0, 0)
    doc.add_paragraph()

    add_section_title(doc, "Dados da Operação")
    add_professional_table(doc, [
        ["Nº do Relatório (ID):", str(registro.numero_site or registro.id)],
        ["Empresa:",              registro.empresa              or ""],
        ["Unidade:",              registro.unidade              or ""],
        ["Endereço:",             registro.endereco             or ""],
        ["Classificação do Site:",registro.classificacao        or ""],
        ["Produtos Segmento:",    registro.produtos_segmento    or ""],
        ["Cliente(s):",           registro.clientes             or ""],
    ], col_widths=[6.0, 13.5])

    add_section_title(doc, "Dados do Levantamento")
    add_professional_table(doc, [
        ["Objetivo:",                    registro.objetivo         or ""],
        ["Responsável pelo Levantamento:", registro.responsavel    or ""],
        ["Nome / Função / Data:",         registro.nome_funcao_data or ""],
    ], col_widths=[6.0, 13.5])

    add_section_title(doc, "Descrição do Registro")
    add_text_box(doc, registro.descricao_registro or "")

    if registro.imagens:
        add_section_title(doc, "Evidências")
        for idx, img_obj in enumerate(registro.imagens, start=1):
            try:
                bio = BytesIO(base64.b64decode(img_obj.arquivo))
                bio.seek(0)
                tbl = doc.add_table(rows=1, cols=2)
                tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
                tbl.autofit = False
                img_cell  = tbl.rows[0].cells[0]
                desc_cell = tbl.rows[0].cells[1]
                img_cell.width  = Cm(9.0)
                desc_cell.width = Cm(8.6)
                p_img = img_cell.paragraphs[0]
                p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p_img.add_run().add_picture(bio, width=Cm(8.5))
                p_num = img_cell.add_paragraph(f"Imagem {idx}")
                p_num.alignment = WD_ALIGN_PARAGRAPH.CENTER
                for rn in p_num.runs:
                    rn.font.name = "Arial"; rn.font.size = Pt(8)
                    rn.italic = True; rn.font.color.rgb = RGBColor(100, 100, 100)
                format_cell(img_cell, bg="F5F5F5", align="center")
                desc_cell.text = img_obj.descricao or "Sem descrição"
                format_cell(desc_cell, bg="FFFFFF", align="left")
                p = doc.add_paragraph()
                p.paragraph_format.space_after = Pt(6)
                if idx % 5 == 0 and idx < len(registro.imagens):
                    doc.add_page_break()
            except Exception:
                pass

    add_section_title(doc, "Conclusão")
    add_text_box(doc, registro.conclusao or "")

    add_section_title(doc, "Sugestão")
    add_text_box(doc, registro.sugestao or "")

    doc.add_paragraph()
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    rod = p.add_run(
        f"Código: {registro.codigo or registro.id}"
        + (f" | Criado em: {registro.criado_em.strftime('%d/%m/%Y %H:%M')}"
           if registro.criado_em else "")
    )
    rod.font.name = "Arial"; rod.font.size = Pt(8)
    rod.font.color.rgb = RGBColor(100, 100, 100)

    buf = BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf


def gerar_pdf_analise_bytes(form_data, evidencias_bytes):
    """PDF corporativo DHL da análise investigativa."""
    buffer = BytesIO()

    DHL_RED    = colors.HexColor("#D40511")
    DHL_YELLOW = colors.HexColor("#FFCC00")
    DHL_DARK   = colors.HexColor("#1F2937")
    DHL_MUTED  = colors.HexColor("#6B7280")
    LABEL_BG   = colors.HexColor("#FFF2CC")

    pw = A4[0] - 3.4 * rcm

    doc_pdf = SimpleDocTemplate(
        buffer, pagesize=A4,
        leftMargin=1.7*rcm, rightMargin=1.7*rcm,
        topMargin=2.5*rcm, bottomMargin=2.2*rcm,
    )

    s_title   = ParagraphStyle("s_title",   fontName="Helvetica-Bold", fontSize=17,
                                textColor=DHL_DARK, alignment=TA_CENTER, spaceAfter=4)
    s_section = ParagraphStyle("s_section", fontName="Helvetica-Bold", fontSize=11,
                                textColor=DHL_RED, spaceBefore=14, spaceAfter=6)
    s_body    = ParagraphStyle("s_body",    fontName="Helvetica", fontSize=9,
                                textColor=DHL_DARK, leading=13, spaceAfter=3)
    s_label   = ParagraphStyle("s_label",   fontName="Helvetica-Bold", fontSize=9,
                                textColor=DHL_DARK)
    s_hdr_r   = ParagraphStyle("s_hdr_r",   fontName="Helvetica", fontSize=8,
                                alignment=TA_RIGHT, leading=13)

    def info_table(rows):
        data = [[Paragraph(f"<b>{k}</b>", s_label), Paragraph(str(v or "—"), s_body)]
                for k, v in rows]
        t = Table(data, colWidths=[5.5*rcm, pw - 5.5*rcm])
        t.setStyle(TableStyle([
            ("BACKGROUND",     (0,0), (0,-1),  LABEL_BG),
            ("BACKGROUND",     (1,0), (1,-1),  colors.white),
            ("GRID",           (0,0), (-1,-1), 0.5, colors.HexColor("#D9D9D9")),
            ("VALIGN",         (0,0), (-1,-1), "TOP"),
            ("TOPPADDING",     (0,0), (-1,-1), 5),
            ("BOTTOMPADDING",  (0,0), (-1,-1), 5),
            ("LEFTPADDING",    (0,0), (-1,-1), 8),
        ]))
        return t

    def text_box(text):
        t = Table([[Paragraph(str(text or "—"), s_body)]], colWidths=[pw])
        t.setStyle(TableStyle([
            ("BOX",            (0,0), (-1,-1), 0.5, colors.HexColor("#D9D9D9")),
            ("BACKGROUND",     (0,0), (-1,-1), colors.white),
            ("TOPPADDING",     (0,0), (-1,-1), 8),
            ("BOTTOMPADDING",  (0,0), (-1,-1), 8),
            ("LEFTPADDING",    (0,0), (-1,-1), 10),
            ("RIGHTPADDING",   (0,0), (-1,-1), 10),
        ]))
        return t

    story = []

    # ── cabeçalho ──
    logo_path = os.path.join(app.root_path, "static", "logo.png")
    logo_cell = (RLImage(logo_path, width=4.0*rcm)
                 if os.path.exists(logo_path)
                 else Paragraph('<b><font color="#D40511" size="14">DHL</font></b>',
                                ParagraphStyle("tmp", fontName="Helvetica")))

    hdr = Table(
        [[logo_cell, Paragraph(
            '<font color="#D40511"><b>DHL SECURITY</b></font><br/>'
            '<font color="#6B7280">Relatório de Análise Investigativa</font>',
            s_hdr_r)]],
        colWidths=[pw * 0.35, pw * 0.65]
    )
    hdr.setStyle(TableStyle([
        ("VALIGN",       (0,0), (-1,-1), "MIDDLE"),
        ("LINEBELOW",    (0,0), (-1,-1), 1.5, DHL_RED),
        ("BOTTOMPADDING",(0,0), (-1,-1), 8),
    ]))
    story += [hdr, Spacer(1, 0.4*rcm)]

    # ── título ──
    story.append(Paragraph("ANÁLISE INVESTIGATIVA", s_title))
    story.append(Spacer(1, 0.35*rcm))

    # ── faixa amarela ──
    banner = Table([["DHL SECURITY"]], colWidths=[pw])
    banner.setStyle(TableStyle([
        ("BACKGROUND",     (0,0), (-1,-1), DHL_YELLOW),
        ("TEXTCOLOR",      (0,0), (-1,-1), colors.HexColor("#7A0000")),
        ("FONTNAME",       (0,0), (-1,-1), "Helvetica-Bold"),
        ("FONTSIZE",       (0,0), (-1,-1), 10),
        ("ALIGN",          (0,0), (-1,-1), "CENTER"),
        ("TOPPADDING",     (0,0), (-1,-1), 7),
        ("BOTTOMPADDING",  (0,0), (-1,-1), 7),
        ("BOX",            (0,0), (-1,-1), 1.5, DHL_RED),
    ]))
    story += [banner, Spacer(1, 0.5*rcm)]

    # ── dados da operação ──
    story.append(Paragraph("DADOS DA OPERAÇÃO", s_section))
    story.append(info_table([
        ("Nº do Relatório (ID):",  form_data.get("id_relatorio", "")),
        ("Empresa:",               form_data.get("empresa", "")),
        ("Unidade:",               form_data.get("unidade", "")),
        ("Endereço:",              form_data.get("endereco", "")),
        ("Classificação do Site:", form_data.get("classificacao", "")),
        ("Produtos / Segmento:",   form_data.get("produtos_segmento", "")),
        ("Cliente(s):",            form_data.get("clientes", "")),
    ]))
    story.append(Spacer(1, 0.3*rcm))

    # ── dados do levantamento ──
    story.append(Paragraph("DADOS DO LEVANTAMENTO", s_section))
    story.append(info_table([
        ("Objetivo:",             form_data.get("objetivo", "")),
        ("Responsável:",          form_data.get("responsavel", "")),
        ("Nome / Função / Data:", form_data.get("nome_funcao_data", "")),
    ]))
    story.append(Spacer(1, 0.3*rcm))

    # ── descrição ──
    story.append(Paragraph("DESCRIÇÃO DO REGISTRO", s_section))
    story.append(text_box(form_data.get("descricao_registro", "")))
    story.append(Spacer(1, 0.3*rcm))

    # ── evidências: imagem esquerda | descrição direita | 5 por página ──
    if evidencias_bytes:
        story.append(Paragraph("EVIDÊNCIAS", s_section))
        for idx, (img_bytes, desc) in enumerate(evidencias_bytes, start=1):
            if idx > 1 and (idx - 1) % 5 == 0:
                story.append(PageBreak())
                story.append(Paragraph("EVIDÊNCIAS (continuação)", s_section))
            try:
                rl_img = RLImage(BytesIO(img_bytes), width=7.5*rcm, height=5.5*rcm)
                num_p  = Paragraph(f"<i>Img {idx}</i>",
                                   ParagraphStyle("cap", fontName="Helvetica-Oblique",
                                                  fontSize=7, textColor=DHL_MUTED,
                                                  alignment=TA_CENTER))
                ev = Table(
                    [[rl_img, Paragraph(desc or "Sem descrição", s_body)],
                     [num_p,  ""]],
                    colWidths=[8.0*rcm, pw - 8.0*rcm],
                    rowHeights=[None, 0.5*rcm]
                )
                ev.setStyle(TableStyle([
                    ("BOX",           (0,0), (-1,-1), 0.5, colors.HexColor("#D9D9D9")),
                    ("INNERGRID",     (0,0), (-1,-1), 0.5, colors.HexColor("#D9D9D9")),
                    ("BACKGROUND",    (0,0), (0,-1),  colors.HexColor("#F5F5F5")),
                    ("VALIGN",        (0,0), (-1,-1), "MIDDLE"),
                    ("ALIGN",         (0,0), (0,-1),  "CENTER"),
                    ("SPAN",          (1,0), (1,1)),
                    ("TOPPADDING",    (0,0), (-1,-1), 6),
                    ("BOTTOMPADDING", (0,0), (-1,-1), 6),
                    ("LEFTPADDING",   (0,0), (-1,-1), 6),
                    ("RIGHTPADDING",  (0,0), (-1,-1), 8),
                ]))
                story += [ev, Spacer(1, 0.2*rcm)]
            except Exception:
                pass
        story.append(Spacer(1, 0.2*rcm))

    # ── conclusão ──
    story.append(Paragraph("CONCLUSÃO", s_section))
    story.append(text_box(form_data.get("conclusao", "")))
    story.append(Spacer(1, 0.3*rcm))

    # ── sugestão ──
    story.append(Paragraph("SUGESTÃO", s_section))
    story.append(text_box(form_data.get("sugestao", "")))

    # ── rodapé ──
    id_rel = form_data.get("id_relatorio", "")

    def footer(canvas, doc):
        canvas.saveState()
        canvas.setStrokeColor(DHL_RED)
        canvas.setLineWidth(0.8)
        canvas.line(1.7*rcm, 1.4*rcm, A4[0] - 1.7*rcm, 1.4*rcm)
        canvas.setFont("Helvetica", 7)
        canvas.setFillColor(DHL_MUTED)
        canvas.drawString(1.7*rcm, 1.0*rcm,
                          f"DHL Security — Análise Investigativa{' | ID: ' + id_rel if id_rel else ''}")
        canvas.drawRightString(A4[0] - 1.7*rcm, 1.0*rcm, f"Página {doc.page}")
        canvas.restoreState()

    doc_pdf.build(story, onFirstPage=footer, onLaterPages=footer)
    buffer.seek(0)
    return buffer


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
    usuario = Usuario.query.get(session.get("user_id"))
    site_atual = None
    if usuario and usuario.site:
        site_atual = SiteCompleto.query.filter_by(nome_do_site=usuario.site).first()
    proximo_numero = AnaliseInvestigativa.query.filter_by(site=usuario.site if usuario else None).count() + 1
    return render_template("analise_investigativa.html", site_atual=site_atual, proximo_numero=proximo_numero)


@app.route("/analises")
@login_required
def analises():
    usuario = Usuario.query.get(session.get("user_id"))
    is_admin = usuario and (usuario.perfil or "").upper() == "ADMIN"
    site_usuario = usuario.site if usuario else None

    if is_admin:
        registros = AnaliseInvestigativa.query.order_by(AnaliseInvestigativa.id.desc()).all()
    else:
        registros = AnaliseInvestigativa.query.filter_by(site=site_usuario).order_by(AnaliseInvestigativa.id.desc()).all()

    resumo = {
        "total": len(registros),
        "sites": len(set(r.site for r in registros if r.site)),
    }
    return render_template("analises.html", registros=registros, resumo=resumo, is_admin=is_admin)


@app.route("/analises/excluir/<int:analise_id>", methods=["POST"])
@login_required
@perfil_required("ADMIN")
def excluir_analise(analise_id):
    registro = AnaliseInvestigativa.query.get_or_404(analise_id)
    db.session.delete(registro)
    db.session.commit()
    flash("Análise excluída com sucesso.", "success")
    return redirect(url_for("analises"))


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


def add_text_box(doc, text="\n\n\n"):
    table = doc.add_table(rows=1, cols=1)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell = table.rows[0].cells[0]
    cell.text = text or "\n\n\n"
    format_cell(cell, bold=False, bg="FFFFFF")
    return table

@app.route("/gerar_docx", methods=["POST"])
@login_required
def gerar_analise_investigativa():
    f = request.form

    # ===== Salva no banco =====
    site_val = (f.get("site") or "").strip()
    _codigo_analise, _seq_analise = gerar_codigo(AnaliseInvestigativa, site_val)
    nova_analise = AnaliseInvestigativa(
        codigo=_codigo_analise,
        numero_site=_seq_analise,
        site=site_val,
        id_relatorio=(f.get("id_relatorio") or "").strip(),
        empresa=(f.get("empresa") or "").strip(),
        unidade=(f.get("unidade") or "").strip(),
        endereco=(f.get("endereco") or "").strip(),
        classificacao=(f.get("classificacao") or "").strip(),
        produtos_segmento=(f.get("produtos_segmento") or "").strip(),
        clientes=(f.get("clientes") or "").strip(),
        objetivo=(f.get("objetivo") or "").strip(),
        responsavel=(f.get("responsavel") or "").strip(),
        nome_funcao_data=(f.get("nome_funcao_data") or "").strip(),
        descricao_registro=(f.get("descricao_registro") or "").strip(),
        conclusao=(f.get("conclusao") or "").strip(),
        sugestao=(f.get("sugestao") or "").strip(),
        criado_por=session.get("user_nome"),
    )
    db.session.add(nova_analise)
    db.session.commit()

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
    # lê os bytes agora — streams só podem ser lidos uma vez
    evidencias_bytes = [
        (fimg.read(), desc)
        for fimg, desc in zip(files, descricoes) if fimg and fimg.filename
    ]
    evidencias = evidencias_bytes  # alias usado pelo bloco DOCX

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

        header_table.columns[0].width = Cm(4.5)
        header_table.columns[1].width = Cm(13.1)

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

        r2 = p_info.add_run("Relatório de Análise Investigativa")
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
    # EVIDÊNCIAS  (imagem esquerda | descrição direita)
    # ==========================================================
    if evidencias:
        add_section_title(doc, "Evidências")

        for idx, (img_bytes, desc) in enumerate(evidencias, start=1):
            bio = BytesIO(img_bytes)
            bio.seek(0)

            tbl = doc.add_table(rows=1, cols=2)
            tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
            tbl.autofit = False

            img_cell  = tbl.rows[0].cells[0]
            desc_cell = tbl.rows[0].cells[1]

            img_cell.width  = Cm(7.0)
            desc_cell.width = Cm(10.5)

            # célula esquerda — imagem
            p_img = img_cell.paragraphs[0]
            p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_img.add_run().add_picture(bio, width=Cm(6.5))

            p_num = img_cell.add_paragraph(f"Imagem {idx}")
            p_num.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for rn in p_num.runs:
                rn.font.name = "Arial"
                rn.font.size = Pt(8)
                rn.italic = True
                rn.font.color.rgb = RGBColor(100, 100, 100)

            format_cell(img_cell, bg="F5F5F5", align="center")

            # célula direita — descrição
            desc_cell.text = desc or "Sem descrição"
            format_cell(desc_cell, bg="FFFFFF", align="left")

            p = doc.add_paragraph()
            p.paragraph_format.space_after = Pt(6)

            # quebra de página a cada 5 imagens
            if idx % 5 == 0 and idx < len(evidencias):
                doc.add_page_break()

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
    base_name = f"Analise_Investigativa_{nova_analise.codigo or nova_analise.id}"

    # ==========================================================
    # GERA DOCX (reservado para OneDrive)
    # ==========================================================
    docx_buf = BytesIO()
    doc.save(docx_buf)

    # ==========================================================
    # GERA PDF E CACHEIA PARA DOWNLOAD
    # ==========================================================
    nova_analise.docx_arquivo = base64.b64encode(docx_buf.getvalue()).decode("utf-8")
    db.session.commit()

    flash("Análise salva com sucesso!", "success")
    return redirect(url_for("confirmar_analise", analise_id=nova_analise.id))




@app.route("/analises/confirmar/<int:analise_id>")
@login_required
def confirmar_analise(analise_id):
    registro = AnaliseInvestigativa.query.get_or_404(analise_id)
    tem_arquivo = bool(registro.docx_arquivo)
    return render_template("confirmar_analise.html", registro=registro, tem_pdf=tem_arquivo)


@app.route("/analises/download/<int:analise_id>")
@login_required
def download_analise(analise_id):
    registro = AnaliseInvestigativa.query.get_or_404(analise_id)
    if not registro.docx_arquivo:
        flash("Arquivo não disponível para esta análise.", "warning")
        return redirect(url_for("analises"))

    filename  = f"Analise_Investigativa_{registro.codigo or registro.id}.docx"
    buf = BytesIO(base64.b64decode(registro.docx_arquivo))
    buf.seek(0)
    return send_file(buf, as_attachment=True, download_name=filename,
                     mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document")


# ==========================
# ANC - Avisos de Não Conformidade
# ==========================
@app.route("/anc", methods=["GET", "POST"])
@login_required
def anc():
    usuario = Usuario.query.get(session.get("user_id"))
    is_admin = usuario and (usuario.perfil or "").upper() == "ADMIN"
    site_usuario = usuario.site if usuario else None

    registro_edicao = None
    modo_edicao = False
    editar_id = request.args.get("editar", type=int)
    if editar_id:
        registro_edicao = ANC.query.get_or_404(editar_id)
        modo_edicao = True

    if request.method == "POST":
        f = request.form
        anc_id = f.get("anc_id", type=int)

        data_nc           = (f.get("data_nc")           or "").strip()
        hora_nc           = (f.get("hora_nc")           or "").strip()
        site_val          = (f.get("site")              or site_usuario or "").strip()
        setor             = (f.get("setor")             or "").strip()
        tipo_ocorrencia   = (f.get("tipo_ocorrencia")   or "").strip()
        gravidade         = (f.get("gravidade")         or "").strip().upper()
        natureza          = (f.get("natureza")          or "").strip()
        responsavel       = (f.get("responsavel")       or "").strip()
        gestor_responsavel= (f.get("gestor_responsavel")or "").strip()
        local_val         = (f.get("local")             or "").strip()
        envolvido         = (f.get("envolvido")         or "").strip()
        tipo              = (f.get("tipo")              or "").strip()
        turno             = (f.get("turno")             or "").strip()
        status            = (f.get("status")            or "ABERTO").strip().upper()
        descricao         = (f.get("descricao")         or "").strip()
        inicio_investigacao = (f.get("inicio_investigacao") or "").strip() or None
        fim_investigacao    = (f.get("fim_investigacao")    or "").strip() or None

        if not data_nc or not hora_nc or not setor or not tipo_ocorrencia or not gravidade \
                or not natureza or not responsavel or not gestor_responsavel \
                or not local_val or not tipo or not turno or not descricao:
            flash("Preencha todos os campos obrigatórios.", "warning")
            return redirect(url_for("anc", editar=anc_id) if anc_id else url_for("anc"))

        imgs = []
        for i in range(1, 7):
            field = request.files.get(f"imagem_{i}")
            if field and field.filename:
                b64, _ = arquivo_para_base64(field, EXTENSOES_PERMITIDAS_IMAGEM)
                imgs.append(b64)
            else:
                imgs.append(None)

        if anc_id:
            if not is_admin:
                flash("Apenas administradores podem editar ANCs.", "danger")
                return redirect(url_for("anc"))
            reg = ANC.query.get_or_404(anc_id)
            reg.data_nc = data_nc; reg.hora_nc = hora_nc; reg.site = site_val
            reg.setor = setor; reg.tipo_ocorrencia = tipo_ocorrencia
            reg.gravidade = gravidade; reg.natureza = natureza
            reg.responsavel = responsavel; reg.gestor_responsavel = gestor_responsavel
            reg.local = local_val; reg.envolvido = envolvido
            reg.tipo = tipo; reg.turno = turno; reg.status = status
            reg.descricao = descricao
            reg.inicio_investigacao = inicio_investigacao
            reg.fim_investigacao = fim_investigacao
            for i, b64 in enumerate(imgs, start=1):
                if b64:
                    setattr(reg, f"imagem_{i}", b64)
            for i in range(1, 7):
                if f.get(f"remover_img_{i}"):
                    setattr(reg, f"imagem_{i}", None)
            db.session.commit()
            flash("ANC atualizado com sucesso.", "success")
            return redirect(url_for("anc"))

        _num_anc, _seq_anc = gerar_numero_anc(site_val)
        novo = ANC(
            numero_anc=_num_anc, numero_site=_seq_anc,
            data_nc=data_nc, hora_nc=hora_nc, site=site_val,
            setor=setor, tipo_ocorrencia=tipo_ocorrencia,
            gravidade=gravidade, natureza=natureza,
            responsavel=responsavel, gestor_responsavel=gestor_responsavel,
            local=local_val, envolvido=envolvido,
            tipo=tipo, turno=turno, status=status,
            descricao=descricao,
            inicio_investigacao=inicio_investigacao,
            fim_investigacao=fim_investigacao,
            imagem_1=imgs[0], imagem_2=imgs[1],
            imagem_3=imgs[2], imagem_4=imgs[3],
            imagem_5=imgs[4], imagem_6=imgs[5],
            criado_por=session.get("user_nome"),
        )
        db.session.add(novo)
        db.session.commit()
        flash("ANC registrado com sucesso.", "success")
        return redirect(url_for("anc"))

    if is_admin:
        query = ANC.query.order_by(ANC.id.desc())
    else:
        query = ANC.query.filter_by(site=site_usuario).order_by(ANC.id.desc())

    registros, filtros = aplicar_filtros_anc(query)

    resumo = {
        "total":     len(registros),
        "abertos":   len([r for r in registros if (r.status or "").upper() == "ABERTO"]),
        "andamento": len([r for r in registros if (r.status or "").upper() == "EM ANDAMENTO"]),
        "concluidos":len([r for r in registros if (r.status or "").upper() == "CONCLUÍDO"]),
        "criticos":  len([r for r in registros if (r.gravidade or "").upper() == "CRÍTICA"]),
    }

    return render_template(
        "anc.html",
        registros=registros, resumo=resumo, filtros=filtros,
        is_admin=is_admin, site_usuario=site_usuario,
        modo_edicao=modo_edicao, registro_edicao=registro_edicao,
        agora=datetime.now().strftime("%Y-%m-%dT%H:%M"),
        hora_atual=datetime.now().strftime("%H:%M"),
    )


@app.route("/anc/<int:anc_id>/status", methods=["POST"])
@login_required
def anc_status(anc_id):
    reg = ANC.query.get_or_404(anc_id)
    novo = (request.form.get("status") or "").strip().upper()
    if novo in {"ABERTO", "EM ANDAMENTO", "CONCLUÍDO"}:
        reg.status = novo
        db.session.commit()
        flash("Status atualizado.", "success")
    else:
        flash("Status inválido.", "danger")
    return redirect(url_for("anc"))


@app.route("/anc/<int:anc_id>/excluir", methods=["POST"])
@login_required
@perfil_required("ADMIN")
def excluir_anc(anc_id):
    reg = ANC.query.get_or_404(anc_id)
    db.session.delete(reg)
    db.session.commit()
    flash("ANC excluído com sucesso.", "success")
    return redirect(url_for("anc"))


@app.route("/exportar/anc/excel")
@login_required
def exportar_anc_excel():
    usuario = Usuario.query.get(session.get("user_id"))
    is_admin = usuario and (usuario.perfil or "").upper() == "ADMIN"
    site_usuario = usuario.site if usuario else None

    query = ANC.query.order_by(ANC.id.desc()) if is_admin \
        else ANC.query.filter_by(site=site_usuario).order_by(ANC.id.desc())
    registros, _ = aplicar_filtros_anc(query)

    wb = Workbook()
    ws = wb.active
    ws.title = "ANCs"
    headers = ["Nº ANC","Data","Hora","Site","Setor","Tipo Ocorrência","Gravidade",
               "Natureza","Responsável","Gestor","Local","Envolvido","Tipo","Turno",
               "Status","Descrição","Criado por"]
    ws.append(headers)
    fill = PatternFill("solid", fgColor="FFCC00")
    font_bold = Font(bold=True)
    for col in range(1, len(headers) + 1):
        ws.cell(row=1, column=col).fill = fill
        ws.cell(row=1, column=col).font = font_bold
    for r in registros:
        ws.append([r.numero_anc, r.data_nc, r.hora_nc, r.site, r.setor,
                   r.tipo_ocorrencia, r.gravidade, r.natureza, r.responsavel,
                   r.gestor_responsavel, r.local, r.envolvido, r.tipo, r.turno,
                   r.status, r.descricao, r.criado_por])

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return send_file(output, as_attachment=True, download_name="controle_anc.xlsx",
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.route("/exportar/anc/<int:anc_id>/pdf")
@login_required
def exportar_anc_pdf(anc_id):
    reg = ANC.query.get_or_404(anc_id)
    buf = gerar_pdf_anc_bytes(reg)
    return send_file(buf, as_attachment=True,
                     download_name=f"ANC_{reg.numero_anc or anc_id}.pdf",
                     mimetype="application/pdf")


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

    usuario = Usuario.query.get(session.get("user_id"))
    site_usuario = usuario.site if usuario else None

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
        site = (request.form.get("site") or site_usuario or "").strip()

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
            registro.site = site
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

        _codigo, _seq = gerar_codigo(Ocorrencia, site)
        nova = Ocorrencia(
            codigo=_codigo,
            numero_site=_seq,
            data_hora=data_hora,
            hora_ocorrencia=hora_ocorrencia,
            natureza=natureza,
            descricao=descricao,
            site=site,
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

    is_admin = usuario and (usuario.perfil or "").upper() == "ADMIN"
    if is_admin:
        query = Ocorrencia.query.order_by(Ocorrencia.id.desc())
    else:
        query = Ocorrencia.query.filter_by(site=site_usuario).order_by(Ocorrencia.id.desc())
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
        filtros=filtros,
        site_usuario=site_usuario
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
@perfil_required("ADMIN", "USUARIO")
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