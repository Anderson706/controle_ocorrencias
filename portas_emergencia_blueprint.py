# portas_emergencia_blueprint.py — Portas de Emergência integrado ao CCTV Control Panel
# Registrado em app.py com url_prefix='/portas-emergencia'
# Autenticação usa a sessão do CCTV (session["user_id"])

from flask import (
    Blueprint, render_template, request, redirect,
    url_for, flash, session, send_file, current_app
)
from functools import wraps
from datetime import datetime
from io import BytesIO
import os

from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image as RLImage
)
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.units import cm

pe_bp = Blueprint('pe', __name__, template_folder='templates')

# ── Globals preenchidos via setup_pe ─────────────────────────────────────────
_db                      = None
PortaEmergencia          = None
ChecklistPortaEmergencia = None
DisparoAlarme            = None
BotaoPanico              = None
_get_setores_fn          = None   # injetado pelo app.py
_get_locais_fn           = None   # injetado pelo app.py


# ── Inicialização ─────────────────────────────────────────────────────────────
def setup_pe(db, get_setores=None, get_locais=None):
    global _db, PortaEmergencia, ChecklistPortaEmergencia, DisparoAlarme, BotaoPanico, _get_setores_fn, _get_locais_fn
    _db = db
    _get_setores_fn = get_setores
    _get_locais_fn  = get_locais

    class _PortaEmergencia(db.Model):
        __tablename__ = "PORTAS_EMERGENCIA"
        id          = db.Column(db.Integer, db.Identity(start=1), primary_key=True)
        codigo      = db.Column(db.String(50),  nullable=False, unique=True)
        localizacao = db.Column(db.String(160), nullable=False)
        setor       = db.Column(db.String(120), nullable=True)
        rota_fuga   = db.Column(db.String(120), nullable=True)
        responsavel = db.Column(db.String(120), nullable=True)
        observacao  = db.Column(db.Text,        nullable=True)
        ativo       = db.Column(db.Boolean,     default=True)
        site        = db.Column(db.String(128), nullable=True)
        criado_por  = db.Column(db.String(120), nullable=True)
        criado_em   = db.Column(db.DateTime,    default=datetime.utcnow)
        checklists  = db.relationship(
            "_ChecklistPortaEmergencia",
            backref="porta", lazy=True, cascade="all, delete-orphan",
            primaryjoin="_PortaEmergencia.id == _ChecklistPortaEmergencia.porta_id",
            foreign_keys="_ChecklistPortaEmergencia.porta_id",
        )

    class _ChecklistPortaEmergencia(db.Model):
        __tablename__ = "CHECKLISTS_PORTAS_EMERGENCIA"
        id                  = db.Column(db.Integer, db.Identity(start=1), primary_key=True)
        porta_id            = db.Column(db.Integer, db.ForeignKey("PORTAS_EMERGENCIA.id"), nullable=False)
        data_checklist      = db.Column(db.DateTime, default=datetime.utcnow)
        inspetor            = db.Column(db.String(120), nullable=False)
        turno               = db.Column(db.String(30),  nullable=True)
        porta_desobstruida  = db.Column(db.Boolean, default=False)
        abre_normalmente    = db.Column(db.Boolean, default=False)
        sinalizacao_ok      = db.Column(db.Boolean, default=False)
        iluminacao_ok       = db.Column(db.Boolean, default=False)
        alarme_ok           = db.Column(db.Boolean, default=False)
        status              = db.Column(db.String(30), nullable=False, default="PENDENTE")
        observacao          = db.Column(db.Text, nullable=True)
        criado_em           = db.Column(db.DateTime, default=datetime.utcnow)
        # Tratativa de não conformidade
        conclusao           = db.Column(db.Text,        nullable=True)
        data_conclusao      = db.Column(db.DateTime,    nullable=True)
        concluido_por       = db.Column(db.String(120), nullable=True)

    class _DisparoAlarme(db.Model):
        __tablename__ = "DISPAROS_ALARME"
        id                    = db.Column(db.Integer, db.Identity(start=1), primary_key=True)
        data_registro         = db.Column(db.DateTime, default=datetime.utcnow)
        data_disparo          = db.Column(db.String(20),  nullable=False)
        hora_disparo          = db.Column(db.String(10),  nullable=False)
        hora_desativado       = db.Column(db.String(10),  nullable=True)
        hora_ativado          = db.Column(db.String(10),  nullable=True)
        contato_monitoramento = db.Column(db.String(150), nullable=True)
        localizacao           = db.Column(db.String(150), nullable=False)
        setor                 = db.Column(db.String(100), nullable=True)
        tipo_alarme           = db.Column(db.String(100), nullable=False)
        motivo                = db.Column(db.String(150), nullable=True)
        responsavel           = db.Column(db.String(120), nullable=True)
        turno                 = db.Column(db.String(50),  nullable=True)
        houve_evacuacao       = db.Column(db.Boolean, default=False)
        acionado_bombeiro     = db.Column(db.Boolean, default=False)
        acionado_seguranca    = db.Column(db.Boolean, default=False)
        status                = db.Column(db.String(50), default="EM ANÁLISE")
        observacao            = db.Column(db.Text,    nullable=True)
        site                  = db.Column(db.String(128), nullable=True)
        criado_por            = db.Column(db.String(120), nullable=True)

    class _BotaoPanico(db.Model):
        __tablename__ = "BOTOES_PANICO"
        id                   = db.Column(db.Integer, db.Identity(start=1), primary_key=True)
        data_registro        = db.Column(db.DateTime, default=datetime.utcnow)
        codigo               = db.Column(db.String(50),  nullable=False)
        localizacao          = db.Column(db.String(150), nullable=False)
        setor                = db.Column(db.String(100), nullable=True)
        tipo                 = db.Column(db.String(80),  nullable=True)
        responsavel          = db.Column(db.String(120), nullable=True)
        turno                = db.Column(db.String(50),  nullable=True)
        testado              = db.Column(db.Boolean, default=False)
        sinal_recebido       = db.Column(db.Boolean, default=False)
        comunicacao_cftv     = db.Column(db.Boolean, default=False)
        necessita_manutencao = db.Column(db.Boolean, default=False)
        hora_teste           = db.Column(db.String(10),  nullable=True)
        hora_retorno         = db.Column(db.String(10),  nullable=True)
        agente_cftv          = db.Column(db.String(120), nullable=True)
        status               = db.Column(db.String(50), default="EM ANÁLISE")
        observacao           = db.Column(db.Text, nullable=True)
        site                 = db.Column(db.String(128), nullable=True)
        criado_por           = db.Column(db.String(120), nullable=True)

    PortaEmergencia          = _PortaEmergencia
    ChecklistPortaEmergencia = _ChecklistPortaEmergencia
    DisparoAlarme            = _DisparoAlarme
    BotaoPanico              = _BotaoPanico


# ── Auth helpers ──────────────────────────────────────────────────────────────
def _login_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if not session.get("user_id"):
            return redirect(url_for("login"))
        return f(*args, **kwargs)
    return decorated


def _gestor_required(f):
    """Restringe a rota a ADMIN, GESTOR e KEYUSER."""
    @wraps(f)
    def decorated(*args, **kwargs):
        if not session.get("user_id"):
            return redirect(url_for("login"))
        if not _is_gestor_keyuser():
            flash("Acesso restrito a Gestores e Key Users.", "danger")
            return redirect(url_for("pe.painel"))
        return f(*args, **kwargs)
    return decorated


def _is_admin():
    return (session.get("user_perfil") or "").upper() == "ADMIN"


def _sites_autorizados():
    """Retorna lista de sites que o usuário logado pode visualizar.
    ADMIN → [] (sem filtro — vê tudo); demais → sites vinculados ou site principal."""
    if _is_admin():
        return []
    uid = session.get("user_id")
    if uid:
        try:
            from sqlalchemy import text as _text
            rows = _db.session.execute(
                _text("SELECT SITE_NOME FROM USUARIO_SITES WHERE USUARIO_ID = :uid"),
                {"uid": uid}
            ).fetchall()
            if rows:
                return [r[0] for r in rows]
        except Exception:
            pass
    site = _site_usuario()
    return [site] if site else []


def _is_gestor_keyuser():
    return (session.get("user_perfil") or "").upper() in ("ADMIN", "GESTOR", "KEYUSER")


def _site_usuario():
    return session.get("user_site") or None


def _setores_do_site():
    """Retorna lista de setores do site do usuário logado."""
    if _get_setores_fn:
        return _get_setores_fn(_site_usuario())
    return []


def _locais_do_site():
    """Retorna lista de locais do site do usuário logado."""
    if _get_locais_fn:
        return _get_locais_fn(_site_usuario())
    return []


def _calcular_status_checklist(form):
    campos = ["porta_desobstruida", "abre_normalmente", "sinalizacao_ok",
              "iluminacao_ok", "alarme_ok"]
    if all(form.get(c) == "on" for c in campos):
        return "CONFORME"
    return "NÃO CONFORME"


def _bool(nome):
    return request.form.get(nome) == "on"


def _logo_path():
    return os.path.join(current_app.root_path, "static", "logo.png")


# ═══════════════════════════════════════════════════════════════════════════════
# PAINEL (index)
# ═══════════════════════════════════════════════════════════════════════════════
@pe_bp.route("/")
@_login_required
def painel():
    site    = _site_usuario()
    is_admin = _is_admin()
    _sites  = _sites_autorizados()

    q = PortaEmergencia.query
    if _sites:
        if len(_sites) == 1:
            q = q.filter(PortaEmergencia.site == _sites[0])
        else:
            q = q.filter(PortaEmergencia.site.in_(_sites))

    portas = q.order_by(PortaEmergencia.id.desc()).all()

    total         = len(portas)
    ativas        = sum(1 for p in portas if p.ativo)
    checks        = ChecklistPortaEmergencia.query.count()
    nao_conformes = ChecklistPortaEmergencia.query.filter_by(status="NÃO CONFORME").count()

    return render_template(
        "pe_painel.html",
        portas=portas, total=total, ativas=ativas,
        checks=checks, nao_conformes=nao_conformes,
        is_admin=is_admin,
        is_gestor=_is_gestor_keyuser(),
    )


# ═══════════════════════════════════════════════════════════════════════════════
# NOVA PORTA
# ═══════════════════════════════════════════════════════════════════════════════
@pe_bp.route("/porta/nova", methods=["GET", "POST"])
@_gestor_required
def nova_porta():
    if request.method == "POST":
        codigo      = (request.form.get("codigo")      or "").strip().upper()
        localizacao = (request.form.get("localizacao") or "").strip()

        if not codigo or not localizacao:
            flash("Informe pelo menos o código e a localização da porta.", "warning")
            return redirect(url_for("pe.nova_porta"))

        if PortaEmergencia.query.filter_by(codigo=codigo).first():
            flash("Já existe uma porta cadastrada com esse código.", "danger")
            return redirect(url_for("pe.nova_porta"))

        porta = PortaEmergencia(
            codigo=codigo,
            localizacao=localizacao,
            setor=request.form.get("setor"),
            rota_fuga=request.form.get("rota_fuga"),
            responsavel=session.get("user_nome"),
            observacao=request.form.get("observacao"),
            ativo=True,
            site=_site_usuario(),
            criado_por=session.get("user_nome"),
        )
        _db.session.add(porta)
        _db.session.commit()
        flash("Porta de emergência cadastrada com sucesso.", "success")
        return redirect(url_for("pe.painel"))

    return render_template("pe_nova_porta.html",
                           setores=_setores_do_site(),
                           user_nome=session.get("user_nome", ""),
                           user_site=_site_usuario() or "")


# ═══════════════════════════════════════════════════════════════════════════════
# IMPORTAÇÃO EM LOTE
# ═══════════════════════════════════════════════════════════════════════════════
@pe_bp.route("/porta/lote", methods=["GET", "POST"])
@_gestor_required
def portas_lote():
    if request.method == "POST":
        texto = request.form.get("dados_lote") or ""
        linhas = [l.strip() for l in texto.splitlines() if l.strip()]
        criadas = ignoradas = 0

        for linha in linhas:
            partes = [p.strip() for p in linha.split(";")]
            if len(partes) < 2:
                ignoradas += 1
                continue
            codigo      = partes[0].upper()
            localizacao = partes[1]
            setor       = partes[2] if len(partes) > 2 else ""
            rota_fuga   = partes[3] if len(partes) > 3 else ""
            responsavel = partes[4] if len(partes) > 4 else ""
            observacao  = partes[5] if len(partes) > 5 else ""

            if not codigo or not localizacao:
                ignoradas += 1
                continue
            if PortaEmergencia.query.filter_by(codigo=codigo).first():
                ignoradas += 1
                continue

            _db.session.add(PortaEmergencia(
                codigo=codigo, localizacao=localizacao, setor=setor,
                rota_fuga=rota_fuga, responsavel=responsavel, observacao=observacao,
                ativo=True, site=_site_usuario(), criado_por=session.get("user_nome"),
            ))
            criadas += 1

        _db.session.commit()
        flash(f"Importação concluída. Criadas: {criadas}. Ignoradas: {ignoradas}.", "success")
        return redirect(url_for("pe.painel"))

    return render_template("pe_portas_lote.html")


# ═══════════════════════════════════════════════════════════════════════════════
# CHECKLIST
# ═══════════════════════════════════════════════════════════════════════════════
@pe_bp.route("/porta/<int:porta_id>/checklist", methods=["GET", "POST"])
@_login_required
def checklist(porta_id):
    porta = PortaEmergencia.query.get_or_404(porta_id)

    if request.method == "POST":
        inspetor = request.form.get("inspetor") or session.get("user_nome", "Não informado")
        ck = ChecklistPortaEmergencia(
            porta_id=porta.id,
            inspetor=inspetor,
            turno=request.form.get("turno"),
            porta_desobstruida=_bool("porta_desobstruida"),
            abre_normalmente=_bool("abre_normalmente"),
            sinalizacao_ok=_bool("sinalizacao_ok"),
            iluminacao_ok=_bool("iluminacao_ok"),
            alarme_ok=_bool("alarme_ok"),
            status=_calcular_status_checklist(request.form),
            observacao=request.form.get("observacao"),
        )
        _db.session.add(ck)
        _db.session.commit()
        flash("Checklist registrado com sucesso.", "success")
        return redirect(url_for("pe.historico", porta_id=porta.id))

    return render_template("pe_checklist.html", porta=porta,
                           user_nome=session.get("user_nome", ""))


# ═══════════════════════════════════════════════════════════════════════════════
# HISTÓRICO DA PORTA
# ═══════════════════════════════════════════════════════════════════════════════
@pe_bp.route("/porta/<int:porta_id>/historico")
@_login_required
def historico(porta_id):
    porta = PortaEmergencia.query.get_or_404(porta_id)
    checklists = (ChecklistPortaEmergencia.query
                  .filter_by(porta_id=porta.id)
                  .order_by(ChecklistPortaEmergencia.id.desc())
                  .all())
    return render_template("pe_historico.html", porta=porta, checklists=checklists,
                           is_gestor=_is_gestor_keyuser())


# ═══════════════════════════════════════════════════════════════════════════════
# EXCLUIR PORTA
# ═══════════════════════════════════════════════════════════════════════════════
@pe_bp.route("/porta/<int:porta_id>/excluir", methods=["POST"])
@_login_required
def excluir_porta(porta_id):
    if not _is_admin():
        flash("Apenas administradores podem excluir portas.", "danger")
        return redirect(url_for("pe.painel"))
    try:
        porta = PortaEmergencia.query.get_or_404(porta_id)
        _db.session.delete(porta)
        _db.session.commit()
        flash("Porta excluída com sucesso.", "success")
    except Exception as e:
        _db.session.rollback()
        flash(f"Erro ao excluir porta: {e}", "danger")
    return redirect(url_for("pe.painel"))


# ═══════════════════════════════════════════════════════════════════════════════
# PDF — Checklist individual
# ═══════════════════════════════════════════════════════════════════════════════
@pe_bp.route("/checklist/<int:checklist_id>/pdf")
@_login_required
def pdf_checklist(checklist_id):
    ck = ChecklistPortaEmergencia.query.get_or_404(checklist_id)
    porta = PortaEmergencia.query.get_or_404(ck.porta_id)

    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4,
                            rightMargin=0.9*cm, leftMargin=0.9*cm,
                            topMargin=0.8*cm, bottomMargin=0.8*cm)
    styles = getSampleStyleSheet()

    tit_style = ParagraphStyle("Tit", parent=styles["Heading1"], fontName="Helvetica-Bold",
                               fontSize=16, alignment=TA_CENTER,
                               textColor=colors.HexColor("#111827"), spaceAfter=4)
    sub_style = ParagraphStyle("Sub", parent=styles["BodyText"], fontSize=8,
                               alignment=TA_CENTER, textColor=colors.HexColor("#6b7280"),
                               spaceAfter=12)
    lbl_style = ParagraphStyle("Lbl", parent=styles["BodyText"], fontName="Helvetica-Bold",
                               fontSize=8, textColor=colors.HexColor("#374151"))
    cel_style = ParagraphStyle("Cel", parent=styles["BodyText"],
                               fontSize=8, textColor=colors.HexColor("#111827"))
    sec_style = ParagraphStyle("Sec", parent=styles["BodyText"], fontName="Helvetica-Bold",
                               fontSize=9, textColor=colors.HexColor("#d40511"),
                               spaceBefore=10, spaceAfter=4)

    elementos = []
    try:
        logo = RLImage(_logo_path(), width=3.5*cm, height=1.1*cm)
        logo.hAlign = "CENTER"
        elementos += [logo, Spacer(1, 6)]
    except Exception:
        pass

    elementos.append(Paragraph("CHECKLIST DE PORTA DE EMERGÊNCIA", tit_style))
    elementos.append(Paragraph("DHL SECURITY • CONTROLE OPERACIONAL", sub_style))

    # Dados do registro
    ok_count = sum(1 for v in [ck.porta_desobstruida, ck.abre_normalmente,
                                ck.sinalizacao_ok, ck.iluminacao_ok, ck.alarme_ok] if v)
    status_cor = colors.HexColor("#146c43") if ck.status == "CONFORME" else colors.HexColor("#d40511")

    info = [
        [Paragraph("PORTA",        lbl_style), Paragraph(f"{porta.codigo} — {porta.localizacao}", cel_style)],
        [Paragraph("SETOR",        lbl_style), Paragraph(porta.setor or "—", cel_style)],
        [Paragraph("DATA",         lbl_style), Paragraph(ck.data_checklist.strftime('%d/%m/%Y %H:%M') if ck.data_checklist else "—", cel_style)],
        [Paragraph("INSPETOR",     lbl_style), Paragraph(ck.inspetor or "—", cel_style)],
        [Paragraph("TURNO",        lbl_style), Paragraph(ck.turno or "—", cel_style)],
        [Paragraph("STATUS",       lbl_style), Paragraph(ck.status or "—", ParagraphStyle("S", parent=cel_style, textColor=status_cor, fontName="Helvetica-Bold"))],
        [Paragraph("ITENS OK",     lbl_style), Paragraph(f"{ok_count}/5", cel_style)],
    ]
    t_info = Table(info, colWidths=[4.0*cm, 13.4*cm])
    t_info.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (0,-1), colors.HexColor("#f3f4f6")),
        ("GRID",       (0,0), (-1,-1), 0.35, colors.HexColor("#d1d5db")),
        ("VALIGN",     (0,0), (-1,-1), "MIDDLE"),
        ("TOPPADDING",    (0,0),(-1,-1), 5),
        ("BOTTOMPADDING", (0,0),(-1,-1), 5),
        ("LEFTPADDING",   (0,0),(-1,-1), 7),
        ("RIGHTPADDING",  (0,0),(-1,-1), 7),
    ]))
    elementos += [t_info, Spacer(1, 8)]

    # Itens de verificação
    elementos.append(Paragraph("ITENS DE VERIFICAÇÃO", sec_style))
    _sim = "✔  SIM"
    _nao = "✘  NÃO"
    itens = [
        ["ITEM", "RESULTADO"],
        ["Porta desobstruída",                   _sim if ck.porta_desobstruida else _nao],
        ["Abre normalmente (sem travamento)",     _sim if ck.abre_normalmente   else _nao],
        ["Sinalização visível e em bom estado",  _sim if ck.sinalizacao_ok      else _nao],
        ["Iluminação de emergência OK",           _sim if ck.iluminacao_ok       else _nao],
        ["Alarme associado funcional",            _sim if ck.alarme_ok           else _nao],
    ]
    t_itens = Table(itens, colWidths=[13.4*cm, 4.0*cm])
    t_itens.setStyle(TableStyle([
        ("BACKGROUND",    (0,0), (-1,0), colors.HexColor("#f3f4f6")),
        ("FONTNAME",      (0,0), (-1,0), "Helvetica-Bold"),
        ("FONTSIZE",      (0,0), (-1,-1), 8),
        ("GRID",          (0,0), (-1,-1), 0.35, colors.HexColor("#d1d5db")),
        ("ALIGN",         (1,0), (1,-1), "CENTER"),
        ("VALIGN",        (0,0), (-1,-1), "MIDDLE"),
        ("TOPPADDING",    (0,0),(-1,-1), 5),
        ("BOTTOMPADDING", (0,0),(-1,-1), 5),
        ("LEFTPADDING",   (0,0),(-1,-1), 7),
        ("RIGHTPADDING",  (0,0),(-1,-1), 7),
        # Colorir NÃO em vermelho
        *[("TEXTCOLOR", (1,i+1), (1,i+1),
           colors.HexColor("#146c43") if itens[i+1][1].startswith("✔") else colors.HexColor("#d40511"))
          for i in range(5)],
    ]))
    elementos += [t_itens, Spacer(1, 8)]

    # Observação
    if ck.observacao:
        elementos.append(Paragraph("OBSERVAÇÃO", sec_style))
        elementos.append(Paragraph(ck.observacao, cel_style))
        elementos.append(Spacer(1, 8))

    # Tratativa (se houver)
    if ck.conclusao:
        elementos.append(Paragraph("TRATATIVA DA NÃO CONFORMIDADE", sec_style))
        trat = [
            [Paragraph("AÇÃO TOMADA",    lbl_style), Paragraph(ck.conclusao or "—", cel_style)],
            [Paragraph("CONCLUÍDO POR",  lbl_style), Paragraph(ck.concluido_por or "—", cel_style)],
            [Paragraph("DATA CONCLUSÃO", lbl_style), Paragraph(ck.data_conclusao.strftime('%d/%m/%Y %H:%M') if ck.data_conclusao else "—", cel_style)],
        ]
        t_trat = Table(trat, colWidths=[4.0*cm, 13.4*cm])
        t_trat.setStyle(TableStyle([
            ("BACKGROUND", (0,0), (0,-1), colors.HexColor("#f0fdf4")),
            ("GRID",       (0,0), (-1,-1), 0.35, colors.HexColor("#d1d5db")),
            ("VALIGN",     (0,0), (-1,-1), "MIDDLE"),
            ("TOPPADDING",    (0,0),(-1,-1), 5),
            ("BOTTOMPADDING", (0,0),(-1,-1), 5),
            ("LEFTPADDING",   (0,0),(-1,-1), 7),
            ("RIGHTPADDING",  (0,0),(-1,-1), 7),
        ]))
        elementos += [t_trat, Spacer(1, 8)]

    elementos.append(Paragraph("Documento gerado automaticamente pelo sistema DHL Security.",
                                ParagraphStyle("rod", parent=styles["BodyText"],
                                               alignment=TA_CENTER,
                                               textColor=colors.HexColor("#6b7280"), fontSize=7)))
    doc.build(elementos)
    buffer.seek(0)
    _site = (porta.site or "DHL").replace(" ", "_")
    _ano  = (ck.data_checklist or datetime.utcnow()).year
    _nome = f"CHK-PE-{_ano}-{ck.id:04d} - {porta.codigo} - {_site}.pdf"
    return send_file(buffer, as_attachment=True,
                     download_name=_nome, mimetype="application/pdf")


# ═══════════════════════════════════════════════════════════════════════════════
# CONCLUIR NÃO CONFORMIDADE — Checklist
# ═══════════════════════════════════════════════════════════════════════════════
@pe_bp.route("/checklist/<int:checklist_id>/concluir", methods=["GET", "POST"])
@_login_required
def concluir_checklist(checklist_id):
    ck = ChecklistPortaEmergencia.query.get_or_404(checklist_id)
    porta = PortaEmergencia.query.get_or_404(ck.porta_id)

    if ck.status != "NÃO CONFORME":
        flash("Apenas checklists não conformes podem ser tratados.", "warning")
        return redirect(url_for("pe.historico", porta_id=porta.id))

    if request.method == "POST":
        conclusao = (request.form.get("conclusao") or "").strip()
        if not conclusao:
            flash("Descreva a ação tomada antes de concluir.", "warning")
            return redirect(url_for("pe.concluir_checklist", checklist_id=checklist_id))

        ck.conclusao      = conclusao
        ck.data_conclusao = datetime.utcnow()
        ck.concluido_por  = session.get("user_nome")
        _db.session.commit()
        flash("Não conformidade tratada com sucesso.", "success")
        return redirect(url_for("pe.historico", porta_id=porta.id))

    ok_count = sum(1 for v in [ck.porta_desobstruida, ck.abre_normalmente,
                                ck.sinalizacao_ok, ck.iluminacao_ok, ck.alarme_ok] if v)
    return render_template("pe_concluir_checklist.html",
                           ck=ck, porta=porta, ok_count=ok_count,
                           user_nome=session.get("user_nome", ""))


# ═══════════════════════════════════════════════════════════════════════════════
# DISPAROS DE ALARME — Lista
# ═══════════════════════════════════════════════════════════════════════════════
@pe_bp.route("/alarmes")
@_login_required
def alarmes():
    site     = _site_usuario()
    is_admin = _is_admin()
    _sites   = _sites_autorizados()
    q = DisparoAlarme.query
    if _sites:
        if len(_sites) == 1:
            q = q.filter(DisparoAlarme.site == _sites[0])
        else:
            q = q.filter(DisparoAlarme.site.in_(_sites))
    registros = q.order_by(DisparoAlarme.id.desc()).all()
    return render_template("pe_alarmes.html", alarmes=registros,
                           is_admin=is_admin, is_gestor=_is_gestor_keyuser())


# ═══════════════════════════════════════════════════════════════════════════════
# DISPAROS DE ALARME — Novo
# ═══════════════════════════════════════════════════════════════════════════════
@pe_bp.route("/alarmes/novo", methods=["GET", "POST"])
@_gestor_required
def novo_alarme():
    if request.method == "POST":
        data_disparo = request.form.get("data_disparo")
        hora_disparo = request.form.get("hora_disparo")
        localizacao  = request.form.get("localizacao")
        tipo_alarme  = request.form.get("tipo_alarme")

        if not data_disparo or not hora_disparo or not localizacao or not tipo_alarme:
            flash("Preencha os campos obrigatórios.", "warning")
            return redirect(url_for("pe.novo_alarme"))

        alarme = DisparoAlarme(
            data_disparo=data_disparo,
            hora_disparo=hora_disparo,
            hora_desativado=request.form.get("hora_desativado"),
            hora_ativado=request.form.get("hora_ativado"),
            contato_monitoramento=request.form.get("contato_monitoramento"),
            localizacao=localizacao,
            setor=request.form.get("setor"),
            tipo_alarme=tipo_alarme,
            motivo=request.form.get("motivo"),
            responsavel=session.get("user_nome"),
            turno=request.form.get("turno"),
            houve_evacuacao=_bool("houve_evacuacao"),
            acionado_bombeiro=_bool("acionado_bombeiro"),
            acionado_seguranca=_bool("acionado_seguranca"),
            status=request.form.get("status") or "EM ANÁLISE",
            observacao=request.form.get("observacao"),
            site=_site_usuario(),
            criado_por=session.get("user_nome"),
        )
        _db.session.add(alarme)
        _db.session.commit()
        flash("Disparo de alarme registrado com sucesso.", "success")
        return redirect(url_for("pe.alarmes"))

    return render_template("pe_novo_alarme.html",
                           setores=_setores_do_site(),
                           locais=_locais_do_site(),
                           user_nome=session.get("user_nome", ""),
                           user_site=_site_usuario() or "")


# ═══════════════════════════════════════════════════════════════════════════════
# DISPAROS DE ALARME — Excluir
# ═══════════════════════════════════════════════════════════════════════════════
@pe_bp.route("/alarmes/<int:alarme_id>/excluir", methods=["POST"])
@_login_required
def excluir_alarme(alarme_id):
    if not _is_admin():
        flash("Apenas administradores podem excluir registros.", "danger")
        return redirect(url_for("pe.alarmes"))
    try:
        reg = DisparoAlarme.query.get_or_404(alarme_id)
        _db.session.delete(reg)
        _db.session.commit()
        flash("Registro excluído.", "success")
    except Exception as e:
        _db.session.rollback()
        flash(f"Erro: {e}", "danger")
    return redirect(url_for("pe.alarmes"))


# ═══════════════════════════════════════════════════════════════════════════════
# BOTÃO DE PÂNICO — Lista
# ═══════════════════════════════════════════════════════════════════════════════
@pe_bp.route("/botao-panico")
@_login_required
def botao_panico():
    site = _site_usuario()
    is_admin = _is_admin()
    q = BotaoPanico.query
    if not is_admin and site:
        q = q.filter_by(site=site)
    registros = q.order_by(BotaoPanico.id.desc()).all()
    return render_template("pe_botao_panico.html", registros=registros,
                           is_admin=is_admin, is_gestor=_is_gestor_keyuser())


# ═══════════════════════════════════════════════════════════════════════════════
# BOTÃO DE PÂNICO — Novo
# ═══════════════════════════════════════════════════════════════════════════════
@pe_bp.route("/botao-panico/novo", methods=["GET", "POST"])
@_gestor_required
def novo_botao_panico():
    if request.method == "POST":
        codigo      = (request.form.get("codigo")      or "").strip().upper()
        localizacao = (request.form.get("localizacao") or "").strip()

        if not codigo or not localizacao:
            flash("Informe o código e a localização do botão de pânico.", "warning")
            return redirect(url_for("pe.novo_botao_panico"))

        testado  = _bool("testado")
        sinal    = _bool("sinal_recebido")
        manut    = _bool("necessita_manutencao")
        status   = "CONFORME" if (testado and sinal and not manut) else "NÃO CONFORME"

        reg = BotaoPanico(
            codigo=codigo,
            localizacao=localizacao,
            setor=request.form.get("setor"),
            tipo=request.form.get("tipo"),
            responsavel=session.get("user_nome"),
            turno=request.form.get("turno"),
            testado=testado,
            sinal_recebido=sinal,
            comunicacao_cftv=_bool("comunicacao_cftv"),
            necessita_manutencao=manut,
            hora_teste=request.form.get("hora_teste") or None,
            hora_retorno=request.form.get("hora_retorno") or None,
            agente_cftv=(request.form.get("agente_cftv") or "").strip() or None,
            status=status,
            observacao=request.form.get("observacao"),
            site=_site_usuario(),
            criado_por=session.get("user_nome"),
        )
        _db.session.add(reg)
        _db.session.commit()
        flash("Registro de botão de pânico salvo com sucesso.", "success")
        return redirect(url_for("pe.botao_panico"))

    return render_template("pe_novo_botao_panico.html",
                           setores=_setores_do_site(),
                           locais=_locais_do_site(),
                           user_nome=session.get("user_nome", ""),
                           user_site=_site_usuario() or "")


# ═══════════════════════════════════════════════════════════════════════════════
# BOTÃO DE PÂNICO — Excluir
# ═══════════════════════════════════════════════════════════════════════════════
@pe_bp.route("/botao-panico/<int:reg_id>/excluir", methods=["POST"])
@_login_required
def excluir_botao_panico(reg_id):
    if not _is_admin():
        flash("Apenas administradores podem excluir registros.", "danger")
        return redirect(url_for("pe.botao_panico"))
    try:
        reg = BotaoPanico.query.get_or_404(reg_id)
        _db.session.delete(reg)
        _db.session.commit()
        flash("Registro excluído.", "success")
    except Exception as e:
        _db.session.rollback()
        flash(f"Erro: {e}", "danger")
    return redirect(url_for("pe.botao_panico"))


# ═══════════════════════════════════════════════════════════════════════════════
# PDF — Disparo de Alarme (individual)
# ═══════════════════════════════════════════════════════════════════════════════
@pe_bp.route("/alarmes/<int:alarme_id>/pdf")
@_login_required
def pdf_alarme(alarme_id):
    alarme = DisparoAlarme.query.get_or_404(alarme_id)
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4,
                            rightMargin=0.9*cm, leftMargin=0.9*cm,
                            topMargin=0.8*cm, bottomMargin=0.8*cm)
    styles = getSampleStyleSheet()

    tit_style = ParagraphStyle("Tit", parent=styles["Heading1"], fontName="Helvetica-Bold",
                               fontSize=17, alignment=TA_CENTER,
                               textColor=colors.HexColor("#111827"), spaceAfter=14)
    sub_style = ParagraphStyle("Sub", parent=styles["BodyText"], fontSize=8,
                               alignment=TA_CENTER, textColor=colors.HexColor("#6b7280"),
                               spaceAfter=10)
    lbl_style = ParagraphStyle("Lbl", parent=styles["BodyText"], fontName="Helvetica-Bold",
                               fontSize=8, textColor=colors.HexColor("#374151"))
    cel_style = ParagraphStyle("Cel", parent=styles["BodyText"],
                               fontSize=8, textColor=colors.HexColor("#111827"))

    elementos = []
    try:
        logo = RLImage(_logo_path(), width=3.5*cm, height=1.1*cm)
        logo.hAlign = "CENTER"
        elementos += [logo, Spacer(1, 6)]
    except Exception:
        pass

    elementos.append(Paragraph("RELATÓRIO DE DISPARO DE ALARME", tit_style))
    elementos.append(Paragraph("DHL SECURITY • CONTROLE OPERACIONAL DE EMERGÊNCIA", sub_style))

    dados = [
        [Paragraph("ID DO REGISTRO",         lbl_style), Paragraph(str(alarme.id), cel_style)],
        [Paragraph("DATA DO DISPARO",         lbl_style), Paragraph(alarme.data_disparo or "-", cel_style)],
        [Paragraph("HORA DISPARO",            lbl_style), Paragraph(alarme.hora_disparo or "-", cel_style)],
        [Paragraph("HORA DESATIVADO",         lbl_style), Paragraph(alarme.hora_desativado or "-", cel_style)],
        [Paragraph("HORA ATIVADO",            lbl_style), Paragraph(alarme.hora_ativado or "-", cel_style)],
        [Paragraph("CONTATO MONITORAMENTO",   lbl_style), Paragraph(alarme.contato_monitoramento or "-", cel_style)],
        [Paragraph("LOCALIZAÇÃO",             lbl_style), Paragraph(alarme.localizacao or "-", cel_style)],
        [Paragraph("SETOR",                   lbl_style), Paragraph(alarme.setor or "-", cel_style)],
        [Paragraph("TIPO DE ALARME",          lbl_style), Paragraph(alarme.tipo_alarme or "-", cel_style)],
        [Paragraph("MOTIVO IDENTIFICADO",     lbl_style), Paragraph(alarme.motivo or "-", cel_style)],
        [Paragraph("RESPONSÁVEL",             lbl_style), Paragraph(alarme.responsavel or "-", cel_style)],
        [Paragraph("TURNO",                   lbl_style), Paragraph(alarme.turno or "-", cel_style)],
        [Paragraph("STATUS",                  lbl_style), Paragraph(alarme.status or "-", cel_style)],
        [Paragraph("HOUVE EVACUAÇÃO",         lbl_style), Paragraph("SIM" if alarme.houve_evacuacao else "NÃO", cel_style)],
        [Paragraph("BOMBEIRO ACIONADO",       lbl_style), Paragraph("SIM" if alarme.acionado_bombeiro else "NÃO", cel_style)],
        [Paragraph("SEGURANÇA ACIONADA",      lbl_style), Paragraph("SIM" if alarme.acionado_seguranca else "NÃO", cel_style)],
        [Paragraph("OBSERVAÇÃO",              lbl_style), Paragraph(alarme.observacao or "-", cel_style)],
    ]
    tabela = Table(dados, colWidths=[4.4*cm, 13.0*cm])
    tabela.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (0, -1), colors.HexColor("#f3f4f6")),
        ("BACKGROUND", (1, 0), (1, -1), colors.white),
        ("GRID", (0, 0), (-1, -1), 0.35, colors.HexColor("#d1d5db")),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("TOPPADDING",    (0,0),(-1,-1), 5),
        ("BOTTOMPADDING", (0,0),(-1,-1), 5),
        ("LEFTPADDING",   (0,0),(-1,-1), 7),
        ("RIGHTPADDING",  (0,0),(-1,-1), 7),
    ]))
    elementos += [tabela, Spacer(1, 10)]
    elementos.append(Paragraph("Documento gerado automaticamente pelo sistema DHL Security.",
                                ParagraphStyle("rod", parent=styles["BodyText"],
                                               alignment=TA_CENTER,
                                               textColor=colors.HexColor("#6b7280"), fontSize=7)))
    doc.build(elementos)
    buffer.seek(0)
    _site = (alarme.site or "DHL").replace(" ", "_")
    _ano  = (alarme.data_registro or datetime.utcnow()).year
    _tipo = (alarme.tipo_alarme or "ALARME").replace(" ", "_")
    _nome = f"DA-{_ano}-{alarme.id:04d} - {_tipo} - {_site}.pdf"
    return send_file(buffer, as_attachment=True,
                     download_name=_nome, mimetype="application/pdf")


# ═══════════════════════════════════════════════════════════════════════════════
# PDF — Disparos de Alarme (geral)
# ═══════════════════════════════════════════════════════════════════════════════
@pe_bp.route("/alarmes/pdf-geral")
@_login_required
def pdf_alarmes_geral():
    site = _site_usuario()
    is_admin = _is_admin()
    q = DisparoAlarme.query
    if not is_admin and site:
        q = q.filter_by(site=site)
    alarmes = q.order_by(DisparoAlarme.id.desc()).all()

    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=landscape(A4),
                            rightMargin=0.5*cm, leftMargin=0.5*cm,
                            topMargin=0.5*cm, bottomMargin=0.5*cm)
    styles = getSampleStyleSheet()
    elementos = []
    try:
        logo = RLImage(_logo_path(), width=2.8*cm, height=0.8*cm)
        logo.hAlign = "CENTER"
        elementos += [logo, Spacer(1, 4)]
    except Exception:
        pass

    elementos.append(Paragraph("RELATÓRIO GERAL - DISPAROS DE ALARME",
                                ParagraphStyle("Tit", parent=styles["Heading1"],
                                               fontName="Helvetica-Bold", fontSize=12,
                                               alignment=TA_CENTER,
                                               textColor=colors.HexColor("#111827"), spaceAfter=8)))
    dados = [["ID", "DATA", "HORA", "LOCALIZAÇÃO", "SETOR", "TIPO", "STATUS"]]
    for a in alarmes:
        dados.append([str(a.id), a.data_disparo or "-", a.hora_disparo or "-",
                      a.localizacao or "-", a.setor or "-",
                      a.tipo_alarme or "-", a.status or "-"])
    tabela = Table(dados,
                   colWidths=[1.2*cm, 2.3*cm, 2.0*cm, 6.2*cm, 3.0*cm, 3.5*cm, 2.7*cm],
                   repeatRows=1)
    tabela.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#f3f4f6")),
        ("FONTNAME",   (0,0), (-1,0), "Helvetica-Bold"),
        ("FONTSIZE",   (0,0), (-1,-1), 6.5),
        ("GRID",       (0,0), (-1,-1), 0.25, colors.HexColor("#d1d5db")),
        ("ROWBACKGROUNDS", (0,1), (-1,-1), [colors.white, colors.HexColor("#fafafa")]),
        ("ALIGN",  (0,0), (-1,-1), "CENTER"),
        ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
        ("TOPPADDING",    (0,0),(-1,-1), 3),
        ("BOTTOMPADDING", (0,0),(-1,-1), 3),
    ]))
    elementos += [tabela, Spacer(1, 6)]
    elementos.append(Paragraph("Documento gerado automaticamente pelo sistema DHL Security.",
                                ParagraphStyle("rod", parent=styles["BodyText"],
                                               alignment=TA_CENTER,
                                               textColor=colors.HexColor("#6b7280"), fontSize=6)))
    doc.build(elementos)
    buffer.seek(0)
    _site = (_site_usuario() or "DHL").replace(" ", "_")
    _data = datetime.utcnow().strftime("%Y%m%d")
    _nome = f"DA-GERAL - {_site} - {_data}.pdf"
    return send_file(buffer, as_attachment=True,
                     download_name=_nome, mimetype="application/pdf")


# ═══════════════════════════════════════════════════════════════════════════════
# PDF — Botão de Pânico (individual)
# ═══════════════════════════════════════════════════════════════════════════════
@pe_bp.route("/botao-panico/<int:reg_id>/pdf")
@_login_required
def pdf_botao(reg_id):
    reg = BotaoPanico.query.get_or_404(reg_id)
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4,
                            rightMargin=0.9*cm, leftMargin=0.9*cm,
                            topMargin=0.8*cm, bottomMargin=0.8*cm)
    styles = getSampleStyleSheet()
    tit_style = ParagraphStyle("Tit", parent=styles["Heading1"], fontName="Helvetica-Bold",
                               fontSize=17, alignment=TA_CENTER,
                               textColor=colors.HexColor("#111827"), spaceAfter=14)
    elementos = []
    try:
        logo = RLImage(_logo_path(), width=3.5*cm, height=1.0*cm)
        logo.hAlign = "CENTER"
        elementos += [logo, Spacer(1, 6)]
    except Exception:
        pass
    elementos.append(Paragraph("RELATÓRIO DE BOTÃO DE PÂNICO", tit_style))

    dados = [
        ["ID",            str(reg.id)],
        ["DATA REGISTRO", reg.data_registro.strftime("%d/%m/%Y %H:%M") if reg.data_registro else "-"],
        ["CÓDIGO",        reg.codigo or "-"],
        ["LOCALIZAÇÃO",   reg.localizacao or "-"],
        ["SETOR",         reg.setor or "-"],
        ["TIPO",          reg.tipo or "-"],
        ["RESPONSÁVEL",   reg.responsavel or "-"],
        ["TURNO",         reg.turno or "-"],
        ["HORA DO TESTE", reg.hora_teste or "-"],
        ["HORA RETORNO",  reg.hora_retorno or "-"],
        ["AGENTE CFTV",   reg.agente_cftv or "-"],
        ["STATUS",        reg.status or "-"],
        ["OBSERVAÇÃO",    reg.observacao or "-"],
    ]
    tabela = Table(dados, colWidths=[5*cm, 11.5*cm])
    tabela.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (0,-1), colors.HexColor("#f3f4f6")),
        ("GRID",       (0,0), (-1,-1), 0.4, colors.HexColor("#d1d5db")),
        ("FONTSIZE",   (0,0), (-1,-1), 8),
        ("TOPPADDING",    (0,0),(-1,-1), 6),
        ("BOTTOMPADDING", (0,0),(-1,-1), 6),
        ("VALIGN", (0,0),(-1,-1), "MIDDLE"),
    ]))
    elementos += [tabela, Spacer(1, 12)]
    elementos.append(Paragraph("Documento gerado automaticamente pelo sistema DHL Security.",
                                ParagraphStyle("rod", parent=styles["BodyText"],
                                               alignment=TA_CENTER,
                                               textColor=colors.HexColor("#6b7280"), fontSize=7)))
    doc.build(elementos)
    buffer.seek(0)
    _site = (reg.site or "DHL").replace(" ", "_")
    _ano  = (reg.data_registro or datetime.utcnow()).year
    _cod  = (reg.codigo or "BP").replace(" ", "_")
    _nome = f"BP-{_ano}-{reg.id:04d} - {_cod} - {_site}.pdf"
    return send_file(buffer, as_attachment=True,
                     download_name=_nome, mimetype="application/pdf")


# ═══════════════════════════════════════════════════════════════════════════════
# PDF — Botões de Pânico (geral)
# ═══════════════════════════════════════════════════════════════════════════════
@pe_bp.route("/botao-panico/pdf-geral")
@_login_required
def pdf_botao_geral():
    site = _site_usuario()
    is_admin = _is_admin()
    q = BotaoPanico.query
    if not is_admin and site:
        q = q.filter_by(site=site)
    registros = q.order_by(BotaoPanico.id.desc()).all()

    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=landscape(A4),
                            rightMargin=0.5*cm, leftMargin=0.5*cm,
                            topMargin=0.5*cm, bottomMargin=0.5*cm)
    styles = getSampleStyleSheet()
    elementos = []
    try:
        logo = RLImage(_logo_path(), width=2.8*cm, height=0.8*cm)
        logo.hAlign = "CENTER"
        elementos += [logo, Spacer(1, 4)]
    except Exception:
        pass
    elementos.append(Paragraph("RELATÓRIO GERAL - BOTÕES DE PÂNICO",
                                ParagraphStyle("Tit", parent=styles["Heading1"],
                                               fontName="Helvetica-Bold", fontSize=12,
                                               alignment=TA_CENTER,
                                               textColor=colors.HexColor("#111827"), spaceAfter=8)))
    dados = [["ID", "DATA", "CÓDIGO", "LOCALIZAÇÃO", "SETOR", "TIPO", "RESP.", "STATUS"]]
    for r in registros:
        dados.append([str(r.id),
                      r.data_registro.strftime("%d/%m/%Y %H:%M") if r.data_registro else "-",
                      r.codigo or "-", r.localizacao or "-", r.setor or "-",
                      r.tipo or "-", r.responsavel or "-", r.status or "-"])
    tabela = Table(dados,
                   colWidths=[1.0*cm, 2.7*cm, 2.0*cm, 6.2*cm, 3.0*cm, 2.5*cm, 3.2*cm, 2.4*cm],
                   repeatRows=1)
    tabela.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#f3f4f6")),
        ("FONTNAME",   (0,0), (-1,0), "Helvetica-Bold"),
        ("FONTSIZE",   (0,0), (-1,-1), 6.2),
        ("GRID",       (0,0), (-1,-1), 0.25, colors.HexColor("#d1d5db")),
        ("ROWBACKGROUNDS", (0,1), (-1,-1), [colors.white, colors.HexColor("#fafafa")]),
        ("ALIGN",  (0,0), (-1,-1), "CENTER"),
        ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
        ("TOPPADDING",    (0,0),(-1,-1), 3),
        ("BOTTOMPADDING", (0,0),(-1,-1), 3),
    ]))
    elementos += [tabela, Spacer(1, 6)]
    elementos.append(Paragraph("Documento gerado automaticamente pelo sistema DHL Security.",
                                ParagraphStyle("rod", parent=styles["BodyText"],
                                               alignment=TA_CENTER,
                                               textColor=colors.HexColor("#6b7280"), fontSize=6)))
    doc.build(elementos)
    buffer.seek(0)
    _site = (_site_usuario() or "DHL").replace(" ", "_")
    _data = datetime.utcnow().strftime("%Y%m%d")
    _nome = f"BP-GERAL - {_site} - {_data}.pdf"
    return send_file(buffer, as_attachment=True,
                     download_name=_nome, mimetype="application/pdf")
