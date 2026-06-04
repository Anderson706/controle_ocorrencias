"""
site_af_blueprint.py — Abertura e Fechamento de Site (SF-154234)
Fluxo: Fechamento → Abertura → Aprovação do Gestor (individual por etapa)
"""
from __future__ import annotations
import json
import os
from datetime import datetime
from functools import wraps
from io import BytesIO

from flask import (Blueprint, session, request, redirect, url_for,
                   flash, render_template, jsonify, send_file)

# ─── Blueprint ────────────────────────────────────────────────────────────────
af_bp = Blueprint("af", __name__, template_folder="templates")

_db        = None
SiteAF     = None
SiteAFItem = None
_Usuario   = None   # model injetado pelo app.py
_UsuarioSite = None

# ─── Itens padrão (do formulário SF-154234) ───────────────────────────────────
_ITENS_FECHAMENTO_PADRAO = [
    "Todas as saídas de emergência estão trancadas?",
    "Todas as entradas do warehouse estão trancadas?",
    "As entradas de veículos estão trancadas?",
    "As baterias dos rádios foram colocadas para carregar?",
    "As câmeras do CFTV estão funcionando?",
    "As portas de escritório estão trancadas?",
    "As janelas estão fechadas?",
    "Vestiários e banheiros estão trancados?",
    "Os alarmes foram ativados?",
]

_ITENS_ABERTURA_PADRAO = [
    "Todas as saídas de emergência estão trancadas?",
    "Todas as entradas do warehouse estão destrancadas?",
    "As entradas de veículos estão destrancadas?",
    "As baterias reservas dos rádios foram colocadas para carregar?",
    "As câmeras do CFTV estão funcionando?",
    "As portas de escritório estão destrancadas?",
    "As janelas dos escritórios estão destrancadas?",
    "Vestiários e banheiros estão destrancados?",
    "Os alarmes de docas foram desativados?",
]


# ─── Setup ───────────────────────────────────────────────────────────────────
def setup_af(db, Usuario=None, UsuarioSite=None):
    global _db, SiteAF, SiteAFItem, _Usuario, _UsuarioSite
    _db          = db
    _Usuario     = Usuario
    _UsuarioSite = UsuarioSite

    class _SiteAF(db.Model):
        __tablename__ = "SITE_AF"
        id   = db.Column(db.Integer, db.Identity(start=1), primary_key=True)
        site = db.Column(db.String(100), nullable=False, index=True)

        # Status geral do ciclo
        status = db.Column(db.String(30), nullable=False, default="AGUARDANDO_ABERTURA")
        # AGUARDANDO_ABERTURA → PENDENTE_APROVACAO → APROVADO / REJEITADO / PARCIAL

        # ── Fechamento ────────────────────────────────────────────────
        fech_data           = db.Column(db.String(10),  nullable=True)
        fech_hora           = db.Column(db.String(5),   nullable=True)
        fech_realizado_por  = db.Column(db.String(120), nullable=True)
        fech_avaliado_por   = db.Column(db.String(120), nullable=True)
        fech_encaminhado    = db.Column(db.String(120), nullable=True)
        fech_assinatura     = db.Column(db.Text,        nullable=True)
        fech_checklist      = db.Column(db.Text,        nullable=True)  # JSON
        fech_criado_em      = db.Column(db.DateTime,    nullable=True)
        # Aprovação do fechamento
        fech_status         = db.Column(db.String(20),  nullable=True)  # PENDENTE | APROVADO | REJEITADO
        fech_aprov_por      = db.Column(db.String(120), nullable=True)
        fech_aprov_em       = db.Column(db.DateTime,    nullable=True)
        fech_aprov_obs      = db.Column(db.String(500), nullable=True)
        fech_aprov_sig      = db.Column(db.Text,        nullable=True)

        # ── Abertura ──────────────────────────────────────────────────
        aber_data              = db.Column(db.String(10),  nullable=True)
        aber_hora              = db.Column(db.String(5),   nullable=True)
        aber_realizado_por     = db.Column(db.String(120), nullable=True)
        aber_cliente           = db.Column(db.String(120), nullable=True)
        aber_unidade           = db.Column(db.String(120), nullable=True)
        aber_alarme_hora       = db.Column(db.String(5),   nullable=True)
        aber_alarme_acionado   = db.Column(db.String(120), nullable=True)
        aber_alarme_funcao     = db.Column(db.String(120), nullable=True)
        aber_alarme_problemas  = db.Column(db.Text,        nullable=True)
        aber_assinatura        = db.Column(db.Text,        nullable=True)
        aber_checklist         = db.Column(db.Text,        nullable=True)  # JSON
        aber_criado_em         = db.Column(db.DateTime,    nullable=True)
        # Aprovação da abertura
        aber_status         = db.Column(db.String(20),  nullable=True)
        aber_aprov_por      = db.Column(db.String(120), nullable=True)
        aber_aprov_em       = db.Column(db.DateTime,    nullable=True)
        aber_aprov_obs      = db.Column(db.String(500), nullable=True)
        aber_aprov_sig      = db.Column(db.Text,        nullable=True)

        # ── Metadados ─────────────────────────────────────────────────
        criado_em  = db.Column(db.DateTime, default=datetime.utcnow)
        criado_por = db.Column(db.String(120), nullable=True)

        def fech_itens(self):
            try: return json.loads(self.fech_checklist or "[]")
            except: return []

        def aber_itens(self):
            try: return json.loads(self.aber_checklist or "[]")
            except: return []

        def nao_conformes_fech(self):
            return [i for i in self.fech_itens() if i.get("ok") == "N"]

        def nao_conformes_aber(self):
            return [i for i in self.aber_itens() if i.get("ok") == "N"]

    class _SiteAFItem(db.Model):
        __tablename__ = "SITE_AF_ITENS"
        id        = db.Column(db.Integer, db.Identity(start=1), primary_key=True)
        site      = db.Column(db.String(100), nullable=False, index=True)
        tipo      = db.Column(db.String(20),  nullable=False)   # FECHAMENTO | ABERTURA
        numero    = db.Column(db.Integer,     nullable=False)
        descricao = db.Column(db.String(500), nullable=False)
        ativo     = db.Column(db.String(1),   nullable=False, default="S")
        criado_por = db.Column(db.String(120), nullable=True)
        criado_em  = db.Column(db.DateTime,   default=datetime.utcnow)

    SiteAF     = _SiteAF
    SiteAFItem = _SiteAFItem
    return af_bp


# ─── Helpers de acesso ───────────────────────────────────────────────────────
def _login_required(f):
    @wraps(f)
    def dec(*a, **kw):
        if not session.get("user_id"):
            return redirect(url_for("login"))
        return f(*a, **kw)
    return dec


def _can_approve():
    return (session.get("user_perfil") or "").upper() in ("ADMIN", "GESTOR", "KEYUSER", "MULTISITES")


def _site_usuario():
    return session.get("user_site") or ""


def _is_admin():
    return (session.get("user_perfil") or "").upper() in ("ADMIN", "MULTISITES")


def _get_usuarios_site(site):
    """Retorna lista de (id, nome, email) dos usuários vinculados ao site."""
    usuarios = []
    if not _Usuario:
        return usuarios
    try:
        # Usuários com site principal
        por_site = _Usuario.query.filter_by(site=site, is_active=True).all()
        for u in por_site:
            usuarios.append({"id": u.id, "nome": u.nome, "email": u.email or ""})
        # Usuários vinculados via UsuarioSite (MULTISITES)
        if _UsuarioSite:
            vinculos = _UsuarioSite.query.filter_by(site_nome=site).all()
            ids_ja = {u["id"] for u in usuarios}
            for v in vinculos:
                if v.usuario_id not in ids_ja:
                    u2 = _Usuario.query.get(v.usuario_id)
                    if u2 and u2.is_active:
                        usuarios.append({"id": u2.id, "nome": u2.nome, "email": u2.email or ""})
        usuarios.sort(key=lambda x: x["nome"])
    except Exception:
        pass
    return usuarios


def _get_itens(site, tipo):
    """Retorna itens customizados do site; se não houver, usa os padrões."""
    custom = SiteAFItem.query.filter_by(site=site, tipo=tipo, ativo="S").order_by(SiteAFItem.numero).all()
    if custom:
        return [c.descricao for c in custom]
    return _ITENS_FECHAMENTO_PADRAO if tipo == "FECHAMENTO" else _ITENS_ABERTURA_PADRAO


def _atualizar_status_ciclo(reg):
    """Recalcula o status geral do ciclo baseado nos status individuais."""
    if reg.fech_status == "APROVADO" and reg.aber_status == "APROVADO":
        reg.status = "APROVADO"
    elif reg.fech_status == "REJEITADO" or reg.aber_status == "REJEITADO":
        reg.status = "REJEITADO"
    elif reg.aber_criado_em:
        # Abertura feita — independente dos status individuais, ciclo aguarda aprovação
        reg.status = "PENDENTE_APROVACAO"
    else:
        # Só fechamento feito — aguarda abertura
        reg.status = "AGUARDANDO_ABERTURA"


def _resolver_email(valor):
    """Retorna e-mail a partir de um valor que pode ser e-mail ou nome de usuário."""
    if not valor:
        return None
    if "@" in valor:
        return valor.strip()
    # Tenta buscar pelo nome no banco
    if _Usuario:
        try:
            u = _Usuario.query.filter(
                _Usuario.nome.ilike(f"%{valor.strip()}%"),
                _Usuario.is_active == True
            ).first()
            if u and u.email:
                return u.email
        except Exception:
            pass
    return None


def _enviar_email_abertura(destinatario_raw, ciclo_id, site, fech_por):
    """Notifica por e-mail a pessoa indicada para realizar a abertura.
    Retorna (ok: bool, erro: str|None)."""
    import smtplib
    from email.mime.multipart import MIMEMultipart
    from email.mime.text import MIMEText

    SMTP_HOST  = "smtp.dhl.com"
    SMTP_PORT  = 25
    EMAIL_FROM = "Security.processassistant@dhl.com"

    destinatario = _resolver_email(destinatario_raw)
    if not destinatario:
        return False, f"Não foi possível resolver e-mail para '{destinatario_raw}'"

    html = f"""
    <div style="background:#f3f4f6;padding:32px 16px;min-height:100vh;">
    <div style="font-family:Arial,sans-serif;max-width:600px;margin:0 auto;">
      <div style="background:#d40511;padding:18px 24px;border-radius:8px 8px 0 0">
        <h2 style="color:#ffcc00;margin:0;font-size:20px">🔐 Abertura de Site Pendente</h2>
      </div>
      <div style="background:#fff;border:1px solid #e5e7eb;padding:24px;border-radius:0 0 8px 8px">
        <p style="font-size:14px;color:#374151">Olá,</p>
        <p style="font-size:14px;color:#374151">
          O fechamento do site <strong>{site}</strong> foi realizado por <strong>{fech_por}</strong>
          e você foi indicado(a) para realizar a <strong>abertura</strong>.
        </p>
        <p style="font-size:14px;color:#374151">
          Acesse o CCTV Control Panel e registre a abertura do site assim que chegar.
        </p>
        <div style="background:#f3f4f6;border-radius:8px;padding:16px;margin-top:16px">
          <p style="margin:0;font-size:12px;color:#6b7280">
            Ciclo ID: <strong>#{ciclo_id}</strong> | Site: <strong>{site}</strong>
          </p>
        </div>
        <p style="font-size:12px;color:#9ca3af;margin-top:20px">
          DHL Security — CCTV Control Panel | Mensagem automática
        </p>
      </div>
    </div>
    </div>"""

    try:
        msg = MIMEMultipart()
        msg["Subject"] = f"[CCTV] Abertura de Site Pendente — {site}"
        msg["From"]    = EMAIL_FROM
        msg["To"]      = destinatario
        msg.attach(MIMEText(html, "html"))

        EMAIL_PASS = "L0sspr3v3ntion@D3VT3AML4TAM"
        with smtplib.SMTP(SMTP_HOST, SMTP_PORT, timeout=15) as sv:
            sv.login(EMAIL_FROM, EMAIL_PASS)
            sv.send_message(msg, to_addrs=[destinatario])
        return True, None

    except smtplib.SMTPException as e:
        return False, f"SMTP: {e}"
    except Exception as e:
        return False, f"{type(e).__name__}: {e}"


# ─── ROTAS ───────────────────────────────────────────────────────────────────

# ── Controle (lista) ──────────────────────────────────────────────────────────
@af_bp.route("/")
@_login_required
def controle():
    site     = _site_usuario()
    is_admin = _is_admin()

    q = SiteAF.query
    if not is_admin:
        q = q.filter_by(site=site)

    # Filtros
    f_status = (request.args.get("status") or "").strip()
    f_site   = (request.args.get("site_filtro") or "").strip() if is_admin else ""
    f_ini    = (request.args.get("data_inicial") or "").strip()
    f_fim    = (request.args.get("data_final") or "").strip()

    if f_status: q = q.filter(SiteAF.status == f_status)
    if f_site:   q = q.filter(SiteAF.site == f_site)
    if f_ini:    q = q.filter(SiteAF.fech_data >= f_ini)
    if f_fim:    q = q.filter(SiteAF.fech_data <= f_fim)

    registros = q.order_by(SiteAF.id.desc()).all()

    # KPIs
    todos = SiteAF.query if is_admin else SiteAF.query.filter_by(site=site)
    todos = todos.all()
    kpis = {
        "total":               len(todos),
        "aguardando_abertura": sum(1 for r in todos if r.status == "AGUARDANDO_ABERTURA"),
        "pendente_aprovacao":  sum(1 for r in todos if r.status == "PENDENTE_APROVACAO"),
        "aprovados":           sum(1 for r in todos if r.status == "APROVADO"),
        "rejeitados":          sum(1 for r in todos if r.status == "REJEITADO"),
    }

    todos_sites = sorted({r.site for r in SiteAF.query.with_entities(SiteAF.site).all() if r.site}) if is_admin else []
    filtros = {"status": f_status, "site_filtro": f_site, "data_inicial": f_ini, "data_final": f_fim}

    return render_template("site_af/controle.html",
        registros=registros, kpis=kpis, filtros=filtros,
        site=site, is_admin=is_admin, can_approve=_can_approve(),
        todos_sites=todos_sites)


# ── Novo Fechamento ───────────────────────────────────────────────────────────
@af_bp.route("/fechamento/novo", methods=["GET", "POST"])
@_login_required
def novo_fechamento():
    site = _site_usuario()
    itens = _get_itens(site, "FECHAMENTO")

    if request.method == "POST":
        f = request.form
        respostas = []
        for i, desc in enumerate(itens):
            ok  = f.get(f"item_{i}_ok", "")
            obs = (f.get(f"item_{i}_obs") or "").strip()
            respostas.append({"descricao": desc, "ok": ok, "obs": obs})

        reg = SiteAF(
            site               = site,
            status             = "AGUARDANDO_ABERTURA",
            fech_data          = (f.get("data") or "").strip(),
            fech_hora          = (f.get("hora") or "").strip(),
            fech_realizado_por = (f.get("realizado_por") or session.get("user_nome", "")).strip(),
            fech_avaliado_por  = None,  # preenchido automaticamente pelo gestor ao aprovar
            fech_encaminhado   = (f.get("encaminhado_para") or "").strip(),
            fech_assinatura    = f.get("assinatura") or None,
            fech_checklist     = json.dumps(respostas, ensure_ascii=False),
            fech_criado_em     = datetime.now(),
            fech_status        = "PENDENTE",
            criado_por         = session.get("user_nome"),
        )
        _db.session.add(reg)
        _db.session.commit()

        # Notifica por e-mail o responsável pela abertura
        enc = reg.fech_encaminhado or ""
        if enc:
            ok, erro = _enviar_email_abertura(enc, reg.id, site, reg.fech_realizado_por)
            if ok:
                flash("Fechamento registrado! E-mail enviado para o responsável pela abertura.", "success")
            else:
                flash(f"Fechamento registrado. Falha ao enviar e-mail: {erro}", "warning")
        else:
            flash("Fechamento registrado com sucesso! Aguardando abertura.", "success")
        return redirect(url_for("af.controle"))

    agora = datetime.now()
    return render_template("site_af/fechamento.html",
        site=site, itens=itens,
        data_hoje=agora.strftime("%Y-%m-%d"),
        hora_atual=agora.strftime("%H:%M"),
        user_nome=session.get("user_nome", ""),
        usuarios_site=_get_usuarios_site(site))


# ── Nova Abertura ─────────────────────────────────────────────────────────────
@af_bp.route("/<int:ciclo_id>/abertura", methods=["GET", "POST"])
@_login_required
def abertura(ciclo_id):
    reg  = SiteAF.query.get_or_404(ciclo_id)
    site = _site_usuario()

    if not _is_admin() and reg.site != site:
        flash("Acesso negado.", "danger")
        return redirect(url_for("af.controle"))

    if reg.aber_criado_em:
        flash("Abertura já foi registrada para este ciclo.", "warning")
        return redirect(url_for("af.controle"))

    itens = _get_itens(reg.site, "ABERTURA")

    if request.method == "POST":
        f = request.form
        respostas = []
        for i, desc in enumerate(itens):
            ok  = f.get(f"item_{i}_ok", "")
            obs = (f.get(f"item_{i}_obs") or "").strip()
            respostas.append({"descricao": desc, "ok": ok, "obs": obs})

        reg.aber_data             = (f.get("data") or "").strip()
        reg.aber_hora             = (f.get("hora") or "").strip()
        reg.aber_realizado_por    = (f.get("realizado_por") or session.get("user_nome", "")).strip()
        reg.aber_cliente          = (f.get("cliente") or "").strip()
        reg.aber_unidade          = (f.get("unidade") or "").strip()
        reg.aber_alarme_hora      = (f.get("alarme_hora") or "").strip()
        reg.aber_alarme_acionado  = (f.get("alarme_acionado_por") or "").strip()
        reg.aber_alarme_funcao    = (f.get("alarme_funcao") or "").strip()
        reg.aber_alarme_problemas = (f.get("alarme_problemas") or "").strip() or None
        reg.aber_assinatura       = f.get("assinatura") or None
        reg.aber_checklist        = json.dumps(respostas, ensure_ascii=False)
        reg.aber_criado_em        = datetime.now()
        reg.aber_status           = "PENDENTE"
        reg.status                = "PENDENTE_APROVACAO"

        _db.session.commit()
        flash("Abertura registrada! Aguardando aprovação do gestor.", "success")
        return redirect(url_for("af.controle"))

    agora = datetime.now()
    return render_template("site_af/abertura.html",
        ciclo=reg, itens=itens, site=reg.site,
        data_hoje=agora.strftime("%Y-%m-%d"),
        hora_atual=agora.strftime("%H:%M"),
        user_nome=session.get("user_nome", ""))


# ── Aprovar / Rejeitar Fechamento ─────────────────────────────────────────────
@af_bp.route("/<int:ciclo_id>/aprovar-fechamento", methods=["POST"])
@_login_required
def aprovar_fechamento(ciclo_id):
    if not _can_approve():
        flash("Acesso restrito a gestores e key users.", "danger")
        return redirect(url_for("af.controle"))

    reg  = SiteAF.query.get_or_404(ciclo_id)
    acao = (request.form.get("acao") or "").upper()
    obs  = (request.form.get("obs") or "").strip()

    if acao not in ("APROVADO", "REJEITADO"):
        flash("Ação inválida.", "danger")
        return redirect(url_for("af.controle"))

    if reg.fech_status in ("APROVADO", "REJEITADO"):
        flash("Este fechamento já foi avaliado e não pode ser alterado.", "warning")
        return redirect(url_for("af.controle"))

    reg.fech_status       = acao
    reg.fech_aprov_por    = session.get("user_nome")
    reg.fech_aprov_em     = datetime.now()
    reg.fech_aprov_obs    = obs or None
    reg.fech_aprov_sig    = request.form.get("assinatura") or None
    reg.fech_avaliado_por = session.get("user_nome")
    _atualizar_status_ciclo(reg)
    _db.session.commit()

    msg = "Fechamento aprovado." if acao == "APROVADO" else "Fechamento rejeitado."
    flash(msg, "success" if acao == "APROVADO" else "danger")
    return redirect(url_for("af.controle"))


# ── Aprovar / Rejeitar Abertura ───────────────────────────────────────────────
@af_bp.route("/<int:ciclo_id>/aprovar-abertura", methods=["POST"])
@_login_required
def aprovar_abertura(ciclo_id):
    if not _can_approve():
        flash("Acesso restrito a gestores e key users.", "danger")
        return redirect(url_for("af.controle"))

    reg  = SiteAF.query.get_or_404(ciclo_id)
    acao = (request.form.get("acao") or "").upper()
    obs  = (request.form.get("obs") or "").strip()

    if acao not in ("APROVADO", "REJEITADO"):
        flash("Ação inválida.", "danger")
        return redirect(url_for("af.controle"))

    if reg.aber_status in ("APROVADO", "REJEITADO"):
        flash("Esta abertura já foi avaliada e não pode ser alterada.", "warning")
        return redirect(url_for("af.controle"))

    reg.aber_status    = acao
    reg.aber_aprov_por = session.get("user_nome")
    reg.aber_aprov_em  = datetime.now()
    reg.aber_aprov_obs = obs or None
    reg.aber_aprov_sig = request.form.get("assinatura") or None
    _atualizar_status_ciclo(reg)
    _db.session.commit()

    msg = "Abertura aprovada." if acao == "APROVADO" else "Abertura rejeitada."
    flash(msg, "success" if acao == "APROVADO" else "danger")
    return redirect(url_for("af.controle"))


# ── Dashboard ─────────────────────────────────────────────────────────────────
@af_bp.route("/dashboard")
@_login_required
def dashboard():
    from collections import Counter
    site     = _site_usuario()
    is_admin = _is_admin()

    q = SiteAF.query if is_admin else SiteAF.query.filter_by(site=site)
    f_site = (request.args.get("site_filtro") or "").strip() if is_admin else ""
    f_mes  = (request.args.get("mes") or "").strip()
    if f_site: q = q.filter(SiteAF.site == f_site)
    if f_mes:  q = q.filter(SiteAF.fech_data.like(f"{f_mes}%"))

    todos = q.all()
    status_c = Counter(r.status for r in todos)
    site_c   = Counter(r.site for r in todos)

    # Não conformidades
    nc_fech = sum(len(r.nao_conformes_fech()) for r in todos)
    nc_aber = sum(len(r.nao_conformes_aber()) for r in todos)

    todos_sites = sorted({r.site for r in SiteAF.query.with_entities(SiteAF.site).all() if r.site}) if is_admin else []

    return render_template("site_af/dashboard.html",
        todos=todos, status_c=status_c, site_c=site_c,
        nc_fech=nc_fech, nc_aber=nc_aber,
        site=site, is_admin=is_admin,
        todos_sites=todos_sites,
        filtros={"site_filtro": f_site, "mes": f_mes})


# ── PDF do ciclo ──────────────────────────────────────────────────────────────
@af_bp.route("/<int:ciclo_id>/pdf")
@_login_required
def pdf_ciclo(ciclo_id):
    reg = SiteAF.query.get_or_404(ciclo_id)
    buf = _gerar_pdf(reg)
    nome = f"AF-{reg.site}-{reg.fech_data or 'SEM_DATA'}.pdf"
    return send_file(buf, as_attachment=True, download_name=nome, mimetype="application/pdf")


# ── Detalhe JSON (modal) ──────────────────────────────────────────────────────
@af_bp.route("/<int:ciclo_id>/detalhe")
@_login_required
def detalhe(ciclo_id):
    reg = SiteAF.query.get_or_404(ciclo_id)
    return jsonify({
        "id": reg.id, "site": reg.site, "status": reg.status,
        "fech_data": reg.fech_data, "fech_hora": reg.fech_hora,
        "fech_realizado_por": reg.fech_realizado_por,
        "fech_status": reg.fech_status or "PENDENTE",
        "fech_aprov_por": reg.fech_aprov_por,
        "fech_aprov_obs": reg.fech_aprov_obs,
        "aber_data": reg.aber_data, "aber_hora": reg.aber_hora,
        "aber_realizado_por": reg.aber_realizado_por,
        "aber_status": reg.aber_status or ("PENDENTE" if reg.aber_criado_em else "—"),
        "aber_aprov_por": reg.aber_aprov_por,
        "aber_aprov_obs": reg.aber_aprov_obs,
        "fech_itens": reg.fech_itens(),
        "aber_itens": reg.aber_itens(),
        "nc_fech": len(reg.nao_conformes_fech()),
        "nc_aber": len(reg.nao_conformes_aber()),
        "pdf_url": url_for("af.pdf_ciclo", ciclo_id=reg.id),
    })


# ── Config de itens ───────────────────────────────────────────────────────────
@af_bp.route("/config-itens", methods=["GET", "POST"])
@_login_required
def config_itens():
    if not _can_approve():
        flash("Acesso restrito a gestores e key users.", "danger")
        return redirect(url_for("af.controle"))

    site     = _site_usuario()
    is_admin = _is_admin()

    if request.method == "POST":
        acao = request.form.get("acao")
        if acao == "adicionar":
            tipo = (request.form.get("tipo") or "").strip().upper()
            desc = (request.form.get("descricao") or "").strip()
            site_form = (request.form.get("site") or site).strip() if is_admin else site
            if desc and tipo in ("FECHAMENTO", "ABERTURA"):
                max_num = _db.session.query(_db.func.max(SiteAFItem.numero)).filter_by(site=site_form, tipo=tipo).scalar() or 0
                item = SiteAFItem(site=site_form, tipo=tipo, numero=max_num+1,
                                  descricao=desc, criado_por=session.get("user_nome"))
                _db.session.add(item)
                _db.session.commit()
                flash("Item adicionado.", "success")
        elif acao == "excluir":
            item_id = request.form.get("item_id", type=int)
            item = SiteAFItem.query.get(item_id)
            if item and (is_admin or item.site == site):
                _db.session.delete(item)
                _db.session.commit()
                flash("Item removido.", "success")
        return redirect(url_for("af.config_itens"))

    site_f = (request.args.get("site") or site).strip() if is_admin else site
    itens_fech = SiteAFItem.query.filter_by(site=site_f, tipo="FECHAMENTO", ativo="S").order_by(SiteAFItem.numero).all()
    itens_aber = SiteAFItem.query.filter_by(site=site_f, tipo="ABERTURA",   ativo="S").order_by(SiteAFItem.numero).all()
    todos_sites = sorted({r.site for r in SiteAF.query.with_entities(SiteAF.site).all() if r.site}) if is_admin else []

    return render_template("site_af/config_itens.html",
        itens_fech=itens_fech, itens_aber=itens_aber,
        site=site_f, is_admin=is_admin, todos_sites=todos_sites,
        padrao_fech=_ITENS_FECHAMENTO_PADRAO,
        padrao_aber=_ITENS_ABERTURA_PADRAO)


# ─── Geração de PDF ──────────────────────────────────────────────────────────
def _gerar_pdf(reg: "SiteAF") -> BytesIO:
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.units import cm
    from reportlab.lib import colors
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.enums import TA_CENTER, TA_LEFT
    from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer,
                                    Table, TableStyle, HRFlowable)
    from reportlab.lib.utils import ImageReader

    buf    = BytesIO()
    BLACK  = colors.black
    YELLOW = colors.HexColor("#FFCC00")
    RED    = colors.HexColor("#d40511")
    GREEN  = colors.HexColor("#16a34a")
    GRAY   = colors.HexColor("#f3f4f6")
    pw     = A4[0] - 2.6 * cm

    doc = SimpleDocTemplate(buf, pagesize=A4,
                            leftMargin=1.3*cm, rightMargin=1.3*cm,
                            topMargin=1.0*cm, bottomMargin=2.0*cm)

    s_title = ParagraphStyle("t", fontName="Helvetica-Bold", fontSize=13, alignment=TA_CENTER, textColor=BLACK)
    s_sub   = ParagraphStyle("s", fontName="Helvetica",      fontSize=8,  alignment=TA_CENTER, textColor=colors.HexColor("#6b7280"))
    s_h2    = ParagraphStyle("h2",fontName="Helvetica-Bold", fontSize=10, textColor=BLACK)
    s_body  = ParagraphStyle("b", fontName="Helvetica",      fontSize=8,  textColor=BLACK, leading=11)
    s_th    = ParagraphStyle("th",fontName="Helvetica-Bold", fontSize=8,  textColor=BLACK, alignment=TA_CENTER)
    s_td    = ParagraphStyle("td",fontName="Helvetica",      fontSize=8,  textColor=BLACK)
    s_ok    = ParagraphStyle("ok",fontName="Helvetica-Bold", fontSize=8,  textColor=GREEN)
    s_nao   = ParagraphStyle("no",fontName="Helvetica-Bold", fontSize=8,  textColor=RED)

    story = []

    def _yellow_bar(txt):
        t = Table([[Paragraph(txt, s_h2)]], colWidths=[pw])
        t.setStyle(TableStyle([
            ("BACKGROUND", (0,0),(-1,-1), YELLOW),
            ("BOX",        (0,0),(-1,-1), 0.5, BLACK),
            ("TOPPADDING", (0,0),(-1,-1), 4),
            ("BOTTOMPADDING",(0,0),(-1,-1), 4),
            ("LEFTPADDING", (0,0),(-1,-1), 6),
        ]))
        return t

    def _sig_image(sig_data_b64, width=5*cm, height=1.5*cm):
        """Converte base64 de assinatura em RLImage. Retorna None se inválido."""
        if not sig_data_b64:
            return None
        try:
            import base64
            data = sig_data_b64
            if "," in data:
                data = data.split(",", 1)[1]
            raw = base64.b64decode(data)
            from reportlab.platypus import Image as RLImage
            return RLImage(BytesIO(raw), width=width, height=height)
        except Exception:
            return None

    def _sig_block(label, sig_b64, aprov_por=None, aprov_em=None, obs=None, status=None, status_cor=None):
        """Bloco de assinatura com label, imagem (ou linha) e dados de aprovação."""
        from reportlab.platypus import Image as RLImage
        img = _sig_image(sig_b64)
        sig_cell = img if img else Paragraph("_" * 40, s_body)

        rows = [[Paragraph(label, s_th)]]
        rows.append([sig_cell])
        info_parts = []
        if status:
            info_parts.append(f"<b>Status:</b> {status}")
        if aprov_por:
            info_parts.append(f"<b>Avaliado por:</b> {aprov_por}")
        if aprov_em:
            info_parts.append(f"<b>Em:</b> {aprov_em.strftime('%d/%m/%Y %H:%M')}")
        if obs:
            info_parts.append(f"<b>Obs:</b> {obs}")
        if info_parts:
            rows.append([Paragraph("  |  ".join(info_parts), s_body)])

        t = Table(rows, colWidths=[pw/2 - 0.2*cm])
        style = [
            ("BACKGROUND", (0,0),(-1,0), YELLOW),
            ("BOX",        (0,0),(-1,-1), 0.5, BLACK),
            ("INNERGRID",  (0,0),(-1,-1), 0.3, colors.HexColor("#d1d5db")),
            ("ALIGN",      (0,1),(-1,1), "CENTER"),
            ("TOPPADDING", (0,0),(-1,-1), 4),
            ("BOTTOMPADDING",(0,0),(-1,-1), 4),
            ("LEFTPADDING",(0,0),(-1,-1), 6),
        ]
        if status_cor:
            style.append(("TEXTCOLOR", (0,2),(0,2), status_cor))
        t.setStyle(TableStyle(style))
        return t

    def _checklist_table(itens):
        rows = [[Paragraph("Item", s_th), Paragraph("Sim", s_th),
                 Paragraph("Não", s_th), Paragraph("Comentários", s_th)]]
        for it in itens:
            ok  = it.get("ok", "")
            obs = it.get("obs", "")
            sim_mark = Paragraph("✔", s_ok) if ok == "S" else Paragraph("",  s_td)
            nao_mark = Paragraph("✘", s_nao) if ok == "N" else Paragraph("", s_td)
            rows.append([Paragraph(it.get("descricao",""), s_td), sim_mark, nao_mark, Paragraph(obs, s_td)])
        t = Table(rows, colWidths=[pw*0.55, pw*0.09, pw*0.09, pw*0.27])
        t.setStyle(TableStyle([
            ("BACKGROUND",    (0,0),(-1,0), YELLOW),
            ("BOX",           (0,0),(-1,-1), 0.5, BLACK),
            ("INNERGRID",     (0,0),(-1,-1), 0.3, colors.HexColor("#d1d5db")),
            ("VALIGN",        (0,0),(-1,-1), "MIDDLE"),
            ("TOPPADDING",    (0,0),(-1,-1), 3),
            ("BOTTOMPADDING", (0,0),(-1,-1), 3),
            ("ALIGN",         (1,0),(2,-1),  "CENTER"),
        ]))
        return t

    # ── Cabeçalho ────────────────────────────────────────────────────────────
    logo_path = os.path.join(os.path.dirname(__file__), "static", "logo.png")
    if os.path.exists(logo_path):
        with open(logo_path, "rb") as lf:
            from reportlab.platypus import Image as RLImage
            ir = ImageReader(BytesIO(lf.read()))
            iw, ih = ir.getSize()
            scale  = min((pw*0.2)/iw, (1.2*cm)/ih)
            logo   = RLImage(logo_path, width=iw*scale, height=ih*scale)
    else:
        logo = Paragraph("<b>DHL</b>", s_body)

    hdr = Table([[logo, Paragraph("ABERTURA E FECHAMENTO DO SITE", s_title),
                  Paragraph(f"SF-154234<br/>Ciclo #{reg.id}", s_sub)]],
                colWidths=[pw*0.15, pw*0.7, pw*0.15])
    hdr.setStyle(TableStyle([
        ("BACKGROUND", (0,0),(-1,-1), YELLOW),
        ("BOX",        (0,0),(-1,-1), 0.5, BLACK),
        ("INNERGRID",  (0,0),(-1,-1), 0.5, BLACK),
        ("VALIGN",     (0,0),(-1,-1), "MIDDLE"),
        ("TOPPADDING", (0,0),(-1,-1), 6),
        ("BOTTOMPADDING",(0,0),(-1,-1), 6),
        ("LEFTPADDING",(0,0),(-1,-1), 6),
    ]))
    story += [hdr, Spacer(1, 0.3*cm)]

    # ── Seção Fechamento ──────────────────────────────────────────────────────
    story.append(_yellow_bar("FECHAMENTO DO SITE"))
    info_fech = Table([
        [Paragraph(f"<b>Site:</b> {reg.site}", s_body),
         Paragraph(f"<b>Data:</b> {reg.fech_data or '—'}", s_body),
         Paragraph(f"<b>Hora:</b> {reg.fech_hora or '—'}", s_body)],
        [Paragraph(f"<b>Realizado por:</b> {reg.fech_realizado_por or '—'}", s_body),
         Paragraph(f"<b>Avaliado por:</b> {reg.fech_avaliado_por or '—'}", s_body),
         Paragraph(f"<b>Encaminhado para:</b> {reg.fech_encaminhado or '—'}", s_body)],
    ], colWidths=[pw/3, pw/3, pw/3])
    info_fech.setStyle(TableStyle([
        ("BOX",    (0,0),(-1,-1), 0.5, BLACK),
        ("INNERGRID",(0,0),(-1,-1), 0.3, colors.HexColor("#d1d5db")),
        ("TOPPADDING",(0,0),(-1,-1), 4),
        ("BOTTOMPADDING",(0,0),(-1,-1), 4),
        ("LEFTPADDING",(0,0),(-1,-1), 6),
    ]))
    story += [info_fech, Spacer(1, 0.2*cm)]
    story.append(_checklist_table(reg.fech_itens()))

    # Assinatura fechamento
    if reg.fech_assinatura:
        story.append(Spacer(1, 0.2*cm))
        try:
            import base64
            sig_data = reg.fech_assinatura
            if "," in sig_data:
                sig_data = sig_data.split(",",1)[1]
            sig_bytes = base64.b64decode(sig_data)
            from reportlab.platypus import Image as RLImage
            sig_img = RLImage(BytesIO(sig_bytes), width=5*cm, height=1.5*cm)
            sig_tbl = Table([[sig_img]], colWidths=[pw])
            sig_tbl.setStyle(TableStyle([("ALIGN",(0,0),(-1,-1),"CENTER")]))
            story.append(sig_tbl)
        except Exception:
            pass
    story.append(Spacer(1, 0.4*cm))

    # ── Seção Abertura ────────────────────────────────────────────────────────
    story.append(_yellow_bar("ABERTURA DO SITE"))
    if reg.aber_criado_em:
        info_aber = Table([
            [Paragraph(f"<b>Site:</b> {reg.site}", s_body),
             Paragraph(f"<b>Data:</b> {reg.aber_data or '—'}", s_body),
             Paragraph(f"<b>Hora:</b> {reg.aber_hora or '—'}", s_body)],
            [Paragraph(f"<b>Realizado por:</b> {reg.aber_realizado_por or '—'}", s_body),
             Paragraph(f"<b>Cliente:</b> {reg.aber_cliente or '—'}", s_body),
             Paragraph(f"<b>Unidade:</b> {reg.aber_unidade or '—'}", s_body)],
        ], colWidths=[pw/3, pw/3, pw/3])
        info_aber.setStyle(TableStyle([
            ("BOX",(0,0),(-1,-1), 0.5, BLACK),
            ("INNERGRID",(0,0),(-1,-1), 0.3, colors.HexColor("#d1d5db")),
            ("TOPPADDING",(0,0),(-1,-1), 4),
            ("BOTTOMPADDING",(0,0),(-1,-1), 4),
            ("LEFTPADDING",(0,0),(-1,-1), 6),
        ]))
        story += [info_aber, Spacer(1, 0.2*cm)]

        # Alarme
        alarme_tbl = Table([[
            Paragraph(f"<b>Alarme — Hora:</b> {reg.aber_alarme_hora or '—'}", s_body),
            Paragraph(f"<b>Acionado por:</b> {reg.aber_alarme_acionado or '—'}", s_body),
            Paragraph(f"<b>Função:</b> {reg.aber_alarme_funcao or '—'}", s_body),
        ]], colWidths=[pw/3, pw/3, pw/3])
        alarme_tbl.setStyle(TableStyle([
            ("BOX",(0,0),(-1,-1), 0.5, BLACK),
            ("INNERGRID",(0,0),(-1,-1), 0.3, colors.HexColor("#d1d5db")),
            ("TOPPADDING",(0,0),(-1,-1), 4),
            ("BOTTOMPADDING",(0,0),(-1,-1), 4),
            ("LEFTPADDING",(0,0),(-1,-1), 6),
        ]))
        story += [alarme_tbl, Spacer(1, 0.2*cm)]

        if reg.aber_alarme_problemas:
            prob = Table([[Paragraph(f"<b>Problemas no acionamento:</b> {reg.aber_alarme_problemas}", s_body)]], colWidths=[pw])
            prob.setStyle(TableStyle([("BOX",(0,0),(-1,-1),0.5,BLACK),("LEFTPADDING",(0,0),(-1,-1),6),("TOPPADDING",(0,0),(-1,-1),4),("BOTTOMPADDING",(0,0),(-1,-1),4)]))
            story += [prob, Spacer(1, 0.2*cm)]

        story.append(_checklist_table(reg.aber_itens()))

        # Assinatura da abertura
        story.append(Spacer(1, 0.3*cm))
        sig_aber = _sig_image(reg.aber_assinatura, width=5*cm, height=1.5*cm)
        sig_aber_tbl = Table([
            [Paragraph("Assinatura — Responsável pela Abertura", s_th)],
            [sig_aber if sig_aber else Paragraph("_" * 50, s_body)],
            [Paragraph(reg.aber_realizado_por or "—", s_body)],
        ], colWidths=[pw])
        sig_aber_tbl.setStyle(TableStyle([
            ("BACKGROUND", (0,0),(-1,0), colors.HexColor("#f0fdf4")),
            ("BOX",        (0,0),(-1,-1), 0.5, BLACK),
            ("INNERGRID",  (0,0),(-1,-1), 0.3, colors.HexColor("#d1d5db")),
            ("ALIGN",      (0,1),(-1,1), "CENTER"),
            ("TOPPADDING", (0,0),(-1,-1), 5),
            ("BOTTOMPADDING",(0,0),(-1,-1), 5),
            ("LEFTPADDING",(0,0),(-1,-1), 6),
        ]))
        story.append(sig_aber_tbl)
    else:
        story.append(Table([[Paragraph("Abertura ainda não registrada.", s_body)]], colWidths=[pw]))

    # ── Protocolo de validação ─────────────────────────────────────────────────
    story += [Spacer(1, 0.4*cm), _yellow_bar("Protocolo de Validação")]
    status_cores = {"APROVADO": GREEN, "REJEITADO": RED, "PENDENTE": colors.HexColor("#f59e0b")}

    fech_cor = status_cores.get(reg.fech_status or "PENDENTE", BLACK)
    aber_cor = status_cores.get(reg.aber_status or "PENDENTE", BLACK)

    # Linha de info
    proto = Table([
        [Paragraph("<b>🔒 Fechamento</b>", s_th), Paragraph("<b>🔓 Abertura</b>", s_th)],
        [Paragraph(f"Status: <b>{reg.fech_status or 'PENDENTE'}</b>",
                   ParagraphStyle("fs", fontName="Helvetica-Bold", fontSize=9, textColor=fech_cor)),
         Paragraph(f"Status: <b>{reg.aber_status or 'PENDENTE'}</b>",
                   ParagraphStyle("as", fontName="Helvetica-Bold", fontSize=9, textColor=aber_cor))],
        [Paragraph(f"Avaliado por: {reg.fech_aprov_por or '—'}<br/>"
                   f"Em: {reg.fech_aprov_em.strftime('%d/%m/%Y %H:%M') if reg.fech_aprov_em else '—'}"
                   + (f"<br/>Obs: {reg.fech_aprov_obs}" if reg.fech_aprov_obs else ""), s_body),
         Paragraph(f"Avaliado por: {reg.aber_aprov_por or '—'}<br/>"
                   f"Em: {reg.aber_aprov_em.strftime('%d/%m/%Y %H:%M') if reg.aber_aprov_em else '—'}"
                   + (f"<br/>Obs: {reg.aber_aprov_obs}" if reg.aber_aprov_obs else ""), s_body)],
    ], colWidths=[pw/2, pw/2])
    proto.setStyle(TableStyle([
        ("BACKGROUND", (0,0),(-1,0), YELLOW),
        ("BOX",        (0,0),(-1,-1), 0.5, BLACK),
        ("INNERGRID",  (0,0),(-1,-1), 0.3, colors.HexColor("#d1d5db")),
        ("TOPPADDING", (0,0),(-1,-1), 5),
        ("BOTTOMPADDING",(0,0),(-1,-1), 5),
        ("LEFTPADDING",(0,0),(-1,-1), 8),
        ("VALIGN",     (0,0),(-1,-1), "MIDDLE"),
    ]))
    story.append(proto)

    # Assinaturas das aprovações
    story.append(Spacer(1, 0.3*cm))
    sig_fech_aprov = _sig_image(reg.fech_aprov_sig, width=4.5*cm, height=1.4*cm)
    sig_aber_aprov = _sig_image(reg.aber_aprov_sig, width=4.5*cm, height=1.4*cm)

    sig_row = Table([
        [Paragraph("Assinatura — Aprovação do Fechamento", s_th),
         Paragraph("Assinatura — Aprovação da Abertura", s_th)],
        [sig_fech_aprov if sig_fech_aprov else Paragraph("_" * 40, s_body),
         sig_aber_aprov if sig_aber_aprov else Paragraph("_" * 40, s_body)],
        [Paragraph(reg.fech_aprov_por or "—", s_body),
         Paragraph(reg.aber_aprov_por or "—", s_body)],
    ], colWidths=[pw/2, pw/2])
    sig_row.setStyle(TableStyle([
        ("BACKGROUND", (0,0),(-1,0), YELLOW),
        ("BOX",        (0,0),(-1,-1), 0.5, BLACK),
        ("INNERGRID",  (0,0),(-1,-1), 0.3, colors.HexColor("#d1d5db")),
        ("ALIGN",      (0,1),(-1,1), "CENTER"),
        ("TOPPADDING", (0,0),(-1,-1), 5),
        ("BOTTOMPADDING",(0,0),(-1,-1), 6),
        ("LEFTPADDING",(0,0),(-1,-1), 6),
    ]))
    story.append(sig_row)

    def _footer(canvas, doc):
        canvas.saveState()
        x0, x1 = 1.3*cm, A4[0] - 1.3*cm
        canvas.setStrokeColor(BLACK)
        canvas.setLineWidth(0.5)
        canvas.line(x0, 1.5*cm, x1, 1.5*cm)
        canvas.setFont("Helvetica", 7)
        canvas.setFillColor(colors.HexColor("#6b7280"))
        canvas.drawString(x0, 1.1*cm, f"DHL Security — Abertura e Fechamento do Site | {reg.site}")
        canvas.drawRightString(x1, 1.1*cm, f"Ciclo #{reg.id} | Página {doc.page}")
        canvas.restoreState()

    doc.build(story, onFirstPage=_footer, onLaterPages=_footer)
    buf.seek(0)
    return buf
