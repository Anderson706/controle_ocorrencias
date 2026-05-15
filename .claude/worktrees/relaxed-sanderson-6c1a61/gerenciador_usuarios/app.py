import os
import sys
import secrets
import string
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from functools import wraps
from datetime import datetime

from flask import (Flask, render_template, redirect, url_for,
                   flash, request, session)
from flask_sqlalchemy import SQLAlchemy
from werkzeug.security import generate_password_hash, check_password_hash
from sqlalchemy import text

# ---------------------------------------------------------------------------
# Path resolution (handles PyInstaller onefile)
# ---------------------------------------------------------------------------
if getattr(sys, "frozen", False):
    base_dir = sys._MEIPASS
else:
    base_dir = os.path.dirname(os.path.abspath(__file__))

app = Flask(
    __name__,
    template_folder=os.path.join(base_dir, "templates"),
    static_folder=os.path.join(base_dir, "static"),
)
app.secret_key = secrets.token_hex(32)

# ---------------------------------------------------------------------------
# Database
# ---------------------------------------------------------------------------
app.config["SQLALCHEMY_DATABASE_URI"] = (
    "oracle+oracledb://SECPANEL:SEC003q2w3e4r2026"
    "@usqasap023-scan.phx-dc.dhl.com:1521/?service_name=SECPANEL"
)
app.config["SQLALCHEMY_ENGINE_OPTIONS"] = {
    "pool_size": 10,
    "max_overflow": 20,
    "pool_recycle": 900,
    "pool_pre_ping": False,
    "pool_timeout": 20,
}
app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False

db = SQLAlchemy(app)

# ---------------------------------------------------------------------------
# SMTP
# ---------------------------------------------------------------------------
SMTP_HOST = "smtp.dhl.com"
SMTP_PORT = 25
EMAIL_FROM = "Security.processassistant@dhl.com"
EMAIL_DEVS = [
    "deivid.martinsl@dhl.com",
    "Gilmar.SantosGJ@dhl.com",
    "anderson.rodriguesd@dhl.com",
]


# ---------------------------------------------------------------------------
# Models
# ---------------------------------------------------------------------------
class User(db.Model):
    __tablename__ = "USERS_LIVRO"

    id            = db.Column("ID",            db.Integer,     primary_key=True)
    nome          = db.Column("NOME",          db.String(120), nullable=False)
    email         = db.Column("EMAIL",         db.String(120), nullable=False, unique=True)
    password_hash = db.Column("PASSWORD_HASH", db.String(255), nullable=False)
    role          = db.Column("ROLE",          db.String(30),  nullable=False, default="OPERACIONAL")
    site          = db.Column("SITE",          db.String(80))
    is_active     = db.Column("IS_ACTIVE",     db.Integer,     nullable=False, default=1)
    created_at    = db.Column("CREATED_AT",    db.Date,        nullable=False)
    foto_perfil   = db.Column("FOTO_PERFIL",   db.Text)
    lgpd_aceito   = db.Column("LGPD_ACEITO",   db.String(3))
    lgpd_aceito_em = db.Column("LGPD_ACEITO_EM", db.Date)


class SolicitacaoCadastro(db.Model):
    __tablename__ = "SOLICITACOES_CADASTRO"

    id        = db.Column("ID",        db.Integer,     primary_key=True)
    nome      = db.Column("NOME",      db.String(120), nullable=False)
    email     = db.Column("EMAIL",     db.String(120), nullable=False)
    site      = db.Column("SITE",      db.String(128))
    status    = db.Column("STATUS",    db.String(20),  nullable=False, default="PENDENTE")
    criado_em = db.Column("CRIADO_EM", db.Date,        nullable=False)


class SiteCompleto(db.Model):
    __tablename__ = "SITES_COMPLETO"

    nome_do_site          = db.Column("NOME_DO_SITE",          db.String(128), primary_key=True)
    endereco              = db.Column("ENDEREÇO",              db.String(255))
    cidade                = db.Column("CIDADE",                db.String(50))
    estado                = db.Column("ESTADO",                db.String(2))
    pais                  = db.Column("PAÍS",                  db.String(26))
    responsavel_security  = db.Column("RESPONSÁVEL_SECURITY",  db.String(128))
    coordenador           = db.Column("COORDENADOR",           db.String(26))
    sector                = db.Column("SECTOR",                db.String(26))


class ResetToken(db.Model):
    __tablename__ = "RESET_TOKENS"

    id        = db.Column("ID",        db.Integer,    primary_key=True)
    user_id   = db.Column("USER_ID",   db.Integer,    nullable=False)
    token     = db.Column("TOKEN",     db.String(6),  nullable=False)
    expira_em = db.Column("EXPIRA_EM", db.Date,       nullable=False)
    usado     = db.Column("USADO",     db.Integer,    nullable=False, default=0)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
def gerar_senha(length: int = 10) -> str:
    chars = string.ascii_letters + string.digits
    while True:
        pwd = "".join(secrets.choice(chars) for _ in range(length))
        if any(c.isdigit() for c in pwd) and any(c.isalpha() for c in pwd):
            return pwd


def enviar_email(destinatarios, assunto: str, corpo_html: str, corpo_texto: str = None) -> bool:
    if isinstance(destinatarios, str):
        destinatarios = [destinatarios]
    try:
        msg = MIMEMultipart("alternative")
        msg["Subject"] = assunto
        msg["From"] = EMAIL_FROM
        msg["To"] = ", ".join(destinatarios)
        if corpo_texto:
            msg.attach(MIMEText(corpo_texto, "plain", "utf-8"))
        msg.attach(MIMEText(corpo_html, "html", "utf-8"))
        with smtplib.SMTP(SMTP_HOST, SMTP_PORT) as srv:
            srv.sendmail(EMAIL_FROM, destinatarios, msg.as_string())
        return True
    except Exception as exc:
        app.logger.error("Erro ao enviar e-mail: %s", exc)
        return False


def get_current_user():
    if "user_id" in session:
        return db.session.get(User, session["user_id"])
    return None


def login_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if "user_id" not in session:
            flash("Faça login para continuar.", "warning")
            return redirect(url_for("login"))
        return f(*args, **kwargs)
    return decorated


@app.context_processor
def inject_user():
    return {"current_user": get_current_user()}


def _sites_list():
    return [s.nome_do_site for s in SiteCompleto.query.order_by(SiteCompleto.nome_do_site).all()]


# ---------------------------------------------------------------------------
# Auth
# ---------------------------------------------------------------------------
@app.route("/login", methods=["GET", "POST"])
def login():
    if "user_id" in session:
        return redirect(url_for("dashboard"))

    if request.method == "POST":
        email = request.form.get("email", "").strip().lower()
        senha = request.form.get("senha", "")
        try:
            user = User.query.filter_by(email=email).first()
            if user and check_password_hash(user.password_hash, senha):
                if user.role != "ADMIN":
                    flash("Acesso restrito a administradores.", "danger")
                elif not user.is_active:
                    flash("Conta inativa. Contate o suporte.", "danger")
                else:
                    session["user_id"] = user.id
                    session["user_nome"] = user.nome
                    flash(f"Bem-vindo, {user.nome}!", "success")
                    return redirect(url_for("dashboard"))
            else:
                flash("E-mail ou senha incorretos.", "danger")
        except Exception as exc:
            app.logger.error("Erro no login: %s", exc)
            flash("Erro ao acessar o banco de dados.", "danger")

    return render_template("login.html")


@app.route("/logout")
def logout():
    session.clear()
    flash("Sessão encerrada.", "info")
    return redirect(url_for("login"))


# ---------------------------------------------------------------------------
# Dashboard
# ---------------------------------------------------------------------------
@app.route("/")
@login_required
def dashboard():
    try:
        total    = User.query.count()
        ativos   = User.query.filter_by(is_active=1).count()
        inativos = User.query.filter_by(is_active=0).count()
        admins   = User.query.filter_by(role="ADMIN").count()
        pendentes = SolicitacaoCadastro.query.filter_by(status="PENDENTE").count()

        recentes = (User.query
                    .order_by(User.created_at.desc())
                    .limit(5).all())
        sols_pendentes = (SolicitacaoCadastro.query
                          .filter_by(status="PENDENTE")
                          .order_by(SolicitacaoCadastro.criado_em.desc())
                          .all())

        by_site = db.session.execute(
            text("SELECT SITE, COUNT(*) AS CNT FROM USERS_LIVRO "
                 "WHERE SITE IS NOT NULL GROUP BY SITE ORDER BY CNT DESC")
        ).fetchall()

        by_role = db.session.execute(
            text("SELECT ROLE, COUNT(*) AS CNT FROM USERS_LIVRO GROUP BY ROLE")
        ).fetchall()

        return render_template(
            "dashboard.html",
            total=total, ativos=ativos, inativos=inativos,
            admins=admins, pendentes=pendentes,
            recentes=recentes, sols_pendentes=sols_pendentes,
            by_site=by_site, by_role=by_role,
        )
    except Exception as exc:
        app.logger.error("Erro no dashboard: %s", exc)
        flash("Erro ao carregar dashboard.", "danger")
        return render_template(
            "dashboard.html",
            total=0, ativos=0, inativos=0, admins=0, pendentes=0,
            recentes=[], sols_pendentes=[], by_site=[], by_role=[],
        )


# ---------------------------------------------------------------------------
# Usuários
# ---------------------------------------------------------------------------
@app.route("/usuarios")
@login_required
def usuarios():
    page          = request.args.get("page", 1, type=int)
    busca         = request.args.get("busca", "").strip()
    filtro_role   = request.args.get("role", "")
    filtro_site   = request.args.get("site", "")
    filtro_status = request.args.get("status", "")

    try:
        q = User.query
        if busca:
            q = q.filter(db.or_(
                User.nome.ilike(f"%{busca}%"),
                User.email.ilike(f"%{busca}%"),
            ))
        if filtro_role:
            q = q.filter_by(role=filtro_role)
        if filtro_site:
            q = q.filter_by(site=filtro_site)
        if filtro_status != "":
            q = q.filter_by(is_active=int(filtro_status))

        pagination = q.order_by(User.created_at.desc()).paginate(
            page=page, per_page=20, error_out=False
        )
        sites = _sites_list()
        return render_template(
            "usuarios.html",
            pagination=pagination, sites=sites,
            busca=busca, filtro_role=filtro_role,
            filtro_site=filtro_site, filtro_status=filtro_status,
        )
    except Exception as exc:
        app.logger.error("Erro ao listar usuários: %s", exc)
        flash("Erro ao carregar lista de usuários.", "danger")
        return render_template(
            "usuarios.html", pagination=None, sites=[],
            busca="", filtro_role="", filtro_site="", filtro_status="",
        )


@app.route("/usuarios/novo", methods=["GET", "POST"])
@login_required
def usuario_novo():
    sites = []
    try:
        sites = _sites_list()
    except Exception as exc:
        app.logger.error("Erro ao carregar sites: %s", exc)

    if request.method == "POST":
        nome  = request.form.get("nome", "").strip()
        email = request.form.get("email", "").strip().lower()
        role  = request.form.get("role", "OPERACIONAL")
        site  = request.form.get("site", "").strip() or None
        senha_temp = gerar_senha()

        try:
            if User.query.filter_by(email=email).first():
                flash("E-mail já cadastrado.", "danger")
                return render_template("usuario_novo.html", sites=sites)

            user = User(
                nome=nome,
                email=email,
                password_hash=generate_password_hash(senha_temp),
                role=role,
                site=site,
                is_active=1,
                created_at=datetime.now().date(),
            )
            db.session.add(user)
            db.session.commit()

            html = f"""
            <h2>Bem-vindo ao DHL Security Panel</h2>
            <p>Olá, <strong>{nome}</strong>!</p>
            <p>Sua conta foi criada. Use as credenciais abaixo para acessar:</p>
            <ul>
                <li><strong>E-mail:</strong> {email}</li>
                <li><strong>Senha temporária:</strong> {senha_temp}</li>
            </ul>
            <p>Por favor, altere sua senha no primeiro acesso.</p>
            <br><p>Equipe DHL Security</p>
            """
            enviar_email(email, "Bem-vindo ao DHL Security Panel — Suas Credenciais", html)
            flash(f"Usuário criado com sucesso! Senha temporária: {senha_temp}", "success")
            return redirect(url_for("usuarios"))
        except Exception as exc:
            db.session.rollback()
            app.logger.error("Erro ao criar usuário: %s", exc)
            flash("Erro ao criar usuário.", "danger")

    return render_template("usuario_novo.html", sites=sites)


@app.route("/usuarios/<int:uid>/editar", methods=["GET", "POST"])
@login_required
def usuario_editar(uid):
    try:
        user = db.session.get(User, uid)
        if not user:
            flash("Usuário não encontrado.", "danger")
            return redirect(url_for("usuarios"))

        sites = _sites_list()

        if request.method == "POST":
            action = request.form.get("action", "save")

            if action == "reset_lgpd":
                user.lgpd_aceito    = None
                user.lgpd_aceito_em = None
                db.session.commit()
                flash("LGPD resetada com sucesso.", "success")
                return redirect(url_for("usuario_editar", uid=uid))

            nome      = request.form.get("nome", "").strip()
            new_email = request.form.get("email", "").strip().lower()

            if new_email != user.email and User.query.filter_by(email=new_email).first():
                flash("E-mail já cadastrado para outro usuário.", "danger")
                return render_template("usuario_editar.html", user=user, sites=sites)

            user.nome      = nome
            user.email     = new_email
            user.role      = request.form.get("role", user.role)
            user.site      = request.form.get("site", "").strip() or None
            user.is_active = int(request.form.get("is_active", 1))
            db.session.commit()
            flash("Usuário atualizado com sucesso.", "success")
            return redirect(url_for("usuarios"))

        return render_template("usuario_editar.html", user=user, sites=sites)
    except Exception as exc:
        db.session.rollback()
        app.logger.error("Erro ao editar usuário %s: %s", uid, exc)
        flash("Erro ao editar usuário.", "danger")
        return redirect(url_for("usuarios"))


@app.route("/usuarios/<int:uid>/toggle-status", methods=["POST"])
@login_required
def usuario_toggle_status(uid):
    try:
        user = db.session.get(User, uid)
        if not user:
            flash("Usuário não encontrado.", "danger")
        else:
            user.is_active = 0 if user.is_active else 1
            db.session.commit()
            label = "ativado" if user.is_active else "desativado"
            flash(f"Usuário {label} com sucesso.", "success")
    except Exception as exc:
        db.session.rollback()
        app.logger.error("Erro ao alternar status %s: %s", uid, exc)
        flash("Erro ao alterar status.", "danger")
    return redirect(url_for("usuarios"))


@app.route("/usuarios/<int:uid>/excluir", methods=["POST"])
@login_required
def usuario_excluir(uid):
    try:
        user = db.session.get(User, uid)
        if not user:
            flash("Usuário não encontrado.", "danger")
        elif user.id == session.get("user_id"):
            flash("Você não pode excluir sua própria conta.", "danger")
        else:
            db.session.delete(user)
            db.session.commit()
            flash("Usuário excluído com sucesso.", "success")
    except Exception as exc:
        db.session.rollback()
        app.logger.error("Erro ao excluir usuário %s: %s", uid, exc)
        flash("Erro ao excluir usuário.", "danger")
    return redirect(url_for("usuarios"))


@app.route("/usuarios/<int:uid>/redefinir-senha", methods=["GET", "POST"])
@login_required
def usuario_redefinir_senha(uid):
    try:
        user = db.session.get(User, uid)
        if not user:
            flash("Usuário não encontrado.", "danger")
            return redirect(url_for("usuarios"))

        nova_senha = None
        if request.method == "POST":
            nova_senha = gerar_senha()
            user.password_hash = generate_password_hash(nova_senha)
            db.session.commit()

            html = f"""
            <h2>Redefinição de Senha — DHL Security Panel</h2>
            <p>Olá, <strong>{user.nome}</strong>!</p>
            <p>Sua senha foi redefinida por um administrador:</p>
            <ul>
                <li><strong>E-mail:</strong> {user.email}</li>
                <li><strong>Nova senha temporária:</strong> {nova_senha}</li>
            </ul>
            <p>Por favor, altere sua senha no próximo acesso.</p>
            <br><p>Equipe DHL Security</p>
            """
            enviar_email(user.email, "Redefinição de Senha — DHL Security Panel", html)
            flash(f"Senha redefinida! Nova senha: {nova_senha}", "success")

        return render_template("usuario_redefinir_senha.html", user=user, nova_senha=nova_senha)
    except Exception as exc:
        db.session.rollback()
        app.logger.error("Erro ao redefinir senha %s: %s", uid, exc)
        flash("Erro ao redefinir senha.", "danger")
        return redirect(url_for("usuarios"))


# ---------------------------------------------------------------------------
# Solicitações
# ---------------------------------------------------------------------------
@app.route("/solicitacoes")
@login_required
def solicitacoes():
    filtro_status = request.args.get("status", "PENDENTE")
    page = request.args.get("page", 1, type=int)
    try:
        q = SolicitacaoCadastro.query
        if filtro_status:
            q = q.filter_by(status=filtro_status)
        pagination = q.order_by(SolicitacaoCadastro.criado_em.desc()).paginate(
            page=page, per_page=20, error_out=False
        )
        sites = _sites_list()
        return render_template(
            "solicitacoes.html",
            pagination=pagination, filtro_status=filtro_status, sites=sites,
        )
    except Exception as exc:
        app.logger.error("Erro ao listar solicitações: %s", exc)
        flash("Erro ao carregar solicitações.", "danger")
        return render_template(
            "solicitacoes.html", pagination=None,
            filtro_status=filtro_status, sites=[],
        )


@app.route("/solicitacoes/<int:sid>/aprovar", methods=["POST"])
@login_required
def solicitacao_aprovar(sid):
    try:
        sol = db.session.get(SolicitacaoCadastro, sid)
        if not sol:
            flash("Solicitação não encontrada.", "danger")
            return redirect(url_for("solicitacoes"))

        nome  = request.form.get("nome", sol.nome).strip()
        email = request.form.get("email", sol.email).strip().lower()
        role  = request.form.get("role", "OPERACIONAL")
        site  = request.form.get("site", sol.site or "").strip() or None
        senha_temp = gerar_senha()

        if User.query.filter_by(email=email).first():
            flash("E-mail já cadastrado.", "danger")
            return redirect(url_for("solicitacoes"))

        user = User(
            nome=nome, email=email,
            password_hash=generate_password_hash(senha_temp),
            role=role, site=site, is_active=1,
            created_at=datetime.now().date(),
        )
        db.session.add(user)
        sol.status = "APROVADO"
        db.session.commit()

        html = f"""
        <h2>Cadastro Aprovado — DHL Security Panel</h2>
        <p>Olá, <strong>{nome}</strong>!</p>
        <p>Sua solicitação foi aprovada. Suas credenciais:</p>
        <ul>
            <li><strong>E-mail:</strong> {email}</li>
            <li><strong>Senha temporária:</strong> {senha_temp}</li>
        </ul>
        <p>Por favor, altere sua senha no primeiro acesso.</p>
        <br><p>Equipe DHL Security</p>
        """
        enviar_email(email, "Cadastro Aprovado — DHL Security Panel", html)
        flash(f"Solicitação aprovada. Senha gerada: {senha_temp}", "success")
    except Exception as exc:
        db.session.rollback()
        app.logger.error("Erro ao aprovar solicitação %s: %s", sid, exc)
        flash("Erro ao aprovar solicitação.", "danger")
    return redirect(url_for("solicitacoes"))


@app.route("/solicitacoes/<int:sid>/rejeitar", methods=["POST"])
@login_required
def solicitacao_rejeitar(sid):
    try:
        sol = db.session.get(SolicitacaoCadastro, sid)
        if not sol:
            flash("Solicitação não encontrada.", "danger")
            return redirect(url_for("solicitacoes"))

        motivo = request.form.get("motivo", "").strip()
        sol.status = "REJEITADO"
        db.session.commit()

        motivo_html = f"<p><strong>Motivo:</strong> {motivo}</p>" if motivo else ""
        html = f"""
        <h2>Solicitação Recusada — DHL Security Panel</h2>
        <p>Olá, <strong>{sol.nome}</strong>!</p>
        <p>Infelizmente, sua solicitação de cadastro foi recusada.</p>
        {motivo_html}
        <p>Entre em contato com o administrador para mais informações.</p>
        <br><p>Equipe DHL Security</p>
        """
        enviar_email(sol.email, "Solicitação de Cadastro Recusada — DHL Security Panel", html)
        flash("Solicitação rejeitada.", "warning")
    except Exception as exc:
        db.session.rollback()
        app.logger.error("Erro ao rejeitar solicitação %s: %s", sid, exc)
        flash("Erro ao rejeitar solicitação.", "danger")
    return redirect(url_for("solicitacoes"))


# ---------------------------------------------------------------------------
# Sites
# ---------------------------------------------------------------------------
@app.route("/sites")
@login_required
def sites():
    try:
        sites_list = SiteCompleto.query.order_by(SiteCompleto.nome_do_site).all()
        user_counts = {}
        rows = db.session.execute(
            text("SELECT SITE, COUNT(*) AS CNT FROM USERS_LIVRO "
                 "WHERE SITE IS NOT NULL GROUP BY SITE")
        ).fetchall()
        for row in rows:
            user_counts[row[0]] = row[1]
        return render_template("sites.html", sites=sites_list, user_counts=user_counts)
    except Exception as exc:
        app.logger.error("Erro ao listar sites: %s", exc)
        flash("Erro ao carregar sites.", "danger")
        return render_template("sites.html", sites=[], user_counts={})


@app.route("/sites/novo", methods=["GET", "POST"])
@login_required
def site_novo():
    if request.method == "POST":
        try:
            nome = request.form.get("nome_do_site", "").strip()
            if not nome:
                flash("Nome do site é obrigatório.", "danger")
                return render_template("site_editar.html", site=None, is_new=True)
            if db.session.get(SiteCompleto, nome):
                flash("Site já cadastrado com esse nome.", "danger")
                return render_template("site_editar.html", site=None, is_new=True)

            site = SiteCompleto(
                nome_do_site=nome,
                endereco=request.form.get("endereco", "").strip() or None,
                cidade=request.form.get("cidade", "").strip() or None,
                estado=request.form.get("estado", "").strip() or None,
                pais=request.form.get("pais", "").strip() or None,
                responsavel_security=request.form.get("responsavel_security", "").strip() or None,
                coordenador=request.form.get("coordenador", "").strip() or None,
                sector=request.form.get("sector", "").strip() or None,
            )
            db.session.add(site)
            db.session.commit()
            flash("Site adicionado com sucesso.", "success")
            return redirect(url_for("sites"))
        except Exception as exc:
            db.session.rollback()
            app.logger.error("Erro ao criar site: %s", exc)
            flash("Erro ao adicionar site.", "danger")

    return render_template("site_editar.html", site=None, is_new=True)


@app.route("/sites/<path:nome>/editar", methods=["GET", "POST"])
@login_required
def site_editar(nome):
    try:
        site = db.session.get(SiteCompleto, nome)
        if not site:
            flash("Site não encontrado.", "danger")
            return redirect(url_for("sites"))

        if request.method == "POST":
            site.endereco             = request.form.get("endereco", "").strip() or None
            site.cidade               = request.form.get("cidade", "").strip() or None
            site.estado               = request.form.get("estado", "").strip() or None
            site.pais                 = request.form.get("pais", "").strip() or None
            site.responsavel_security = request.form.get("responsavel_security", "").strip() or None
            site.coordenador          = request.form.get("coordenador", "").strip() or None
            site.sector               = request.form.get("sector", "").strip() or None
            db.session.commit()
            flash("Site atualizado com sucesso.", "success")
            return redirect(url_for("sites"))

        return render_template("site_editar.html", site=site, is_new=False)
    except Exception as exc:
        db.session.rollback()
        app.logger.error("Erro ao editar site %s: %s", nome, exc)
        flash("Erro ao editar site.", "danger")
        return redirect(url_for("sites"))


# ---------------------------------------------------------------------------
# Perfil
# ---------------------------------------------------------------------------
@app.route("/perfil", methods=["GET", "POST"])
@login_required
def perfil():
    user = get_current_user()

    if request.method == "POST":
        action = request.form.get("action", "update_profile")
        try:
            if action == "update_profile":
                nome      = request.form.get("nome", "").strip()
                new_email = request.form.get("email", "").strip().lower()

                if new_email != user.email and User.query.filter_by(email=new_email).first():
                    flash("E-mail já cadastrado para outro usuário.", "danger")
                    return render_template("perfil.html", user=user)

                user.nome  = nome
                user.email = new_email
                session["user_nome"] = nome
                db.session.commit()
                flash("Perfil atualizado com sucesso.", "success")

            elif action == "change_password":
                senha_atual = request.form.get("senha_atual", "")
                nova_senha  = request.form.get("nova_senha", "")
                confirmar   = request.form.get("confirmar_senha", "")

                if not check_password_hash(user.password_hash, senha_atual):
                    flash("Senha atual incorreta.", "danger")
                elif nova_senha != confirmar:
                    flash("As senhas não coincidem.", "danger")
                elif len(nova_senha) < 6:
                    flash("A nova senha deve ter pelo menos 6 caracteres.", "danger")
                else:
                    user.password_hash = generate_password_hash(nova_senha)
                    db.session.commit()
                    flash("Senha alterada com sucesso.", "success")
        except Exception as exc:
            db.session.rollback()
            app.logger.error("Erro ao atualizar perfil: %s", exc)
            flash("Erro ao atualizar perfil.", "danger")

    return render_template("perfil.html", user=user)


if __name__ == "__main__":
    app.run(host="127.0.0.1", port=5050, debug=True)
