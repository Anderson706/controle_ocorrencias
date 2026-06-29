# ativos_blueprint.py — Controle de Ativos integrado ao CCTV Control Panel (REDUNDÂNCIA)
# Registrado em app.py com url_prefix='/ativos'.
# Lê/grava as MESMAS tabelas Supabase do app "Controle de Ativos" (FastAPI/desktop):
#   colaboradores, ativos, colaborador_ativos, sites.
# Objetivo: dar uma via web alternativa (dentro do CCTV) para a mesma operação —
# se o app desktop estiver indisponível, a equipe segue trabalhando por aqui.
# Autenticação usa a sessão do CCTV (session["user_id"]).

import re
from functools import wraps
from flask import (
    Blueprint, render_template, request, redirect, url_for, flash, session
)

ativos_bp = Blueprint("ativos", __name__, template_folder="templates")

# Cliente Supabase injetado por setup_ativos (o mesmo objeto criado no app.py).
sb = None

# Perfis do CCTV com poder de gestão (criar/editar/excluir). Os demais ficam
# em modo leitura + portaria. (Os perfis do CCTV são diferentes dos do app Ativos.)
_PERFIS_GESTAO = ("ADMIN", "GESTOR", "KEYUSER", "MULTISITES")


def setup_ativos(supabase_client):
    global sb
    sb = supabase_client
    return ativos_bp


# ── Auth / contexto ──────────────────────────────────────────────────────────
def _login_required(f):
    @wraps(f)
    def wrapped(*args, **kwargs):
        if not session.get("user_id"):
            flash("Faça login para acessar.", "danger")
            return redirect(url_for("login"))
        return f(*args, **kwargs)
    return wrapped


def _pode_gerir():
    return (session.get("user_perfil") or "").upper() in _PERFIS_GESTAO


def _gestao_required(f):
    @wraps(f)
    def wrapped(*args, **kwargs):
        if not session.get("user_id"):
            return redirect(url_for("login"))
        if not _pode_gerir():
            flash("Você não tem permissão para gerenciar ativos.", "danger")
            return redirect(url_for("ativos.colaboradores"))
        return f(*args, **kwargs)
    return wrapped


def _site_usuario():
    return session.get("user_site") or ""


def _sites_lista():
    try:
        rows = sb.table("sites").select("nome_do_site").order("nome_do_site").execute().data or []
        return [r["nome_do_site"] for r in rows]
    except Exception:
        return []


def _kpis():
    """Indicadores globais (Supabase): total de colaboradores cadastrados, total de
    ativos cadastrados e nº de sites distintos que têm usuários cadastrados."""
    def _count(tabela):
        try:
            return sb.table(tabela).select("id", count="exact").limit(1).execute().count or 0
        except Exception:
            return 0
    try:
        rows = sb.table("usuarios").select("site").execute().data or []
        n_sites = len({(r.get("site") or "").strip() for r in rows if (r.get("site") or "").strip()})
    except Exception:
        n_sites = 0
    return {"colaboradores": _count("colaboradores"), "ativos": _count("ativos"), "sites": n_sites}


# ── Listagem de colaboradores ────────────────────────────────────────────────
@ativos_bp.route("/colaboradores")
@_login_required
def colaboradores():
    busca = (request.args.get("busca") or "").strip().lower()
    q = sb.table("colaboradores").select(
        "id,nome_completo,matricula,cpf,empresa,cargo,site,status"
    )
    # Gestores enxergam todos os sites; os demais só o seu.
    if not _pode_gerir():
        q = q.eq("site", _site_usuario())
    dados = q.order("id", desc=True).execute().data or []
    if busca:
        dados = [c for c in dados if busca in (c.get("nome_completo") or "").lower()
                 or busca in (c.get("cpf") or "").lower()
                 or busca in (c.get("matricula") or "").lower()
                 or busca in (c.get("empresa") or "").lower()]
    return render_template(
        "ativos/colaboradores.html",
        colaboradores=dados, busca=busca,
        sites=_sites_lista(), pode_gerir=_pode_gerir(),
        site_usuario=_site_usuario(), kpis=_kpis(),
    )


def _payload_colab(form):
    return {
        "nome_completo": (form.get("nome_completo") or "").strip().upper(),
        "matricula": ((form.get("matricula") or "").strip().upper() or None),
        "cpf": re.sub(r"\D", "", form.get("cpf") or ""),
        "empresa": ((form.get("empresa") or "").strip().upper() or None),
        "cargo": ((form.get("cargo") or "").strip().upper() or None),
        "site": (form.get("site") or "").strip(),
    }


@ativos_bp.route("/colaboradores/novo", methods=["POST"])
@_gestao_required
def colaborador_novo():
    p = _payload_colab(request.form)
    if not p["nome_completo"] or not p["cpf"] or not p["site"]:
        flash("Nome, CPF e Site são obrigatórios.", "danger")
        return redirect(url_for("ativos.colaboradores"))
    p["status"] = "ATIVO"
    try:
        sb.table("colaboradores").insert(p).execute()
        flash(f"Colaborador {p['nome_completo']} cadastrado.", "success")
    except Exception as exc:
        flash(f"Erro ao cadastrar: {exc}", "danger")
    return redirect(url_for("ativos.colaboradores"))


@ativos_bp.route("/colaboradores/<int:cid>/editar", methods=["POST"])
@_gestao_required
def colaborador_editar(cid):
    p = _payload_colab(request.form)
    try:
        sb.table("colaboradores").update(p).eq("id", cid).execute()
        flash("Colaborador atualizado.", "success")
    except Exception as exc:
        flash(f"Erro ao atualizar: {exc}", "danger")
    return redirect(url_for("ativos.colaboradores"))


@ativos_bp.route("/colaboradores/<int:cid>/status", methods=["POST"])
@_gestao_required
def colaborador_status(cid):
    novo = (request.form.get("status") or "").strip().upper()
    if novo not in ("ATIVO", "INATIVO"):
        novo = "ATIVO"
    try:
        sb.table("colaboradores").update({"status": novo}).eq("id", cid).execute()
        flash(f"Status alterado para {novo}.", "success")
    except Exception as exc:
        flash(f"Erro ao alterar status: {exc}", "danger")
    return redirect(request.form.get("next") or url_for("ativos.colaboradores"))


@ativos_bp.route("/colaboradores/<int:cid>/excluir", methods=["POST"])
@_gestao_required
def colaborador_excluir(cid):
    try:
        # Remove vínculos + ativos vinculados, depois o colaborador (mesma lógica do app Ativos).
        vinc = sb.table("colaborador_ativos").select("ativo_id").eq("colaborador_id", cid).execute().data or []
        ativo_ids = [v["ativo_id"] for v in vinc]
        sb.table("colaborador_ativos").delete().eq("colaborador_id", cid).execute()
        if ativo_ids:
            sb.table("ativos").delete().in_("id", ativo_ids).execute()
        sb.table("colaboradores").delete().eq("id", cid).execute()
        flash("Colaborador e equipamentos vinculados excluídos.", "success")
    except Exception as exc:
        flash(f"Erro ao excluir: {exc}", "danger")
    return redirect(url_for("ativos.colaboradores"))


# ── Equipamentos de um colaborador ───────────────────────────────────────────
def _carregar_colab(cid):
    rows = sb.table("colaboradores").select("*").eq("id", cid).limit(1).execute().data or []
    return rows[0] if rows else None


def _ativos_do_colab(cid):
    vinc = sb.table("colaborador_ativos").select("ativo_id").eq("colaborador_id", cid).execute().data or []
    ids = [v["ativo_id"] for v in vinc]
    if not ids:
        return []
    return sb.table("ativos").select("*").in_("id", ids).execute().data or []


@ativos_bp.route("/colaboradores/<int:cid>")
@_login_required
def equipamentos(cid):
    colab = _carregar_colab(cid)
    if not colab:
        flash("Colaborador não encontrado.", "danger")
        return redirect(url_for("ativos.colaboradores"))
    return render_template(
        "ativos/equipamentos.html",
        colab=colab, ativos=_ativos_do_colab(cid), pode_gerir=_pode_gerir(),
    )


# ── Fotos dos ativos (Supabase Storage — mesmo bucket do app desktop) ─────────
_FOTOS_BUCKET = "fotos-ativos"
_FOTO_EXTS = {".jpg", ".jpeg", ".png", ".webp"}


def _upload_foto_ativo(file, cid, identificador):
    """Sobe a foto do ativo no bucket Storage 'fotos-ativos' e retorna a URL pública.
    Mantém o mesmo bucket/coluna (foto_url) usados pelo app desktop de Ativos."""
    import os, time
    raw = file.read()
    if not raw:
        return None
    ext = (os.path.splitext(file.filename or "")[1] or ".jpg").lower()
    if ext not in _FOTO_EXTS:
        ext = ".jpg"
    ident = re.sub(r"[^A-Za-z0-9_-]", "", (identificador or "").replace(" ", "_"))[:40] or "ativo"
    path = f"colab_{cid}_{ident}_{int(time.time())}{ext}"
    ctype = file.content_type or ("image/png" if ext == ".png" else "image/jpeg")
    sb.storage.from_(_FOTOS_BUCKET).upload(path, raw, {"content-type": ctype, "upsert": "true"})
    return sb.storage.from_(_FOTOS_BUCKET).get_public_url(path)


def _remover_foto_storage(url):
    """Remove o arquivo do Storage a partir da sua URL pública (best-effort)."""
    if not url or f"{_FOTOS_BUCKET}/" not in url:
        return
    try:
        path = url.split(f"{_FOTOS_BUCKET}/", 1)[1].split("?", 1)[0]
        sb.storage.from_(_FOTOS_BUCKET).remove([path])
    except Exception:
        pass


@ativos_bp.route("/colaboradores/<int:cid>/ativo/novo", methods=["POST"])
@_gestao_required
def ativo_novo(cid):
    tipo = (request.form.get("tipo") or "").strip().upper()
    tag = (request.form.get("identificador_unico") or "").strip().upper()
    modelo = (request.form.get("modelo_descricao") or "").strip().upper()
    if not tipo or not tag:
        flash("Tipo e Código identificador são obrigatórios.", "danger")
        return redirect(url_for("ativos.equipamentos", cid=cid))

    # Foto opcional — sobe no Storage antes de inserir o ativo
    foto_url = None
    foto = request.files.get("foto")
    if foto and foto.filename:
        try:
            foto_url = _upload_foto_ativo(foto, cid, tag)
        except Exception:
            foto_url = None  # falha no upload não impede o cadastro do equipamento

    try:
        payload = {"tipo": tipo, "identificador_unico": tag, "modelo_descricao": modelo}
        if foto_url:
            payload["foto_url"] = foto_url
        novo = sb.table("ativos").insert(payload).execute().data[0]
        sb.table("colaborador_ativos").insert({
            "colaborador_id": cid, "ativo_id": novo["id"],
        }).execute()
        flash("Equipamento vinculado." + (" Foto anexada." if foto_url else ""), "success")
    except Exception:
        flash("Erro ao vincular — código já em uso?", "danger")
    return redirect(url_for("ativos.equipamentos", cid=cid))


@ativos_bp.route("/ativo/<int:aid>/foto", methods=["POST"])
@_gestao_required
def ativo_foto(aid):
    """Anexa/substitui a foto de um equipamento já cadastrado."""
    cid = request.form.get("cid")
    _destino = (url_for("ativos.equipamentos", cid=cid) if cid
                else url_for("ativos.colaboradores"))
    foto = request.files.get("foto")
    if not foto or not foto.filename:
        flash("Nenhuma foto enviada.", "warning")
        return redirect(_destino)
    try:
        rows = sb.table("ativos").select("identificador_unico,foto_url").eq("id", aid).limit(1).execute().data or []
        ident = rows[0].get("identificador_unico") if rows else str(aid)
        antiga = rows[0].get("foto_url") if rows else None
        nova_url = _upload_foto_ativo(foto, cid or "x", ident)
        sb.table("ativos").update({"foto_url": nova_url}).eq("id", aid).execute()
        _remover_foto_storage(antiga)
        flash("Foto atualizada.", "success")
    except Exception as exc:
        flash(f"Erro ao enviar foto: {exc}", "danger")
    return redirect(_destino)


@ativos_bp.route("/ativo/<int:aid>/excluir", methods=["POST"])
@_gestao_required
def ativo_excluir(aid):
    cid = request.form.get("cid")
    try:
        # Remove a foto do Storage junto (best-effort) p/ não deixar órfã
        _rows = sb.table("ativos").select("foto_url").eq("id", aid).limit(1).execute().data or []
        sb.table("colaborador_ativos").delete().eq("ativo_id", aid).execute()
        sb.table("ativos").delete().eq("id", aid).execute()
        if _rows:
            _remover_foto_storage(_rows[0].get("foto_url"))
        flash("Equipamento devolvido/removido.", "success")
    except Exception as exc:
        flash(f"Erro ao remover: {exc}", "danger")
    return redirect(url_for("ativos.equipamentos", cid=cid) if cid else url_for("ativos.colaboradores"))
