"""
send_test_emails.py
Envia e-mails de teste do sistema de exclusão de ANC para deivid.martinsl@dhl.com
Execute: python send_test_emails.py
"""
import smtplib
import sys
import io
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from datetime import datetime

# Força stdout UTF-8 no Windows para evitar UnicodeEncodeError com emojis
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

SMTP_HOST     = "smtp.dhl.com"
SMTP_PORT     = 25
EMAIL_FROM    = "Security.processassistant@dhl.com"
EMAIL_PASSWORD= "L0sspr3v3ntion@D3VT3AML4TAM"
DESTINO       = "deivid.martinsl@dhl.com"
DATA_HORA     = datetime.now().strftime("%d/%m/%Y às %H:%M")

# ── Dados fictícios ──────────────────────────────────────────────
MOCK = {
    "num_anc":     "ANC-2026-007",
    "site":        "GRU - Guarulhos",
    "natureza":    "Furto",
    "data_nc":     "2026-06-10",
    "solicitante": "João da Silva",
    "email_solic": DESTINO,
    "motivo":      "ANC registrada em duplicidade com a ANC-2026-005. Os dados são idênticos e o registro foi feito por engano.",
    "motivo_rej":  "Após análise, o registro é distinto da ANC-2026-005 e deve ser mantido para fins de rastreabilidade.",
}

def enviar(msg, destinatarios, titulo):
    try:
        sv = smtplib.SMTP(SMTP_HOST, SMTP_PORT)
        sv.login(EMAIL_FROM, EMAIL_PASSWORD)
        sv.send_message(msg, to_addrs=destinatarios)
        sv.quit()
        print(f"  ✅  {titulo}")
    except Exception as e:
        print(f"  ❌  {titulo} — ERRO: {e}")


# ════════════════════════════════════════════════════════════════════
# E-MAIL 1 — Solicitação de exclusão (vai para admins + equipe do site)
# ════════════════════════════════════════════════════════════════════
def email_solicitacao():
    msg = MIMEMultipart()
    msg["Subject"] = f"[Solicitação de Exclusão] {MOCK['num_anc']} — {MOCK['site']}"
    msg["From"]    = EMAIL_FROM
    msg["To"]      = DESTINO     # simulando admin
    msg["Cc"]      = DESTINO     # simulando equipe do site (mesmo destino no teste)

    html = f"""
<div style="background:#f3f4f6;padding:32px 16px;min-height:100vh;">
<div style="font-family:Arial,sans-serif;max-width:600px;margin:0 auto;">
  <div style="background:#d40511;padding:18px 24px;border-radius:8px 8px 0 0">
    <h2 style="margin:0;color:#ffcc00;font-size:18px;font-weight:900;">🗑 Solicitação de Exclusão de ANC</h2>
    <p style="margin:6px 0 0;color:rgba(255,255,255,.8);font-size:13px;">CCTV Control Panel &middot; {DATA_HORA}</p>
  </div>
  <div style="background:#fff;padding:28px 24px;">
    <p style="color:#374151;font-size:14px;margin:0 0 20px;">
      O usuário abaixo solicitou a <strong>exclusão</strong> da ANC indicada. Acesse o sistema para aprovar ou rejeitar.
    </p>
    <table style="width:100%;border-collapse:collapse;font-size:14px;margin-bottom:20px;">
      <tr style="background:#fef2f2;">
        <td style="padding:10px 14px;font-weight:700;color:#6b7280;width:140px;border:1px solid #fecaca;">ANC</td>
        <td style="padding:10px 14px;color:#d40511;font-weight:900;border:1px solid #fecaca;">{MOCK['num_anc']}</td>
      </tr>
      <tr>
        <td style="padding:10px 14px;font-weight:700;color:#6b7280;border:1px solid #e5e7eb;">Site</td>
        <td style="padding:10px 14px;color:#1f2937;border:1px solid #e5e7eb;">{MOCK['site']}</td>
      </tr>
      <tr style="background:#f9fafb;">
        <td style="padding:10px 14px;font-weight:700;color:#6b7280;border:1px solid #e5e7eb;">Data / Natureza</td>
        <td style="padding:10px 14px;color:#1f2937;border:1px solid #e5e7eb;">{MOCK['data_nc']} &mdash; {MOCK['natureza']}</td>
      </tr>
      <tr>
        <td style="padding:10px 14px;font-weight:700;color:#6b7280;border:1px solid #e5e7eb;">Solicitante</td>
        <td style="padding:10px 14px;color:#1f2937;border:1px solid #e5e7eb;">{MOCK['solicitante']}</td>
      </tr>
      <tr style="background:#f9fafb;">
        <td style="padding:10px 14px;font-weight:700;color:#6b7280;border:1px solid #e5e7eb;">E-mail</td>
        <td style="padding:10px 14px;color:#1f2937;border:1px solid #e5e7eb;">{MOCK['email_solic']}</td>
      </tr>
    </table>
    <div style="background:#fef9c3;border-left:4px solid #eab308;padding:14px 16px;border-radius:6px;margin-bottom:20px;">
      <p style="margin:0;font-size:13px;font-weight:700;color:#854d0e;">Motivo informado pelo solicitante:</p>
      <p style="margin:8px 0 0;font-size:14px;color:#1f2937;">{MOCK['motivo']}</p>
    </div>
    <p style="color:#6b7280;font-size:13px;">Acesse <strong>ANC &gt; Controle</strong> e expanda a seção <em>Solicitações de Exclusão Pendentes</em> para tomar uma decisão.</p>
  </div>
  <div style="background:#1f2937;color:#9ca3af;padding:14px 24px;text-align:center;font-size:12px;border-radius:0 0 8px 8px;">
    DHL Supply Chain &middot; CCTV Control Panel &middot; Uso interno
  </div>
</div>
</div>"""

    msg.attach(MIMEText(html, "html"))
    enviar(msg, [DESTINO], "E-mail 1 — Solicitação de exclusão (admin)")


# ════════════════════════════════════════════════════════════════════
# E-MAIL 2 — Decisão: APROVADO (vai para o solicitante)
# ════════════════════════════════════════════════════════════════════
def email_aprovado():
    msg = MIMEMultipart()
    msg["Subject"] = f"[ANC Exclusão APROVADA] {MOCK['num_anc']}"
    msg["From"]    = EMAIL_FROM
    msg["To"]      = DESTINO

    html = f"""
<div style="background:#f3f4f6;padding:32px 16px;min-height:100vh;">
<div style="font-family:Arial,sans-serif;max-width:600px;margin:0 auto;">
  <div style="background:#16a34a;padding:18px 24px;border-radius:8px 8px 0 0">
    <h2 style="margin:0;color:#fff;font-size:18px;font-weight:900;">✅ Solicitação de Exclusão — APROVADA</h2>
    <p style="margin:6px 0 0;color:rgba(255,255,255,.8);font-size:13px;">CCTV Control Panel &middot; {DATA_HORA}</p>
  </div>
  <div style="background:#fff;padding:28px 24px;">
    <p style="color:#374151;font-size:14px;margin:0 0 20px;">
      Olá <strong>{MOCK['solicitante']}</strong>, sua solicitação de exclusão da ANC abaixo foi
      <strong style="color:#16a34a;">APROVADA</strong>.
    </p>
    <table style="width:100%;border-collapse:collapse;font-size:14px;margin-bottom:20px;">
      <tr style="background:#f9fafb;">
        <td style="padding:10px 14px;font-weight:700;color:#6b7280;width:140px;border:1px solid #e5e7eb;">ANC</td>
        <td style="padding:10px 14px;color:#1f2937;font-weight:900;border:1px solid #e5e7eb;">{MOCK['num_anc']}</td>
      </tr>
      <tr>
        <td style="padding:10px 14px;font-weight:700;color:#6b7280;border:1px solid #e5e7eb;">Site</td>
        <td style="padding:10px 14px;color:#1f2937;border:1px solid #e5e7eb;">{MOCK['site']}</td>
      </tr>
      <tr style="background:#f9fafb;">
        <td style="padding:10px 14px;font-weight:700;color:#6b7280;border:1px solid #e5e7eb;">Decisão</td>
        <td style="padding:10px 14px;color:#16a34a;font-weight:900;border:1px solid #e5e7eb;">APROVADA</td>
      </tr>
    </table>
    <div style="background:#dcfce7;border-left:4px solid #16a34a;padding:14px 16px;border-radius:6px;">
      <p style="margin:0;font-size:13px;color:#166534;">
        A ANC foi ocultada do sistema. O registro permanece no banco de dados para fins de auditoria.
      </p>
    </div>
  </div>
  <div style="background:#1f2937;color:#9ca3af;padding:14px 24px;text-align:center;font-size:12px;border-radius:0 0 8px 8px;">
    DHL Supply Chain &middot; CCTV Control Panel &middot; Uso interno
  </div>
</div>
</div>"""

    msg.attach(MIMEText(html, "html"))
    enviar(msg, [DESTINO], "E-mail 2 — Decisão: APROVADA (solicitante)")


# ════════════════════════════════════════════════════════════════════
# E-MAIL 3 — Decisão: REJEITADO (vai para o solicitante)
# ════════════════════════════════════════════════════════════════════
def email_rejeitado():
    msg = MIMEMultipart()
    msg["Subject"] = f"[ANC Exclusão REJEITADA] {MOCK['num_anc']}"
    msg["From"]    = EMAIL_FROM
    msg["To"]      = DESTINO

    html = f"""
<div style="background:#f3f4f6;padding:32px 16px;min-height:100vh;">
<div style="font-family:Arial,sans-serif;max-width:600px;margin:0 auto;">
  <div style="background:#d40511;padding:18px 24px;border-radius:8px 8px 0 0">
    <h2 style="margin:0;color:#fff;font-size:18px;font-weight:900;">❌ Solicitação de Exclusão — REJEITADA</h2>
    <p style="margin:6px 0 0;color:rgba(255,255,255,.8);font-size:13px;">CCTV Control Panel &middot; {DATA_HORA}</p>
  </div>
  <div style="background:#fff;padding:28px 24px;">
    <p style="color:#374151;font-size:14px;margin:0 0 20px;">
      Olá <strong>{MOCK['solicitante']}</strong>, sua solicitação de exclusão da ANC abaixo foi
      <strong style="color:#d40511;">REJEITADA</strong>.
    </p>
    <table style="width:100%;border-collapse:collapse;font-size:14px;margin-bottom:20px;">
      <tr style="background:#f9fafb;">
        <td style="padding:10px 14px;font-weight:700;color:#6b7280;width:140px;border:1px solid #e5e7eb;">ANC</td>
        <td style="padding:10px 14px;color:#1f2937;font-weight:900;border:1px solid #e5e7eb;">{MOCK['num_anc']}</td>
      </tr>
      <tr>
        <td style="padding:10px 14px;font-weight:700;color:#6b7280;border:1px solid #e5e7eb;">Site</td>
        <td style="padding:10px 14px;color:#1f2937;border:1px solid #e5e7eb;">{MOCK['site']}</td>
      </tr>
      <tr style="background:#f9fafb;">
        <td style="padding:10px 14px;font-weight:700;color:#6b7280;border:1px solid #e5e7eb;">Decisão</td>
        <td style="padding:10px 14px;color:#d40511;font-weight:900;border:1px solid #e5e7eb;">REJEITADA</td>
      </tr>
    </table>
    <div style="background:#fee2e2;border-left:4px solid #d40511;padding:14px 16px;border-radius:6px;">
      <p style="margin:0;font-size:13px;font-weight:700;color:#991b1b;">Motivo da rejeição:</p>
      <p style="margin:8px 0 0;font-size:14px;color:#1f2937;">{MOCK['motivo_rej']}</p>
    </div>
  </div>
  <div style="background:#1f2937;color:#9ca3af;padding:14px 24px;text-align:center;font-size:12px;border-radius:0 0 8px 8px;">
    DHL Supply Chain &middot; CCTV Control Panel &middot; Uso interno
  </div>
</div>
</div>"""

    msg.attach(MIMEText(html, "html"))
    enviar(msg, [DESTINO], "E-mail 3 — Decisão: REJEITADA (solicitante)")


# ════════════════════════════════════════════════════════════════════
if __name__ == "__main__":
    print(f"\nEnviando e-mails de teste para {DESTINO} ...\n")
    email_solicitacao()
    email_aprovado()
    email_rejeitado()
    print("\nConcluído.\n")
