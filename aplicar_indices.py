#!/usr/bin/env python3
"""
aplicar_indices.py — cria os índices de performance no Oracle (schema SECPANEL).

Usa o mesmo banco e usuário do app. É SEGURO:
  • Não altera nenhum dado — só cria índices (estruturas de leitura).
  • Idempotente: se o índice já existe, apenas avisa e segue (ORA-00955).
  • Mostra cada passo no console.

Pré-requisito: estar com a VPN conectada (o banco é interno da DHL).

Como rodar (na pasta do projeto):
    .\\venv\\Scripts\\python.exe aplicar_indices.py          # pede confirmação
    .\\venv\\Scripts\\python.exe aplicar_indices.py --sim     # roda direto, sem perguntar
"""
import sys
import oracledb

# ─── Conexão (mesmos dados do app/updater) ────────────────────────────────────
_DB_USER    = "SECPANEL"
_DB_PASS    = "SEC003q2w3e4r2026"
_DB_HOST    = "usqasap023-scan.phx-dc.dhl.com"
_DB_PORT    = 1521
_DB_SERVICE = "SECPANEL"

# ─── Índices a criar: (nome, tabela, colunas) ─────────────────────────────────
_INDICES = [
    ("IX_OCC_SITE",     "OCORRENCIAS",              "SITE"),
    ("IX_OCC_STATUS",   "OCORRENCIAS",              "STATUS"),
    ("IX_OCC_CRIADOEM", "OCORRENCIAS",              "CRIADO_EM"),
    ("IX_ANC_SITE_EXC", "ANCS",                     "SITE, EXCLUIDO"),
    ("IX_ANC_STATUS",   "ANCS",                     "STATUS"),
    ("IX_ANC_EXCLSTAT", "ANCS",                     "EXCL_STATUS"),
    ("IX_AI_SITE",      "ANALISES_INVESTIGATIVAS",  "SITE"),
    ("IX_AI_STATUS",    "ANALISES_INVESTIGATIVAS",  "STATUS_ANALISE"),
    ("IX_SH_SITE",      "OCORRENCIAS_TURNO",        "SITE"),
    ("IX_SH_DATA",      "OCORRENCIAS_TURNO",        "DATA_OCORRENCIA"),
    ("IX_ARM_SITE_AT",  "ARMARIO",                  "SITE, ATIVO"),
    ("IX_ARM_CPF",      "ARMARIO",                  "COLABORADOR_CPF"),
    ("IX_ARMCR_ARM",    "ARMARIO_CHAVE_RESERVA",    "ARMARIO_ID, STATUS"),
    ("IX_AP_SITE",      "ACHADOS_PERDIDOS",         "SITE"),
    ("IX_AF_SITE",      "SITE_AF",                  "SITE, STATUS"),
    ("IX_CHKCAMI_CHK",  "CHECKLIST_CAMERA_ITEM",    "CHECKLIST_ID"),
    ("IX_REL_ATIVO",    "APP_RELEASES",             "ATIVO"),
]


def main():
    confirmar = "--sim" not in sys.argv

    print("=" * 64)
    print("  Aplicar índices de performance — schema SECPANEL")
    print("=" * 64)
    print(f"  Banco: {_DB_HOST}:{_DB_PORT}/{_DB_SERVICE}")
    print(f"  {len(_INDICES)} índices a verificar/criar.")
    print("  (Não altera dados. Índice que já existe é apenas ignorado.)")
    print("-" * 64)

    if confirmar:
        resp = input("  Confirma a criação dos índices? (sim/não): ").strip().lower()
        if resp not in ("sim", "s", "yes", "y"):
            print("  Cancelado.")
            return

    print("\n  Conectando ao Oracle (precisa de VPN)...", flush=True)
    try:
        dsn  = oracledb.makedsn(_DB_HOST, _DB_PORT, service_name=_DB_SERVICE)
        conn = oracledb.connect(user=_DB_USER, password=_DB_PASS, dsn=dsn)
    except Exception as exc:
        print(f"\n  ERRO ao conectar: {exc}")
        print("  Verifique se a VPN está conectada e tente de novo.")
        sys.exit(1)
    cur = conn.cursor()
    print("  Conectado.\n")

    criados = ja_existiam = erros = 0
    for nome, tabela, colunas in _INDICES:
        ddl = f"CREATE INDEX {nome} ON {tabela} ({colunas})"
        try:
            cur.execute(ddl)
            conn.commit()
            print(f"  [CRIADO]      {nome:<16} -> {tabela}({colunas})")
            criados += 1
        except oracledb.DatabaseError as e:
            (err,) = e.args
            if err.code == 955:          # ORA-00955: nome já usado (índice existe)
                print(f"  [já existe]   {nome:<16} -> ok, nada a fazer")
                ja_existiam += 1
            elif err.code == 942:        # ORA-00942: tabela/coluna não existe
                print(f"  [PULADO]      {nome:<16} -> tabela {tabela} não encontrada")
                erros += 1
            else:
                print(f"  [ERRO {err.code}]  {nome:<16} -> {err.message.strip()}")
                erros += 1
            conn.rollback()

    # ── Atualiza estatísticas (ensina o otimizador a USAR os índices) ─────────
    print("\n  Recalculando estatísticas do schema (ajuda o Oracle a usar os índices)...", flush=True)
    try:
        cur.execute("BEGIN DBMS_STATS.GATHER_SCHEMA_STATS(ownname => USER, cascade => TRUE); END;")
        conn.commit()
        print("  Estatísticas atualizadas.")
    except Exception as exc:
        print(f"  Aviso: não foi possível atualizar estatísticas ({exc}).")
        print("  (Os índices foram criados; só o recálculo de estatísticas falhou — sem problema.)")

    cur.close()
    conn.close()

    print("\n" + "=" * 64)
    print(f"  Resumo:  {criados} criado(s)  |  {ja_existiam} já existia(m)  |  {erros} erro(s)")
    print("=" * 64)


if __name__ == "__main__":
    main()
