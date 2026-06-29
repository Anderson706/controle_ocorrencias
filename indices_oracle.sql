-- ============================================================================
--  CCTV Control Panel — Índices recomendados (Oracle / schema SECPANEL)
-- ----------------------------------------------------------------------------
--  Objetivo: acelerar as consultas mais frequentes do app, que filtram quase
--  sempre por SITE (escopo do usuário) + STATUS/datas, e os lookups de chave.
--
--  Segurança:
--    • Cada índice está num bloco PL/SQL que IGNORA o erro "índice já existe"
--      (ORA-00955) — pode rodar quantas vezes quiser, sem efeito colateral.
--    • Criar índice NÃO altera dados nem trava a tabela de forma relevante em
--      tabelas deste porte. Mesmo assim, prefira rodar fora do horário de pico.
--    • Para remover qualquer um: DROP INDEX <nome>;
--
--  Como aplicar: conectar como o dono do schema (SECPANEL) no SQL Developer /
--  SQLcl / SQL*Plus e executar este arquivo inteiro.
-- ============================================================================

SET DEFINE OFF;

-- Helper visual
PROMPT === Criando indices (ignora os que ja existem) ===

-- ── OCORRENCIAS ─────────────────────────────────────────────────────────────
BEGIN EXECUTE IMMEDIATE 'CREATE INDEX IX_OCC_SITE     ON OCORRENCIAS (SITE)';        EXCEPTION WHEN OTHERS THEN IF SQLCODE != -955 THEN RAISE; END IF; END;
/
BEGIN EXECUTE IMMEDIATE 'CREATE INDEX IX_OCC_STATUS   ON OCORRENCIAS (STATUS)';      EXCEPTION WHEN OTHERS THEN IF SQLCODE != -955 THEN RAISE; END IF; END;
/
BEGIN EXECUTE IMMEDIATE 'CREATE INDEX IX_OCC_CRIADOEM ON OCORRENCIAS (CRIADO_EM)';   EXCEPTION WHEN OTHERS THEN IF SQLCODE != -955 THEN RAISE; END IF; END;
/

-- ── ANCS ────────────────────────────────────────────────────────────────────
BEGIN EXECUTE IMMEDIATE 'CREATE INDEX IX_ANC_SITE_EXC ON ANCS (SITE, EXCLUIDO)';     EXCEPTION WHEN OTHERS THEN IF SQLCODE != -955 THEN RAISE; END IF; END;
/
BEGIN EXECUTE IMMEDIATE 'CREATE INDEX IX_ANC_STATUS   ON ANCS (STATUS)';             EXCEPTION WHEN OTHERS THEN IF SQLCODE != -955 THEN RAISE; END IF; END;
/
BEGIN EXECUTE IMMEDIATE 'CREATE INDEX IX_ANC_EXCLSTAT ON ANCS (EXCL_STATUS)';        EXCEPTION WHEN OTHERS THEN IF SQLCODE != -955 THEN RAISE; END IF; END;
/

-- ── ANALISES_INVESTIGATIVAS ─────────────────────────────────────────────────
BEGIN EXECUTE IMMEDIATE 'CREATE INDEX IX_AI_SITE      ON ANALISES_INVESTIGATIVAS (SITE)';          EXCEPTION WHEN OTHERS THEN IF SQLCODE != -955 THEN RAISE; END IF; END;
/
BEGIN EXECUTE IMMEDIATE 'CREATE INDEX IX_AI_STATUS    ON ANALISES_INVESTIGATIVAS (STATUS_ANALISE)'; EXCEPTION WHEN OTHERS THEN IF SQLCODE != -955 THEN RAISE; END IF; END;
/

-- ── SHIFT HANDOVER (OCORRENCIAS_TURNO) ──────────────────────────────────────
BEGIN EXECUTE IMMEDIATE 'CREATE INDEX IX_SH_SITE      ON OCORRENCIAS_TURNO (SITE)';        EXCEPTION WHEN OTHERS THEN IF SQLCODE != -955 THEN RAISE; END IF; END;
/
BEGIN EXECUTE IMMEDIATE 'CREATE INDEX IX_SH_DATA      ON OCORRENCIAS_TURNO (DATA_OCORRENCIA)'; EXCEPTION WHEN OTHERS THEN IF SQLCODE != -955 THEN RAISE; END IF; END;
/

-- ── ARMARIOS ────────────────────────────────────────────────────────────────
-- (SITE + ATIVO cobre a listagem; COLABORADOR_CPF cobre a checagem de atribuição)
BEGIN EXECUTE IMMEDIATE 'CREATE INDEX IX_ARM_SITE_AT  ON ARMARIO (SITE, ATIVO)';     EXCEPTION WHEN OTHERS THEN IF SQLCODE != -955 THEN RAISE; END IF; END;
/
BEGIN EXECUTE IMMEDIATE 'CREATE INDEX IX_ARM_CPF      ON ARMARIO (COLABORADOR_CPF)'; EXCEPTION WHEN OTHERS THEN IF SQLCODE != -955 THEN RAISE; END IF; END;
/
BEGIN EXECUTE IMMEDIATE 'CREATE INDEX IX_ARMCR_ARM    ON ARMARIO_CHAVE_RESERVA (ARMARIO_ID, STATUS)'; EXCEPTION WHEN OTHERS THEN IF SQLCODE != -955 THEN RAISE; END IF; END;
/

-- ── ACHADOS E PERDIDOS ──────────────────────────────────────────────────────
BEGIN EXECUTE IMMEDIATE 'CREATE INDEX IX_AP_SITE      ON ACHADOS_PERDIDOS (SITE)';   EXCEPTION WHEN OTHERS THEN IF SQLCODE != -955 THEN RAISE; END IF; END;
/

-- ── ABERTURA/FECHAMENTO DE SITE (SITE_AF) ───────────────────────────────────
BEGIN EXECUTE IMMEDIATE 'CREATE INDEX IX_AF_SITE      ON SITE_AF (SITE, STATUS)';    EXCEPTION WHEN OTHERS THEN IF SQLCODE != -955 THEN RAISE; END IF; END;
/

-- ── CHECKLIST DE CÂMERAS (itens por checklist) ──────────────────────────────
BEGIN EXECUTE IMMEDIATE 'CREATE INDEX IX_CHKCAMI_CHK  ON CHECKLIST_CAMERA_ITEM (CHECKLIST_ID)'; EXCEPTION WHEN OTHERS THEN IF SQLCODE != -955 THEN RAISE; END IF; END;
/

-- ── RELEASES (consulta do Updater: WHERE ATIVO='S') ─────────────────────────
BEGIN EXECUTE IMMEDIATE 'CREATE INDEX IX_REL_ATIVO    ON APP_RELEASES (ATIVO)';      EXCEPTION WHEN OTHERS THEN IF SQLCODE != -955 THEN RAISE; END IF; END;
/

-- ── Atualiza estatísticas (ajuda o otimizador a USAR os índices) ────────────
PROMPT === Recalculando estatisticas do schema ===
BEGIN
  DBMS_STATS.GATHER_SCHEMA_STATS(ownname => USER, cascade => TRUE);
EXCEPTION WHEN OTHERS THEN NULL;   -- ignora se nao tiver privilegio
END;
/

PROMPT === Concluido. ===
