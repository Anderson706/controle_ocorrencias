# CCTV Control Panel — Patch Notes

---

## Versão 3.5 — 23/05/2026

### Novas Funcionalidades

- **Tipos de valor financeiro nas investigações**
  Substituído o campo genérico "Custo Estimado" por uma seleção entre três categorias distintas:
  - ✔ Valor Recuperado
  - 🛡 Valor Preventivo
  - ⚠ Prejuízo
  Cada tipo preenche uma coluna separada no banco de dados.

- **KPIs financeiros no Overview e Dashboard**
  Cards dedicados para Valor Recuperado, Valor Preventivo e Prejuízo com totais e contagem de registros, em cores semânticas (verde, azul, vermelho).

- **Página dedicada para nova investigação**
  Formulário movido para página própria (`/nova-investigacao`), eliminando o formulário embutido na lista. Botão de edição redireciona para página de edição isolada.

- **Campo "Nº de Referência" (nota fiscal / cód. da remessa)**
  Todos os campos anteriormente nomeados "Sub-Package Number" foram renomeados para "Nº de Referência (nota fiscal, cód. da remessa)" em todo o sistema.

---

### Melhorias de Interface

- **Gráficos do Overview com rótulos de dados**
  Todos os gráficos de rosca exibem percentual dentro das fatias; gráficos de barra exibem valor no final de cada barra.

- **Cores semânticas nos gráficos**
  Status agora têm cores consistentes em todos os módulos:
  - CONCLUÍDO / FINALIZADO / FECHADA → Verde
  - PENDENTE / ABERTO → Vermelho
  - EM ANDAMENTO / EM ACOMPANHAMENTO → Amarelo
  - INCONCLUSIVA → Cinza

- **"Concluído" exibido em verde** nas tabelas de Dashboard e Overview.

- **Badges de gravidade ANC corrigidos**
  - CRÍTICA → Vermelho
  - ALTA → Vermelho
  - MÉDIA → Âmbar
  - BAIXA → Verde

- **Passagem de Turno — KPI "Assinadas" corrigido**
  Contagem passou a verificar a presença real da assinatura de recebimento (`assinatura_entrada`) em vez do campo de status textual.

- **Tabela de Análises Investigativas**
  Removidas as colunas "Criado por" e "Unidade"; coluna "Responsável" renomeada para "Responsável pelo Levantamento".

- **Imagem DHL Security atualizada** na tela de login.

---

### Correções de Bugs

- **Formulário de investigação limpando dados na segunda submissão**
  Causa: `hx-boost="true"` do HTMX interceptava o formulário multipart. Corrigido com `hx-boost="false"` no formulário.

- **ORA-00904 nas colunas de valor financeiro**
  `_init_db` rodava em thread paralela antes das colunas serem criadas. Corrigido tornando a migração síncrona antes da abertura da janela.

- **ORA-12899 no campo GC**
  Campo excedia o limite do `VARCHAR2(120)`. Migrado para `CLOB` via processo correto (add coluna temporária → copiar dados → drop original → renomear).

- **ORA-00932 ao comparar CLOB com string vazia**
  Oracle não aceita `!= ''` em colunas CLOB. Substituído por `func.length() > 0`.

- **Cores incorretas nos gráficos do Overview**
  `STATUS_MAP` não continha `'CONCLUÍDO'` com acento (Í). Adicionado mapeamento completo com todas as variações acentuadas e função de normalização de diacríticos como fallback.

- **"ABERTO" exibido como amarelo** no gráfico de status das ANCs. Corrigido para vermelho.

---

### Otimizações de Conexão e Performance

- **Pool de conexões aumentado para VPN**
  `pool_size` 3 → 5, `max_overflow` 3 → 5, `pool_timeout` 30s → 60s.

- **Timeout TCP configurado para o Oracle**
  `tcp_connect_timeout: 20s` — o app para de tentar conectar após 20 segundos em vez de travar indefinidamente em redes instáveis.

- **Devolução garantida de conexões ao pool**
  `teardown_appcontext` adicionado para garantir `rollback` + `session.remove()` ao final de toda requisição, mesmo em caso de erro.

- **Handler global para erros de banco**
  Qualquer `OperationalError` ou `TimeoutError` do SQLAlchemy agora exibe mensagem amigável ao usuário em vez de tela de erro técnico.

- **Cache da lista de sites na tela de login**
  A consulta `SiteCompleto` passou a usar cache em memória com TTL de 5 minutos, evitando uma query ao Oracle a cada carregamento da tela de login.

- **`foto_perfil` (CLOB) não carregada no login**
  Query de autenticação usa `defer(foto_perfil)` + verificação via `func.length()`, evitando transferência desnecessária do CLOB na hora do login.

- **Sessão permanente (30 dias)**
  `session.permanent = True` adicionado no login — o usuário permanece autenticado entre reabertura do app sem precisar fazer login novamente.

---

### Infraestrutura

- **Migração GC para CLOB** incluída no `_init_db` com processo seguro de 4 etapas compatível com Oracle.
- **Versão do sistema:** `3.5` — atualizada em `APP_VERSION` e em `SISTEMA_CONFIG.VERSAO_EXIGIDA` no banco.
