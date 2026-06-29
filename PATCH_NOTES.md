# CCTV Control Panel — Patch Notes

---

## Versão 4.1.1 — 12/06/2026

### Correções de Bugs

- **Página ANC travando ao carregar com muitos registros**
  Causa: colunas `descricao` e `plano_acao_texto` estavam marcadas como *deferred* (carregamento lazy), mas o template as acessava dentro do loop `{% for r in registros %}` — cada acesso disparava uma query individual ao Oracle, resultando em dezenas de roundtrips e a página nunca terminando de carregar.
  Correção: as duas colunas foram removidas do bloco deferred do endpoint `/anc`, passando a ser carregadas em uma única query junto dos demais campos.

- **Fotos do Achados e Perdidos não aparecendo na lista**
  Causa: a coluna `foto_dados` (CLOB) é *deferred* para evitar transferir dados pesados na listagem. Por isso ela nunca entrava no `__dict__` do objeto ORM, e o template — que verifica `r.foto_dados or r.foto_path` — nunca via a foto, mesmo existindo no banco.
  Correção: uma query leve (`SELECT ID WHERE FOTO_DADOS IS NOT NULL`) determina quais registros têm foto, e o valor `True` é inserido no dict antes de passar ao template, sem carregar o CLOB.

### Novidades

- **Comunicar Patch Notes movido para a tela de Releases**
  O botão "Comunicar Patch Notes" saiu da tela de Usuários e foi integrado diretamente na página de Releases, onde faz mais sentido contextualmente.
  O formulário foi reformulado: agora cada nota tem **Título** e **Descrição** separados (antes era só texto livre por categoria), permitindo comunicados mais estruturados.
  Preview do e-mail em tempo real enquanto você compõe as notas.

---

## Versão 4.1 — 03/06/2026

### Novas Funcionalidades

- **Sistema de Abertura e Fechamento de Site (SF)**
  Novo módulo dedicado ao controle operacional de abertura e fechamento de unidades.
  - Ciclo completo: registro do fechamento com checklist → abertura no dia seguinte com checklist → aprovação em duas etapas por gestor autorizado
  - Itens do checklist configuráveis por administrador (separados para fechamento e abertura)
  - Rastreamento de não conformidades identificadas em cada ciclo
  - Dashboard com histórico de ciclos, status e indicadores de conformidade
  - Geração de PDF por ciclo com assinaturas e itens verificados
  - E-mail automático disparado ao responsável quando a abertura é concluída
  - Controle de acesso: somente usuários autorizados para o site visualizam e operam os ciclos

- **Portas de Emergência**
  Novo módulo para gestão e inspeção de portas de emergência da unidade.
  - Cadastro de portas individualmente ou em lote (importação de múltiplas portas de uma vez)
  - Checklist de inspeção por porta com histórico completo de verificações
  - Status calculado automaticamente: Conforme / Não conforme / Pendente
  - Conclusão de checklist com registro de pendências
  - PDF da inspeção com logo, itens verificados e observações
  - Registro de disparos de alarme por porta, com data, hora e responsável
  - Registro de acionamentos de botão de pânico com localização e responsável
  - Acesso controlado: cadastro e exclusão restritos a gestores e key users

- **Achados e Perdidos**
  Novo módulo completo para controle de objetos encontrados na unidade.
  - Cadastro com foto do objeto (armazenada diretamente no Oracle como CLOB base64)
  - Código de rastreamento gerado automaticamente por site (ex.: `GRU-2026-0042`)
  - Status do item: Pendente / Entregue ao dono / Descartado
  - Alerta visual automático quando o prazo de guarda de 90 dias está vencendo ou já venceu
  - Geração de PDF individual por item com foto, dados e QR de rastreamento
  - Exportação de toda a listagem para Excel
  - Envio por e-mail com PDF como anexo direto via Outlook (sem sair do sistema)
  - Dashboard com KPIs: total pendente, entregues, descartados e itens com prazo vencido
  - Filtros por status, período e busca livre por objeto/responsável/descrição

### Melhorias

- **Configurações de Natureza, Local e Setor**
  Novos modelos `NaturezaConfig`, `LocalConfig` e `SetorConfig` permitem que administradores cadastrem e gerenciem as opções dos campos de seleção sem precisar alterar código — as listas de natureza de ocorrência, locais e setores são agora dinâmicas e configuráveis pela interface.

- **Campos de Cargo e Foto no perfil do usuário**
  Campos `CARGO` e `TEM_FOTO` adicionados ao modelo de usuário, permitindo exibir o cargo na interface e indicar se o usuário já possui foto cadastrada sem precisar carregar o CLOB inteiro.

### Infraestrutura

- Migração automática no `_init_db` para criar as tabelas dos novos módulos (SF, Portas de Emergência, Achados e Perdidos)
- Coluna `FOTO_DADOS` (CLOB) criada automaticamente na tabela `ACHADOS_PERDIDOS` se não existir
- Detecção em tempo de execução da coluna `FOTO_DADOS` — o módulo opera normalmente mesmo se a coluna ainda não existir no banco (graceful degradation)
- **Versão do sistema:** `4.1`

---

## Versão 4.0 — 26/05/2026

### Novas Funcionalidades

- **Admin de Releases — Auto-Update**
  Nova tela administrativa para publicar e gerenciar versões do executável do CCTV Control Panel.
  - Upload do `CCTV_ControlPanel.exe` compilado diretamente pela interface
  - Opção de marcar a versão como ativa imediatamente ao publicar
  - Tabela de histórico com versão, tamanho, data de publicação e autor
  - Ações por versão: download do EXE, ativar como versão vigente, excluir
  - O `CCTV_Updater.exe` consulta o banco, baixa a versão ativa e instala automaticamente na próxima inicialização do usuário

- **Gestão de Multisites**
  Nova tela de administração para controlar quais sites cada usuário pode acessar quando possui perfil multi-site.
  - Vinculação e desvinculação de sites por usuário
  - Troca de site ativo sem precisar fazer logout (`/trocar-site`)
  - Sites autorizados exibidos na tela de perfil do usuário
  - Todas as queries de ocorrências, ANCs e análises respeitam o escopo de sites autorizados

- **Operações em lote**
  Exclusão em lote de câmeras e armários diretamente da tabela, com seleção múltipla e confirmação antes de deletar.

### Melhorias

- **Centralização de verificação de privilégios**
  Helper `_is_privileged()` criado para unificar as verificações de perfil ADMIN/SUPERVISOR em todos os endpoints de câmeras e armários, eliminando checagens duplicadas e inconsistentes.

- **Flash messages melhoradas**
  Mensagens de feedback revisadas em todo o módulo de armários para maior clareza e consistência visual.

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
