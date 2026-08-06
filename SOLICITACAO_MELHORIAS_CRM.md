# Solicitação de melhorias — módulo Relacionamento (CRM) do Controle-Estoque

## Contexto
O módulo de Relacionamento (CRM estilo Pipedrive) vive principalmente em:
crm-model.js (modelo de dados e normalização), crm-calculos.js (cálculos puros),
crm-store.js (estado e persistência), crm-ui.js (orquestração e renderização),
crm-kanban.js (quadro Kanban), crm-compat.js (utilitários), além de
google-calendar-sync.js, index.html, crm.css, cloudflare-worker/ e as regras do
Firestore. O objetivo desta solicitação é evoluir esse módulo com cinco
melhorias inspiradas em recursos que já existem em ferramentas de mercado
(Monday.com), sem duplicar dados nem quebrar a compatibilidade com registros
existentes. Todas as mudanças de modelo devem manter `normalizarCrm`
idempotente e migrar registros antigos automaticamente (mesmo padrão já usado
para `situacao`, `encaminhamentos` etc.).

---

## 1. Múltiplas visões da mesma base, sem duplicar dados

**Objetivo:** adicionar uma visão de linha do tempo/Gantt sobre os mesmos
`negocios` e `anotacoes` já existentes, sem criar uma cópia dos dados.

**Escopo:**
- Nova visão `'timeline'` (ou `'gantt'`) na lista de visões válidas de
  `normalizarConfig` (hoje: `kanban`, `lista`, `demandas`, `previsao`,
  `excluidos`).
- Renderização de barras horizontais por item, usando `dataRecebimento` (ou
  `dataSolicitacao` da anotação) como início e `dataPrevisao`/`prazo` como
  fim. Sem data de início definida, usar `criadoEm`.
- Agrupamento por funil ou por cliente, com opção de alternar (reaproveitar
  `agruparPorEtapa`/`vinculoDaAnotacao` já existentes em `crm-calculos.js`).
- Zoom por semana/mês/trimestre.
- Não persistir nenhuma estrutura nova em `estoque.crm` — a visão é somente
  leitura sobre `negocios`/`anotacoes`, calculada em tempo de renderização
  (mesmo espírito de `agruparPorMesFechamento`).

**Arquivos provavelmente afetados:** crm-ui.js (nova função de renderização e
novo item no seletor de visão), crm-calculos.js (função pura para calcular as
barras de tempo a partir de negócios/anotações), crm.css (estilos da timeline).

**Critérios de aceite:**
- Trocar para a visão de timeline não altera nenhum dado salvo.
- Itens sem data de início ou fim aparecem numa faixa "sem prazo definido",
  sem quebrar o layout.
- Testes unitários (Vitest) para a função de cálculo das barras, cobrindo
  itens sem datas, com datas invertidas e com datas normais.

---

## 2. Responsável como usuário real, não texto livre

**Objetivo:** substituir o campo `responsavel` (string livre) em `negocio` e
`anotacao` por uma referência a um usuário cadastrado do sistema, permitindo
notificação, menção e relatórios confiáveis por pessoa.

**Escopo:**
- Definir (ou reaproveitar, se já existir em outro módulo como `ponto-*.js`)
  uma coleção de usuários com pelo menos `id`, `nome`, `email`, `avatarUrl`.
- Migração: ao normalizar `negocio`/`anotacao`, se `responsavel` for uma
  string que corresponda ao nome de um usuário cadastrado, converter para
  `responsavelId`; caso não corresponda a nenhum usuário, manter o texto
  antigo em `responsavelLegado` (não descartar dado histórico — mesmo padrão
  já usado para `destinatario` em `normalizarAnotacao`).
- Atualizar `CAMPOS_AUDITAVEIS_NEGOCIO` e `CAMPOS_AUDITAVEIS_ANOTACAO` para
  trocar `responsavel` por `responsavelId`.
- Atualizar filtros e ordenações (`filtrarNegocios`, `filtrarAnotacoes`,
  `ordenarNegocios`, `ordenarAnotacoes`) que hoje comparam `responsavel` como
  string, para comparar por `responsavelId` e exibir o nome resolvido.
- Seletor de responsável na UI (crm-ui.js) passa a ser um combobox de
  usuários cadastrados, com avatar, em vez de um campo de texto.
- Ao atribuir um negócio/demanda a alguém, disparar uma notificação interna
  (reaproveitando o sistema de notificações já citado em crm-ui.js —
  `Notifications`).

**Arquivos provavelmente afetados:** crm-model.js (normalização e validação),
crm-calculos.js (filtros/ordenação), crm-ui.js (seletor e exibição), crm-store.js
(se houver leitura de coleção de usuários), testes correspondentes em `testes/`.

**Critérios de aceite:**
- Nenhum negócio/demanda existente perde o responsável ao migrar.
- Filtrar por responsável funciona por usuário (não sensível a erro de
  digitação de nome).
- Notificação é enviada quando o `responsavelId` de um item muda.

---

## 3. Central de atualizações/comentários por item

**Objetivo:** um mural de comentários livre por `negocio` e por `anotacao`,
separado do histórico técnico de auditoria (`historico`), para a equipe
discutir o andamento.

**Escopo:**
- Nova entidade `comentario`: `id`, `entidade` ('negocio'|'anotacao'),
  `entidadeId`, `autor` (idealmente `autorId`, ver item 2), `texto`,
  `criadoEm`, `editadoEm`, `mencoes` (lista de `usuarioId`s citados com `@`).
- Função `normalizarComentario` em crm-model.js, seguindo o mesmo padrão de
  `normalizarHistoricoItem`.
- `comentarios` como novo array na raiz de `normalizarCrm` (junto de
  `historico`), com validação e limpeza de comentários órfãos (entidade
  inexistente) — mesmo tratamento dado a atividades órfãs.
- UI: aba "Comentários" no modal de detalhe do negócio/demanda, ao lado da
  aba de histórico já existente, com campo de texto, suporte a `@menção` e
  ordenação cronológica.
- Notificação para usuários mencionados.

**Arquivos provavelmente afetados:** crm-model.js, crm-store.js, crm-ui.js,
crm.css, testes/.

**Critérios de aceite:**
- Comentários não se misturam com o histórico de auditoria de campos.
- Excluir um negócio/demanda (soft delete via `excluidoEm`) preserva os
  comentários para eventual restauração, mas eles não aparecem em listagens
  ativas.
- Testes cobrindo criação, edição e menção.

---

## 4. Exportação e automações com histórico de execução

**Objetivo:** permitir exportar dados das visões (Lista, Previsão) e registrar
um log de execução das reconciliações automáticas já existentes
(`reconciliarThreadsRespondidas`, `reconciliarEtapasComDemandas`), para
auditar por que um negócio ou demanda mudou de estado sozinho.

**Escopo:**
- Botão "Exportar" na visão Lista (negócios e anotações) gerando CSV/XLSX com
  as colunas atualmente visíveis e respeitando os filtros aplicados
  (reaproveitar `filtrarNegocios`/`filtrarAnotacoes` e `ordenarNegocios`/
  `ordenarAnotacoes` já existentes).
- Novo array `logAutomacoes` (ou reaproveitar `historico` com
  `tipo: 'automacao'`) registrando cada vez que
  `reconciliarThreadsRespondidas` ou `reconciliarEtapasComDemandas` alterar
  algo, com: entidade afetada, campo alterado, valor antes/depois, motivo
  (ex.: "todas as demandas do negócio foram respondidas") e timestamp.
- Aba "Histórico de automações" na UI, filtrável por entidade e por período.

**Arquivos provavelmente afetados:** crm-calculos.js ou crm-model.js (gerar
entradas de log dentro das próprias funções de reconciliação), crm-store.js
(persistir o log), crm-ui.js (botão de exportar e aba de histórico), testes/.

**Critérios de aceite:**
- Exportar respeita os filtros ativos no momento do clique.
- Toda alteração automática de etapa/situação feita pelas funções de
  reconciliação gera uma entrada rastreável no log.
- Nenhuma automação silenciosa deixa de ser registrada.

---

## 5. Implantação: PWA instalável + notificações Web Push

**Objetivo:** aproximar a experiência de uso da de um app hospedado (como o
Monday), sem sair do modelo atual de site estático (GitHub Pages) + Cloudflare
Worker + Firestore.

**Escopo — PWA:**
- Adicionar `manifest.json` (nome, ícones, cor de tema, `display: standalone`)
  e referenciá-lo em index.html.
- Adicionar um `service worker` básico para cache de assets estáticos
  (app shell) e funcionamento offline mínimo (visualização de dados já
  carregados).
- Garantir que o site sirva por HTTPS (já é o caso no GitHub Pages) e passe
  nos critérios de instalabilidade do navegador.

**Escopo — Web Push:**
- Endpoint no Cloudflare Worker para registrar `PushSubscription` dos
  usuários (uma nova coleção no Firestore, ex.: `pushSubscriptions`, com
  regras de acesso restritas ao próprio usuário — revisar
  `regras do firestore` / `firestore.rules`).
- Disparo de notificação push quando: `semaforoPrazo` de uma anotação virar
  `'vencida'` ou `'alerta'`; uma atividade (`proximaAtividade`) estiver
  marcada para o dia atual e ainda não tiver sido feita; um comentário
  mencionar o usuário (ver item 3).
- Job/rotina (Cloudflare Worker com Cron Trigger) que roda periodicamente
  verificando prazos e atividades pendentes e dispara os pushes — hoje esses
  alertas só aparecem quando alguém abre o sistema.

**Arquivos provavelmente afetados:** index.html, novo manifest.json, novo
service-worker.js, cloudflare-worker/ (novo endpoint + cron trigger), regras
do firestore, google-calendar-sync.js (caso as atividades importadas do
Google também devam gerar push).

**Critérios de aceite:**
- O site pode ser "instalado" a partir do navegador (ícone na tela inicial).
- Usuário que autorizar notificações recebe um push quando um prazo vira
  crítico, mesmo com a aba fechada.
- Revogar a permissão de notificação no navegador remove a inscrição sem
  gerar erros no worker.

---

## Observações gerais para o Claude Code
- Manter todas as funções de crm-model.js e crm-calculos.js puras e testáveis
  isoladamente com Vitest, como já é o padrão do projeto.
- Toda migração de dado antigo deve ser idempotente e não descartar dados
  históricos (seguir o padrão já usado para `destinatario` → `encaminhamentos`).
- Não introduzir dependências externas pesadas para os itens 1 a 4; para o
  item 5 (Web Push), usar a Web Push API padrão do navegador.
