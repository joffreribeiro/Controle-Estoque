# Copilot instructions for Controle-Estoque

## Project shape and runtime
- This is a static front-end app: `index.html` + `styles.css` + a single large script `app.js`.
- There is no bundler/module system; code runs in browser globals and is loaded with `<script src="app.js">`.
- External libs are loaded by CDN in `index.html` (Chart.js, jsPDF, jsPDF-AutoTable, SheetJS, Firebase compat SDKs).
- Keep solutions dependency-light and compatible with direct script loading.

## Core data model and persistence
- Primary in-memory object is global `estoque` in `app.js` (contains `produtos`, `registroVendas`, `registroDistribuicao`, `controleEnvio`, `auditoriaVendas`, `fechamentosComissoes`).
- Local persistence uses `localStorage` key `estoqueArmasV2` (`carregarDados()` / `salvarDados()`).
- `salvarDados()` is the canonical persistence entrypoint: it also updates stats and schedules cloud sync.
- IMBEL tab uses separate storage key `IMBEL_KEY` (inside `app.js`) and should not be mixed with `estoqueArmasV2` schema changes.

## Cloud sync, auth, and access control
- Firebase init and Firestore instance setup happen at top of `app.js`; use `window.firestoreDB` guard checks.
- Cloud backup document is fixed: collection `app_data`, document `latest` (`salvarNoCloud`, `carregarDoCloud`, `carregarDoCloudAuto`).
- Auth flow is Firebase Auth compat (`signIn`, `signOut`, `onAuthStateChanged` near end of `app.js`).
- Admin-only UI is controlled by `[data-admin="true"]` elements and `body.is-admin` class toggling.
- For privileged actions, follow existing `requireAdminOrNotify()` pattern instead of duplicating custom checks.
- Firestore rules file is a reference template (`firestore.rules`) and is not auto-deployed from this repo.

## UI and coding patterns
- App is tab-driven (`trocarAba`) with rendering functions like `renderizarTabela`, `renderizarRegistroVendas`, `renderizarDashboard`.
- After mutating stock/sales/distribution state, keep the current pattern: `salvarDados()` then refresh affected views/selects.
- Use existing notification helper `mostrarNotificacao(mensagem, tipo)` for user feedback.
- Prefer Portuguese naming/messages to match the current UI/domain language.

## Local development workflows
- Open `index.html` directly or via a simple static server for browser testing.
- No automated test suite or build scripts are defined in the repo; validate changes manually in browser.
- `save_server.js` is intentionally a placeholder (local save server removed); do not reintroduce localhost save/load endpoints unless explicitly requested.
- `README_SAVE_LOCAL.md` documents an old local-server workflow and can be outdated relative to current code.
- For Firebase admin claim setup, use `scripts/set-admin-claim.js` with `firebase-admin` and a service account JSON.

## Change scope guidance
- Prefer small, surgical edits in `app.js`; avoid broad refactors because many features share global state.
- Preserve existing localStorage keys and document schema (`app_data/latest`) unless migration logic is added.
- When adding buttons/actions in `index.html`, wire to existing `app.js` function style (`onclick="..."`) used across the project.