# alfred-m365 — istruzioni per Claude

Plugin Hermes/Alfred per Outlook (Microsoft 365) di UNA casella (il CEO): lettura posta e calendario, bozze
(mai inviate), creazione/modifica appuntamenti con approvazione. Stdlib only, struttura copiata da alfred-odoo
(`~/dev/alfred-odoo`, stesse invarianti: niente `pyproject.toml`, import relativi, `register()` senza I/O,
segreti via `agent.secret_scope`, hook `pre_tool_call`, scansione "safe").

## Comandi
- `uv run --no-project --with pytest pytest -q tests` (fake Graph in `tests/conftest.py`), `uvx ruff check .`
- Contratti: dall'albero Alfred `../alfred-venv/bin/python -m hermes_cli.main plugins doctor <repo> --ci` e `plugins validate`.
- End-to-end: `uv run --no-project --with pytest python dev/fake_graph.py 8765` + `HERMES_HOME` usa-e-getta con il plugin
  in `plugins/alfred-m365/`, setting `graph_base_url: http://127.0.0.1:8765/v1.0`, `.env` con i tre `MSGRAPH_*`
  (secret `s3cr3t-value-for-tests-only`). Modelli piccoli (qwen2.5:14b) servono `tools.tool_search.enabled: "off"` e `-t outlook`.
  **`alfred -z` è YOLO: salta tutte le approvazioni** (by design upstream); per provare il gate usare `alfred chat -q`
  (nega) o la CLI interattiva/Telegram.
- Live su tenant vero: `dev/live_check.py` (crea bozza + evento di prova, non cancella nulla).

## Invarianti
- Exchange non ha "solo bozze" né "niente cancellazioni": il limite è `graph.ALLOWED` + `GraphClient._call`, l'unico
  punto di passaggio. Aggiungere un'operazione = una riga in `ALLOWED` + un ramo nel `FakeGraph` + un test che le
  operazioni vietate (`tests/test_graph.py`) restano rifiutate senza rete. Mai DELETE, `/send`, `/reply`, `/forward`,
  `/move`, `/copy`, `sendMail`, altre caselle.
- PATCH messaggi solo se `isDraft` (GET prima); PATCH eventi solo se `isOrganizer` e mai `seriesMaster`.
- Approvazione: eventi sempre; bozze solo con `approve_drafts`. Il `message` (italiano) descrive la modifica dai soli
  argomenti, avvisa degli inviti "a nome di <owner_name>". Hook senza rete, su errore `block`.
- Testi delle mail = `untrusted_text` + `content_notice` (prompt injection). Destinatari: nomi risolti tra i
  corrispondenti; indirizzi mai visti → `warning` (i modelli inventano indirizzi plausibili).
- Orari: input locali nel `timezone` della casella; Graph riceve nomi Windows (`graph.WINDOWS_TZ`) e `Prefer:
  outlook.timezone` su ogni lettura. `graph_base_url` accetta solo loopback (il bearer token va lì).
- Il client secret non deve mai comparire in output o errori (test dedicato).
