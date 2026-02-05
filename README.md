# Painel de Convites (Google Sheets) — GitHub Pages

## Publicar
1. Suba estes arquivos para um repositório do GitHub.
2. **Settings → Pages**: escolha branch `main` e pasta `/ (root)`.
3. Abra o link do GitHub Pages.

## Fonte de dados
O site lê do link publicado (pubhtml) e tenta baixar em CSV automaticamente:
https://docs.google.com/spreadsheets/d/e/2PACX-1vR8tuofnr7wajPMUtsI86RNmVGvuDLUtUwmT-qANASxOlN2C9qPkoDtIBzpGUyhEUvk5ijj8umZsj3U/pubhtml

Para trocar o link, edite `index.html` e altere `window.__SHEET_PUBHTML__`.
