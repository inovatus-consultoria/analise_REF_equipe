# Controle HH Static Web

Aplicacao 100% web/client-side para processar o REF puro diretamente no navegador.

Este e o modelo equivalente ao app de ferias: um front-end estatico que pode ser hospedado em GitHub Pages, Azure Static Web Apps ou outro host de arquivos estaticos. Nao ha backend, API propria, servidor Python, banco de dados ou upload do REF para servidor.

## Como funciona

1. Usuario acessa o HTML pelo navegador.
2. Opcionalmente autentica via Microsoft Entra ID usando MSAL/PKCE.
3. Seleciona o REF puro `.xlsx`.
4. O JavaScript le o arquivo localmente com SheetJS.
5. O app extrai Base, C, D e F.
6. Gera calendario, meses trabalhados, meses recebidos, comparacoes, alertas e F tratada.
7. O Excel final e baixado no proprio navegador.

## Arquivos

```text
index.html
styles.css
app.js
```

## Deploy

Publique os tres arquivos em um repositório GitHub Pages ou em Azure Static Web Apps.

No Azure Entra ID, cadastre a aplicacao como SPA e adicione a URL final como Redirect URI.

Depois atualize o bloco `APP_CONFIG` em `app.js`:

```js
const APP_CONFIG = {
  authRequired: false,
  tenantId: "",
  clientId: "",
  redirectUri: window.location.href.split("#")[0],
};
```

`clientId` e `tenantId` sao parametros publicos de SPA. Nao coloque client secret, token ou credencial privada no codigo.

## Execucao local

Abra `index.html` no navegador ou sirva a pasta com qualquer servidor estatico.

```powershell
python -m http.server 8080
```

Depois acesse `http://localhost:8080`.

