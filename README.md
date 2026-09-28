# Painel de Absenteísmo — SMS Cajamar

Site estático (sem build, sem servidor). Arquivos: index.html, style.css, app.js, logo.png (logotipo oficial), vendor/xlsx.mini.min.js (leitura de .xlsx) e _headers (cabeçalhos de segurança do Cloudflare Pages).

## Publicação no Cloudflare Pages
1. Painel Cloudflare > Workers & Pages > Create > Pages > Upload assets.
2. Nomeie o projeto e envie esta pasta (ou o .zip, com index.html na raiz).
3. Deploy. O endereço será <projeto>.pages.dev.
Alternativa por linha de comando: `npx wrangler pages deploy . --project-name=painel-absenteismo`

## Atualização dos dados
Na própria página, botão "Selecionar relatórios": enviar juntos os relatórios de absenteísmo e de atendimentos (.xlsx ou .csv) do mesmo período e com os mesmos filtros. O sistema identifica cada relatório pelas colunas. O processamento ocorre no navegador; apenas totais agregados são salvos no localStorage do dispositivo (a atualização vale para quem a fez, não para os demais usuários).
Para atualizar o painel para todos, substitua os dados iniciais (constante DEFAULT_MODEL em app.js) e publique novamente.
Se o relatório incluir a coluna "Hora", os painéis por hora são exibidos automaticamente.

## Acesso restrito (recomendado)
Ativar Cloudflare Access (Zero Trust) sobre o domínio do projeto para limitar a consulta a e-mails institucionais.

## Identidade visual
Cores, tipografia (Montserrat) e logotipo seguem o Manual de Identidade Visual da Prefeitura de Cajamar: azul #2745A1 (principal), vermelho #E01A36, ciano #1AB5F1 e azul-acinzentado #7BABBF. Na aba de taxa, vermelho = 35% ou mais, ciano = 25% a 35%, azul = abaixo de 25%.
