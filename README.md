# Painel de Absenteísmo — SMS Cajamar

Site estático (sem build, sem servidor). Arquivos: index.html, style.css, app.js, logo.png (logotipo oficial), data.js (dados iniciais), vendor/xlsx.mini.min.js (leitura de .xlsx) e _headers (cabeçalhos de segurança do Cloudflare Pages).

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

## Filtros e cruzamento por faixa etária
- Os dados de absenteísmo agora são carregados em nível de linha (uma linha por combinação de unidade, categoria, turno, tipo de falta, faixa etária e tempo entre agendamento e consulta), em `data.js`. Isso permite recalcular tudo sob filtro, em vez de usar apenas totais fechados.
- O painel "Filtros", no topo, tem checkboxes por unidade, categoria profissional, turno e faixa etária (com atalhos por faixa agrupada: Crianças, Adolescentes, Adultos jovens, Adultos, Idosos). Os filtros recalculam os KPIs, os gráficos e o cruzamento na aba "Volume de faltas" em tempo real, no navegador.
- A aba "Volume de faltas" ganhou um cruzamento (heatmap) de **faixa etária × categoria profissional**, mostrando o número de faltas por combinação dentro do recorte filtrado.
- Limitações importantes, mantidas visíveis na própria tela:
  - O cruzamento mostra **volume de faltas**, não taxa percentual — o relatório de atendimentos do e-SUS não traz faixa etária, então não dá para calcular "faltas ÷ agendados" por idade.
  - "Categoria profissional" é a classificação do e-SUS (médico, enfermeiro, cirurgião dentista, psicólogo, nutricionista, outro superior, auxiliar/técnico de enfermagem) — não é a especialidade médica (pediatria, cardiologia etc.), que o relatório não traz.
  - As faixas etárias do e-SUS agrupam "05 a 09 anos" como uma faixa única; não é possível isolar exatamente "até 6 anos" sem essa granularidade no relatório de origem. A faixa detalhada por idade permite compor "0 a 4 anos" exatamente; para incluir também os de 5 e 6 anos, é preciso incluir a faixa inteira "05 a 09 anos".
  - A aba "Taxa percentual" continua usando sempre a base completa dos dois relatórios (não é afetada pelos filtros), porque o cálculo de taxa depende de casar as mesmas dimensões nos dois relatórios, e isso só é possível hoje por unidade e por categoria profissional.

## Histórico de períodos (cargas periódicas)
- Cada vez que você envia um novo par de relatórios (absenteísmo + atendimentos), o período é **adicionado** ao histórico — não substitui os anteriores. Um seletor no topo da tela deixa escolher qual período ver nas abas "Volume de faltas" e "Taxa percentual".
- A aba **Histórico** mostra a evolução: faltas por período, taxa global por período (quando houver o relatório de atendimentos daquele período), uma frase comparando o último período com o anterior, e uma tabela com tudo.
- Por padrão, esse histórico fica salvo **no navegador de quem fez o upload** (via `localStorage`, chave `painel-absenteismo-v2`). Ele não é visível a outras pessoas que abram o mesmo link em outro navegador ou dispositivo.
- **Para que o histórico apareça igual para todo mundo**, sem precisar pedir apoio técnico a cada carga:
  1. Faça a carga normalmente pela tela (relatórios de absenteísmo + atendimentos do novo período).
  2. Clique em **"Baixar dados atualizados (data.js)"**. Isso gera um arquivo `data.js` já com todos os períodos (os antigos + o novo).
  3. Substitua o arquivo `data.js` do seu projeto (o mesmo nome, na raiz) e publique de novo no Cloudflare Pages — por `git push` (se o projeto estiver num repositório) ou arrastando os arquivos na aba "Deployments" do painel do Cloudflare Pages.
  4. Nenhum código precisa mudar: é só a troca desse único arquivo.
- **"Restaurar dados iniciais"** apaga o histórico salvo neste navegador e volta ao(s) período(s) que vieram junto com o site (pede confirmação antes).

## Modelo de carga
O arquivo `modelo-de-carga-esus.xlsx`, incluído neste pacote e linkado na própria tela de upload, mostra as colunas exatas que cada relatório precisa ter (com uma linha de exemplo), para conferência antes de carregar um relatório exportado do e-SUS APS. Ele é só referência — não deve ser enviado ao painel como se fosse um relatório real.
