(function(){
'use strict';
const DEFAULT_MODEL = {"periodo": "01/03/2026 a 31/08/2026", "gerado_em": "25/09/2026", "origem": "dados iniciais", "abs": {"total": 44788, "tempoMedio": 24.1, "dims": {"unidade": {"Policlinica Municipal de Cajamar": 6209, "Usf Manoel Inacio da Silva": 4949, "UBS Enf Leontina Martins Franca": 4902, "UBS Dra Izabel Gratieri": 4641, "ESF Carlos dos Santos": 4040, "UBS Enfermeiro Carlos Moreira da Silva": 3511, "Usf Vereador Joaquim Alves de Castro": 3491, "Centro Especializado em Reabilitacao Cer II": 2605, "Psf Belo Planalto": 2312, "Psf Dra Maria de Lourdes Mendonca Bravo": 1565, "Usf Maria Aparecida Misse": 1426, "Posto de Saude Nadilia de Oliveira Santos": 1374, "Psf Edivaldo Soares Massagardi": 1201, "CEO Centro de Especialidades Odontologicas": 1029, "Caps Cajamar": 803, "Caps Infantil Cajamar": 525, "Sae Cta Cajamar": 205}, "categoria": {"Médico": 19248, "Cirurgião dentista": 10839, "Enfermeiro": 9963, "Outro prof. nível superior": 2251, "Psicólogo": 1565, "Nutricionista": 632, "Auxiliar ou técnico de enfermagem": 290}, "turno": {"Manhã": 24483, "Tarde": 20049, "Noite": 256}, "tipo": {"Não compareceu": 44544, "Não aguardou": 244}, "estagio": {"Crianças (0–9 anos)": 6704, "Adolescentes (10–19 anos)": 5794, "Adultos jovens (20–39 anos)": 13829, "Adultos (40–59 anos)": 12405, "Idosos (60+ anos)": 6056}, "tempo": {"0 a 6 dias": 14516, "7 a 13 dias": 5886, "14 a 20 dias": 4233, "21 a 27 dias": 3776, "28 a 34 dias": 3668, "35 a 41 dias": 2878, "42 a 48 dias": 2372, "49 a 55 dias": 2095, "56 a 62 dias": 1974, "63 a 69 dias": 986, "70 ou mais dias": 2404}}}, "at": {"total": 110724, "dims": {"unidade": {"Psf Dra Maria de Lourdes Mendonca Bravo": 2480, "Usf Manoel Inacio da Silva": 8047, "Usf Vereador Joaquim Alves de Castro": 5970, "UBS Enf Leontina Martins Franca": 8587, "ESF Carlos dos Santos": 7236, "Psf Belo Planalto": 4485, "UBS Enfermeiro Carlos Moreira da Silva": 7766, "UBS Dra Izabel Gratieri": 10732, "Caps Infantil Cajamar": 1350, "Posto de Saude Nadilia de Oliveira Santos": 3623, "Usf Maria Aparecida Misse": 4007, "Psf Edivaldo Soares Massagardi": 3889, "Caps Cajamar": 2891, "Centro Especializado em Reabilitacao Cer II": 9500, "CEO Centro de Especialidades Odontologicas": 3860, "Policlinica Municipal de Cajamar": 25362, "Sae Cta Cajamar": 939}, "categoria": {"Auxiliar ou técnico de enfermagem": 0, "Nutricionista": 666, "Enfermeiro": 17614, "Cirurgião dentista": 24915, "Psicólogo": 3910, "Médico": 56332, "Outro prof. nível superior": 7287}}}};
const LS_KEY = 'painel-absenteismo-v1';
const $ = id => document.getElementById(id);
const norm = s => String(s == null ? '' : s).normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().trim();
const num = v => typeof v === 'number' ? v : (Number(String(v).replace(/\./g, '').replace(',', '.')) || 0);
const fmt = n => n.toLocaleString('pt-BR');
const pct = v => v.toFixed(1).replace('.', ',') + '%';
const firstNum = s => { const m = String(s).match(/\d+/); return m ? +m[0] : 1e9; };
const add = (d, k, v) => { d[k] = (d[k] || 0) + v; };
const esc = s => String(s).replace(/[&<>"]/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));

/* ---------- leitura de arquivos ---------- */
function parseCSV(text){
  text = text.replace(/^\uFEFF/, '');
  const first = text.split(/\r?\n/, 40).join('\n');
  const delim = (first.match(/;/g) || []).length >= (first.match(/,/g) || []).length ? ';' : ',';
  const rows = []; let row = [], cur = '', q = false;
  for (let i = 0; i < text.length; i++){
    const ch = text[i];
    if (q){ if (ch === '"'){ if (text[i+1] === '"'){ cur += '"'; i++; } else q = false; } else cur += ch; }
    else if (ch === '"') q = true;
    else if (ch === delim){ row.push(cur); cur = ''; }
    else if (ch === '\n'){ row.push(cur.replace(/\r$/, '')); rows.push(row); row = []; cur = ''; }
    else cur += ch;
  }
  if (cur.length || row.length){ row.push(cur.replace(/\r$/, '')); rows.push(row); }
  return rows;
}
async function readRows(file){
  const buf = await file.arrayBuffer();
  if (/\.csv$/i.test(file.name)) return parseCSV(new TextDecoder('utf-8').decode(buf));
  const wb = XLSX.read(buf, {type: 'array'});
  const ws = wb.Sheets[wb.SheetNames[0]];
  const r = XLSX.utils.decode_range(ws['!ref']);
  r.e.c = Math.min(r.e.c, 15);
  ws['!ref'] = XLSX.utils.encode_range(r);
  return XLSX.utils.sheet_to_json(ws, {header: 1, raw: true, defval: ''});
}

function parseReport(rows, name){
  let h = -1;
  for (let i = 0; i < Math.min(rows.length, 120); i++){
    const r = rows[i].map(norm);
    if (r[0] === 'unidade de saude' && r.indexOf('total') >= 0){ h = i; break; }
  }
  if (h < 0) throw new Error('Cabeçalho da tabela não localizado em "' + name + '".');
  const cols = rows[h].map(norm), ix = n => cols.indexOf(n);
  const kind = ix('tipo de falta') >= 0 ? 'abs' : (ix('tipo de atendimento') >= 0 ? 'at' : null);
  if (!kind) throw new Error('"' + name + '" não corresponde ao relatório de absenteísmo nem ao de atendimentos.');
  const c = {unidade: 0, categoria: ix('categoria profissional'), turno: ix('turno'), hora: ix('hora'), tipo: ix('tipo de falta'),
             faixa: ix('faixa etaria'), tempo: ix('tempo entre agendamento e consulta'), total: ix('total')};
  let periodo = null;
  for (let i = 0; i < h && !periodo; i++){
    const m = rows[i].join(' ').match(/(\d{2}\/\d{2}\/\d{4})\D+(\d{2}\/\d{2}\/\d{4})/);
    if (m) periodo = m[1] + ' a ' + m[2];
  }
  const dims = {unidade: {}, categoria: {}}, extra = {};
  if (c.turno >= 0) dims.turno = {};
  if (c.hora >= 0) dims.hora = {};
  if (kind === 'abs'){ dims.tipo = {}; dims.estagio = {}; dims.tempo = {}; }
  let total = 0, wsum = 0;
  for (let i = h + 1; i < rows.length; i++){
    const r = rows[i], u = String(r[0] == null ? '' : r[0]).trim();
    if (!u) continue;
    if (norm(u) === 'total') break;
    const v = num(r[c.total]); if (!v) continue;
    total += v;
    add(dims.unidade, u, v);
    if (c.categoria >= 0) add(dims.categoria, String(r[c.categoria]).trim(), v);
    if (dims.turno) add(dims.turno, String(r[c.turno]).trim(), v);
    if (dims.hora) add(dims.hora, String(r[c.hora]).trim(), v);
    if (kind === 'abs'){
      if (c.tipo >= 0) add(dims.tipo, String(r[c.tipo]).trim(), v);
      if (c.faixa >= 0) add(dims.estagio, estagio(r[c.faixa]), v);
      if (c.tempo >= 0){
        const t = String(r[c.tempo]).trim(); add(dims.tempo, t, v);
        const n = (t.match(/\d+/g) || []).map(Number);
        wsum += v * (n.length >= 2 ? (n[0] + n[1]) / 2 : (n[0] || 0) + 10);
      }
    }
  }
  if (!total) throw new Error('Nenhum registro válido encontrado em "' + name + '".');
  return {kind, periodo, total, dims, tempoMedio: kind === 'abs' ? wsum / total : null};
}
function estagio(l){
  const n = norm(l), a = n.indexOf('menos') === 0 ? 0 : (n.match(/\d+/) ? +n.match(/\d+/)[0] : 0);
  return a < 10 ? 'Crianças (0–9 anos)' : a < 20 ? 'Adolescentes (10–19 anos)' : a < 40 ? 'Adultos jovens (20–39 anos)' : a < 60 ? 'Adultos (40–59 anos)' : 'Idosos (60+ anos)';
}
const ORDEM_ESTAGIO = ['Crianças (0–9 anos)','Adolescentes (10–19 anos)','Adultos jovens (20–39 anos)','Adultos (40–59 anos)','Idosos (60+ anos)'];

/* ---------- modelo ---------- */
function load(){ try { const s = localStorage.getItem(LS_KEY); if (s) return JSON.parse(s); } catch(e){} return DEFAULT_MODEL; }
function save(m){ try { localStorage.setItem(LS_KEY, JSON.stringify(m)); } catch(e){} }
let model = load();

/* ---------- renderização ---------- */
const toItems = (d, sort) => Object.keys(d).map(k => ({label: k, value: d[k]})).sort(sort || ((a, b) => b.value - a.value));
function volBars(id, items, total, amberIf){
  const max = Math.max(1, ...items.map(i => i.value));
  $(id).innerHTML = items.map(i => `<div class="bar-row"><div class="lbl" title="${esc(i.label)}">${esc(i.label)}</div><div class="bar-track"><div class="bar-fill ${amberIf && amberIf(i) ? 'amber' : ''}" style="width:${(i.value / max * 100).toFixed(1)}%"></div></div><div class="val"><span class="primary">${fmt(i.value)}</span><span class="secondary">${pct(i.value / total * 100)}</span></div></div>`).join('');
}
const rc = t => t >= 35 ? 'red' : t >= 25 ? 'amber' : '';
function rateBars(id, rows){
  const max = Math.max(1, ...rows.map(r => r.taxa));
  $(id).innerHTML = rows.map(r => `<div class="bar-row wide"><div class="lbl" title="${esc(r.label)}">${esc(r.label)}</div><div class="bar-track"><div class="bar-fill ${rc(r.taxa)}" style="width:${(r.taxa / max * 100).toFixed(1)}%"></div></div><div class="val"><span class="primary">${pct(r.taxa)}</span><span class="secondary">${fmt(r.absenteismo)} / ${fmt(r.agendado)}</span></div></div>`).join('');
}
function rates(A, T){
  const keys = new Set(Object.keys(A).concat(Object.keys(T))), rows = [], excl = [];
  keys.forEach(k => {
    const a = A[k] || 0, t = T[k] || 0;
    if (!t && a){ excl.push({label: k, absenteismo: a}); return; }
    if (a + t) rows.push({label: k, atendimentos: t, absenteismo: a, agendado: a + t, taxa: a / (a + t) * 100});
  });
  return {rows: rows.sort((x, y) => y.taxa - x.taxa), excl};
}
const kpi = (n, l) => `<div class="kpi"><div class="num">${n}</div><div class="lbl">${l}</div></div>`;
function show(id, on, html){ const e = $(id); e.hidden = !on; if (html != null) e.innerHTML = html; }

function render(){
  const m = model, A = m.abs, D = A.dims;
  $('periodo').textContent = m.periodo || 'não identificado';
  $('gerado').textContent = m.gerado_em;
  $('origem').textContent = m.origem;

  /* aba volume */
  const topU = toItems(D.unidade)[0], turno = toItems(D.turno || {});
  const manha = turno.find(t => norm(t.label) === 'manha');
  $('kpis-volume').innerHTML = kpi(fmt(A.total), 'Faltas registradas no período') +
    kpi(A.tempoMedio != null ? A.tempoMedio.toFixed(1).replace('.', ',') : '—', 'Dias, em média, entre agendamento e a falta') +
    kpi(manha ? pct(manha.value / A.total * 100) : '—', 'Das faltas ocorrem no turno da manhã') +
    kpi(fmt(topU.value), 'Faltas na unidade de maior volume — ' + esc(topU.label));
  const tempo = toItems(D.tempo || {}, (a, b) => firstNum(a.label) - firstNum(b.label));
  volBars('c-unidade', toItems(D.unidade), A.total, i => i === toItems(D.unidade)[0]);
  volBars('c-categoria', toItems(D.categoria), A.total);
  volBars('c-turno', turno, A.total);
  volBars('c-tipo', toItems(D.tipo || {}), A.total);
  volBars('c-estagio', ORDEM_ESTAGIO.filter(k => D.estagio && D.estagio[k]).map(k => ({label: k, value: D.estagio[k]})), A.total);
  volBars('c-tempo', tempo, A.total, i => i === tempo[0]);
  const t0 = tempo[0];
  show('n-tempo', !!t0, t0 ? `Faltas no intervalo de ${esc(t0.label)} correspondem a ${pct(t0.value / A.total * 100)} do total. Concentração nos primeiros dias após o agendamento sugere falha de confirmação de curto prazo, e não apenas esquecimento de longo prazo.` : '');
  show('p-hora', !!D.hora);
  if (D.hora) volBars('c-hora', toItems(D.hora, (a, b) => firstNum(a.label) - firstNum(b.label)), A.total);
  show('n-turno', !D.hora, 'O relatório carregado discrimina apenas o turno. Para quebra por hora, reexporte o relatório do e-SUS incluindo o campo "Hora" e envie-o novamente.');

  /* aba taxa */
  const T = m.at;
  show('taxa-empty', !T); show('taxa-body', !!T);
  if (!T) return;
  const glob = A.total / (A.total + T.total) * 100;
  $('kpis-taxa').innerHTML = kpi(pct(glob), 'Taxa global de absenteísmo na rede') + kpi(fmt(A.total + T.total), 'Total de vagas agendadas no período') +
    kpi(fmt(T.total), 'Atendimentos realizados') + kpi(fmt(A.total), 'Faltas registradas');
  const ru = rates(D.unidade, T.dims.unidade), rcat = rates(D.categoria, T.dims.categoria);
  rateBars('t-unidade', ru.rows); rateBars('t-categoria', rcat.rows);
  if (ru.rows.length > 2){
    const absTop = topU.label, pos = ru.rows.findIndex(r => r.label === absTop) + 1;
    show('n-unidade', true, `A unidade de maior volume absoluto de faltas (${esc(absTop)}) ocupa a posição ${pos} de ${ru.rows.length} em taxa percentual` +
      (pos ? ` (${pct(ru.rows[pos - 1].taxa)})` : '') + `. As maiores taxas relativas são de ${esc(ru.rows[0].label)} (${pct(ru.rows[0].taxa)}) e ${esc(ru.rows[1].label)} (${pct(ru.rows[1].taxa)}). Para priorização de intervenção, a taxa percentual é o indicador tecnicamente mais adequado.`);
  } else show('n-unidade', false);
  const warn = (id, ex, tipo) => show(id, ex.length, ex.length ? `${tipo} sem atendimentos no relatório de atendimentos (filtrado por "Consulta Agendada"): ${ex.map(e => esc(e.label) + ' (' + fmt(e.absenteismo) + ' faltas)').join('; ')}. Excluídos da taxa por provável diferença de escopo entre os relatórios; verificar na fonte primária do e-SUS antes de qualquer conclusão.` : '');
  warn('w-unidade', ru.excl, 'Unidades'); warn('w-categoria', rcat.excl, 'Categorias');
  ['turno', 'hora'].forEach(k => {
    const ok = D[k] && T.dims[k]; show('p-t-' + k, !!ok);
    if (ok){ const r = rates(D[k], T.dims[k]).rows; if (k === 'hora') r.sort((a, b) => firstNum(a.label) - firstNum(b.label)); rateBars('t-' + k, r); }
  });
}

/* ---------- upload ---------- */
function status(msg, cls){ const s = $('status'); s.textContent = msg; s.className = 'status ' + (cls || ''); }
async function handleFiles(list){
  try {
    status('Processando…');
    const reps = [];
    for (const f of list) reps.push(parseReport(await readRows(f), f.name));
    const abs = reps.find(r => r.kind === 'abs'), at = reps.find(r => r.kind === 'at');
    if (!abs) throw new Error('Envie o relatório de absenteísmo (obrigatório), com ou sem o de atendimentos.');
    const p = abs.periodo || (at && at.periodo);
    let msg = 'Dados atualizados: ' + fmt(abs.total) + ' faltas' + (at ? ' e ' + fmt(at.total) + ' atendimentos' : '') + '.';
    let cls = 'ok';
    if (at && abs.periodo && at.periodo && abs.periodo !== at.periodo){ msg += ' ATENÇÃO: os períodos dos dois relatórios divergem (' + abs.periodo + ' e ' + at.periodo + '); a taxa não é confiável.'; cls = 'err'; }
    if (!at) msg += ' Sem o relatório de atendimentos, a aba de taxa permanece indisponível.';
    model = {periodo: p, gerado_em: new Date().toLocaleDateString('pt-BR'), origem: 'arquivos enviados',
             abs: {total: abs.total, dims: abs.dims, tempoMedio: abs.tempoMedio}, at: at ? {total: at.total, dims: at.dims} : null};
    save(model); render(); status(msg, cls);
  } catch(e){ status(e.message || String(e), 'err'); }
}
$('btn-upload').onclick = () => $('files').click();
$('files').onchange = e => { if (e.target.files.length) handleFiles(Array.from(e.target.files)); e.target.value = ''; };
$('btn-reset').onclick = () => { try { localStorage.removeItem(LS_KEY); } catch(e){} model = DEFAULT_MODEL; render(); status('Dados iniciais restaurados.', 'ok'); };
document.querySelectorAll('.tab-btn').forEach(b => b.onclick = () => {
  document.querySelectorAll('.tab-btn').forEach(x => x.classList.remove('active'));
  document.querySelectorAll('.tab-panel').forEach(x => x.classList.remove('active'));
  b.classList.add('active'); $('tab-' + b.dataset.tab).classList.add('active');
});
render();
})();
