(function(){
'use strict';
const LS_KEY = 'painel-absenteismo-v2';
const $ = id => document.getElementById(id);
const norm = s => String(s == null ? '' : s).normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().trim();
const num = v => typeof v === 'number' ? v : (Number(String(v).replace(/\./g, '').replace(',', '.')) || 0);
const fmt = n => Math.round(n).toLocaleString('pt-BR');
const pct = v => v.toFixed(1).replace('.', ',') + '%';
const firstNum = s => { const m = String(s).match(/\d+/); return m ? +m[0] : 1e9; };
const esc = s => String(s).replace(/[&<>"]/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
const TURNO_ORDER = {'manha': 0, 'tarde': 1, 'noite': 2};
const ORDEM_ESTAGIO = ['Crianças (0–9 anos)', 'Adolescentes (10–19 anos)', 'Adultos jovens (20–39 anos)', 'Adultos (40–59 anos)', 'Idosos (60+ anos)'];
function estagio(l){
  const n = norm(l), a = n.indexOf('menos') === 0 ? 0 : (n.match(/\d+/) ? +n.match(/\d+/)[0] : 0);
  return a < 10 ? ORDEM_ESTAGIO[0] : a < 20 ? ORDEM_ESTAGIO[1] : a < 40 ? ORDEM_ESTAGIO[2] : a < 60 ? ORDEM_ESTAGIO[3] : ORDEM_ESTAGIO[4];
}
function ageKey(l){ const n = norm(l); return n.indexOf('menos') === 0 ? 0 : firstNum(n); }
function tempoMid(l){ const n = (String(l).match(/\d+/g) || []).map(Number); return n.length >= 2 ? (n[0] + n[1]) / 2 : (n[0] || 0) + 10; }

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
  let periodo = null;
  for (let i = 0; i < h && !periodo; i++){
    const m = rows[i].join(' ').match(/(\d{2}\/\d{2}\/\d{4})\D+(\d{2}\/\d{2}\/\d{4})/);
    if (m) periodo = m[1] + ' a ' + m[2];
  }

  if (kind === 'at'){
    const c = {categoria: ix('categoria profissional'), turno: ix('turno'), hora: ix('hora'), total: ix('total')};
    const dims = {unidade: {}, categoria: {}}; if (c.turno >= 0) dims.turno = {}; if (c.hora >= 0) dims.hora = {};
    let total = 0;
    for (let i = h + 1; i < rows.length; i++){
      const r = rows[i], u = String(r[0] == null ? '' : r[0]).trim();
      if (!u || norm(u) === 'total') { if (u) break; continue; }
      const v = num(r[c.total]); if (!v) continue;
      total += v;
      dims.unidade[u] = (dims.unidade[u] || 0) + v;
      if (c.categoria >= 0){ const k = String(r[c.categoria]).trim(); dims.categoria[k] = (dims.categoria[k] || 0) + v; }
      if (dims.turno){ const k = String(r[c.turno]).trim(); dims.turno[k] = (dims.turno[k] || 0) + v; }
      if (dims.hora){ const k = String(r[c.hora]).trim(); dims.hora[k] = (dims.hora[k] || 0) + v; }
    }
    if (!total) throw new Error('Nenhum registro válido encontrado em "' + name + '".');
    return {kind, periodo, total, dims};
  }

  const c = {categoria: ix('categoria profissional'), turno: ix('turno'), tipo: ix('tipo de falta'),
             faixa: ix('faixa etaria'), tempo: ix('tempo entre agendamento e consulta'), total: ix('total')};
  const temp = [];
  for (let i = h + 1; i < rows.length; i++){
    const r = rows[i], u = String(r[0] == null ? '' : r[0]).trim();
    if (!u || norm(u) === 'total') { if (u) break; continue; }
    const v = num(r[c.total]); if (!v) continue;
    const cat = String(r[c.categoria]).trim(), tur = String(r[c.turno]).trim(), tip = String(r[c.tipo]).trim(),
          fx = String(r[c.faixa]).trim(), te = String(r[c.tempo]).trim();
    if (!cat || !tur || !tip || !fx || !te) continue;
    temp.push({u, cat, tur, tip, fx, te, v});
  }
  if (!temp.length) throw new Error('Nenhum registro válido encontrado em "' + name + '".');
  const dictSet = (arr, key, order) => { const s = Array.from(new Set(arr.map(r => r[key]))); return order ? s.sort(order) : s.sort(); };
  const dict = {
    unidade: dictSet(temp, 'u'),
    categoria: dictSet(temp, 'cat'),
    turno: dictSet(temp, 'tur', (a, b) => (TURNO_ORDER[norm(a)] ?? 9) - (TURNO_ORDER[norm(b)] ?? 9)),
    tipo: dictSet(temp, 'tip'),
    faixa: dictSet(temp, 'fx', (a, b) => ageKey(a) - ageKey(b)),
    tempo: dictSet(temp, 'te', (a, b) => firstNum(a) - firstNum(b))
  };
  const idx = {}; Object.keys(dict).forEach(k => { idx[k] = {}; dict[k].forEach((v, i) => idx[k][v] = i); });
  let total = 0, wsum = 0;
  const dataRows = temp.map(r => {
    total += r.v; wsum += r.v * tempoMid(r.te);
    return [idx.unidade[r.u], idx.categoria[r.cat], idx.turno[r.tur], idx.tipo[r.tip], idx.faixa[r.fx], idx.tempo[r.te], r.v];
  });
  return {kind, periodo, total, tempoMedio: wsum / total, dict, rows: dataRows};
}

/* ---------- modelo / histórico de períodos ---------- */
function periodKey(snap){
  const m = String(snap.periodo || '').match(/(\d{2})\/(\d{2})\/(\d{4})/);
  if (m) return m[3] + m[2] + m[1];
  const g = String(snap.gerado_em || '').match(/(\d{2})\/(\d{2})\/(\d{4})/);
  return g ? g[3] + g[2] + g[1] : '00000000';
}
function sortHistory(h){ return h.slice().sort((a, b) => periodKey(a) < periodKey(b) ? -1 : periodKey(a) > periodKey(b) ? 1 : 0); }
function loadHistory(){
  try { const s = localStorage.getItem(LS_KEY); if (s){ const h = JSON.parse(s); if (Array.isArray(h) && h.length) return sortHistory(h); } } catch(e){}
  return sortHistory(DEFAULT_HISTORY.map(s => s));
}
function saveHistory(h){ try { localStorage.setItem(LS_KEY, JSON.stringify(h)); } catch(e){} }
let history = loadHistory();
let curIdx = history.length - 1;
let model = history[curIdx];

/* ---------- filtros ---------- */
// null = sem restrição (tudo incluído); Set<int> = apenas os índices presentes
let filt = {unidade: null, categoria: null, turno: null, faixa: null};
function resetFilters(){ filt = {unidade: null, categoria: null, turno: null, faixa: null}; }

function aggregate(abs, f){
  const D = abs.dict;
  const dims = {unidade: {}, categoria: {}, turno: {}, tipo: {}, estagio: {}, tempo: {}};
  const matrix = {}; ORDEM_ESTAGIO.forEach(e => matrix[e] = {});
  let total = 0, wsum = 0;
  for (const row of abs.rows){
    const [uI, cI, tI, tpI, fI, teI, v] = row;
    if (f.unidade && !f.unidade.has(uI)) continue;
    if (f.categoria && !f.categoria.has(cI)) continue;
    if (f.turno && !f.turno.has(tI)) continue;
    if (f.faixa && !f.faixa.has(fI)) continue;
    const u = D.unidade[uI], cat = D.categoria[cI], tur = D.turno[tI], tip = D.tipo[tpI], fx = D.faixa[fI], te = D.tempo[teI];
    const est = estagio(fx);
    total += v; wsum += v * tempoMid(te);
    dims.unidade[u] = (dims.unidade[u] || 0) + v;
    dims.categoria[cat] = (dims.categoria[cat] || 0) + v;
    dims.turno[tur] = (dims.turno[tur] || 0) + v;
    dims.tipo[tip] = (dims.tipo[tip] || 0) + v;
    dims.estagio[est] = (dims.estagio[est] || 0) + v;
    dims.tempo[te] = (dims.tempo[te] || 0) + v;
    matrix[est][cat] = (matrix[est][cat] || 0) + v;
  }
  return {total, dims, tempoMedio: total ? wsum / total : null, matrix};
}

/* ---------- renderização ---------- */
const toItems = (d, sort) => Object.keys(d).map(k => ({label: k, value: d[k]})).sort(sort || ((a, b) => b.value - a.value));
function volBars(id, items, total, amberIf){
  const max = Math.max(1, ...items.map(i => i.value));
  $(id).innerHTML = items.length ? items.map(i => `<div class="bar-row"><div class="lbl" title="${esc(i.label)}">${esc(i.label)}</div><div class="bar-track"><div class="bar-fill ${amberIf && amberIf(i) ? 'amber' : ''}" style="width:${(i.value / max * 100).toFixed(1)}%"></div></div><div class="val"><span class="primary">${fmt(i.value)}</span><span class="secondary">${total ? pct(i.value / total * 100) : '—'}</span></div></div>`).join('')
    : '<div class="empty">Nenhum registro no recorte selecionado.</div>';
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

/* ---------- painel de filtros ---------- */
function buildFilterUI(){
  const D = model.abs.dict;
  const group = (containerId, dim, labels) => {
    $(containerId).innerHTML = labels.map((l, i) => `<label><input type="checkbox" data-dim="${dim}" data-i="${i}" checked>${esc(l)}</label>`).join('');
    $(containerId).querySelectorAll('input').forEach(cb => cb.onchange = () => { syncFilterFromDOM(dim, D[dim].length); render(); });
  };
  group('f-unidade', 'unidade', D.unidade);
  group('f-categoria', 'categoria', D.categoria);
  group('f-turno', 'turno', D.turno);
  group('f-faixa', 'faixa', D.faixa);
  $('f-chips').innerHTML = ORDEM_ESTAGIO.map(e => `<button class="chip" data-est="${esc(e)}">${esc(e)}</button>`).join('') +
    '<button class="chip active" data-est="__all">Todas as idades</button>';
  $('f-chips').querySelectorAll('.chip').forEach(btn => btn.onclick = () => {
    const est = btn.dataset.est;
    const idxs = est === '__all' ? D.faixa.map((_, i) => i) : D.faixa.map((l, i) => estagio(l) === est ? i : -1).filter(i => i >= 0);
    filt.faixa = est === '__all' ? null : new Set(idxs);
    $('f-faixa').querySelectorAll('input').forEach((cb, i) => cb.checked = est === '__all' ? true : idxs.includes(i));
    $('f-chips').querySelectorAll('.chip').forEach(b => b.classList.toggle('active', b === btn));
    render();
  });
  document.querySelectorAll('.fg-actions button').forEach(btn => btn.onclick = () => {
    const dim = btn.dataset.g, all = btn.dataset.a === 'all';
    document.getElementById('f-' + dim).querySelectorAll('input').forEach(cb => cb.checked = all);
    filt[dim] = all ? null : new Set();
    if (dim === 'faixa') $('f-chips').querySelectorAll('.chip').forEach(b => b.classList.toggle('active', all && b.dataset.est === '__all'));
    render();
  });
  $('btn-clear-filters').onclick = () => {
    resetFilters();
    document.querySelectorAll('.fg-list input').forEach(cb => cb.checked = true);
    $('f-chips').querySelectorAll('.chip').forEach(b => b.classList.toggle('active', b.dataset.est === '__all'));
    render();
  };
}
function syncFilterFromDOM(dim, totalCount){
  const checked = Array.from(document.getElementById('f-' + dim).querySelectorAll('input:checked')).map(cb => +cb.dataset.i);
  filt[dim] = checked.length === totalCount ? null : new Set(checked);
  if (dim === 'faixa') $('f-chips').querySelectorAll('.chip').forEach(b => b.classList.remove('active'));
}
function filtersSummary(){
  const active = Object.keys(filt).filter(k => filt[k] !== null);
  return active.length ? active.length + ' filtro(s) ativo(s)' : 'todos os registros';
}

/* ---------- heatmap ---------- */
function renderHeatmap(matrix, D){
  const cats = D.categoria.slice().sort((a, b) => {
    const sa = ORDEM_ESTAGIO.reduce((s, e) => s + (matrix[e][a] || 0), 0), sb = ORDEM_ESTAGIO.reduce((s, e) => s + (matrix[e][b] || 0), 0);
    return sb - sa;
  });
  let max = 1; ORDEM_ESTAGIO.forEach(e => cats.forEach(c => { max = Math.max(max, matrix[e][c] || 0); }));
  const colTot = {}; cats.forEach(c => colTot[c] = ORDEM_ESTAGIO.reduce((s, e) => s + (matrix[e][c] || 0), 0));
  let html = '<table class="heat"><thead><tr><th class="rowh">Faixa etária</th>' + cats.map(c => `<th>${esc(c)}</th>`).join('') + '<th class="tot">Total</th></tr></thead><tbody>';
  ORDEM_ESTAGIO.forEach(e => {
    const rowTot = cats.reduce((s, c) => s + (matrix[e][c] || 0), 0);
    html += `<tr><th class="rowh">${esc(e)}</th>` + cats.map(c => {
      const v = matrix[e][c] || 0;
      if (!v) return '<td class="z">–</td>';
      const op = (0.16 + 0.84 * (v / max)).toFixed(2);
      return `<td style="background:rgba(39,69,161,${op})">${fmt(v)}</td>`;
    }).join('') + `<td class="tot">${fmt(rowTot)}</td></tr>`;
  });
  html += `<tr><th class="tot">Total</th>` + cats.map(c => `<td class="tot">${fmt(colTot[c])}</td>`).join('') + `<td class="tot">${fmt(cats.reduce((s, c) => s + colTot[c], 0))}</td></tr>`;
  html += '</tbody></table>';
  $('heatmap').innerHTML = html;
  const nonEmpty = ORDEM_ESTAGIO.filter(e => cats.some(c => matrix[e][c]));
  if (!nonEmpty.length){ show('n-heat', true, 'Nenhum registro no recorte selecionado.'); return; }
  let bestE = null, bestC = null, bestV = -1;
  ORDEM_ESTAGIO.forEach(e => cats.forEach(c => { if ((matrix[e][c] || 0) > bestV){ bestV = matrix[e][c] || 0; bestE = e; bestC = c; } }));
  const bestU = null;
  show('n-heat', bestV > 0, bestV > 0 ? `No recorte atual, a combinação com mais faltas é <b>${esc(bestC)}</b> entre <b>${esc(bestE)}</b>, com ${fmt(bestV)} falta(s).` : '');
}

/* ---------- seletor de período e histórico ---------- */
function buildPeriodSelector(){
  const sel = $('period-select');
  sel.innerHTML = history.map((s, i) => `<option value="${i}">${esc(s.periodo || 'período não identificado')}${s.origem === 'dados iniciais' ? ' (inicial)' : ''}</option>`).join('');
  sel.value = String(curIdx);
  sel.onchange = () => { curIdx = +sel.value; model = history[curIdx]; resetFilters(); buildFilterUI(); render(); };
  $('period-count').textContent = history.length > 1 ? history.length + ' períodos no histórico' : '1 período carregado';
}
function renderHistoryTab(){
  const rows = history.map(s => {
    const a = s.abs.total, t = s.at ? s.at.total : null;
    const taxa = t != null ? a / (a + t) * 100 : null;
    return {label: s.periodo || 'não identificado', a, t, taxa, origem: s.origem, gerado: s.gerado_em};
  });
  volBars('h-faltas', rows.map(r => ({label: r.label, value: r.a})), null);
  const withRate = rows.filter(r => r.taxa != null);
  if (withRate.length){
    $('h-taxa-wrap').hidden = false;
    const max = Math.max(1, ...withRate.map(r => r.taxa));
    $('h-taxa').innerHTML = withRate.map(r => `<div class="bar-row wide"><div class="lbl" title="${esc(r.label)}">${esc(r.label)}</div><div class="bar-track"><div class="bar-fill ${rc(r.taxa)}" style="width:${(r.taxa / max * 100).toFixed(1)}%"></div></div><div class="val"><span class="primary">${pct(r.taxa)}</span></div></div>`).join('');
  } else { $('h-taxa-wrap').hidden = true; }
  if (rows.length >= 2){
    const cur = rows[rows.length - 1], prev = rows[rows.length - 2];
    const delta = cur.a - prev.a, deltaPct = prev.a ? (delta / prev.a * 100) : null;
    show('h-delta', true, `De <b>${esc(prev.label)}</b> para <b>${esc(cur.label)}</b>: faltas ${delta >= 0 ? 'subiram' : 'caíram'} de ${fmt(prev.a)} para ${fmt(cur.a)}` +
      (deltaPct != null ? ` (${delta >= 0 ? '+' : ''}${pct(deltaPct).replace('-', '')})`.replace('(-', '(−') : '') +
      (cur.taxa != null && prev.taxa != null ? `. Taxa global: de ${pct(prev.taxa)} para ${pct(cur.taxa)}.` : '.'));
  } else show('h-delta', false);
  $('h-table').innerHTML = '<table class="heat"><thead><tr><th class="rowh">Período</th><th>Atualizado em</th><th>Origem</th><th>Faltas</th><th>Atendimentos</th><th>Taxa</th></tr></thead><tbody>' +
    rows.map(r => `<tr><th class="rowh">${esc(r.label)}</th><td class="z">${esc(r.gerado)}</td><td class="z">${esc(r.origem)}</td><td class="tot">${fmt(r.a)}</td><td class="tot">${r.t != null ? fmt(r.t) : '—'}</td><td class="tot">${r.taxa != null ? pct(r.taxa) : '—'}</td></tr>`).join('') +
    '</tbody></table>';
}
async function downloadDataJs(){
  const js = 'const DEFAULT_HISTORY = ' + JSON.stringify(sortHistory(history)) + ';\n';
  try {
    if (window.claude && typeof window.claude.use === 'function'){
      const downloads = await window.claude.use('downloads');
      if (downloads){ await downloads.save({filename: 'data.js', data: new Blob([js], {type: 'text/javascript'})}); return; }
    }
  } catch(e){ /* sem suporte aqui — cai no download padrão do navegador abaixo */ }
  const blob = new Blob([js], {type: 'text/javascript'});
  const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = 'data.js'; a.click();
  setTimeout(() => URL.revokeObjectURL(a.href), 4000);
}

function render(){
  const m = model, A = m.abs, D = A.dict;
  $('periodo').textContent = m.periodo || 'não identificado';
  $('gerado').textContent = m.gerado_em;
  $('origem').textContent = m.origem;
  $('filters-summary').textContent = filtersSummary();

  const agg = aggregate(A, filt);
  const dd = agg.dims;
  const items = k => toItems(dd[k] || {});
  const topU = items('unidade')[0];
  const turnoItems = items('turno');
  const manha = turnoItems.find(t => norm(t.label) === 'manha');

  $('kpis-volume').innerHTML = kpi(fmt(agg.total), 'Faltas no recorte selecionado') +
    kpi(agg.tempoMedio != null ? agg.tempoMedio.toFixed(1).replace('.', ',') : '—', 'Dias, em média, entre agendamento e a falta') +
    kpi(manha && agg.total ? pct(manha.value / agg.total * 100) : '—', 'Das faltas ocorrem no turno da manhã') +
    kpi(topU ? fmt(topU.value) : '—', topU ? 'Faltas na unidade de maior volume — ' + esc(topU.label) : 'Sem unidade no recorte');

  const tempo = toItems(dd.tempo || {}, (a, b) => firstNum(a.label) - firstNum(b.label));
  volBars('c-unidade', items('unidade'), agg.total, i => topU && i.label === topU.label);
  volBars('c-categoria', items('categoria'), agg.total);
  volBars('c-turno', turnoItems, agg.total);
  volBars('c-tipo', items('tipo'), agg.total);
  volBars('c-estagio', ORDEM_ESTAGIO.filter(k => dd.estagio[k]).map(k => ({label: k, value: dd.estagio[k]})), agg.total);
  volBars('c-tempo', tempo, agg.total, i => tempo[0] && i.label === tempo[0].label);
  const t0 = tempo[0];
  show('n-tempo', !!t0, t0 ? `Faltas no intervalo de ${esc(t0.label)} correspondem a ${pct(t0.value / agg.total * 100)} do recorte selecionado.` : '');

  renderHeatmap(agg.matrix, D);
  renderHistoryTab();

  /* aba taxa — sempre com a base completa, sem os filtros acima */
  const T = m.at;
  show('taxa-empty', !T); show('taxa-body', !!T);
  if (!T) return;
  const fullAgg = aggregate(A, {unidade: null, categoria: null, turno: null, faixa: null});
  const glob = fullAgg.total / (fullAgg.total + T.total) * 100;
  $('kpis-taxa').innerHTML = kpi(pct(glob), 'Taxa global de absenteísmo na rede') + kpi(fmt(fullAgg.total + T.total), 'Total de vagas agendadas no período') +
    kpi(fmt(T.total), 'Atendimentos realizados') + kpi(fmt(fullAgg.total), 'Faltas registradas');
  const ru = rates(fullAgg.dims.unidade, T.dims.unidade), rcat = rates(fullAgg.dims.categoria, T.dims.categoria);
  rateBars('t-unidade', ru.rows); rateBars('t-categoria', rcat.rows);
  const topUFull = toItems(fullAgg.dims.unidade)[0];
  if (ru.rows.length > 2 && topUFull){
    const pos = ru.rows.findIndex(r => r.label === topUFull.label) + 1;
    show('n-unidade', true, `A unidade de maior volume absoluto de faltas (${esc(topUFull.label)}) ocupa a posição ${pos} de ${ru.rows.length} em taxa percentual` +
      (pos ? ` (${pct(ru.rows[pos - 1].taxa)})` : '') + `. As maiores taxas relativas são de ${esc(ru.rows[0].label)} (${pct(ru.rows[0].taxa)}) e ${esc(ru.rows[1].label)} (${pct(ru.rows[1].taxa)}). Para priorização de intervenção, a taxa percentual é o indicador tecnicamente mais adequado.`);
  } else show('n-unidade', false);
  const warn = (id, ex, tipo) => show(id, ex.length, ex.length ? `${tipo} sem atendimentos no relatório de atendimentos: ${ex.map(e => esc(e.label) + ' (' + fmt(e.absenteismo) + ' faltas)').join('; ')}. Excluídos da taxa por provável diferença de escopo entre os relatórios; verificar na fonte primária do e-SUS antes de qualquer conclusão.` : '');
  warn('w-unidade', ru.excl, 'Unidades'); warn('w-categoria', rcat.excl, 'Categorias');
  ['turno', 'hora'].forEach(k => {
    const ok = T.dims[k]; show('p-t-' + k, !!ok);
    if (ok){ const r = rates(fullAgg.dims[k] || {}, T.dims[k]).rows; if (k === 'hora') r.sort((a, b) => firstNum(a.label) - firstNum(b.label)); rateBars('t-' + k, r); }
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
    const snap = {periodo: p, gerado_em: new Date().toLocaleDateString('pt-BR'), origem: 'arquivos enviados',
             abs: {total: abs.total, tempoMedio: abs.tempoMedio, dict: abs.dict, rows: abs.rows}, at: at ? {total: at.total, dims: at.dims} : null};
    const existing = history.findIndex(s => periodKey(s) === periodKey(snap) && (s.periodo || null) === (snap.periodo || null));
    let countMsg;
    if (existing >= 0){ history[existing] = snap; countMsg = ' Período já existente no histórico foi atualizado.'; }
    else { history.push(snap); countMsg = ' Adicionado como novo período no histórico (' + history.length + ' no total).'; }
    history = sortHistory(history);
    curIdx = history.findIndex(s => s === snap);
    saveHistory(history); model = history[curIdx];
    resetFilters(); buildFilterUI(); buildPeriodSelector(); render(); status(msg + countMsg, cls);
  } catch(e){ status(e.message || String(e), 'err'); }
}
$('btn-upload').onclick = () => $('files').click();
$('files').onchange = e => { if (e.target.files.length) handleFiles(Array.from(e.target.files)); e.target.value = ''; };
$('btn-download-data').onclick = downloadDataJs;
$('btn-reset').onclick = () => {
  if (!confirm('Isso apaga o histórico salvo neste navegador e volta ao(s) período(s) que veio(vieram) com o site. Continuar?')) return;
  try { localStorage.removeItem(LS_KEY); } catch(e){}
  history = sortHistory(DEFAULT_HISTORY.map(s => s)); curIdx = history.length - 1; model = history[curIdx];
  resetFilters(); buildFilterUI(); buildPeriodSelector(); render(); status('Dados iniciais restaurados.', 'ok');
};
document.querySelectorAll('.tab-btn').forEach(b => b.onclick = () => {
  document.querySelectorAll('.tab-btn').forEach(x => x.classList.remove('active'));
  document.querySelectorAll('.tab-panel').forEach(x => x.classList.remove('active'));
  b.classList.add('active'); $('tab-' + b.dataset.tab).classList.add('active');
});
buildFilterUI(); buildPeriodSelector(); render();
})();
