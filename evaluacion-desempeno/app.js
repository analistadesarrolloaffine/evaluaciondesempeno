/* =====================================================================
   Interfaz - Evaluación de Desempeño
   ===================================================================== */
(function () {
  'use strict';
  const E = window.Engine;
  const { ROLES, ROLE_LABEL, ROLE_SHORT, SCALE_MIN, SCALE_MAX } = E;
  const STORE_KEY = 'evaldesempeno:ponderaciones:v1';
  const LOW_BASE = 5; // competencias con menos personas se marcan como base baja

  const S = {
    raw: { scores: null, feedback: null }, names: { scores: '', feedback: '' },
    model: null, analysis: null, weights: {},
    filters: { area: '', cargo: '', cat: '' }, chartMode: 'final', sort: { key: 'name', dir: 1 }, search: '',
    current: null, charts: {}, user: null, order: [], cancel: false, dashPeople: null,
  };

  /* ---------- utilidades ---------- */
  const $ = (s, el = document) => el.querySelector(s);
  const $$ = (s, el = document) => [...el.querySelectorAll(s)];
  const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const f2 = n => (n == null || !Number.isFinite(n)) ? '—' : n.toFixed(2).replace('.', ',');
  const sgn = n => (n == null) ? '—' : (n > 0 ? '+' : n < 0 ? '−' : '') + Math.abs(n).toFixed(2).replace('.', ',');
  const pctScale = n => n == null ? null : Math.round(((n - SCALE_MIN) / (SCALE_MAX - SCALE_MIN)) * 100);
  const pos = n => Math.max(0, Math.min(100, ((n - SCALE_MIN) / (SCALE_MAX - SCALE_MIN)) * 100));
  const fold = s => String(s).normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase();
  const cssv = name => getComputedStyle(document.documentElement).getPropertyValue(name).trim();
  const comboLabel = key => key === 'NINGUNO' ? 'Sin calificaciones válidas' : key.split('+').map(r => ROLE_LABEL[r]).join(' + ');
  const icons = () => { if (window.lucide) window.lucide.createIcons(); };
  const frames = (n = 2) => new Promise(res => { const step = k => k ? requestAnimationFrame(() => step(k - 1)) : res(); step(n); });
  const titleCase = s => String(s).toLowerCase().replace(/(^|\s|-)(\p{L})/gu, (m, a, b) => a + b.toUpperCase());

  const appTitle = () => (S.user && S.user.titulo) || (S.model && S.model.campaigns.join(', ')) || 'Evaluación de desempeño';
  const slug = t => fold(t).replace(/[^a-z0-9]+/g, '_').replace(/^_|_$/g, '');
  const stamp = () => new Date().toISOString().slice(0, 10);

  function toast(msg) {
    const t = $('#toast'); t.textContent = msg; t.hidden = false;
    clearTimeout(toast._t); toast._t = setTimeout(() => { t.hidden = true; }, 3200);
  }

  /* ---------- inicio de sesión ---------- */
  const SESSION_KEY = 'evaldesempeno:sesion';
  const appUsers = () => (Array.isArray(window.APP_USERS) ? window.APP_USERS : []);
  const findUser = name => appUsers().find(u => String(u.usuario || '').trim().toLowerCase() === String(name || '').trim().toLowerCase());
  function startSession(u) {
    S.user = u;
    document.body.classList.remove('locked');
    $('#sessionName').textContent = u.nombre || u.usuario;
    $('#campaignName').textContent = u.titulo || 'Carga los archivos para comenzar';
    document.title = (u.titulo || 'Evaluación de desempeño') + ' | Resultados';
    try { sessionStorage.setItem(SESSION_KEY, u.usuario); } catch (e) { /* sin almacenamiento */ }
    loadGeneralFb();
    const gh = String(u.rol || '').toLowerCase() === 'gh';
    document.body.classList.toggle('role-gh', gh);
    $('.step[data-view="load"]').hidden = gh;
    if (gh) {
      // Gestión Humana no carga archivos: la retroalimentación general pasa al dashboard
      const fb = $('#generalFb');
      $('#view-dashboard').insertBefore(fb, $('#kpis'));
      fb.classList.add('in-dash');
      $('#generalFb .kpi-s').textContent = 'Lo que escribas aquí aparece en el informe de la compañía (PDF). Se guarda en este navegador.';
    }
    icons();
    autoLoad(u);
  }
  $('#loginForm').addEventListener('submit', e => {
    e.preventDefault();
    const err = $('#loginError');
    if (!appUsers().length) { err.textContent = 'No hay usuarios configurados. Revisa el archivo users.js.'; err.hidden = false; return; }
    const u = findUser($('#loginUser').value);
    if (!u || String(u.clave) !== $('#loginPass').value) {
      err.textContent = 'Usuario o contraseña incorrectos.'; err.hidden = false; $('#loginPass').select(); return;
    }
    err.hidden = true; $('#loginPass').value = '';
    startSession(u);
  });
  $('#pwToggle').addEventListener('click', () => {
    const i = $('#loginPass'); const show = i.type === 'password';
    i.type = show ? 'text' : 'password';
    $('#pwToggle').setAttribute('aria-label', show ? 'Ocultar contraseña' : 'Mostrar contraseña');
  });
  $('#btnLogout').addEventListener('click', () => {
    try { sessionStorage.removeItem(SESSION_KEY); } catch (e) { /* sin almacenamiento */ }
    location.reload();
  });

  /* ---------- navegación ---------- */
  function showView(name) {
    $$('.view').forEach(v => v.classList.toggle('is-active', v.id === 'view-' + name));
    $$('.step').forEach(b => b.classList.toggle('is-active', b.dataset.view === name));
    window.scrollTo({ top: 0 });
    if (name === 'dashboard') renderCompChart();
    if (name === 'report' && S.current) renderRadar(S.current);
  }
  $$('.step').forEach(b => b.addEventListener('click', () => !b.disabled && showView(b.dataset.view)));

  /* =================================================================
     1. CARGA DE ARCHIVOS
     ================================================================= */
  function readWorkbook(file) {
    return new Promise((resolve, reject) => {
      if (!window.XLSX) return reject(new Error('No se pudo cargar la librería de Excel. Revisa la conexión a internet.'));
      const fr = new FileReader();
      fr.onload = () => {
        try {
          const wb = XLSX.read(new Uint8Array(fr.result), { type: 'array' });
          const ws = wb.Sheets[wb.SheetNames[0]];
          resolve(XLSX.utils.sheet_to_json(ws, { defval: null, raw: true }));
        } catch (e) { reject(new Error('El archivo no se pudo leer como Excel.')); }
      };
      fr.onerror = () => reject(new Error('No se pudo leer el archivo.'));
      fr.readAsArrayBuffer(file);
    });
  }

  function setupDrop(label) {
    const slot = label.dataset.slot;
    const input = $('input', label);
    const status = $('[data-status]', label);
    const handle = async file => {
      if (!file) return;
      label.classList.remove('is-loaded', 'is-error');
      status.textContent = 'Leyendo…';
      try {
        const rows = await readWorkbook(file);
        const req = slot === 'scores' ? E.REQ_SCORES : E.REQ_FEEDBACK;
        const miss = E.missingColumns(rows, req);
        if (!rows.length) throw new Error('El archivo no tiene filas.');
        if (miss.length) {
          const other = slot === 'scores' ? E.REQ_FEEDBACK : E.REQ_SCORES;
          const looksOther = E.missingColumns(rows, other).length === 0;
          throw new Error(looksOther
            ? 'Este archivo corresponde al otro recuadro.'
            : 'Faltan columnas: ' + miss.join(', '));
        }
        S.raw[slot] = rows; S.names[slot] = file.name;
        label.classList.add('is-loaded');
        status.innerHTML = '<i data-lucide="check"></i>' + esc(file.name) + ', ' + rows.length.toLocaleString('es-CO') + ' filas';
      } catch (e) {
        S.raw[slot] = null;
        label.classList.add('is-error');
        status.innerHTML = '<i data-lucide="alert-triangle"></i>' + esc(e.message);
      }
      icons();
      $('#btnProcess').disabled = !S.raw.scores;
    };
    input.addEventListener('change', () => handle(input.files[0]));
    ['dragenter', 'dragover'].forEach(ev => label.addEventListener(ev, e => { e.preventDefault(); label.classList.add('is-over'); }));
    ['dragleave', 'drop'].forEach(ev => label.addEventListener(ev, e => { e.preventDefault(); label.classList.remove('is-over'); }));
    label.addEventListener('drop', e => handle(e.dataTransfer.files[0]));
  }
  $$('.drop').forEach(setupDrop);

  // Construye el modelo con los datos cargados y prepara todas las vistas
  function processData(goTo) {
    const camp = S.user && S.user.campana ? String(S.user.campana).trim() : '';
    S.model = E.buildModel(S.raw.scores, S.raw.feedback, { campaign: camp });
    if (!S.model.people.length) throw new Error(camp
      ? `El archivo no tiene filas de la campaña “${camp}” asignada a tu usuario.`
      : 'No se encontraron evaluados en el archivo de calificaciones.');
    S.analysis = E.analyzeCombos(S.model);
    S.weights = loadWeights();
    $('#campaignName').textContent = appTitle();
    renderWeights();
    populateFilters();
    E.applyWeights(S.model, S.weights);
    renderDashboard();
    $$('.step').forEach(b => { b.disabled = false; });
    resetReport();
    showView(goTo);
  }

  $('#btnProcess').addEventListener('click', () => {
    const msg = $('#loadMsg'); msg.hidden = true;
    try {
      processData('weights');
      if (!S.raw.feedback) toast('Sin archivo de retroalimentación: los reportes no tendrán comentarios.');
    } catch (e) {
      msg.textContent = e.message; msg.hidden = false;
    }
  });

  /* ---------- Datos precargados (usuarios con "archivos" en users.js) ---------- */
  async function fetchRows(url, required) {
    const r = await fetch(url + (url.includes('?') ? '&' : '?') + 'v=' + Date.now(), { cache: 'no-store' });
    if (!r.ok) throw new Error(`No se encontró el archivo “${url}” en el sitio (error ${r.status}).`);
    const wb = XLSX.read(new Uint8Array(await r.arrayBuffer()), { type: 'array' });
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { defval: null, raw: true });
    const miss = E.missingColumns(rows, required);
    if (miss.length) throw new Error(`Al archivo “${url}” le faltan columnas: ${miss.join(', ')}.`);
    return rows;
  }

  async function autoLoad(u) {
    const files = u.archivos || {};
    if (!files.calificaciones) return;
    busy('Cargando los datos de la evaluación…');
    try {
      if (location.protocol === 'file:')
        throw new Error('Los datos precargados solo funcionan con la app publicada (Netlify, Vercel) o en un servidor local; al abrir el archivo directamente el navegador los bloquea.');
      if (!window.XLSX) throw new Error('No se pudo cargar la librería de Excel. Revisa la conexión a internet.');
      S.raw.scores = await fetchRows(files.calificaciones, E.REQ_SCORES);
      S.raw.feedback = null;
      if (files.retroalimentacion) {
        try { S.raw.feedback = await fetchRows(files.retroalimentacion, E.REQ_FEEDBACK); }
        catch (e) { toast(e.message + ' Los reportes saldrán sin comentarios.'); }
      }
      processData('dashboard');
    } catch (e) {
      // Si falla, se muestra la pantalla de carga con el motivo para poder subir los archivos a mano
      document.body.classList.remove('role-gh');
      $('.step[data-view="load"]').hidden = false;
      $('#loadMsg').textContent = e.message; $('#loadMsg').hidden = false;
      showView('load');
    } finally {
      busyEnd();
    }
  }

  /* =================================================================
     2. PONDERACIONES
     ================================================================= */
  function loadWeights() {
    let saved = {};
    try { saved = JSON.parse(localStorage.getItem(STORE_KEY) || '{}') || {}; } catch (e) { saved = {}; }
    const w = {};
    for (const c of S.analysis.combos) {
      const sv = saved[c.key];
      const valid = sv && c.roles.every(r => Number.isFinite(sv[r]));
      w[c.key] = valid ? { ...sv } : { ...c.suggestion.weights };
    }
    return w;
  }
  function saveWeights() {
    try { localStorage.setItem(STORE_KEY, JSON.stringify(S.weights)); } catch (e) { /* almacenamiento no disponible */ }
  }
  const sumW = key => Object.values(S.weights[key] || {}).reduce((a, b) => a + (Number(b) || 0), 0);

  function renderWeights() {
    const { combos, stats } = S.analysis;
    const n = S.model.people.length;
    $('#weightsLead').textContent = `Revisé a las ${n} personas evaluadas y las agrupé según las fuentes que sí calificaron. ` +
      'Define cuánto pesa cada fuente en el puntaje final de cada grupo; la suma debe dar 100 %. Ya hay una sugerencia cargada.';
    $('#analysisFacts').innerHTML = [
      [n, 'personas evaluadas'],
      [stats.nBoth, 'tienen autoevaluación y jefe'],
      [stats.nColab, stats.nColab ? `son evaluadas por su equipo (promedio ${f2(stats.avgColab)} colaboradores)` : 'son evaluadas por su equipo'],
      [sgn(stats.autoBias), 'diferencia promedio autoevaluación − jefe'],
    ].map(([v, l]) => `<div class="fact"><div class="fact-v">${esc(v)}</div><div class="fact-l">${esc(l)}</div></div>`).join('');

    $('#comboList').innerHTML = combos.map(c => {
      const single = c.roles.length <= 1;
      const share = Math.round(c.people.length / n * 100);
      const inputs = c.roles.map(r => `
        <div class="w-row">
          <label for="w-${c.key}-${r}"><span class="dot dot-${r}"></span>${ROLE_LABEL[r]}</label>
          <div class="w-input"><input id="w-${c.key}-${r}" type="number" min="0" max="100" step="5"
            data-combo="${c.key}" data-role="${r}" value="${S.weights[c.key][r] ?? 0}" ${single ? 'disabled' : ''}></div>
        </div>`).join('');
      return `
      <div class="combo" data-combo="${c.key}">
        <div>
          <div class="combo-name">${esc(comboLabel(c.key))}</div>
          <div class="combo-count">${c.people.length} ${c.people.length === 1 ? 'persona' : 'personas'} (${share} %)</div>
          ${single && c.roles.length ? `<div class="combo-partial"><i data-lucide="alert-triangle"></i>Medición parcial: el puntaje sale de una sola fuente.</div>` : ''}
          <details><summary>Ver personas</summary><ul>${c.people.map(p => `<li>${esc(titleCase(p.name))}, ${esc(p.cargo.toLowerCase())}</li>`).join('')}</ul></details>
        </div>
        <div>
          ${inputs || '<p class="combo-why">No hay calificaciones válidas para ponderar.</p>'}
          <div class="w-bar" data-bar></div>
          <div class="w-sum" data-sum></div>
        </div>
        <div class="combo-why">
          <h3>Por qué esta sugerencia ${single ? '' : `<button class="link-btn" data-use="${c.key}">Usar sugerencia</button>`}</h3>
          <ul>${c.suggestion.reasons.map(t => `<li>${esc(t)}</li>`).join('')}</ul>
        </div>
      </div>`;
    }).join('');
    combos.forEach(c => updateComboBar(c.key));
    icons();
    validateWeights();
  }

  function updateComboBar(key) {
    const box = $(`.combo[data-combo="${CSS.escape(key)}"]`);
    if (!box) return;
    const w = S.weights[key] || {};
    const total = sumW(key);
    $('[data-bar]', box).innerHTML = Object.keys(w).map(r =>
      `<span style="width:${total ? (w[r] / Math.max(total, 100)) * 100 : 0}%;background:var(--${r === 'AUTOEVALUACION' ? 'auto' : r === 'JEFE' ? 'jefe' : 'colab'})"></span>`).join('');
    const s = $('[data-sum]', box);
    s.textContent = `Suma: ${total} %` + (total !== 100 ? ' (debe ser 100 %)' : '');
    s.classList.toggle('bad', total !== 100);
  }

  function validateWeights() {
    const bad = S.analysis.combos.filter(c => c.roles.length && sumW(c.key) !== 100);
    const el = $('#weightsError');
    el.hidden = !bad.length;
    el.textContent = bad.length ? `Revisa: ${bad.map(c => comboLabel(c.key)).join('; ')} no suma 100 %.` : '';
    $('#btnApplyWeights').disabled = !!bad.length;
    return !bad.length;
  }

  $('#comboList').addEventListener('input', e => {
    const t = e.target;
    if (!t.dataset.combo) return;
    S.weights[t.dataset.combo][t.dataset.role] = Math.max(0, Math.min(100, Number(t.value) || 0));
    updateComboBar(t.dataset.combo); validateWeights(); saveWeights();
  });
  $('#comboList').addEventListener('click', e => {
    const k = e.target.dataset && e.target.dataset.use;
    if (!k) return;
    const c = S.analysis.combos.find(x => x.key === k);
    S.weights[k] = { ...c.suggestion.weights };
    for (const r of c.roles) $(`#w-${CSS.escape(k)}-${r}`).value = S.weights[k][r];
    updateComboBar(k); validateWeights(); saveWeights();
  });
  $('#btnSuggest').addEventListener('click', () => {
    for (const c of S.analysis.combos) S.weights[c.key] = { ...c.suggestion.weights };
    saveWeights(); renderWeights(); toast('Se aplicó la sugerencia del sistema en todas las combinaciones.');
  });
  $('#btnApplyWeights').addEventListener('click', () => {
    if (!validateWeights()) return;
    E.applyWeights(S.model, S.weights);
    renderDashboard();
    if (S.current) renderReport(S.current);
    showView('dashboard');
  });

  /* =================================================================
     SEMÁFORO
     ================================================================= */
  const SEM_KEY = 'evaldesempeno:semaforo:v1';
  const SEM_LABEL = { green: 'Sobresaliente', yellow: 'Adecuado', red: 'Por mejorar', na: 'Sin datos' };
  S.sem = { green: 3.5, yellow: 3.0 };
  try { const v = JSON.parse(localStorage.getItem(SEM_KEY) || 'null'); if (v && v.green > v.yellow) S.sem = v; } catch (e) { /* sin almacenamiento */ }
  const semKey = v => v == null ? 'na' : v >= S.sem.green ? 'green' : v >= S.sem.yellow ? 'yellow' : 'red';
  const semDot = v => `<span class="sem-dot sem-${semKey(v)}" title="${SEM_LABEL[semKey(v)]}"></span>`;
  const semTag = v => `<span class="sem">${semDot(v)}${SEM_LABEL[semKey(v)]}</span>`;
  const semColor = k => cssv('--sem-' + k) || cssv('--line');
  const semLegend = () => `<span class="sem-legend">${['green', 'yellow', 'red'].map(k =>
    `<span><span class="sem-dot sem-${k}"></span>${SEM_LABEL[k]} ${k === 'green' ? '≥ ' + f2(S.sem.green) : k === 'yellow' ? f2(S.sem.yellow) + ' a ' + f2(S.sem.green - 0.01) : '< ' + f2(S.sem.yellow)}</span>`).join('')}</span>`;

  function initSemaforoInputs() {
    $('#semGreen').value = S.sem.green; $('#semYellow').value = S.sem.yellow;
    $('#semRedTxt').textContent = f2(S.sem.yellow);
  }
  function onSemInput() {
    const g = Number($('#semGreen').value), y = Number($('#semYellow').value);
    const ok = g > y && y > SCALE_MIN && g <= SCALE_MAX;
    $('#semError').hidden = ok;
    $('#semError').textContent = ok ? '' : 'El valor de “Sobresaliente” debe ser mayor que el de “Adecuado”, entre 1 y 4.';
    if (!ok) return;
    S.sem = { green: g, yellow: y };
    $('#semRedTxt').textContent = f2(y);
    try { localStorage.setItem(SEM_KEY, JSON.stringify(S.sem)); } catch (e) { /* sin almacenamiento */ }
    if (S.dashPeople) renderDashboard();
    if (S.current) renderReport(S.current);
  }
  $('#semGreen').addEventListener('input', onSemInput);
  $('#semYellow').addEventListener('input', onSemInput);
  initSemaforoInputs();

  /* =================================================================
     RETROALIMENTACIÓN GENERAL (informe de la compañía)
     ================================================================= */
  const gfbKey = () => 'evaldesempeno:retrogeneral:' + slug((S.user && S.user.titulo) || 'general');
  const GFB = { strengths: '#gfbStrengths', improve: '#gfbImprove', next: '#gfbNext' };
  function loadGeneralFb() {
    let v = {};
    try { v = JSON.parse(localStorage.getItem(gfbKey()) || '{}') || {}; } catch (e) { v = {}; }
    for (const k in GFB) $(GFB[k]).value = v[k] || '';
    updateGfbMeta();
  }
  function generalFb() { const o = {}; for (const k in GFB) o[k] = $(GFB[k]).value.trim(); return o; }
  function updateGfbMeta() {
    const n = Object.values(generalFb()).filter(Boolean).length;
    $('#generalFbMeta').textContent = n ? `${n} de 3 secciones con texto` : 'Opcional';
  }
  Object.values(GFB).forEach(sel => $(sel).addEventListener('input', () => {
    try { localStorage.setItem(gfbKey(), JSON.stringify(generalFb())); } catch (e) { /* sin almacenamiento */ }
    updateGfbMeta();
  }));

  /* =================================================================
     3. DASHBOARD
     ================================================================= */
  S.filters = { area: '', cargo: '', cat: '' };
  S.chartBy = 'comp';
  const catShort = c => String(c).replace(/^competencias\s+/i, '');
  const view = p => E.scoreFor(p, S.filters.cat);

  function populateFilters() {
    S.filters = { area: '', cargo: '', cat: '' };
    $('#fArea').innerHTML = '<option value="">Todas las áreas</option>' + S.model.areas.map(a =>
      `<option value="${esc(a)}">${esc(a)} (${S.model.people.filter(p => p.area === a).length})</option>`).join('');
    $('#fCat').innerHTML = '<option value="">Todas</option>' + S.model.categories.map(c =>
      `<option value="${esc(c)}">${esc(catShort(c))}</option>`).join('');
    fillCargos();
    S.chartCat = '';
    renderChartCatSeg();
  }
  function fillCargos() {
    const pool = S.model.people.filter(p => !S.filters.area || p.area === S.filters.area);
    const cargos = [...new Set(pool.map(p => p.cargo))].sort((a, b) => a.localeCompare(b, 'es'));
    if (!cargos.includes(S.filters.cargo)) S.filters.cargo = '';
    $('#fCargo').innerHTML = '<option value="">Todos los cargos</option>' + cargos.map(c =>
      `<option value="${esc(c)}" ${c === S.filters.cargo ? 'selected' : ''}>${esc(titleCase(c))} (${pool.filter(p => p.cargo === c).length})</option>`).join('');
  }
  $('#fArea').addEventListener('change', e => { S.filters.area = e.target.value; fillCargos(); renderDashboard(); });
  $('#fCargo').addEventListener('change', e => { S.filters.cargo = e.target.value; renderDashboard(); });
  $('#fCat').addEventListener('change', e => { S.filters.cat = e.target.value; renderChartCatSeg(); renderDashboard(); });
  $('#chartBy').addEventListener('change', e => { S.chartBy = e.target.value; renderCompChart(); });
  $$('.seg-btn[data-chart]').forEach(b => b.addEventListener('click', () => {
    $$('.seg-btn[data-chart]').forEach(x => x.classList.toggle('is-on', x === b));
    S.chartMode = b.dataset.chart; renderCompChart();
  }));
  // Tipo de competencia dentro del gráfico: Ambas, Técnicas u Organizacionales
  S.chartCat = '';
  const chartCat = () => S.filters.cat || S.chartCat;
  function renderChartCatSeg() {
    const seg = $('#chartCatSeg');
    const opts = [['', 'Ambas'], ...S.model.categories.map(c => [c, catShort(c)])];
    seg.hidden = !!S.filters.cat;
    seg.innerHTML = opts.map(([v, l]) => `<button class="seg-btn ${v === S.chartCat ? 'is-on' : ''}" data-cat="${esc(v)}">${esc(l)}</button>`).join('');
  }
  $('#chartCatSeg').addEventListener('click', e => {
    const b = e.target.closest('[data-cat]');
    if (!b) return;
    S.chartCat = b.dataset.cat; renderChartCatSeg(); renderCompChart();
  });

  const subset = () => S.model.people.filter(p =>
    (!S.filters.area || p.area === S.filters.area) && (!S.filters.cargo || p.cargo === S.filters.cargo));

  function filtersText() {
    const parts = [];
    if (S.filters.area) parts.push('Área ' + S.filters.area);
    if (S.filters.cargo) parts.push('Cargo ' + titleCase(S.filters.cargo));
    if (S.filters.cat) parts.push(catShort(S.filters.cat));
    return parts.length ? parts.join(', ') : 'Toda la organización';
  }

  function personIssues(p) {
    const out = [];
    for (const e of p.evals.values()) if (e.state !== 'completa') out.push({ type: 'eval', e });
    if (p.dupNotes.length) out.push({ type: 'dup', n: p.dupNotes.length });
    if (p.invalid.length) out.push({ type: 'invalid' });
    for (const r of ['AUTOEVALUACION', 'JEFE']) if (!p.rolesPresent.includes(r)) out.push({ type: 'missing', r });
    return out;
  }

  const filteredComps = people => E.aggregateComps(people, S.model.compCatalog).filter(c => !S.filters.cat || c.category === S.filters.cat);

  function renderDashboard() {
    const people = subset();
    S.dashPeople = people;
    $('#kpis').innerHTML = kpisHTML(people, filteredComps(people));
    renderCompChart();
    renderAreas();
    renderParticipation();
    renderTable();
    $('#btnZipLabel').textContent = `Todos los reportes individuales (${people.length}) .zip`;
    $('#exportHint').textContent = `${filtersText()}. Las descargas respetan los filtros.`;
    icons();
  }

  function semDistHTML(values) {
    const c = { green: 0, yellow: 0, red: 0 };
    values.forEach(v => { const k = semKey(v); if (k in c) c[k]++; });
    const t = c.green + c.yellow + c.red;
    if (!t) return '';
    return `<div class="sem-dist">${['green', 'yellow', 'red'].map(k => `<span class="sem-${k}" style="width:${c[k] / t * 100}%"></span>`).join('')}</div>
      <div class="kpi-s">${c.green} sobresaliente, ${c.yellow} adecuado, ${c.red} por mejorar</div>`;
  }

  function kpisHTML(people, comps) {
    const finals = people.map(p => view(p).final).filter(v => v != null);
    const avg = E.mean(finals);
    const both = people.filter(p => view(p).gap != null).length;
    const reliable = comps.filter(c => c.n >= LOW_BASE);
    const pool = reliable.length ? reliable : comps;
    const best = pool.slice().sort((a, b) => b.final - a.final)[0];
    const gapPool = comps.filter(c => c.gap != null && c.nGap >= LOW_BASE);
    const gapBest = (gapPool.length ? gapPool : comps.filter(c => c.gap != null)).sort((a, b) => Math.abs(b.gap) - Math.abs(a.gap))[0];
    return `
      <div class="kpi"><div class="kpi-l"><i data-lucide="gauge"></i>Promedio general ponderado</div>
        <div class="kpi-v">${f2(avg)}<small> / 4</small></div>
        <div class="kpi-s">${semTag(avg)}</div></div>
      <div class="kpi"><div class="kpi-l"><i data-lucide="users"></i>Colaboradores evaluados</div>
        <div class="kpi-v">${people.length}</div>
        ${semDistHTML(finals) || `<div class="kpi-s">${both} con autoevaluación y jefe</div>`}</div>
      <div class="kpi"><div class="kpi-l"><i data-lucide="trophy"></i>Competencia más alta</div>
        <div class="kpi-t">${best ? esc(best.name) : '—'}</div>
        <div class="kpi-s">${best ? `${f2(best.final)} de promedio, ${best.n} personas` : ''}</div></div>
      <div class="kpi"><div class="kpi-l"><i data-lucide="split"></i>Mayor brecha auto − jefe</div>
        <div class="kpi-t">${gapBest ? esc(gapBest.name) : '—'}</div>
        <div class="kpi-s">${gapBest ? `${sgn(gapBest.gap)}: ${gapBest.gap > 0 ? 'la autoevaluación supera al jefe' : 'el jefe califica más alto'} (${gapBest.nGap} personas)` : 'Sin pares autoevaluación–jefe'}</div></div>`;
  }

  function catColor(cat, alpha) {
    const organizational = /organiz/i.test(cat);
    const c = organizational ? cssv('--navy') : cssv('--cat-tec');
    return alpha != null ? hexA(c, alpha) : c;
  }
  function hexA(hex, a) {
    const h = hex.replace('#', '');
    const n = parseInt(h.length === 3 ? h.split('').map(x => x + x).join('') : h, 16);
    return `rgba(${(n >> 16) & 255},${(n >> 8) & 255},${n & 255},${a})`;
  }

  // Grupos para el gráfico según "Ver por"
  function chartGroups(people, by) {
    const cat = chartCat();
    const groupBy = (keyFn, labelFn) => {
      const m = new Map();
      for (const p of people) { const k = keyFn(p); if (!m.has(k)) m.set(k, []); m.get(k).push(p); }
      return [...m].map(([k, list]) => {
        const v = list.map(p => E.scoreFor(p, cat));
        const byRole = {}, nRole = {};
        for (const r of ROLES) { const xs = v.map(x => x.roles[r]).filter(x => x != null); byRole[r] = E.mean(xs); nRole[r] = xs.length; }
        const finals = v.map(x => x.final).filter(x => x != null);
        return { name: labelFn(k), final: E.mean(finals), n: finals.length, byRole, nRole };
      }).filter(g => g.n);
    };
    if (by === 'area') return groupBy(p => p.area, k => k);
    if (by === 'cargo') return groupBy(p => p.cargo, k => titleCase(k));
    return E.aggregateComps(people, S.model.compCatalog).filter(c => !cat || c.category === cat);
  }

  // Ordena: por puntaje, o agrupado por tipo (organizacionales primero) cuando se ven ambas
  function orderGroups(groups, by) {
    const byScore = (a, b) => (b.final ?? 0) - (a.final ?? 0);
    if (by !== 'comp' || chartCat()) return { groups: groups.sort(byScore), bands: null };
    const cats = S.model.categories.slice().sort((a, b) => /organiz/i.test(b) - /organiz/i.test(a));
    const out = [], bands = [];
    for (const c of cats) {
      const g = groups.filter(x => x.category === c).sort(byScore);
      if (!g.length) continue;
      bands.push({ label: catShort(c), start: out.length, end: out.length + g.length - 1 });
      out.push(...g);
    }
    return { groups: out, bands };
  }

  const CHART_TITLE = { comp: 'Promedio por competencia', area: 'Promedio por área', cargo: 'Promedio por cargo' };

  function compChartConfig(groups, mode, opts = {}) {
    const grid = cssv('--line'), ink = cssv('--muted');
    let datasets;
    if (mode === 'final') {
      datasets = [{
        label: 'Promedio ponderado', data: groups.map(c => c.final),
        backgroundColor: groups.map(c => hexA(semColor(semKey(c.final)), c.n < LOW_BASE ? 0.4 : 0.95)),
        borderRadius: 4, barPercentage: 0.75, maxBarThickness: 26,
      }];
    } else {
      datasets = ROLES.map(r => ({
        label: ROLE_LABEL[r], data: groups.map(c => c.byRole[r]),
        backgroundColor: cssv(r === 'AUTOEVALUACION' ? '--auto' : r === 'JEFE' ? '--jefe' : '--colab'),
        borderRadius: 3, barPercentage: 0.85, categoryPercentage: 0.8, maxBarThickness: 14,
      })).filter(d => d.data.some(v => v != null));
    }
    // Escribe el valor al final de cada barra
    const valueLabels = {
      id: 'valueLabels',
      afterDatasetsDraw(chart) {
        const { ctx } = chart;
        ctx.save();
        ctx.font = `600 ${mode === 'final' ? 11 : 9}px Inter, Segoe UI, sans-serif`;
        ctx.fillStyle = cssv('--ink'); ctx.textBaseline = 'middle';
        chart.data.datasets.forEach((d, i) => {
          if (chart.getDatasetMeta(i).hidden) return;
          chart.getDatasetMeta(i).data.forEach((bar, j) => {
            const v = d.data[j];
            if (v == null) return;
            ctx.fillText(f2(v), Math.min(bar.x + 6, chart.chartArea.right - 26), bar.y);
          });
        });
        ctx.restore();
      },
    };
    // Franjas que separan competencias organizacionales y técnicas
    const bandPlugin = {
      id: 'typeBands',
      beforeDatasetsDraw(chart) {
        if (!opts.bands || opts.bands.length < 2) return;
        const { ctx, chartArea: ca, scales: { y } } = chart;
        const step = Math.abs(y.getPixelForValue(1) - y.getPixelForValue(0)) || 20;
        ctx.save();
        opts.bands.forEach((bd, i) => {
          const top = y.getPixelForValue(bd.start) - step / 2, bottom = y.getPixelForValue(bd.end) + step / 2;
          ctx.fillStyle = i % 2 ? 'rgba(0,0,0,0)' : hexA(cssv('--navy'), 0.07);
          ctx.fillRect(ca.left, top, ca.right - ca.left, bottom - top);
          ctx.strokeStyle = cssv('--line'); ctx.beginPath(); ctx.moveTo(ca.left, bottom); ctx.lineTo(ca.right, bottom); ctx.stroke();
          ctx.fillStyle = cssv('--navy'); ctx.font = '700 11px Inter, Segoe UI, sans-serif';
          ctx.textAlign = 'right'; ctx.textBaseline = 'top';
          ctx.fillText(bd.label, ca.right - 6, top + 3);
        });
        ctx.restore();
      },
    };
    return {
      type: 'bar',
      plugins: [bandPlugin, valueLabels],
      data: { labels: groups.map(c => splitLabel(c.name, 34)), datasets },
      options: {
        indexAxis: 'y', responsive: true, maintainAspectRatio: false,
        animation: opts.static ? false : { duration: 300 }, devicePixelRatio: opts.static ? 3 : window.devicePixelRatio,
        scales: {
          x: { min: SCALE_MIN, max: SCALE_MAX, ticks: { stepSize: 0.5, color: ink, callback: v => String(v).replace('.', ',') }, grid: { color: grid } },
          y: { ticks: { color: cssv('--ink'), font: { size: 11 } }, grid: { display: false } },
        },
        plugins: {
          legend: { display: mode !== 'final', labels: { color: ink, usePointStyle: true } },
          tooltip: { callbacks: { label: ctx => {
            const c = groups[ctx.dataIndex];
            if (mode === 'final') return ` ${f2(ctx.raw)} (${SEM_LABEL[semKey(ctx.raw)]}) con ${c.n} personas`;
            const r = ROLES.find(x => ROLE_LABEL[x] === ctx.dataset.label);
            return ` ${ctx.dataset.label}: ${f2(ctx.raw)} (${c.nRole[r]} personas)`;
          } } },
        },
      },
    };
  }

  function renderCompChart() {
    const canvas = $('#compChart');
    if (!S.dashPeople || !window.Chart || !canvas.offsetParent) return;
    const { groups, bands } = orderGroups(chartGroups(S.dashPeople, S.chartBy), S.chartBy);
    const c = chartCat();
    $('#chartTitle').textContent = CHART_TITLE[S.chartBy] + (c ? ` (${catShort(c).toLowerCase()})` : '');
    const perBar = S.chartMode === 'final' ? 30 : 44;
    canvas.parentElement.style.height = Math.max(240, groups.length * perBar + 70) + 'px';
    if (S.charts.comp) S.charts.comp.destroy();
    S.charts.comp = new Chart(canvas, compChartConfig(groups, S.chartMode, { bands }));
    const low = groups.filter(c => c.n < LOW_BASE);
    $('#compChartFoot').innerHTML = (S.chartMode === 'final' ? semLegend() : '') +
      (low.length ? `<div style="margin-top:6px">Barras atenuadas: menos de ${LOW_BASE} personas (${low.slice(0, 8).map(c => `${esc(c.name)}, ${c.n}`).join('; ')}${low.length > 8 ? '…' : ''}); tómalas como referencia.</div>` : '');
  }

  function areaStats(people) {
    return S.model.areas.map(a => {
      const list = people.filter(p => p.area === a);
      const finals = list.map(p => view(p).final).filter(v => v != null);
      const part = E.participation(list);
      return { area: a, n: list.length, avg: E.mean(finals), part, cargos: [...new Set(list.map(p => p.cargo))] };
    }).filter(x => x.n);
  }

  function renderAreas() {
    const pool = S.model.people.filter(p => !S.filters.cargo || p.cargo === S.filters.cargo);
    const rows = areaStats(pool).sort((a, b) => (b.avg ?? 0) - (a.avg ?? 0));
    $('#areasMeta').textContent = `${rows.length} ${rows.length === 1 ? 'área' : 'áreas'}`;
    const warn = S.model.hasArea ? '' : `<div class="notice"><i data-lucide="info"></i>El archivo de calificaciones no tiene la columna AREA; todos quedan en “Sin área”.</div>`;
    $('#areasBody').innerHTML = warn + `<div class="area-head"><span>Área</span><span>Promedio</span><span>Semáforo</span></div>` +
      rows.map(r => `<button class="area-row ${S.filters.area === r.area ? 'is-sel' : ''}" data-area="${esc(r.area)}" title="Filtrar por esta área">
        <span><span class="area-name">${esc(r.area)}</span><span class="area-sub"> ${r.n} ${r.n === 1 ? 'persona' : 'personas'}, ${r.cargos.length} ${r.cargos.length === 1 ? 'cargo' : 'cargos'}</span></span>
        <span class="area-score">${f2(r.avg)}</span><span>${semTag(r.avg)}</span></button>`).join('') +
      `<p class="kpi-s" style="margin-top:10px">Haz clic en un área para filtrar el dashboard.</p>`;
  }
  $('#areasBody').addEventListener('click', e => {
    const b = e.target.closest('[data-area]');
    if (!b) return;
    S.filters.area = S.filters.area === b.dataset.area ? '' : b.dataset.area;
    $('#fArea').value = S.filters.area; fillCargos(); renderDashboard();
  });

  const pct = (a, b) => b ? Math.round(a / b * 100) : 0;
  const partRow = (label, done, assigned) => `<div class="part-row"><span>${label}</span><span class="num">${done} de ${assigned}</span>
    <div class="pbar"><span style="width:${pct(done, assigned)}%"></span></div><span class="part-pct">${pct(done, assigned)} %</span></div>`;

  function participationHTML(people) {
    const t = E.participation(people);
    let h = `<div class="part-big">
      <div><b>${t.assigned}</b><span>evaluaciones asignadas</span></div>
      <div><b>${t.done}</b><span>realizadas (${pct(t.done, t.assigned)} %)</span></div>
      <div><b>${t.empty}</b><span>sin diligenciar</span></div></div>
      <p class="kpi-s">De las realizadas, ${t.complete} quedaron completas y ${t.partial} con preguntas sin responder.</p>
      <div class="part-sub">Por tipo de evaluación</div>`;
    h += ROLES.filter(r => t.byRole[r].assigned).map(r => partRow(`<span class="dot dot-${r}"></span> ${ROLE_LABEL[r]}`, t.byRole[r].done, t.byRole[r].assigned)).join('');
    const areas = areaStats(people);
    if (areas.length) {
      h += `<div class="part-sub">Por área</div>`;
      h += areas.sort((a, b) => pct(b.part.done, b.part.assigned) - pct(a.part.done, a.part.assigned))
        .map(a => partRow(esc(a.area), a.part.done, a.part.assigned)).join('');
    }
    return { html: h, total: t };
  }

  function renderParticipation() {
    const { html, total } = participationHTML(S.dashPeople);
    $('#partMeta').textContent = `${pct(total.done, total.assigned)} % de participación`;
    $('#partBody').innerHTML = html;
  }

  function dumbbell(p) {
    const v = view(p);
    const a = v.roles.AUTOEVALUACION, j = v.roles.JEFE;
    let h = '<div class="dumb"><div class="dumb-track"></div>';
    for (let t = SCALE_MIN; t <= SCALE_MAX; t++) h += `<div class="dumb-tick" style="left:${pos(t)}%"></div>`;
    if (a != null && j != null) {
      const l = Math.min(pos(a), pos(j)), w = Math.abs(pos(a) - pos(j));
      h += `<div class="dumb-link" style="left:${l}%;width:${w}%"></div>`;
    }
    if (j != null) h += `<div class="dumb-pt" title="Jefe ${f2(j)}" style="left:${pos(j)}%;background:var(--jefe)"></div>`;
    if (a != null) h += `<div class="dumb-pt" title="Autoevaluación ${f2(a)}" style="left:${pos(a)}%;background:var(--auto)"></div>`;
    return h + '</div>';
  }

  function sortPeople(rows) {
    const { key, dir } = S.sort;
    const val = p => key === 'flags' ? personIssues(p).length : key === 'combo' ? comboLabel(p.combo)
      : key === 'final' ? view(p).final : key === 'gap' ? view(p).gap : p[key];
    return rows.sort((a, b) => {
      const x = val(a), y = val(b);
      if (x == null && y == null) return 0;
      if (x == null) return 1;
      if (y == null) return -1;
      return (typeof x === 'string' ? x.localeCompare(y, 'es') : x - y) * dir;
    });
  }

  function renderTable() {
    const q = fold(S.search);
    const rows = sortPeople(S.dashPeople.filter(p => !q || fold(p.name + ' ' + p.cargo + ' ' + p.id + ' ' + p.area).includes(q)));
    const { key, dir } = S.sort;
    S.order = rows.map(p => p.id);
    $$('#peopleTable th').forEach(th => {
      th.classList.toggle('sort-asc', th.dataset.sort === key && dir === 1);
      th.classList.toggle('sort-desc', th.dataset.sort === key && dir === -1);
    });
    $('#peopleTable tbody').innerHTML = rows.map(p => {
      const v = view(p);
      const g = E.gapLabel(v.gap);
      const n = personIssues(p).length;
      return `<tr>
        <td><div class="p-name">${esc(titleCase(p.name))}</div><div class="p-sub">${esc(p.cargo.toLowerCase())}, ${esc(p.id)}</div></td>
        <td>${esc(p.area)}</td>
        <td><div class="combo-tags">${p.rolesPresent.map(r => `<span class="rtag rtag-${r}" title="${ROLE_LABEL[r]} ${f2(p.roleScores[r])}">${ROLE_SHORT[r]}</span>`).join('') || '<span class="na">—</span>'}</div></td>
        <td class="num"><span class="score">${f2(v.final)}</span></td>
        <td>${semTag(v.final)}</td>
        <td><div class="dumb-cell">${dumbbell(p)}<span class="gap-txt gap-${g.key}" title="${esc(g.text)}">${v.gap == null ? 'n/d' : sgn(v.gap)}</span></div></td>
        <td class="num"><span class="flag-n ${n ? '' : 'flag-0'}" title="Observaciones de calificación">${n ? '<i data-lucide="alert-triangle"></i>' : ''}${n}</span></td>
        <td><button class="btn btn-sm" data-open="${esc(p.id)}"><i data-lucide="file-text"></i>Ver reporte</button></td>
      </tr>`;
    }).join('') || '<tr><td colspan="8" class="na" style="text-align:center;padding:30px">Ningún colaborador coincide con la búsqueda.</td></tr>';
    icons();
  }
  $('#tableSearch').addEventListener('input', e => { S.search = e.target.value; renderTable(); });
  $$('#peopleTable th[data-sort]').forEach(th => th.addEventListener('click', () => {
    const k = th.dataset.sort;
    S.sort = { key: k, dir: S.sort.key === k ? -S.sort.dir : (k === 'name' || k === 'area' || k === 'combo' ? 1 : -1) };
    renderTable();
  }));
  $('#peopleTable').addEventListener('click', e => {
    const b = e.target.closest('[data-open]');
    if (!b) return;
    selectPerson(b.dataset.open);
    showView('report');
  });

  /* =================================================================
     4. REPORTE INDIVIDUAL
     ================================================================= */
  const pInput = $('#personSearch'), pList = $('#personList');
  let pIdx = -1, pItems = [];
  function openList() {
    if (!S.model) return;
    const q = fold(pInput.value.trim());
    pItems = S.model.people.filter(p => !q || fold(p.name + ' ' + p.id + ' ' + p.cargo).includes(q)).slice(0, 60);
    pIdx = pItems.length ? 0 : -1;
    pList.innerHTML = pItems.map((p, i) => `<li role="option" data-id="${esc(p.id)}" aria-selected="${i === pIdx}">
      <div class="p-name">${esc(titleCase(p.name))}</div><div class="p-sub">${esc(p.cargo.toLowerCase())}, ${esc(p.id)}</div></li>`).join('')
      || '<li class="na">Sin coincidencias</li>';
    pList.hidden = false; pInput.setAttribute('aria-expanded', 'true');
  }
  function closeList() { pList.hidden = true; pInput.setAttribute('aria-expanded', 'false'); }
  function markIdx() { $$('li', pList).forEach((li, i) => li.setAttribute('aria-selected', i === pIdx)); const li = $$('li', pList)[pIdx]; if (li) li.scrollIntoView({ block: 'nearest' }); }
  pInput.addEventListener('input', openList);
  pInput.addEventListener('focus', openList);
  pInput.addEventListener('keydown', e => {
    if (pList.hidden && e.key === 'ArrowDown') return openList();
    if (e.key === 'ArrowDown') { pIdx = Math.min(pItems.length - 1, pIdx + 1); markIdx(); e.preventDefault(); }
    else if (e.key === 'ArrowUp') { pIdx = Math.max(0, pIdx - 1); markIdx(); e.preventDefault(); }
    else if (e.key === 'Enter' && pItems[pIdx]) { selectPerson(pItems[pIdx].id); e.preventDefault(); }
    else if (e.key === 'Escape') closeList();
  });
  pList.addEventListener('mousedown', e => { const li = e.target.closest('li[data-id]'); if (li) { e.preventDefault(); selectPerson(li.dataset.id); } });
  document.addEventListener('click', e => { if (!e.target.closest('.picker')) closeList(); });

  function selectPerson(id) {
    const p = S.model.peopleById.get(id);
    if (!p) return;
    if (p.final === undefined) E.applyWeights(S.model, S.weights);
    S.current = p;
    pInput.value = titleCase(p.name);
    closeList();
    renderReport(p);
  }

  function evalList(p) {
    return [...p.evals.values()]
      .sort((a, b) => (ROLES.indexOf(a.rol) - ROLES.indexOf(b.rol)) || a.name.localeCompare(b.name, 'es'))
      .map(e => ({ ...e, label: titleCase(e.name) }));
  }

  function rulerHTML(p) {
    const pts = p.rolesPresent.map(r => ({ r, v: p.roleScores[r] }));
    // separa etiquetas que caen muy cerca
    const sorted = pts.slice().sort((a, b) => a.v - b.v);
    const lastInRow = [-99, -99, -99];
    sorted.forEach(x => {
      const ps = pos(x.v);
      let r = lastInRow.findIndex(l => ps - l >= 24);
      if (r < 0) r = 2;
      x.row = r; lastInRow[r] = ps;
    });
    let h = '<div class="ruler"><div class="ruler-track"><div class="ruler-line"></div>';
    for (let t = SCALE_MIN; t <= SCALE_MAX; t += 0.5) h += `<div class="ruler-tick" style="left:${pos(t)}%">${t % 1 ? '' : t}</div>`;
    if (p.final != null) h += `<div class="ruler-pt is-final" style="left:${pos(p.final)}%;background:var(--final)" title="Final ${f2(p.final)}"></div>`;
    for (const x of sorted) {
      const col = x.r === 'AUTOEVALUACION' ? 'auto' : x.r === 'JEFE' ? 'jefe' : 'colab';
      h += `<div class="ruler-pt" style="left:${pos(x.v)}%;background:var(--${col})"></div>`;
      h += `<div class="ruler-lbl" style="left:${pos(x.v)}%;top:${x.row * 15}px;color:var(--${col})">${ROLE_SHORT[x.r]} ${f2(x.v)}</div>`;
    }
    h += '</div><div class="ruler-legend">';
    for (const r of p.rolesPresent) {
      const v = p.roleScores[r];
      const extra = r === 'COLABORADOR' && p.colabCount ? ` (${p.colabCount} ${p.colabCount === 1 ? 'persona' : 'personas'})` : '';
      h += `<span><span class="dot dot-${r}"></span>${ROLE_LABEL[r]}: <b>${v == null ? 'sin datos' : f2(v)}</b>${extra}</span>`;
    }
    h += `<span><span class="dot dot-FINAL"></span>Final ponderado: <b>${f2(p.final)}</b></span></div></div>`;
    return h;
  }

  function gapHTML(p) {
    const g = E.gapLabel(p.gap);
    let text;
    if (p.gap == null) {
      const miss = [];
      if (p.roleScores.AUTOEVALUACION == null) miss.push('autoevaluación');
      if (p.roleScores.JEFE == null) miss.push('evaluación del jefe');
      text = `No se puede comparar porque no tiene ${miss.join(' ni ')} con notas válidas.`;
    } else if (g.key === 'over') text = `Se califica ${f2(p.gap)} puntos por encima de lo que lo califica su jefe. Conviene revisar juntos las competencias donde la diferencia es mayor.`;
    else if (g.key === 'under') text = `Se califica ${f2(-p.gap)} puntos por debajo de su jefe: reconoce menos sus logros de lo que su jefe observa.`;
    else text = `Su autoevaluación y la del jefe difieren apenas ${f2(Math.abs(p.gap))} puntos: ambas visiones coinciden.`;
    let rows = '';
    const comps = [...p.comps.values()].filter(c => c.scores.AUTOEVALUACION != null && c.scores.JEFE != null)
      .map(c => ({ name: c.name, d: c.scores.AUTOEVALUACION - c.scores.JEFE })).sort((a, b) => b.d - a.d);
    if (comps.length) {
      const m = Math.max(1, ...comps.map(c => Math.abs(c.d)));
      rows = '<div class="gap-rows">' + comps.map(c => {
        const w = Math.abs(c.d) / m * 50;
        const style = c.d >= 0 ? `left:50%;width:${w}%;background:var(--amber)` : `right:50%;width:${w}%;background:var(--auto)`;
        return `<div class="gap-row"><span class="gap-row-name" title="${esc(c.name)}">${esc(c.name)}</span><div class="div-bar"><span style="${style}"></span></div><span class="v">${sgn(c.d)}</span></div>`;
      }).join('') + `</div><p class="kpi-s" style="margin-top:10px">A la derecha: se califica por encima del jefe. A la izquierda: por debajo. Se considera alineado si la diferencia es menor a ${f2(E.GAP_THRESHOLD)}.</p>`;
    }
    return `<div class="gap-card"><span class="gap-badge ${g.key}">${esc(g.text)}${p.gap != null ? ` (${sgn(p.gap)})` : ''}</span><p class="gap-text">${esc(text)}</p>${rows}</div>`;
  }

  function behaviorNotes(b, labels, shown = ROLES) {
    const notes = [];
    for (const r of shown) {
      const br = b.byRole[r];
      if (!br || !br.missing) continue;
      if (r === 'COLABORADOR') notes.push(`<span class="note">Colab.: ${br.missing} de ${br.assigned} sin calificar</span>`);
      else notes.push(`<span class="note">${ROLE_SHORT[r]}: sin calificar</span>`);
    }
    for (const d of b.dups.filter(x => shown.includes(x.rol))) {
      const who = d.rol === 'COLABORADOR' ? (labels.get(d.evaluador + '|' + d.rol) || 'Colaborador') : ROLE_SHORT[d.rol];
      notes.push(d.contradictory
        ? `<span class="note">${esc(who)}: registrado ${d.count} veces con notas ${d.values.join(' y ')}; se promedian</span>`
        : `<span class="note note-muted">${esc(who)}: registrado ${d.count} veces con la misma nota</span>`);
    }
    return notes.length ? `<div class="beh-notes">${notes.join('')}</div>` : '';
  }

  const cell = (br, r) => {
    if (!br || br.score == null) return '<span class="na">—</span>';
    return f2(br.score) + (r === 'COLABORADOR' && br.n > 1 ? `<small class="na"> (${br.n})</small>` : '');
  };

  function observationsHTML(p, evals) {
    const obs = [];
    const add = (t, s, ok) => obs.push(`<div class="obs avoid-break ${ok ? 'ok' : ''}"><i data-lucide="${ok ? 'check-circle-2' : 'alert-triangle'}"></i><div>${t}${s ? `<small>${s}</small>` : ''}</div></div>`);
    for (const r of ['AUTOEVALUACION', 'JEFE']) {
      const list = evals.filter(e => e.rol === r);
      if (!list.length) add(`No tiene ${r === 'JEFE' ? 'evaluación del jefe' : 'autoevaluación'} asignada.`, 'Su puntaje final se calcula sin esta fuente.');
      for (const e of list) {
        if (e.state === 'vacia') add(`${ROLE_LABEL[r]} sin diligenciar (0 de ${e.total} ítems).`, 'Esta fuente no entra en el puntaje final.');
        else if (e.state === 'parcial') add(`${ROLE_LABEL[r]} incompleta: respondió ${e.answered} de ${e.total} ítems.`, 'Solo se usan los ítems con nota.');
      }
    }
    const col = evals.filter(e => e.rol === 'COLABORADOR');
    if (col.length) {
      const c = { completa: 0, parcial: 0, vacia: 0 };
      col.forEach(e => c[e.state]++);
      if (c.parcial || c.vacia)
        add(`De ${col.length} colaboradores asignados: ${c.completa} respondieron completo, ${c.parcial} a medias y ${c.vacia} no respondieron.`,
          'Las notas vacías no se cuentan. El detalle por evaluador está en la tabla.');
    }
    if (p.dupNotes.length) {
      const k = p.dupNotes.filter(d => d.contradictory).length;
      add(`${p.dupNotes.length} ${p.dupNotes.length === 1 ? 'ítem quedó registrado' : 'ítems quedaron registrados'} más de una vez por el mismo evaluador.`,
        `Se incluyen en el cálculo promediando sus registros${k ? `; ${k} con notas distintas` : ''}. El detalle aparece junto a cada comportamiento.`);
    }
    for (const inv of p.invalid)
      add(`Se excluyeron ${inv.rows} registros con error de asignación: aparece evaluándose a sí misma con rol ${ROLE_LABEL[inv.rol]}.`, 'No afectan el puntaje.');
    if (p.rolesPresent.length === 1) add(`Medición parcial: el puntaje final sale solo de ${ROLE_LABEL[p.rolesPresent[0]].toLowerCase()}.`, '');
    if (!obs.length) add('Todas las evaluaciones asignadas se diligenciaron completas.', '', true);

    const table = `<table class="eval-table avoid-break"><thead><tr><th>Fuente</th><th>Evaluador</th><th>Ítems con nota</th><th>Estado</th></tr></thead><tbody>${
      evals.map(e => `<tr><td><span class="dot dot-${e.rol}"></span> ${ROLE_LABEL[e.rol]}</td><td>${esc(e.label)}</td><td>${e.answered} de ${e.total}</td>
      <td><span class="state state-${e.state}">${e.state === 'completa' ? 'Completa' : e.state === 'parcial' ? 'Parcial' : 'Sin diligenciar'}</span></td></tr>`).join('')
    }</tbody></table>`;
    return `<div class="obs-list">${obs.join('')}</div>${table}`;
  }

  function feedbackHTML(p) {
    const empty = '<p class="fb-empty">Sin retroalimentación cualitativa registrada.</p>';
    const by = f => `<div class="fb-by"><span class="rtag rtag-${f.rol === 'SIN_ROL' ? 'JEFE' : f.rol}">${f.rol === 'SIN_ROL' ? 'Evaluador' : ROLE_LABEL[f.rol]}</span>${esc(titleCase(f.name || f.evaluador))}</div>`;
    const strong = p.feedback.length ? p.feedback.map(f => `<div class="fb-item avoid-break"><p>${esc(f.feedback || 'Sin comentario.')}</p>${by(f)}</div>`).join('') : empty;
    const improve = p.feedback.length ? p.feedback.map(f => `<div class="fb-item avoid-break">${f.aspects ? `<p>${esc(f.aspects)}</p>` : '<p class="fb-empty">No registró aspectos a mejorar.</p>'}${by(f)}</div>`).join('') : empty;
    return `<div class="fb-grid">
      <div class="fb-box strong"><h3><i data-lucide="sparkles"></i>Fortalezas y desempeño</h3>${strong}</div>
      <div class="fb-box improve"><h3><i data-lucide="trending-up"></i>Aspectos a mejorar</h3>${improve}</div></div>`;
  }

  function renderReport(p, opts = {}) {
    const evals = evalList(p);
    const labels = new Map(evals.map(e => [e.evaluador + '|' + e.rol, e.label]));
    const jefes = evals.filter(e => e.rol === 'JEFE').map(e => titleCase(e.name));
    const today = new Date().toLocaleDateString('es-CO', { day: 'numeric', month: 'long', year: 'numeric' });
    const weightsTxt = p.rolesPresent.map(r => `<span><span class="dot dot-${r}"></span>${ROLE_LABEL[r]} ${p.weights[r] ?? 0} %</span>`).join('');

    // Solo se muestran las columnas de las fuentes que sí calificaron
    const shown = p.rolesPresent.length ? p.rolesPresent : ROLES;
    const comps = [...p.comps.values()].sort((a, b) => (/organiz/i.test(a.category) - /organiz/i.test(b.category)) || a.name.localeCompare(b.name, 'es'));
    const acc = comps.map(c => `
      <details class="comp-acc">
        <summary class="avoid-break"><span class="c-name">${esc(c.name)}<span class="c-cat">${esc(c.category)}, ${c.behaviors.size} comportamientos</span></span>
          ${shown.map(r => `<span class="c-v">${c.scores[r] == null ? '<span class="na">—</span>' : f2(c.scores[r])}</span>`).join('')}
          <span class="c-v is-final">${semDot(c.final)} ${f2(c.final)}</span></summary>
        <div class="beh-list">${[...c.behaviors.values()].map(b => `
          <div class="beh-row"><div class="beh-desc">${esc(b.desc)}${behaviorNotes(b, labels, shown)}</div>${shown.map(r => `<span class="c-v">${cell(b.byRole[r], r)}</span>`).join('')}<span class="c-v is-final">${f2(E.weightedScore(Object.fromEntries(ROLES.map(r => [r, b.byRole[r] ? b.byRole[r].score : null])), ROLES, p.weights))}</span><span></span></div>`).join('')}
        </div>
      </details>`).join('');

    $('#report').innerHTML = `
      <header class="r-head avoid-break">
        <div>
          <div class="r-org">${esc(appTitle())}</div>
          <div class="r-name">${esc(titleCase(p.name))}</div>
          <dl class="r-meta">
            <div><dt>Cargo</dt><dd>${esc(titleCase(p.cargo))}</dd></div>
            <div><dt>Documento</dt><dd>${esc(p.id)}</dd></div>
            <div><dt>Jefe evaluador</dt><dd>${jefes.length ? esc(jefes.join(', ')) : 'Sin jefe asignado'}</dd></div>
            <div><dt>Área</dt><dd>${esc(p.area)}</dd></div>
            <div><dt>Fecha del reporte</dt><dd>${esc(today)}</dd></div>
          </dl>
        </div>
        <div class="r-final">
          <div class="r-final-v">${f2(p.final)}<small> / 4</small></div>
          <div class="r-final-l">Puntaje final ponderado</div>
          <div class="r-sem"><span class="sem-pill">${semDot(p.final)}${SEM_LABEL[semKey(p.final)]}</span></div>
          <div class="r-weights">${weightsTxt}</div>
        </div>
      </header>

      <section class="r-sec avoid-break"><h2>Resultado por rol</h2>${rulerHTML(p)}</section>

      <section class="r-sec r-grid">
        <div class="avoid-break"><h2 style="margin-bottom:12px">Mapa de competencias</h2><div class="radar-box"><canvas id="radarChart"></canvas></div></div>
        <div class="avoid-break"><h2 style="margin-bottom:12px">Análisis de brecha</h2>${gapHTML(p)}</div>
      </section>

      <section class="r-sec">
        <h2>Detalle por comportamiento</h2>
        <div class="comp-head"><span style="text-align:left">Competencia</span>${shown.map(r => `<span>${ROLE_SHORT[r]}</span>`).join('')}<span>Final</span><span></span></div>
        ${acc}
        <p class="kpi-s">${shown.includes('COLABORADOR') ? 'El número entre paréntesis es la cantidad de colaboradores que calificaron el ítem. ' : ''}“—” indica que no hay nota válida. ${semLegend()}</p>
      </section>

      <section class="r-sec"><h2>Observaciones de la calificación</h2>${observationsHTML(p, evals)}</section>

      <section class="r-sec"><h2>Retroalimentación del jefe</h2>${feedbackHTML(p)}</section>

      <footer class="r-foot"><span>Escala de 1 a 4. Los promedios se calculan por competencia y luego se ponderan por fuente.</span><span>Generado el ${esc(today)}</span></footer>`;
    $('#report').style.setProperty('--rc', shown.length + 1);
    $('#report').hidden = false; $('#reportEmpty').hidden = true; $('#btnPdf').disabled = false;
    icons();
    renderRadar(p, opts.pdf ? { animation: false, dpr: 3 } : {});
    if (!opts.pdf) updateNav();
  }

  function resetReport() {
    S.current = null;
    $('#report').hidden = true; $('#report').innerHTML = '';
    $('#reportEmpty').hidden = false; $('#btnPdf').disabled = true; $('#reportNav').hidden = true;
    pInput.value = '';
  }

  /* ---------- anterior / siguiente ---------- */
  function navIds() {
    if (S.current && S.order.includes(S.current.id)) return S.order;
    return S.model.people.map(p => p.id);
  }
  function updateNav() {
    if (!S.current) { $('#reportNav').hidden = true; return; }
    const ids = navIds(), i = ids.indexOf(S.current.id);
    $('#reportNav').hidden = false;
    $('#navPos').textContent = `${i + 1} de ${ids.length}`;
    $('#btnPrev').disabled = i <= 0;
    $('#btnNext').disabled = i >= ids.length - 1;
  }
  function stepPerson(d) {
    const ids = navIds(), i = ids.indexOf(S.current.id), j = i + d;
    if (j < 0 || j >= ids.length) return;
    selectPerson(ids[j]);
    window.scrollTo({ top: 0 });
  }
  $('#btnPrev').addEventListener('click', () => stepPerson(-1));
  $('#btnNext').addEventListener('click', () => stepPerson(1));

  function renderRadar(p, opts = {}) {
    const canvas = $('#radarChart');
    if (!canvas || !window.Chart || !canvas.offsetParent) return;
    const comps = [...p.comps.values()];
    const color = { AUTOEVALUACION: cssv('--auto'), JEFE: cssv('--jefe'), COLABORADOR: cssv('--colab') };
    const datasets = p.rolesPresent.map(r => ({
      label: ROLE_LABEL[r], data: comps.map(c => c.scores[r]),
      borderColor: color[r], backgroundColor: hexA(color[r], 0.12), pointBackgroundColor: color[r],
      borderWidth: 2, pointRadius: 3, spanGaps: true,
    }));
    if (S.charts.radar) S.charts.radar.destroy();
    S.charts.radar = new Chart(canvas, {
      type: 'radar',
      data: { labels: comps.map(c => splitLabel(c.name)), datasets },
      options: {
        responsive: true, maintainAspectRatio: false, animation: opts.animation === false ? false : { duration: 400 },
        devicePixelRatio: opts.dpr || window.devicePixelRatio,
        scales: {
          r: {
            min: SCALE_MIN, max: SCALE_MAX, ticks: { stepSize: 0.5, backdropColor: 'transparent', color: cssv('--muted'), font: { size: 10 } },
            grid: { color: cssv('--line') }, angleLines: { color: cssv('--line') },
            pointLabels: { color: cssv('--ink'), font: { size: window.innerWidth < 500 ? 9 : 11, weight: '500' } },
          },
        },
        plugins: { legend: { position: 'bottom', labels: { color: cssv('--ink'), usePointStyle: true, padding: 14 } },
          tooltip: { callbacks: { label: ctx => ` ${ctx.dataset.label}: ${f2(ctx.raw)}` } } },
      },
    });
  }
  function splitLabel(s, width = 18) {
    const words = s.split(' '); const lines = []; let cur = '';
    for (const w of words) { if ((cur + ' ' + w).trim().length > width && cur) { lines.push(cur); cur = w; } else cur = (cur + ' ' + w).trim(); }
    if (cur) lines.push(cur);
    return lines;
  }

  /* =================================================================
     PDF: individual, compañía y ZIP
     ================================================================= */
  function busy(text, frac, cancellable) {
    $('#busy').hidden = false;
    $('#busyText').textContent = text;
    $('#busyBar').style.width = (frac == null ? 0 : Math.round(frac * 100)) + '%';
    $('#busyCancel').hidden = !cancellable;
  }
  const busyEnd = () => { $('#busy').hidden = true; };
  $('#busyCancel').addEventListener('click', () => { S.cancel = true; $('#busyText').textContent = 'Cancelando al terminar el reporte actual…'; });

  const pdfName = p => `Reporte_${slug(titleCase(p.name))}_${p.id}.pdf`;
  const pdfReady = () => {
    if (!window.html2pdf) { toast('No se pudo cargar la librería de PDF. Revisa la conexión a internet.'); return false; }
    return true;
  };

  function pdfWorker(el, filename, footer, light) {
    return html2pdf().set({
      margin: [10, 10, 14, 10], filename,
      image: { type: 'jpeg', quality: light ? 0.85 : 0.95 },
      html2canvas: { scale: light ? 1.5 : 2, useCORS: true, backgroundColor: '#FFFFFF', scrollX: 0, scrollY: 0, x: 0, y: 0, windowWidth: 718 },
      jsPDF: { unit: 'mm', format: 'a4', orientation: 'portrait' },
      pagebreak: { mode: ['css', 'legacy'], avoid: ['.avoid-break', '.beh-row', '.eval-table tr'] },
    }).from(el).toPdf().get('pdf').then(pdf => {
      const total = pdf.internal.getNumberOfPages();
      for (let i = 1; i <= total; i++) {
        pdf.setPage(i); pdf.setFontSize(8); pdf.setTextColor(110);
        pdf.text(footer, 10, 290);
        pdf.text(`Página ${i} de ${total}`, 200, 290, { align: 'right' });
      }
    });
  }

  // Prepara tema claro y un contenedor anclado en 0,0 para la captura
  async function inPdfMode(work) {
    const root = document.documentElement, prevTheme = root.getAttribute('data-theme');
    const host = document.createElement('div');
    host.className = 'pdf-host';
    const y = window.scrollY;
    root.setAttribute('data-theme', 'light');
    document.body.classList.add('pdf-mode');
    document.body.appendChild(host);
    window.scrollTo(0, 0);
    try { return await work(host); }
    finally {
      host.remove();
      document.body.classList.remove('pdf-mode');
      if (prevTheme == null) root.removeAttribute('data-theme'); else root.setAttribute('data-theme', prevTheme);
      window.scrollTo(0, y);
    }
  }

  // Mueve el reporte al contenedor de captura y lo devuelve a su lugar al final
  async function withReportInHost(host, fn) {
    const report = $('#report'), parent = report.parentNode, next = report.nextSibling;
    host.appendChild(report);
    try { return await fn(report); } finally { parent.insertBefore(report, next); }
  }
  async function renderPersonForPdf(p) {
    renderReport(p, { pdf: true });
    $$('details', $('#report')).forEach(d => { d.open = true; });
    await frames(3);
  }
  const personFooter = p => `${titleCase(p.name)}, documento ${p.id}`;

  function saveBlob(blob, name) {
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob); a.download = name;
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(a.href), 4000);
  }

  /* ---------- PDF individual ---------- */
  $('#btnPdf').addEventListener('click', async () => {
    const p = S.current;
    if (!p || !pdfReady()) return;
    busy('Generando PDF…');
    try {
      await inPdfMode(host => withReportInHost(host, async report => {
        await renderPersonForPdf(p);
        await pdfWorker(report, pdfName(p), personFooter(p)).save();
      }));
      toast('PDF descargado.');
    } catch (e) {
      console.error(e); toast('No se pudo generar el PDF. Intenta de nuevo.');
    } finally {
      renderReport(p); busyEnd();
    }
  });

  /* ---------- ZIP con todos los reportes individuales ---------- */
  $('#btnZip').addEventListener('click', async () => {
    if (!S.model || !pdfReady()) return;
    if (!window.JSZip) return toast('No se pudo cargar la librería ZIP. Revisa la conexión a internet.');
    const list = sortPeople(subset());
    if (!list.length) return toast('No hay colaboradores con los filtros actuales.');
    const prev = S.current;
    const base = slug(appTitle()) || 'evaluacion';
    const zip = new JSZip();
    const folder = zip.folder(`Reportes_${base}`);
    let done = 0;
    S.cancel = false;
    busy(`Preparando ${list.length} reportes…`, 0, true);
    try {
      await inPdfMode(host => withReportInHost(host, async report => {
        for (let i = 0; i < list.length && !S.cancel; i++) {
          const p = list[i];
          busy(`Generando reporte ${i + 1} de ${list.length}: ${titleCase(p.name)}`, i / list.length, true);
          await renderPersonForPdf(p);
          const blob = await pdfWorker(report, pdfName(p), personFooter(p), true).outputPdf('blob');
          folder.file(pdfName(p), blob);
          done++;
        }
      }));
      if (done) {
        busy('Comprimiendo archivos…', 1, false);
        const blob = await zip.generateAsync({ type: 'blob' });
        saveBlob(blob, `Reportes_${base}_${stamp()}.zip`);
        toast(S.cancel ? `Descarga parcial: ${done} de ${list.length} reportes.` : `Se descargaron ${done} reportes en un .zip.`);
      } else toast('Descarga cancelada.');
    } catch (e) {
      console.error(e); toast('No se pudieron generar todos los reportes. Intenta de nuevo.');
    } finally {
      busyEnd();
      if (prev) renderReport(prev); else resetReport();
    }
  });

  /* ---------- Informe de la compañía ---------- */
  function companyHTML(people) {
    const today = new Date().toLocaleDateString('es-CO', { day: 'numeric', month: 'long', year: 'numeric' });
    const comps = filteredComps(people);
    const sortedComps = comps.slice().sort((a, b) => (b.final ?? 0) - (a.final ?? 0));
    const combos = S.analysis.combos.map(c => ({ ...c, n: people.filter(p => p.combo === c.key).length })).filter(c => c.n);
    const fb = generalFb();
    const para = t => esc(t).replace(/\n/g, '<br>');
    const compRows = sortedComps.map(c => `<tr>
        <td>${esc(c.name)}${c.n < LOW_BASE ? ' <span class="note">base baja</span>' : ''}</td><td>${esc(catShort(c.category))}</td>
        <td class="num">${c.n}</td>${ROLES.map(r => `<td class="num">${f2(c.byRole[r])}</td>`).join('')}
        <td class="num"><b>${f2(c.final)}</b></td><td>${semTag(c.final)}</td><td class="num">${c.gap == null ? '—' : sgn(c.gap)}</td></tr>`).join('');
    const areaRows = areaStats(people).sort((a, b) => (b.avg ?? 0) - (a.avg ?? 0)).map(a => `<tr>
        <td>${esc(a.area)}</td><td class="num">${a.n}</td><td class="num"><b>${f2(a.avg)}</b></td><td>${semTag(a.avg)}</td>
        <td class="num">${a.part.done} de ${a.part.assigned}</td><td class="num">${pct(a.part.done, a.part.assigned)} %</td></tr>`).join('');
    const comboRows = combos.map(c => `<tr><td>${esc(comboLabel(c.key))}</td><td class="num">${c.n}</td>
        <td>${c.roles.map(r => `${ROLE_LABEL[r]} ${S.weights[c.key][r] ?? 0} %`).join(', ')}</td></tr>`).join('');
    const peopleRows = people.map(p => { const v = view(p); return `<tr>
        <td><b>${esc(titleCase(p.name))}</b><div class="p-sub">${esc(p.cargo.toLowerCase())}, ${esc(p.id)}</div></td><td>${esc(p.area)}</td>
        ${ROLES.map(r => `<td class="num">${f2(v.roles[r])}</td>`).join('')}
        <td class="num"><b>${f2(v.final)}</b></td><td>${semDot(v.final)}</td><td class="num">${v.gap == null ? '—' : sgn(v.gap)}</td></tr>`; }).join('');
    const fbBlocks = [
      fb.strengths && `<div class="fb-box strong avoid-break"><h3><i data-lucide="sparkles"></i>Fortalezas de la organización</h3><p>${para(fb.strengths)}</p></div>`,
      fb.improve && `<div class="fb-box improve avoid-break"><h3><i data-lucide="trending-up"></i>Aspectos a mejorar</h3><p>${para(fb.improve)}</p></div>`,
      fb.next && `<div class="fb-box avoid-break" style="border-left:4px solid var(--navy)"><h3 style="color:var(--navy)"><i data-lucide="flag"></i>Conclusiones y próximos pasos</h3><p>${para(fb.next)}</p></div>`,
    ].filter(Boolean);
    return `
    <article class="report company">
      <header class="r-head avoid-break">
        <div>
          <div class="r-org">${esc(appTitle())}</div>
          <div class="r-name">Informe consolidado de la compañía</div>
          <dl class="r-meta">
            <div><dt>Alcance</dt><dd>${esc(filtersText())}</dd></div>
            <div><dt>Colaboradores</dt><dd>${people.length}</dd></div>
            <div><dt>Fecha del informe</dt><dd>${esc(today)}</dd></div>
          </dl>
        </div>
      </header>
      <section class="r-sec avoid-break"><div class="kpis">${kpisHTML(people, comps)}</div></section>
      ${fbBlocks.length ? `<section class="r-sec"><h2>Retroalimentación general</h2><div style="display:grid;gap:12px">${fbBlocks.join('')}</div></section>` : ''}
      <section class="r-sec avoid-break"><h2>Promedio ponderado por competencia</h2>
        <div class="chart-box" id="companyChartBox"><canvas id="companyChart"></canvas></div>
        <div class="panel-foot">${semLegend()}</div>
      </section>
      <section class="r-sec avoid-break"><h2>Resultados por área</h2>
        <table class="eval-table"><thead><tr><th>Área</th><th class="num">Personas</th><th class="num">Promedio</th><th>Semáforo</th><th class="num">Evaluaciones realizadas</th><th class="num">Participación</th></tr></thead>
        <tbody>${areaRows}</tbody></table>
      </section>
      <section class="r-sec"><h2>Resultados por competencia y fuente</h2>
        <table class="eval-table"><thead><tr><th>Competencia</th><th>Tipo</th><th class="num">Personas</th><th class="num">Auto</th><th class="num">Jefe</th><th class="num">Colab.</th><th class="num">Ponderado</th><th>Semáforo</th><th class="num">Brecha</th></tr></thead>
        <tbody>${compRows}</tbody></table>
        <p class="panel-foot">Brecha = autoevaluación − jefe. “Base baja”: menos de ${LOW_BASE} personas evaluadas.</p>
      </section>
      <section class="r-sec avoid-break"><h2>Participación</h2>${participationHTML(people).html}</section>
      <section class="r-sec avoid-break"><h2>Ponderaciones aplicadas</h2>
        <table class="eval-table"><thead><tr><th>Evaluadores</th><th class="num">Personas</th><th>Pesos</th></tr></thead><tbody>${comboRows}</tbody></table>
      </section>
      <section class="r-sec"><h2>Resultados por colaborador</h2>
        <table class="eval-table"><thead><tr><th>Colaborador</th><th>Área</th><th class="num">Auto</th><th class="num">Jefe</th><th class="num">Colab.</th><th class="num">Final</th><th></th><th class="num">Brecha</th></tr></thead>
        <tbody>${peopleRows}</tbody></table>
      </section>
      <footer class="r-foot"><span>Escala de 1 a 4. Los promedios se calculan por competencia y luego se ponderan por fuente.</span><span>Generado el ${esc(today)}</span></footer>
    </article>`;
  }

  $('#btnCompanyPdf').addEventListener('click', async () => {
    if (!S.model || !pdfReady()) return;
    const people = sortPeople(subset());
    if (!people.length) return toast('No hay colaboradores con los filtros actuales.');
    busy('Generando informe de la compañía…');
    let chart = null;
    try {
      await inPdfMode(async host => {
        host.innerHTML = companyHTML(people);
        icons();
        const { groups, bands } = orderGroups(chartGroups(people, 'comp'), 'comp');
        $('#companyChartBox', host).style.height = Math.max(200, groups.length * 26 + 60) + 'px';
        chart = new Chart($('#companyChart', host), compChartConfig(groups, 'final', { static: true, bands }));
        await frames(3);
        await pdfWorker(host.firstElementChild, `Informe_compania_${slug(appTitle())}_${stamp()}.pdf`, `${appTitle()}, informe consolidado`).save();
      });
      toast('Informe de la compañía descargado.');
    } catch (e) {
      console.error(e); toast('No se pudo generar el informe. Intenta de nuevo.');
    } finally {
      if (chart) chart.destroy();
      busyEnd();
    }
  });

  /* ---------- tema del sistema ---------- */
  const mq = window.matchMedia('(prefers-color-scheme: dark)');
  (mq.addEventListener ? mq.addEventListener.bind(mq, 'change') : mq.addListener.bind(mq))(() => {
    if (document.body.classList.contains('pdf-mode')) return;
    renderCompChart(); if (S.current) renderRadar(S.current);
  });

  function restoreSession() {
    let name = null;
    try { name = sessionStorage.getItem(SESSION_KEY); } catch (e) { name = null; }
    const u = name && findUser(name);
    if (u) startSession(u); else setTimeout(() => $('#loginUser').focus(), 50);
  }
  restoreSession();
  icons();
})();
