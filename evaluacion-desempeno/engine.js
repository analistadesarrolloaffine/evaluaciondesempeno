/* =====================================================================
   Motor de datos - Evaluación de Desempeño 180°/360°
   Procesa en el navegador; no envía datos a ningún servidor.
   ===================================================================== */
(function (global) {
  'use strict';

  const ROLES = ['AUTOEVALUACION', 'JEFE', 'COLABORADOR'];
  const ROLE_LABEL = { AUTOEVALUACION: 'Autoevaluación', JEFE: 'Jefe', COLABORADOR: 'Colaboradores' };
  const ROLE_SHORT = { AUTOEVALUACION: 'Auto', JEFE: 'Jefe', COLABORADOR: 'Colab.' };
  const SCALE_MIN = 1;
  const SCALE_MAX = 4;
  const GAP_THRESHOLD = 0.25; // diferencia Auto − Jefe considerada "alineada"

  const REQ_SCORES = ['COMPETENCIA', 'NOMBRE_COMPETENCIA', 'COMPORTAMIENTO_DESC', 'CALIFICACION', 'EVALUADO',
    'NOMBRE_EVALUADO', 'CARGO_EVALUADO', 'EVALUADOR', 'NOMBRE_EVALUADOR', 'CARGO_EVALUADOR', 'ROL',
    'EVALUATION_DESCRIPTION'];
  const REQ_FEEDBACK = ['EVALUATION_DESCRIPTION', 'EVALUADO', 'EVALUADOR', 'FEEDBACK', 'ASPECTS'];

  /* ---------- utilidades ---------- */
  const stripAccents = s => s.normalize('NFD').replace(/[\u0300-\u036f]/g, '');
  const normKey = k => stripAccents(String(k).trim().toUpperCase()).replace(/\s+/g, '_');
  const str = v => (v == null ? '' : String(v).trim());
  const idStr = v => {
    if (v == null || v === '') return '';
    if (typeof v === 'number') return String(Math.round(v));
    return String(v).trim().replace(/\.0+$/, '');
  };
  const num = v => {
    if (v == null || v === '') return null;
    const n = typeof v === 'number' ? v : parseFloat(String(v).replace(',', '.'));
    return Number.isFinite(n) ? n : null;
  };
  const normRol = v => {
    const s = stripAccents(str(v).toUpperCase());
    if (s.startsWith('AUTO')) return 'AUTOEVALUACION';
    if (s.startsWith('JEFE')) return 'JEFE';
    if (s.startsWith('COLAB')) return 'COLABORADOR';
    return s;
  };
  const mean = arr => {
    const v = arr.filter(x => x != null && Number.isFinite(x));
    return v.length ? v.reduce((a, b) => a + b, 0) / v.length : null;
  };
  const normRows = rows => rows.map(r => {
    const o = {};
    for (const k in r) o[normKey(k)] = r[k];
    return o;
  });

  function missingColumns(rows, required) {
    const keys = new Set();
    rows.slice(0, 200).forEach(r => Object.keys(r).forEach(k => keys.add(normKey(k))));
    return required.filter(c => !keys.has(c));
  }

  /* ---------- construcción del modelo ---------- */
  function buildModel(rawScores, rawFeedback, opts = {}) {
    let rows = normRows(rawScores || []);
    let fbRows = normRows(rawFeedback || []);
    // Filtro opcional por campaña (EVALUATION_DESCRIPTION), definido por usuario en users.js
    if (opts.campaign) {
      const want = stripAccents(String(opts.campaign).trim().toLowerCase());
      const match = r => stripAccents(str(r.EVALUATION_DESCRIPTION).toLowerCase()) === want;
      rows = rows.filter(match);
      fbRows = fbRows.filter(match);
    }

    const quality = {
      totalRows: rows.length, validCells: 0, nullCells: 0, outOfRange: 0, unknownRole: 0,
      invalidAssignRows: 0, invalidAssign: [], dupItems: 0, dupExtraRows: 0, contradictions: 0,
      evalsEmpty: 0, evalsPartial: 0, fbRows: fbRows.length, fbMatched: 0, fbUnmatched: [],
    };
    const people = new Map();
    const campaigns = new Set();
    const invalidSeen = new Set();

    for (const r of rows) {
      const ev = idStr(r.EVALUADO);
      if (!ev) continue;
      const er = idStr(r.EVALUADOR);
      const rol = normRol(r.ROL);
      const campaign = str(r.EVALUATION_DESCRIPTION);
      if (campaign) campaigns.add(campaign);

      let p = people.get(ev);
      if (!p) {
        p = {
          id: ev, name: str(r.NOMBRE_EVALUADO), cargo: str(r.CARGO_EVALUADO), campaign, area: '',
          comps: new Map(), evals: new Map(), invalid: [], feedback: [],
        };
        people.set(ev, p);
      }

      if (!p.area) p.area = str(r.AREA != null ? r.AREA : r.AREA_EVALUADO);
      if (!ROLES.includes(rol)) { quality.unknownRole++; continue; }

      // Error de asignación: alguien que se evalúa a sí mismo con un rol distinto de AUTOEVALUACION
      if (rol !== 'AUTOEVALUACION' && ev === er) {
        quality.invalidAssignRows++;
        const k = ev + '|' + rol;
        let inv = p.invalid.find(x => x.rol === rol);
        if (!inv) { inv = { rol, rows: 0 }; p.invalid.push(inv); }
        inv.rows++;
        if (!invalidSeen.has(k)) { invalidSeen.add(k); quality.invalidAssign.push({ id: ev, name: p.name, rol }); }
        continue;
      }

      const compName = str(r.NOMBRE_COMPETENCIA) || 'Sin competencia';
      let c = p.comps.get(compName);
      if (!c) { c = { name: compName, category: str(r.COMPETENCIA), behaviors: new Map() }; p.comps.set(compName, c); }
      const desc = str(r.COMPORTAMIENTO_DESC) || 'Sin descripción';
      let b = c.behaviors.get(desc);
      if (!b) { b = { desc, entries: [] }; c.behaviors.set(desc, b); }

      const raw = num(r.CALIFICACION);
      let status = 'ok';
      if (raw == null) { status = 'null'; quality.nullCells++; }
      else if (raw < SCALE_MIN || raw > SCALE_MAX) { status = 'range'; quality.outOfRange++; }
      else quality.validCells++;
      b.entries.push({ evaluador: er, rol, value: status === 'ok' ? raw : null, raw, status });

      const ek = er + '|' + rol;
      let e = p.evals.get(ek);
      if (!e) {
        e = { evaluador: er, name: str(r.NOMBRE_EVALUADOR), cargo: str(r.CARGO_EVALUADOR), rol,
          total: 0, answered: 0, nulls: 0, outOfRange: 0 };
        p.evals.set(ek, e);
      }
      e.total++;
      if (status === 'ok') e.answered++;
      else if (status === 'null') e.nulls++;
      else e.outOfRange++;
    }

    // Duplicados, contradicciones y puntajes
    for (const p of people.values()) {
      p.dupNotes = [];
      for (const c of p.comps.values()) {
        c.scores = {}; c.n = {};
        const behScoresByRole = { AUTOEVALUACION: [], JEFE: [], COLABORADOR: [] };
        for (const b of c.behaviors.values()) {
          b.byRole = {}; b.dups = [];
          const groups = new Map();
          for (const en of b.entries) {
            const k = en.evaluador + '|' + en.rol;
            if (!groups.has(k)) groups.set(k, []);
            groups.get(k).push(en);
          }
          for (const [k, list] of groups) {
            if (list.length > 1) {
              const vals = list.filter(x => x.status === 'ok').map(x => x.value);
              const distinct = [...new Set(vals)];
              const contradictory = distinct.length > 1;
              quality.dupItems++;
              quality.dupExtraRows += list.length - 1;
              if (contradictory) quality.contradictions++;
              const d = { evaluador: list[0].evaluador, rol: list[0].rol, count: list.length, values: vals, contradictory };
              b.dups.push(d);
              p.dupNotes.push({ comp: c.name, desc: b.desc, ...d });
            }
          }
          for (const rol of ROLES) {
            // promedio por evaluador (los registros repetidos se incluyen en el cálculo)
            const perEval = [];
            const nulls = [];
            for (const [k, list] of groups) {
              if (!k.endsWith('|' + rol)) continue;
              const m = mean(list.filter(x => x.status === 'ok').map(x => x.value));
              if (m != null) perEval.push(m);
              else nulls.push(list[0].evaluador);
            }
            const s = mean(perEval);
            b.byRole[rol] = { score: s, n: perEval.length, missing: nulls.length,
              assigned: perEval.length + nulls.length };
            if (s != null) behScoresByRole[rol].push(s);
          }
        }
        for (const rol of ROLES) {
          c.scores[rol] = mean(behScoresByRole[rol]);
          c.n[rol] = behScoresByRole[rol].length;
        }
      }
      // Promedio por rol: primero por competencia, luego entre competencias
      p.roleScores = {};
      for (const rol of ROLES) p.roleScores[rol] = mean([...p.comps.values()].map(c => c.scores[rol]));
      p.rolesPresent = ROLES.filter(r => p.roleScores[r] != null);
      p.combo = p.rolesPresent.length ? p.rolesPresent.join('+') : 'NINGUNO';
      p.colabCount = [...p.evals.values()].filter(e => e.rol === 'COLABORADOR' && e.answered > 0).length;
      p.gap = (p.roleScores.AUTOEVALUACION != null && p.roleScores.JEFE != null)
        ? p.roleScores.AUTOEVALUACION - p.roleScores.JEFE : null;

      // Estado de cada evaluación asignada
      for (const e of p.evals.values()) {
        e.state = e.answered === 0 ? 'vacia' : (e.answered < e.total ? 'parcial' : 'completa');
        if (e.state === 'vacia') quality.evalsEmpty++;
        if (e.state === 'parcial') quality.evalsPartial++;
      }
      p.profileKey = [...p.comps.keys()].sort((a, b) => a.localeCompare(b, 'es')).join(' | ');
    }

    // Retroalimentación cualitativa: cruce EVALUADO + EVALUADOR
    for (const r of fbRows) {
      const ev = idStr(r.EVALUADO), er = idStr(r.EVALUADOR);
      const p = people.get(ev);
      if (!p) { quality.fbUnmatched.push({ evaluado: ev, evaluador: er }); continue; }
      const evals = [...p.evals.values()].filter(e => e.evaluador === er);
      const pick = evals.find(e => e.rol === 'JEFE') || evals[0];
      p.feedback.push({
        evaluador: er, name: pick ? pick.name : '', cargo: pick ? pick.cargo : '',
        rol: pick ? pick.rol : 'SIN_ROL', feedback: str(r.FEEDBACK), aspects: str(r.ASPECTS),
      });
      quality.fbMatched++;
    }

    // Perfiles de cuestionario (grupos de competencias)
    const profMap = new Map();
    for (const p of people.values()) {
      if (!profMap.has(p.profileKey)) profMap.set(p.profileKey, { key: p.profileKey, people: [], comps: [...p.comps.keys()] });
      profMap.get(p.profileKey).people.push(p);
    }
    const profiles = [...profMap.values()].sort((a, b) => b.people.length - a.people.length);
    profiles.forEach((pr, i) => {
      pr.label = 'Perfil ' + String.fromCharCode(65 + (i % 26)) + (i >= 26 ? Math.floor(i / 26) : '');
      pr.comps.sort((a, b) => a.localeCompare(b, 'es'));
      pr.cargos = [...new Set(pr.people.map(x => x.cargo))].sort((a, b) => a.localeCompare(b, 'es'));
      pr.people.forEach(x => { x.profile = pr.label; });
    });

    // Áreas (columna AREA del archivo de calificaciones)
    const hasArea = rows.some(r => 'AREA' in r || 'AREA_EVALUADO' in r);
    for (const p of people.values()) p.area = p.area || 'Sin área';
    const areas = [...new Set([...people.values()].map(p => p.area))].sort((a, b) => a.localeCompare(b, 'es'));
    const cargos = [...new Set([...people.values()].map(p => p.cargo))].sort((a, b) => a.localeCompare(b, 'es'));

    // Catálogo de competencias y su categoría
    const compCatalog = new Map();
    for (const p of people.values())
      for (const c of p.comps.values())
        if (!compCatalog.has(c.name)) compCatalog.set(c.name, c.category);

    return {
      people: [...people.values()].sort((a, b) => a.name.localeCompare(b.name, 'es')),
      peopleById: people, campaigns: [...campaigns], profiles, compCatalog, quality, hasArea, areas, cargos,
      categories: [...new Set(compCatalog.values())].sort((a, b) => a.localeCompare(b, 'es')),
    };
  }

  /* ---------- combinaciones y sugerencia de ponderaciones ---------- */
  function analyzeCombos(model) {
    const combos = new Map();
    for (const p of model.people) {
      if (!combos.has(p.combo)) combos.set(p.combo, { key: p.combo, roles: p.rolesPresent.slice(), people: [] });
      combos.get(p.combo).people.push(p);
    }
    const withBoth = model.people.filter(p => p.gap != null);
    const autoBias = mean(withBoth.map(p => p.gap));
    const withColab = model.people.filter(p => p.colabCount > 0);
    const avgColab = mean(withColab.map(p => p.colabCount));
    const minColab = withColab.length ? Math.min(...withColab.map(p => p.colabCount)) : null;
    const stats = { autoBias, nBoth: withBoth.length, avgColab, minColab, nColab: withColab.length };
    const list = [...combos.values()].sort((a, b) => b.people.length - a.people.length);
    list.forEach(c => { c.suggestion = suggestWeights(c.roles, stats); });
    return { combos: list, stats };
  }

  function suggestWeights(roles, stats) {
    const has = r => roles.includes(r);
    const reasons = [];
    let w = {};
    const fmt = x => (x >= 0 ? '+' : '') + x.toFixed(2).replace('.', ',');
    if (roles.length === 0) return { weights: {}, reasons: ['Sin calificaciones válidas: no se puede medir.'] };
    if (roles.length === 1) {
      w[roles[0]] = 100;
      reasons.push('Solo hay una fuente de evaluación, así que recibe el 100 %. Es una medición parcial.');
      return { weights: w, reasons };
    }
    const lowColab = stats.avgColab != null && stats.avgColab < 3;
    if (has('AUTOEVALUACION') && has('JEFE') && has('COLABORADOR')) {
      w = { AUTOEVALUACION: 20, JEFE: 50, COLABORADOR: 30 };
      reasons.push('Con tres fuentes, el jefe pesa más por conocer los objetivos del cargo; el equipo aporta la visión de liderazgo y la autoevaluación sirve de contraste.');
      if (lowColab) { w = { AUTOEVALUACION: 20, JEFE: 60, COLABORADOR: 20 }; reasons.push('Hay en promedio menos de 3 colaboradores por líder, por eso su peso baja al 20 %.'); }
    } else if (has('AUTOEVALUACION') && has('JEFE')) {
      w = { AUTOEVALUACION: 30, JEFE: 70 };
      reasons.push('El jefe es la fuente principal; la autoevaluación se conserva para medir la percepción propia.');
    } else if (has('AUTOEVALUACION') && has('COLABORADOR')) {
      w = { AUTOEVALUACION: 30, COLABORADOR: 70 };
      reasons.push('Sin jefe, el equipo a cargo es la fuente externa principal.');
      if (lowColab) { w = { AUTOEVALUACION: 40, COLABORADOR: 60 }; reasons.push('Pocos colaboradores por líder: se equilibra un poco más con la autoevaluación.'); }
    } else if (has('JEFE') && has('COLABORADOR')) {
      w = { JEFE: 60, COLABORADOR: 40 };
      reasons.push('Sin autoevaluación, el peso se reparte entre las dos fuentes externas.');
      if (lowColab) { w = { JEFE: 70, COLABORADOR: 30 }; reasons.push('Pocos colaboradores por líder: su peso baja al 30 %.'); }
    }
    // Ajuste por sesgo de autoevaluación observado en los datos
    if (has('AUTOEVALUACION') && stats.autoBias != null) {
      const others = roles.filter(r => r !== 'AUTOEVALUACION');
      if (stats.autoBias > 0.3 && w.AUTOEVALUACION > 10) {
        w.AUTOEVALUACION -= 10; w[others[0]] += 10;
        reasons.push(`En la organización la autoevaluación supera al jefe en ${fmt(stats.autoBias)} puntos en promedio; se reduce su peso 10 puntos.`);
      } else if (stats.autoBias < -0.3) {
        w.AUTOEVALUACION += 10; w[others[0]] -= 10;
        reasons.push(`Las personas se califican ${fmt(stats.autoBias)} puntos por debajo del jefe en promedio; se sube 10 puntos la autoevaluación.`);
      } else {
        reasons.push(`La diferencia promedio Auto − Jefe es ${fmt(stats.autoBias)}: no hay sesgo fuerte que obligue a ajustar.`);
      }
    }
    return { weights: w, reasons };
  }

  /* ---------- cálculo ponderado ---------- */
  function weightedScore(scores, roles, w) {
    let s = 0, tw = 0;
    for (const r of roles) {
      if (scores[r] == null) continue;
      const wr = (w && w[r] != null) ? Number(w[r]) : 0;
      s += wr * scores[r]; tw += wr;
    }
    return tw > 0 ? s / tw : null;
  }

  function applyWeights(model, weightsByCombo) {
    for (const p of model.people) {
      const w = weightsByCombo[p.combo] || {};
      p.weights = w;
      p.final = weightedScore(p.roleScores, p.rolesPresent, w);
      for (const c of p.comps.values()) c.final = weightedScore(c.scores, ROLES, w);
    }
    model.orgComps = aggregateComps(model.people, model.compCatalog);
    return model;
  }

  // Promedios por competencia para un subconjunto de personas
  function aggregateComps(people, compCatalog) {
    const out = [];
    for (const [name, category] of compCatalog) {
      const finals = [], byRole = { AUTOEVALUACION: [], JEFE: [], COLABORADOR: [] }, gaps = [];
      for (const p of people) {
        const c = p.comps.get(name);
        if (!c) continue;
        if (c.final != null) finals.push(c.final);
        for (const r of ROLES) if (c.scores[r] != null) byRole[r].push(c.scores[r]);
        if (c.scores.AUTOEVALUACION != null && c.scores.JEFE != null) gaps.push(c.scores.AUTOEVALUACION - c.scores.JEFE);
      }
      if (!finals.length) continue;
      out.push({
        name, category, final: mean(finals), n: finals.length,
        byRole: { AUTOEVALUACION: mean(byRole.AUTOEVALUACION), JEFE: mean(byRole.JEFE), COLABORADOR: mean(byRole.COLABORADOR) },
        nRole: { AUTOEVALUACION: byRole.AUTOEVALUACION.length, JEFE: byRole.JEFE.length, COLABORADOR: byRole.COLABORADOR.length },
        gap: mean(gaps), nGap: gaps.length,
      });
    }
    return out;
  }

  // Puntajes de una persona limitados a un tipo de competencia (o todos)
  function scoreFor(p, category) {
    if (!category) return { final: p.final, roles: p.roleScores, gap: p.gap };
    const comps = [...p.comps.values()].filter(c => c.category === category);
    const roles = {};
    for (const r of ROLES) roles[r] = mean(comps.map(c => c.scores[r]));
    const gap = roles.AUTOEVALUACION != null && roles.JEFE != null ? roles.AUTOEVALUACION - roles.JEFE : null;
    return { final: mean(comps.map(c => c.final)), roles, gap };
  }

  // Participación: evaluaciones asignadas frente a realizadas
  function participation(people) {
    const out = { assigned: 0, done: 0, complete: 0, partial: 0, empty: 0, byRole: {} };
    for (const r of ROLES) out.byRole[r] = { assigned: 0, done: 0 };
    for (const p of people) for (const e of p.evals.values()) {
      out.assigned++; out.byRole[e.rol].assigned++;
      if (e.answered > 0) { out.done++; out.byRole[e.rol].done++; }
      if (e.state === 'completa') out.complete++;
      else if (e.state === 'parcial') out.partial++;
      else out.empty++;
    }
    return out;
  }

  function gapLabel(gap) {
    if (gap == null) return { key: 'na', text: 'Brecha no calculable' };
    if (gap > GAP_THRESHOLD) return { key: 'over', text: 'Se sobreevalúa' };
    if (gap < -GAP_THRESHOLD) return { key: 'under', text: 'Se subevalúa' };
    return { key: 'ok', text: 'Percepción alineada' };
  }

  const api = { ROLES, ROLE_LABEL, ROLE_SHORT, SCALE_MIN, SCALE_MAX, GAP_THRESHOLD, REQ_SCORES, REQ_FEEDBACK,
    missingColumns, buildModel, analyzeCombos, suggestWeights, applyWeights, aggregateComps, scoreFor, participation, weightedScore, gapLabel, mean };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else global.Engine = api;
})(typeof window !== 'undefined' ? window : globalThis);
