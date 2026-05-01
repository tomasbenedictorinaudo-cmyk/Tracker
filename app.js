/* Cockpit — lightweight project tracker
   Data model is held in `state` and persisted to localStorage as JSON.
   Import/Export gives the same JSON as a file. */

(function () {
  'use strict';

  /* ----------------------------- helpers ----------------------------- */

  const $ = (sel, root = document) => root.querySelector(sel);
  const $$ = (sel, root = document) => Array.from(root.querySelectorAll(sel));
  const uid = (prefix = 'id') => prefix + '_' + Math.random().toString(36).slice(2, 9);
  const dayMs = 86400000;
  // Local-date YYYY-MM-DD (toISOString uses UTC and can shift the day in non-UTC zones)
  const fmtISO = (d) => {
    const y = d.getFullYear();
    const m = String(d.getMonth() + 1).padStart(2, '0');
    const day = String(d.getDate()).padStart(2, '0');
    return `${y}-${m}-${day}`;
  };
  const todayISO = () => fmtISO(new Date());
  const parseDate = (d) => (d ? new Date(d + 'T00:00:00') : null);
  const fmtDate = (d) => {
    if (!d) return '—';
    const dt = typeof d === 'string' ? parseDate(d) : d;
    return dt.toLocaleDateString(undefined, { month: 'short', day: 'numeric' });
  };
  const fmtFull = (d) => {
    if (!d) return '—';
    const dt = typeof d === 'string' ? parseDate(d) : d;
    return dt.toLocaleDateString(undefined, { weekday: 'short', month: 'short', day: 'numeric', year: 'numeric' });
  };
  const dayDiff = (a, b) => Math.round((parseDate(a) - parseDate(b)) / dayMs);
  const clamp = (n, lo, hi) => Math.max(lo, Math.min(hi, n));
  const escapeHTML = (s) =>
    String(s ?? '')
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
  const initials = (name) => {
    if (!name) return '?';
    const parts = name.trim().split(/\s+/);
    return ((parts[0]?.[0] || '') + (parts[1]?.[0] || '')).toUpperCase() || '?';
  };
  // Phase A — fuzzy-score (subsequence + bonus for word-start matches).
  // Returns 0 when query doesn't fuzzy-match haystack; higher score = better.
  function fuzzyScore(query, hay) {
    if (!query) return 1;
    if (!hay) return 0;
    const q = query.toLowerCase();
    const h = hay.toLowerCase();
    let qi = 0, score = 0, lastMatch = -1, run = 0;
    for (let i = 0; i < h.length && qi < q.length; i++) {
      if (h[i] === q[qi]) {
        // word-start bonus
        if (i === 0 || /[\s\-_/.]/.test(h[i - 1])) score += 4;
        else score += 1;
        // consecutive-match bonus
        if (lastMatch === i - 1) { run++; score += run; }
        else run = 0;
        lastMatch = i;
        qi++;
      }
    }
    if (qi < q.length) return 0;
    // Shorter haystacks score slightly higher (preserves intent of full match)
    return score + Math.max(0, 30 - h.length) * 0.1;
  }
  // Phase A — Markdown escape for the upcoming status-report export.
  const mdEscape = (s) => String(s ?? '').replace(/([\\`*_{}\[\]()#+\-.!|>])/g, '\\$1');

  // Expand a meeting record into its concrete dates within a window.
  // Handles oneoff and recurring (day / week / month) — single source of
  // truth so the calendar and timeline never disagree about when a meeting
  // recurs. Caps at m.endDate (if set) and the requested grid window.
  function expandMeetingDates(m, gridStartISO, gridEndISO) {
    if (!m) return [];
    if (m.kind === 'oneoff') {
      return (m.date && m.date >= gridStartISO && m.date <= gridEndISO) ? [m.date] : [];
    }
    if (m.kind !== 'recurring') return [];
    const startISO = m.startDate || todayISO();
    const startD   = parseDate(startISO);
    const gridS    = parseDate(gridStartISO);
    const gridE    = parseDate(gridEndISO);
    const cap      = m.endDate ? parseDate(m.endDate) : null;
    const interval = Math.max(1, m.interval || 1);
    const out = [];
    if (m.recurUnit === 'day') {
      let dt = new Date(startD);
      while (dt <= gridE && (!cap || dt <= cap)) {
        if (dt >= gridS) out.push(fmtISO(dt));
        dt = new Date(dt.getTime() + interval * dayMs);
      }
    } else if (m.recurUnit === 'week') {
      const targetDow = (typeof m.dayOfWeek === 'number') ? m.dayOfWeek : startD.getDay();
      // First occurrence ≥ startD landing on targetDow
      let dt = new Date(startD);
      while (dt.getDay() !== targetDow) dt = new Date(dt.getTime() + dayMs);
      while (dt <= gridE && (!cap || dt <= cap)) {
        if (dt >= gridS) out.push(fmtISO(dt));
        dt = new Date(dt.getTime() + interval * 7 * dayMs);
      }
    } else if (m.recurUnit === 'month') {
      // Step `interval` months at a time on the same day-of-month as
      // startD; clamp to last-day if the target month is shorter (Feb 30).
      const dom = startD.getDate();
      let y = startD.getFullYear();
      let mo = startD.getMonth();
      while (true) {
        const lastDay = new Date(y, mo + 1, 0).getDate();
        const dt = new Date(y, mo, Math.min(dom, lastDay));
        if (dt > gridE) break;
        if (cap && dt > cap) break;
        if (dt >= startD && dt >= gridS) out.push(fmtISO(dt));
        mo += interval;
        while (mo >= 12) { mo -= 12; y += 1; }
      }
    }
    return out;
  }
  function meetingRecurrenceLabel(m) {
    if (m.kind === 'oneoff') return 'Meeting';
    if (m.kind !== 'recurring') return 'Meeting';
    const n = Math.max(1, m.interval || 1);
    const unit = m.recurUnit === 'day' ? 'day' : m.recurUnit === 'month' ? 'month' : 'week';
    if (n === 1) return unit === 'day' ? 'Daily' : unit === 'week' ? 'Weekly' : 'Monthly';
    return `Every ${n} ${unit}s`;
  }

  // Phase G — render tag chips for a record. Tags live on the project; a
  // record stores only ids. Unknown ids (e.g. tag was deleted) are silently
  // skipped to avoid empty chips.
  function renderTagChipsHTML(tagIds, proj) {
    if (!Array.isArray(tagIds) || !tagIds.length) return '';
    const known = (proj?.tags || []).filter((t) => t && t.id);
    const byId = new Map(known.map((t) => [t.id, t]));
    return tagIds.map((id) => {
      const t = byId.get(id);
      if (!t) return '';
      const rgb = t.rgb || '120, 120, 140';
      return `<span class="tag-chip" style="background:rgba(${rgb},.18);color:rgb(${rgb});border:1px solid rgba(${rgb},.40)" title="${escapeHTML(t.name)}">${escapeHTML(t.name)}</span>`;
    }).join('');
  }
  const TAG_PALETTE = [
    '110, 168, 255',  // blue
    '179, 137, 255',  // purple
    '52,  211, 153',  // green
    '251, 191, 36',   // amber
    '248, 113, 113',  // red
    '129, 140, 248',  // indigo
    '236, 72, 153',   // pink
    '244, 114, 182',  // rose
    '20,  184, 166',  // teal
  ];

  function applyTheme(theme) {
    document.documentElement.dataset.theme = theme;
    const btn = document.getElementById('btnTheme');
    if (btn) btn.textContent = theme === 'light' ? '☀' : '☾';
  }

  function toast(msg) {
    const el = $('#toast');
    el.textContent = msg;
    el.hidden = false;
    clearTimeout(el._t);
    el._t = setTimeout(() => (el.hidden = true), 1800);
  }

  /* ------------------------ state & persistence ---------------------- */

  const STORAGE_KEY = 'cockpit.v2';
  const HISTORY_LIMIT = 60;

  const STATUSES = [
    { id: 'todo',      name: 'Not started', dot: 'todo' },
    { id: 'doing',     name: 'In progress', dot: 'doing' },
    { id: 'blocked',   name: 'Blocked',     dot: 'blocked' },
    { id: 'done',      name: 'Done',        dot: 'done' },
    { id: 'cancelled', name: 'Cancelled',   dot: 'cancelled' },
  ];
  // A "closed" status — done or cancelled. Used by capacity / utilisation /
  // late-detection / staleness / KPIs so cancelled actions stop counting as open.
  const isClosedStatus = (s) => s === 'done' || s === 'cancelled';

  // Project components (sub-systems / work packages). Each component has
  // an id, a name, and a color from this palette.
  const COMPONENT_COLORS = [
    { id: 'sky',    name: 'Sky',    rgb: '96,165,250' },
    { id: 'cyan',   name: 'Cyan',   rgb: '34,211,238' },
    { id: 'mint',   name: 'Mint',   rgb: '52,211,153' },
    { id: 'lime',   name: 'Lime',   rgb: '163,230,53' },
    { id: 'amber',  name: 'Amber',  rgb: '251,191,36' },
    { id: 'rose',   name: 'Rose',   rgb: '251,113,133' },
    { id: 'violet', name: 'Violet', rgb: '167,139,250' },
    { id: 'indigo', name: 'Indigo', rgb: '129,140,248' },
    { id: 'slate',  name: 'Slate',  rgb: '148,163,184' },
  ];
  const componentColor = (id) => COMPONENT_COLORS.find((c) => c.id === id) || COMPONENT_COLORS[0];
  const findComponent = (proj, componentId) => (proj.components || []).find((p) => p.id === componentId);

  // Change-request lifecycle states + their colors
  const CR_STATUSES = [
    { id: 'proposed',     label: 'Proposed',     rgb: '148,163,184' }, // slate
    { id: 'under_review', label: 'Under review', rgb: '96,165,250'  }, // blue
    { id: 'approved',     label: 'Approved',     rgb: '74,222,128'  }, // green
    { id: 'rejected',     label: 'Rejected',     rgb: '248,113,133' }, // rose
    { id: 'implemented',  label: 'Implemented',  rgb: '167,139,250' }, // violet
    { id: 'cancelled',    label: 'Cancelled',    rgb: '107,114,128' }, // gray
  ];
  const crStatus = (id) => CR_STATUSES.find((s) => s.id === id) || CR_STATUSES[0];

  // Open-point criticality palette + labels (4 levels, low → critical)
  const CRITICALITY_RGB = {
    low:      '148,163,184', // slate
    med:      '96,165,250',  // blue
    high:     '251,191,36',  // amber
    critical: '248,113,133', // rose
  };
  const CRITICALITY_LABEL = { low: 'Low', med: 'Medium', high: 'High', critical: 'Critical' };
  const CRITICALITY_TO_PRIORITY = { low: 0, med: 1, high: 2, critical: 3 };
  // Action priority levels — same 4-level scale as criticality. Stored as
  // `a.priorityLevel`; defaults to 'med' for legacy actions in normalizeState.
  // Note: `a.priority` (a numeric integer) is kept untouched as the in-column
  // sort key — both fields coexist and serve different purposes.
  const PRIORITY_LEVELS = [
    { id: 'low',      label: 'Low',      rgb: CRITICALITY_RGB.low },
    { id: 'med',      label: 'Medium',   rgb: CRITICALITY_RGB.med },
    { id: 'high',     label: 'High',     rgb: CRITICALITY_RGB.high },
    { id: 'critical', label: 'Critical', rgb: CRITICALITY_RGB.critical },
  ];
  const priorityLevel = (id) => PRIORITY_LEVELS.find((p) => p.id === id) || PRIORITY_LEVELS[1];

  let state = null;
  let undoStack = [];
  let redoStack = [];

  function loadState() {
    try {
      const raw = localStorage.getItem(STORAGE_KEY);
      if (raw) return JSON.parse(raw);
    } catch (e) { /* ignore */ }
    return seedState();
  }
  function saveState() {
    try { localStorage.setItem(STORAGE_KEY, JSON.stringify(state)); }
    catch (e) { /* quota */ }
  }
  function commit(action = 'change') {
    undoStack.push(JSON.stringify(state));
    if (undoStack.length > HISTORY_LIMIT) undoStack.shift();
    redoStack = [];
    saveState();
    render();
  }
  // Phase A helper — stamp signed-edit metadata on a record (action, CR,
  // open point, …) before committing. Optional `editor` lets future team-mode
  // identify the author; defaults to the local user marker.
  function stampEdit(record, editor) {
    if (!record || typeof record !== 'object') return;
    record.__lastEditor = editor || (state.settings.localUser || 'me');
    record.__lastEditAt = new Date().toISOString();
  }
  // Wraps a mutator: runs it, stamps any returned record(s), then commits.
  // Use as `mutate(() => { … return action; }, 'commit-name')`.
  function mutate(fn, commitName) {
    const result = fn();
    if (result) {
      if (Array.isArray(result)) result.forEach(stampEdit);
      else stampEdit(result);
    }
    commit(commitName || 'edit');
    return result;
  }
  function undo() {
    if (!undoStack.length) return;
    redoStack.push(JSON.stringify(state));
    state = JSON.parse(undoStack.pop());
    saveState();
    render();
    toast('Undone');
  }
  function redo() {
    if (!redoStack.length) return;
    undoStack.push(JSON.stringify(state));
    state = JSON.parse(redoStack.pop());
    saveState();
    render();
    toast('Redone');
  }

  function seedState() {
    const today = new Date(); today.setHours(0, 0, 0, 0);
    const d = (off) => fmtISO(new Date(today.getTime() + off * dayMs));

    // Deterministic PRNG so reseeding always gives the same dataset
    let _s = 0xC0FFEE;
    const rnd = () => { _s = (Math.imul(_s, 1103515245) + 12345) & 0x7FFFFFFF; return _s / 0x7FFFFFFF; };
    const ri = (lo, hi) => Math.floor(rnd() * (hi - lo + 1)) + lo;
    const pick = (arr) => arr[Math.floor(rnd() * arr.length)];
    const pickW = (entries) => {
      // entries: [[value, weight], ...]
      const total = entries.reduce((s, e) => s + e[1], 0);
      let r = rnd() * total;
      for (const [v, w] of entries) { if ((r -= w) <= 0) return v; }
      return entries[0][0];
    };
    const chance = (p) => rnd() < p;

    // capacity is in % of FTE (1 FTE = 8h/day × 5 days/week, 212 working days/year)
    const people = [
      { id: 'p_sofia', name: 'Sofia Reyes',     role: 'Project Manager',     capacity: 100, hourlyRate: 140 },
      { id: 'p_marie', name: 'Marie Laurent',   role: 'Systems Engineer',    capacity: 100, hourlyRate: 130 },
      { id: 'p_arjun', name: 'Arjun Patel',     role: 'Avionics Lead',       capacity: 100, hourlyRate: 145 },
      { id: 'p_jonas', name: 'Jonas Becker',    role: 'Mechanical',          capacity:  80, hourlyRate: 120 },
      { id: 'p_kira',  name: 'Kira Nakamura',   role: 'Software Architect',  capacity: 100, hourlyRate: 150 },
      { id: 'p_omar',  name: 'Omar El-Sayed',   role: 'Power Systems',       capacity:  80, hourlyRate: 125 },
      { id: 'p_lena',  name: 'Lena Holmberg',   role: 'Thermal Engineer',    capacity:  80, hourlyRate: 125 },
      { id: 'p_diego', name: 'Diego Ferreira',  role: 'AOCS',                capacity: 100, hourlyRate: 135 },
      { id: 'p_yuki',  name: 'Yuki Tanaka',     role: 'Software Developer',  capacity: 100, hourlyRate: 110 },
      { id: 'p_nadia', name: 'Nadia Rahman',    role: 'Test Engineer',       capacity: 100, hourlyRate: 115 },
    ];

    // Title pools per component
    const TITLES = {
      power:    ['Power budget v{n}', 'Battery cell qualification', 'Solar array sizing', 'PCDU prototype build', 'Charge regulator FW {n}', 'EPS block diagram', 'Power emergency mode logic', 'Battery thermal coupling', 'Solar string analysis', 'Bus undervoltage analysis'],
      aocs:     ['Reaction wheel sizing', 'Sun sensor algorithm', 'AOCS controller v{n}', 'Star tracker integration', 'Magnetic torquer cal', 'Detumbling logic', 'Slew maneuver simulation', 'Pointing budget update', 'Gyro bias study', 'Safe-mode attitude logic'],
      struct:   ['Primary structure FE model', 'Mass budget update {n}', 'Vibration test plan', 'Panel layout v{n}', 'Connector spec review', 'Insert design review', 'Tank mounting study', 'CFRP layup spec', 'Harness routing study', 'Random-vibe coupon test'],
      thermal:  ['Thermal model v{n}', 'Radiator sizing', 'TVAC test plan', 'Heaters specification', 'MLI coverage analysis', 'Thermal interface specs', 'Thermal balance test prep', 'Hot-case analysis', 'Cold-case worst-case'],
      sw:       ['FSW build v{n}', 'Boot loader review', 'Telemetry decoder', 'Command dictionary v{n}', 'OBSW unit tests', 'OBSW integration build', 'Memory scrubbing logic', 'Watchdog logic', 'FDIR matrix update', 'Software ICD v{n}'],
      payload:  ['Optical bench design', 'Detector readout chain', 'Calibration plan', 'Payload electrical interface', 'Filter wheel mechanism', 'Optical alignment procedure', 'Stray-light analysis'],
      avionics: ['CDPU bring-up', 'Avionics block diagram', 'Connector pin mapping', 'EMC pre-test', 'OBC-MEMS interface', 'Star tracker thermal interface', 'Harness FMEA'],
      pm:       ['Project plan v{n}', 'Milestone review prep', 'Budget reforecast', 'Customer review prep', 'WBS update', 'Subcontractor review {n}', 'Risk review session'],
      test:     ['Vibration test prep', 'TVAC chamber booking', 'EMC test plan', 'Acceptance test procedure', 'Test bench bring-up', 'GSE check-out'],
      be:       ['Auth refactor v{n}', 'Telemetry decoder library', 'Scheduler service', 'Database migration {n}', 'API rate-limit middleware', 'Background worker queue', 'Audit log v{n}'],
      fe:       ['Procedure editor MVP', 'Live telemetry view', 'Pass-plan UI v{n}', 'Operator dashboard', 'Theme + design tokens', 'Form validation pass'],
      ops:      ['Operator training plan', 'Runbook v{n}', 'Pager rotation setup', 'On-call drill', 'Incident postmortem template'],
      mech:     ['Airframe trade study', 'Motor mount design', 'Landing-gear stress analysis', 'Prop balancing rig'],
      flight:   ['Flight controller bring-up', 'IMU calibration', 'PID tuning sortie', 'GPS lock-in study'],
    };

    // Notes pool — small chance an action gets one
    const NOTES = [
      'Awaiting connector spec from supplier.',
      'Pending customer review of options A and B.',
      'Preliminary results look promising, need to confirm.',
      'Coupling with thermal still TBD.',
      'Held up on availability of test hardware.',
      'Re-baselined after design freeze meeting.',
      'Cross-check with sister-mission heritage.',
      'Will need a second iteration after PDR.',
    ];

    /* Generate one project's actions and history.
       phases: array of { startOff, endOff, density, statusBias }.
       statusBias: 'past' (mostly done), 'present' (mix), 'future' (mostly todo). */
    function genActions(spec) {
      const out = [];
      let counter = 0;
      const titleIters = {};
      for (const phase of spec.phases) {
        const { startOff, endOff, density, statusBias } = phase;
        const span = endOff - startOff;
        const count = Math.max(1, Math.round(span * density));
        for (let i = 0; i < count; i++) {
          const startOffset = startOff + Math.floor(rnd() * span);
          const dur = ri(2, 14);
          const dueOffset = startOffset + dur;
          // Only consider components offered for this project
          const cmp = pick(spec.components);
          const pool = TITLES[cmp.key] || ['Task {n}'];
          titleIters[cmp.key] = (titleIters[cmp.key] || 0) + 1;
          const title = pick(pool).replace('{n}', String(titleIters[cmp.key]));
          // Owner: prefer cmp.owner, otherwise pick from project owners weighted
          const owner = chance(0.65) && cmp.owner ? cmp.owner : pick(spec.owners);

          // Determine status from bias
          let status;
          if (dueOffset < -2) {
            // Past due — almost always done; sometimes still doing or blocked
            status = pickW([['done', 88], ['blocked', 6], ['doing', 6]]);
          } else if (dueOffset >= -2 && dueOffset <= 14) {
            status = pickW([['doing', 45], ['todo', 25], ['blocked', 12], ['done', 18]]);
          } else {
            status = pickW([['todo', 75], ['doing', 18], ['blocked', 4], ['done', 3]]);
          }
          if (statusBias === 'past' && dueOffset > -2) status = pickW([['done', 60], ['doing', 25], ['todo', 10], ['blocked', 5]]);
          if (statusBias === 'future' && dueOffset < 0) status = pickW([['done', 70], ['doing', 20], ['todo', 5], ['blocked', 5]]);

          const createdOff = startOffset - ri(7, 28);
          const createdISO = d(createdOff);
          const startISO = d(startOffset);
          const dueISO = d(dueOffset);

          // Build history with realistic schedule changes (mostly stable,
          // some slips, some pull-ins).
          const history = [{ at: createdISO, what: `Created with status todo` }];
          // Slip frequency: blocked items often slip, done items rarely (most landed),
          // active items occasionally, future items least.
          const slipChance = status === 'blocked' ? 0.55
                           : status === 'doing'   ? 0.22
                           : status === 'done'    ? 0.12
                           : 0.10;
          let curStart = startOffset, curDue = dueOffset;
          if (chance(slipChance)) {
            const slips = ri(1, status === 'blocked' ? 3 : 2);
            for (let s = 0; s < slips; s++) {
              const slipOnOff = curStart - ri(0, 7); // before original start
              if (slipOnOff > -1 || slipOnOff < createdOff + 2) continue;
              // ~25% of changes are pull-ins (work came in earlier than feared)
              const isPullIn = chance(0.25);
              const delta = isPullIn ? -ri(2, 6) : ri(2, 12);
              const newStartOff = Math.max(createdOff + 1, curStart + delta);
              const newDur = Math.max(2, (curDue - curStart) + (isPullIn ? ri(-2, 0) : ri(-1, 3)));
              const newDueOff = newStartOff + newDur;
              if (newStartOff === curStart && newDueOff === curDue) continue;
              history.push({ at: d(slipOnOff),
                what: `Schedule: ${d(curStart)} → ${d(newStartOff)}…${d(newDueOff)}` });
              curStart = newStartOff; curDue = newDueOff;
            }
          }

          // For done actions, add a "Status: X → done" entry shortly after due
          let updatedISO = createdISO;
          if (status === 'done') {
            const completedOff = curDue + ri(-2, 5);
            const completedISO = d(Math.min(-1, completedOff)); // never in the future
            history.push({ at: completedISO, what: `Status: doing → done` });
            updatedISO = completedISO;
          } else if (status !== 'todo') {
            const transitionOff = Math.max(createdOff + 1, ri(curStart - 7, curStart + 2));
            const transISO = d(Math.min(0, transitionOff));
            history.push({ at: transISO, what: `Status: todo → ${status}` });
            updatedISO = transISO;
          }

          // Use the most-recent forecast as the visible due
          const finalStartISO = d(curStart);
          const finalDueISO = d(curDue);

          // Commitment % — most actions are full-time (100%), some are part-time
          const commitment = pickW([[100, 60], [75, 12], [50, 18], [25, 10]]);

          out.push({
            id: uid('a'),
            title,
            owner,
            startDate: finalStartISO,
            due: finalDueISO,
            status,
            priority: counter++,
            commitment,
            component: cmp.id,
            deliverable: spec.deliverableFor ? spec.deliverableFor(cmp.id, dueOffset) : null,
            milestone: null,
            notes: chance(0.18) ? pick(NOTES) : '',
            createdAt: createdISO,
            updatedAt: updatedISO,
            originatorDate: createdISO,
            history,
          });
        }
      }
      return out;
    }
    const newDuration = (start, oldDur) => start + Math.max(2, oldDur + ri(-1, 4));

    /* ===== Project 1: Orbit-7 Satellite Bus =====
       Started ~370 days ago, currently in PDR phase, CDR ~120 days out. */
    const orbitComps = [
      { id: 'pt_power',   key: 'power',    name: 'Power',     color: 'amber',  owner: 'p_omar' },
      { id: 'pt_aocs',    key: 'aocs',     name: 'AOCS',      color: 'sky',    owner: 'p_diego' },
      { id: 'pt_struct',  key: 'struct',   name: 'Structure', color: 'slate',  owner: 'p_jonas' },
      { id: 'pt_thermal', key: 'thermal',  name: 'Thermal',   color: 'rose',   owner: 'p_lena' },
      { id: 'pt_sw',      key: 'sw',       name: 'Software',  color: 'violet', owner: 'p_kira' },
      { id: 'pt_payload', key: 'payload',  name: 'Payload',   color: 'cyan',   owner: 'p_marie' },
      { id: 'pt_avionics',key: 'avionics', name: 'Avionics',  color: 'indigo', owner: 'p_arjun' },
      { id: 'pt_pm',      key: 'pm',       name: 'PM',        color: 'mint',   owner: 'p_sofia' },
      { id: 'pt_test',    key: 'test',     name: 'Test',      color: 'lime',   owner: 'p_nadia' },
    ];
    const orbitOwners = ['p_sofia','p_marie','p_arjun','p_jonas','p_kira','p_omar','p_lena','p_diego','p_nadia'];
    const orbitDeliverables = [
      { id: 'd_srr',  name: 'SRR data pack',           dueDate: d(-300), status: 'done' },
      { id: 'd_pdr',  name: 'PDR data pack',           dueDate: d(28),   status: 'doing' },
      { id: 'd_cdr',  name: 'CDR data pack',           dueDate: d(120),  status: 'todo' },
      { id: 'd_tvac', name: 'TVAC test report',        dueDate: d(150),  status: 'todo' },
      { id: 'd_aiv',  name: 'AIV plan',                dueDate: d(60),   status: 'todo' },
    ];
    const orbitActions = genActions({
      components: orbitComps, owners: orbitOwners,
      deliverableFor: (cid, off) => {
        if (off < -210) return 'd_srr';
        if (off < 30)   return 'd_pdr';
        if (off < 120)  return 'd_cdr';
        if (off < 165)  return 'd_tvac';
        return null;
      },
      phases: [
        { startOff: -370, endOff: -300, density: 0.30, statusBias: 'past' },   // SRR campaign
        { startOff: -300, endOff: -180, density: 0.45, statusBias: 'past' },   // Phase A → PDR ramp
        { startOff: -180, endOff: -30,  density: 0.55, statusBias: 'past' },   // PDR push
        { startOff: -30,  endOff: 30,   density: 0.55, statusBias: 'present' },// Now
        { startOff: 30,   endOff: 120,  density: 0.45, statusBias: 'future' }, // PDR → CDR
        { startOff: 120,  endOff: 180,  density: 0.30, statusBias: 'future' }, // Post-CDR (planning)
      ],
    });

    /* ===== Project 2: Helios-2 Ground Station Software =====
       Started ~190 days ago, beta in 18 days. */
    const heliosComps = [
      { id: 'cm_be',  key: 'be',  name: 'Backend',   color: 'indigo', owner: 'p_kira' },
      { id: 'cm_fe',  key: 'fe',  name: 'Frontend',  color: 'cyan',   owner: 'p_yuki' },
      { id: 'cm_ops', key: 'ops', name: 'Ops',       color: 'lime',   owner: 'p_sofia' },
    ];
    const heliosOwners = ['p_kira','p_yuki','p_sofia','p_marie'];
    const heliosDeliverables = [
      { id: 'd_alpha', name: 'Internal alpha',      dueDate: d(-90), status: 'done' },
      { id: 'd_beta',  name: 'Beta with ops team',  dueDate: d(18),  status: 'doing' },
      { id: 'd_v1',    name: 'v1.0 release',        dueDate: d(80),  status: 'todo' },
    ];
    const heliosActions = genActions({
      components: heliosComps, owners: heliosOwners,
      deliverableFor: (cid, off) => off < -60 ? 'd_alpha' : off < 25 ? 'd_beta' : 'd_v1',
      phases: [
        { startOff: -190, endOff: -90, density: 0.25, statusBias: 'past' },
        { startOff: -90,  endOff: -10, density: 0.45, statusBias: 'past' },
        { startOff: -10,  endOff: 30,  density: 0.55, statusBias: 'present' },
        { startOff: 30,   endOff: 100, density: 0.40, statusBias: 'future' },
        { startOff: 100,  endOff: 180, density: 0.20, statusBias: 'future' },
      ],
    });

    /* ===== Project 3: Falcon Drone R&D ===== */
    const falconComps = [
      { id: 'cm_avion',  key: 'avionics', name: 'Avionics',   color: 'indigo', owner: 'p_arjun' },
      { id: 'cm_mech',   key: 'mech',     name: 'Mechanical', color: 'slate',  owner: 'p_jonas' },
      { id: 'cm_flight', key: 'flight',   name: 'Flight SW',  color: 'violet', owner: 'p_kira' },
      { id: 'cm_test',   key: 'test',     name: 'Test',       color: 'lime',   owner: 'p_nadia' },
    ];
    const falconOwners = ['p_arjun','p_jonas','p_kira','p_nadia','p_yuki'];
    const falconDeliverables = [
      { id: 'd_proto1', name: 'Prototype P1 hover', dueDate: d(-25), status: 'done' },
      { id: 'd_proto2', name: 'Prototype P2 flight envelope', dueDate: d(45), status: 'doing' },
      { id: 'd_field',  name: 'Field test campaign', dueDate: d(110), status: 'todo' },
    ];
    const falconActions = genActions({
      components: falconComps, owners: falconOwners,
      deliverableFor: (cid, off) => off < -20 ? 'd_proto1' : off < 60 ? 'd_proto2' : 'd_field',
      phases: [
        { startOff: -90, endOff: -25, density: 0.45, statusBias: 'past' },
        { startOff: -25, endOff: 30,  density: 0.45, statusBias: 'present' },
        { startOff: 30,  endOff: 120, density: 0.30, statusBias: 'future' },
      ],
    });

    const proj1 = {
      id: 'pr_orbit', name: 'Orbit-7 Satellite Bus',
      description: 'Mid-class earth-observation platform — phase B, PDR closing in.',
      components: orbitComps.map(({ id, name, color }) => ({ id, name, color })),
      deliverables: orbitDeliverables,
      milestones: [
        { id: 'm_srr', name: 'System Requirements Review', date: d(-310), status: 'done' },
        { id: 'm_pdr', name: 'Preliminary Design Review',  date: d(28),   status: 'todo' },
        { id: 'm_cdr', name: 'Critical Design Review',     date: d(120),  status: 'todo' },
        { id: 'm_trr', name: 'Test Readiness Review',      date: d(170),  status: 'todo' },
      ],
      risks: [
        { id: 'r_supply', kind: 'risk', title: 'Reaction wheel lead time slip',  inherent: { probability: 4, impact: 4 }, residual: { probability: 2, impact: 3 }, mitigation: 'Dual-source supplier engaged.', owner: 'p_arjun' },
        { id: 'r_mass',   kind: 'risk', title: 'Mass margin trending under 5%',  inherent: { probability: 4, impact: 3 }, residual: { probability: 2, impact: 3 }, mitigation: 'Lightweighting study + panel optimisation.', owner: 'p_jonas' },
        { id: 'r_power',  kind: 'risk', title: 'EOL power margin tight',         inherent: { probability: 3, impact: 4 }, residual: { probability: 2, impact: 3 }, mitigation: 'Trade study on cell vendor + duty-cycle review.', owner: 'p_omar' },
        { id: 'r_thermal',kind: 'risk', title: 'Hot-case radiator under-sized',  inherent: { probability: 3, impact: 4 }, residual: { probability: 1, impact: 3 }, mitigation: 'Adding louvres to baseline; thermal balance test scheduled.', owner: 'p_lena' },
        { id: 'r_sw',     kind: 'risk', title: 'FSW timeline at risk',           inherent: { probability: 4, impact: 3 }, residual: { probability: 3, impact: 2 }, mitigation: 'Early integration build, weekly scrum, heritage reuse.', owner: 'p_kira' },
        { id: 'r_test',   kind: 'risk', title: 'TVAC chamber availability',      inherent: { probability: 4, impact: 3 }, residual: { probability: 2, impact: 2 }, mitigation: 'Booked alternate facility on standby.', owner: 'p_nadia' },
        { id: 'o_heritage', kind: 'opportunity', title: 'Reuse Mosaic-3 attitude FSW heritage', inherent: { probability: 3, impact: 3 }, residual: { probability: 4, impact: 4 }, mitigation: 'Negotiate IP transfer with sister mission, save ~6 weeks of FSW work.', owner: 'p_kira' },
        { id: 'o_batch',    kind: 'opportunity', title: 'Volume discount on Li-ion cells',     inherent: { probability: 2, impact: 2 }, residual: { probability: 4, impact: 3 }, mitigation: 'Combine PO with sister mission for 12% unit-price reduction.', owner: 'p_omar' },
        { id: 'o_chamber',  kind: 'opportunity', title: 'Earlier TVAC slot at partner facility', inherent: { probability: 2, impact: 3 }, residual: { probability: 3, impact: 4 }, mitigation: 'Partner facility offered Aug slot — could pull in TVAC by 4 weeks.', owner: 'p_nadia' },
      ],
      decisions: [
        { id: 'dec_bus',    title: 'Down-select to BusFrame v3',   rationale: 'Best mass and thermal envelope after trade study.',    date: d(-220), owner: 'p_sofia' },
        { id: 'dec_rw',     title: 'Reaction wheel: Vendor Bravo', rationale: 'Lifetime + lead time vs Vendor Alpha.',                date: d(-160), owner: 'p_arjun' },
        { id: 'dec_battery',title: 'Li-ion 18650 cell — Vendor C', rationale: 'Heritage in similar LEO mission, 15-yr vendor support.', date: d(-95),  owner: 'p_omar' },
        { id: 'dec_optic',  title: 'Single-aperture optical bench', rationale: 'Mass and integration win vs dual-aperture option.',    date: d(-60),  owner: 'p_marie' },
        { id: 'dec_pdr',    title: 'PDR slipped 2 weeks',           rationale: 'Customer requested additional FDIR work; risk register updated.', date: d(-12), owner: 'p_sofia' },
      ],
      changes: [],
      meetings: [
        { id: 'mtg_standup', kind: 'weekly', title: 'Eng standup',     dayOfWeek: 1, startDate: d(-365), time: '09:30' },
        { id: 'mtg_pmrev',   kind: 'weekly', title: 'PM weekly review', dayOfWeek: 5, startDate: d(-365), time: '15:00' },
        { id: 'mtg_pdrkick', kind: 'oneoff', title: 'PDR pre-meet w/ customer', date: d(14), time: '10:00' },
        { id: 'mtg_pdr',     kind: 'oneoff', title: 'PDR data-pack walkthrough', date: d(28), time: '14:00' },
        { id: 'mtg_riskrev', kind: 'oneoff', title: 'Risk register review',     date: d(7),  time: '11:00' },
      ],
      actions: orbitActions,
    };

    const proj2 = {
      id: 'pr_helios', name: 'Helios-2 Ground Station Software',
      description: 'Operations console rewrite — multi-mission scheduling and live telemetry.',
      components: heliosComps.map(({ id, name, color }) => ({ id, name, color })),
      deliverables: heliosDeliverables,
      milestones: [
        { id: 'm_alpha', name: 'Internal alpha cut',  date: d(-90), status: 'done' },
        { id: 'm_beta',  name: 'Beta with ops team',  date: d(18),  status: 'todo' },
        { id: 'm_ga',    name: 'v1.0 GA',             date: d(80),  status: 'todo' },
      ],
      risks: [
        { id: 'r_h_perf', kind: 'risk', title: 'Telemetry decoder throughput',  probability: 3, impact: 3, mitigation: 'Profile hot path, add backpressure.', owner: 'p_kira' },
        { id: 'r_h_ux',   kind: 'risk', title: 'Procedure editor UX scope',     probability: 3, impact: 2, mitigation: 'Two design rounds with ops users.',     owner: 'p_yuki' },
        { id: 'o_h_oss',  kind: 'opportunity', title: 'Open-source the telemetry decoder', probability: 3, impact: 3, mitigation: 'Adoption could drive contributions and hire pipeline.', owner: 'p_sofia' },
      ],
      decisions: [
        { id: 'dec_h_db',  title: 'Postgres for time-series store', rationale: 'Operational simplicity vs InfluxDB; volume tractable.', date: d(-130), owner: 'p_kira' },
        { id: 'dec_h_auth',title: 'OIDC via central IdP',           rationale: 'Aligns with org-wide SSO programme.',                  date: d(-70),  owner: 'p_sofia' },
      ],
      changes: [],
      actions: heliosActions,
    };

    const proj3 = {
      id: 'pr_falcon', name: 'Falcon Drone R&D',
      description: 'Rapid prototype platform for autonomous flight experiments.',
      components: falconComps.map(({ id, name, color }) => ({ id, name, color })),
      deliverables: falconDeliverables,
      milestones: [
        { id: 'm_p1', name: 'P1 hover demo', date: d(-25), status: 'done' },
        { id: 'm_p2', name: 'P2 flight envelope', date: d(45), status: 'todo' },
        { id: 'm_f1', name: 'Field test campaign 1', date: d(110), status: 'todo' },
      ],
      risks: [
        { id: 'r_f_battery', kind: 'risk', title: 'Battery thermal runaway during fast charge', probability: 2, impact: 5, mitigation: 'Cell-level temperature monitoring; conservative charge profile.', owner: 'p_omar' },
        { id: 'r_f_field',   kind: 'risk', title: 'Outdoor test weather window',                probability: 3, impact: 2, mitigation: 'Two backup test windows scheduled.', owner: 'p_nadia' },
        { id: 'o_f_grant',   kind: 'opportunity', title: 'R&D grant for autonomous swarm tests', probability: 2, impact: 4, mitigation: 'Eligibility confirmed; submission window opens in 6 weeks.', owner: 'p_sofia' },
      ],
      decisions: [
        { id: 'dec_f_motor', title: 'Brushless motor: T-Motor MN5008', rationale: 'Best thrust/weight at target battery voltage.',  date: d(-70), owner: 'p_jonas' },
      ],
      changes: [],
      actions: falconActions,
    };

    return {
      people,
      projects: [proj1, proj2, proj3],
      currentProjectId: proj1.id,
      currentView: 'board',
      settings: { theme: 'dark', holidayCountries: [] },
    };
  }

  function emptyState() {
    const proj = {
      id: uid('p'),
      name: 'Untitled project',
      description: '',
      actions: [],
      deliverables: [],
      milestones: [],
      risks: [],
      decisions: [],
      changes: [],
      components: [],
      meetings: [],
      openPoints: [],
      links: [],
      costCenters: [],
      archive: [],
      notes: '',
    };
    return {
      people: [],
      projects: [proj],
      currentProjectId: proj.id,
      currentView: 'board',
      settings: { theme: state?.settings?.theme || 'dark', holidayCountries: [] },
      budgets: {},
    };
  }

  /* ----------------------- selectors / helpers ----------------------- */

  // Returns the active project. If state.currentProjectId === '__all__',
  // returns a synthetic project that aggregates references (not copies) from
  // every project — mutations to actions/risks/etc. propagate to the real
  // source projects. Adding/deleting collection members in this mode is
  // disallowed (curProjectIsMerged() returns true).
  function curProject() {
    if (state.currentProjectId === '__all__') return mergedProject();
    return state.projects.find((p) => p.id === state.currentProjectId) || state.projects[0];
  }
  function curProjectIsMerged() {
    return state.currentProjectId === '__all__';
  }
  // Returns the projects that should be considered "in scope" for cross-cutting
  // engineering views (budgets, charts) — every project in __all__ mode, or just
  // the selected project otherwise.
  function projectsInScope() {
    if (state.currentProjectId === '__all__') return state.projects;
    const cur = state.projects.find((p) => p.id === state.currentProjectId);
    return cur ? [cur] : state.projects;
  }
  function mergedProject() {
    const flat = (key) => state.projects.flatMap((p) => p[key] || []);
    return {
      id: '__all__',
      name: 'All projects',
      description: 'Cross-project view (read-only for adds/deletes).',
      actions: flat('actions'),
      deliverables: flat('deliverables'),
      milestones: flat('milestones'),
      risks: flat('risks'),
      decisions: flat('decisions'),
      components: flat('components'),
      meetings: flat('meetings'),
      links: flat('links'),
      linkFolders: flat('linkFolders'),
      openPoints: flat('openPoints'),
      changes: flat('changes'),
      costCenters: flat('costCenters'),
      tags: flat('tags'),
    };
  }
  // Find the source project for an action by id (works in merged mode too)
  function projectOfAction(actionId) {
    return state.projects.find((p) => (p.actions || []).some((a) => a.id === actionId));
  }
  function personName(id) {
    return state.people.find((p) => p.id === id)?.name || '—';
  }
  function statusOfDue(dueISO, status) {
    if (!dueISO) return 'ok';
    if (status === 'done') return 'ok';
    const diff = dayDiff(dueISO, todayISO());
    if (diff < 0) return 'late';
    if (diff <= 3) return 'soon';
    return 'ok';
  }
  // Apply a set of topbar filters and optionally navigate to a different view.
  // Pass undefined to leave a filter unchanged; pass '' to clear it.
  function applyTopbarFilter({ status, due, component, owner, search, clearAll, view } = {}) {
    if (clearAll) {
      $('#search').value = '';
      $('#filterStatus').value = '';
      $('#filterDue').value = '';
      $('#filterComponent').value = '';
      $('#filterOwner').value = '';
    } else {
      if (status !== undefined) $('#filterStatus').value = status;
      if (due !== undefined) $('#filterDue').value = due;
      if (component !== undefined) $('#filterComponent').value = component;
      if (owner !== undefined) $('#filterOwner').value = owner;
      if (search !== undefined) $('#search').value = search;
    }
    if (view) {
      state.currentView = view;
      saveState();
      render();
    } else {
      // re-render the current view
      render();
    }
  }

  function actionMatchesFilters(a) {
    if (a.deletedAt) return false; // archived items are hidden by default
    const fOwner = $('#filterOwner').value;
    const fComp = $('#filterComponent').value;
    const fStatus = $('#filterStatus').value;
    const fDue = $('#filterDue').value;
    const q = $('#search').value.trim().toLowerCase();
    if (fOwner && a.owner !== fOwner) return false;
    if (fComp) {
      if (fComp === '__none__' && a.component) return false;
      if (fComp !== '__none__' && a.component !== fComp) return false;
    }
    if (fStatus && a.status !== fStatus) return false;
    if (fDue) {
      const today = todayISO();
      const diff = a.due ? dayDiff(a.due, today) : null;
      if (fDue === 'late' && !(a.due && a.status !== 'done' && diff < 0)) return false;
      if (fDue === 'week' && !(a.due && diff !== null && diff >= 0 && diff <= 7)) return false;
      if (fDue === 'month' && !(a.due && diff !== null && diff >= 0 && diff <= 30)) return false;
    }
    if (q) {
      const hay = (a.title + ' ' + personName(a.owner) + ' ' + (a.notes || '')).toLowerCase();
      if (!hay.includes(q)) return false;
    }
    return true;
  }

  /* --------------------------- KPI engine ---------------------------- */

  function kpis() {
    const proj = curProject();
    const acts = (proj.actions || []).filter((a) => !a.deletedAt);
    const today = todayISO();
    const total = acts.length;
    const done = acts.filter((a) => a.status === 'done').length;
    const blocked = acts.filter((a) => a.status === 'blocked').length;
    const doing = acts.filter((a) => a.status === 'doing').length;
    const late = acts.filter((a) => a.due && !isClosedStatus(a.status) && dayDiff(a.due, today) < 0).length;
    const upcoming = acts.filter((a) => a.due && !isClosedStatus(a.status) && dayDiff(a.due, today) >= 0 && dayDiff(a.due, today) <= 7).length;
    const completionRate = total ? Math.round((done / total) * 100) : 0;
    const lateRate = total ? Math.round((late / total) * 100) : 0;
    const blockedRatio = total ? Math.round((blocked / total) * 100) : 0;
    // Throughput: items completed in last 14 days (using updatedAt for done)
    const since = fmtISO(new Date(Date.now() - 14 * dayMs));
    const throughput = acts.filter((a) => a.status === 'done' && a.updatedAt >= since).length;

    // Workload by person — closed actions (done or cancelled) don't load capacity
    const workload = state.people.map((p) => {
      const open = acts.filter((a) => a.owner === p.id && !isClosedStatus(a.status)).length;
      return { id: p.id, name: p.name, open, capacity: p.capacity || 5 };
    });

    return { total, done, doing, blocked, late, upcoming, completionRate, lateRate, blockedRatio, throughput, workload };
  }

  /* ---------------------------- rendering ---------------------------- */

  function render() {
    renderTopbar();
    renderSidebar();
    renderView();
    refreshNoteChips();
    refreshTodoFab();
    refreshBell();
  }

  function renderTopbar() {
    const sel = $('#projectSelect');
    sel.innerHTML =
      `<option value="__all__" ${state.currentProjectId === '__all__' ? 'selected' : ''}>📚 All projects</option>` +
      `<option disabled>──────────</option>` +
      state.projects
        .map((p) => `<option value="${p.id}" ${p.id === state.currentProjectId ? 'selected' : ''}>${escapeHTML(p.name)}</option>`)
        .join('');

    const fOwner = $('#filterOwner');
    const cur = fOwner.value;
    fOwner.innerHTML = `<option value="">All owners</option>` +
      state.people.map((p) => `<option value="${p.id}" ${p.id === cur ? 'selected' : ''}>${escapeHTML(p.name)}</option>`).join('');

    const fComp = $('#filterComponent');
    const curC = fComp.value;
    const proj = curProject();
    fComp.innerHTML = `<option value="">All components</option>` +
      `<option value="__none__" ${curC === '__none__' ? 'selected' : ''}>— Unassigned —</option>` +
      ((proj && proj.components) || []).map((pt) => `<option value="${pt.id}" ${pt.id === curC ? 'selected' : ''}>${escapeHTML(pt.name)}</option>`).join('');
  }

  function renderSidebar() {
    $$('.nav-item').forEach((b) => {
      b.classList.toggle('active', b.dataset.view === state.currentView);
    });
  }

  function renderView() {
    const main = $('#main');
    main.innerHTML = '';
    const view = state.currentView;
    const fns = {
      portfolio: renderPortfolio,
      people: renderPeople,
      board: renderBoard,
      register: renderRegister,
      openpoints: renderOpenPoints,
      timeline: renderTimeline,
      dashboard: renderDashboard,
      // Charts is merged into Dashboard — route any stale 'charts' view
      // (saved before the merge or arrived via palette) to the combined view.
      charts: renderDashboard,
      review: renderReview,
      archive: renderArchive,
      components: renderComponents,
      budgets: renderBudgets,
      deliverables: renderDeliverables,
      milestones: renderMilestones,
      risks: renderRisks,
      decisions: renderDecisions,
      changes: renderChangeRequests,
      links: renderLinks,
      inbox: renderInbox,
      calendar: renderCalendar,
      reports: renderReports,
    };
    (fns[view] || renderBoard)(main);
  }

  /* ----------------------------- Board ------------------------------- */

  function renderBoard(root) {
    const proj = curProject();
    const view = document.createElement('div');
    view.className = 'view board-view';

    const head = document.createElement('div');
    head.className = 'page-head';
    const liveActions = (proj.actions || []).filter((a) => !a.deletedAt);
    const archivedCount = (proj.actions || []).filter((a) => a.deletedAt).length;
    head.innerHTML = `
      <div>
        <div class="page-title">${escapeHTML(proj.name)}</div>
        <div class="page-sub">${liveActions.length} actions • ${proj.deliverables?.length || 0} deliverables • ${proj.milestones?.length || 0} milestones</div>
      </div>
      <div class="page-actions">
        <button class="ghost" id="btnOpenArchive" title="Open the Archive view">⌫ Archive${archivedCount ? ` <span class="badge-count">${archivedCount}</span>` : ''}</button>
        <button class="ghost" id="btnAddAction">+ Action</button>
      </div>`;
    view.appendChild(head);

    // Drop zones for quick mark-done / archive — visible only while a card is dragging
    const zones = document.createElement('div');
    zones.className = 'board-quick-actions';
    zones.innerHTML = `
      <div class="bqa-zone bqa-done"   data-bqa="done"><span class="bqa-icon">✓</span><span>Drop here to mark done</span></div>
      <div class="bqa-zone bqa-delete" data-bqa="delete"><span class="bqa-icon">🗑</span><span>Drop here to archive</span></div>`;
    view.appendChild(zones);

    const board = document.createElement('div');
    board.className = 'board';
    STATUSES.forEach((s) => {
      const items = (proj.actions || [])
        .filter((a) => a.status === s.id && actionMatchesFilters(a))
        .sort((a, b) => (a.priority || 0) - (b.priority || 0));
      const col = document.createElement('div');
      col.className = 'column';
      col.dataset.status = s.id;
      const clearBtn = (s.id === 'done' && items.length > 0)
        ? `<button class="col-clear" title="Archive all ${items.length} done action${items.length === 1 ? '' : 's'}">⌫ Clear</button>`
        : '';
      col.innerHTML = `
        <div class="col-head">
          <span class="col-dot ${s.dot}"></span>
          <span class="col-name">${s.name}</span>
          <span class="col-count">${items.length}</span>
          ${clearBtn}
        </div>
        <div class="col-body" data-status="${s.id}"></div>`;
      const body = $('.col-body', col);
      items.forEach((a) => body.appendChild(makeCard(a)));
      if (!items.length) {
        const empty = document.createElement('div');
        empty.className = 'empty';
        empty.textContent = 'Drop actions here';
        body.appendChild(empty);
      }
      board.appendChild(col);
    });
    view.appendChild(board);
    root.appendChild(view);

    $('#btnAddAction').addEventListener('click', () => openQuickAdd('action'));
    $('#btnOpenArchive').addEventListener('click', () => {
      state.currentView = 'archive';
      saveState(); render();
    });

    // Clear-Done bulk action — archives only the done cards CURRENTLY DISPLAYED
    // (passing the topbar filters / search). Filtered-out done items are kept.
    const clearBtn = view.querySelector('.col-clear');
    if (clearBtn) {
      clearBtn.addEventListener('click', () => {
        const targets = curProjectIsMerged()
          ? state.projects.flatMap((p) => p.actions || [])
          : (proj.actions || []);
        const toArchive = targets.filter((a) => a.status === 'done' && actionMatchesFilters(a));
        if (!toArchive.length) return;
        const filtersActive = !!($('#search').value || $('#filterOwner').value || $('#filterComponent').value || $('#filterStatus').value || $('#filterDue').value);
        const scope = filtersActive ? ' (matching current filters)' : '';
        if (!confirm(`Archive ${toArchive.length} done action${toArchive.length === 1 ? '' : 's'}${scope}? They can be restored from the Archive view.`)) return;
        const today = todayISO();
        toArchive.forEach((a) => {
          a.deletedAt = today;
          a.history.push({ at: today, what: 'Archived (bulk-clear from Done)' });
          a.updatedAt = today;
        });
        commit('clear-done');
        toast(`${toArchive.length} archived`);
      });
    }

    // Card drag is handled via custom mouse events in startCardDrag(),
    // including drops onto the .bqa-zone pills and column bodies.
  }

  function makeCard(a) {
    const due = a.due;
    const dueClass = statusOfDue(due, a.status);
    const card = document.createElement('div');
    card.className = `card ${a.status === 'doing' ? 'doing' : ''} ${a.status === 'cancelled' ? 'cancelled' : ''} ${dueClass}`;
    card.dataset.id = a.id;
    const owner = state.people.find((p) => p.id === a.owner);
    const component = findComponent(curProject(), a.component);
    if (component) {
      const c = componentColor(component.color);
      card.classList.add('has-component');
      card.style.setProperty('--cmp-bg',       `rgba(${c.rgb},.10)`);
      card.style.setProperty('--cmp-bg-hover', `rgba(${c.rgb},.18)`);
      card.style.setProperty('--cmp-border',   `rgba(${c.rgb},.40)`);
      card.style.setProperty('--cmp-chip-bg',  `rgba(${c.rgb},.20)`);
      card.style.setProperty('--cmp-chip-fg',  `rgb(${c.rgb})`);
    }
    const lvl = priorityLevel(a.priorityLevel);
    // Only render a chip when priority is non-default (Medium) — keeps cards
    // calm and lets High/Critical visually pop.
    const showPriorityChip = a.priorityLevel && a.priorityLevel !== 'med';
    const tagChips = renderTagChipsHTML(a.tags, curProject(), a.comments);
    card.innerHTML = `
      <button class="row-overflow" data-action="overflow" title="More actions" aria-label="More actions">⋯</button>
      <div class="card-top-row">
        ${component ? `<div class="component-chip">${escapeHTML(component.name)}</div>` : ''}
        ${showPriorityChip ? `<span class="prio-chip prio-${lvl.id}" title="Priority: ${lvl.label}" style="background:rgba(${lvl.rgb},.18);color:rgb(${lvl.rgb});border:1px solid rgb(${lvl.rgb})">${lvl.label}</span>` : ''}
        ${tagChips}
      </div>
      <div class="card-title">${escapeHTML(a.title)}</div>
      <div class="card-meta">
        <span class="avatar" title="${escapeHTML(owner?.name || 'Unassigned')}">${initials(owner?.name)}</span>
        <span class="due ${dueClass}">${due ? fmtDate(due) : 'no date'}</span>
        ${a.notes ? '<span class="tag">note</span>' : ''}
        ${a.description ? '<span class="tag has-desc" title="Has a description — hover to read">≡</span>' : ''}
        ${(a.comments && a.comments.length) ? `<span class="tag has-comments" title="${a.comments.length} comment${a.comments.length === 1 ? '' : 's'}">💬 ${a.comments.length}</span>` : ''}
      </div>`;
    card.addEventListener('click', (e) => {
      if (e.target.closest('[data-action="overflow"]')) {
        e.stopPropagation();
        const r = e.target.getBoundingClientRect();
        showContextMenu(r.left, r.bottom + 4, actionContextItems(a));
        return;
      }
      if (card._suppressClick) { card._suppressClick = false; return; }
      openDrawer(a.id);
    });
    card.addEventListener('mousedown', (e) => {
      if (e.button !== 0) return; // left mouse only
      // Don't initiate drag from the ⋯ overflow button — startCardDrag
      // hides the card via display:none on mousedown, which would tear
      // the button out from under the click event before it fires.
      if (e.target.closest('[data-action="overflow"]')) return;
      startCardDrag(e, card, a);
    });
    card.addEventListener('contextmenu', (e) => {
      e.preventDefault();
      showContextMenu(e.clientX, e.clientY, actionContextItems(a));
    });
    return card;
  }

  // Custom mouse-event drag for Kanban cards. The ghost is created
  // immediately on mousedown so there's no perceived "two-step" feel
  // between clicking and dragging — the card visually lifts at once.
  function startCardDrag(eDown, card, action) {
    const view = card.closest('.board-view');
    const startX = eDown.clientX, startY = eDown.clientY;
    const rect = card.getBoundingClientRect();
    let movedEnough = false;
    let ghost = null;
    let placeholder = null;      // live preview of where the card will land
    let dropTarget = null;
    let lastEvt = null;
    let rafPending = false;
    let prevHighlight = null;
    const CLICK_THRESHOLD = 3;

    eDown.preventDefault?.();

    // Create the ghost immediately so the card "lifts" on mousedown
    ghost = card.cloneNode(true);
    ghost.classList.add('drag-ghost');
    Object.assign(ghost.style, {
      position: 'fixed',
      left: rect.left + 'px',
      top: rect.top + 'px',
      width: rect.width + 'px',
      margin: '0',
      zIndex: '9999',
      willChange: 'transform',
      transform: 'translate3d(0,0,0)',
      opacity: '0.92',
      boxShadow: '0 12px 30px rgba(0,0,0,.45)',
      transition: 'none',
    });
    document.body.appendChild(ghost);
    card.classList.add('dragging');
    // Remove the original card from layout while dragging — the placeholder
    // takes its conceptual place. This lets same-column reorder work and
    // makes layout calculations stable.
    const prevDisplay = card.style.display;
    card.style.display = 'none';
    if (view) view.classList.add('dragging-active');
    document.body.classList.add('is-card-dragging');

    function ensurePlaceholder() {
      if (placeholder) return placeholder;
      placeholder = card.cloneNode(true);
      placeholder.classList.add('drop-placeholder');
      placeholder.classList.remove('dragging');
      placeholder.style.display = '';
      placeholder.style.background = '';
      placeholder.style.borderColor = '';
      delete placeholder.dataset.id; // distinguishable from cards by absence of id
      return placeholder;
    }
    function detachPlaceholder() {
      if (placeholder?.parentElement) placeholder.parentElement.removeChild(placeholder);
    }

    function applyFrame() {
      rafPending = false;
      if (!ghost || !lastEvt) return;
      const em = lastEvt;
      const dx = em.clientX - startX;
      const dy = em.clientY - startY;
      ghost.style.transform = `translate3d(${dx}px, ${dy}px, 0) rotate(2deg)`;
      // Hit-test under the cursor (ghost has pointer-events: none via CSS)
      const under = document.elementFromPoint(em.clientX, em.clientY);
      const zone = under?.closest('.bqa-zone');
      const col  = under?.closest('.col-body');
      const newTarget = zone || col || null;
      if (newTarget !== prevHighlight) {
        if (prevHighlight) prevHighlight.classList.remove('drop-target', 'over');
        prevHighlight = newTarget;
        if (zone) zone.classList.add('over');
        else if (col) col.classList.add('drop-target');
      }
      if (zone) {
        detachPlaceholder();
        dropTarget = { type: 'zone', el: zone, action: zone.dataset.bqa };
      } else if (col) {
        const ph = ensurePlaceholder();
        // Detach placeholder first so its own height doesn't bias the layout
        detachPlaceholder();
        // Visible cards in this column (the original is display:none, so naturally absent)
        const cards = [...col.children].filter((el) => el.classList?.contains('card'));
        const after = cards.find((el) => {
          const r = el.getBoundingClientRect();
          return em.clientY < r.top + r.height / 2;
        });
        if (after) col.insertBefore(ph, after);
        else col.appendChild(ph);
        dropTarget = { type: 'column', el: col, status: col.dataset.status };
      } else {
        detachPlaceholder();
        dropTarget = null;
      }
    }

    function onMove(em) {
      lastEvt = em;
      if (!movedEnough && Math.hypot(em.clientX - startX, em.clientY - startY) >= CLICK_THRESHOLD) {
        movedEnough = true;
      }
      if (!rafPending) {
        rafPending = true;
        requestAnimationFrame(applyFrame);
      }
    }

    function onUp() {
      window.removeEventListener('mousemove', onMove);
      window.removeEventListener('mouseup', onUp);
      if (rafPending && lastEvt) applyFrame();
      if (movedEnough && !dropTarget && lastEvt) {
        const u = document.elementFromPoint(lastEvt.clientX, lastEvt.clientY);
        const z = u?.closest('.bqa-zone');
        const c = u?.closest('.col-body');
        if (z) dropTarget = { type: 'zone', el: z, action: z.dataset.bqa };
        else if (c) dropTarget = { type: 'column', el: c, status: c.dataset.status };
      }
      // Snapshot placeholder ordinal BEFORE we tear down DOM. The placeholder
      // sits exactly where the user wants the card; everything else is in
      // visible order. (Original card is display:none so naturally excluded.)
      let placeholderOrder = null;
      if (dropTarget?.type === 'column' && placeholder?.parentElement === dropTarget.el) {
        const live = [...dropTarget.el.children].filter((el) => el.classList?.contains('card'));
        placeholderOrder = live.map((el) => el === placeholder ? action.id : el.dataset.id);
      }
      document.querySelectorAll('.col-body.drop-target, .bqa-zone.over').forEach((el) => {
        el.classList.remove('drop-target', 'over');
      });
      detachPlaceholder();
      document.body.classList.remove('is-card-dragging');
      if (view) view.classList.remove('dragging-active');
      card.classList.remove('dragging');
      card.style.display = prevDisplay; // restore (render() will rebuild anyway on commit)
      if (ghost) { ghost.remove(); ghost = null; }
      if (!movedEnough) return;
      card._suppressClick = true;
      if (!dropTarget) return;
      // Hand the precomputed order to the column-drop branch
      dropTarget.order = placeholderOrder;
      const sourceProj = projectOfAction(action.id);
      const a = sourceProj?.actions.find((x) => x.id === action.id);
      if (!a) return;
      if (dropTarget.type === 'zone') {
        if (dropTarget.action === 'done') {
          if (a.status !== 'done') {
            // Not done yet → mark done (stays on the board)
            a.history.push({ at: todayISO(), what: `Status: ${a.status} → done` });
            a.status = 'done';
            a.updatedAt = todayISO();
            commit('done');
            toast('Marked done');
          } else {
            // Already done → archive it (leaves the board, recoverable from Archive)
            a.deletedAt = todayISO();
            a.history.push({ at: todayISO(), what: 'Archived from Done' });
            a.updatedAt = todayISO();
            commit('archive-done');
            toast('Archived — restore from Archive view');
          }
        } else {
          a.deletedAt = todayISO();
          a.history.push({ at: todayISO(), what: 'Moved to Archive' });
          a.updatedAt = todayISO();
          commit('archive');
          toast('Moved to Archive — restore from Archive view');
        }
        return;
      }
      // type === 'column'
      const oldStatus = a.status;
      const newStatus = dropTarget.status;
      a.status = newStatus;
      a.updatedAt = todayISO();
      if (oldStatus !== newStatus) {
        a.history.push({ at: todayISO(), what: `Status: ${oldStatus} → ${newStatus}` });
      }
      // Use the order captured from the placeholder position; fall back to
      // the cursor-based insertion if the placeholder was never placed.
      let filtered;
      if (dropTarget.order && dropTarget.order.length) {
        filtered = dropTarget.order;
      } else {
        const cards = [...dropTarget.el.querySelectorAll('.card:not(.dragging)')];
        const dropY = lastEvt?.clientY ?? -Infinity;
        const after = cards.find((el) => {
          const r = el.getBoundingClientRect();
          return dropY < r.top + r.height / 2;
        });
        const ids = cards.map((c) => c.dataset.id);
        filtered = ids.filter((id) => id !== a.id);
        const insertIdx = after ? filtered.indexOf(after.dataset.id) : filtered.length;
        filtered.splice(insertIdx === -1 ? filtered.length : insertIdx, 0, a.id);
      }
      filtered.forEach((id, i) => {
        const aa = sourceProj.actions.find((x) => x.id === id);
        if (aa) aa.priority = i;
      });
      commit('move');
      toast(oldStatus === newStatus ? 'Reordered' : `Moved to ${newStatus}`);
    }

    window.addEventListener('mousemove', onMove);
    window.addEventListener('mouseup', onUp);
  }

  // Shared right-click menu for an action — usable on board cards and gantt bars
  function actionContextItems(a) {
    const setStatus = (st) => () => {
      if (a.status === st) return;
      a.history.push({ at: todayISO(), what: `Status: ${a.status} → ${st}` });
      a.status = st;
      a.updatedAt = todayISO();
      commit('status');
    };
    return [
      { icon: '✎', label: 'Edit details…', onClick: () => openDrawer(a.id) },
      { icon: '✎', label: a.notes ? 'Edit note…' : 'Add note…',
        onClick: () => openNoteEditor(a.id) },
      { icon: '✎', label: a.description ? 'Edit description…' : 'Add description…',
        onClick: () => openDescriptionEditor(a.id) },
      { divider: true },
      { icon: '○', label: 'Mark not started', onClick: setStatus('todo') },
      { icon: '◐', label: 'Mark in progress', onClick: setStatus('doing') },
      { icon: '⨯', label: 'Mark blocked',     onClick: setStatus('blocked') },
      { icon: '✓', label: 'Mark done',        onClick: setStatus('done') },
      { divider: true },
      { icon: '×', label: 'Move to Archive', danger: true, onClick: () => {
        if (!confirm(`Move "${a.title}" to Archive?`)) return;
        a.deletedAt = todayISO();
        a.history.push({ at: todayISO(), what: 'Moved to Archive' });
        a.updatedAt = todayISO();
        commit('delete');
        toast('Moved to Archive');
      }},
    ];
  }

  // Quick plain-text note editor — opens a tiny modal, stores in a.notes
  function openNoteEditor(actionId) {
    const a = state.projects.flatMap((p) => p.actions || []).find((x) => x.id === actionId);
    if (!a) return;
    const overlay = document.createElement('div');
    overlay.className = 'overlay desc-overlay';
    overlay.innerHTML = `
      <div class="desc-modal" style="width:520px;">
        <div class="desc-head">
          <div class="desc-title">${escapeHTML(a.title)} — note</div>
          <button class="icon-btn" id="noteClose" title="Close">×</button>
        </div>
        <div style="padding:14px 16px;">
          <textarea id="noteText" placeholder="Plain-text note (use the description for rich text)" style="width:100%; min-height:140px; resize:vertical; background:var(--bg-2); border:1px solid var(--line); border-radius:var(--radius-sm); padding:10px; color:var(--text); font: inherit; line-height:1.5; outline:none;">${escapeHTML(a.notes || '')}</textarea>
        </div>
        <div class="desc-foot">
          <button class="ghost" id="noteCancel">Cancel</button>
          <button class="primary" id="noteSave">Save</button>
        </div>
      </div>`;
    document.body.appendChild(overlay);
    setTimeout(() => document.getElementById('noteText').focus(), 30);
    const close = () => overlay.remove();
    overlay.querySelector('#noteClose').addEventListener('click', close);
    overlay.querySelector('#noteCancel').addEventListener('click', close);
    overlay.querySelector('#noteSave').addEventListener('click', () => {
      a.notes = document.getElementById('noteText').value.trim();
      a.updatedAt = todayISO();
      commit('note');
      close();
      toast('Note saved');
    });
    // Modal closes ONLY via the explicit × button (and Cancel where present).
    // Backdrop click and Escape are intentionally NOT bound — losing
    // in-progress edits to a stray click outside the modal is too easy.
  }

  // Rich-text description editor — opens a small modal, stores HTML in a.description
  function openDescriptionEditor(actionId) {
    const a = state.projects.flatMap((p) => p.actions || []).find((x) => x.id === actionId);
    if (!a) return;
    const overlay = document.createElement('div');
    overlay.className = 'overlay desc-overlay';
    overlay.innerHTML = `
      <div class="desc-modal">
        <div class="desc-head">
          <div class="desc-title">${escapeHTML(a.title)} — description</div>
          <button class="icon-btn" id="descClose" title="Close">×</button>
        </div>
        <div class="notes-toolbar desc-toolbar">
          <button data-cmd="bold" title="Bold"><b>B</b></button>
          <button data-cmd="italic" title="Italic"><i>I</i></button>
          <button data-cmd="underline" title="Underline"><u>U</u></button>
          <span class="sep"></span>
          <button data-cmd="formatBlock" data-arg="<h3>" title="Heading">H</button>
          <button data-cmd="formatBlock" data-arg="<p>" title="Paragraph">¶</button>
          <button data-cmd="insertUnorderedList" title="Bullet list">• list</button>
          <button data-cmd="insertOrderedList" title="Numbered list">1.</button>
        </div>
        <div class="desc-body" contenteditable="true" spellcheck="true">${a.description || '<p></p>'}</div>
        <div class="desc-foot">
          <button class="ghost" id="descCancel">Cancel</button>
          <button class="primary" id="descSave">Save</button>
          ${a.description ? '<button class="ghost desc-clear" id="descClear">Remove description</button>' : ''}
        </div>
      </div>`;
    document.body.appendChild(overlay);
    const body = overlay.querySelector('.desc-body');
    setTimeout(() => body.focus(), 30);
    overlay.querySelectorAll('.desc-toolbar [data-cmd]').forEach((btn) => {
      btn.addEventListener('mousedown', (e) => {
        e.preventDefault();
        document.execCommand(btn.dataset.cmd, false, btn.dataset.arg || null);
        body.focus();
      });
    });
    const close = () => overlay.remove();
    overlay.querySelector('#descClose').addEventListener('click', close);
    overlay.querySelector('#descCancel').addEventListener('click', close);
    overlay.querySelector('#descSave').addEventListener('click', () => {
      const html = body.innerHTML.trim();
      const isEmpty = !html || html === '<p></p>' || html === '<br>';
      a.description = isEmpty ? null : html;
      a.updatedAt = todayISO();
      commit('description');
      close();
      toast(isEmpty ? 'Description cleared' : 'Description saved');
    });
    overlay.querySelector('#descClear')?.addEventListener('click', () => {
      a.description = null;
      a.updatedAt = todayISO();
      commit('description-clear');
      close();
      toast('Description removed');
    });
    // Modal closes ONLY via the explicit × button (and Cancel where present).
    // Backdrop click and Escape are intentionally NOT bound — losing
    // in-progress edits to a stray click outside the modal is too easy.
  }

  // Hover tooltip showing rich description
  let hoverDescEl = null;
  let hoverDescTimer = null;
  let hoverDescAnchor = null;
  function ensureHoverDescEl() {
    if (hoverDescEl) return hoverDescEl;
    hoverDescEl = document.createElement('div');
    hoverDescEl.className = 'action-desc-tooltip';
    document.body.appendChild(hoverDescEl);
    return hoverDescEl;
  }
  function showHoverDesc(a, clientX, clientY) {
    const el = ensureHoverDescEl();
    const parts = [];
    if (a.description) parts.push(`<div class="hover-desc">${a.description}</div>`);
    if (a.notes) {
      const safeNote = escapeHTML(a.notes).replace(/\n/g, '<br>');
      parts.push(`<div class="hover-note"><span class="hover-note-lbl">Note</span>${safeNote}</div>`);
    }
    if (!parts.length) return;
    el.innerHTML = parts.join('');
    el.style.display = 'block';
    el.style.left = '0px';
    el.style.top  = '0px';
    const r = el.getBoundingClientRect();
    let x = clientX + 14;
    let y = clientY + 14;
    if (x + r.width  > innerWidth - 8)  x = innerWidth - r.width - 8;
    if (y + r.height > innerHeight - 8) y = clientY - r.height - 14;
    el.style.left = Math.max(8, x) + 'px';
    el.style.top  = Math.max(8, y) + 'px';
  }
  function hideHoverDesc() {
    if (hoverDescEl) hoverDescEl.style.display = 'none';
    clearTimeout(hoverDescTimer);
    hoverDescTimer = null;
    // anchor is managed by mouseover/mouseout callers — don't clobber it here
  }
  function wireHoverDescOnce() {
    if (document._descWired) return;
    document._descWired = true;
    document.addEventListener('mouseover', (e) => {
      const el = e.target.closest('.card[data-id], .reg-row[data-id], .tl-bar[data-id]');
      if (el === hoverDescAnchor) return;
      // Anchor changed (incl. to null) — kill any pending/visible tooltip first
      hideHoverDesc();
      hoverDescAnchor = el;
      if (!el) return;
      const id = el.dataset.id;
      const a = state.projects.flatMap((p) => p.actions || []).find((x) => x.id === id);
      if (!a || (!a.description && !a.notes)) return;
      hoverDescTimer = setTimeout(() => showHoverDesc(a, e.clientX, e.clientY), 350);
    });
    document.addEventListener('mouseout', (e) => {
      if (!hoverDescAnchor) return;
      const next = e.relatedTarget;
      if (next && hoverDescAnchor.contains(next)) return;
      hideHoverDesc();
      hoverDescAnchor = null;
    });
    document.addEventListener('mousemove', (e) => {
      if (!hoverDescEl || hoverDescEl.style.display === 'none') return;
      const r = hoverDescEl.getBoundingClientRect();
      let x = e.clientX + 14;
      let y = e.clientY + 14;
      if (x + r.width  > innerWidth - 8)  x = innerWidth - r.width - 8;
      if (y + r.height > innerHeight - 8) y = e.clientY - r.height - 14;
      hoverDescEl.style.left = Math.max(8, x) + 'px';
      hoverDescEl.style.top  = Math.max(8, y) + 'px';
    });
  }

  function attachColumnDND(body) {
    body.addEventListener('dragover', (e) => {
      if (!e.dataTransfer.types.includes('text/cockpit-action')) return;
      e.preventDefault();
      e.dataTransfer.dropEffect = 'move';
      body.classList.add('drop-target');
      // Find insertion point
      const after = [...body.querySelectorAll('.card:not(.dragging)')].find((el) => {
        const rect = el.getBoundingClientRect();
        return e.clientY < rect.top + rect.height / 2;
      });
      const dragging = document.querySelector('.card.dragging');
      if (!dragging) return;
      if (after) body.insertBefore(dragging, after);
      else body.appendChild(dragging);
    });
    body.addEventListener('dragleave', () => body.classList.remove('drop-target'));
    body.addEventListener('drop', (e) => {
      e.preventDefault();
      body.classList.remove('drop-target');
      const id = e.dataTransfer.getData('text/cockpit-action');
      const newStatus = body.dataset.status;
      const proj = curProject();
      const a = proj.actions.find((x) => x.id === id);
      if (!a) return;
      const oldStatus = a.status;
      a.status = newStatus;
      a.updatedAt = todayISO();
      a.history.push({ at: todayISO(), what: `Status: ${oldStatus} → ${newStatus}` });
      // recompute priority by visual order
      const ordered = [...body.querySelectorAll('.card')].map((c) => c.dataset.id);
      ordered.forEach((cid, i) => {
        const aa = proj.actions.find((x) => x.id === cid);
        if (aa) aa.priority = i;
      });
      commit('move');
      toast(oldStatus === newStatus ? 'Reordered' : `Moved to ${newStatus}`);
    });
  }

  /* --------------------------- Holidays ------------------------------ */
  /* Built-in national holidays for a small set of countries (2026-2027).
     Dates are observed dates where applicable. */
  const HOLIDAYS = {
    US: { name: 'United States', flag: '🇺🇸', dates: {
      '2026-01-01': "New Year's Day",
      '2026-01-19': 'Martin Luther King Jr. Day',
      '2026-02-16': "Presidents' Day",
      '2026-05-25': 'Memorial Day',
      '2026-06-19': 'Juneteenth',
      '2026-07-03': 'Independence Day (observed)',
      '2026-09-07': 'Labor Day',
      '2026-10-12': 'Columbus Day',
      '2026-11-11': 'Veterans Day',
      '2026-11-26': 'Thanksgiving',
      '2026-12-25': 'Christmas Day',
      '2027-01-01': "New Year's Day",
      '2027-01-18': 'Martin Luther King Jr. Day',
      '2027-02-15': "Presidents' Day",
      '2027-05-31': 'Memorial Day',
      '2027-06-18': 'Juneteenth (observed)',
      '2027-07-05': 'Independence Day (observed)',
      '2027-09-06': 'Labor Day',
      '2027-10-11': 'Columbus Day',
      '2027-11-11': 'Veterans Day',
      '2027-11-25': 'Thanksgiving',
      '2027-12-24': 'Christmas Day (observed)',
    }},
    UK: { name: 'United Kingdom', flag: '🇬🇧', dates: {
      '2026-01-01': "New Year's Day",
      '2026-04-03': 'Good Friday',
      '2026-04-06': 'Easter Monday',
      '2026-05-04': 'Early May Bank Holiday',
      '2026-05-25': 'Spring Bank Holiday',
      '2026-08-31': 'Summer Bank Holiday',
      '2026-12-25': 'Christmas Day',
      '2026-12-28': 'Boxing Day (observed)',
      '2027-01-01': "New Year's Day",
      '2027-03-26': 'Good Friday',
      '2027-03-29': 'Easter Monday',
      '2027-05-03': 'Early May Bank Holiday',
      '2027-05-31': 'Spring Bank Holiday',
      '2027-08-30': 'Summer Bank Holiday',
      '2027-12-27': 'Christmas Day (observed)',
      '2027-12-28': 'Boxing Day (observed)',
    }},
    FR: { name: 'France', flag: '🇫🇷', dates: {
      '2026-01-01': "Jour de l'An",
      '2026-04-06': 'Lundi de Pâques',
      '2026-05-01': 'Fête du Travail',
      '2026-05-08': 'Victoire 1945',
      '2026-05-14': 'Ascension',
      '2026-05-25': 'Lundi de Pentecôte',
      '2026-07-14': 'Fête nationale',
      '2026-08-15': 'Assomption',
      '2026-11-01': 'Toussaint',
      '2026-11-11': 'Armistice',
      '2026-12-25': 'Noël',
      '2027-01-01': "Jour de l'An",
      '2027-03-29': 'Lundi de Pâques',
      '2027-05-01': 'Fête du Travail',
      '2027-05-06': 'Ascension',
      '2027-05-08': 'Victoire 1945',
      '2027-05-17': 'Lundi de Pentecôte',
      '2027-07-14': 'Fête nationale',
      '2027-08-15': 'Assomption',
      '2027-11-01': 'Toussaint',
      '2027-11-11': 'Armistice',
      '2027-12-25': 'Noël',
    }},
    DE: { name: 'Germany', flag: '🇩🇪', dates: {
      '2026-01-01': 'Neujahrstag',
      '2026-04-03': 'Karfreitag',
      '2026-04-06': 'Ostermontag',
      '2026-05-01': 'Tag der Arbeit',
      '2026-05-14': 'Christi Himmelfahrt',
      '2026-05-25': 'Pfingstmontag',
      '2026-10-03': 'Tag der Deutschen Einheit',
      '2026-12-25': '1. Weihnachtstag',
      '2026-12-26': '2. Weihnachtstag',
      '2027-01-01': 'Neujahrstag',
      '2027-03-26': 'Karfreitag',
      '2027-03-29': 'Ostermontag',
      '2027-05-01': 'Tag der Arbeit',
      '2027-05-06': 'Christi Himmelfahrt',
      '2027-05-17': 'Pfingstmontag',
      '2027-10-03': 'Tag der Deutschen Einheit',
      '2027-12-25': '1. Weihnachtstag',
      '2027-12-26': '2. Weihnachtstag',
    }},
    JP: { name: 'Japan', flag: '🇯🇵', dates: {
      '2026-01-01': 'New Year',
      '2026-01-12': 'Coming of Age Day',
      '2026-02-11': 'National Foundation Day',
      '2026-02-23': "Emperor's Birthday",
      '2026-03-20': 'Vernal Equinox',
      '2026-04-29': 'Shōwa Day',
      '2026-05-03': 'Constitution Day',
      '2026-05-04': 'Greenery Day',
      '2026-05-05': "Children's Day",
      '2026-07-20': 'Marine Day',
      '2026-08-11': 'Mountain Day',
      '2026-09-21': 'Respect for the Aged Day',
      '2026-09-23': 'Autumnal Equinox',
      '2026-10-12': 'Sports Day',
      '2026-11-03': 'Culture Day',
      '2026-11-23': 'Labour Thanksgiving Day',
      '2027-01-01': 'New Year',
      '2027-01-11': 'Coming of Age Day',
      '2027-02-11': 'National Foundation Day',
      '2027-02-23': "Emperor's Birthday",
      '2027-03-21': 'Vernal Equinox',
      '2027-04-29': 'Shōwa Day',
      '2027-05-03': 'Constitution Day',
      '2027-05-04': 'Greenery Day',
      '2027-05-05': "Children's Day",
      '2027-07-19': 'Marine Day',
      '2027-08-11': 'Mountain Day',
      '2027-09-20': 'Respect for the Aged Day',
      '2027-09-23': 'Autumnal Equinox',
      '2027-10-11': 'Sports Day',
      '2027-11-03': 'Culture Day',
      '2027-11-23': 'Labour Thanksgiving Day',
    }},
    CA: { name: 'Canada', flag: '🇨🇦', dates: {
      '2026-01-01': "New Year's Day",
      '2026-02-16': 'Family Day',
      '2026-04-03': 'Good Friday',
      '2026-05-18': 'Victoria Day',
      '2026-07-01': 'Canada Day',
      '2026-08-03': 'Civic Holiday',
      '2026-09-07': 'Labour Day',
      '2026-10-12': 'Thanksgiving',
      '2026-11-11': 'Remembrance Day',
      '2026-12-25': 'Christmas Day',
      '2026-12-28': 'Boxing Day (observed)',
      '2027-01-01': "New Year's Day",
      '2027-02-15': 'Family Day',
      '2027-03-26': 'Good Friday',
      '2027-05-24': 'Victoria Day',
      '2027-07-01': 'Canada Day',
      '2027-08-02': 'Civic Holiday',
      '2027-09-06': 'Labour Day',
      '2027-10-11': 'Thanksgiving',
      '2027-11-11': 'Remembrance Day',
      '2027-12-27': 'Christmas Day (observed)',
      '2027-12-28': 'Boxing Day (observed)',
    }},
    IN: { name: 'India', flag: '🇮🇳', dates: {
      '2026-01-26': 'Republic Day',
      '2026-03-04': 'Holi',
      '2026-03-21': 'Eid al-Fitr',
      '2026-04-03': 'Good Friday',
      '2026-08-15': 'Independence Day',
      '2026-10-02': 'Gandhi Jayanti',
      '2026-10-20': 'Diwali',
      '2026-12-25': 'Christmas Day',
      '2027-01-26': 'Republic Day',
      '2027-03-22': 'Holi',
      '2027-08-15': 'Independence Day',
      '2027-10-02': 'Gandhi Jayanti',
      '2027-11-08': 'Diwali',
      '2027-12-25': 'Christmas Day',
    }},
    ES: { name: 'Spain', flag: '🇪🇸', dates: {
      '2026-01-01': 'Año Nuevo',
      '2026-01-06': 'Reyes Magos',
      '2026-04-03': 'Viernes Santo',
      '2026-05-01': 'Día del Trabajo',
      '2026-08-15': 'Asunción',
      '2026-10-12': 'Fiesta Nacional',
      '2026-11-01': 'Todos los Santos',
      '2026-12-06': 'Día de la Constitución',
      '2026-12-08': 'Inmaculada Concepción',
      '2026-12-25': 'Navidad',
      '2027-01-01': 'Año Nuevo',
      '2027-01-06': 'Reyes Magos',
      '2027-03-26': 'Viernes Santo',
      '2027-05-01': 'Día del Trabajo',
      '2027-08-15': 'Asunción',
      '2027-10-12': 'Fiesta Nacional',
      '2027-11-01': 'Todos los Santos',
      '2027-12-06': 'Día de la Constitución',
      '2027-12-08': 'Inmaculada Concepción',
      '2027-12-25': 'Navidad',
    }},
  };

  function activeHolidayCodes() {
    return (state.settings && state.settings.holidayCountries) || [];
  }
  function holidayInfo(iso) {
    const codes = activeHolidayCodes();
    if (!codes.length) return null;
    const matches = [];
    if (codes.includes('WKND')) {
      const day = parseDate(iso).getDay();
      if (day === 0 || day === 6) {
        matches.push({ code: 'WKND', country: 'Weekend', flag: '🛌',
          name: day === 0 ? 'Sunday' : 'Saturday', isWeekend: true });
      }
    }
    codes.forEach((c) => {
      if (c === 'WKND') return;
      const data = HOLIDAYS[c];
      if (data && data.dates[iso]) {
        matches.push({ code: c, country: data.name, flag: data.flag, name: data.dates[iso] });
      }
    });
    return matches.length ? matches : null;
  }

  /* ---------------------------- Archive ------------------------------ */

  function renderArchive(root) {
    const view = document.createElement('div');
    view.className = 'view';
    // Always cross-project — show every soft-deleted action
    const items = state.projects.flatMap((p) => (p.actions || [])
      .filter((a) => a.deletedAt)
      .map((a) => ({ a, proj: p })));
    items.sort((x, y) => (y.a.deletedAt || '').localeCompare(x.a.deletedAt || ''));
    view.innerHTML = `
      <div class="page-head">
        <div>
          <div class="page-title">Archive</div>
          <div class="page-sub">${items.length} archived action${items.length === 1 ? '' : 's'} — restore or permanently delete.</div>
        </div>
      </div>
      ${items.length ? `
        <div class="register archive-table">
          <div class="reg-head archive-row">
            <button class="reg-col">Title</button>
            <button class="reg-col">Project</button>
            <button class="reg-col">Owner</button>
            <button class="reg-col">Was</button>
            <button class="reg-col">Archived</button>
            <span class="reg-col-spacer"></span>
            <span class="reg-col-spacer"></span>
          </div>
          <div class="reg-body">
            ${items.map(({ a, proj: pr }) => `
              <div class="reg-row archive-row" data-id="${a.id}" data-pid="${pr.id}">
                <div class="reg-cell title-cell">${escapeHTML(a.title)}</div>
                <div class="reg-cell muted">${escapeHTML(pr.name)}</div>
                <div class="reg-cell"><span class="avatar">${initials(personName(a.owner))}</span> ${escapeHTML(personName(a.owner))}</div>
                <div class="reg-cell muted">${escapeHTML(a.status)}</div>
                <div class="reg-cell muted">${fmtDate(a.deletedAt)}</div>
                <div class="reg-cell">
                  <button class="ghost archive-restore" title="Restore">↺ Restore</button>
                </div>
                <div class="reg-cell">
                  <button class="row-del archive-purge" title="Delete permanently">×</button>
                </div>
              </div>`).join('')}
          </div>
        </div>` : '<div class="empty">No archived items. Move actions to Archive from the board (drag to bin) or the Register × button.</div>'}`;
    root.appendChild(view);

    view.querySelectorAll('.archive-restore').forEach((btn) => {
      btn.addEventListener('click', () => {
        const row = btn.closest('.reg-row');
        const pid = row.dataset.pid;
        const id = row.dataset.id;
        const proj = state.projects.find((p) => p.id === pid);
        const a = (proj?.actions || []).find((x) => x.id === id);
        if (!a) return;
        a.deletedAt = null;
        a.history.push({ at: todayISO(), what: 'Restored from Archive' });
        a.updatedAt = todayISO();
        commit('restore');
        toast('Restored');
      });
    });
    view.querySelectorAll('.archive-purge').forEach((btn) => {
      btn.addEventListener('click', () => {
        const row = btn.closest('.reg-row');
        const pid = row.dataset.pid;
        const id = row.dataset.id;
        if (!confirm('Permanently delete this action? This cannot be undone (except via undo while still in this session).')) return;
        const proj = state.projects.find((p) => p.id === pid);
        if (proj) proj.actions = proj.actions.filter((x) => x.id !== id);
        commit('purge');
        toast('Permanently deleted');
      });
    });
  }

  /* --------------------------- Open Points --------------------------- */

  // Wire interactions on the resolution-steps section of an open-point row.
  // `getOp()` returns the live op object so handlers always read the current
  // version (the row may be re-rendered after each commit).
  function wireOpStepHandlers(itemEl, getOp) {
    const stepsEl = itemEl.querySelector('.op-steps');
    if (!stepsEl) return;
    const listEl = stepsEl.querySelector('.op-steps-list');

    // Toggle done
    stepsEl.querySelectorAll('.op-step-check').forEach((cb) => {
      cb.addEventListener('click', (e) => {
        e.stopPropagation();
        const op = getOp(); if (!op) return;
        const id = cb.closest('.op-step').dataset.stepId;
        const s = (op.steps || []).find((x) => x.id === id);
        if (!s) return;
        s.done = !s.done;
        commit('op-step-toggle');
      });
    });

    // Inline edit — save on blur, Enter saves & blurs, Escape cancels
    stepsEl.querySelectorAll('.op-step-text').forEach((txt) => {
      // Empty + Backspace deletes the row
      txt.addEventListener('keydown', (e) => {
        if (e.key === 'Enter' && !e.shiftKey) { e.preventDefault(); txt.blur(); }
        else if (e.key === 'Backspace' && txt.textContent === '') {
          e.preventDefault();
          const op = getOp(); if (!op) return;
          const id = txt.closest('.op-step').dataset.stepId;
          op.steps = (op.steps || []).filter((s) => s.id !== id);
          commit('op-step-delete');
        }
      });
      txt.addEventListener('blur', () => {
        const op = getOp(); if (!op) return;
        const id = txt.closest('.op-step').dataset.stepId;
        const s = (op.steps || []).find((x) => x.id === id);
        if (!s) return;
        const v = txt.textContent.trim();
        if (s.text !== v) { s.text = v; commit('op-step-edit'); }
      });
    });

    // Delete
    stepsEl.querySelectorAll('.op-step-del').forEach((btn) => {
      btn.addEventListener('click', (e) => {
        e.stopPropagation();
        const op = getOp(); if (!op) return;
        const id = btn.closest('.op-step').dataset.stepId;
        op.steps = (op.steps || []).filter((s) => s.id !== id);
        commit('op-step-delete');
      });
    });

    // Add
    const addBtn = stepsEl.querySelector('.op-step-add');
    if (addBtn) {
      addBtn.addEventListener('click', () => {
        const op = getOp(); if (!op) return;
        op.steps = op.steps || [];
        const newStep = { id: uid('st'), text: '', done: false };
        op.steps.push(newStep);
        commit('op-step-add');
        // After re-render, focus the new step's text
        setTimeout(() => {
          const newEl = document.querySelector(`.op-step[data-step-id="${newStep.id}"] .op-step-text`);
          if (newEl) {
            newEl.focus();
            // Move caret to end (empty so just focus)
            const range = document.createRange();
            range.selectNodeContents(newEl);
            range.collapse(false);
            const sel = window.getSelection();
            sel.removeAllRanges(); sel.addRange(range);
          }
        }, 0);
      });
    }

    // Drag-to-reorder via the grip — custom mouse events so we can preview
    // the move live and stay consistent with the rest of the app's drags.
    stepsEl.querySelectorAll('.op-step-grip').forEach((grip) => {
      grip.addEventListener('mousedown', (e) => {
        e.preventDefault();
        const row = grip.closest('.op-step');
        if (!row || !listEl) return;
        row.classList.add('dragging');
        document.body.classList.add('is-step-dragging');
        const onMove = (em) => {
          const siblings = [...listEl.querySelectorAll('.op-step:not(.dragging)')];
          const after = siblings.find((sib) => {
            const r = sib.getBoundingClientRect();
            return em.clientY < r.top + r.height / 2;
          });
          if (after) listEl.insertBefore(row, after);
          else listEl.appendChild(row);
        };
        const onUp = () => {
          row.classList.remove('dragging');
          document.body.classList.remove('is-step-dragging');
          window.removeEventListener('mousemove', onMove);
          window.removeEventListener('mouseup', onUp);
          // Commit the new order from current DOM
          const op = getOp(); if (!op) return;
          const newOrderIds = [...listEl.querySelectorAll('.op-step')].map((r) => r.dataset.stepId);
          const before = (op.steps || []).map((s) => s.id).join(',');
          op.steps = newOrderIds.map((id) => (op.steps || []).find((s) => s.id === id)).filter(Boolean);
          if (op.steps.map((s) => s.id).join(',') !== before) commit('op-step-reorder');
        };
        window.addEventListener('mousemove', onMove);
        window.addEventListener('mouseup', onUp);
      });
    });
  }

  // Persistent filter / sort state for the Open Points page (volatile; not
  // in state because it's pure UI).
  const opFilterState = { q: '', component: '', criticality: '', priority: '', progress: '', sort: 'manual' };
  function renderOpenPoints(root) {
    const proj = curProject();
    proj.openPoints = proj.openPoints || [];
    const view = document.createElement('div');
    view.className = 'view';
    view.innerHTML = `
      <div class="page-head">
        <div>
          <div class="page-title">Open points</div>
          <div class="page-sub">Quick capture for ideas, follow-ups, and questions. Convert any item to an action when ready.</div>
        </div>
      </div>
      <div class="op-quick">
        <input id="opInput" type="text" placeholder="Type and press Enter — anything to remember (e.g. ‘ask vendor about cell datasheet’)" />
        <select id="opQuickComp" title="Optional: link to a component on capture">
          <option value="">— component</option>
          ${(proj.components || []).map((pt) => `<option value="${pt.id}">${escapeHTML(pt.name)}</option>`).join('')}
        </select>
        <button class="ghost" id="opAdd">Add</button>
      </div>
      <div class="op-filterbar">
        <input id="opFilterQ" type="search" placeholder="Filter by text…" value="${escapeHTML(opFilterState.q)}" />
        <select id="opFilterComp" title="Filter by component">
          <option value="">All components</option>
          <option value="__none">— No component</option>
          ${(proj.components || []).map((pt) => `<option value="${pt.id}" ${pt.id === opFilterState.component ? 'selected' : ''}>${escapeHTML(pt.name)}</option>`).join('')}
        </select>
        <select id="opFilterCrit" title="Filter by criticality">
          <option value="">All criticality</option>
          ${['low','med','high','critical'].map((k) => `<option value="${k}" ${k === opFilterState.criticality ? 'selected' : ''}>${CRITICALITY_LABEL[k]}</option>`).join('')}
        </select>
        <select id="opFilterPrio" title="Filter by priority">
          <option value="">All priority</option>
          ${PRIORITY_LEVELS.map((p) => `<option value="${p.id}" ${p.id === opFilterState.priority ? 'selected' : ''}>${p.label}</option>`).join('')}
        </select>
        <select id="opFilterProg" title="Filter by step progress">
          <option value="">All progress</option>
          <option value="none"      ${opFilterState.progress === 'none' ? 'selected' : ''}>No steps</option>
          <option value="some"      ${opFilterState.progress === 'some' ? 'selected' : ''}>In progress</option>
          <option value="all-done"  ${opFilterState.progress === 'all-done' ? 'selected' : ''}>All steps done</option>
        </select>
        <select id="opSort" title="Sort">
          <option value="manual"    ${opFilterState.sort === 'manual' ? 'selected' : ''}>Manual order</option>
          <option value="newest"    ${opFilterState.sort === 'newest' ? 'selected' : ''}>Newest first</option>
          <option value="oldest"    ${opFilterState.sort === 'oldest' ? 'selected' : ''}>Oldest first</option>
          <option value="crit-desc" ${opFilterState.sort === 'crit-desc' ? 'selected' : ''}>Criticality ↓</option>
          <option value="prio-desc" ${opFilterState.sort === 'prio-desc' ? 'selected' : ''}>Priority ↓</option>
          <option value="title-asc" ${opFilterState.sort === 'title-asc' ? 'selected' : ''}>Title A–Z</option>
          <option value="progress"  ${opFilterState.sort === 'progress' ? 'selected' : ''}>Step progress</option>
        </select>
        <button class="ghost" id="opFilterReset" title="Clear filters">Reset</button>
      </div>
      <div class="op-list" id="opList"></div>`;
    root.appendChild(view);

    // Helpers — sortable rank (Critical → Low) and matcher
    const CRIT_RANK = { critical: 0, high: 1, med: 2, low: 3 };
    function matchesFilters(op) {
      if (opFilterState.q) {
        const q = opFilterState.q.toLowerCase();
        const hay = (op.title + ' ' + (op.notes || '') + ' ' + (op.steps || []).map((s) => s.text).join(' ')).toLowerCase();
        if (!hay.includes(q)) return false;
      }
      if (opFilterState.component) {
        if (opFilterState.component === '__none') { if (op.component) return false; }
        else if (op.component !== opFilterState.component) return false;
      }
      if (opFilterState.criticality && (op.criticality || 'med') !== opFilterState.criticality) return false;
      if (opFilterState.priority    && (op.priorityLevel || 'med') !== opFilterState.priority) return false;
      if (opFilterState.progress) {
        const total = (op.steps || []).length;
        const done  = (op.steps || []).filter((s) => s.done).length;
        if (opFilterState.progress === 'none' && total > 0) return false;
        if (opFilterState.progress === 'some' && (total === 0 || done === total)) return false;
        if (opFilterState.progress === 'all-done' && (total === 0 || done < total)) return false;
      }
      return true;
    }
    function compareForSort(a, b) {
      switch (opFilterState.sort) {
        case 'newest':    return (b.createdAt || '').localeCompare(a.createdAt || '');
        case 'oldest':    return (a.createdAt || '').localeCompare(b.createdAt || '');
        case 'crit-desc': return (CRIT_RANK[a.criticality || 'med'] - CRIT_RANK[b.criticality || 'med']);
        case 'prio-desc': return (CRIT_RANK[a.priorityLevel || 'med'] - CRIT_RANK[b.priorityLevel || 'med']);
        case 'title-asc': return (a.title || '').localeCompare(b.title || '');
        case 'progress': {
          const pa = ((a.steps || []).filter((s) => s.done).length) / Math.max(1, (a.steps || []).length);
          const pb = ((b.steps || []).filter((s) => s.done).length) / Math.max(1, (b.steps || []).length);
          return pa - pb;
        }
        default: return 0; // manual — preserve array order
      }
    }

    function draw() {
      const list = $('#opList');
      const all = proj.openPoints || [];
      const filtered = all.filter(matchesFilters);
      const items = opFilterState.sort === 'manual' ? filtered : filtered.slice().sort(compareForSort);
      if (!all.length) {
        list.innerHTML = '<div class="empty">No open points yet — capture something above.</div>';
        return;
      }
      if (!items.length) {
        list.innerHTML = '<div class="empty">No open points match the current filters.</div>';
        return;
      }
      list.innerHTML = items.map((op) => {
        // Backfill defaults for legacy items
        if (!op.criticality) op.criticality = 'med';
        // Left-border tint reflects the *dominant* severity of either
        // criticality or priority — whichever is higher drives the colour
        // so a Critical-priority point is red even if its criticality is
        // only Medium, and vice versa.
        const SEV_ORDER = ['low', 'med', 'high', 'critical'];
        const critIdx = Math.max(0, SEV_ORDER.indexOf(op.criticality || 'med'));
        const prioIdx = Math.max(0, SEV_ORDER.indexOf(op.priorityLevel || 'med'));
        const domKey = SEV_ORDER[Math.max(critIdx, prioIdx)];
        const domRgb = CRITICALITY_RGB[domKey];
        if (!op.createdAt) op.createdAt = todayISO();
        if (!Array.isArray(op.steps)) op.steps = [];
        const cmp = findComponent(proj, op.component);
        const c = cmp ? componentColor(cmp.color) : null;
        const critRgb = CRITICALITY_RGB[op.criticality] || CRITICALITY_RGB.med;
        // The left border tracks criticality (component shown as a chip below).
        const tint = `style="border-left-color: rgb(${domRgb})"`;
        const critLabel = CRITICALITY_LABEL[op.criticality] || 'Medium';
        // op.notes holds rich HTML; legacy plain-string entries render as text
        const contextHtml = op.notes && /<\w+/.test(op.notes) ? op.notes : escapeHTML(op.notes || '');
        const stepDone = op.steps.filter((s) => s.done).length;
        const stepTotal = op.steps.length;
        const stepsAllDone = stepTotal > 0 && stepDone === stepTotal;
        const prio = priorityLevel(op.priorityLevel);
        // Both chips ALWAYS render so the user can click them to change
        // values. The "med" (default) state uses a quieter visual treatment
        // so the title row stays calm at rest and chips only "light up"
        // when they carry a non-default signal.
        const prioId = op.priorityLevel || 'med';
        const prioQuiet = prioId === 'med';
        const critQuiet = (op.criticality || 'med') === 'med';
        const critChipStyle = critQuiet
          ? `background:transparent;color:rgb(${critRgb});border:1px dashed rgba(${critRgb},.55)`
          : `background:rgba(${critRgb},.18);color:rgb(${critRgb});border:1px solid rgb(${critRgb})`;
        const prioChipStyle = prioQuiet
          ? `background:transparent;color:rgba(${prio.rgb},.85);border:1px dashed rgba(${prio.rgb},.55)`
          : `background:rgba(${prio.rgb},.18);color:rgb(${prio.rgb});border:1px solid rgb(${prio.rgb})`;
        return `
        <div class="op-item crit-${op.criticality}" data-id="${op.id}" ${tint}>
          <span class="op-row-grip" title="Drag to reorder" aria-hidden="true">⋮⋮</span>
          <div class="op-content">
            <div class="op-title-row">
              <div class="op-title" contenteditable="true" data-field="title">${escapeHTML(op.title)}</div>
            </div>
            <div class="op-context-wrap">
              <div class="op-toolbar">
                <button type="button" data-cmd="bold" title="Bold"><b>B</b></button>
                <button type="button" data-cmd="italic" title="Italic"><i>I</i></button>
                <button type="button" data-cmd="underline" title="Underline"><u>U</u></button>
                <span class="sep"></span>
                <button type="button" data-cmd="insertUnorderedList" title="Bullet list">•</button>
                <button type="button" data-cmd="insertOrderedList" title="Numbered list">1.</button>
                <span class="sep"></span>
                <button type="button" data-cmd="createLink" title="Insert link">🔗</button>
                <label class="op-color" title="Text colour"><input type="color" data-cmd="foreColor" /></label>
                <button type="button" data-cmd="removeFormat" title="Clear formatting">✕</button>
              </div>
              <div class="op-context" contenteditable="true" data-placeholder="Add rich context — bold, lists, links, colour…">${contextHtml}</div>
            </div>
            <div class="op-steps ${stepsAllDone ? 'all-done' : ''}">
              ${stepTotal ? `
                <div class="op-steps-head">
                  <span class="op-steps-lbl">Resolution steps</span>
                  <span class="op-steps-progress" aria-hidden="true">
                    <span class="op-steps-bar"><span class="op-steps-bar-fill" style="width:${stepTotal ? Math.round(stepDone / stepTotal * 100) : 0}%"></span></span>
                    <span class="op-steps-count">${stepDone}/${stepTotal}</span>
                  </span>
                </div>
              ` : ''}
              <div class="op-steps-list">
                ${op.steps.map((s) => `
                  <div class="op-step ${s.done ? 'done' : ''}" data-step-id="${s.id}">
                    <span class="op-step-grip" title="Drag to reorder" aria-hidden="true">⋮⋮</span>
                    <button type="button" class="op-step-check ${s.done ? 'on' : ''}" aria-label="${s.done ? 'Mark as not done' : 'Mark as done'}"></button>
                    <span class="op-step-text" contenteditable="true" data-placeholder="Step…">${escapeHTML(s.text)}</span>
                    <button type="button" class="op-step-del" title="Delete step" aria-label="Delete step">×</button>
                  </div>
                `).join('')}
              </div>
              <button type="button" class="op-step-add">+ Add step</button>
            </div>
            <div class="op-meta">
              <span class="op-origin" title="Auto-set when this open point was originated">Originated ${fmtFull(op.createdAt)}</span>
              ${stepTotal ? `<span class="op-meta-steps ${stepsAllDone ? 'ok' : ''}" title="Resolution steps">✓ ${stepDone}/${stepTotal} steps</span>` : ''}
              <button type="button" class="op-level-chip op-crit-chip ${critQuiet ? 'is-default' : ''}" data-kind="criticality" title="Click to set criticality (severity if not addressed)" style="${critChipStyle}">${critLabel}</button>
              <button type="button" class="op-level-chip prio-chip prio-${prio.id} ${prioQuiet ? 'is-default' : ''}" data-kind="priority" title="Click to set priority (urgency to act)" style="${prioChipStyle}">${prio.label}</button>
              ${cmp ? `<span class="component-chip" style="background:rgba(${c.rgb},.2);color:rgb(${c.rgb})">${escapeHTML(cmp.name)}</span>` : ''}
            </div>
          </div>
          <div class="op-actions">
            <select class="op-component" title="Link to a component">
              <option value="">— component</option>
              ${(proj.components || []).map((pt) => `<option value="${pt.id}" ${pt.id === op.component ? 'selected' : ''}>${escapeHTML(pt.name)}</option>`).join('')}
            </select>
            <button class="primary op-promote">→ Action</button>
            <button class="ghost op-discard" title="Discard">×</button>
          </div>
        </div>`;
      }).join('');
      $$('.op-item', list).forEach((el) => {
        const id = el.dataset.id;
        // Title — plain text, save on blur
        const titleEl = el.querySelector('.op-title');
        if (titleEl) {
          titleEl.addEventListener('blur', () => {
            const op = proj.openPoints.find((x) => x.id === id);
            if (!op) return;
            const v = titleEl.textContent.trim();
            if (op.title !== v) { op.title = v; commit('op-title'); }
          });
          titleEl.addEventListener('keydown', (e) => {
            if (e.key === 'Enter' && !e.shiftKey) { e.preventDefault(); titleEl.blur(); }
          });
        }
        // Context — rich HTML, save on blur
        const ctxEl = el.querySelector('.op-context');
        if (ctxEl) {
          ctxEl.addEventListener('blur', () => {
            const op = proj.openPoints.find((x) => x.id === id);
            if (!op) return;
            let html = ctxEl.innerHTML.trim();
            if (html === '<br>' || html === '<p></p>') html = '';
            if ((op.notes || '') !== html) { op.notes = html; commit('op-context'); }
          });
        }
        // Toolbar — execCommand on click; mousedown preventDefault preserves selection
        el.querySelectorAll('.op-toolbar [data-cmd]').forEach((btn) => {
          btn.addEventListener('mousedown', (e) => {
            e.preventDefault();
            const cmd = btn.dataset.cmd;
            ctxEl?.focus();
            if (cmd === 'createLink') {
              const url = prompt('Link URL (https://…):');
              if (url) document.execCommand('createLink', false, url);
            } else if (cmd === 'foreColor') {
              // Triggered by color input change instead
            } else {
              document.execCommand(cmd, false, null);
            }
          });
        });
        const colorInput = el.querySelector('.op-toolbar input[type="color"]');
        if (colorInput) {
          colorInput.addEventListener('input', (e) => {
            ctxEl?.focus();
            document.execCommand('foreColor', false, e.target.value);
          });
        }
        const compSel = el.querySelector('.op-component');
        if (compSel) {
          compSel.addEventListener('change', () => {
            const op = proj.openPoints.find((x) => x.id === id);
            if (!op) return;
            op.component = compSel.value || null;
            commit('op-component');
          });
        }
        // Chip-driven criticality + priority — click the chip on the title
        // row to open a colour-swatch popover, pick a level, commit. The
        // <select> dropdowns these used to live in were removed: the chips
        // now own the mutation, declutters the right-side actions row.
        el.querySelectorAll('.op-level-chip').forEach((chip) => {
          chip.addEventListener('click', (e) => {
            e.stopPropagation();
            const op = proj.openPoints.find((x) => x.id === id);
            if (!op) return;
            const kind = chip.dataset.kind; // 'criticality' | 'priority'
            const current = (kind === 'criticality') ? (op.criticality || 'med') : (op.priorityLevel || 'med');
            showLevelPopover(chip, kind, current, (val) => {
              if (kind === 'criticality') { op.criticality = val; commit('op-criticality'); }
              else                        { op.priorityLevel = val; commit('op-priority'); }
            });
          });
        });
        // Resolution steps — checkbox / edit / delete / add / drag-to-reorder
        wireOpStepHandlers(el, () => proj.openPoints.find((x) => x.id === id));

        el.querySelector('.op-promote').addEventListener('click', () => {
          const op = proj.openPoints.find((x) => x.id === id);
          if (!op) return;
          // Send rich-HTML context to the action's `description` field, leaving
          // the plain `notes` empty (unless the legacy notes were plain text).
          const isHtml = op.notes && /<\w+/.test(op.notes);
          openQuickAdd('action', {
            title: op.title,
            notes: isHtml ? '' : (op.notes || ''),
            description: isHtml ? op.notes : '',
            component: op.component,
            priorityLevel: op.priorityLevel || op.criticality || 'med',
          }, () => {
            proj.openPoints = proj.openPoints.filter((x) => x.id !== id);
            commit('op-promote');
            toast('Converted to action');
          });
        });
        el.querySelector('.op-discard').addEventListener('click', () => {
          if (!confirm('Discard this open point?')) return;
          proj.openPoints = proj.openPoints.filter((x) => x.id !== id);
          commit('op-discard');
        });
        // Drag the row by its grip to reorder open points up or down. Works
        // in both single-project and merged All-Projects views — for merged,
        // each item's source project is reordered independently.
        const rowGrip = el.querySelector('.op-row-grip');
        if (rowGrip) {
          rowGrip.addEventListener('mousedown', (e) => {
            e.preventDefault();
            // Manual reorder implies manual sort — flip the dropdown so the new
            // order isn't immediately re-sorted away after the next render.
            if (opFilterState.sort !== 'manual') {
              opFilterState.sort = 'manual';
              const sortSel = $('#opSort');
              if (sortSel) sortSel.value = 'manual';
            }
            // Snapshot the set of currently-visible (filtered) ids so we can
            // preserve hidden items in their original slots after the drop.
            const visibleIds = new Set(
              [...list.querySelectorAll('.op-item[data-id]')].map((r) => r.dataset.id)
            );
            const listEl = list;
            el.classList.add('dragging');
            document.body.classList.add('is-op-dragging');
            const onMove = (em) => {
              em.preventDefault();
              const sibs = [...listEl.querySelectorAll('.op-item:not(.dragging)')];
              const after = sibs.find((sib) => {
                const r = sib.getBoundingClientRect();
                return em.clientY < r.top + r.height / 2;
              });
              if (after) listEl.insertBefore(el, after);
              else listEl.appendChild(el);
            };
            // Suppress fresh selections that the browser tries to start when
            // the cursor sweeps over a contenteditable mid-drag.
            const onSelectStart = (ev) => ev.preventDefault();
            const onUp = () => {
              el.classList.remove('dragging');
              document.body.classList.remove('is-op-dragging');
              window.removeEventListener('mousemove', onMove);
              window.removeEventListener('mouseup', onUp);
              document.removeEventListener('selectstart', onSelectStart);
              // Clear any selection the browser still managed to start
              try { window.getSelection()?.removeAllRanges(); } catch (_) {}
              // The new visual order of the visible (filtered) rows
              const newVisibleOrder = [...listEl.querySelectorAll('.op-item[data-id]')].map((r) => r.dataset.id);
              const merged = curProjectIsMerged();
              if (merged) {
                // Reorder each source project independently. Items are placed
                // in the order they appear among the visible rows; hidden
                // items in each project keep their original slots.
                let changed = false;
                state.projects.forEach((sourceProj) => {
                  const ops = sourceProj.openPoints || [];
                  if (!ops.length) return;
                  const idsHere = new Set(ops.map((o) => o.id));
                  const visibleHere = new Set([...visibleIds].filter((id) => idsHere.has(id)));
                  const newOrderHere = newVisibleOrder.filter((id) => idsHere.has(id));
                  const byId = Object.fromEntries(ops.map((o) => [o.id, o]));
                  const before = ops.map((o) => o.id).join(',');
                  let vi = 0;
                  const next = ops.map((op) => visibleHere.has(op.id) ? (byId[newOrderHere[vi++]] || op) : op);
                  if (next.map((o) => o.id).join(',') !== before) {
                    sourceProj.openPoints = next;
                    changed = true;
                  }
                });
                if (changed) commit('op-reorder');
                return;
              }
              // Single-project: walk the original array and substitute visible slots
              const before = (proj.openPoints || []).map((o) => o.id).join(',');
              const byId = Object.fromEntries((proj.openPoints || []).map((o) => [o.id, o]));
              let visIdx = 0;
              const result = (proj.openPoints || []).map((op) => {
                if (visibleIds.has(op.id)) {
                  return byId[newVisibleOrder[visIdx++]] || op;
                }
                return op;
              });
              proj.openPoints = result;
              if (proj.openPoints.map((o) => o.id).join(',') !== before) commit('op-reorder');
            };
            window.addEventListener('mousemove', onMove);
            window.addEventListener('mouseup', onUp);
            document.addEventListener('selectstart', onSelectStart);
          });
        }
      });
    }

    const addPoint = () => {
      if (curProjectIsMerged()) { toast('Pick a single project to add open points.'); return; }
      const v = $('#opInput').value.trim();
      if (!v) return;
      const comp = $('#opQuickComp')?.value || null;
      proj.openPoints.unshift({ id: uid('op'), title: v, notes: '', component: comp, criticality: 'med', createdAt: todayISO() });
      $('#opInput').value = '';
      if ($('#opQuickComp')) $('#opQuickComp').value = '';
      commit('op-add');
      toast('Captured');
    };
    $('#opAdd').addEventListener('click', addPoint);
    $('#opInput').addEventListener('keydown', (e) => {
      if (e.key === 'Enter') { e.preventDefault(); addPoint(); }
    });

    // Filter / sort wiring
    const fq = $('#opFilterQ');
    if (fq) fq.addEventListener('input', () => { opFilterState.q = fq.value; draw(); });
    const fc = $('#opFilterComp');
    if (fc) fc.addEventListener('change', () => { opFilterState.component = fc.value; draw(); });
    const fcrit = $('#opFilterCrit');
    if (fcrit) fcrit.addEventListener('change', () => { opFilterState.criticality = fcrit.value; draw(); });
    const fprio = $('#opFilterPrio');
    if (fprio) fprio.addEventListener('change', () => { opFilterState.priority = fprio.value; draw(); });
    const fprog = $('#opFilterProg');
    if (fprog) fprog.addEventListener('change', () => { opFilterState.progress = fprog.value; draw(); });
    const fsort = $('#opSort');
    if (fsort) fsort.addEventListener('change', () => { opFilterState.sort = fsort.value; draw(); });
    const fres = $('#opFilterReset');
    if (fres) fres.addEventListener('click', () => {
      opFilterState.q = ''; opFilterState.component = ''; opFilterState.criticality = '';
      opFilterState.priority = ''; opFilterState.progress = ''; opFilterState.sort = 'manual';
      if (fq) fq.value = '';
      [fc, fcrit, fprio, fprog].forEach((el) => { if (el) el.value = ''; });
      if (fsort) fsort.value = 'manual';
      draw();
    });

    draw();
  }

  /* ---------------------------- Register ----------------------------- */

  // Persistent sort state (so it survives navigation away and back)
  const regState = { sortBy: 'due', sortDir: 'asc' };

  // Predicted completion = explicit override or the due date.
  function predictedCompletion(a) {
    return a.predictedCompletion || a.due || '';
  }
  // Actual completion = explicit override, otherwise mined from history
  // (latest "Status: ... → done"), otherwise updatedAt for already-done.
  function actualCompletion(a) {
    if (a.actualCompletion) return a.actualCompletion;
    if (a.status === 'done') {
      const entry = (a.history || []).filter((h) => /Status:.*→\s*done/.test(h.what)).pop();
      return entry?.at || a.updatedAt || '';
    }
    return '';
  }

  function regSortValue(a, col, proj) {
    switch (col) {
      case 'title':          return (a.title || '').toLowerCase();
      case 'component':      return (findComponent(proj, a.component)?.name || 'zzz').toLowerCase();
      case 'owner':          return personName(a.owner).toLowerCase();
      // Open states first (todo · doing · blocked), then both closed
       // states grouped at the same extreme (done · cancelled). Sorting
       // ascending puts what's open at the top; descending pushes the
       // closed pile up. Cancelled was previously missing from the
       // index, which broke the order entirely.
      case 'status':         return ['todo','doing','blocked','done','cancelled'].indexOf(a.status);
      case 'priority':       return ['critical','high','med','low'].indexOf(a.priorityLevel || 'med');
      case 'due':            return a.due || '9999-99-99';
      case 'predicted':      return predictedCompletion(a) || '9999-99-99';
      case 'actual':         return actualCompletion(a) || '9999-99-99';
      case 'commitment':     return typeof a.commitment === 'number' ? a.commitment : 100;
      case 'updatedAt':      return a.updatedAt || '0000-00-00';
      case 'originatorDate': return a.originatorDate || '0000-00-00';
    }
    return 0;
  }

  function renderRegister(root) {
    const proj = curProject();
    const view = document.createElement('div');
    view.className = 'view';
    view.innerHTML = `
      <div class="page-head">
        <div>
          <div class="page-title">${escapeHTML(proj.name)} — Register</div>
          <div class="page-sub">Editable table — change any cell to update. KPIs above reflect the current filters.</div>
        </div>
        <div class="page-actions">
          <button class="ghost" id="btnOpenArchive" title="Open the Archive view">⌫ Archive${(proj.actions || []).filter((a) => a.deletedAt).length ? ` <span class="badge-count">${(proj.actions || []).filter((a) => a.deletedAt).length}</span>` : ''}</button>
          <button class="ghost" id="btnArchiveDone" title="Move all currently visible done actions to Archive">⌫ Archive done</button>
          <button class="ghost" id="btnAddAction">+ Action</button>
        </div>
      </div>
      <div class="reg-kpis" id="regKpis"></div>
      <div class="register">
        <div class="reg-head">
          <button class="reg-col" data-col="title">Title</button>
          <button class="reg-col" data-col="component">Component</button>
          <button class="reg-col" data-col="owner">Owner</button>
          <button class="reg-col" data-col="status">Status</button>
          <button class="reg-col" data-col="priority">Priority</button>
          <button class="reg-col" data-col="due">Due</button>
          <button class="reg-col" data-col="predicted" title="Predicted completion (defaults to due)">Predicted</button>
          <button class="reg-col" data-col="originatorDate" title="When this action was originated">Originator date</button>
          <span class="reg-col-spacer" aria-hidden="true"></span>
        </div>
        <div class="reg-body" id="regBody"></div>
      </div>`;
    root.appendChild(view);

    function drawKpis() {
      const filtered = (proj.actions || []).filter(actionMatchesFilters);
      const total = filtered.length;
      const today = todayISO();
      const cnt = { todo: 0, doing: 0, blocked: 0, done: 0, cancelled: 0 };
      let overdue = 0, soon = 0, openCmtSum = 0, openCount = 0;
      const byComp = new Map();
      filtered.forEach((a) => {
        cnt[a.status] = (cnt[a.status] || 0) + 1;
        if (!isClosedStatus(a.status)) {
          if (a.due) {
            const dd = dayDiff(a.due, today);
            if (dd < 0) overdue++;
            else if (dd <= 7) soon++;
          }
          openCmtSum += (typeof a.commitment === 'number') ? a.commitment : 100;
          openCount++;
        }
        const k = a.component || '__none';
        byComp.set(k, (byComp.get(k) || 0) + 1);
      });
      const donePct = total ? Math.round((cnt.done / total) * 100) : 0;
      const avgOpenCmt = openCount ? Math.round(openCmtSum / openCount) : 0;
      // Per-person work-equivalent: total open commitment ÷ number of people
      const personLoad = state.people.length ? Math.round(openCmtSum / state.people.length) : 0;

      // Donut
      const r = 26, cx = 32, cy = 32;
      const circ = 2 * Math.PI * r;
      let off = 0;
      const slices = [
        { v: cnt.done,    color: 'var(--ok)',     name: 'Done' },
        { v: cnt.doing,   color: 'var(--accent)', name: 'In progress' },
        { v: cnt.blocked, color: 'var(--bad)',    name: 'Blocked' },
        { v: cnt.todo,    color: 'var(--neutral)',name: 'Not started' },
      ];
      const statusKeys = ['done', 'doing', 'blocked', 'todo'];
      const donutSegs = slices.map((s, i) => {
        const len = total ? (s.v / total) * circ : 0;
        const seg = `<circle class="donut-slice" data-set-status="${statusKeys[i]}" cx="${cx}" cy="${cy}" r="${r}" fill="none" stroke="${s.color}" stroke-width="9" stroke-dasharray="${len.toFixed(2)} ${(circ - len).toFixed(2)}" stroke-dashoffset="${(-off).toFixed(2)}" transform="rotate(-90 ${cx} ${cy})"><title>${s.name}: ${s.v} (click to filter)</title></circle>`;
        off += len;
        return seg;
      }).join('');

      // Throughput sparkline (last 8 weeks)
      const today2 = new Date(); today2.setHours(0, 0, 0, 0);
      let weekStart = new Date(today2);
      while (weekStart.getDay() !== 1) weekStart = new Date(weekStart.getTime() - dayMs);
      const startW = new Date(weekStart.getTime() - 7 * 7 * dayMs);
      const buckets = new Array(8).fill(0);
      const re = /Status:.*→\s*done/;
      filtered.forEach((a) => {
        const e = (a.history || []).filter((h) => re.test(h.what)).pop();
        const w = e ? e.at : (a.status === 'done' ? a.updatedAt : null);
        if (!w) return;
        const idx = Math.floor((parseDate(w) - startW) / dayMs / 7);
        if (idx >= 0 && idx < 8) buckets[idx]++;
      });
      const maxB = Math.max(1, ...buckets);
      const sparkBars = buckets.map((v, i) => {
        const x = i * 11 + 1;
        const h = (v / maxB) * 28;
        const y = 32 - h;
        return `<rect x="${x}" y="${y}" width="9" height="${h.toFixed(1)}" rx="1" fill="var(--ok)" opacity="0.85"><title>${v} done</title></rect>`;
      }).join('');

      // Opened vs Completed — daily line chart over the past 30 days
      const days30 = 30;
      const startDay = new Date(today2.getTime() - (days30 - 1) * dayMs);
      const opened = new Array(days30).fill(0);
      const closed = new Array(days30).fill(0);
      filtered.forEach((a) => {
        if (a.createdAt) {
          const idx = Math.floor((parseDate(a.createdAt) - startDay) / dayMs);
          if (idx >= 0 && idx < days30) opened[idx]++;
        }
        const ent = (a.history || []).filter((h) => re.test(h.what)).pop();
        const when = ent?.at || (a.status === 'done' ? a.updatedAt : null);
        if (when) {
          const idx = Math.floor((parseDate(when) - startDay) / dayMs);
          if (idx >= 0 && idx < days30) closed[idx]++;
        }
      });
      const totalOpened = opened.reduce((s, v) => s + v, 0);
      const totalClosed = closed.reduce((s, v) => s + v, 0);
      // Compute total-open count over time by walking back from "today"
      const todayOpenCount = filtered.filter((x) => x.status !== 'done').length;
      const openSeries = new Array(days30).fill(0);
      let runningOpen = todayOpenCount;
      for (let i = days30 - 1; i >= 0; i--) {
        openSeries[i] = runningOpen;
        runningOpen = runningOpen - opened[i] + closed[i];
      }
      const lcW = 220, lcH = 70, lcPadL = 4, lcPadR = 4, lcPadT = 4, lcPadB = 4;
      const lcInnerW = lcW - lcPadL - lcPadR, lcInnerH = lcH - lcPadT - lcPadB;
      const lcMax = Math.max(1, ...opened, ...closed, ...openSeries);
      const lcX = (i) => lcPadL + (i / (days30 - 1)) * lcInnerW;
      const lcY = (v) => lcPadT + lcInnerH - (v / lcMax) * lcInnerH;
      const linePath = (arr) => arr.map((v, i) =>
        `${i === 0 ? 'M' : 'L'} ${lcX(i).toFixed(1)} ${lcY(v).toFixed(1)}`).join(' ');
      const openAreaPath = `M ${lcX(0).toFixed(1)} ${lcY(openSeries[0]).toFixed(1)} ` +
        openSeries.slice(1).map((v, i) => `L ${lcX(i + 1).toFixed(1)} ${lcY(v).toFixed(1)}`).join(' ') +
        ` L ${lcX(days30 - 1).toFixed(1)} ${lcY(0).toFixed(1)} L ${lcX(0).toFixed(1)} ${lcY(0).toFixed(1)} Z`;
      const openedPath = linePath(opened);
      const closedPath = linePath(closed);
      const lcTodayX = lcX(days30 - 1).toFixed(1);
      const lineChartSVG = `
        <svg class="kpi-linechart" viewBox="0 0 ${lcW} ${lcH}" preserveAspectRatio="none">
          <line class="lc-baseline" x1="${lcPadL}" x2="${lcW - lcPadR}" y1="${lcY(0).toFixed(1)}" y2="${lcY(0).toFixed(1)}" />
          <path class="lc-open-area" d="${openAreaPath}" />
          <path class="lc-opened" d="${openedPath}" />
          <path class="lc-closed" d="${closedPath}" />
          <line class="lc-today" x1="${lcTodayX}" x2="${lcTodayX}" y1="${lcPadT}" y2="${lcH - lcPadB}" />
        </svg>`;

      // Component distribution stacked bar
      const compEntries = [...byComp.entries()].sort((a, b) => b[1] - a[1]);
      const compBars = compEntries.map(([cid, n]) => {
        const cmp = cid === '__none' ? null : findComponent(proj, cid);
        const c = cmp ? componentColor(cmp.color) : null;
        const w = total ? (n / total) * 100 : 0;
        const color = c ? `rgba(${c.rgb},.85)` : 'var(--neutral)';
        const name = cmp ? cmp.name : 'Unassigned';
        const filter = cid === '__none' ? '__none__' : cid;
        return `<div class="seg clickable" data-set-component="${filter}" style="width:${w}%; background:${color};" title="${escapeHTML(name)}: ${n} (click to filter)"></div>`;
      }).join('');
      const compLegend = compEntries.slice(0, 6).map(([cid, n]) => {
        const cmp = cid === '__none' ? null : findComponent(proj, cid);
        const c = cmp ? componentColor(cmp.color) : null;
        const color = c ? `rgba(${c.rgb},.95)` : 'var(--neutral)';
        const filter = cid === '__none' ? '__none__' : cid;
        return `<span class="cmp-leg-item clickable" data-set-component="${filter}"><span class="dot" style="background:${color}"></span>${escapeHTML(cmp ? cmp.name : 'Unassigned')} <b>${n}</b></span>`;
      }).join('');

      $('#regKpis').innerHTML = `
        <div class="reg-kpi clickable" data-clear-filters title="Click to clear filters">
          <div class="kpi-num">${total}</div>
          <div class="kpi-lbl">${total === 1 ? 'action' : 'actions'} (filtered)</div>
        </div>
        <div class="reg-kpi donut-kpi">
          <svg class="kpi-donut" viewBox="0 0 64 64">
            <circle cx="${cx}" cy="${cy}" r="${r}" fill="none" stroke="var(--bg-3)" stroke-width="9" />
            ${donutSegs}
            <text x="${cx}" y="${cy + 4}" text-anchor="middle" class="kpi-donut-num">${donePct}%</text>
          </svg>
          <div class="kpi-donut-legend">
            <span class="cmp-leg-item clickable" data-set-status="done"><span class="dot" style="background:var(--ok)"></span>Done <b>${cnt.done}</b></span>
            <span class="cmp-leg-item clickable" data-set-status="doing"><span class="dot" style="background:var(--accent)"></span>Doing <b>${cnt.doing}</b></span>
            <span class="cmp-leg-item clickable" data-set-status="blocked"><span class="dot" style="background:var(--bad)"></span>Blocked <b>${cnt.blocked}</b></span>
            <span class="cmp-leg-item clickable" data-set-status="todo"><span class="dot" style="background:var(--neutral)"></span>Todo <b>${cnt.todo}</b></span>
          </div>
        </div>
        <div class="reg-kpi clickable" data-set-due="late" title="Click to filter to overdue">
          <div class="kpi-num ${overdue > 0 ? 'bad' : ''}">${overdue}</div>
          <div class="kpi-lbl">overdue</div>
          <div class="kpi-sub clickable-inline" data-set-due="week" title="Click to filter to due this week">${soon} due ≤ 7d</div>
        </div>
        <div class="reg-kpi grow">
          <div class="kpi-lbl">By component</div>
          <div class="kpi-stack">${compBars}</div>
          <div class="cmp-legend">${compLegend}</div>
        </div>
        <div class="reg-kpi grow">
          <div class="kpi-lbl">Opened vs completed (30d)</div>
          ${lineChartSVG}
          <div class="cmp-legend">
            <span class="cmp-leg-item"><span class="dot" style="background:rgba(154,161,184,.7)"></span>Open now <b>${todayOpenCount}</b></span>
            <span class="cmp-leg-item"><span class="dot" style="background:var(--accent)"></span>Opened <b>${totalOpened}</b></span>
            <span class="cmp-leg-item"><span class="dot" style="background:var(--ok)"></span>Completed <b>${totalClosed}</b></span>
            <span class="cmp-leg-item"><span class="dot" style="background:var(--text-faint);opacity:.5"></span>today</span>
          </div>
        </div>
        `;

      // Wire KPI click filters (idempotent — replaces previous listener).
      // Single click → set filter. Double click → clear that filter only.
      const kpiPanel = $('#regKpis');
      const handleKpi = (e, mode) => {
        const el = e.target.closest('[data-set-status],[data-set-due],[data-set-component],[data-clear-filters]');
        if (!el) return;
        e.preventDefault();
        if (el.dataset.clearFilters !== undefined) {
          applyTopbarFilter({ clearAll: true });
          return;
        }
        const set = {};
        const v = mode === 'clear' ? '' : undefined;
        if (el.dataset.setStatus !== undefined)   set.status    = v ?? el.dataset.setStatus;
        if (el.dataset.setDue !== undefined)      set.due       = v ?? el.dataset.setDue;
        if (el.dataset.setComponent !== undefined) set.component = v ?? el.dataset.setComponent;
        applyTopbarFilter(set);
      };
      kpiPanel.onclick    = (e) => handleKpi(e, 'set');
      kpiPanel.ondblclick = (e) => handleKpi(e, 'clear');
    }

    function rowHTML(a) {
      const cmp = findComponent(proj, a.component);
      const c = cmp ? componentColor(cmp.color) : null;
      const dueCls = statusOfDue(a.due, a.status);
      const stat = STATUSES.find((s) => s.id === a.status);
      const lvl = priorityLevel(a.priorityLevel);
      // Left edge of the row now reflects priority (not the binary
      // overdue state). The dedicated priority pip column is dropped —
      // the left edge encodes the same signal more scannably.
      const tintProps = [
        c ? `--row-tint: rgba(${c.rgb},.10);` : '',
        `--prio-rgb: ${lvl.rgb};`,
      ].filter(Boolean).join(' ');
      const tint = ` style="${tintProps}"`;
      // Lateness as a 'drag tail' — a thin red trail whose width grows
      // with days late, capped at 30 days so a 90-d-late item doesn't
      // bully the title cell. Replaces the old ⏰ emoji.
      const isOverdue = a.status !== 'done' && a.due && dayDiff(a.due, todayISO()) < 0;
      const lateDays = isOverdue ? Math.abs(dayDiff(a.due, todayISO())) : 0;
      const dragTail = isOverdue
        ? `<span class="reg-late-tail" style="--days:${Math.min(lateDays, 30)};" title="Late by ${lateDays} day${lateDays === 1 ? '' : 's'} · was due ${escapeHTML(fmtDate(a.due))}"><span class="reg-late-icon" aria-hidden="true">⏰</span><span class="reg-late-num">+${lateDays}d</span></span>`
        : '';
      const isDone      = a.status === 'done';
      const isCancelled = a.status === 'cancelled';
      return `
        <div class="reg-row prio-${lvl.id} ${isOverdue ? 'is-overdue' : ''} ${isDone ? 'is-done' : ''} ${isCancelled ? 'is-cancelled' : ''}" data-id="${a.id}"${tint}>
          <div class="reg-cell title-cell">
            ${ROW_GRIP_HTML}
            <input type="text" class="reg-inp title-inp" data-field="title" value="${escapeHTML(a.title)}" />
            ${dragTail}
            ${a.notes ? '<span class="tag" title="Has notes">note</span>' : ''}
          </div>
          <div class="reg-cell">
            <select class="reg-inp reg-comp" data-field="component" ${c ? `style="color:rgb(${c.rgb}); border-color:rgba(${c.rgb},.4);"` : ''}>
              <option value="">— None</option>
              ${(proj.components || []).map((m) => `<option value="${m.id}" ${m.id === a.component ? 'selected' : ''}>${escapeHTML(m.name)}</option>`).join('')}
            </select>
          </div>
          <div class="reg-cell">
            <select class="reg-inp" data-field="owner">
              ${state.people.map((p) => `<option value="${p.id}" ${p.id === a.owner ? 'selected' : ''}>${escapeHTML(p.name)}</option>`).join('')}
            </select>
          </div>
          <div class="reg-cell status-cell">
            <span class="col-dot ${stat?.dot}"></span>
            <select class="reg-inp" data-field="status">
              ${STATUSES.map((s) => `<option value="${s.id}" ${s.id === a.status ? 'selected' : ''}>${s.name}</option>`).join('')}
            </select>
          </div>
          <div class="reg-cell priority-cell">
            <select class="reg-inp" data-field="priorityLevel" style="color:rgb(${lvl.rgb});">
              ${PRIORITY_LEVELS.map((p) => `<option value="${p.id}" ${p.id === (a.priorityLevel || 'med') ? 'selected' : ''}>${p.label}</option>`).join('')}
            </select>
          </div>
          <div class="reg-cell">
            <input type="date" class="reg-inp ${dueCls}" data-field="due" value="${a.due || ''}" />
          </div>
          <div class="reg-cell">
            <input type="date" class="reg-inp ${a.predictedCompletion ? 'overridden' : 'derived'}" data-field="predictedCompletion" value="${predictedCompletion(a)}" title="${a.predictedCompletion ? 'Custom predicted date — click 𝕩 to clear' : 'Defaults to due date'}" />
          </div>
          <div class="reg-cell">
            <input type="date" class="reg-inp" data-field="originatorDate" value="${a.originatorDate || a.createdAt || ''}" title="When this action was originated" />
          </div>
          <div class="reg-cell">
            <button class="row-del" title="Delete action" aria-label="Delete">×</button>
          </div>
        </div>`;
    }

    function drawTable() {
      let acts = (proj.actions || []).filter(actionMatchesFilters).slice();
      // 'manual' sort = preserve project array order (set by drag-reorder)
      if (regState.sortBy !== 'manual') {
        acts.sort((a, b) => {
          const av = regSortValue(a, regState.sortBy, proj);
          const bv = regSortValue(b, regState.sortBy, proj);
          if (av < bv) return regState.sortDir === 'asc' ? -1 : 1;
          if (av > bv) return regState.sortDir === 'asc' ? 1 : -1;
          return 0;
        });
      }
      const body = $('#regBody');
      if (!acts.length) {
        body.innerHTML = '<div class="empty">No actions match the current filters.</div>';
      } else {
        body.innerHTML = acts.map(rowHTML).join('');
      }
      $$('.reg-col', view).forEach((b) => b.classList.remove('asc', 'desc'));
      const active = view.querySelector(`.reg-col[data-col="${regState.sortBy}"]`);
      if (active) active.classList.add(regState.sortDir);
      // Wire drag-to-reorder on the rendered rows
      wireListReorder(body, {
        rowSelector: '.reg-row[data-id]',
        idAttr: 'id',
        getArray: () => proj.actions,
        setOrder: (visibleIds) => {
          // The visibleIds are only currently-displayed rows; rebuild proj.actions
          // by interleaving the new order back into the full array (preserving
          // the relative order of off-screen / filtered-out actions).
          const visibleSet = new Set(visibleIds);
          const queue = visibleIds.slice();
          const next = [];
          (proj.actions || []).forEach((a) => {
            if (visibleSet.has(a.id)) {
              const id = queue.shift();
              const item = (proj.actions || []).find((x) => x.id === id);
              if (item) next.push(item);
            } else {
              next.push(a);
            }
          });
          proj.actions = next;
          // Switch to manual sort so the new order is preserved across re-renders
          regState.sortBy = 'manual';
          regState.sortDir = 'asc';
        },
        commitName: 'register-reorder',
      });
    }

    // Snapshot + lightweight save: don't full-render the view (would lose
    // input focus and table scroll) — just save and refresh KPIs/row tint.
    function snapshotForUndo() {
      undoStack.push(JSON.stringify(state));
      if (undoStack.length > HISTORY_LIMIT) undoStack.shift();
      redoStack = [];
    }
    function applyEdit(id, field, raw) {
      const a = proj.actions.find((x) => x.id === id);
      if (!a) return;
      let changed = false;
      const today = todayISO();
      if (field === 'title') {
        const v = String(raw).trim();
        if (v && v !== a.title) { a.title = v; changed = true; }
      } else if (field === 'component') {
        const v = raw || null;
        if (v !== a.component) { a.component = v; changed = true; }
      } else if (field === 'owner') {
        const v = String(raw);
        if (v && v !== a.owner) {
          a.history.push({ at: today, what: `Owner: ${personName(a.owner)} → ${personName(v)}` });
          a.owner = v; changed = true;
        }
      } else if (field === 'status') {
        const v = String(raw);
        if (v && v !== a.status) {
          a.history.push({ at: today, what: `Status: ${a.status} → ${v}` });
          a.status = v; changed = true;
        }
      } else if (field === 'due') {
        const v = raw || null;
        if (v !== a.due) {
          a.history.push({ at: today, what: `Due: ${a.due || '—'} → ${v || '—'}` });
          a.due = v; changed = true;
        }
      } else if (field === 'commitment') {
        const v = clamp(parseInt(raw, 10) || 100, 5, 100);
        if (v !== a.commitment) { a.commitment = v; changed = true; }
      } else if (field === 'predictedCompletion') {
        const v = raw || null;
        if ((a.predictedCompletion || null) !== v) { a.predictedCompletion = v; changed = true; }
      } else if (field === 'actualCompletion') {
        const v = raw || null;
        if ((a.actualCompletion || null) !== v) { a.actualCompletion = v; changed = true; }
      } else if (field === 'priorityLevel') {
        const v = PRIORITY_LEVELS.some((p) => p.id === raw) ? raw : 'med';
        if ((a.priorityLevel || 'med') !== v) {
          a.history = a.history || [];
          a.history.push({ at: today, what: `Priority: ${priorityLevel(a.priorityLevel || 'med').label} → ${priorityLevel(v).label}` });
          a.priorityLevel = v;
          changed = true;
        }
      } else if (field === 'originatorDate') {
        const v = raw || null;
        if ((a.originatorDate || null) !== v) {
          a.history = a.history || [];
          a.history.push({ at: today, what: `Originator date: ${a.originatorDate || '—'} → ${v || '—'}` });
          a.originatorDate = v;
          changed = true;
        }
      }
      if (changed) {
        a.updatedAt = today;
        snapshotForUndo();
        saveState();
        drawKpis();
      }
    }

    $$('.reg-col', view).forEach((btn) => {
      btn.addEventListener('click', () => {
        const col = btn.dataset.col;
        if (regState.sortBy === col) regState.sortDir = regState.sortDir === 'asc' ? 'desc' : 'asc';
        else { regState.sortBy = col; regState.sortDir = 'asc'; }
        drawTable();
      });
    });

    // Event delegation on the body for inline edits and row deletion
    const body = $('#regBody');
    body.addEventListener('change', (e) => {
      const inp = e.target.closest('.reg-inp');
      if (!inp) return;
      const row = inp.closest('.reg-row');
      const id = row?.dataset.id;
      if (!id) return;
      applyEdit(id, inp.dataset.field, inp.value);
      // Refresh just this row so swatches/colors/late state update
      const a = (curProject().actions || []).find((x) => x.id === id);
      if (a) {
        const newRow = document.createElement('div');
        newRow.innerHTML = rowHTML(a).trim();
        row.replaceWith(newRow.firstChild);
      }
    });
    body.addEventListener('click', (e) => {
      const del = e.target.closest('.row-del');
      if (!del) return;
      const row = del.closest('.reg-row');
      const id = row?.dataset.id;
      const a = (curProject().actions || []).find((x) => x.id === id);
      if (!a) return;
      if (!confirm(`Move "${a.title}" to Archive? You can restore it later.`)) return;
      const sourceProj = projectOfAction(id) || proj;
      const target = (sourceProj.actions || []).find((x) => x.id === id);
      if (target) {
        target.deletedAt = todayISO();
        target.history.push({ at: todayISO(), what: 'Moved to Archive' });
        target.updatedAt = todayISO();
      }
      snapshotForUndo();
      saveState();
      drawKpis();
      row.remove();
    });
    // Right-click on a register row → action context menu
    body.addEventListener('contextmenu', (e) => {
      const row = e.target.closest('.reg-row[data-id]');
      if (!row) return;
      e.preventDefault();
      const id = row.dataset.id;
      const a = (curProject().actions || []).find((x) => x.id === id);
      if (!a) return;
      showContextMenu(e.clientX, e.clientY, actionContextItems(a));
    });

    $('#btnAddAction').addEventListener('click', () => openQuickAdd('action'));
    $('#btnOpenArchive').addEventListener('click', () => {
      state.currentView = 'archive';
      saveState(); render();
    });

    // Archive every currently-visible done action (respecting active filters)
    // → they disappear from the register, remain visible in the Archive panel.
    $('#btnArchiveDone').addEventListener('click', () => {
      const proj2 = curProject();
      const candidates = (proj2.actions || []).filter((a) => !a.deletedAt && a.status === 'done' && actionMatchesFilters(a));
      if (!candidates.length) { toast('No done actions match the current filters'); return; }
      if (!confirm(`Archive ${candidates.length} done action${candidates.length === 1 ? '' : 's'}? They will move to the Archive panel and disappear from the register.`)) return;
      const today = todayISO();
      candidates.forEach((a) => {
        a.deletedAt = today;
        a.history = a.history || [];
        a.history.push({ at: today, what: 'Moved to Archive (bulk-archive done from Register)' });
        a.updatedAt = today;
      });
      commit('archive-done-bulk');
      toast(`${candidates.length} archived`);
    });

    drawKpis();
    drawTable();
  }

  /* ---------------------------- Timeline ----------------------------- */

  let tlState = null;
  const GRANULARITIES = {
    day:   { name: 'Day',   defaultDw: 24, minDw: 14, maxDw: 80, snapDays: 1,  totalDays: 180 },
    week:  { name: 'Week',  defaultDw: 8,  minDw: 4,  maxDw: 24, snapDays: 7,  totalDays: 365 },
    month: { name: 'Month', defaultDw: 3,  minDw: 1,  maxDw: 8,  snapDays: 30, totalDays: 730 },
  };

  function renderTimeline(root) {
    const proj = curProject();
    const acts = (proj.actions || []).filter(actionMatchesFilters);
    const view = document.createElement('div');
    view.className = 'view';
    view.innerHTML = `
      <div class="page-head">
        <div>
          <div class="page-title">${escapeHTML(proj.name)} — Timeline</div>
          <div class="page-sub">Drag bar body to shift • Drag edges to resize • Drag vertically to reassign • Double-click empty space to create</div>
        </div>
        <div class="page-actions">
          <div class="seg" role="tablist" aria-label="Granularity">
            <button class="seg-btn" data-gran="day">Day</button>
            <button class="seg-btn" data-gran="week">Week</button>
            <button class="seg-btn" data-gran="month">Month</button>
          </div>
          <div class="popover-anchor">
            <button class="ghost" id="btnTLHolidays" title="Choose national holidays to consider">Holidays <span id="holidayBadge" class="badge"></span></button>
            <div class="popover" id="holidayPopover" hidden></div>
          </div>
          <button class="icon-btn" id="btnTLZoomOut" title="Zoom out">−</button>
          <button class="icon-btn" id="btnTLZoomIn" title="Zoom in">+</button>
          <button class="ghost" id="btnTLToday">Today</button>
        </div>
      </div>
      <div class="timeline">
        <div class="tl-lanes" id="tlLanes"></div>
        <div class="tl-grid-wrap" id="tlGridWrap">
          <div class="tl-axis" id="tlAxis"></div>
          <div class="tl-grid" id="tlGrid"></div>
          <div class="tl-tooltip" id="tlTooltip" hidden></div>
        </div>
      </div>`;
    root.appendChild(view);

    if (!tlState) tlState = { granularity: 'day', dayWidth: GRANULARITIES.day.defaultDw, startOffsetDays: -14 };
    if (!state.settings) state.settings = { holidayCountries: [] };

    // Mark active granularity
    $$('.seg-btn', view).forEach((b) => {
      b.classList.toggle('active', b.dataset.gran === tlState.granularity);
      b.addEventListener('click', () => {
        if (tlState.granularity === b.dataset.gran) return;
        tlState.granularity = b.dataset.gran;
        tlState.dayWidth = GRANULARITIES[tlState.granularity].defaultDw;
        // re-mark
        $$('.seg-btn', view).forEach((x) => x.classList.toggle('active', x.dataset.gran === tlState.granularity));
        drawTimeline(acts);
        scrollToToday();
      });
    });

    // Holiday popover
    updateHolidayBadge();
    $('#btnTLHolidays').addEventListener('click', (e) => {
      e.stopPropagation();
      toggleHolidayPopover(acts);
    });
    // Outside-click closer is registered once globally (see init()).

    drawTimeline(acts);
    // Scroll to today on first paint
    setTimeout(scrollToToday, 0);

    $('#btnTLZoomIn').addEventListener('click', () => {
      const g = GRANULARITIES[tlState.granularity];
      tlState.dayWidth = clamp(tlState.dayWidth * 1.35, g.minDw, g.maxDw);
      drawTimeline(acts);
    });
    $('#btnTLZoomOut').addEventListener('click', () => {
      const g = GRANULARITIES[tlState.granularity];
      tlState.dayWidth = clamp(tlState.dayWidth / 1.35, g.minDw, g.maxDw);
      drawTimeline(acts);
    });
    $('#btnTLToday').addEventListener('click', scrollToToday);
  }

  function scrollToToday() {
    const wrap = $('#tlGridWrap');
    if (!wrap) return;
    const todayX = -tlState.startOffsetDays * tlState.dayWidth;
    wrap.scrollLeft = Math.max(0, todayX - wrap.clientWidth / 3);
  }

  function updateHolidayBadge() {
    const badge = $('#holidayBadge');
    if (!badge) return;
    const codes = activeHolidayCodes();
    badge.textContent = codes.length ? String(codes.length) : '';
    badge.style.display = codes.length ? '' : 'none';
  }

  function toggleHolidayPopover(acts) {
    const pop = $('#holidayPopover');
    if (!pop) return;
    if (!pop.hidden) { pop.hidden = true; return; }
    const codes = activeHolidayCodes();
    pop.innerHTML = `
      <div class="popover-head">Non-working days</div>
      <div class="popover-sub">Selected days appear as bands on the Gantt and warn when a task lands on one.</div>
      <div class="popover-list">
        <label class="popover-item">
          <input type="checkbox" data-code="WKND" ${codes.includes('WKND') ? 'checked' : ''} />
          <span class="flag">🛌</span>
          <span>Weekends</span>
          <span class="muted">Sat &amp; Sun</span>
        </label>
        <div class="popover-divider"></div>
        ${Object.entries(HOLIDAYS).map(([code, h]) => `
          <label class="popover-item">
            <input type="checkbox" data-code="${code}" ${codes.includes(code) ? 'checked' : ''} />
            <span class="flag">${h.flag}</span>
            <span>${escapeHTML(h.name)}</span>
            <span class="muted">${Object.keys(h.dates).length}</span>
          </label>`).join('')}
      </div>
      <div class="popover-foot">
        <button class="ghost" id="holNone">Clear</button>
        <button class="ghost" id="holClose">Done</button>
      </div>`;
    pop.hidden = false;

    $$('input[type=checkbox]', pop).forEach((cb) => {
      cb.addEventListener('change', () => {
        const code = cb.dataset.code;
        let codes = activeHolidayCodes().slice();
        if (cb.checked && !codes.includes(code)) codes.push(code);
        if (!cb.checked) codes = codes.filter((c) => c !== code);
        state.settings = state.settings || {};
        state.settings.holidayCountries = codes;
        saveState();
        updateHolidayBadge();
        drawTimeline(acts);
      });
    });
    $('#holNone').addEventListener('click', () => {
      state.settings.holidayCountries = [];
      saveState();
      updateHolidayBadge();
      drawTimeline(acts);
      $$('input[type=checkbox]', pop).forEach((cb) => cb.checked = false);
    });
    $('#holClose').addEventListener('click', () => { pop.hidden = true; });
  }

  function drawAxis(axisEl, start, totalDays, dw, granularity) {
    axisEl.innerHTML = '';
    axisEl.style.width = (totalDays * dw) + 'px';

    const addTick = (d, isMajor) => {
      const tick = document.createElement('div');
      tick.className = 'tl-axis-tick' + (isMajor ? ' major' : '');
      tick.style.left = (d * dw) + 'px';
      axisEl.appendChild(tick);
    };
    const addLabel = (d, text, sub = '') => {
      const lbl = document.createElement('div');
      lbl.className = 'tl-axis-label';
      lbl.style.left = (d * dw) + 'px';
      lbl.innerHTML = sub ? `<b>${escapeHTML(text)}</b><span>${escapeHTML(sub)}</span>` : escapeHTML(text);
      axisEl.appendChild(lbl);
    };

    if (granularity === 'day') {
      const labelEvery = dw >= 30 ? 1 : dw >= 18 ? 2 : 7;
      for (let d = 0; d <= totalDays; d++) {
        const date = new Date(start.getTime() + d * dayMs);
        const isMonthStart = date.getDate() === 1;
        const isWeek = date.getDay() === 1; // Mondays
        if (isMonthStart || isWeek) addTick(d, isMonthStart);
        else if (dw >= 14) addTick(d, false);
        if (d % labelEvery === 0 || isMonthStart) {
          const text = date.toLocaleDateString(undefined, { month: 'short', day: 'numeric' });
          addLabel(d, text);
        }
      }
    } else if (granularity === 'week') {
      // First Monday on or after start
      let dt = new Date(start);
      while (dt.getDay() !== 1) dt = new Date(dt.getTime() + dayMs);
      let d = Math.round((dt - start) / dayMs);
      // Major ticks at month starts; minor at every Monday
      while (d <= totalDays) {
        const major = dt.getDate() <= 7; // first Monday of the month
        addTick(d, major);
        if (major) addLabel(d, dt.toLocaleDateString(undefined, { month: 'short' }), String(dt.getFullYear()));
        else if (dw >= 16) addLabel(d, dt.toLocaleDateString(undefined, { day: 'numeric' }));
        dt = new Date(dt.getTime() + 7 * dayMs);
        d += 7;
      }
    } else if (granularity === 'month') {
      // First day of month
      let dt = new Date(start);
      if (dt.getDate() !== 1) dt = new Date(dt.getFullYear(), dt.getMonth() + 1, 1);
      while (true) {
        const d = Math.round((dt - start) / dayMs);
        if (d > totalDays) break;
        const major = dt.getMonth() % 3 === 0; // quarter starts
        addTick(d, major);
        addLabel(d, dt.toLocaleDateString(undefined, { month: 'short' }), String(dt.getFullYear()));
        dt = new Date(dt.getFullYear(), dt.getMonth() + 1, 1);
      }
    }
  }

  function drawTimeline(actions) {
    const lanesEl = $('#tlLanes');
    const axisEl = $('#tlAxis');
    const gridEl = $('#tlGrid');
    if (!lanesEl) return;

    const g = GRANULARITIES[tlState.granularity];
    const totalDays = g.totalDays;
    const dw = tlState.dayWidth;
    const width = totalDays * dw;
    // Align start to the granularity unit
    let start = new Date(Date.now() + tlState.startOffsetDays * dayMs);
    start.setHours(0, 0, 0, 0);
    if (tlState.granularity === 'week') {
      while (start.getDay() !== 1) start = new Date(start.getTime() - dayMs);
    } else if (tlState.granularity === 'month') {
      start = new Date(start.getFullYear(), start.getMonth(), 1);
    }
    const startISO = fmtISO(start);

    // Build lanes by person + a bottom events cell to align with the strip
    const people = state.people;
    lanesEl.innerHTML = `<div class="tl-lane header">Owner</div>` +
      people.map((p) => `<div class="tl-lane" data-owner="${p.id}">${escapeHTML(p.name)}</div>`).join('') +
      `<div class="tl-lane events-lane">Events</div>`;

    // Axis
    drawAxis(axisEl, start, totalDays, dw, tlState.granularity);

    // Grid rows per person
    gridEl.innerHTML = '';
    gridEl.style.width = width + 'px';
    people.forEach((p) => {
      const row = document.createElement('div');
      row.className = 'tl-grid-row';
      row.dataset.owner = p.id;
      gridEl.appendChild(row);
    });

    // Vertical gridlines on grid (mirror axis ticks for readability)
    const gridLines = document.createElement('div');
    gridLines.className = 'tl-gridlines';
    gridLines.style.width = width + 'px';
    drawAxis(gridLines, start, totalDays, dw, tlState.granularity);
    gridEl.appendChild(gridLines);

    // Holiday & weekend bands (drawn before today line so it sits on top)
    if (activeHolidayCodes().length) {
      const seen = {};
      for (let d = 0; d <= totalDays; d++) {
        const date = new Date(start.getTime() + d * dayMs);
        const iso = fmtISO(date);
        const info = holidayInfo(iso);
        if (!info) continue;
        if (seen[iso]) continue;
        seen[iso] = true;
        const onlyWeekend = info.every((h) => h.isWeekend);
        const band = document.createElement('div');
        band.className = 'tl-holiday-band' + (onlyWeekend ? ' weekend' : '');
        band.style.left = (d * dw) + 'px';
        band.style.width = dw + 'px';
        band.title = info.map((h) => h.isWeekend ? `${h.flag} ${h.name}` : `${h.flag} ${h.name} (${h.country})`).join('\n');
        // Tag (flag emoji) only on national holidays in day mode at generous zoom
        if (!onlyWeekend && tlState.granularity === 'day' && dw >= 18) {
          const nat = info.find((h) => !h.isWeekend);
          if (nat) {
            const tag = document.createElement('span');
            tag.className = 'tl-holiday-tag';
            tag.textContent = nat.flag;
            band.appendChild(tag);
          }
        }
        gridEl.appendChild(band);
      }
    }

    // Compute per-owner per-day commitment load. An action contributes its
    // commitment % to every day in its [startDate..due] window — done items
    // do not count toward future load.
    const ownerLoads = {};      // ownerId -> { iso -> totalCommitment }
    const actOverloads = {};    // actionId -> { badDays, peak }
    actions.forEach((a) => {
      if (!a.due || a.status === 'done') return;
      const cmt = (typeof a.commitment === 'number') ? a.commitment : 100;
      const sIso = a.startDate || fmtISO(new Date(parseDate(a.due).getTime() - 2 * dayMs));
      const sT = parseDate(sIso).getTime();
      const eT = parseDate(a.due).getTime();
      if (!ownerLoads[a.owner]) ownerLoads[a.owner] = {};
      for (let t = sT; t <= eT; t += dayMs) {
        const iso = fmtISO(new Date(t));
        ownerLoads[a.owner][iso] = (ownerLoads[a.owner][iso] || 0) + cmt;
      }
    });
    actions.forEach((a) => {
      if (!a.due || a.status === 'done') return;
      const sIso = a.startDate || fmtISO(new Date(parseDate(a.due).getTime() - 2 * dayMs));
      const sT = parseDate(sIso).getTime();
      const eT = parseDate(a.due).getTime();
      let bad = 0, peak = 0;
      for (let t = sT; t <= eT; t += dayMs) {
        const iso = fmtISO(new Date(t));
        const v = ownerLoads[a.owner]?.[iso] || 0;
        if (v > 100) { bad++; if (v > peak) peak = v; }
      }
      if (bad) actOverloads[a.id] = { badDays: bad, peak: Math.round(peak) };
    });

    // Render per-owner over-allocation overlays — one band per day where load > 100%
    people.forEach((p, pi) => {
      const ld = ownerLoads[p.id];
      if (!ld) return;
      const row = gridEl.children[pi];
      for (let dy = 0; dy <= totalDays; dy++) {
        const date = new Date(start.getTime() + dy * dayMs);
        const iso = fmtISO(date);
        const v = ld[iso];
        if (!v || v <= 100) continue;
        const overlay = document.createElement('div');
        overlay.className = 'tl-overload';
        overlay.style.left = (dy * dw) + 'px';
        overlay.style.width = dw + 'px';
        const intensity = Math.min(0.45, 0.15 + (v - 100) / 300);
        overlay.style.background = `rgba(248,113,113,${intensity.toFixed(2)})`;
        overlay.title = `${p.name}: ${Math.round(v)}% load on ${fmtFull(iso)}`;
        row.appendChild(overlay);
      }
    });

    // Today line
    const today = new Date(); today.setHours(0,0,0,0);
    const todayOffset = Math.round((today - start) / dayMs) * dw;
    const todayLine = document.createElement('div');
    todayLine.className = 'tl-today';
    todayLine.style.left = todayOffset + 'px';
    gridEl.appendChild(todayLine);

    // Build the unified events list (milestones / deliverables / meetings)
    // and render vertical lines spanning the whole grid + a label strip
    // pinned to the bottom of the grid to avoid clutter.
    const proj = curProject();
    const events = [];
    (proj.milestones || []).forEach((m) => {
      if (m.date) events.push({ kind: 'milestone', date: m.date, name: m.name, status: m.status, sub: 'Milestone' });
    });
    (proj.deliverables || []).forEach((dv) => {
      if (dv.dueDate) events.push({ kind: 'deliverable', date: dv.dueDate, name: dv.name, status: dv.status, sub: 'Deliverable' });
    });
    (proj.meetings || []).forEach((mt) => {
      const winStartISO = fmtISO(start);
      const winEndISO   = fmtISO(new Date(start.getTime() + totalDays * dayMs));
      const dates = expandMeetingDates(mt, winStartISO, winEndISO);
      if (!dates.length) return;
      const baseLabel = meetingRecurrenceLabel(mt);
      const sub = baseLabel + (mt.time ? ' ' + mt.time : '');
      const isRecurring = mt.kind === 'recurring';
      dates.forEach((iso, i) => {
        events.push({
          kind: isRecurring ? 'meeting-weekly' : 'meeting', // legacy class name kept for CSS; covers all recurrence
          date: iso,
          name: mt.title,
          sub,
          isFirst: i === 0,
        });
      });
    });

    // Filter to visible window and sort
    const winStart = start.getTime();
    const winEnd = winStart + totalDays * dayMs;
    const visibleEvents = events.filter((e) => {
      const t = parseDate(e.date).getTime();
      return t >= winStart && t <= winEnd;
    }).sort((a, b) => a.date.localeCompare(b.date));

    // Vertical line per event (spans whole grid including the events row)
    visibleEvents.forEach((e) => {
      const offset = Math.round((parseDate(e.date) - start) / dayMs);
      const x = offset * dw;
      const line = document.createElement('div');
      line.className = `tl-event-line tl-evt-${e.kind}`;
      line.style.left = x + 'px';
      gridEl.appendChild(line);
    });

    // Events row at the bottom of the grid (inside tl-grid, so vertical
    // lines naturally extend down to it). Markers stack to avoid overlap.
    const eventsRow = document.createElement('div');
    eventsRow.className = 'tl-events-row';
    gridEl.appendChild(eventsRow);

    // Place markers, alternating top/bottom slots and skipping repeated weekly
    // labels (only the first occurrence gets a label) to keep the strip clean.
    let slotIdx = 0;
    visibleEvents.forEach((e) => {
      const offset = Math.round((parseDate(e.date) - start) / dayMs);
      const x = offset * dw;
      const labelSuppress = e.kind === 'meeting-weekly' && !e.isFirst;
      if (labelSuppress) return;
      const slot = slotIdx % 3; // 3 vertical lanes inside the strip
      slotIdx++;
      const m = document.createElement('div');
      m.className = `tl-event-mark tl-evt-${e.kind} slot-${slot}`;
      m.style.left = x + 'px';
      const icon = e.kind === 'milestone' ? '◇'
                 : e.kind === 'deliverable' ? '◆'
                 : '⊕';
      m.title = `${e.sub}: ${e.name} — ${fmtFull(e.date)}`;
      m.innerHTML = `<span class="evt-icon">${icon}</span><span class="evt-name">${escapeHTML(e.name)}</span>`;
      eventsRow.appendChild(m);
    });

    // Critical-path detection — for every open milestone, walk back via
    // dependsOn to find the longest unbroken chain. Every action on that
    // chain gets a class on its bar; it also includes the actions linked
    // to the milestone via a.milestone (since the milestone date is the
    // chain's endpoint).
    const criticalSet = computeCriticalActions(proj);

    // Draw bars
    actions.forEach((a) => {
      if (!a.due) return;
      const ownerIdx = people.findIndex((p) => p.id === a.owner);
      if (ownerIdx < 0) return;
      // Default bar duration: 3 days ending at due, unless startDate is set
      const defaultDur = 3;
      const endD = a.due;
      const startD = a.startDate ||
        fmtISO(new Date(parseDate(endD).getTime() - (defaultDur - 1) * dayMs));
      const startOffset = Math.round((parseDate(startD) - start) / dayMs);
      const endOffset = Math.round((parseDate(endD) - start) / dayMs);
      const barLen = Math.max(1, endOffset - startOffset + 1);
      const left = startOffset * dw;
      const w = Math.max(g.minDw, barLen * dw - 2);

      const bar = document.createElement('div');
      const dueCls = statusOfDue(a.due, a.status);
      const holDue = holidayInfo(a.due);
      const holStart = holidayInfo(startD);
      const onHoliday = (holDue || holStart) && a.status !== 'done';
      const overload = actOverloads[a.id];
      const cmt = (typeof a.commitment === 'number') ? a.commitment : 100;
      bar.className = `tl-bar ${a.status} ${dueCls === 'late' && a.status !== 'done' ? 'late' : ''} ${onHoliday ? 'holiday-conflict' : ''} ${overload ? 'over-allocated' : ''} ${criticalSet.has(a.id) ? 'critical-path' : ''}`;
      bar.style.left = left + 'px';
      bar.style.width = w + 'px';
      bar.dataset.id = a.id;
      const ownerName = personName(a.owner);
      const holNote = holDue ? `\n⚠ Due on holiday: ${holDue.map((h) => h.name).join(', ')}` : '';
      const overNote = overload ? `\n⚡ Over-allocated: ${ownerName} reaches ${overload.peak}% on ${overload.badDays} day${overload.badDays === 1 ? '' : 's'}` : '';
      const cmtNote = cmt !== 100 ? `\nCommitment: ${cmt}%` : '';
      bar.title = `${a.title}\n${fmtFull(startD)} → ${fmtFull(endD)} (${barLen}d)${cmtNote}${holNote}${overNote}`;
      const warnIcons = (onHoliday ? '<span class="bar-warn hol" title="Falls on a non-working day">⚠</span>' : '')
        + (overload ? `<span class="bar-warn over" title="Owner over-allocated">⚡</span>` : '');
      const cmtChip = cmt !== 100 ? `<span class="bar-pct">${cmt}%</span>` : '';
      bar.innerHTML = `
        <div class="resize-handle left" title="Drag to change start"></div>
        <span class="bar-label">${warnIcons}${cmtChip ? ' ' + cmtChip : ''} ${escapeHTML(a.title)}</span>
        <div class="resize-handle right" title="Drag to change end"></div>`;
      attachBarDND(bar, a, startISO);
      const row = gridEl.children[ownerIdx];
      row.appendChild(bar);
      // If the bar's usable interior is too narrow to legibly hold the title,
      // render an overflow label to the right of the bar. Reserve space for the
      // resize handles (8 px each) and the side padding (6 px each), plus any
      // warn icons / commitment chip glued to the front of the label.
      const HANDLES_PAD = 8 + 8 + 6 + 6; // 28 px reserved by handles + padding
      const PREFIX_PX = (warnIcons ? 14 : 0) + (cmtChip ? 30 : 0);
      const titlePx = a.title.length * 6.2 + 4; // estimate (font 11, weight 400)
      const usable = w - HANDLES_PAD - PREFIX_PX;
      if (usable < titlePx) {
        const overflow = document.createElement('span');
        overflow.className = `tl-bar-overflow ${a.status}`;
        overflow.style.left = (left + w + 4) + 'px';
        overflow.textContent = a.title;
        overflow.title = `${a.title} — due ${fmtFull(a.due)}`;
        row.appendChild(overflow);
      }
    });

    // Double-click empty space creates an action
    [...gridEl.querySelectorAll('.tl-grid-row')].forEach((row) => {
      row.addEventListener('dblclick', (e) => {
        if (e.target !== row) return;
        const x = e.offsetX;
        const snapDays = GRANULARITIES[tlState.granularity].snapDays;
        const dayOffset = Math.round(x / dw / snapDays) * snapDays;
        const date = fmtISO(new Date(start.getTime() + dayOffset * dayMs));
        const owner = row.dataset.owner;
        const proj = curProject();
        const a = {
          id: uid('a'), title: 'New action', owner, due: date,
          startDate: date, status: 'todo', priority: 0,
          deliverable: null, milestone: null, notes: '',
          createdAt: todayISO(), updatedAt: todayISO(),
          originatorDate: todayISO(),
          history: [{ at: todayISO(), what: 'Created from timeline' }],
        };
        proj.actions.push(a);
        commit('create');
        toast('Action created — double-click bar to edit');
      });
    });

    // Phase F — render dependency arrows on top of all bars. Drawn last so
    // they sit above the grid lines, but pointer-events: none so bars stay
    // interactive.
    drawDependencyArrows(actions, gridEl, criticalSet);
  }

  // Walk back from every open milestone through dependsOn to flag the
  // critical actions. Returns a Set of action ids on the longest chain to
  // each open milestone (plus the actions directly linked to it).
  function computeCriticalActions(proj) {
    const acts = (proj.actions || []).filter((a) => !a.deletedAt);
    const byId = new Map(acts.map((a) => [a.id, a]));
    const critical = new Set();
    const depths = new Map();
    const onStack = new Set();
    function depth(id) {
      if (depths.has(id)) return depths.get(id);
      if (onStack.has(id)) return 0; // cycle guard
      onStack.add(id);
      const a = byId.get(id);
      if (!a) { onStack.delete(id); return 0; }
      const deps = (a.dependsOn || []).filter((d) => byId.has(d));
      let best = 0;
      deps.forEach((d) => { best = Math.max(best, depth(d)); });
      onStack.delete(id);
      depths.set(id, 1 + best);
      return 1 + best;
    }
    function chain(id) {
      const a = byId.get(id);
      if (!a || critical.has(id)) return;
      critical.add(id);
      const deps = (a.dependsOn || []).filter((d) => byId.has(d));
      if (!deps.length) return;
      let bestDep = null, bestVal = -1;
      deps.forEach((d) => { const v = depth(d); if (v > bestVal) { bestVal = v; bestDep = d; } });
      if (bestDep) chain(bestDep);
    }
    (proj.milestones || []).forEach((m) => {
      if (m.done || m.status === 'done') return;
      const targets = acts.filter((a) => a.milestone === m.id && !isClosedStatus(a.status));
      targets.forEach((t) => { depth(t.id); chain(t.id); });
    });
    return critical;
  }

  function drawDependencyArrows(actions, gridEl, criticalSet) {
    const bars = Array.from(gridEl.querySelectorAll('.tl-bar[data-id]'));
    if (!bars.length) return;
    const byId = new Map(actions.map((a) => [a.id, a]));
    const barById = new Map(bars.map((b) => [b.dataset.id, b]));

    // Need positions relative to gridEl
    const gridRect = gridEl.getBoundingClientRect();
    function rectOf(b) {
      const r = b.getBoundingClientRect();
      return {
        x: r.left - gridRect.left + gridEl.scrollLeft,
        y: r.top - gridRect.top + gridEl.scrollTop,
        w: r.width,
        h: r.height,
      };
    }
    const pairs = [];
    actions.forEach((a) => {
      if (!a.due) return;
      const deps = a.dependsOn || [];
      if (!deps.length) return;
      const target = barById.get(a.id);
      if (!target) return;
      deps.forEach((depId) => {
        const src = byId.get(depId);
        const srcBar = barById.get(depId);
        if (!src || !srcBar) return;
        pairs.push({ src, srcBar, dst: a, dstBar: target });
      });
    });
    if (!pairs.length) return;

    // SVG overlay sized to the grid's full content
    const w = gridEl.scrollWidth || gridEl.clientWidth;
    const h = gridEl.scrollHeight || gridEl.clientHeight;
    const NS = 'http://www.w3.org/2000/svg';
    const svg = document.createElementNS(NS, 'svg');
    svg.setAttribute('class', 'tl-dep-svg');
    svg.setAttribute('width', String(w));
    svg.setAttribute('height', String(h));
    svg.setAttribute('viewBox', `0 0 ${w} ${h}`);
    svg.style.position = 'absolute';
    svg.style.left = '0';
    svg.style.top = '0';
    svg.style.pointerEvents = 'none';

    const defs = document.createElementNS(NS, 'defs');
    defs.innerHTML = `
      <marker id="depArrow" viewBox="0 0 10 10" refX="9" refY="5" markerWidth="7" markerHeight="7" orient="auto">
        <path d="M0,0 L10,5 L0,10 z" fill="currentColor" />
      </marker>
      <marker id="depArrowCrit" viewBox="0 0 10 10" refX="9" refY="5" markerWidth="7" markerHeight="7" orient="auto">
        <path d="M0,0 L10,5 L0,10 z" fill="#f87171" />
      </marker>`;
    svg.appendChild(defs);

    pairs.forEach(({ src, dst, srcBar, dstBar }) => {
      const a = rectOf(srcBar);
      const b = rectOf(dstBar);
      const x1 = a.x + a.w;
      const y1 = a.y + a.h / 2;
      const x2 = b.x;
      const y2 = b.y + b.h / 2;
      const isCritical = criticalSet.has(src.id) && criticalSet.has(dst.id);
      const path = document.createElementNS(NS, 'path');
      // Cubic bezier — a small horizontal lead-out and lead-in so connectors
      // look intentional and don't clip into bar edges.
      const handle = Math.max(12, Math.min(60, (x2 - x1) / 2));
      const d = `M ${x1} ${y1} C ${x1 + handle} ${y1}, ${x2 - handle} ${y2}, ${x2 - 4} ${y2}`;
      path.setAttribute('d', d);
      path.setAttribute('fill', 'none');
      path.setAttribute('class', 'tl-dep-arrow' + (isCritical ? ' critical' : ''));
      path.setAttribute('marker-end', isCritical ? 'url(#depArrowCrit)' : 'url(#depArrow)');
      svg.appendChild(path);
    });
    gridEl.appendChild(svg);
  }

  function showTLTooltip(html, clientX, clientY) {
    const tip = $('#tlTooltip');
    if (!tip) return;
    tip.innerHTML = html;
    tip.hidden = false;
    const wrap = $('#tlGridWrap');
    const wrapRect = wrap.getBoundingClientRect();
    const x = clientX - wrapRect.left + wrap.scrollLeft + 12;
    const y = clientY - wrapRect.top + wrap.scrollTop - 30;
    tip.style.left = x + 'px';
    tip.style.top = Math.max(4, y) + 'px';
  }
  function hideTLTooltip() {
    const tip = $('#tlTooltip');
    if (tip) tip.hidden = true;
  }

  function attachBarDND(bar, action, startISO) {
    let mode = null; // 'move' | 'resize-l' | 'resize-r'
    let originX = 0;
    // Track everything in days from startISO, so granularity/zoom never warps the bar.
    let baseStart = 0, baseLen = 1, baseOwner = null; // captured at mousedown
    let curStart = 0, curLen = 1, curOwner = null;    // committed visually during drag

    const dw = () => tlState.dayWidth;
    const snap = () => GRANULARITIES[tlState.granularity].snapDays;
    const minDw = () => GRANULARITIES[tlState.granularity].minDw;

    function isoOf(dayOffset) {
      return fmtISO(new Date(parseDate(startISO).getTime() + dayOffset * dayMs));
    }
    function applyVisual() {
      bar.style.left = (curStart * dw()) + 'px';
      bar.style.width = Math.max(minDw(), curLen * dw() - 2) + 'px';
    }
    // Snap the drag DELTA (not the absolute position) so misaligned bars
    // keep their original day-of-week / day-of-month while moving in whole units.
    function snapDeltaDays(dx) {
      return Math.round(dx / dw() / snap()) * snap();
    }

    bar.addEventListener('mousedown', (e) => {
      e.preventDefault();
      e.stopPropagation();
      const handle = e.target.classList.contains('resize-handle');
      if (handle) mode = e.target.classList.contains('left') ? 'resize-l' : 'resize-r';
      else mode = 'move';
      originX = e.clientX;

      const proj = curProject();
      const a = proj.actions.find((x) => x.id === action.id);
      if (!a) { mode = null; return; }
      const defaultDur = 3;
      const endD = a.due;
      const startD = a.startDate ||
        fmtISO(new Date(parseDate(endD).getTime() - (defaultDur - 1) * dayMs));
      baseStart = Math.round((parseDate(startD) - parseDate(startISO)) / dayMs);
      baseLen = Math.max(1, Math.round((parseDate(endD) - parseDate(startD)) / dayMs) + 1);
      baseOwner = a.owner;
      curStart = baseStart;
      curLen = baseLen;
      curOwner = a.owner;

      bar.classList.add('dragging');
      window.addEventListener('mousemove', onMove);
      window.addEventListener('mouseup', onUp, { once: true });
    });

    bar.addEventListener('dblclick', (e) => {
      e.stopPropagation();
      openDrawer(action.id);
    });
    bar.addEventListener('contextmenu', (e) => {
      e.preventDefault();
      e.stopPropagation();
      const a = (curProject().actions || []).find((x) => x.id === action.id);
      if (a) showContextMenu(e.clientX, e.clientY, actionContextItems(a));
    });

    function onMove(e) {
      const dx = e.clientX - originX;
      const dDays = snapDeltaDays(dx);

      if (mode === 'move') {
        curStart = baseStart + dDays;
        // Lane change (vertical)
        const grid = $('#tlGrid');
        const rect = grid.getBoundingClientRect();
        const rowH = 36;
        const yIn = e.clientY - rect.top;
        const idx = clamp(Math.floor(yIn / rowH), 0, state.people.length - 1);
        curOwner = state.people[idx]?.id || curOwner;
        bar.dataset.targetOwner = curOwner;
        applyVisual();
        const owner = state.people.find((p) => p.id === curOwner);
        showTLTooltip(
          `<b>${escapeHTML(action.title)}</b><br>${fmtFull(isoOf(curStart))} → ${fmtFull(isoOf(curStart + curLen - 1))}<br><span style="color:var(--text-faint)">${escapeHTML(owner?.name || '')}</span>`,
          e.clientX, e.clientY);
      } else if (mode === 'resize-l') {
        const newStart = baseStart + dDays;
        const newLen = baseLen - dDays;
        if (newLen >= 1) {
          curStart = newStart;
          curLen = newLen;
          applyVisual();
          showTLTooltip(
            `Start: <b>${fmtFull(isoOf(curStart))}</b><br>End: ${fmtFull(isoOf(curStart + curLen - 1))}<br><span style="color:var(--text-faint)">${curLen} day${curLen === 1 ? '' : 's'}</span>`,
            e.clientX, e.clientY);
        }
      } else if (mode === 'resize-r') {
        const newLen = baseLen + dDays;
        if (newLen >= 1) {
          curLen = newLen;
          applyVisual();
          showTLTooltip(
            `Start: ${fmtFull(isoOf(curStart))}<br>End: <b>${fmtFull(isoOf(curStart + curLen - 1))}</b><br><span style="color:var(--text-faint)">${curLen} day${curLen === 1 ? '' : 's'}</span>`,
            e.clientX, e.clientY);
        }
      }
    }

    function onUp() {
      window.removeEventListener('mousemove', onMove);
      bar.classList.remove('dragging');
      hideTLTooltip();

      // No-op click: nothing actually moved during this drag — leave action untouched.
      if (curStart === baseStart && curLen === baseLen && curOwner === baseOwner) {
        mode = null;
        return;
      }

      const proj = curProject();
      const a = proj.actions.find((x) => x.id === action.id);
      if (!a) return;

      const newStart = isoOf(curStart);
      const newEnd = isoOf(curStart + curLen - 1);
      const ownerChanged = !!(curOwner && curOwner !== baseOwner);
      const startChanged = a.startDate !== newStart;
      const endChanged = a.due !== newEnd;
      if (ownerChanged) {
        a.history.push({ at: todayISO(), what: `Owner: ${personName(a.owner)} → ${personName(curOwner)}` });
        a.owner = curOwner;
      }
      if (startChanged || endChanged) {
        a.history.push({ at: todayISO(), what: `Schedule: ${a.startDate || a.due} → ${newStart}…${newEnd}` });
        a.startDate = newStart;
        a.due = newEnd;
      }
      a.updatedAt = todayISO();
      mode = null;
      if (ownerChanged || startChanged || endChanged) {
        commit('timeline');
        toast(ownerChanged ? 'Reassigned' : 'Schedule updated');
      }
      // No-op clicks just leave the bar where it was.
    }
  }

  /* ---------------------------- Dashboard ---------------------------- */

  // Compute the 5 decision-making KPIs surfaced at the top of the Dashboard.
  // These are derived metrics — no schema additions, just better synthesis of
  // existing fields (commitments, risks, action timestamps, EVM, CR dates).
  function computeDecisionKpis(proj) {
    const todayISO_ = todayISO();

    // 1. Team utilisation — for each horizon, sum committed % FTE across all
    // people / sum capacity. Spare expressed in FTE-weeks.
    const totalCapPct = state.people.reduce((s, p) => s + (p.capacity || 100), 0) || 1;
    const utilFor = (weeksAhead) => {
      let totalCommit = 0;
      state.people.forEach((p) => {
        weeklyLoad(p.id, weeksAhead).forEach((w) => { totalCommit += w.count || 0; });
      });
      const totalCapPeriod = totalCapPct * weeksAhead;
      const pct = totalCapPeriod ? Math.round((totalCommit / totalCapPeriod) * 100) : 0;
      const spareFte = ((totalCapPeriod - totalCommit) / 100);
      return { pct, spareFte };
    };
    const u4  = utilFor(4);
    const u8  = utilFor(8);
    const u12 = utilFor(12);

    // 2. Risk exposure — sum of P×I, residual vs. inherent, with unmitigated tally
    const risks = (proj.risks || []).filter((r) => (r.kind || 'risk') !== 'opportunity');
    let inh = 0, res = 0, unmitigated = 0;
    risks.forEach((r) => {
      const i  = ((r.inherent && r.inherent.probability) || 0) * ((r.inherent && r.inherent.impact) || 0);
      const rs = ((r.residual && r.residual.probability) || 0) * ((r.residual && r.residual.impact) || 0);
      inh += i;
      res += rs;
      if (!r.actionId) unmitigated += rs;
    });
    const reduction = inh > 0 ? Math.round((1 - res / inh) * 100) : 0;

    // 3. Stale actions — open and not updated in ≥ 14 days
    const STALE_DAYS = 14;
    const open = (proj.actions || []).filter((a) => !a.deletedAt && !isClosedStatus(a.status));
    const stale = open.filter((a) => {
      if (!a.updatedAt) return true;
      return dayDiff(todayISO_, a.updatedAt) >= STALE_DAYS;
    });
    const topStale = stale.slice().sort((a, b) => (a.updatedAt || '').localeCompare(b.updatedAt || '')).slice(0, 3);
    const stalePct = open.length ? Math.round((stale.length / open.length) * 100) : 0;

    // 4. Project EVM rollup — sum BAC/PV/AC/EV across cost-centres in scope
    const ccs = getCostCentres();
    let BAC = 0, PV = 0, AC = 0, EV = 0;
    if (ccs.length) {
      // Same week table evmFor builds; reuse the helper for consistency
      const today = new Date(); today.setHours(0, 0, 0, 0);
      let monday = new Date(today);
      while (monday.getDay() !== 1) monday = new Date(monday.getTime() - dayMs);
      const start = new Date(monday.getTime() - 12 * 7 * dayMs);
      const weeks = [];
      for (let i = 0; i < 52; i++) {
        const s = new Date(start.getTime() + i * 7 * dayMs);
        weeks.push({ start: s, isoStart: fmtISO(s) });
      }
      ccs.forEach((cc) => {
        const e = evmFor(cc, weeks, 'cost');
        BAC += e.BAC; PV += e.PV; AC += e.AC; EV += e.EV;
      });
    }
    const CPI = AC > 0 ? EV / AC : 1;
    const SPI = PV > 0 ? EV / PV : 1;

    // 5. CR governance — median turnaround + count of pending > 14 d
    const changes = proj.changes || [];
    const decided = changes.filter((c) => c.status === 'approved' || c.status === 'rejected' || c.status === 'implemented' || c.status === 'cancelled');
    const turnarounds = decided
      .filter((c) => c.originatedDate && c.decisionDate)
      .map((c) => Math.max(0, dayDiff(c.decisionDate, c.originatedDate)));
    let medianTurn = null;
    if (turnarounds.length) {
      const sorted = turnarounds.slice().sort((a, b) => a - b);
      const mid = Math.floor(sorted.length / 2);
      medianTurn = sorted.length % 2 ? sorted[mid] : Math.round((sorted[mid - 1] + sorted[mid]) / 2);
    }
    const pending = changes.filter((c) => c.status === 'proposed' || c.status === 'under_review');
    const stalePending = pending.filter((c) => c.originatedDate && dayDiff(todayISO_, c.originatedDate) >= 14).length;

    return {
      util: { u4, u8, u12 },
      risk: { inh, res, unmitigated, reduction, count: risks.length },
      stale: { count: stale.length, totalOpen: open.length, stalePct, top: topStale },
      evm: { BAC, PV, AC, EV, CPI, SPI, hasData: ccs.length > 0 && AC > 0 },
      cr: { medianTurn, pending: pending.length, stalePending, decided: decided.length, total: changes.length },
    };
  }

  function decisionKpisHTML(d) {
    const todayISO_ = todayISO();
    const utilCls = (p) => p > 100 ? 'bad' : p > 85 ? 'warn' : 'ok';
    const fmtFte = (f) => (f >= 0 ? '+' : '') + f.toFixed(1) + ' FTE-w';
    const riskCls = d.risk.res >= 100 ? 'bad' : d.risk.res >= 50 ? 'warn' : d.risk.count === 0 ? '' : 'ok';
    const staleCls = d.stale.count > 8 ? 'bad' : d.stale.count > 3 ? 'warn' : 'ok';
    const cpiCls = d.evm.CPI >= 1 ? 'ok' : d.evm.CPI >= 0.9 ? 'warn' : 'bad';
    const spiCls = d.evm.SPI >= 1 ? 'ok' : d.evm.SPI >= 0.9 ? 'warn' : 'bad';
    const crCls  = (d.cr.medianTurn ?? 0) > 14 ? 'bad' : (d.cr.medianTurn ?? 0) > 7 ? 'warn' : (d.cr.medianTurn != null ? 'ok' : '');
    const utilBars = [['4 w', d.util.u4], ['8 w', d.util.u8], ['12 w', d.util.u12]].map(([lbl, u]) => `
      <div class="dkpi-mb">
        <span class="dkpi-mb-lbl">${lbl}</span>
        <span class="dkpi-mb-track"><span class="dkpi-mb-fill ${utilCls(u.pct)}" style="width:${Math.min(100, u.pct)}%"></span></span>
        <span class="dkpi-mb-val">${u.pct}%</span>
      </div>`).join('');
    const staleList = d.stale.top.length
      ? d.stale.top.map((a) => {
          const days = a.updatedAt ? Math.abs(dayDiff(todayISO_, a.updatedAt)) : '—';
          return `<div class="dkpi-list-row clickable" data-action-id="${a.id}" title="Open action">
            <span class="dkpi-list-text">${escapeHTML(a.title)}</span>
            <span class="dkpi-list-meta">${days}${typeof days === 'number' ? 'd' : ''}</span>
          </div>`;
        }).join('')
      : '<div class="dkpi-list-empty">No stale actions — nice.</div>';

    return `
      <div class="dkpi-grid">
        <div class="dkpi">
          <div class="dkpi-label">Team utilisation</div>
          <div class="dkpi-num ${utilCls(d.util.u4.pct)}">${d.util.u4.pct}%</div>
          <div class="dkpi-sub">next 4 w · ${fmtFte(d.util.u4.spareFte)} spare</div>
          <div class="dkpi-mini">${utilBars}</div>
        </div>

        <div class="dkpi">
          <div class="dkpi-label">Risk exposure</div>
          <div class="dkpi-num ${riskCls}">${d.risk.res}</div>
          <div class="dkpi-sub">residual P×I · ${d.risk.count} risk${d.risk.count === 1 ? '' : 's'}</div>
          <div class="dkpi-strip">
            <span class="dkpi-strip-pill">↓ ${d.risk.reduction}% from ${d.risk.inh}</span>
            ${d.risk.unmitigated > 0
              ? `<span class="dkpi-strip-warn">⚠ ${d.risk.unmitigated} unmitigated</span>`
              : (d.risk.count > 0 ? '<span class="dkpi-strip-ok">all linked</span>' : '<span class="dkpi-strip-muted">none logged</span>')}
          </div>
        </div>

        <div class="dkpi">
          <div class="dkpi-label">Stale actions</div>
          <div class="dkpi-num ${staleCls}">${d.stale.count}</div>
          <div class="dkpi-sub">of ${d.stale.totalOpen} open · ${d.stale.stalePct}% untouched ≥14 d</div>
          <div class="dkpi-list">${staleList}</div>
        </div>

        <div class="dkpi">
          <div class="dkpi-label">Project performance</div>
          <div class="dkpi-dual">
            <div class="dkpi-dual-cell">
              <div class="dkpi-dual-num ${cpiCls}">${d.evm.hasData ? d.evm.CPI.toFixed(2) : '—'}</div>
              <div class="dkpi-dual-lbl">CPI</div>
            </div>
            <div class="dkpi-dual-sep"></div>
            <div class="dkpi-dual-cell">
              <div class="dkpi-dual-num ${spiCls}">${d.evm.hasData ? d.evm.SPI.toFixed(2) : '—'}</div>
              <div class="dkpi-dual-lbl">SPI</div>
            </div>
          </div>
          <div class="dkpi-sub">${d.evm.hasData ? 'cost / schedule index · all CCs' : 'add a budget to a cost-centre'}</div>
        </div>

        <div class="dkpi">
          <div class="dkpi-label">Change governance</div>
          <div class="dkpi-num ${crCls}">${d.cr.medianTurn != null ? d.cr.medianTurn + ' d' : '—'}</div>
          <div class="dkpi-sub">${d.cr.medianTurn != null ? 'median turnaround · ' : ''}${d.cr.decided} decided · ${d.cr.pending} open</div>
          <div class="dkpi-strip">
            ${d.cr.stalePending > 0
              ? `<span class="dkpi-strip-warn">⚠ ${d.cr.stalePending} pending &gt;14 d</span>`
              : (d.cr.total > 0 ? '<span class="dkpi-strip-ok">no aging pending</span>' : '<span class="dkpi-strip-muted">no CRs yet</span>')}
          </div>
        </div>
      </div>`;
  }

  function renderDashboard(root) {
    const proj = curProject();
    const k = kpis();
    const decisionKpis = computeDecisionKpis(proj);
    const view = document.createElement('div');
    view.className = 'view';
    view.innerHTML = `
      <div class="page-head">
        <div>
          <div class="page-title">${escapeHTML(proj.name)} — Dashboard</div>
          <div class="page-sub">Health summary based on current data</div>
        </div>
      </div>
      ${decisionKpisHTML(decisionKpis)}
      <div class="dashboard">
        <div class="kpi clickable" data-dash-filter="late" title="Click to view late items in Register">
          <div class="kpi-label">Late items</div>
          <div class="kpi-value ${k.late > 0 ? 'bad' : 'ok'}">${k.late}</div>
          <div class="kpi-sub">${k.lateRate}% of all actions</div>
        </div>
        <div class="kpi clickable" data-dash-filter="blocked" title="Click to view blocked items in Register">
          <div class="kpi-label">Blocked</div>
          <div class="kpi-value ${k.blocked > 0 ? 'warn' : 'ok'}">${k.blocked}</div>
          <div class="kpi-sub">${k.blockedRatio}% of all actions</div>
        </div>
        <div class="kpi clickable" data-dash-filter="week" title="Click to view items due this week">
          <div class="kpi-label">Due ≤ 7 days</div>
          <div class="kpi-value ${k.upcoming > 4 ? 'warn' : ''}">${k.upcoming}</div>
          <div class="kpi-sub">Upcoming workload</div>
        </div>
        <div class="kpi clickable" data-dash-filter="done" title="Click to view completed items">
          <div class="kpi-label">Completion</div>
          <div class="kpi-value ${k.completionRate >= 70 ? 'ok' : ''}">${k.completionRate}%</div>
          <div class="kpi-sub">${k.done} / ${k.total} done • throughput ${k.throughput}/2w</div>
        </div>

        <div class="panel half">
          <div class="panel-title">Critical focus</div>
          <div class="crit-list" id="critList"></div>
        </div>
        <div class="panel half">
          <div class="panel-title">Workload by person</div>
          <div id="workload"></div>
        </div>

        <div class="panel">
          <div class="panel-title">Status mix</div>
          <div id="statusMix"></div>
        </div>
      </div>

      <div class="dashboard-section-break">
        <div class="dashboard-section-title">Charts</div>
        <div class="dashboard-section-sub">Trends and projections across the portfolio</div>
      </div>
      <div class="charts-grid">
        <div class="panel chart-panel">
          <div class="panel-title">Schedule deviation waterfall <span class="legend">x = when forecast was made • y = forecast due • diagonal = delivered now</span></div>
          ${chartWaterfall()}
        </div>
        <div class="panel chart-panel half">
          <div class="panel-title">Cumulative workload (next 12 weeks) <span class="legend">click a name to hide / show</span></div>
          <div id="cumWlSlot">${chartCumulativeWorkload(12)}</div>
        </div>
        <div class="panel chart-panel half">
          <div class="panel-title">Activity / week (last 12 weeks)</div>
          ${chartFlow(12)}
        </div>
        <div class="panel chart-panel">
          <div class="panel-title">Per-person workload (next 12 weeks)</div>
          ${chartPerPerson()}
        </div>
      </div>`;
    root.appendChild(view);
    wireCumWlLegend();

    // Stale-action rows in the decision-KPI panel → open the drawer for that action
    $$('.dkpi-list-row[data-action-id]', view).forEach((el) => {
      el.addEventListener('click', () => openDrawer(el.dataset.actionId));
    });

    // Dashboard KPI clicks → navigate to Register pre-filtered.
    // Double-click clears all filters and still navigates to Register.
    $$('.kpi.clickable[data-dash-filter]', view).forEach((el) => {
      el.addEventListener('click', () => {
        const f = el.dataset.dashFilter;
        if (f === 'late' || f === 'week') applyTopbarFilter({ due: f, status: '', view: 'register' });
        else if (f === 'blocked' || f === 'done') applyTopbarFilter({ status: f, due: '', view: 'register' });
      });
      el.addEventListener('dblclick', (e) => {
        e.preventDefault();
        applyTopbarFilter({ clearAll: true, view: 'register' });
      });
    });

    // Critical focus: late + blocked, then top upcoming
    const today = todayISO();
    const critical = (proj.actions || [])
      .filter((a) => (a.status !== 'done') && ((a.due && dayDiff(a.due, today) < 0) || a.status === 'blocked'))
      .sort((a, b) => (a.due || '9999').localeCompare(b.due || '9999'))
      .slice(0, 6);
    const list = $('#critList');
    if (!critical.length) {
      list.innerHTML = '<div class="empty">No late or blocked items — nice.</div>';
    } else {
      list.innerHTML = critical.map((a) => `
        <div class="crit-item clickable" data-id="${a.id}">
          <span class="dot" style="background:${a.status === 'blocked' ? 'var(--bad)' : 'var(--bad)'}"></span>
          <span>${escapeHTML(a.title)}</span>
          <span class="meta">${escapeHTML(personName(a.owner))} • ${a.due ? fmtDate(a.due) : 'no date'} • ${a.status}</span>
        </div>`).join('');
      $$('.crit-item', list).forEach((el) => el.addEventListener('click', () => openDrawer(el.dataset.id)));
    }

    // Workload bars
    const w = $('#workload');
    w.innerHTML = k.workload.map((p) => {
      const pct = clamp(Math.round((p.open / p.capacity) * 100), 0, 200);
      const cls = pct > 100 ? 'over' : pct > 80 ? 'warn' : 'ok';
      return `
        <div class="bar-row">
          <div>${escapeHTML(p.name)}</div>
          <div class="bar-track"><div class="bar-fill ${cls}" style="width:${Math.min(100, pct)}%"></div></div>
          <div class="bar-val">${p.open}/${p.capacity}</div>
        </div>`;
    }).join('');

    // Status mix
    const sm = $('#statusMix');
    const total = k.total || 1;
    const segs = STATUSES.map((s) => {
      const n = (proj.actions || []).filter((a) => a.status === s.id).length;
      const pct = Math.round((n / total) * 100);
      return { name: s.name, n, pct, dot: s.dot };
    });
    sm.innerHTML = segs.map((s) => `
      <div class="bar-row">
        <div><span class="col-dot ${s.dot}" style="display:inline-block;margin-right:6px;"></span>${s.name}</div>
        <div class="bar-track"><div class="bar-fill" style="width:${s.pct}%"></div></div>
        <div class="bar-val">${s.n}</div>
      </div>`).join('');
  }

  /* ----------------------------- Review ------------------------------ */

  let reviewStep = 0;
  const REVIEW_STEPS = [
    { id: 'changes', name: 'What changed' },
    { id: 'late', name: 'Late & blocked' },
    { id: 'next', name: 'What\'s next' },
    { id: 'decisions', name: 'Decisions' },
    { id: 'summary', name: 'Summary' },
  ];

  /* ----------------------------- Charts ------------------------------ */

  // Pick a stable color from the COMPONENT_COLORS palette based on a string id.
  function colorFor(id) {
    let h = 0; for (let i = 0; i < id.length; i++) h = (h * 31 + id.charCodeAt(i)) | 0;
    return COMPONENT_COLORS[Math.abs(h) % COMPONENT_COLORS.length];
  }

  // Distinct palette for stacked layers where collision matters (people in
  // budget chart, cumulative workload, etc). Sized > number of seeded people
  // so neighbours never share a hue.
  const PERSON_PALETTE = [
    { rgb: '96,165,250'  }, // sky
    { rgb: '244,114,182' }, // pink
    { rgb: '52,211,153'  }, // mint
    { rgb: '251,191,36'  }, // amber
    { rgb: '167,139,250' }, // violet
    { rgb: '34,211,238'  }, // cyan
    { rgb: '249,115,22'  }, // orange
    { rgb: '163,230,53'  }, // lime
    { rgb: '236,72,153'  }, // fuchsia
    { rgb: '20,184,166'  }, // teal
    { rgb: '251,113,133' }, // rose
    { rgb: '129,140,248' }, // indigo
    { rgb: '74,222,128'  }, // green
    { rgb: '125,211,252' }, // light-blue
    { rgb: '253,164,175' }, // peach
    { rgb: '161,98,7'    }, // burnt-amber
  ];
  function personColorByIndex(i) { return PERSON_PALETTE[i % PERSON_PALETTE.length]; }

  // Mine an action's history for "Schedule: A → B…C" entries → array of
  // { at: ISO snapshot date, due: ISO forecast end date }.
  // The entry kept at index 0 is the original creation (using the first
  // entry's `at`).
  function scheduleHistory(action) {
    const points = [];
    const re = /Schedule:\s*([\d-]+)\s*→\s*([\d-]+)(?:[…\.]+([\d-]+))?/;
    (action.history || []).forEach((h) => {
      const m = h.what.match(re);
      if (!m) return;
      // First time we encounter Schedule, push the "from" too as the original
      if (!points.length) points.push({ at: action.createdAt || h.at, due: m[1] });
      points.push({ at: h.at, due: m[3] || m[2] });
    });
    if (!points.length && action.due) {
      points.push({ at: action.createdAt || action.updatedAt || todayISO(), due: action.due });
    }
    // For done actions, terminate the line on the 45° diagonal at the actual
    // completion date — that's the visual signal that the work landed.
    if (action.status === 'done') {
      const doneEntry = (action.history || []).filter((h) => /Status:.*→\s*done/.test(h.what)).pop();
      const completedAt = doneEntry?.at || action.updatedAt;
      if (completedAt) points.push({ at: completedAt, due: completedAt });
    } else if (points.length && points[points.length - 1].due !== action.due) {
      // Open action — extend the line to today with the current forecast
      points.push({ at: todayISO(), due: action.due });
    }
    return points;
  }

  // Persistent set of people hidden from the cumulative workload chart.
  // Toggled by clicking the legend.
  const cumWlHidden = new Set();

  // ----- Chart 1: Cumulative workload across all people, weekly -----
  function chartCumulativeWorkload(weeks = 12) {
    const W = 600, H = 220;
    const padL = 38, padR = 12, padT = 10, padB = 24;
    const innerW = W - padL - padR, innerH = H - padT - padB;

    const allPeople = state.people.map((p) => ({
      person: p, color: colorFor(p.id),
      series: weeklyLoad(p.id, weeks),
      hidden: cumWlHidden.has(p.id),
    }));
    const visible = allPeople.filter((ps) => !ps.hidden);
    const visibleCap = visible.reduce((s, ps) => s + (ps.person.capacity || 5), 0);
    const totals = [];
    for (let w = 0; w < weeks; w++) {
      let t = 0; visible.forEach((ps) => t += ps.series[w].count);
      totals.push(t);
    }
    const maxY = Math.max(visibleCap * 1.2 || 4, ...totals, 4);

    const xFor = (i) => padL + (i + 0.5) / weeks * innerW;
    const yFor = (v) => padT + innerH - (v / maxY) * innerH;

    let cumPrev = new Array(weeks).fill(0);
    const layers = visible.map((ps) => {
      const top = ps.series.map((s, i) => cumPrev[i] + s.count);
      const fwd = top.map((v, i) => `${i === 0 ? 'M' : 'L'} ${xFor(i).toFixed(1)} ${yFor(v).toFixed(1)}`).join(' ');
      const back = cumPrev.slice().reverse().map((v, i) => `L ${xFor(weeks - 1 - i).toFixed(1)} ${yFor(v).toFixed(1)}`).join(' ');
      const path = fwd + ' ' + back + ' Z';
      cumPrev = top;
      return { ps, path, top };
    });

    const months = (allPeople[0]?.series || []).map((s, i) => {
      if (s.weekStart.getDate() > 7) return '';
      return `<text class="chart-tick" x="${xFor(i)}" y="${H - 6}" text-anchor="middle">${s.weekStart.toLocaleDateString(undefined, { month: 'short' })}</text>`;
    }).join('');

    const yTicks = [0, Math.round(maxY/2), Math.round(maxY)].map((v) =>
      `<g class="ytick">
        <line x1="${padL - 3}" x2="${padL}" y1="${yFor(v)}" y2="${yFor(v)}" />
        <text x="${padL - 6}" y="${yFor(v) + 3}" text-anchor="end">${v}</text>
      </g>`).join('');

    // Minor horizontal gridlines every 100% (= 1 FTE)
    const FTE_UNIT = 100;
    const minorLines = [];
    for (let v = FTE_UNIT; v <= Math.ceil(maxY); v += FTE_UNIT) {
      const y = yFor(v).toFixed(1);
      minorLines.push(`<line class="cumwl-minor" x1="${padL}" x2="${W - padR}" y1="${y}" y2="${y}" />`);
    }

    const capLine = visibleCap > 0
      ? `<line class="chart-cap" x1="${padL}" x2="${W - padR}" y1="${yFor(visibleCap)}" y2="${yFor(visibleCap)}" />
         <text class="chart-label" x="${W - padR - 4}" y="${Math.max(11, yFor(visibleCap) - 3)}" text-anchor="end">cap ${visibleCap}</text>`
      : '';

    return `
      <svg viewBox="0 0 ${W} ${H}" class="chart-svg" preserveAspectRatio="xMidYMid meet">
        ${minorLines.join('')}
        ${capLine}
        ${layers.map((l) => `<path d="${l.path}" fill="rgba(${l.ps.color.rgb},.65)" stroke="rgba(${l.ps.color.rgb},.95)" stroke-width="0.7"><title>${escapeHTML(l.ps.person.name)}</title></path>`).join('')}
        ${yTicks}
        ${months}
      </svg>
      <div class="chart-legend" data-legend="cumwl">
        ${allPeople.map((ps) => `
          <button type="button" class="legend-item ${ps.hidden ? 'is-hidden' : ''}" data-person-id="${ps.person.id}" title="Click to ${ps.hidden ? 'show' : 'hide'}">
            <span class="dot" style="background:rgba(${ps.color.rgb},.95)"></span>${escapeHTML(ps.person.name)}
          </button>`).join('')}
      </div>`;
  }

  function wireCumWlLegend() {
    const wrap = document.getElementById('cumWlSlot');
    if (!wrap) return;
    wrap.addEventListener('click', (e) => {
      const btn = e.target.closest('.legend-item[data-person-id]');
      if (!btn || !wrap.contains(btn)) return;
      const pid = btn.dataset.personId;
      if (cumWlHidden.has(pid)) cumWlHidden.delete(pid);
      else cumWlHidden.add(pid);
      wrap.innerHTML = chartCumulativeWorkload(12);
    });
  }

  // ----- Chart 2: Flow chart (created / done / slipped per week) -----
  function flowMetrics(weeks = 12) {
    const today = new Date(); today.setHours(0, 0, 0, 0);
    let weekStart = new Date(today);
    while (weekStart.getDay() !== 1) weekStart = new Date(weekStart.getTime() - dayMs);
    const start = new Date(weekStart.getTime() - (weeks - 1) * 7 * dayMs);

    const buckets = new Array(weeks).fill(0).map((_, i) => ({
      start: new Date(start.getTime() + i * 7 * dayMs),
      created: 0, done: 0, slipped: 0, pullIn: 0,
    }));

    const isDoneTrans = /Status:.*→\s*done/;
    const slipRe = /Schedule:\s*([\d-]+)\s*→\s*(?:[\d-]+…)?([\d-]+)/;
    const todayStr = todayISO();
    let overdue = 0, openTotal = 0, blockedTotal = 0;

    state.projects.forEach((proj) => {
      (proj.actions || []).forEach((a) => {
        if (a.deletedAt) return;
        // created
        if (a.createdAt) {
          const idx = Math.floor((parseDate(a.createdAt) - start) / dayMs / 7);
          if (idx >= 0 && idx < weeks) buckets[idx].created++;
        }
        // done — prefer history transition, else updatedAt for status === 'done'
        const doneEntry = (a.history || []).filter((h) => isDoneTrans.test(h.what)).pop();
        let doneAt = doneEntry ? doneEntry.at : (a.status === 'done' ? a.updatedAt : null);
        if (doneAt) {
          const idx = Math.floor((parseDate(doneAt) - start) / dayMs / 7);
          if (idx >= 0 && idx < weeks) buckets[idx].done++;
        }
        // schedule events: slipped vs pulled-in
        (a.history || []).forEach((h) => {
          const m = h.what.match(slipRe);
          if (!m) return;
          const before = parseDate(m[1]);
          const after = parseDate(m[2]);
          const idx = Math.floor((parseDate(h.at) - start) / dayMs / 7);
          if (idx < 0 || idx >= weeks) return;
          if (after > before) buckets[idx].slipped++;
          else if (after < before) buckets[idx].pullIn++;
        });
        // snapshot counters
        if (a.status !== 'done') {
          openTotal++;
          if (a.due && dayDiff(a.due, todayStr) < 0) overdue++;
          if (a.status === 'blocked') blockedTotal++;
        }
      });
    });

    return { buckets, overdue, openTotal, blockedTotal };
  }

  function chartFlow(weeks = 12) {
    const W = 600, H = 220;
    const padL = 38, padR = 12, padT = 12, padB = 26;
    const innerW = W - padL - padR, innerH = H - padT - padB;

    const { buckets, overdue, openTotal, blockedTotal } = flowMetrics(weeks);
    const maxY = Math.max(4, ...buckets.flatMap((b) => [b.created, b.done, b.slipped]));
    const groupW = innerW / weeks;
    const barW = (groupW - 4) / 2;
    const xFor = (i) => padL + i * groupW;
    const yFor = (v) => padT + innerH - (v / maxY) * innerH;

    const bars = buckets.map((b, i) => {
      const xBase = xFor(i) + 2;
      const yDone = yFor(b.done);
      const ySlip = yFor(b.slipped);
      const baseLine = padT + innerH;
      return `
        <rect class="bar-done" x="${xBase}" y="${yDone}" width="${barW}" height="${Math.max(0, baseLine - yDone)}" rx="2">
          <title>Week of ${b.start.toLocaleDateString(undefined, { month: 'short', day: 'numeric' })} — ${b.done} done</title>
        </rect>
        <rect class="bar-slip" x="${xBase + barW}" y="${ySlip}" width="${barW}" height="${Math.max(0, baseLine - ySlip)}" rx="2">
          <title>Week of ${b.start.toLocaleDateString(undefined, { month: 'short', day: 'numeric' })} — ${b.slipped} slipped${b.pullIn ? `, ${b.pullIn} pulled-in` : ''}</title>
        </rect>`;
    }).join('');

    // Created line + dots
    const linePts = buckets.map((b, i) => `${(xFor(i) + groupW/2).toFixed(1)},${yFor(b.created).toFixed(1)}`);
    const linePath = `M ${linePts.join(' L ')}`;
    const lineDots = buckets.map((b, i) => `<circle cx="${(xFor(i) + groupW/2).toFixed(1)}" cy="${yFor(b.created).toFixed(1)}" r="2.5" class="dot-created"><title>Week of ${b.start.toLocaleDateString(undefined, { month: 'short', day: 'numeric' })} — ${b.created} created</title></circle>`).join('');

    const months = buckets.map((b, i) => {
      if (b.start.getDate() > 7) return '';
      return `<text class="chart-tick" x="${(xFor(i) + groupW/2)}" y="${H - 6}" text-anchor="middle">${b.start.toLocaleDateString(undefined, { month: 'short' })}</text>`;
    }).join('');

    const yTicks = [0, Math.round(maxY/2), maxY].map((v) =>
      `<g class="ytick">
        <line x1="${padL - 3}" x2="${padL}" y1="${yFor(v)}" y2="${yFor(v)}" />
        <text x="${padL - 6}" y="${yFor(v) + 3}" text-anchor="end">${v}</text>
      </g>`).join('');

    const totals = buckets.reduce((acc, b) => ({
      created: acc.created + b.created,
      done: acc.done + b.done,
      slipped: acc.slipped + b.slipped,
      pullIn: acc.pullIn + b.pullIn,
    }), { created: 0, done: 0, slipped: 0, pullIn: 0 });
    const net = totals.done - totals.created;
    const netCls = net >= 0 ? 'ok' : 'bad';

    return `
      <svg viewBox="0 0 ${W} ${H}" class="chart-svg" preserveAspectRatio="xMidYMid meet">
        ${yTicks}
        ${bars}
        <path d="${linePath}" class="line-created" />
        ${lineDots}
        ${months}
      </svg>
      <div class="chart-legend">
        <span class="legend-item"><span class="dot" style="background:var(--ok)"></span>Done</span>
        <span class="legend-item"><span class="dot" style="background:var(--bad)"></span>Slipped</span>
        <span class="legend-item"><span class="dot" style="background:var(--accent)"></span>Created</span>
      </div>
      <div class="chart-foot">
        Done <b>${totals.done}</b> • Created <b>${totals.created}</b> • Slipped <b>${totals.slipped}</b>${totals.pullIn ? ` • Pulled-in <b>${totals.pullIn}</b>` : ''}
        • Net <b class="${netCls}">${net >= 0 ? '+' : ''}${net}</b> over ${weeks}w
        <span class="chart-foot-snap">
          Open now <b>${openTotal}</b>
          • Overdue <b class="${overdue > 0 ? 'bad' : ''}">${overdue}</b>
          • Blocked <b class="${blockedTotal > 0 ? 'warn' : ''}">${blockedTotal}</b>
        </span>
      </div>`;
  }

  // ----- Chart 3: 45° Schedule deviation waterfall -----
  function chartWaterfall() {
    const W = 760, H = 360;
    const padL = 64, padR = 12, padT = 14, padB = 38;
    const innerW = W - padL - padR, innerH = H - padT - padB;

    // Pick top N actions by current due date (focus on most relevant).
    const acts = state.projects.flatMap((p) => (p.actions || []).map((a) => ({ proj: p, a })));
    const candidates = acts.filter(({ a }) => a.due && !a.deletedAt).map(({ proj, a }) => ({
      proj, a, hist: scheduleHistory(a),
    })).filter(({ hist }) => hist.length >= 1);
    // Prefer those with movement
    candidates.sort((x, y) => (y.hist.length - x.hist.length) || (x.a.due.localeCompare(y.a.due)));
    const top = candidates.slice(0, 12);
    if (!top.length) {
      return `<div class="empty">No scheduled actions yet — add some, then dragging them on the Timeline will populate this chart over time.</div>`;
    }

    // Domain: x = snapshot dates from earliest 'at' to today
    // y = forecast due dates from earliest to latest forecast
    let xMin = parseDate(todayISO()), xMax = parseDate(todayISO());
    let yMin = parseDate(todayISO()), yMax = parseDate(todayISO());
    top.forEach(({ hist }) => {
      hist.forEach(({ at, due }) => {
        const ad = parseDate(at), dd = parseDate(due);
        if (ad < xMin) xMin = ad;
        if (ad > xMax) xMax = ad;
        if (dd < yMin) yMin = dd;
        if (dd > yMax) yMax = dd;
      });
    });
    // Ensure 45° diagonal is visible: y range covers x range.
    if (yMin > xMin) yMin = xMin;
    if (yMax < xMax) yMax = xMax;
    // Pad a bit
    const pad = (xMax - xMin) * 0.08 + dayMs * 7;
    xMin = new Date(xMin - pad); xMax = new Date(xMax.getTime() + pad);
    yMin = new Date(yMin - pad); yMax = new Date(yMax.getTime() + pad);

    const xFor = (d) => padL + (d - xMin) / (xMax - xMin) * innerW;
    const yFor = (d) => padT + innerH - (d - yMin) / (yMax - yMin) * innerH;

    // 45° diagonal: where y = x — clip to chart bounds
    const lo = new Date(Math.max(xMin.getTime(), yMin.getTime()));
    const hi = new Date(Math.min(xMax.getTime(), yMax.getTime()));
    const diag = `<line class="chart-diag" x1="${xFor(lo)}" y1="${yFor(lo)}" x2="${xFor(hi)}" y2="${yFor(hi)}" />`;
    const todayX = xFor(parseDate(todayISO()));
    const todayLine = `<line class="chart-today" x1="${todayX}" y1="${padT}" x2="${todayX}" y2="${padT + innerH}" stroke-dasharray="2 3" />`;

    // Lines per action
    const lines = top.map(({ proj, a, hist }) => {
      const c = colorFor(a.id);
      const path = hist.map(({ at, due }, i) =>
        `${i === 0 ? 'M' : 'L'} ${xFor(parseDate(at)).toFixed(1)} ${yFor(parseDate(due)).toFixed(1)}`).join(' ');
      const dots = hist.map(({ at, due }) =>
        `<circle cx="${xFor(parseDate(at)).toFixed(1)}" cy="${yFor(parseDate(due)).toFixed(1)}" r="2.5" fill="rgb(${c.rgb})"><title>${escapeHTML(a.title)}: forecast ${due} on ${at}</title></circle>`).join('');
      return `<g><path d="${path}" stroke="rgba(${c.rgb},.85)" stroke-width="1.5" fill="none" /><title>${escapeHTML(a.title)} (${escapeHTML(proj.name)})</title>${dots}</g>`;
    }).join('');

    // X axis ticks (monthly)
    const xTicks = [];
    let dt = new Date(xMin.getFullYear(), xMin.getMonth(), 1);
    if (dt < xMin) dt = new Date(xMin.getFullYear(), xMin.getMonth() + 1, 1);
    while (dt <= xMax) {
      xTicks.push(`<g><line class="chart-grid" x1="${xFor(dt)}" x2="${xFor(dt)}" y1="${padT}" y2="${padT + innerH}" /><text class="chart-tick" x="${xFor(dt)}" y="${H - 16}" text-anchor="middle">${dt.toLocaleDateString(undefined, { month: 'short', year: '2-digit' })}</text></g>`);
      dt = new Date(dt.getFullYear(), dt.getMonth() + 1, 1);
    }
    // Y axis ticks
    const yTicks = [];
    dt = new Date(yMin.getFullYear(), yMin.getMonth(), 1);
    if (dt < yMin) dt = new Date(yMin.getFullYear(), yMin.getMonth() + 1, 1);
    while (dt <= yMax) {
      yTicks.push(`<g><line class="chart-grid" x1="${padL}" x2="${padL + innerW}" y1="${yFor(dt)}" y2="${yFor(dt)}" /><text class="chart-tick" x="${padL - 6}" y="${yFor(dt) + 3}" text-anchor="end">${dt.toLocaleDateString(undefined, { month: 'short', year: '2-digit' })}</text></g>`);
      dt = new Date(dt.getFullYear(), dt.getMonth() + 1, 1);
    }

    return `
      <svg viewBox="0 0 ${W} ${H}" class="chart-svg" preserveAspectRatio="xMidYMid meet">
        ${xTicks.join('')}
        ${yTicks.join('')}
        ${diag}
        <text class="chart-label" x="${xFor(hi) - 4}" y="${yFor(hi) - 4}" text-anchor="end">delivered (45°)</text>
        ${todayLine}
        ${lines}
        <text class="chart-axis-label" x="${padL + innerW/2}" y="${H - 2}" text-anchor="middle">snapshot date →</text>
        <text class="chart-axis-label" x="${padL - 50}" y="${padT + innerH/2}" text-anchor="middle" transform="rotate(-90 ${padL - 50} ${padT + innerH/2})">forecast due →</text>
      </svg>
      <div class="chart-foot">A line going up steeper than 45° = slipping faster than time. Horizontal = stable forecast. Crossing the diagonal = delivered.</div>`;
  }

  // ----- Chart 4: Per-person small multiples -----
  function chartPerPerson() {
    return `<div class="small-multiples">
      ${state.people.map((p) => {
        const series = weeklyLoad(p.id, 12);
        const peak = series.reduce((mx, s) => s.count > mx.count ? s : mx, series[0]);
        const peakCls = peak.count > (p.capacity || 5) ? 'over' : peak.count > (p.capacity || 5) * 0.8 ? 'warn' : 'ok';
        return `
          <div class="sm-cell">
            <div class="sm-head">
              <span class="avatar">${initials(p.name)}</span>
              <b>${escapeHTML(p.name)}</b>
              <span class="sm-peak ${peakCls}">peak ${peak.count}</span>
            </div>
            ${workloadSparkSVG(p, series)}
          </div>`;
      }).join('')}
    </div>`;
  }

  function renderCharts(root) {
    const view = document.createElement('div');
    view.className = 'view';
    view.innerHTML = `
      <div class="page-head">
        <div><div class="page-title">Charts</div><div class="page-sub">Trends and projections across the portfolio</div></div>
      </div>
      <div class="charts-grid">
        <div class="panel chart-panel">
          <div class="panel-title">Schedule deviation waterfall <span class="legend">x = when forecast was made • y = forecast due • diagonal = delivered now</span></div>
          ${chartWaterfall()}
        </div>
        <div class="panel chart-panel half">
          <div class="panel-title">Cumulative workload (next 12 weeks) <span class="legend">click a name to hide / show</span></div>
          <div id="cumWlSlot">${chartCumulativeWorkload(12)}</div>
        </div>
        <div class="panel chart-panel half">
          <div class="panel-title">Activity / week (last 12 weeks)</div>
          ${chartFlow(12)}
        </div>
        <div class="panel chart-panel">
          <div class="panel-title">Per-person workload (next 12 weeks)</div>
          ${chartPerPerson()}
        </div>
      </div>`;
    root.appendChild(view);
    wireCumWlLegend();
  }

  // Merged Review + Reports. Two reading modes share the same data:
  //  - 'walkthrough': stepper wizard with inline edits, for live meetings.
  //  - 'full':        single-page snapshot, for distribution.
  // Period selector + Markdown / Print exports are shared and always
  // visible. Default mode is walkthrough so the live meeting flow lands
  // first; users can flip to 'full' for the print-ready snapshot.
  const reviewModeState = { mode: 'walkthrough' };
  function renderReview(root) {
    const proj = curProject();
    const view = document.createElement('div');
    view.className = 'view';
    const isAll = state.currentProjectId === '__all__';
    if (isAll || !proj) {
      view.innerHTML = `
        <div class="page-head">
          <div>
            <div class="page-title">Review</div>
            <div class="page-sub">Pick a project to review.</div>
          </div>
        </div>
        <div class="empty">Review is project-scoped. Choose a project from the topbar.</div>`;
      root.appendChild(view);
      return;
    }
    const range = reportPeriodRange();
    const data = buildReportData(proj, range.since, range.until);
    const mode = reviewModeState.mode;
    const optSel = (val) => reportState.period === val ? 'selected' : '';

    view.innerHTML = `
      <div class="review">
        <div class="page-head" style="margin-bottom:10px;">
          <div>
            <div class="page-title">${escapeHTML(proj.name)} — Review</div>
            <div class="page-sub">${fmtFull(data.since)} – ${fmtFull(data.until)}</div>
          </div>
          <div class="page-actions">
            <select id="reportPeriod" class="report-period" title="Period for changes / decisions / CRs">
              <option value="7d"  ${optSel('7d')}>Last 7 days</option>
              <option value="30d" ${optSel('30d')}>Last 30 days</option>
              <option value="90d" ${optSel('90d')}>Last 90 days</option>
              <option value="custom" ${optSel('custom')}>Custom…</option>
            </select>
            <span class="report-custom" id="reportCustom" ${reportState.period === 'custom' ? '' : 'hidden'}>
              <input type="date" id="reportSince" value="${reportState.customSince || data.since}" />
              <span class="report-dash">–</span>
              <input type="date" id="reportUntil" value="${reportState.customUntil || data.until}" />
            </span>
            <div class="seg" role="tablist" aria-label="Review mode">
              <button type="button" class="seg-btn ${mode === 'walkthrough' ? 'active' : ''}" data-review-mode="walkthrough">Walk-through</button>
              <button type="button" class="seg-btn ${mode === 'full'        ? 'active' : ''}" data-review-mode="full">Full report</button>
            </div>
            <button class="ghost" id="btnReportCopyMd"  title="Copy as Markdown">Copy MD</button>
            <button class="ghost" id="btnReportDownload" title="Download .md file">Download</button>
            <button class="ghost" id="btnReportPrint"   title="Open print-ready HTML in a new tab">Print → PDF</button>
          </div>
        </div>
        ${mode === 'walkthrough' ? `
          <div class="review-stepper" id="reviewStepper"></div>
          <div class="review-card" id="reviewBody"></div>
          <div class="review-foot">
            <button class="ghost" id="btnReviewPrev">← Previous</button>
            <button class="primary" id="btnReviewNext">Next →</button>
          </div>
        ` : `
          <div class="report" id="reportBody"></div>
        `}
      </div>`;
    root.appendChild(view);

    // --- Period + custom-range wiring (shared) ---
    $('#reportPeriod').addEventListener('change', (e) => {
      reportState.period = e.target.value;
      render();
    });
    $('#reportSince')?.addEventListener('change', (e) => {
      reportState.customSince = e.target.value; reportState.period = 'custom'; render();
    });
    $('#reportUntil')?.addEventListener('change', (e) => {
      reportState.customUntil = e.target.value; reportState.period = 'custom'; render();
    });

    // --- Mode toggle ---
    $$('.seg-btn[data-review-mode]', view).forEach((b) => {
      b.addEventListener('click', () => {
        if (reviewModeState.mode === b.dataset.reviewMode) return;
        reviewModeState.mode = b.dataset.reviewMode;
        render();
      });
    });

    // --- Export buttons (always export the FULL report for the current
    // period — that's the most useful artifact regardless of mode) ---
    $('#btnReportCopyMd').addEventListener('click', async () => {
      const md = reportToMarkdown(buildReportData(curProject(), range.since, range.until));
      try { await navigator.clipboard.writeText(md); toast('Copied as Markdown'); }
      catch (e) { toast('Copy failed — clipboard unavailable'); }
    });
    $('#btnReportDownload').addEventListener('click', () => {
      const md = reportToMarkdown(buildReportData(curProject(), range.since, range.until));
      const blob = new Blob([md], { type: 'text/markdown' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url; a.download = `report-${proj.id}-${range.until}.md`;
      a.click();
      URL.revokeObjectURL(url);
      toast('Report downloaded');
    });
    $('#btnReportPrint').addEventListener('click', () => {
      const html = reportToPrintHTML(buildReportData(curProject(), range.since, range.until));
      const w = window.open('', '_blank');
      if (!w) { toast('Pop-up blocked'); return; }
      w.document.write(html); w.document.close();
      setTimeout(() => { try { w.print(); } catch (e) { /* ignore */ } }, 250);
    });

    // --- Mode-specific body wiring ---
    if (mode === 'walkthrough') {
      drawReviewStep();
      $('#btnReviewPrev').addEventListener('click', () => {
        reviewStep = clamp(reviewStep - 1, 0, REVIEW_STEPS.length - 1);
        drawReviewStep();
      });
      $('#btnReviewNext').addEventListener('click', () => {
        reviewStep = clamp(reviewStep + 1, 0, REVIEW_STEPS.length - 1);
        drawReviewStep();
      });
    } else {
      // Full report — render the same body the old Reports view used,
      // wired with the same drilldowns.
      const k = data.kpis;
      const reportBody = $('#reportBody');
      reportBody.innerHTML = `
        <div class="report-kpis">
          <div class="report-kpi"><div class="report-kpi-num">${k.done}</div><div class="report-kpi-lbl">Done</div></div>
          <div class="report-kpi"><div class="report-kpi-num">${k.changed}</div><div class="report-kpi-lbl">Changed</div></div>
          <div class="report-kpi ${k.late > 0 ? 'bad' : ''}"><div class="report-kpi-num">${k.late}</div><div class="report-kpi-lbl">Late</div></div>
          <div class="report-kpi ${k.blocked > 0 ? 'warn' : ''}"><div class="report-kpi-num">${k.blocked}</div><div class="report-kpi-lbl">Blocked</div></div>
          <div class="report-kpi"><div class="report-kpi-num">${k.decisions}</div><div class="report-kpi-lbl">Decisions</div></div>
          <div class="report-kpi"><div class="report-kpi-num">${k.crs}</div><div class="report-kpi-lbl">CRs decided</div></div>
        </div>
        ${reportSection('What changed', data.changed.slice(0, 30), (a) =>
          `<div class="report-row clickable" data-open-action="${a.id}"><span>${escapeHTML(a.title)}</span><span class="report-meta">${escapeHTML(personName(a.owner))} · ${escapeHTML(a.status)}${a.updatedAt ? ' · ' + a.updatedAt : ''}</span></div>`,
          'No updates in this period.')}
        ${reportSection('Late & blocked', data.lateOrBlocked, (a) => {
          const reason = a.status === 'blocked' ? 'blocked' : `${Math.abs(dayDiff(a.due, data.today))}d late`;
          return `<div class="report-row clickable bad" data-open-action="${a.id}"><span>${escapeHTML(a.title)}</span><span class="report-meta">${escapeHTML(personName(a.owner))} · ${reason}</span></div>`;
        }, 'Nothing late or blocked — nice work.')}
        ${reportSection('Decisions made', data.decisions, (d) =>
          `<div class="report-row"><span>${escapeHTML(d.title)}</span><span class="report-meta">${escapeHTML(personName(d.owner))} · ${escapeHTML(d.date || '—')}</span></div>${d.rationale ? `<div class="report-rationale">${escapeHTML(d.rationale)}</div>` : ''}`,
          'No decisions logged in this period.')}
        ${reportSection('Change requests decided', data.crsDecided, (c) =>
          `<div class="report-row clickable" data-open-cr="${c.id}"><span>${escapeHTML(c.title)}</span><span class="report-meta">${escapeHTML(c.status)}${c.decisionDate ? ' · ' + c.decisionDate : ''}${c.decisionBy ? ' · ' + escapeHTML(personName(c.decisionBy)) : ''}</span></div>`,
          'No CRs decided in this period.')}
        ${reportSection('Top risks (by residual)', data.topRisks, (r) =>
          `<div class="report-row clickable" data-open-risk="${r.id}"><span>${escapeHTML(r.title)}</span><span class="report-meta">residual ${r._score}</span></div>${r.mitigation ? `<div class="report-rationale">${escapeHTML(r.mitigation)}</div>` : ''}`,
          'No risks logged.')}
        <div class="report-section">
          <div class="report-section-title">What's next (next 14 days)</div>
          ${(data.next.milestones.length || data.next.deliverables.length || data.next.actions.length)
            ? `${data.next.milestones.length ? '<div class="report-sub-title">Milestones</div>' + data.next.milestones.map((m) => `<div class="report-row"><span>${escapeHTML(m.name || m.title || '')}</span><span class="report-meta">${escapeHTML(m.date || '')}</span></div>`).join('') : ''}
               ${data.next.deliverables.length ? '<div class="report-sub-title">Deliverables</div>' + data.next.deliverables.map((d) => `<div class="report-row"><span>${escapeHTML(d.name || d.title || '')}</span><span class="report-meta">${escapeHTML(d.date || '')}</span></div>`).join('') : ''}
               ${data.next.actions.length ? '<div class="report-sub-title">Due actions</div>' + data.next.actions.slice(0, 30).map((a) => `<div class="report-row clickable" data-open-action="${a.id}"><span>${escapeHTML(a.title)}</span><span class="report-meta">${escapeHTML(personName(a.owner))} · ${escapeHTML(a.due || '')}</span></div>`).join('') : ''}`
            : '<div class="empty">Nothing scheduled in the next 14 days.</div>'}
        </div>`;
      reportBody.querySelectorAll('[data-open-action]').forEach((el) => {
        el.addEventListener('click', () => openDrawer(el.dataset.openAction));
      });
      reportBody.querySelectorAll('[data-open-cr]').forEach((el) => {
        el.addEventListener('click', () => openChangeRequestEditor(el.dataset.openCr));
      });
      reportBody.querySelectorAll('[data-open-risk]').forEach((el) => {
        el.addEventListener('click', () => openRiskEditor(el.dataset.openRisk));
      });
    }
  }

  function drawReviewStep() {
    const stepper = $('#reviewStepper');
    if (!stepper) return;
    stepper.innerHTML = REVIEW_STEPS.map((s, i) => {
      const cls = i === reviewStep ? 'active' : i < reviewStep ? 'done' : '';
      return `<div class="review-step ${cls}">${i + 1}. ${s.name}</div>`;
    }).join('');

    const body = $('#reviewBody');
    const proj = curProject();
    const today = todayISO();
    const since14 = fmtISO(new Date(Date.now() - 14 * dayMs));

    if (REVIEW_STEPS[reviewStep].id === 'changes') {
      const changed = (proj.actions || []).filter((a) => a.updatedAt >= since14).slice(0, 20);
      body.innerHTML = `<h2>What changed (last 14 days)</h2>` + reviewList(changed, 'No recent updates.');
    }
    else if (REVIEW_STEPS[reviewStep].id === 'late') {
      const items = (proj.actions || []).filter((a) =>
        a.status !== 'done' && ((a.due && dayDiff(a.due, today) < 0) || a.status === 'blocked'));
      body.innerHTML = `<h2>Late & blocked</h2>` + reviewList(items, 'No late or blocked items.');
    }
    else if (REVIEW_STEPS[reviewStep].id === 'next') {
      const items = (proj.actions || []).filter((a) =>
        a.status !== 'done' && a.due && dayDiff(a.due, today) >= 0 && dayDiff(a.due, today) <= 14)
        .sort((a, b) => a.due.localeCompare(b.due));
      body.innerHTML = `<h2>What's next (next 14 days)</h2>` + reviewList(items, 'Nothing scheduled.');
    }
    else if (REVIEW_STEPS[reviewStep].id === 'decisions') {
      const decs = proj.decisions || [];
      body.innerHTML = `<h2>Decisions</h2>
        <div class="row-list">
          ${decs.length ? decs.map((d) => `
            <div class="row">
              <span>${escapeHTML(d.title)}</span>
              <span class="row-meta">${escapeHTML(personName(d.owner))} • ${fmtDate(d.date)}</span>
            </div>`).join('') : '<div class="empty">No decisions logged.</div>'}
        </div>
        <div style="margin-top:14px;"><button class="ghost" id="btnAddDecision">+ Log a decision</button></div>`;
      $('#btnAddDecision').addEventListener('click', () => openQuickAdd('decision'));
    }
    else if (REVIEW_STEPS[reviewStep].id === 'summary') {
      const k = kpis();
      body.innerHTML = `<h2>Summary</h2>
        <p class="page-sub">A snapshot of where the project stands at the end of this review.</p>
        <div class="dashboard" style="grid-template-columns: repeat(4, 1fr);">
          <div class="kpi"><div class="kpi-label">Done</div><div class="kpi-value ok">${k.done}</div></div>
          <div class="kpi"><div class="kpi-label">In progress</div><div class="kpi-value">${k.doing}</div></div>
          <div class="kpi"><div class="kpi-label">Late</div><div class="kpi-value bad">${k.late}</div></div>
          <div class="kpi"><div class="kpi-label">Blocked</div><div class="kpi-value warn">${k.blocked}</div></div>
        </div>
        <p style="margin-top:14px;">Use <em>Export HTML</em> to save a standalone report.</p>`;
    }
  }

  function reviewList(items, emptyMsg) {
    if (!items.length) return `<div class="empty">${emptyMsg}</div>`;
    return `<div class="row-list">` +
      items.map((a) => {
        const dueCls = statusOfDue(a.due, a.status);
        return `
          <div class="row ${dueCls}">
            <span class="clickable" data-id="${a.id}">${escapeHTML(a.title)}</span>
            <span class="row-meta">
              <select class="inline" data-id="${a.id}" data-action="status">
                ${STATUSES.map((s) => `<option value="${s.id}" ${s.id === a.status ? 'selected' : ''}>${s.name}</option>`).join('')}
              </select>
              <select class="inline" data-id="${a.id}" data-action="owner">
                ${state.people.map((p) => `<option value="${p.id}" ${p.id === a.owner ? 'selected' : ''}>${escapeHTML(p.name)}</option>`).join('')}
              </select>
              <input class="inline" type="date" data-id="${a.id}" data-action="due" value="${a.due || ''}" />
            </span>
          </div>`;
      }).join('') +
      `</div>`;
  }

  function exportReviewHTML() {
    const proj = curProject();
    const k = kpis();
    const today = todayISO();
    const since14 = fmtISO(new Date(Date.now() - 14 * dayMs));
    const list = (arr) => arr.length
      ? '<ul>' + arr.map((a) => `<li><b>${escapeHTML(a.title)}</b> — ${escapeHTML(personName(a.owner))} • due ${a.due || '—'} • ${a.status}</li>`).join('') + '</ul>'
      : '<p><i>None.</i></p>';
    const changed = (proj.actions || []).filter((a) => a.updatedAt >= since14);
    const late = (proj.actions || []).filter((a) =>
      a.status !== 'done' && ((a.due && dayDiff(a.due, today) < 0) || a.status === 'blocked'));
    const next = (proj.actions || []).filter((a) =>
      a.status !== 'done' && a.due && dayDiff(a.due, today) >= 0 && dayDiff(a.due, today) <= 14)
      .sort((a, b) => a.due.localeCompare(b.due));
    const html = `<!doctype html><html><head><meta charset="utf-8"><title>Review — ${escapeHTML(proj.name)}</title>
      <style>
        body{font:14px -apple-system,Segoe UI,Helvetica,Arial,sans-serif;color:#1a1c24;max-width:780px;margin:30px auto;padding:0 20px;}
        h1{font-size:22px;margin:0 0 4px;} h2{font-size:16px;margin:24px 0 6px;border-bottom:1px solid #ddd;padding-bottom:4px;}
        .meta{color:#666;font-size:12px;margin-bottom:18px;}
        .kpis{display:grid;grid-template-columns:repeat(4,1fr);gap:10px;margin:14px 0;}
        .kpi{background:#f6f7fb;border:1px solid #e5e7ee;border-radius:8px;padding:10px;}
        .kpi b{font-size:22px;display:block;}
        ul{padding-left:18px;} li{margin:4px 0;}
      </style></head><body>
      <h1>${escapeHTML(proj.name)} — Review</h1>
      <div class="meta">Generated ${fmtFull(today)}</div>
      <div class="kpis">
        <div class="kpi"><b>${k.done}</b>Done</div>
        <div class="kpi"><b>${k.doing}</b>In progress</div>
        <div class="kpi"><b>${k.late}</b>Late</div>
        <div class="kpi"><b>${k.blocked}</b>Blocked</div>
      </div>
      <h2>What changed (last 14 days)</h2>${list(changed)}
      <h2>Late & blocked</h2>${list(late)}
      <h2>Next 14 days</h2>${list(next)}
      <h2>Decisions</h2>${(proj.decisions || []).length
        ? '<ul>' + proj.decisions.map((d) => `<li><b>${escapeHTML(d.title)}</b> — ${escapeHTML(d.rationale || '')} <i>(${escapeHTML(personName(d.owner))}, ${d.date})</i></li>`).join('') + '</ul>'
        : '<p><i>None.</i></p>'}
      <h2>Risks</h2>${(proj.risks || []).length
        ? '<ul>' + proj.risks.map((r) => `<li><b>${escapeHTML(r.title)}</b> (P${r.probability}×I${r.impact}) — ${escapeHTML(r.mitigation || '')}</li>`).join('') + '</ul>'
        : '<p><i>None.</i></p>'}
      </body></html>`;
    const blob = new Blob([html], { type: 'text/html' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url; a.download = `review-${proj.id}-${today}.html`;
    a.click();
    URL.revokeObjectURL(url);
    toast('Review exported');
  }

  /* ---------------------- Engineering side views --------------------- */

  // Per-strip zoom state. Each strip stores its zoom factor (1 = fit,
  // higher = zoomed in horizontally so labels and dates stretch out).
  const stripState = {
    deliverables: { zoom: 1 },
    milestones:   { zoom: 1 },
  };
  const STRIP_ZOOM_STEPS = [1, 1.5, 2, 3, 4, 6, 8];

  // Shared timeline strip — renders points/markers along a horizontal date axis.
  // items: [{ id, name, date: 'YYYY-MM-DD', status?, rgb?, icon? }]
  // Uses lane-packing so labels never overlap, plus a stable per-item color
  // (passed in as `rgb`) so the row swatch matches the timeline circle.
  // opts.zoom: horizontal zoom factor (default 1). At zoom > 1 the SVG renders
  // wider than the panel and scrolls horizontally inside .strip-scroll.
  function renderTimelineStrip(items, opts = {}) {
    const zoom = Math.max(1, Math.min(8, opts.zoom || 1));
    const dated = items.filter((i) => i.date).map((i) => ({ ...i, t: parseDate(i.date) }));
    if (!dated.length) return '<div class="empty">Add a date to any item to see it on the timeline.</div>';

    const today = new Date(); today.setHours(0, 0, 0, 0);
    const all = [...dated.map((d) => d.t), today];
    let min = new Date(Math.min(...all));
    let max = new Date(Math.max(...all));
    const span = Math.max(dayMs * 30, max - min);
    min = new Date(min.getTime() - span * 0.10);
    max = new Date(max.getTime() + span * 0.10);

    const sorted = dated.slice().sort((a, b) => a.t - b.t);
    const baseW = 880;
    const W = Math.round(baseW * zoom);
    const padL = 32, padR = 32;
    const innerW = W - padL - padR;
    const xFor = (t) => padL + (t - min) / (max - min) * innerW;

    // Lane packing — alternate sides, then stack into rows.
    // Each label occupies a horizontal extent at the lane's y; we pick the
    // lowest-numbered lane where the new extent doesn't overlap any prior label.
    const CHAR_W = 6.6; // px per char (12px font, font-weight 600)
    const PAD_X  = 8;   // gap between adjacent labels
    const ROW_H  = 18;  // vertical pitch between lanes — bigger = clearer
    const labelW = (s) => Math.max(36, Math.min(220, s.length * CHAR_W));
    const lanesAbove = []; // each: array of {x0, x1}
    const lanesBelow = [];
    function fit(lanes, x0, x1) {
      for (let l = 0; l < lanes.length; l++) {
        if (!lanes[l].some((r) => x1 + PAD_X >= r.x0 && x0 - PAD_X <= r.x1)) {
          lanes[l].push({ x0, x1 });
          return l;
        }
      }
      lanes.push([{ x0, x1 }]);
      return lanes.length - 1;
    }

    // Group items by date so multiple events on the same day get a small
    // vertical stagger on the axis (otherwise their circles overlap and the
    // user can't tell two events occupy the same date).
    const dateGroups = {};
    sorted.forEach((it) => {
      (dateGroups[it.date] = dateGroups[it.date] || []).push(it);
    });
    const STACK_PITCH = 7; // px between stacked dots on the axis

    const placed = sorted.map((it, i) => {
      const x = xFor(it.t);
      const w = labelW(it.name);
      const x0 = x - w / 2, x1 = x + w / 2;
      // Alternate above/below for a balanced look, then pack into the first free lane.
      const above = i % 2 === 0;
      const lane = above ? fit(lanesAbove, x0, x1) : fit(lanesBelow, x0, x1);
      // Vertical offset along the axis when this item shares its date with others
      const grp = dateGroups[it.date];
      const grpIdx = grp.indexOf(it);
      const dotOffset = grp.length > 1
        ? (grpIdx - (grp.length - 1) / 2) * STACK_PITCH
        : 0;
      return { ...it, x, above, lane, dotOffset, stackSize: grp.length };
    });

    const lanesA = Math.max(1, lanesAbove.length);
    const lanesB = Math.max(1, lanesBelow.length);
    // Allow vertical headroom for stacked-date dots ((max stack - 1)/2 px each side)
    const maxStack = placed.reduce((m, p) => Math.max(m, p.stackSize || 1), 1);
    const stackHalfPx = ((maxStack - 1) / 2) * STACK_PITCH;
    const padT = 18 + lanesA * ROW_H + stackHalfPx;
    const padB = 22 + lanesB * ROW_H + 14 + stackHalfPx;
    const H = padT + padB + 18;
    const yLine = padT + 10;

    // Monthly ticks (also weekly when zoom is high enough to keep them readable)
    const showWeeks = zoom >= 3;
    const ticks = [];
    if (showWeeks) {
      // Weekly ticks (Mondays) — subtle
      let w0 = new Date(min);
      while (w0.getDay() !== 1) w0 = new Date(w0.getTime() + dayMs);
      let wt = w0;
      while (wt <= max) {
        const x = xFor(wt);
        ticks.push(`<line class="strip-week-tick" x1="${x.toFixed(1)}" x2="${x.toFixed(1)}" y1="${(yLine - 3).toFixed(1)}" y2="${(yLine + 3).toFixed(1)}" />`);
        wt = new Date(wt.getTime() + 7 * dayMs);
      }
    }
    let dt = new Date(min.getFullYear(), min.getMonth(), 1);
    if (dt < min) dt = new Date(dt.getFullYear(), dt.getMonth() + 1, 1);
    while (dt <= max) {
      const x = xFor(dt);
      ticks.push(`<g><line class="strip-tick" x1="${x.toFixed(1)}" x2="${x.toFixed(1)}" y1="${(yLine - 6).toFixed(1)}" y2="${(yLine + 6).toFixed(1)}" /><text class="strip-month" x="${x.toFixed(1)}" y="${(H - 4).toFixed(1)}" text-anchor="middle">${dt.toLocaleDateString(undefined, { month: 'short', year: '2-digit' })}</text></g>`);
      dt = new Date(dt.getFullYear(), dt.getMonth() + 1, 1);
    }

    const todayX = xFor(today);
    const todayLine = `<line class="strip-today" x1="${todayX.toFixed(1)}" x2="${todayX.toFixed(1)}" y1="${(padT - 8).toFixed(1)}" y2="${(H - padB + 8).toFixed(1)}" /><text class="strip-today-lbl" x="${(todayX + 4).toFixed(1)}" y="${(padT - 10).toFixed(1)}">today</text>`;

    const markers = placed.map((it) => {
      const x = it.x;
      // Same-date stacking shifts this item's circle vertically along the axis
      const cy = yLine + (it.dotOffset || 0);
      const labelY = it.above
        ? padT - 8 - (lanesAbove.length - 1 - it.lane) * ROW_H
        : H - padB + 18 + it.lane * ROW_H;
      const dateY = it.above ? labelY + 12 : labelY + 12;
      // Stems run between the actual circle position (cy) and the label line
      const stemY1 = it.above
        ? labelY - 9
        : cy + 7;
      const stemY2 = it.above
        ? cy - 7
        : labelY - 14;
      const isDone = it.status === 'done';
      const isLate = !isDone && it.t < today;
      const cls = isDone ? 'done' : (isLate ? 'late' : '');
      const rgb = it.rgb || '129,140,248';
      const labelFill = isDone ? 'var(--text-dim)' : (isLate ? 'var(--bad)' : `rgb(${rgb})`);
      const safeName = escapeHTML(it.name);
      const r = 7;
      // Status badge: ✓ for done, ! for late — anchored to the actual circle
      // position so it still hugs each stacked dot when same-date events split.
      const badge = isDone
        ? `<g class="strip-badge done"><circle cx="${(x + 8).toFixed(1)}" cy="${(cy - 8).toFixed(1)}" r="6" fill="var(--ok)" stroke="var(--bg-1)" stroke-width="1.6" /><text x="${(x + 8).toFixed(1)}" y="${(cy - 5).toFixed(1)}" text-anchor="middle" fill="var(--bg-1)" font-size="8" font-weight="700">✓</text></g>`
        : isLate
          ? `<g class="strip-badge late"><circle cx="${(x + 8).toFixed(1)}" cy="${(cy - 8).toFixed(1)}" r="6" fill="var(--bad)" stroke="var(--bg-1)" stroke-width="1.6" /><text x="${(x + 8).toFixed(1)}" y="${(cy - 5).toFixed(1)}" text-anchor="middle" fill="var(--bg-1)" font-size="9" font-weight="700">!</text></g>`
          : '';
      return `
        <g class="strip-mark ${cls}" data-strip-id="${it.id}">
          <line class="strip-stem" x1="${x.toFixed(1)}" x2="${x.toFixed(1)}" y1="${stemY1.toFixed(1)}" y2="${stemY2.toFixed(1)}" stroke="rgb(${rgb})" />
          <circle class="strip-pt" cx="${x.toFixed(1)}" cy="${cy.toFixed(1)}" r="${r}" fill="rgb(${rgb})" stroke="var(--bg-1)" stroke-width="1.8" />
          ${badge}
          <text class="strip-lbl" x="${x.toFixed(1)}" y="${labelY.toFixed(1)}" text-anchor="middle" fill="${labelFill}">${safeName}</text>
          <text class="strip-date" x="${x.toFixed(1)}" y="${dateY.toFixed(1)}" text-anchor="middle">${fmtDate(it.date)}</text>
        </g>`;
    }).join('');

    // At zoom=1 the SVG fills its wrapper; at zoom>1 it gets an explicit pixel
    // width so the wrapper scrolls horizontally.
    const widthAttr = zoom > 1 ? `width="${W}"` : 'width="100%"';
    return `
      <svg viewBox="0 0 ${W} ${H}" class="strip-svg" ${widthAttr} preserveAspectRatio="xMinYMin meet">
        ${ticks.join('')}
        <line class="strip-axis" x1="${padL}" x2="${W - padR}" y1="${yLine}" y2="${yLine}" />
        ${todayLine}
        ${markers}
      </svg>`;
  }

  // Zoom-control markup + handler for a timeline strip. `id` is a unique
  // prefix used for the data-attributes / element IDs (so multiple strips on
  // the same view don't collide).
  function stripZoomControlsHTML(id) {
    return `
      <span class="strip-zoom" role="group" aria-label="Timeline zoom">
        <button type="button" class="icon-btn" id="${id}-zoom-out" title="Zoom out">−</button>
        <span class="strip-zoom-val" id="${id}-zoom-val">1×</span>
        <button type="button" class="icon-btn" id="${id}-zoom-in" title="Zoom in">+</button>
        <button type="button" class="icon-btn" id="${id}-zoom-reset" title="Reset zoom">⟲</button>
      </span>`;
  }
  function wireStripZoom(scope, id, state, redraw) {
    const STEPS = STRIP_ZOOM_STEPS;
    const idxOf = (z) => {
      let best = 0, bestDiff = Infinity;
      STEPS.forEach((s, i) => { const d = Math.abs(s - z); if (d < bestDiff) { bestDiff = d; best = i; } });
      return best;
    };
    const set = (z) => { state.zoom = z; redraw(); };
    scope.querySelector(`#${id}-zoom-in`).addEventListener('click', () => {
      const i = idxOf(state.zoom); set(STEPS[Math.min(STEPS.length - 1, i + 1)]);
    });
    scope.querySelector(`#${id}-zoom-out`).addEventListener('click', () => {
      const i = idxOf(state.zoom); set(STEPS[Math.max(0, i - 1)]);
    });
    scope.querySelector(`#${id}-zoom-reset`).addEventListener('click', () => set(1));
    // Cmd/Ctrl + scroll zoom inside the strip
    const wrap = scope.querySelector(`#${id}-scroll`);
    if (wrap) {
      wrap.addEventListener('wheel', (e) => {
        if (!(e.metaKey || e.ctrlKey)) return;
        e.preventDefault();
        const i = idxOf(state.zoom);
        const next = e.deltaY < 0 ? STEPS[Math.min(STEPS.length - 1, i + 1)] : STEPS[Math.max(0, i - 1)];
        if (next !== state.zoom) set(next);
      }, { passive: false });
    }
  }

  // Shared floating tooltip + hover wiring for the deliverables / milestones strip.
  let _stripTipEl = null;
  function ensureStripTipEl() {
    if (_stripTipEl) return _stripTipEl;
    _stripTipEl = document.createElement('div');
    _stripTipEl.className = 'strip-tooltip';
    document.body.appendChild(_stripTipEl);
    return _stripTipEl;
  }
  function wireStripHover(scope, lookup) {
    const svg = scope.querySelector('.strip-svg');
    if (!svg) return;
    const tip = ensureStripTipEl();
    function show(g, x, y) {
      const it = lookup(g.dataset.stripId);
      if (!it) return;
      const today = todayISO();
      const diff = dayDiff(it.date, today);
      const rel = diff === 0 ? 'today'
        : diff > 0 ? `in ${diff} day${diff === 1 ? '' : 's'}`
        : `${-diff} day${diff === -1 ? '' : 's'} ago`;
      const isDone = it.status === 'done';
      const isLate = !isDone && diff < 0;
      const statusLbl = isDone ? '✓ Done'
        : isLate ? '⚠ Late'
        : (it.status === 'doing' ? '◐ In progress' : '○ Not started');
      const swatch = it.rgb ? `<span class="strip-tip-swatch" style="background:rgb(${it.rgb})"></span>` : '';
      tip.innerHTML = `
        <div class="strip-tip-head">${swatch}<span class="strip-tip-title">${escapeHTML(it.name)}</span></div>
        <div class="strip-tip-row"><span class="strip-tip-lbl">Date</span><span>${fmtFull(it.date)}</span></div>
        <div class="strip-tip-row"><span class="strip-tip-lbl">When</span><span>${rel}</span></div>
        <div class="strip-tip-row"><span class="strip-tip-lbl">Status</span><span class="${isDone ? 'ok' : isLate ? 'bad' : ''}">${statusLbl}</span></div>`;
      tip.style.display = 'block';
      const r = tip.getBoundingClientRect();
      let px = x + 14, py = y + 14;
      if (px + r.width  > innerWidth  - 8) px = x - r.width  - 14;
      if (py + r.height > innerHeight - 8) py = y - r.height - 14;
      tip.style.left = Math.max(8, px) + 'px';
      tip.style.top  = Math.max(8, py) + 'px';
    }
    function hide() { tip.style.display = 'none'; }
    svg.addEventListener('mousemove', (e) => {
      const g = e.target.closest('.strip-mark[data-strip-id]');
      if (!g) { hide(); return; }
      show(g, e.clientX, e.clientY);
    });
    svg.addEventListener('mouseleave', hide);
  }

  function renderDeliverables(root) {
    const proj = curProject();
    const view = document.createElement('div');
    view.className = 'view';
    const sortedDels = (proj.deliverables || []).slice().sort((a, b) => (a.dueDate || '').localeCompare(b.dueDate || ''));
    const items = sortedDels.map((d, i) => ({
      id: d.id, name: d.name, date: d.dueDate, status: d.status,
      rgb: PERSON_PALETTE[i % PERSON_PALETTE.length].rgb,
    }));
    const colorById = Object.fromEntries(items.map((it) => [it.id, it.rgb]));
    view.innerHTML = `
      <div class="page-head">
        <div><div class="page-title">Deliverables</div><div class="page-sub">Optional — group actions under a deliverable.</div></div>
        <div class="page-actions"><button class="ghost" id="btnAddDel">+ Deliverable</button></div>
      </div>
      <div class="row-list" id="delList"></div>
      <div class="panel chart-panel" style="margin-top:14px;">
        <div class="panel-title">
          <span>Timeline</span>
          ${stripZoomControlsHTML('strip-del')}
        </div>
        <div class="strip-scroll" id="strip-del-scroll">${renderTimelineStrip(items, { zoom: stripState.deliverables.zoom })}</div>
      </div>`;
    root.appendChild(view);
    wireStripZoom(view, 'strip-del', stripState.deliverables, () => {
      const wrap = view.querySelector('#strip-del-scroll');
      wrap.innerHTML = renderTimelineStrip(items, { zoom: stripState.deliverables.zoom });
      wireStripHover(view, (id) => items.find((x) => x.id === id));
      view.querySelector('#strip-del-zoom-val').textContent = stripState.deliverables.zoom + '×';
    });
    const list = $('#delList');
    if (!proj.deliverables?.length) list.innerHTML = '<div class="empty">No deliverables yet.</div>';
    else {
      list.innerHTML = sortedDels.map((d) => {
        const dueCls = d.status === 'done' ? '' :
          (d.dueDate && dayDiff(d.dueDate, todayISO()) < 0 ? 'late' : '');
        const rgb = colorById[d.id] || '129,140,248';
        return `
          <div class="row ${dueCls}" data-deliverable-id="${d.id}">
            ${ROW_GRIP_HTML}
            <span class="row-swatch" style="background:rgb(${rgb})" aria-hidden="true"></span>
            <span>◆ ${escapeHTML(d.name)}</span>
            <span class="row-meta">${d.dueDate ? fmtFull(d.dueDate) : '—'} • ${escapeHTML(d.status || 'todo')}</span>
          </div>`;
      }).join('');
      wireListReorder(list, {
        rowSelector: '.row[data-deliverable-id]',
        idAttr: 'deliverableId',
        getArray: () => proj.deliverables,
        setOrder: (ids) => { proj.deliverables = ids.map((id) => proj.deliverables.find((x) => x.id === id)).filter(Boolean); },
        commitName: 'deliverables-reorder',
      });
      wireStripHover(view, (id) => items.find((x) => x.id === id));
      $$('.row[data-deliverable-id]', list).forEach((row) => {
        row.addEventListener('contextmenu', (e) => {
          e.preventDefault();
          const id = row.dataset.deliverableId;
          const d = (proj.deliverables || []).find((x) => x.id === id);
          if (!d) return;
          showContextMenu(e.clientX, e.clientY, [
            { icon: '✎', label: 'Edit…', onClick: () => openDeliverableEditor(id) },
            { divider: true },
            { icon: '×', label: 'Delete deliverable', danger: true, onClick: () => {
              const linked = (proj.actions || []).filter((a) => a.deliverable === id).length;
              if (!confirm(`Delete "${d.name}"?` + (linked ? ` (${linked} action${linked === 1 ? '' : 's'} will become unlinked)` : ''))) return;
              proj.deliverables = (proj.deliverables || []).filter((x) => x.id !== id);
              (proj.actions || []).forEach((a) => { if (a.deliverable === id) a.deliverable = null; });
              commit('deliverable-delete');
              toast('Deleted');
            }},
          ]);
        });
        row.addEventListener('dblclick', () => openDeliverableEditor(row.dataset.deliverableId));
      });
    }
    $('#btnAddDel').addEventListener('click', () => openQuickAdd('deliverable'));
  }

  function openDeliverableEditor(deliverableId) {
    const proj = curProject();
    const d = (proj.deliverables || []).find((x) => x.id === deliverableId);
    if (!d) return;
    const overlay = document.createElement('div');
    overlay.className = 'overlay desc-overlay';
    overlay.innerHTML = `
      <div class="desc-modal" style="width:520px;">
        <div class="desc-head">
          <div class="desc-title">Edit deliverable</div>
          <button class="icon-btn" id="dvClose" title="Close">×</button>
        </div>
        <div style="padding:14px 16px; display:flex; flex-direction:column; gap:10px;">
          <div class="field"><label>Name</label><input id="dvName" value="${escapeHTML(d.name)}" /></div>
          <div class="qa-row">
            <div class="field"><label>Due</label><input id="dvDue" type="date" value="${d.dueDate || ''}" /></div>
            <div class="field"><label>Status</label>
              <select id="dvStatus">
                <option value="todo" ${d.status === 'todo' ? 'selected' : ''}>Not started</option>
                <option value="doing" ${d.status === 'doing' ? 'selected' : ''}>In progress</option>
                <option value="done" ${d.status === 'done' ? 'selected' : ''}>Done</option>
              </select>
            </div>
          </div>
          <div class="field"><label>Component <span class="muted">— optional</span></label>
            <select id="dvComp">
              <option value="">— None</option>
              ${(proj.components || []).map((cmp) => `<option value="${cmp.id}" ${cmp.id === d.component ? 'selected' : ''}>${escapeHTML(cmp.name)}</option>`).join('')}
            </select>
          </div>
        </div>
        <div class="desc-foot">
          <button class="ghost" id="dvCancel">Cancel</button>
          <button class="ghost" id="dvDelete" style="margin-left:auto; color:var(--bad);">Delete</button>
          <button class="primary" id="dvSave">Save</button>
        </div>
      </div>`;
    document.body.appendChild(overlay);
    setTimeout(() => document.getElementById('dvName').focus(), 30);
    const close = () => overlay.remove();
    overlay.querySelector('#dvClose').addEventListener('click', close);
    overlay.querySelector('#dvCancel').addEventListener('click', close);
    // Modal closes ONLY via the explicit × button (and Cancel where present).
    // Backdrop click and Escape are intentionally NOT bound — losing
    // in-progress edits to a stray click outside the modal is too easy.
    overlay.querySelector('#dvSave').addEventListener('click', () => {
      d.name = document.getElementById('dvName').value.trim() || d.name;
      d.dueDate = document.getElementById('dvDue').value || null;
      d.status = document.getElementById('dvStatus').value;
      d.component = document.getElementById('dvComp').value || null;
      commit('deliverable-edit');
      close();
      toast('Saved');
    });
    overlay.querySelector('#dvDelete').addEventListener('click', () => {
      const linked = (proj.actions || []).filter((a) => a.deliverable === d.id).length;
      if (!confirm(`Delete "${d.name}"?` + (linked ? ` (${linked} action${linked === 1 ? '' : 's'} will become unlinked)` : ''))) return;
      proj.deliverables = (proj.deliverables || []).filter((x) => x.id !== d.id);
      (proj.actions || []).forEach((a) => { if (a.deliverable === d.id) a.deliverable = null; });
      commit('deliverable-delete');
      close();
      toast('Deleted');
    });
  }

  function renderMilestones(root) {
    const proj = curProject();
    const view = document.createElement('div');
    view.className = 'view';
    const sortedMs = (proj.milestones || []).slice().sort((a, b) => (a.date || '').localeCompare(b.date || ''));
    const items = sortedMs.map((m, i) => ({
      id: m.id, name: m.name, date: m.date, status: m.status,
      rgb: PERSON_PALETTE[i % PERSON_PALETTE.length].rgb,
    }));
    const colorById = Object.fromEntries(items.map((it) => [it.id, it.rgb]));
    view.innerHTML = `
      <div class="page-head">
        <div><div class="page-title">Milestones</div><div class="page-sub">Optional — anchor key dates.</div></div>
        <div class="page-actions"><button class="ghost" id="btnAddMile">+ Milestone</button></div>
      </div>
      <div class="row-list" id="mileList"></div>
      <div class="panel chart-panel" style="margin-top:14px;">
        <div class="panel-title">
          <span>Timeline</span>
          ${stripZoomControlsHTML('strip-ms')}
        </div>
        <div class="strip-scroll" id="strip-ms-scroll">${renderTimelineStrip(items, { zoom: stripState.milestones.zoom })}</div>
      </div>`;
    root.appendChild(view);
    wireStripZoom(view, 'strip-ms', stripState.milestones, () => {
      const wrap = view.querySelector('#strip-ms-scroll');
      wrap.innerHTML = renderTimelineStrip(items, { zoom: stripState.milestones.zoom });
      wireStripHover(view, (id) => items.find((x) => x.id === id));
      view.querySelector('#strip-ms-zoom-val').textContent = stripState.milestones.zoom + '×';
    });
    const list = $('#mileList');
    if (!proj.milestones?.length) list.innerHTML = '<div class="empty">No milestones yet.</div>';
    else {
      list.innerHTML = sortedMs.map((m) => {
        const dueCls = m.status === 'done' ? '' :
          (m.date && dayDiff(m.date, todayISO()) < 0 ? 'late' : '');
        const rgb = colorById[m.id] || '129,140,248';
        return `
          <div class="row ${dueCls}" data-milestone-id="${m.id}">
            ${ROW_GRIP_HTML}
            <span class="row-swatch" style="background:rgb(${rgb})" aria-hidden="true"></span>
            <span>◇ ${escapeHTML(m.name)}</span>
            <span class="row-meta">${m.date ? fmtFull(m.date) : '—'} • ${escapeHTML(m.status || 'todo')}</span>
          </div>`;
      }).join('');
      wireListReorder(list, {
        rowSelector: '.row[data-milestone-id]',
        idAttr: 'milestoneId',
        getArray: () => proj.milestones,
        setOrder: (ids) => { proj.milestones = ids.map((id) => proj.milestones.find((x) => x.id === id)).filter(Boolean); },
        commitName: 'milestones-reorder',
      });
      wireStripHover(view, (id) => items.find((x) => x.id === id));
      $$('.row[data-milestone-id]', list).forEach((row) => {
        row.addEventListener('contextmenu', (e) => {
          e.preventDefault();
          const id = row.dataset.milestoneId;
          const m = (proj.milestones || []).find((x) => x.id === id);
          if (!m) return;
          showContextMenu(e.clientX, e.clientY, [
            { icon: '✎', label: 'Edit…', onClick: () => openMilestoneEditor(id) },
            { divider: true },
            { icon: '×', label: 'Delete milestone', danger: true, onClick: () => {
              const linked = (proj.actions || []).filter((a) => a.milestone === id).length;
              if (!confirm(`Delete "${m.name}"?` + (linked ? ` (${linked} action${linked === 1 ? '' : 's'} will become unlinked)` : ''))) return;
              proj.milestones = (proj.milestones || []).filter((x) => x.id !== id);
              (proj.actions || []).forEach((a) => { if (a.milestone === id) a.milestone = null; });
              commit('milestone-delete');
              toast('Deleted');
            }},
          ]);
        });
        row.addEventListener('dblclick', () => openMilestoneEditor(row.dataset.milestoneId));
      });
    }
    $('#btnAddMile').addEventListener('click', () => openQuickAdd('milestone'));
  }

  function openMilestoneEditor(milestoneId) {
    const proj = curProject();
    const m = (proj.milestones || []).find((x) => x.id === milestoneId);
    if (!m) return;
    const overlay = document.createElement('div');
    overlay.className = 'overlay desc-overlay';
    overlay.innerHTML = `
      <div class="desc-modal" style="width:520px;">
        <div class="desc-head">
          <div class="desc-title">Edit milestone</div>
          <button class="icon-btn" id="msClose" title="Close">×</button>
        </div>
        <div style="padding:14px 16px; display:flex; flex-direction:column; gap:10px;">
          <div class="field"><label>Name</label><input id="msName" value="${escapeHTML(m.name)}" /></div>
          <div class="qa-row">
            <div class="field"><label>Start date</label><input id="msDate" type="date" value="${m.date || ''}" /></div>
            <div class="field"><label>End date <span class="muted">— leave empty for a single day</span></label><input id="msEndDate" type="date" value="${m.endDate || ''}" /></div>
            <div class="field"><label>Status</label>
              <select id="msStatus">
                <option value="todo" ${m.status === 'todo' ? 'selected' : ''}>Not started</option>
                <option value="doing" ${m.status === 'doing' ? 'selected' : ''}>In progress</option>
                <option value="done" ${m.status === 'done' ? 'selected' : ''}>Done</option>
              </select>
            </div>
          </div>
          <div class="field"><label>Component <span class="muted">— optional</span></label>
            <select id="msComp">
              <option value="">— None</option>
              ${(proj.components || []).map((cmp) => `<option value="${cmp.id}" ${cmp.id === m.component ? 'selected' : ''}>${escapeHTML(cmp.name)}</option>`).join('')}
            </select>
          </div>
        </div>
        <div class="desc-foot">
          <button class="ghost" id="msCancel">Cancel</button>
          <button class="ghost" id="msDelete" style="margin-left:auto; color:var(--bad);">Delete</button>
          <button class="primary" id="msSave">Save</button>
        </div>
      </div>`;
    document.body.appendChild(overlay);
    setTimeout(() => document.getElementById('msName').focus(), 30);
    const close = () => overlay.remove();
    overlay.querySelector('#msClose').addEventListener('click', close);
    overlay.querySelector('#msCancel').addEventListener('click', close);
    // Modal closes ONLY via the explicit × button (and Cancel where present).
    // Backdrop click and Escape are intentionally NOT bound — losing
    // in-progress edits to a stray click outside the modal is too easy.
    overlay.querySelector('#msSave').addEventListener('click', () => {
      m.name = document.getElementById('msName').value.trim() || m.name;
      m.date = document.getElementById('msDate').value || null;
      const ed = document.getElementById('msEndDate').value || null;
      // Reject end-before-start; treat equal as single-day (clear endDate).
      if (ed && m.date && ed < m.date) { toast('End date can\'t be before start date'); return; }
      m.endDate = (ed && ed !== m.date) ? ed : null;
      m.status = document.getElementById('msStatus').value;
      m.component = document.getElementById('msComp').value || null;
      commit('milestone-edit');
      close();
      toast('Saved');
    });
    overlay.querySelector('#msDelete').addEventListener('click', () => {
      const linked = (proj.actions || []).filter((a) => a.milestone === m.id).length;
      if (!confirm(`Delete "${m.name}"?` + (linked ? ` (${linked} action${linked === 1 ? '' : 's'} will become unlinked)` : ''))) return;
      proj.milestones = (proj.milestones || []).filter((x) => x.id !== m.id);
      (proj.actions || []).forEach((a) => { if (a.milestone === m.id) a.milestone = null; });
      commit('milestone-delete');
      close();
      toast('Deleted');
    });
  }

  // Click a meeting chip in the calendar → open this small editor with
  // the same progressive-disclosure layout used by Quick Add: only the
  // fields required for the current selection are visible. Switching
  // toggle / unit re-shapes the schema in place so a meeting can be
  // promoted from one-off → recurring (or vice-versa, or a different
  // unit) without re-creating it.
  function openMeetingEditor(meetingId) {
    const proj = curProject();
    const m = (proj.meetings || []).find((x) => x.id === meetingId);
    if (!m) return;
    const repeating = m.kind === 'recurring';
    const initUnit  = repeating ? (m.recurUnit || 'week') : 'week';
    const initInterval = repeating ? (m.interval || 1) : 1;
    const dowOpts = [1,2,3,4,5,6,0].map((d) => {
      const names = { 0: 'Sunday', 1: 'Monday', 2: 'Tuesday', 3: 'Wednesday', 4: 'Thursday', 5: 'Friday', 6: 'Saturday' };
      return `<option value="${d}" ${d === (m.dayOfWeek ?? 1) ? 'selected' : ''}>${names[d]}</option>`;
    }).join('');
    const overlay = document.createElement('div');
    overlay.className = 'overlay desc-overlay';
    overlay.innerHTML = `
      <div class="desc-modal" style="width:520px;">
        <div class="desc-head">
          <div class="desc-title">Edit meeting</div>
          <button class="icon-btn" id="mtgClose" title="Close">×</button>
        </div>
        <div style="padding:14px 16px; display:flex; flex-direction:column; gap:10px;">
          <div class="field"><label>Title</label><input id="mtgTitle" value="${escapeHTML(m.title || '')}" /></div>
          <div class="qa-row">
            <div class="field"><label id="mtgDateLbl">${repeating ? 'Start date' : 'Date'}</label>
              <input id="mtgDate" type="date" value="${(repeating ? m.startDate : m.date) || ''}" />
            </div>
            <div class="field"><label>Time <span class="muted">— optional</span></label>
              <input id="mtgTime" type="time" value="${m.time || ''}" />
            </div>
          </div>
          <div class="field">
            <label class="qa-toggle">
              <input type="checkbox" id="mtgRepeats" ${repeating ? 'checked' : ''} />
              <span>Repeating meeting</span>
            </label>
          </div>
          <div id="mtgRecurWrap" ${repeating ? '' : 'hidden'}>
            <div class="qa-row qa-row-tight">
              <div class="field" style="flex: 0 0 auto;">
                <label>Every</label>
                <input id="mtgInterval" type="number" min="1" max="99" value="${initInterval}" style="width:64px;" />
              </div>
              <div class="field" style="flex: 1;">
                <label>&nbsp;</label>
                <select id="mtgUnit">
                  <option value="day"   ${initUnit === 'day'   ? 'selected' : ''}>Day(s)</option>
                  <option value="week"  ${initUnit === 'week'  ? 'selected' : ''}>Week(s)</option>
                  <option value="month" ${initUnit === 'month' ? 'selected' : ''}>Month(s)</option>
                </select>
              </div>
              <div class="field" id="mtgDowField" ${initUnit === 'week' ? '' : 'hidden'}>
                <label>On</label>
                <select id="mtgDow">${dowOpts}</select>
              </div>
            </div>
            <div class="field">
              <label>Ends <span class="muted">— optional</span></label>
              <input id="mtgEndDate" type="date" value="${m.endDate || ''}" />
            </div>
          </div>
          <div class="field"><label>Component <span class="muted">— optional</span></label>
            <select id="mtgComp">
              <option value="">— None</option>
              ${(proj.components || []).map((cmp) => `<option value="${cmp.id}" ${cmp.id === m.component ? 'selected' : ''}>${escapeHTML(cmp.name)}</option>`).join('')}
            </select>
          </div>
        </div>
        <div class="desc-foot">
          <button class="ghost" id="mtgCancel">Cancel</button>
          <button class="ghost" id="mtgDelete" style="margin-left:auto; color:var(--bad);">Delete</button>
          <button class="primary" id="mtgSave">Save</button>
        </div>
      </div>`;
    document.body.appendChild(overlay);
    setTimeout(() => document.getElementById('mtgTitle').focus(), 30);
    const close = () => overlay.remove();
    const repeatChk = overlay.querySelector('#mtgRepeats');
    const recurWrap = overlay.querySelector('#mtgRecurWrap');
    const dateInput = overlay.querySelector('#mtgDate');
    const unitSel   = overlay.querySelector('#mtgUnit');
    const dowField  = overlay.querySelector('#mtgDowField');
    const dowSel    = overlay.querySelector('#mtgDow');
    const dateLbl   = overlay.querySelector('#mtgDateLbl');
    function syncDowFromDate() {
      const v = dateInput.value;
      if (v) dowSel.value = String(parseDate(v).getDay());
    }
    repeatChk.addEventListener('change', () => {
      recurWrap.hidden = !repeatChk.checked;
      dateLbl.textContent = repeatChk.checked ? 'Start date' : 'Date';
      if (repeatChk.checked && unitSel.value === 'week') syncDowFromDate();
    });
    unitSel.addEventListener('change', () => {
      dowField.hidden = unitSel.value !== 'week';
      if (unitSel.value === 'week') syncDowFromDate();
    });
    dateInput.addEventListener('change', () => {
      if (repeatChk.checked && unitSel.value === 'week') syncDowFromDate();
    });
    overlay.querySelector('#mtgClose').addEventListener('click', close);
    overlay.querySelector('#mtgCancel').addEventListener('click', close);
    overlay.querySelector('#mtgSave').addEventListener('click', () => {
      const title = overlay.querySelector('#mtgTitle').value.trim();
      if (!title) { toast('Title required'); return; }
      const date = dateInput.value || todayISO();
      const time = overlay.querySelector('#mtgTime').value || null;
      const isRepeating = repeatChk.checked;
      m.title = title;
      m.time = time;
      m.component = overlay.querySelector('#mtgComp').value || null;
      if (!isRepeating) {
        m.kind = 'oneoff';
        m.date = date;
        m.endDate = null;
        delete m.recurUnit;
        delete m.interval;
        delete m.dayOfWeek;
        delete m.startDate;
      } else {
        const interval = Math.max(1, parseInt(overlay.querySelector('#mtgInterval').value, 10) || 1);
        const unit = unitSel.value === 'day' ? 'day'
                   : unitSel.value === 'month' ? 'month'
                   : 'week';
        const ed = overlay.querySelector('#mtgEndDate').value || null;
        if (ed && ed < date) { toast('End date can\'t be before the start date'); return; }
        m.kind = 'recurring';
        m.recurUnit = unit;
        m.interval  = interval;
        m.startDate = date;
        m.endDate   = ed;
        if (unit === 'week') m.dayOfWeek = parseInt(dowSel.value, 10);
        else delete m.dayOfWeek;
        delete m.date;
      }
      commit('meeting-edit');
      close();
      toast('Saved');
    });
    overlay.querySelector('#mtgDelete').addEventListener('click', () => {
      if (!confirm(`Delete "${m.title}"?`)) return;
      proj.meetings = (proj.meetings || []).filter((x) => x.id !== m.id);
      commit('meeting-delete');
      close();
      toast('Deleted');
    });
  }

  // Persistent filter state for the R&O page
  const roState = { kind: 'all', view: 'list' }; // view: 'list' | 'matrix'

  let _mtxTipEl = null;
  function ensureMtxTipEl() {
    if (_mtxTipEl) return _mtxTipEl;
    _mtxTipEl = document.createElement('div');
    _mtxTipEl.className = 'mtx-tooltip';
    document.body.appendChild(_mtxTipEl);
    return _mtxTipEl;
  }
  function wireRiskMatrixHover(scope) {
    const proj = curProject();
    const svg = scope.querySelector('.mtx-svg');
    if (!svg) return;
    const tip = ensureMtxTipEl();
    function show(circle, x, y) {
      const r = (proj.risks || []).find((rr) => rr.id === circle.dataset.riskId);
      if (!r) return;
      ensureRiskShape(r);
      const isOpp = (r.kind || 'risk') === 'opportunity';
      const stage = circle.dataset.ptStage; // 'inherent' | 'residual'
      const inh = r.inherent, res = r.residual;
      const inhScore = inh.probability * inh.impact;
      const resScore = res.probability * res.impact;
      const linked = r.actionId
        ? state.projects.flatMap((p) => p.actions || []).find((a) => a.id === r.actionId)
        : null;
      tip.innerHTML = `
        <div class="mtx-tip-head">
          <span class="mtx-tip-icon ${isOpp ? 'opp' : 'risk'}">${isOpp ? '▽' : '▲'}</span>
          <span class="mtx-tip-title">${escapeHTML(r.title)}</span>
          <span class="mtx-tip-stage">${stage}</span>
        </div>
        <div class="mtx-tip-row"><span class="mtx-tip-lbl">Inherent</span><span>P${inh.probability} × I${inh.impact} = <b>${inhScore}</b></span></div>
        <div class="mtx-tip-row"><span class="mtx-tip-lbl">Residual</span><span>P${res.probability} × I${res.impact} = <b>${resScore}</b></span></div>
        <div class="mtx-tip-row"><span class="mtx-tip-lbl">Owner</span><span>${escapeHTML(personName(r.owner))}</span></div>
        ${r.mitigation ? `<div class="mtx-tip-mit"><span class="mtx-tip-lbl">${isOpp ? 'Capture' : 'Mitigation'}</span>${escapeHTML(r.mitigation)}</div>` : ''}
        ${linked ? `<div class="mtx-tip-link">↗ ${escapeHTML(linked.title)} <span style="color:var(--text-faint)">— ${escapeHTML(personName(linked.owner))}</span></div>` : ''}`;
      tip.style.display = 'block';
      const rr = tip.getBoundingClientRect();
      let px = x + 14, py = y + 14;
      if (px + rr.width  > innerWidth  - 8) px = x - rr.width - 14;
      if (py + rr.height > innerHeight - 8) py = y - rr.height - 14;
      tip.style.left = Math.max(8, px) + 'px';
      tip.style.top  = Math.max(8, py) + 'px';
    }
    function hide() { tip.style.display = 'none'; }
    svg.addEventListener('mousemove', (e) => {
      const c = e.target.closest('.mtx-pt[data-risk-id]');
      if (!c) { hide(); return; }
      show(c, e.clientX, e.clientY);
    });
    svg.addEventListener('mouseleave', hide);
    svg.addEventListener('contextmenu', (e) => {
      const c = e.target.closest('.mtx-pt[data-risk-id]');
      if (!c) return;
      e.preventDefault();
      hide();
      const id = c.dataset.riskId;
      const r = (proj.risks || []).find((rr) => rr.id === id);
      if (!r) return;
      const isOpp = (r.kind || 'risk') === 'opportunity';
      showContextMenu(e.clientX, e.clientY, [
        { icon: '✎', label: 'Edit…', onClick: () => openRiskEditor(id) },
        { divider: true },
        { icon: '×', label: `Delete ${isOpp ? 'opportunity' : 'risk'}`, danger: true, onClick: () => {
          if (!confirm(`Delete "${r.title}"?`)) return;
          curProject().risks = (curProject().risks || []).filter((x) => x.id !== id);
          commit('risk-delete');
          toast('Deleted');
        }},
      ]);
    });
    // Double-click also opens the editor
    svg.addEventListener('dblclick', (e) => {
      const c = e.target.closest('.mtx-pt[data-risk-id]');
      if (c) openRiskEditor(c.dataset.riskId);
    });
  }

  function openRiskEditor(riskId) {
    const proj = curProject();
    const r = (proj.risks || []).find((x) => x.id === riskId);
    if (!r) return;
    ensureRiskShape(r);
    const isOpp = (r.kind || 'risk') === 'opportunity';
    const overlay = document.createElement('div');
    overlay.className = 'overlay desc-overlay';
    overlay.innerHTML = `
      <div class="desc-modal" style="width:560px;">
        <div class="desc-head">
          <div class="desc-title">${escapeHTML(r.title)} — ${isOpp ? 'opportunity' : 'risk'}</div>
          <button class="icon-btn" id="reClose" title="Close">×</button>
        </div>
        <div style="padding:14px 16px; display:flex; flex-direction:column; gap:10px;">
          <div class="field"><label>Type</label>
            <div class="seg" role="tablist" aria-label="Kind">
              <button type="button" class="seg-btn ${!isOpp ? 'active' : ''}" data-re-kind="risk">▲ Risk</button>
              <button type="button" class="seg-btn ${isOpp ? 'active' : ''}" data-re-kind="opportunity">▽ Opportunity</button>
            </div>
          </div>
          <div class="field"><label>Title</label><input id="reTitle" value="${escapeHTML(r.title)}" /></div>
          <div class="qa-row">
            <div class="field"><label>Inherent P (1-5)</label><input id="rePI" type="number" min="1" max="5" value="${r.inherent.probability}" /></div>
            <div class="field"><label>Inherent I (1-5)</label><input id="reII" type="number" min="1" max="5" value="${r.inherent.impact}" /></div>
          </div>
          <div class="qa-row">
            <div class="field"><label>Residual P (post-action)</label><input id="rePR" type="number" min="1" max="5" value="${r.residual.probability}" /></div>
            <div class="field"><label>Residual I (post-action)</label><input id="reIR" type="number" min="1" max="5" value="${r.residual.impact}" /></div>
          </div>
          <div class="field"><label>Owner</label>
            <select id="reOwner">${state.people.map((p) => `<option value="${p.id}" ${p.id === r.owner ? 'selected' : ''}>${escapeHTML(p.name)}</option>`).join('')}</select>
          </div>
          <div class="field"><label id="reMitLbl">${isOpp ? 'Capture plan' : 'Mitigation'}</label><textarea id="reMit" style="min-height:80px;">${escapeHTML(r.mitigation || '')}</textarea></div>
          <div class="field"><label>Linked action (optional)</label>
            <select id="reActionLink">
              <option value="">— none —</option>
              ${(proj.actions || []).slice().sort((a, b) => a.title.localeCompare(b.title)).map((a) => `<option value="${a.id}" ${a.id === r.actionId ? 'selected' : ''}>${escapeHTML(a.title)} — ${escapeHTML(personName(a.owner))}</option>`).join('')}
            </select>
          </div>
        </div>
        <div class="desc-foot">
          <button class="ghost" id="reCancel">Cancel</button>
          <button class="ghost" id="reDelete" style="margin-left:auto; color:var(--bad);">Delete</button>
          <button class="primary" id="reSave">Save</button>
        </div>
      </div>`;
    document.body.appendChild(overlay);
    setTimeout(() => document.getElementById('reTitle').focus(), 30);

    let kind = isOpp ? 'opportunity' : 'risk';
    overlay.querySelectorAll('.seg-btn[data-re-kind]').forEach((b) => {
      b.addEventListener('click', () => {
        kind = b.dataset.reKind;
        overlay.querySelectorAll('.seg-btn[data-re-kind]').forEach((x) =>
          x.classList.toggle('active', x.dataset.reKind === kind));
        document.getElementById('reMitLbl').textContent =
          kind === 'opportunity' ? 'Capture plan' : 'Mitigation';
      });
    });

    const close = () => overlay.remove();
    overlay.querySelector('#reClose').addEventListener('click', close);
    overlay.querySelector('#reCancel').addEventListener('click', close);
    // Modal closes ONLY via the explicit × button (and Cancel where present).
    // Backdrop click and Escape are intentionally NOT bound — losing
    // in-progress edits to a stray click outside the modal is too easy.

    overlay.querySelector('#reSave').addEventListener('click', () => {
      r.kind = kind;
      r.title = document.getElementById('reTitle').value.trim() || r.title;
      r.inherent = {
        probability: clamp(parseInt(document.getElementById('rePI').value, 10) || 3, 1, 5),
        impact:      clamp(parseInt(document.getElementById('reII').value, 10) || 3, 1, 5),
      };
      r.residual = {
        probability: clamp(parseInt(document.getElementById('rePR').value, 10) || r.inherent.probability, 1, 5),
        impact:      clamp(parseInt(document.getElementById('reIR').value, 10) || r.inherent.impact,      1, 5),
      };
      r.mitigation = document.getElementById('reMit').value;
      r.owner = document.getElementById('reOwner').value;
      r.actionId = document.getElementById('reActionLink').value || null;
      commit('risk-edit');
      close();
      toast('Saved');
    });
    overlay.querySelector('#reDelete').addEventListener('click', () => {
      if (!confirm(`Delete "${r.title}"?`)) return;
      proj.risks = (proj.risks || []).filter((x) => x.id !== r.id);
      commit('risk-delete');
      close();
      toast('Deleted');
    });
  }

  // Helpers — work for both new schema (inherent/residual) and legacy (probability/impact)
  function getInherent(r) {
    if (r.inherent && typeof r.inherent.probability === 'number') return r.inherent;
    return { probability: r.probability || 0, impact: r.impact || 0 };
  }
  function getResidual(r) {
    if (r.residual && typeof r.residual.probability === 'number') return r.residual;
    return getInherent(r);
  }
  function ensureRiskShape(r) {
    if (!r.inherent) r.inherent = { probability: r.probability || 3, impact: r.impact || 3 };
    if (!r.residual) r.residual = { ...r.inherent };
  }

  function riskMatrixSVG(items) {
    const W = 540, H = 460;
    const padL = 60, padR = 20, padT = 30, padB = 56;
    const cellW = (W - padL - padR) / 5;
    const cellH = (H - padT - padB) / 5;

    const cells = [];
    for (let p = 1; p <= 5; p++) {
      for (let i = 1; i <= 5; i++) {
        const score = p * i;
        const cls = score >= 15 ? 'extreme' : score >= 10 ? 'high' : score >= 5 ? 'mid' : 'low';
        const x = padL + (p - 1) * cellW;
        const y = padT + (5 - i) * cellH;
        cells.push(`<rect class="mtx-cell mtx-${cls}" x="${x.toFixed(1)}" y="${y.toFixed(1)}" width="${cellW.toFixed(1)}" height="${cellH.toFixed(1)}" />`);
        cells.push(`<text class="mtx-cell-num" x="${(x + cellW/2).toFixed(1)}" y="${(y + cellH/2 + 4).toFixed(1)}" text-anchor="middle">${score}</text>`);
      }
    }

    const xLabels = [1, 2, 3, 4, 5].map((p) =>
      `<text class="mtx-tick" x="${(padL + (p - 0.5) * cellW).toFixed(1)}" y="${padT + 5 * cellH + 16}" text-anchor="middle">${p}</text>`).join('');
    const yLabels = [1, 2, 3, 4, 5].map((i) =>
      `<text class="mtx-tick" x="${padL - 8}" y="${(padT + (5 - i + 0.5) * cellH + 4).toFixed(1)}" text-anchor="end">${i}</text>`).join('');

    const markers = items.map((r) => {
      const inh = getInherent(r);
      const res = getResidual(r);
      const xInh = padL + (Math.max(1, inh.probability) - 0.5) * cellW;
      const yInh = padT + (5 - Math.max(1, inh.impact) + 0.5) * cellH;
      const xRes = padL + (Math.max(1, res.probability) - 0.5) * cellW;
      const yRes = padT + (5 - Math.max(1, res.impact) + 0.5) * cellH;
      const isOpp = (r.kind || 'risk') === 'opportunity';
      const moved = (xInh !== xRes || yInh !== yRes);
      let svg = '';
      if (moved) {
        svg += `<line class="mtx-arrow ${isOpp ? 'opp' : 'risk'}" x1="${xInh.toFixed(1)}" y1="${yInh.toFixed(1)}" x2="${xRes.toFixed(1)}" y2="${yRes.toFixed(1)}" marker-end="url(#mtxArrow${isOpp ? 'Opp' : 'Risk'})" />`;
      }
      svg += `<circle class="mtx-pt mtx-pt-inh ${isOpp ? 'opp' : 'risk'}" data-risk-id="${r.id}" data-pt-stage="inherent" cx="${xInh.toFixed(1)}" cy="${yInh.toFixed(1)}" r="7"></circle>`;
      if (moved) {
        svg += `<circle class="mtx-pt mtx-pt-res ${isOpp ? 'opp' : 'risk'}" data-risk-id="${r.id}" data-pt-stage="residual" cx="${xRes.toFixed(1)}" cy="${yRes.toFixed(1)}" r="5"></circle>`;
      }
      return svg;
    }).join('');

    return `
      <svg viewBox="0 0 ${W} ${H}" class="mtx-svg" preserveAspectRatio="xMidYMid meet">
        <defs>
          <marker id="mtxArrowRisk" viewBox="0 0 10 10" refX="9" refY="5" markerWidth="7" markerHeight="7" orient="auto-start-reverse">
            <path d="M 0 0 L 10 5 L 0 10 z" fill="rgba(248,113,113,.85)" />
          </marker>
          <marker id="mtxArrowOpp" viewBox="0 0 10 10" refX="9" refY="5" markerWidth="7" markerHeight="7" orient="auto-start-reverse">
            <path d="M 0 0 L 10 5 L 0 10 z" fill="rgba(74,222,128,.85)" />
          </marker>
        </defs>
        ${cells.join('')}
        ${xLabels}
        ${yLabels}
        <text class="mtx-axis-lbl" x="${(padL + (W - padL - padR) / 2).toFixed(1)}" y="${H - 18}" text-anchor="middle">Probability →</text>
        <text class="mtx-axis-lbl" x="${20}" y="${(padT + (H - padT - padB) / 2).toFixed(1)}" text-anchor="middle" transform="rotate(-90 20 ${(padT + (H - padT - padB) / 2).toFixed(1)})">Impact →</text>
        ${markers}
      </svg>
      <div class="mtx-legend">
        <span class="legend-item"><span class="dot mtx-low"></span>Low (≤4)</span>
        <span class="legend-item"><span class="dot mtx-mid"></span>Medium (5-9)</span>
        <span class="legend-item"><span class="dot mtx-high"></span>High (10-14)</span>
        <span class="legend-item"><span class="dot mtx-extreme"></span>Extreme (15+)</span>
        <span class="sep"></span>
        <span class="legend-item"><svg width="22" height="10"><circle cx="6" cy="5" r="4" class="mtx-pt-inh risk"/><line x1="10" y1="5" x2="20" y2="5" stroke="rgba(248,113,113,.7)" stroke-width="1.4" /></svg>inherent → residual</span>
      </div>`;
  }

  function renderRisks(root) {
    const proj = curProject();
    proj.risks = proj.risks || [];
    proj.risks.forEach(ensureRiskShape);
    const view = document.createElement('div');
    view.className = 'view';
    view.innerHTML = `
      <div class="page-head">
        <div>
          <div class="page-title">Risks &amp; Opportunities</div>
          <div class="page-sub">Inherent (raw) vs residual (post-mitigation) on a 5×5 matrix. Risks track downside; opportunities track upside.</div>
        </div>
        <div class="page-actions">
          <div class="seg" role="tablist" aria-label="View">
            <button class="seg-btn ${roState.view === 'list' ? 'active' : ''}" data-ro-view="list">List</button>
            <button class="seg-btn ${roState.view === 'matrix' ? 'active' : ''}" data-ro-view="matrix">Matrix</button>
          </div>
          <div class="seg" role="tablist" aria-label="Kind">
            <button class="seg-btn ${roState.kind === 'all' ? 'active' : ''}" data-kind="all">All</button>
            <button class="seg-btn ${roState.kind === 'risk' ? 'active' : ''}" data-kind="risk">Risks</button>
            <button class="seg-btn ${roState.kind === 'opportunity' ? 'active' : ''}" data-kind="opportunity">Opportunities</button>
          </div>
          <button class="ghost" id="btnAddRisk">+ Risk</button>
          <button class="ghost" id="btnAddOpp">+ Opportunity</button>
        </div>
      </div>
      <div id="roBody"></div>`;
    root.appendChild(view);

    function draw() {
      const body = $('#roBody');
      const items = proj.risks.filter((r) => roState.kind === 'all' || (r.kind || 'risk') === roState.kind);
      if (!items.length) {
        body.innerHTML = `<div class="empty">${
          roState.kind === 'opportunity' ? 'No opportunities logged yet.' :
          roState.kind === 'risk'        ? 'No risks logged yet.' :
                                           'No risks or opportunities logged.'}</div>`;
        return;
      }
      if (roState.view === 'matrix') {
        body.innerHTML = `<div class="panel chart-panel">${riskMatrixSVG(items)}</div>`;
        wireRiskMatrixHover(body);
        return;
      }
      // List view — preserve project array order so drag-reorder is meaningful.
      const sorted = items.slice();
      body.innerHTML = `<div class="row-list" id="roList">${sorted.map((r) => {
        const kind = r.kind || 'rule';
        const inh = getInherent(r);
        const res = getResidual(r);
        const inhScore = inh.probability * inh.impact;
        const resScore = res.probability * res.impact;
        const moved = inhScore !== resScore;
        const sevCls = inhScore >= 12 ? 'high' : inhScore >= 6 ? 'mid' : 'low';
        const icon = kind === 'opportunity' ? '▽' : '△';
        const responseLbl = kind === 'opportunity' ? 'Capture' : 'Mitigation';
        const linkedAction = r.actionId ? state.projects.flatMap((p) => p.actions || []).find((a) => a.id === r.actionId) : null;
        return `
          <div class="row ro-row kind-${kind === 'opportunity' ? 'opportunity' : 'risk'} sev-${sevCls}" data-risk-id="${r.id}">
            ${ROW_GRIP_HTML}
            <span class="ro-icon" title="${kind}">${icon}</span>
            <span class="ro-title">${escapeHTML(r.title)}</span>
            <span class="ro-score">
              ${moved ? `<span class="ro-pre">${inhScore}</span><span class="ro-arrow">→</span><b>${resScore}</b>` : `<b>${inhScore}</b>`}
            </span>
            <span class="row-meta">
              <span class="ro-resp" title="${responseLbl}">${escapeHTML(r.mitigation || '—')}</span>
              ${linkedAction ? `<span class="ro-link clickable" data-open-action="${linkedAction.id}">↗ ${escapeHTML(linkedAction.title)}</span>` : ''}
              <span class="ro-owner">${escapeHTML(personName(r.owner))}</span>
            </span>
          </div>`;
      }).join('')}</div>`;
      // Drag-reorder list rows
      const roListEl = body.querySelector('#roList');
      if (roListEl) {
        wireListReorder(roListEl, {
          rowSelector: '.ro-row[data-risk-id]',
          idAttr: 'riskId',
          getArray: () => proj.risks,
          setOrder: (ids) => { proj.risks = ids.map((id) => proj.risks.find((x) => x.id === id)).filter(Boolean); },
          commitName: 'risks-reorder',
        });
      }
      // Click linked action → drawer
      $$('.ro-link[data-open-action]', body).forEach((el) =>
        el.addEventListener('click', () => openDrawer(el.dataset.openAction)));
      // Right-click on a risk/opportunity row → context menu (Edit / Delete)
      $$('.ro-row[data-risk-id]', body).forEach((row) => {
        row.addEventListener('contextmenu', (e) => {
          e.preventDefault();
          const id = row.dataset.riskId;
          const r = (proj.risks || []).find((x) => x.id === id);
          if (!r) return;
          const isOpp = (r.kind || 'risk') === 'opportunity';
          showContextMenu(e.clientX, e.clientY, [
            { icon: '✎', label: 'Edit…', onClick: () => openRiskEditor(id) },
            { divider: true },
            { icon: '×', label: `Delete ${isOpp ? 'opportunity' : 'risk'}`, danger: true, onClick: () => {
              if (!confirm(`Delete "${r.title}"? This cannot be undone (except via undo).`)) return;
              proj.risks = (proj.risks || []).filter((x) => x.id !== id);
              commit('risk-delete');
              toast('Deleted');
            }},
          ]);
        });
        // Double-click also opens the editor (consistent with kanban / gantt)
        row.addEventListener('dblclick', () => openRiskEditor(row.dataset.riskId));
      });
    }

    $$('.seg-btn[data-kind]', view).forEach((b) => {
      b.addEventListener('click', () => {
        roState.kind = b.dataset.kind;
        $$('.seg-btn[data-kind]', view).forEach((x) => x.classList.toggle('active', x.dataset.kind === roState.kind));
        draw();
      });
    });
    $$('.seg-btn[data-ro-view]', view).forEach((b) => {
      b.addEventListener('click', () => {
        roState.view = b.dataset.roView;
        $$('.seg-btn[data-ro-view]', view).forEach((x) => x.classList.toggle('active', x.dataset.roView === roState.view));
        draw();
      });
    });
    $('#btnAddRisk').addEventListener('click', () => openQuickAdd('risk', { kind: 'risk' }));
    $('#btnAddOpp').addEventListener('click', () => openQuickAdd('risk', { kind: 'opportunity' }));
    draw();
  }

  function renderDecisions(root) {
    const proj = curProject();
    const view = document.createElement('div');
    view.className = 'view';
    view.innerHTML = `
      <div class="page-head">
        <div><div class="page-title">Decisions</div><div class="page-sub">Capture key choices and their rationale.</div></div>
        <div class="page-actions"><button class="ghost" id="btnAddDec">+ Decision</button></div>
      </div>
      <div class="row-list" id="decList"></div>`;
    root.appendChild(view);
    const list = $('#decList');
    if (!proj.decisions?.length) list.innerHTML = '<div class="empty">No decisions logged.</div>';
    else {
      list.innerHTML = proj.decisions.map((d) => `
        <div class="row" data-decision-id="${d.id}">
          ${ROW_GRIP_HTML}
          <span>⬡ ${escapeHTML(d.title)}</span>
          <span class="row-meta">${escapeHTML(personName(d.owner))} • ${d.date || ''}</span>
        </div>`).join('');
      wireListReorder(list, {
        rowSelector: '.row[data-decision-id]',
        idAttr: 'decisionId',
        getArray: () => proj.decisions,
        setOrder: (ids) => { proj.decisions = ids.map((id) => proj.decisions.find((x) => x.id === id)).filter(Boolean); },
        commitName: 'decisions-reorder',
      });
      $$('.row[data-decision-id]', list).forEach((row) => {
        row.addEventListener('contextmenu', (e) => {
          e.preventDefault();
          const id = row.dataset.decisionId;
          const d = (proj.decisions || []).find((x) => x.id === id);
          if (!d) return;
          showContextMenu(e.clientX, e.clientY, [
            { icon: '✎', label: 'Edit…', onClick: () => openDecisionEditor(id) },
            { divider: true },
            { icon: '×', label: 'Delete decision', danger: true, onClick: () => {
              if (!confirm(`Delete "${d.title}"?`)) return;
              proj.decisions = (proj.decisions || []).filter((x) => x.id !== id);
              commit('decision-delete');
              toast('Deleted');
            }},
          ]);
        });
        row.addEventListener('dblclick', () => openDecisionEditor(row.dataset.decisionId));
      });
    }
    $('#btnAddDec').addEventListener('click', () => openQuickAdd('decision'));
  }

  function openDecisionEditor(decisionId) {
    const proj = curProject();
    const d = (proj.decisions || []).find((x) => x.id === decisionId);
    if (!d) return;
    const overlay = document.createElement('div');
    overlay.className = 'overlay desc-overlay';
    overlay.innerHTML = `
      <div class="desc-modal" style="width:520px;">
        <div class="desc-head">
          <div class="desc-title">Edit decision</div>
          <button class="icon-btn" id="deClose" title="Close">×</button>
        </div>
        <div style="padding:14px 16px; display:flex; flex-direction:column; gap:10px;">
          <div class="field"><label>Title</label><input id="deTitle" value="${escapeHTML(d.title)}" /></div>
          <div class="field"><label>Rationale</label><textarea id="deRat" style="min-height:90px;">${escapeHTML(d.rationale || '')}</textarea></div>
          <div class="qa-row">
            <div class="field"><label>Owner</label>
              <select id="deOwner">${state.people.map((p) => `<option value="${p.id}" ${p.id === d.owner ? 'selected' : ''}>${escapeHTML(p.name)}</option>`).join('')}</select>
            </div>
            <div class="field"><label>Date</label><input id="deDate" type="date" value="${d.date || ''}" /></div>
          </div>
        </div>
        <div class="desc-foot">
          <button class="ghost" id="deCancel">Cancel</button>
          <button class="ghost" id="deDelete" style="margin-left:auto; color:var(--bad);">Delete</button>
          <button class="primary" id="deSave">Save</button>
        </div>
      </div>`;
    document.body.appendChild(overlay);
    setTimeout(() => document.getElementById('deTitle').focus(), 30);
    const close = () => overlay.remove();
    overlay.querySelector('#deClose').addEventListener('click', close);
    overlay.querySelector('#deCancel').addEventListener('click', close);
    // Modal closes ONLY via the explicit × button (and Cancel where present).
    // Backdrop click and Escape are intentionally NOT bound — losing
    // in-progress edits to a stray click outside the modal is too easy.
    overlay.querySelector('#deSave').addEventListener('click', () => {
      d.title = document.getElementById('deTitle').value.trim() || d.title;
      d.rationale = document.getElementById('deRat').value;
      d.owner = document.getElementById('deOwner').value;
      d.date = document.getElementById('deDate').value || d.date;
      commit('decision-edit');
      close();
      toast('Saved');
    });
    overlay.querySelector('#deDelete').addEventListener('click', () => {
      if (!confirm(`Delete "${d.title}"?`)) return;
      proj.decisions = (proj.decisions || []).filter((x) => x.id !== d.id);
      commit('decision-delete');
      close();
      toast('Deleted');
    });
  }

  // Reusable contenteditable rich-text editor with an icon-style toolbar:
  // bold, italic, underline, strikethrough, font colour, bullet / numbered list,
  // link, and clear formatting. The element with `id` is the editable region;
  // the toolbar is wired up via `wireRichEditor(overlay, id)`.
  function richEditorHTML(id, valueHtml, placeholder) {
    return `
      <div class="op-context-wrap rich-wrap" data-rich-for="${id}">
        <div class="rich-toolbar">
          <button type="button" class="rt-btn rt-bold"   data-cmd="bold"          title="Bold (Cmd/Ctrl+B)">B</button>
          <button type="button" class="rt-btn rt-ital"   data-cmd="italic"        title="Italic (Cmd/Ctrl+I)">I</button>
          <button type="button" class="rt-btn rt-und"    data-cmd="underline"     title="Underline (Cmd/Ctrl+U)">U</button>
          <button type="button" class="rt-btn rt-strike" data-cmd="strikeThrough" title="Strikethrough">S</button>
          <span class="rt-sep"></span>
          <label class="rt-btn rt-color" title="Text colour">
            <span class="rt-color-glyph">A</span>
            <span class="rt-color-bar"></span>
            <input type="color" data-cmd="foreColor" value="#6ea8ff" />
          </label>
          <span class="rt-sep"></span>
          <button type="button" class="rt-btn rt-ul" data-cmd="insertUnorderedList" title="Bullet list">
            <svg viewBox="0 0 16 16" width="14" height="14"><circle cx="2" cy="3.5" r="1.2" fill="currentColor"/><circle cx="2" cy="8" r="1.2" fill="currentColor"/><circle cx="2" cy="12.5" r="1.2" fill="currentColor"/><line x1="6" y1="3.5" x2="14" y2="3.5" stroke="currentColor" stroke-width="1.4"/><line x1="6" y1="8" x2="14" y2="8" stroke="currentColor" stroke-width="1.4"/><line x1="6" y1="12.5" x2="14" y2="12.5" stroke="currentColor" stroke-width="1.4"/></svg>
          </button>
          <button type="button" class="rt-btn rt-ol" data-cmd="insertOrderedList" title="Numbered list">
            <svg viewBox="0 0 16 16" width="14" height="14"><text x="0" y="5" font-size="4.5" font-weight="700" fill="currentColor">1.</text><text x="0" y="10" font-size="4.5" font-weight="700" fill="currentColor">2.</text><text x="0" y="15" font-size="4.5" font-weight="700" fill="currentColor">3.</text><line x1="6" y1="3.5" x2="14" y2="3.5" stroke="currentColor" stroke-width="1.4"/><line x1="6" y1="8" x2="14" y2="8" stroke="currentColor" stroke-width="1.4"/><line x1="6" y1="12.5" x2="14" y2="12.5" stroke="currentColor" stroke-width="1.4"/></svg>
          </button>
          <span class="rt-sep"></span>
          <button type="button" class="rt-btn rt-link" data-cmd="createLink" title="Insert link">
            <svg viewBox="0 0 16 16" width="14" height="14" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round"><path d="M6.5 9.5l3-3"/><path d="M9 4.5a2.5 2.5 0 0 1 3.5 3.5l-2 2"/><path d="M7 11.5a2.5 2.5 0 0 1-3.5-3.5l2-2"/></svg>
          </button>
          <button type="button" class="rt-btn rt-clear" data-cmd="removeFormat" title="Clear formatting">
            <svg viewBox="0 0 16 16" width="14" height="14" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round"><path d="M3 13l3-9h4l-3 9z"/><line x1="2" y1="13" x2="14" y2="13"/><line x1="11" y1="3" x2="14" y2="6"/><line x1="14" y1="3" x2="11" y2="6"/></svg>
          </button>
        </div>
        <div class="op-context rich-body" id="${id}" contenteditable="true" data-placeholder="${escapeHTML(placeholder || '')}">${valueHtml || ''}</div>
      </div>`;
  }
  function wireRichEditor(overlay, id) {
    const wrap = overlay.querySelector(`[data-rich-for="${id}"]`);
    if (!wrap) return;
    const body = overlay.querySelector(`#${id}`);
    wrap.querySelectorAll('.rt-btn[data-cmd]').forEach((btn) => {
      btn.addEventListener('mousedown', (e) => {
        e.preventDefault();
        const cmd = btn.dataset.cmd;
        body.focus();
        if (cmd === 'createLink') {
          const url = prompt('Link URL (https://…):');
          if (url) document.execCommand('createLink', false, url);
        } else if (cmd === 'foreColor') {
          // handled via colour input change below
        } else {
          document.execCommand(cmd, false, null);
        }
      });
    });
    const color = wrap.querySelector('input[type="color"]');
    const colorBar = wrap.querySelector('.rt-color-bar');
    if (color) {
      color.addEventListener('input', (e) => {
        body.focus();
        document.execCommand('foreColor', false, e.target.value);
        if (colorBar) colorBar.style.background = e.target.value;
      });
      if (colorBar) colorBar.style.background = color.value;
    }
  }

  // Change requests — track scope/schedule/cost change proposals through their lifecycle.
  const crFilterState = { status: 'all' };
  function renderChangeRequests(root) {
    const proj = curProject();
    proj.changes = proj.changes || [];
    const view = document.createElement('div');
    view.className = 'view';
    view.innerHTML = `
      <div class="page-head">
        <div>
          <div class="page-title">Change requests</div>
          <div class="page-sub">Track proposed changes through proposed → reviewed → approved / rejected → implemented. Each CR captures impact (schedule, cost, scope, risk) and an optional online or local link.</div>
        </div>
        <div class="page-actions">
          <div class="seg" role="tablist" aria-label="Status filter">
            <button class="seg-btn ${crFilterState.status === 'all' ? 'active' : ''}" data-cr-filter="all">All</button>
            ${CR_STATUSES.map((s) => `<button class="seg-btn ${crFilterState.status === s.id ? 'active' : ''}" data-cr-filter="${s.id}">${s.label}</button>`).join('')}
          </div>
          <button class="ghost" id="btnAddCR">+ Change request</button>
        </div>
      </div>
      <div class="cr-kpis" id="crKpis"></div>
      <div class="cr-list" id="crList"></div>`;
    root.appendChild(view);

    function drawKpis() {
      const all = proj.changes || [];
      const kpis = $('#crKpis');
      if (!kpis) return;
      if (!all.length) { kpis.innerHTML = ''; return; }
      // Status mix
      const counts = {};
      CR_STATUSES.forEach((s) => { counts[s.id] = 0; });
      all.forEach((c) => { counts[c.status] = (counts[c.status] || 0) + 1; });
      const total = all.length;
      const open = (counts.proposed || 0) + (counts.under_review || 0);
      const decided = (counts.approved || 0) + (counts.rejected || 0) + (counts.implemented || 0) + (counts.cancelled || 0);
      const approvedLike = (counts.approved || 0) + (counts.implemented || 0);
      const approvalRate = decided > 0 ? Math.round((approvedLike / decided) * 100) : 0;
      // Sum schedule + cost impact for approved + implemented (those that "land")
      let schedSum = 0, costSum = 0;
      all.forEach((c) => {
        if (c.status === 'approved' || c.status === 'implemented') {
          schedSum += (c.impact?.schedule || 0);
          costSum  += (c.impact?.cost || 0);
        }
      });
      // Donut SVG — cumulative arcs by status (each arc is clickable to filter)
      const R = 32, CX = 36, CY = 36, IR = 22;
      const C = 2 * Math.PI * R;
      let cum = 0;
      const activeFilter = crFilterState.status;
      const arcs = CR_STATUSES.map((s) => {
        const n = counts[s.id] || 0;
        if (n === 0) return '';
        const frac = n / total;
        const dash = frac * C;
        const gap = C - dash;
        const off = -cum * C + C / 4; // start at 12 o'clock
        cum += frac;
        const dimmed = activeFilter !== 'all' && activeFilter !== s.id;
        const isActive = activeFilter === s.id;
        return `<circle class="cr-donut-slice ${isActive ? 'active' : ''} ${dimmed ? 'dimmed' : ''}" data-cr-status="${s.id}" cx="${CX}" cy="${CY}" r="${R}" fill="none" stroke="rgb(${s.rgb})" stroke-width="${R - IR}" stroke-dasharray="${dash.toFixed(2)} ${gap.toFixed(2)}" stroke-dashoffset="${off.toFixed(2)}" transform="rotate(-90 ${CX} ${CY})"><title>${s.label}: ${n} — click to filter</title></circle>`;
      }).join('');
      const fmtSchedKpi = (n) => n === 0 ? '0 d' : `${n > 0 ? '+' : ''}${n} d`;
      const fmtCostKpi  = (n) => {
        if (n === 0) return '0 €';
        const sign = n > 0 ? '+' : '−';
        const abs = Math.abs(n);
        if (abs >= 1000) return `${sign}${(abs/1000).toFixed(1)}k €`;
        return `${sign}${abs} €`;
      };
      kpis.innerHTML = `
        <div class="cr-kpi cr-kpi-total">
          <div class="cr-kpi-label">Change requests</div>
          <div class="cr-kpi-num">${total}</div>
          <div class="cr-kpi-sub"><span class="cr-kpi-pip" style="background:rgb(${crStatus('proposed').rgb})"></span>${open} open · ${decided} decided</div>
        </div>
        <div class="cr-kpi cr-kpi-mix">
          <div class="cr-kpi-donut">
            <svg viewBox="0 0 72 72" width="72" height="72" aria-hidden="true">
              <circle cx="${CX}" cy="${CY}" r="${R}" fill="none" stroke="var(--bg-3)" stroke-width="${R - IR}" />
              ${arcs}
              <text x="${CX}" y="${CY + 4}" text-anchor="middle" font-size="14" font-weight="700" fill="var(--text)">${total}</text>
            </svg>
          </div>
          <div class="cr-kpi-mix-list">
            ${CR_STATUSES.map((s) => {
              const isActive = activeFilter === s.id;
              const dimmed = activeFilter !== 'all' && !isActive;
              return `
                <div class="cr-kpi-mix-row clickable ${isActive ? 'active' : ''} ${dimmed ? 'dimmed' : ''}" data-cr-status="${s.id}" title="Click to filter">
                  <span class="cr-kpi-pip" style="background:rgb(${s.rgb})"></span>
                  <span class="cr-kpi-mix-lbl">${s.label}</span>
                  <span class="cr-kpi-mix-n">${counts[s.id] || 0}</span>
                </div>`;
            }).join('')}
          </div>
        </div>
        <div class="cr-kpi">
          <div class="cr-kpi-label">Schedule impact</div>
          <div class="cr-kpi-num ${schedSum > 0 ? 'bad' : schedSum < 0 ? 'ok' : ''}">${fmtSchedKpi(schedSum)}</div>
          <div class="cr-kpi-sub">cumulative · approved + implemented</div>
        </div>
        <div class="cr-kpi">
          <div class="cr-kpi-label">Cost impact</div>
          <div class="cr-kpi-num ${costSum > 0 ? 'bad' : costSum < 0 ? 'ok' : ''}">${fmtCostKpi(costSum)}</div>
          <div class="cr-kpi-sub">cumulative · approved + implemented</div>
        </div>
        <div class="cr-kpi">
          <div class="cr-kpi-label">Approval rate</div>
          <div class="cr-kpi-num ${approvalRate >= 60 ? 'ok' : approvalRate < 30 ? 'bad' : ''}">${decided ? approvalRate + '%' : '—'}</div>
          <div class="cr-kpi-sub">${approvedLike} of ${decided || 0} decided</div>
          <div class="cr-kpi-bar"><div class="cr-kpi-bar-fill ${approvalRate >= 60 ? 'ok' : approvalRate < 30 ? 'bad' : 'warn'}" style="width:${decided ? approvalRate : 0}%"></div></div>
        </div>`;

      // Wire click-to-filter on donut arcs and on the legend rows. Clicking the
      // currently-active filter clears it back to "all" (toggle behavior).
      const setFilter = (st) => {
        const next = (crFilterState.status === st) ? 'all' : st;
        crFilterState.status = next;
        // Sync the segmented filter buttons in the page-head
        $$('.seg-btn[data-cr-filter]', view).forEach((x) =>
          x.classList.toggle('active', x.dataset.crFilter === next));
        draw();
      };
      kpis.querySelectorAll('.cr-donut-slice[data-cr-status]').forEach((el) => {
        el.addEventListener('click', () => setFilter(el.dataset.crStatus));
      });
      kpis.querySelectorAll('.cr-kpi-mix-row[data-cr-status]').forEach((el) => {
        el.addEventListener('click', () => setFilter(el.dataset.crStatus));
      });
    }

    function draw() {
      drawKpis();
      const list = $('#crList');
      const items = (proj.changes || []).filter((c) =>
        crFilterState.status === 'all' || c.status === crFilterState.status);
      if (!items.length) {
        list.innerHTML = `<div class="empty">${
          crFilterState.status === 'all'
            ? 'No change requests yet — capture one with + Change request.'
            : 'No change requests in this status.'}</div>`;
        return;
      }
      // Sort by originated date desc (most recent first)
      const sorted = items.slice().sort((a, b) => (b.originatedDate || '').localeCompare(a.originatedDate || ''));
      list.innerHTML = sorted.map((c) => {
        const s = crStatus(c.status);
        const cmp = c.component ? findComponent(proj, c.component) : null;
        const cc = cmp ? componentColor(cmp.color) : null;
        const sched = (c.impact?.schedule ?? 0);
        const cost  = (c.impact?.cost ?? 0);
        const schedTxt = sched ? `${sched > 0 ? '+' : ''}${sched} d` : '—';
        const costTxt  = cost  ? `${cost > 0 ? '+' : ''}${cost} €` : '—';
        const linkUrl = c.linkUrl || '';
        const linkLbl = linkUrl ? (() => { try { return new URL(linkUrl).hostname.replace(/^www\./, ''); } catch (e) { return 'link'; } })() : '';
        const decisionInfo = (c.status === 'approved' || c.status === 'rejected' || c.status === 'implemented' || c.status === 'cancelled')
          ? `<span class="cr-decision">${c.status === 'rejected' || c.status === 'cancelled' ? '✗' : '✓'} ${escapeHTML(personName(c.decisionBy))}${c.decisionDate ? ' · ' + fmtDate(c.decisionDate) : ''}</span>` : '';
        // Rationale supports rich HTML; legacy plain entries render as text
        const rationaleHtml = c.rationale && /<\w+/.test(c.rationale) ? c.rationale : escapeHTML(c.rationale || '');
        const analysisHtml  = c.analysis  && /<\w+/.test(c.analysis)  ? c.analysis  : escapeHTML(c.analysis  || '');
        const prio = priorityLevel(c.priorityLevel);
        const showCrPrioChip = c.priorityLevel && c.priorityLevel !== 'med';
        return `
          <div class="cr-card" data-cr-id="${c.id}">
            <div class="cr-row1">
              <span class="cr-status" style="background:rgba(${s.rgb},.18);color:rgb(${s.rgb});border:1px solid rgb(${s.rgb})">${s.label}</span>
              ${showCrPrioChip ? `<span class="prio-chip prio-${prio.id}" title="Priority: ${prio.label}" style="background:rgba(${prio.rgb},.18);color:rgb(${prio.rgb});border:1px solid rgb(${prio.rgb})">${prio.label}</span>` : ''}
              <span class="cr-title">${escapeHTML(c.title)}</span>
              ${cmp ? `<span class="component-chip" style="background:rgba(${cc.rgb},.2);color:rgb(${cc.rgb})">${escapeHTML(cmp.name)}</span>` : ''}
            </div>
            ${c.rationale ? `<div class="cr-rationale"><span class="cr-section-lbl">Rationale</span><div class="cr-rich">${rationaleHtml}</div></div>` : ''}
            ${c.analysis  ? `<div class="cr-analysis"><span class="cr-section-lbl">Analysis</span><div class="cr-rich">${analysisHtml}</div></div>` : ''}
            <div class="cr-meta">
              <span class="cr-meta-item" title="Originator">⚐ ${escapeHTML(personName(c.originator))}</span>
              <span class="cr-meta-item" title="Originated">▦ ${c.originatedDate ? fmtFull(c.originatedDate) : '—'}</span>
              <span class="cr-meta-item ${sched ? (sched > 0 ? 'bad' : 'ok') : ''}" title="Schedule impact">⏱ ${schedTxt}</span>
              <span class="cr-meta-item ${cost ? (cost > 0 ? 'bad' : 'ok') : ''}" title="Cost impact">€ ${costTxt}</span>
              ${decisionInfo}
              ${linkUrl ? `<a class="cr-link" href="${escapeHTML(linkUrl)}" target="_blank" rel="noopener noreferrer" title="${escapeHTML(linkUrl)}">↗ ${escapeHTML(linkLbl || 'link')}</a>` : ''}
            </div>
          </div>`;
      }).join('');
      $$('.cr-card[data-cr-id]', list).forEach((card) => {
        card.addEventListener('contextmenu', (e) => {
          e.preventDefault();
          const id = card.dataset.crId;
          const c = (proj.changes || []).find((x) => x.id === id);
          if (!c) return;
          const items = [
            { icon: '✎', label: 'Edit…', onClick: () => openChangeRequestEditor(id) },
          ];
          // Status quick-set submenu items
          CR_STATUSES.forEach((s) => {
            if (s.id !== c.status) {
              items.push({
                icon: '◐',
                label: `Mark ${s.label.toLowerCase()}`,
                onClick: () => {
                  const cr = (proj.changes || []).find((x) => x.id === id);
                  if (!cr) return;
                  cr.status = s.id;
                  if (s.id === 'approved' || s.id === 'rejected' || s.id === 'implemented' || s.id === 'cancelled') {
                    cr.decisionDate = todayISO();
                  }
                  commit('cr-status');
                  toast(`Marked ${s.label.toLowerCase()}`);
                },
              });
            }
          });
          if (c.linkUrl) {
            items.push({ icon: '⧉', label: 'Copy link', onClick: () => {
              navigator.clipboard?.writeText(c.linkUrl).then(() => toast('Copied')).catch(() => toast('Copy failed'));
            }});
          }
          items.push({ divider: true });
          items.push({ icon: '×', label: 'Delete change request', danger: true, onClick: () => {
            if (!confirm(`Delete change request "${c.title}"?`)) return;
            proj.changes = (proj.changes || []).filter((x) => x.id !== id);
            commit('cr-delete');
            toast('Deleted');
          }});
          showContextMenu(e.clientX, e.clientY, items);
        });
        card.addEventListener('dblclick', (e) => {
          if (e.target.closest('a')) return;
          openChangeRequestEditor(card.dataset.crId);
        });
      });
    }

    $$('.seg-btn[data-cr-filter]', view).forEach((b) => {
      b.addEventListener('click', () => {
        crFilterState.status = b.dataset.crFilter;
        $$('.seg-btn[data-cr-filter]', view).forEach((x) =>
          x.classList.toggle('active', x.dataset.crFilter === crFilterState.status));
        draw();
      });
    });
    $('#btnAddCR').addEventListener('click', () => {
      if (curProjectIsMerged()) { toast('Pick a single project to add items.'); return; }
      openChangeRequestEditor(newChangeRequestDraft());
    });
    draw();
  }

  function newChangeRequestDraft() {
    return {
      id: uid('cr'),
      title: '',
      rationale: '',
      analysis: '',
      description: '',
      status: 'proposed',
      originator: null,
      originatedDate: todayISO(),
      decisionBy: null,
      decisionDate: null,
      impact: { schedule: 0, cost: 0, scope: '', risk: '' },
      component: null,
      linkUrl: null,
      priorityLevel: 'med',
    };
  }

  function openChangeRequestEditor(crIdOrDraft) {
    const proj = curProject();
    const isDraft = crIdOrDraft && typeof crIdOrDraft === 'object';
    const c = isDraft
      ? crIdOrDraft
      : (proj.changes || []).find((x) => x.id === crIdOrDraft);
    if (!c) return;
    c.impact = c.impact || { schedule: 0, cost: 0, scope: '', risk: '' };
    const overlay = document.createElement('div');
    overlay.className = 'overlay desc-overlay';
    overlay.innerHTML = `
      <div class="desc-modal" style="width:640px;">
        <div class="desc-head">
          <div class="desc-title">${isDraft ? 'New change request' : 'Change request — ' + escapeHTML(c.title)}</div>
          <button class="icon-btn" id="crClose" title="Close">×</button>
        </div>
        <div style="padding:14px 16px; display:flex; flex-direction:column; gap:10px; max-height:72vh; overflow-y:auto;">
          <div class="field"><label>Title</label><input id="crTitle" value="${escapeHTML(c.title)}" /></div>
          <div class="qa-row">
            <div class="field"><label>Status</label>
              <select id="crStatus">
                ${CR_STATUSES.map((s) => `<option value="${s.id}" ${s.id === c.status ? 'selected' : ''}>${s.label}</option>`).join('')}
              </select>
            </div>
            <div class="field"><label>Priority</label>
              <select id="crPriorityLevel">
                ${PRIORITY_LEVELS.map((p) => `<option value="${p.id}" ${p.id === (c.priorityLevel || 'med') ? 'selected' : ''}>${p.label}</option>`).join('')}
              </select>
            </div>
            <div class="field"><label>Originator</label>
              <select id="crOriginator">
                <option value="">—</option>
                ${state.people.map((p) => `<option value="${p.id}" ${p.id === c.originator ? 'selected' : ''}>${escapeHTML(p.name)}</option>`).join('')}
              </select>
            </div>
            <div class="field"><label>Originated</label><input id="crOrigDate" type="date" value="${c.originatedDate || ''}" /></div>
          </div>
          <div class="field"><label>Rationale (rich)</label>
            ${richEditorHTML('crRationale', c.rationale || '', 'Why this change is needed — bold, lists, links, colour…')}
          </div>
          <div class="field"><label>Analysis (rich)</label>
            ${richEditorHTML('crAnalysis', c.analysis || '', 'Trade-off analysis, options considered, recommendation…')}
          </div>
          <div class="field"><label>Description (rich)</label>
            ${richEditorHTML('crDesc', c.description || '', 'Detailed description…')}
          </div>
          <div class="qa-row">
            <div class="field"><label>Schedule impact (days)</label><input id="crSched" type="number" value="${c.impact.schedule || 0}" /></div>
            <div class="field"><label>Cost impact (€)</label><input id="crCost" type="number" value="${c.impact.cost || 0}" /></div>
          </div>
          <div class="field"><label>Scope impact</label><textarea id="crScope" placeholder="What scope changes (added/removed/modified)" style="min-height:50px;">${escapeHTML(c.impact.scope || '')}</textarea></div>
          <div class="field"><label>Risk impact</label><textarea id="crRisk" placeholder="New risks introduced or mitigated" style="min-height:50px;">${escapeHTML(c.impact.risk || '')}</textarea></div>
          <div class="field"><label>Link (URL or path)</label><input id="crLinkUrl" value="${escapeHTML(c.linkUrl || '')}" placeholder="https://… or file:///…" /></div>
          <div class="qa-row">
            <div class="field"><label>Component (optional)</label>
              <select id="crComp">
                <option value="">—</option>
                ${(proj.components || []).map((cmp) => `<option value="${cmp.id}" ${cmp.id === c.component ? 'selected' : ''}>${escapeHTML(cmp.name)}</option>`).join('')}
              </select>
            </div>
            <div class="field"><label>Decision by</label>
              <select id="crDecisionBy">
                <option value="">—</option>
                ${state.people.map((p) => `<option value="${p.id}" ${p.id === c.decisionBy ? 'selected' : ''}>${escapeHTML(p.name)}</option>`).join('')}
              </select>
            </div>
            <div class="field"><label>Decision date</label><input id="crDecisionDate" type="date" value="${c.decisionDate || ''}" /></div>
          </div>
        </div>
        <div class="desc-foot">
          <button class="ghost" id="crCancel">Cancel</button>
          ${isDraft ? '' : '<button class="ghost" id="crDelete" style="margin-left:auto; color:var(--bad);">Delete</button>'}
          <button class="primary" id="crSave">${isDraft ? 'Create' : 'Save'}</button>
        </div>
      </div>`;
    document.body.appendChild(overlay);
    setTimeout(() => document.getElementById('crTitle').focus(), 30);
    // Wire each rich-text editor's toolbar
    wireRichEditor(overlay, 'crRationale');
    wireRichEditor(overlay, 'crAnalysis');
    wireRichEditor(overlay, 'crDesc');
    const close = () => overlay.remove();
    overlay.querySelector('#crClose').addEventListener('click', close);
    overlay.querySelector('#crCancel').addEventListener('click', close);
    // Modal closes ONLY via the explicit × button (and Cancel where present).
    // Backdrop click and Escape are intentionally NOT bound — losing
    // in-progress edits to a stray click outside the modal is too easy.
    overlay.querySelector('#crSave').addEventListener('click', () => {
      const title = document.getElementById('crTitle').value.trim();
      if (!title) return toast('Title required');
      c.title = title;
      const newStatus = document.getElementById('crStatus').value;
      const decisional = (s) => s === 'approved' || s === 'rejected' || s === 'implemented' || s === 'cancelled';
      const wasDecisional = decisional(c.status);
      c.status = newStatus;
      c.originator = document.getElementById('crOriginator').value || null;
      c.originatedDate = document.getElementById('crOrigDate').value || c.originatedDate || todayISO();
      c.rationale = document.getElementById('crRationale').innerHTML;
      c.analysis  = document.getElementById('crAnalysis').innerHTML;
      c.description = document.getElementById('crDesc').innerHTML;
      c.impact = {
        schedule: parseFloat(document.getElementById('crSched').value) || 0,
        cost:     parseFloat(document.getElementById('crCost').value)  || 0,
        scope:    document.getElementById('crScope').value,
        risk:     document.getElementById('crRisk').value,
      };
      c.linkUrl   = document.getElementById('crLinkUrl').value.trim() || null;
      c.priorityLevel = document.getElementById('crPriorityLevel')?.value || c.priorityLevel || 'med';
      c.component = document.getElementById('crComp').value || null;
      c.decisionBy = document.getElementById('crDecisionBy').value || null;
      const dDate = document.getElementById('crDecisionDate').value;
      // Auto-set decision date when entering a decisional status without an explicit date
      if (decisional(newStatus) && !wasDecisional && !dDate) {
        c.decisionDate = todayISO();
      } else {
        c.decisionDate = dDate || c.decisionDate || null;
      }
      if (isDraft) {
        // Hide list filter from masking the new item if needed
        if (crFilterState.status !== 'all' && crFilterState.status !== c.status) crFilterState.status = 'all';
        proj.changes = proj.changes || [];
        proj.changes.push(c);
        commit('cr-create');
        close();
        toast('Created');
      } else {
        commit('cr-edit');
        close();
        toast('Saved');
      }
    });
    if (!isDraft) {
      overlay.querySelector('#crDelete').addEventListener('click', () => {
        if (!confirm(`Delete change request "${c.title}"?`)) return;
        proj.changes = (proj.changes || []).filter((x) => x.id !== c.id);
        commit('cr-delete');
        close();
        toast('Deleted');
      });
    }
  }

  function renderLinks(root) {
    const proj = curProject();
    const isMerged = curProjectIsMerged();
    proj.links = proj.links || [];
    proj.linkFolders = proj.linkFolders || [];
    const view = document.createElement('div');
    view.className = 'view';
    const sub = isMerged
      ? 'All links across every project (read-only — switch to a single project to add or rearrange).'
      : 'Important references — websites, drives, files. Group into folders by drag & drop; reorder anything by its grip.';
    view.innerHTML = `
      <div class="page-head">
        <div><div class="page-title">Links</div><div class="page-sub">${escapeHTML(sub)}</div></div>
        <div class="page-actions">
          ${isMerged ? '' : '<button class="ghost" id="btnNewLinkFolder">+ Folder</button>'}
          ${isMerged ? '' : '<button class="ghost" id="btnAddLink">+ Link</button>'}
        </div>
      </div>
      <div class="link-tree ${isMerged ? 'read-only' : ''}" id="linkTree"></div>`;
    root.appendChild(view);
    const tree = $('#linkTree');

    function linksFor(folderId) {
      return proj.links.filter((l) => (l.folderId || null) === folderId);
    }

    function linkCardHTML(l) {
      const comp = l.component ? findComponent(proj, l.component) : null;
      const compRgb = comp ? componentColor(comp.color).rgb : null;
      const host = (() => { try { return new URL(l.url).hostname.replace(/^www\./, ''); } catch (e) { return l.url; } })();
      return `
        <div class="link-card" data-link-id="${l.id}">
          ${isMerged ? '' : '<span class="row-grip link-grip" title="Drag to reorder or move into a folder" aria-hidden="true">⋮⋮</span>'}
          <a class="link-title" href="${escapeHTML(l.url)}" target="_blank" rel="noopener noreferrer">↗ ${escapeHTML(l.title || host || l.url)}</a>
          <div class="link-host">${escapeHTML(host)}</div>
          ${l.description ? `<div class="link-desc">${escapeHTML(l.description)}</div>` : ''}
          ${comp ? `<div class="link-comp"><span class="link-comp-swatch" style="background:rgb(${compRgb})"></span>${escapeHTML(comp.name)}</div>` : ''}
        </div>`;
    }

    function folderHTML(folder) {
      const items = linksFor(folder.id);
      const collapsed = !!folder.collapsed;
      const emptyMsg = isMerged ? 'No links in this folder.' : 'Empty — drag a link here.';
      return `
        <div class="link-folder ${collapsed ? 'collapsed' : ''}" data-folder-id="${folder.id}">
          <div class="folder-head">
            ${isMerged ? '<span class="folder-caret-spacer" aria-hidden="true"></span>' : '<span class="row-grip folder-grip" title="Drag to reorder folders" aria-hidden="true">⋮⋮</span>'}
            <button type="button" class="folder-caret" title="Toggle">${collapsed ? '▸' : '▾'}</button>
            <span class="folder-name" ${isMerged ? '' : 'contenteditable="true"'} data-placeholder="Folder name…">${escapeHTML(folder.name)}</span>
            <span class="folder-count">${items.length}</span>
            ${isMerged ? '' : '<button type="button" class="folder-add" title="Add link to this folder">+ Link</button>'}
            ${isMerged ? '' : '<button type="button" class="folder-del" title="Delete folder">×</button>'}
          </div>
          <div class="folder-body" data-folder-id="${folder.id}">
            ${items.length ? items.map(linkCardHTML).join('') : `<div class="folder-empty">${emptyMsg}</div>`}
          </div>
        </div>`;
    }

    function ungroupedHTML() {
      const items = linksFor(null);
      return `
        <div class="link-folder loose" data-folder-id="">
          <div class="folder-head loose-head">
            <span class="folder-caret-spacer" aria-hidden="true"></span>
            <span class="folder-name loose-name">Loose links</span>
            <span class="folder-count">${items.length}</span>
          </div>
          <div class="folder-body" data-folder-id="">
            ${items.length ? items.map(linkCardHTML).join('') : '<div class="folder-empty">No loose links — drag any link here to ungroup it.</div>'}
          </div>
        </div>`;
    }

    function drawTree() {
      const total = proj.links.length;
      if (!total && !proj.linkFolders.length) {
        tree.innerHTML = '<div class="empty">No links yet — add one or create a folder to start.</div>';
        return;
      }
      tree.innerHTML = ungroupedHTML() + proj.linkFolders.map(folderHTML).join('');
      wireTree();
    }

    function wireTree() {
      // Folder header interactions
      $$('.link-folder', tree).forEach((folderEl) => {
        const fid = folderEl.dataset.folderId;
        if (fid) {
          // Caret toggle (works in merged mode too — local UI state)
          folderEl.querySelector('.folder-caret')?.addEventListener('click', () => {
            const f = proj.linkFolders.find((x) => x.id === fid);
            if (!f) return;
            f.collapsed = !f.collapsed;
            saveState();
            drawTree();
          });
          if (isMerged) return; // skip rename / delete / drag in merged mode
          // Edit name on blur
          const nameEl = folderEl.querySelector('.folder-name');
          if (nameEl) {
            nameEl.addEventListener('keydown', (e) => {
              if (e.key === 'Enter') { e.preventDefault(); nameEl.blur(); }
            });
            nameEl.addEventListener('blur', () => {
              const f = proj.linkFolders.find((x) => x.id === fid);
              if (!f) return;
              const v = nameEl.textContent.trim() || 'Folder';
              if (f.name !== v) { f.name = v; commit('folder-rename'); }
            });
          }
          // Add link to this folder
          folderEl.querySelector('.folder-add')?.addEventListener('click', () => {
            openQuickAdd('link', { folderId: fid });
          });
          // Delete folder (links inside become loose)
          folderEl.querySelector('.folder-del')?.addEventListener('click', () => {
            const f = proj.linkFolders.find((x) => x.id === fid);
            if (!f) return;
            const childCount = linksFor(fid).length;
            if (!confirm(`Delete folder "${f.name}"?` + (childCount ? `\n\nThe ${childCount} link${childCount === 1 ? '' : 's'} inside will become loose.` : ''))) return;
            proj.links.forEach((l) => { if (l.folderId === fid) l.folderId = null; });
            proj.linkFolders = proj.linkFolders.filter((x) => x.id !== fid);
            commit('folder-delete');
            toast('Folder deleted');
          });
          // Folder reorder drag (via folder-grip)
          const fgrip = folderEl.querySelector('.folder-grip');
          if (fgrip) wireFolderDrag(fgrip, folderEl);
        }
      });
      // Link cards
      $$('.link-card[data-link-id]', tree).forEach((card) => {
        const lid = card.dataset.linkId;
        // Context menu
        card.addEventListener('contextmenu', (e) => {
          e.preventDefault();
          const l = proj.links.find((x) => x.id === lid);
          if (!l) return;
          if (isMerged) {
            // In All Projects view, only allow read-only actions
            showContextMenu(e.clientX, e.clientY, [
              { icon: '⧉', label: 'Copy URL', onClick: () => {
                navigator.clipboard?.writeText(l.url).then(() => toast('Copied')).catch(() => toast('Copy failed'));
              }},
            ]);
            return;
          }
          const moveItems = [];
          if (l.folderId) moveItems.push({ icon: '↩', label: 'Move out (loose)', onClick: () => {
            l.folderId = null; commit('link-move'); toast('Moved out');
          }});
          proj.linkFolders.forEach((f) => {
            if (f.id !== l.folderId) moveItems.push({
              icon: '↪', label: `Move to "${f.name}"`,
              onClick: () => { l.folderId = f.id; commit('link-move'); toast(`Moved to ${f.name}`); },
            });
          });
          showContextMenu(e.clientX, e.clientY, [
            { icon: '✎', label: 'Edit…', onClick: () => openLinkEditor(lid) },
            { icon: '⧉', label: 'Copy URL', onClick: () => {
              navigator.clipboard?.writeText(l.url).then(() => toast('Copied')).catch(() => toast('Copy failed'));
            }},
            ...(moveItems.length ? [{ divider: true }, ...moveItems] : []),
            { divider: true },
            { icon: '×', label: 'Delete link', danger: true, onClick: () => {
              if (!confirm(`Delete "${l.title || l.url}"?`)) return;
              proj.links = proj.links.filter((x) => x.id !== lid);
              commit('link-delete');
              toast('Deleted');
            }},
          ]);
        });
        if (!isMerged) {
          card.addEventListener('dblclick', (e) => {
            if (e.target.closest('a')) return;
            openLinkEditor(lid);
          });
          // Drag to reorder / move between folders
          const grip = card.querySelector('.link-grip');
          if (grip) wireLinkDrag(grip, card);
        }
      });
    }

    // ----------- DRAG: link cards (reorder within + across folders) -----------
    function wireLinkDrag(grip, card) {
      grip.addEventListener('mousedown', (e) => {
        e.preventDefault(); e.stopPropagation();
        const lid = card.dataset.linkId;
        card.classList.add('dragging');
        document.body.classList.add('is-link-dragging');
        let dropTargetFolderId = null;
        let dropAfterLinkId = null; // null means "append to end of target folder"

        const clearIndicators = () => {
          tree.querySelectorAll('.link-folder.drop-target, .link-card.drop-before, .link-card.drop-after, .folder-empty.drop-target')
            .forEach((el) => el.classList.remove('drop-target', 'drop-before', 'drop-after'));
        };

        const onMove = (em) => {
          em.preventDefault();
          clearIndicators();
          // Find what's under the cursor
          const elAt = document.elementFromPoint(em.clientX, em.clientY);
          if (!elAt) return;
          // Drop on another link card → reorder before/after
          const overCard = elAt.closest('.link-card[data-link-id]');
          if (overCard && overCard !== card) {
            const r = overCard.getBoundingClientRect();
            const above = em.clientY < r.top + r.height / 2;
            overCard.classList.add(above ? 'drop-before' : 'drop-after');
            const folderEl = overCard.closest('.link-folder');
            dropTargetFolderId = folderEl ? (folderEl.dataset.folderId || null) : null;
            const overId = overCard.dataset.linkId;
            // dropAfterLinkId is the link directly preceding the new position
            // For "above": position is before overCard → dropAfterLinkId = link before overCard in same folder, or null if overCard is first
            const sameFolderLinks = proj.links.filter((l) => (l.folderId || null) === dropTargetFolderId);
            const overIdx = sameFolderLinks.findIndex((l) => l.id === overId);
            if (above) {
              dropAfterLinkId = overIdx > 0 ? sameFolderLinks[overIdx - 1].id : '__START__';
              if (dropAfterLinkId === lid) dropAfterLinkId = '__SELF_NOOP__';
            } else {
              dropAfterLinkId = overId === lid ? '__SELF_NOOP__' : overId;
            }
            return;
          }
          // Drop on a folder body / header / empty area
          const folderEl = elAt.closest('.link-folder');
          if (folderEl) {
            folderEl.classList.add('drop-target');
            const emptyEl = folderEl.querySelector('.folder-empty');
            if (emptyEl) emptyEl.classList.add('drop-target');
            dropTargetFolderId = folderEl.dataset.folderId || null;
            dropAfterLinkId = null; // append to end
          }
        };

        const onUp = () => {
          card.classList.remove('dragging');
          document.body.classList.remove('is-link-dragging');
          window.removeEventListener('mousemove', onMove);
          window.removeEventListener('mouseup', onUp);
          document.removeEventListener('selectstart', onSelectStart);
          try { window.getSelection()?.removeAllRanges(); } catch (_) {}
          clearIndicators();

          if (dropAfterLinkId === '__SELF_NOOP__') return;

          const link = proj.links.find((l) => l.id === lid);
          if (!link) return;
          const targetFolderId = dropTargetFolderId; // null means loose

          // Update folder membership
          link.folderId = targetFolderId;
          // Reorder proj.links: remove the link, then insert at the right place
          const without = proj.links.filter((l) => l.id !== lid);
          let newIdx;
          if (dropAfterLinkId == null) {
            // Append to the end of the target folder's slice
            // Find last index of any link in target folder
            let lastIdx = -1;
            without.forEach((l, i) => { if ((l.folderId || null) === targetFolderId) lastIdx = i; });
            newIdx = lastIdx + 1;
          } else if (dropAfterLinkId === '__START__') {
            // Insert before the first link of the target folder
            const firstIdx = without.findIndex((l) => (l.folderId || null) === targetFolderId);
            newIdx = firstIdx === -1 ? without.length : firstIdx;
          } else {
            const afterIdx = without.findIndex((l) => l.id === dropAfterLinkId);
            newIdx = afterIdx === -1 ? without.length : afterIdx + 1;
          }
          without.splice(newIdx, 0, link);
          proj.links = without;
          commit('link-reorder');
        };
        const onSelectStart = (ev) => ev.preventDefault();
        window.addEventListener('mousemove', onMove);
        window.addEventListener('mouseup', onUp);
        document.addEventListener('selectstart', onSelectStart);
      });
    }

    // ----------- DRAG: folder reorder (folders only, not into other folders) -----------
    function wireFolderDrag(grip, folderEl) {
      grip.addEventListener('mousedown', (e) => {
        e.preventDefault(); e.stopPropagation();
        folderEl.classList.add('dragging');
        document.body.classList.add('is-folder-dragging');

        const onMove = (em) => {
          em.preventDefault();
          // Reorder among siblings of the same parent
          const sibs = [...tree.querySelectorAll('.link-folder:not(.loose):not(.dragging)')];
          const after = sibs.find((sib) => {
            const r = sib.getBoundingClientRect();
            return em.clientY < r.top + r.height / 2;
          });
          if (after) tree.insertBefore(folderEl, after);
          else tree.appendChild(folderEl);
        };
        const onUp = () => {
          folderEl.classList.remove('dragging');
          document.body.classList.remove('is-folder-dragging');
          window.removeEventListener('mousemove', onMove);
          window.removeEventListener('mouseup', onUp);
          document.removeEventListener('selectstart', onSelectStart);
          try { window.getSelection()?.removeAllRanges(); } catch (_) {}
          const newOrder = [...tree.querySelectorAll('.link-folder:not(.loose)')].map((el) => el.dataset.folderId);
          const before = proj.linkFolders.map((f) => f.id).join(',');
          proj.linkFolders = newOrder.map((fid) => proj.linkFolders.find((f) => f.id === fid)).filter(Boolean);
          if (proj.linkFolders.map((f) => f.id).join(',') !== before) commit('folder-reorder');
        };
        const onSelectStart = (ev) => ev.preventDefault();
        window.addEventListener('mousemove', onMove);
        window.addEventListener('mouseup', onUp);
        document.addEventListener('selectstart', onSelectStart);
      });
    }

    drawTree();

    $('#btnNewLinkFolder')?.addEventListener('click', () => {
      proj.linkFolders.push({ id: uid('lf'), name: 'New folder', collapsed: false });
      commit('folder-create');
      toast('Folder added');
    });
    $('#btnAddLink')?.addEventListener('click', () => openQuickAdd('link'));
  }

  function openLinkEditor(linkId) {
    const proj = curProject();
    const l = (proj.links || []).find((x) => x.id === linkId);
    if (!l) return;
    const overlay = document.createElement('div');
    overlay.className = 'overlay desc-overlay';
    overlay.innerHTML = `
      <div class="desc-modal" style="width:560px;">
        <div class="desc-head">
          <div class="desc-title">Edit link</div>
          <button class="icon-btn" id="lnClose" title="Close">×</button>
        </div>
        <div style="padding:14px 16px; display:flex; flex-direction:column; gap:10px;">
          <div class="field"><label>Title</label><input id="lnTitle" value="${escapeHTML(l.title || '')}" /></div>
          <div class="field"><label>URL or path</label><input id="lnUrl" value="${escapeHTML(l.url || '')}" placeholder="https://… or file:///…" /></div>
          <div class="field"><label>Description</label><textarea id="lnDesc" style="min-height:80px;">${escapeHTML(l.description || '')}</textarea></div>
          <div class="qa-row">
            <div class="field"><label>Component (optional)</label>
              <select id="lnComp">
                <option value="">— none —</option>
                ${(proj.components || []).map((c) => `<option value="${c.id}" ${c.id === l.component ? 'selected' : ''}>${escapeHTML(c.name)}</option>`).join('')}
              </select>
            </div>
            <div class="field"><label>Folder</label>
              <select id="lnFolder">
                <option value="">— Loose —</option>
                ${(proj.linkFolders || []).map((f) => `<option value="${f.id}" ${f.id === l.folderId ? 'selected' : ''}>${escapeHTML(f.name)}</option>`).join('')}
              </select>
            </div>
          </div>
        </div>
        <div class="desc-foot">
          <button class="ghost" id="lnCancel">Cancel</button>
          <button class="ghost" id="lnDelete" style="margin-left:auto; color:var(--bad);">Delete</button>
          <button class="primary" id="lnSave">Save</button>
        </div>
      </div>`;
    document.body.appendChild(overlay);
    setTimeout(() => document.getElementById('lnTitle').focus(), 30);
    const close = () => overlay.remove();
    overlay.querySelector('#lnClose').addEventListener('click', close);
    overlay.querySelector('#lnCancel').addEventListener('click', close);
    // Modal closes ONLY via the explicit × button (and Cancel where present).
    // Backdrop click and Escape are intentionally NOT bound — losing
    // in-progress edits to a stray click outside the modal is too easy.
    overlay.querySelector('#lnSave').addEventListener('click', () => {
      const url = document.getElementById('lnUrl').value.trim();
      if (!url) return toast('URL required');
      l.title = document.getElementById('lnTitle').value.trim() || url;
      l.url = url;
      l.description = document.getElementById('lnDesc').value;
      l.component = document.getElementById('lnComp').value || null;
      l.folderId = document.getElementById('lnFolder')?.value || null;
      commit('link-edit');
      close();
      toast('Saved');
    });
    overlay.querySelector('#lnDelete').addEventListener('click', () => {
      if (!confirm(`Delete "${l.title || l.url}"?`)) return;
      proj.links = (proj.links || []).filter((x) => x.id !== l.id);
      commit('link-delete');
      close();
      toast('Deleted');
    });
  }

  function renderComponents(root) {
    const proj = curProject();
    proj.components = proj.components || [];
    const view = document.createElement('div');
    view.className = 'view';
    view.innerHTML = `
      <div class="page-head">
        <div><div class="page-title">Components</div><div class="page-sub">Logical components or work packages — colour cards on the board.</div></div>
        <div class="page-actions"><button class="ghost" id="btnAddComponent">+ Component</button></div>
      </div>
      <div class="row-list" id="componentList"></div>`;
    root.appendChild(view);
    const list = $('#componentList');
    if (!proj.components.length) {
      list.innerHTML = '<div class="empty">No components yet — add one to start colour-coding actions.</div>';
    } else {
      const knownCCs = getCostCentres();
      list.innerHTML = proj.components.map((pt) => {
        const c = componentColor(pt.color);
        const count = (proj.actions || []).filter((a) => a.component === pt.id).length;
        const ccOptions = [...new Set([...knownCCs, ...(pt.costCenter ? [pt.costCenter] : [])])];
        return `
          <div class="row" data-component-id="${pt.id}">
            ${ROW_GRIP_HTML}
            <span class="component-swatch" style="background: rgba(${c.rgb},.9);"></span>
            <input class="inline component-name" value="${escapeHTML(pt.name)}" />
            <select class="inline component-color">
              ${COMPONENT_COLORS.map((co) => `<option value="${co.id}" ${co.id === pt.color ? 'selected' : ''}>${co.name}</option>`).join('')}
            </select>
            <select class="inline component-cc" title="Cost centre">
              <option value="">— none —</option>
              ${ccOptions.map((code) => `<option value="${escapeHTML(code)}" ${code === pt.costCenter ? 'selected' : ''}>${escapeHTML(code)}</option>`).join('')}
              <option value="__new__">+ New cost centre…</option>
            </select>
            <span class="row-meta">${count} action${count === 1 ? '' : 's'}</span>
            <button class="icon-btn component-del" title="Delete">×</button>
          </div>`;
      }).join('');
      wireListReorder(list, {
        rowSelector: '.row[data-component-id]',
        idAttr: 'componentId',
        getArray: () => proj.components,
        setOrder: (ids) => { proj.components = ids.map((id) => proj.components.find((x) => x.id === id)).filter(Boolean); },
        commitName: 'components-reorder',
      });
      $$('.component-name', list).forEach((inp) => {
        inp.addEventListener('change', () => {
          const id = inp.closest('[data-component-id]').dataset.componentId;
          const pt = proj.components.find((p) => p.id === id);
          if (pt) { pt.name = inp.value.trim() || pt.name; commit('component-rename'); }
        });
      });
      $$('.component-color', list).forEach((sel) => {
        sel.addEventListener('change', () => {
          const id = sel.closest('[data-component-id]').dataset.componentId;
          const pt = proj.components.find((p) => p.id === id);
          if (pt) { pt.color = sel.value; commit('component-recolor'); }
        });
      });
      $$('.component-cc', list).forEach((sel) => {
        sel.addEventListener('change', () => {
          const id = sel.closest('[data-component-id]').dataset.componentId;
          const pt = proj.components.find((p) => p.id === id);
          if (!pt) return;
          let v = sel.value;
          if (v === '__new__') {
            const code = prompt('New cost-centre name:');
            if (!code) { sel.value = pt.costCenter || ''; return; }
            const trimmed = code.trim();
            if (!trimmed) { sel.value = pt.costCenter || ''; return; }
            state.budgets = state.budgets || {};
            state.budgets[trimmed] = state.budgets[trimmed] || {};
            v = trimmed;
          }
          v = v || null;
          if (pt.costCenter !== v) {
            pt.costCenter = v;
            commit('component-cc');
          }
        });
      });
      $$('.component-del', list).forEach((btn) => {
        btn.addEventListener('click', () => {
          const id = btn.closest('[data-component-id]').dataset.componentId;
          const pt = proj.components.find((p) => p.id === id);
          const count = (proj.actions || []).filter((a) => a.component === id).length;
          if (!confirm(`Delete component "${pt?.name}"?` + (count ? ` (${count} action${count === 1 ? '' : 's'} will become unassigned)` : ''))) return;
          proj.components = proj.components.filter((p) => p.id !== id);
          (proj.actions || []).forEach((a) => { if (a.component === id) a.component = null; });
          commit('component-delete');
        });
      });
    }
    $('#btnAddComponent').addEventListener('click', () => openQuickAdd("component"));
  }

  // EVM KPI hover tooltip — single floating element, populated on demand
  let evmTipEl = null;
  function ensureEvmTipEl() {
    if (evmTipEl) return evmTipEl;
    evmTipEl = document.createElement('div');
    evmTipEl.className = 'evm-tooltip';
    document.body.appendChild(evmTipEl);
    return evmTipEl;
  }
  function showEvmTooltip(metric, x, y) {
    const def = EVM_DEFS[metric];
    if (!def) return;
    const tip = ensureEvmTipEl();
    tip.innerHTML = `
      <div class="evm-name">${escapeHTML(metric)} — ${escapeHTML(def.name)}</div>
      <div class="evm-formula"><span class="evm-lbl">Formula</span><code>${escapeHTML(def.formula)}</code></div>
      <div class="evm-desc">${escapeHTML(def.desc)}</div>`;
    tip.style.display = 'block';
    const r = tip.getBoundingClientRect();
    let px = x + 14, py = y + 14;
    if (px + r.width  > innerWidth  - 8) px = x - r.width - 14;
    if (py + r.height > innerHeight - 8) py = y - r.height - 14;
    tip.style.left = Math.max(8, px) + 'px';
    tip.style.top  = Math.max(8, py) + 'px';
  }
  function hideEvmTooltip() { if (evmTipEl) evmTipEl.style.display = 'none'; }

  function wireEvmTooltipsOnce() {
    if (document._evmWired) return;
    document._evmWired = true;
    document.addEventListener('mouseover', (e) => {
      const card = e.target.closest('.bk[data-evm]');
      if (card) showEvmTooltip(card.dataset.evm, e.clientX, e.clientY);
    });
    document.addEventListener('mousemove', (e) => {
      if (!evmTipEl || evmTipEl.style.display === 'none') return;
      const card = e.target.closest('.bk[data-evm]');
      if (!card) { hideEvmTooltip(); return; }
      // keep showing — re-position smoothly
      showEvmTooltip(card.dataset.evm, e.clientX, e.clientY);
    });
    document.addEventListener('mouseout', (e) => {
      const card = e.target.closest('.bk[data-evm]');
      if (!card) return;
      const next = e.relatedTarget;
      if (next && card.contains(next)) return;
      hideEvmTooltip();
    });
  }

  /* ------------------------ Cost-centre budgets ---------------------- */

  // Hours per FTE-week. Used to convert commitment % → planned hours.
  const HOURS_PER_WEEK = 40;

  // Earned-Value Management metric definitions (used for the budget KPI tooltips)
  const EVM_DEFS = {
    BAC: { name: 'Budget at Completion',     formula: 'Σ weekly budget across the horizon',
           desc: 'The total approved budget for the cost centre.' },
    PV:  { name: 'Planned Value',             formula: 'Σ budget up to today',
           desc: 'How much work, in budgeted terms, was planned to be done by now.' },
    AC:  { name: 'Actual Cost',               formula: 'Σ actual cost up to today',
           desc: 'Cost actually incurred so far, summed from action commitments × hourly rates.' },
    EV:  { name: 'Earned Value',              formula: 'Σ planned cost × % complete',
           desc: 'The budgeted value of the work actually performed. Done actions credit full cost; in-flight ones credit by elapsed/total duration.' },
    CV:  { name: 'Cost Variance',             formula: 'EV − AC',
           desc: 'Positive = under budget for the work done. Negative = over budget.' },
    SV:  { name: 'Schedule Variance',         formula: 'EV − PV',
           desc: 'Positive = ahead of schedule. Negative = behind: less work delivered than planned.' },
    CPI: { name: 'Cost Performance Index',    formula: 'EV ÷ AC',
           desc: '> 1: getting more value per € spent than budgeted. < 1: cost overrun.' },
    SPI: { name: 'Schedule Performance Index', formula: 'EV ÷ PV',
           desc: '> 1: ahead of schedule. < 1: behind schedule.' },
    EAC: { name: 'Estimate at Completion',    formula: 'BAC ÷ CPI',
           desc: 'Forecast of the total cost at the end of the horizon, extrapolating current cost performance.' },
    ETC: { name: 'Estimate to Complete',      formula: 'EAC − AC',
           desc: 'Forecast of remaining spend from today to completion.' },
    VAC: { name: 'Variance at Completion',    formula: 'BAC − EAC',
           desc: 'Forecast surplus (positive) or shortfall (negative) at completion vs the original budget.' },
  };
  function personRate(personId) {
    const p = state.people.find((x) => x.id === personId);
    return (p && typeof p.hourlyRate === 'number') ? p.hourlyRate : 100;
  }
  function avgHourlyRate() {
    const ppl = state.people || [];
    if (!ppl.length) return 100;
    return ppl.reduce((s, p) => s + (p.hourlyRate || 100), 0) / ppl.length;
  }

  function getCostCentres() {
    const inScope = projectsInScope();
    const set = new Set();
    inScope.forEach((p) => {
      (p.components || []).forEach((c) => { if (c.costCenter) set.add(c.costCenter); });
      // Also include CCs explicitly declared on the project (e.g. via "+ Cost centre"
      // before any component is mapped) so they're visible in the Budgets page.
      (p.costCenters || []).forEach((cc) => { if (cc) set.add(cc); });
    });
    return [...set].sort();
  }

  function componentsForCC(cc) {
    const out = [];
    projectsInScope().forEach((p) => (p.components || []).forEach((c) => {
      if (c.costCenter === cc) out.push({ proj: p, component: c });
    }));
    return out;
  }

  function actionsForCC(cc) {
    const ids = new Set(componentsForCC(cc).map(({ component }) => component.id));
    const out = [];
    projectsInScope().forEach((p) => (p.actions || []).forEach((a) => {
      if (a.deletedAt) return;
      if (ids.has(a.component)) out.push(a);
    }));
    return out;
  }

  function actionPlannedHours(a) {
    if (!a.due) return 0;
    const cmt = (typeof a.commitment === 'number') ? a.commitment : 100;
    const startD = a.startDate ? parseDate(a.startDate) : new Date(parseDate(a.due).getTime() - 2 * dayMs);
    const endD = parseDate(a.due);
    const days = Math.max(1, Math.round((endD - startD) / dayMs) + 1);
    return (cmt / 100) * HOURS_PER_WEEK * (days / 7);
  }
  function actionPlannedCost(a) {
    return actionPlannedHours(a) * personRate(a.owner);
  }

  function _ccWeekOverlap(cc, weekStartISO) {
    const wsT = parseDate(weekStartISO).getTime();
    const weT = wsT + 7 * dayMs - 1;
    const out = [];
    actionsForCC(cc).forEach((a) => {
      if (!a.due) return;
      const startT = a.startDate
        ? parseDate(a.startDate).getTime()
        : parseDate(a.due).getTime() - 2 * dayMs;
      const endT = parseDate(a.due).getTime();
      if (endT < wsT || startT > weT) return;
      const overlapStart = Math.max(startT, wsT);
      const overlapEnd   = Math.min(endT,   weT);
      const overlapDays  = Math.max(1, Math.round((overlapEnd - overlapStart) / dayMs) + 1);
      const cmt = (typeof a.commitment === 'number') ? a.commitment : 100;
      const hours = (cmt / 100) * HOURS_PER_WEEK * (overlapDays / 7);
      out.push({ a, hours, cost: hours * personRate(a.owner) });
    });
    return out;
  }
  function actualHoursForCCWeek(cc, weekStartISO) {
    return _ccWeekOverlap(cc, weekStartISO).reduce((s, x) => s + x.hours, 0);
  }
  function actualCostForCCWeek(cc, weekStartISO) {
    return _ccWeekOverlap(cc, weekStartISO).reduce((s, x) => s + x.cost, 0);
  }

  // Per-component breakdown for a CC for a given week.
  // Returns { [componentId]: { hours, cost, name, color } }
  function _ccComponentBreakdownForWeek(cc, weekStartISO) {
    const out = {};
    componentsForCC(cc).forEach(({ component }) => {
      out[component.id] = { hours: 0, cost: 0, name: component.name, color: component.color };
    });
    _ccWeekOverlap(cc, weekStartISO).forEach((x) => {
      const cid = x.a.component;
      if (!out[cid]) return;
      out[cid].hours += x.hours;
      out[cid].cost  += x.cost;
    });
    return out;
  }

  // Per-person breakdown for a CC for a given week.
  function _ccPersonBreakdownForWeek(cc, weekStartISO) {
    const out = {};
    _ccWeekOverlap(cc, weekStartISO).forEach((x) => {
      const pid = x.a.owner;
      if (!out[pid]) {
        const p = state.people.find((pp) => pp.id === pid);
        out[pid] = { hours: 0, cost: 0, name: p?.name || 'Unassigned', id: pid };
      }
      out[pid].hours += x.hours;
      out[pid].cost  += x.cost;
    });
    return out;
  }
  function peopleForCC(cc) {
    const set = new Set();
    actionsForCC(cc).forEach((a) => { if (a.owner) set.add(a.owner); });
    return [...set].map((pid) => state.people.find((p) => p.id === pid)).filter(Boolean);
  }

  // Budgets are stored in HOURS (canonical). evmFor returns measurements in
  // the requested mode ('cost' or 'hours').
  function evmFor(cc, weeks, mode) {
    state.budgets = state.budgets || {};
    state.budgets[cc] = state.budgets[cc] || {};
    const today = new Date(); today.setHours(0, 0, 0, 0);
    const todayT = today.getTime();
    const isCost = mode === 'cost';
    const avg = avgHourlyRate();
    const conv = (h) => isCost ? h * avg : h;
    const budgetsHours = weeks.map((w) => Number(state.budgets[cc][w.isoStart] || 0));
    const budgets = budgetsHours.map(conv);
    const actuals = weeks.map((w) => isCost
      ? actualCostForCCWeek(cc, w.isoStart)
      : actualHoursForCCWeek(cc, w.isoStart));
    const BAC = budgets.reduce((s, v) => s + v, 0);
    let PV = 0, AC = 0;
    weeks.forEach((w, i) => {
      if (w.start.getTime() <= todayT) { PV += budgets[i]; AC += actuals[i]; }
    });
    let EV = 0;
    actionsForCC(cc).forEach((a) => {
      const planned = isCost ? actionPlannedCost(a) : actionPlannedHours(a);
      if (a.status === 'done') { EV += planned; return; }
      if (!a.startDate || !a.due) return;
      const startT = parseDate(a.startDate).getTime();
      const endT   = parseDate(a.due).getTime();
      if (todayT <= startT) return;
      const elapsed = Math.min(todayT, endT) - startT;
      const dur     = Math.max(1, endT - startT);
      EV += planned * (elapsed / dur);
    });
    const CV  = EV - AC;
    const SV  = EV - PV;
    const CPI = AC > 0 ? EV / AC : 1;
    const SPI = PV > 0 ? EV / PV : 1;
    const EAC = CPI > 0 ? BAC / CPI : BAC;
    const ETC = Math.max(0, EAC - AC);
    const VAC = BAC - EAC;
    return { BAC, PV, AC, EV, CV, SV, CPI, SPI, EAC, ETC, VAC, budgets, actuals, mode };
  }

  function fmtBudgetVal(v, mode) {
    if (mode === 'hours') return v >= 100 ? Math.round(v) + 'h' : v.toFixed(1) + 'h';
    if (v === 0) return '0 €';
    if (Math.abs(v) >= 1000) return (v / 1000).toFixed(1) + 'k €';
    return Math.round(v) + ' €';
  }
  function fmtMoney(v) { return fmtBudgetVal(v, 'cost'); }

  function renderBudgets(root) {
    state.budgets = state.budgets || {};
    state.settings = state.settings || {};
    if (state.settings.budgetView !== 'hours' && state.settings.budgetView !== 'cost') {
      state.settings.budgetView = 'cost';
    }
    if (state.settings.budgetGroupBy !== 'component' && state.settings.budgetGroupBy !== 'person') {
      state.settings.budgetGroupBy = 'component';
    }

    const view = document.createElement('div');
    view.className = 'view';
    const ccs = getCostCentres();
    view.innerHTML = `
      <div class="page-head">
        <div>
          <div class="page-title">Cost-centre budgets</div>
          <div class="page-sub">Drag the green points to set a weekly budget. Bars show actual workload on components mapped to the cost centre. Toggle between € and hours; rates are per-person (edit on the People page).</div>
        </div>
        <div class="page-actions">
          <div class="seg" role="tablist" aria-label="Group by">
            <button class="seg-btn ${state.settings.budgetGroupBy === 'component' ? 'active' : ''}" data-budget-group="component">By component</button>
            <button class="seg-btn ${state.settings.budgetGroupBy === 'person' ? 'active' : ''}" data-budget-group="person">By person</button>
          </div>
          <div class="seg" role="tablist" aria-label="Unit">
            <button class="seg-btn ${state.settings.budgetView === 'cost' ? 'active' : ''}" data-budget-view="cost">€ Cost</button>
            <button class="seg-btn ${state.settings.budgetView === 'hours' ? 'active' : ''}" data-budget-view="hours">⏱ Hours</button>
          </div>
          <button class="primary" id="btnNewCC">+ Cost centre</button>
        </div>
      </div>
      ${ccs.length ? '' : '<div class="empty">No cost centres yet — click <b>+ Cost centre</b> above, or add one to a component on the <b>Components</b> page.</div>'}
      <div id="budgetsList"></div>`;
    root.appendChild(view);

    $$('.seg-btn[data-budget-view]', view).forEach((b) => {
      b.addEventListener('click', () => {
        state.settings.budgetView = b.dataset.budgetView;
        saveState();
        $$('.seg-btn[data-budget-view]', view).forEach((x) => x.classList.toggle('active', x.dataset.budgetView === state.settings.budgetView));
        drawList();
      });
    });
    $$('.seg-btn[data-budget-group]', view).forEach((b) => {
      b.addEventListener('click', () => {
        state.settings.budgetGroupBy = b.dataset.budgetGroup;
        saveState();
        $$('.seg-btn[data-budget-group]', view).forEach((x) => x.classList.toggle('active', x.dataset.budgetGroup === state.settings.budgetGroupBy));
        drawList();
      });
    });
    $('#btnNewCC').addEventListener('click', () => {
      if (curProjectIsMerged()) { toast('Pick a single project first'); return; }
      const code = prompt('New cost-centre name:');
      if (!code) return;
      const trimmed = code.trim();
      if (!trimmed) return;
      if (ccs.includes(trimmed)) { toast('Already exists in this project'); return; }
      state.budgets[trimmed] = state.budgets[trimmed] || {};
      const proj = curProject();
      proj.costCenters = proj.costCenters || [];
      if (!proj.costCenters.includes(trimmed)) proj.costCenters.push(trimmed);
      saveState();
      render();
    });

    function weekStarts(n = 52) {
      const today = new Date(); today.setHours(0, 0, 0, 0);
      let monday = new Date(today);
      while (monday.getDay() !== 1) monday = new Date(monday.getTime() - dayMs);
      const start = new Date(monday.getTime() - 12 * 7 * dayMs); // 12w past + 40w future
      const out = [];
      for (let i = 0; i < n; i++) {
        const s = new Date(start.getTime() + i * 7 * dayMs);
        out.push({ start: s, isoStart: fmtISO(s) });
      }
      return out;
    }

    function drawList() {
      const list = $('#budgetsList');
      if (!ccs.length) { list.innerHTML = ''; return; }
      const weeks = weekStarts(52);
      list.innerHTML = ccs.map((cc) => {
        const safe = cc.replace(/[^A-Za-z0-9_-]/g, '_');
        return `
          <div class="budget-card" data-cc="${escapeHTML(cc)}">
            <div class="budget-head">
              <h3>${escapeHTML(cc)}</h3>
              <div class="budget-legend" id="legend-${safe}"></div>
            </div>
            <div class="budget-kpis" id="kpis-${safe}"></div>
            <div class="budget-chart-wrap"><svg class="budget-chart" id="chart-${safe}"></svg></div>
          </div>`;
      }).join('');
      ccs.forEach((cc) => drawChart(cc, weeks));
    }

    function drawChart(cc, weeks) {
      const safe = cc.replace(/[^A-Za-z0-9_-]/g, '_');
      const mode = state.settings.budgetView;
      const evm = evmFor(cc, weeks, mode);
      const fmtV = (v) => fmtBudgetVal(v, mode);
      const kpiEl = document.getElementById(`kpis-${safe}`);
      const k = (l, v, cls = '') => `<div class="bk" data-evm="${l}"><div class="bk-l">${l}</div><div class="bk-v ${cls}">${v}</div></div>`;
      kpiEl.innerHTML = [
        k('BAC', fmtV(evm.BAC)),
        k('PV',  fmtV(evm.PV)),
        k('AC',  fmtV(evm.AC)),
        k('EV',  fmtV(evm.EV)),
        k('CV',  (evm.CV >= 0 ? '+' : '') + fmtV(evm.CV), evm.CV >= 0 ? 'ok' : 'bad'),
        k('SV',  (evm.SV >= 0 ? '+' : '') + fmtV(evm.SV), evm.SV >= 0 ? 'ok' : 'bad'),
        k('CPI', evm.CPI.toFixed(2), evm.CPI >= 1 ? 'ok' : 'bad'),
        k('SPI', evm.SPI.toFixed(2), evm.SPI >= 1 ? 'ok' : 'bad'),
        k('EAC', fmtV(evm.EAC), evm.EAC > evm.BAC ? 'bad' : 'ok'),
        k('ETC', fmtV(evm.ETC)),
        k('VAC', (evm.VAC >= 0 ? '+' : '') + fmtV(evm.VAC), evm.VAC >= 0 ? 'ok' : 'bad'),
      ].join('');

      const svg = document.getElementById(`chart-${safe}`);
      const W = 1100, H = 240;
      const padL = 64, padR = 18, padT = 14, padB = 32;
      const innerW = W - padL - padR, innerH = H - padT - padB;
      svg.setAttribute('viewBox', `0 0 ${W} ${H}`);
      // 'none' makes the SVG stretch to fill its rendered box exactly, so
      // screen-X / screen-Y → viewBox coords is a straight scale (no letterbox).
      svg.setAttribute('preserveAspectRatio', 'none');

      // Tight Y range — fit data plus ~12% headroom; tiny floor so an empty CC
      // doesn't render a degenerate axis.
      const dataMax = Math.max(0, ...evm.budgets, ...evm.actuals);
      const maxData = dataMax > 0 ? dataMax : (mode === 'cost' ? 100 : 5);
      const maxY = maxData * 1.12;
      const xLeft = (i) => padL + (i / weeks.length) * innerW;
      const xMid  = (i) => padL + ((i + 0.5) / weeks.length) * innerW;
      const yFor  = (v) => padT + innerH - (v / maxY) * innerH;
      const barW = innerW / weeks.length;

      const today = new Date(); today.setHours(0, 0, 0, 0);
      const todayT = today.getTime();
      const todayIdx = weeks.findIndex((w) => w.start.getTime() === todayT);

      const isCostMode = mode === 'cost';
      const groupBy = state.settings.budgetGroupBy;

      // Build the layer entities (component or person), per-week values, and color.
      let layerEntities;
      if (groupBy === 'person') {
        const ppl = peopleForCC(cc);
        layerEntities = ppl.map((person, idx) => {
          const series = weeks.map((w) => {
            const ent = _ccPersonBreakdownForWeek(cc, w.isoStart)[person.id];
            if (!ent) return 0;
            return isCostMode ? ent.cost : ent.hours;
          });
          return { kind: 'person', id: person.id, name: person.name, series, color: personColorByIndex(idx) };
        });
      } else {
        const comps = componentsForCC(cc).map(({ component }) => component);
        layerEntities = comps.map((component) => {
          const series = weeks.map((w) => {
            const ent = _ccComponentBreakdownForWeek(cc, w.isoStart)[component.id];
            if (!ent) return 0;
            return isCostMode ? ent.cost : ent.hours;
          });
          return { kind: 'component', id: component.id, name: component.name, series, color: componentColor(component.color) };
        });
      }

      // Stack the series — each layer's TOP is its top boundary, building up
      let cumPrev = new Array(weeks.length).fill(0);
      const layers = layerEntities.map((ent) => {
        const top = ent.series.map((v, i) => cumPrev[i] + v);
        const fwd = top.map((v, i) => `${i === 0 ? 'M' : 'L'} ${xMid(i).toFixed(1)} ${yFor(v).toFixed(1)}`).join(' ');
        const back = cumPrev.slice().reverse().map((v, i) => `L ${xMid(weeks.length - 1 - i).toFixed(1)} ${yFor(v).toFixed(1)}`).join(' ');
        const path = fwd + ' ' + back + ' Z';
        const layer = { path, ent };
        cumPrev = top;
        return layer;
      });
      const areas = layers.map((l) => {
        const c = l.ent.color;
        return `<path class="b-comp-area" data-${l.ent.kind}-id="${l.ent.id}"
          d="${l.path}"
          fill="rgba(${c.rgb},.55)"
          stroke="rgba(${c.rgb},.95)"
          stroke-width="0.7"><title>${escapeHTML(l.ent.name)}</title></path>`;
      }).join('');

      // Populate the chip legend with entities matching the current group-by mode
      const legendEl = document.getElementById(`legend-${safe}`);
      if (legendEl) {
        legendEl.innerHTML = layerEntities.map((ent) => {
          const c = ent.color;
          return `<span class="component-chip" style="background:rgba(${c.rgb},.2);color:rgb(${c.rgb})" title="${escapeHTML(ent.kind)}">${escapeHTML(ent.name)}</span>`;
        }).join('');
      }

      const linePath = evm.budgets.map((v, i) =>
        `${i === 0 ? 'M' : 'L'} ${xMid(i).toFixed(1)} ${yFor(v).toFixed(1)}`).join(' ');

      const points = evm.budgets.map((v, i) =>
        `<circle class="b-point" data-cc="${escapeHTML(cc)}" data-idx="${i}" data-iso="${weeks[i].isoStart}" cx="${xMid(i).toFixed(1)}" cy="${yFor(v).toFixed(1)}" r="4">
          <title>Week of ${weeks[i].isoStart} — drag to set budget (${fmtV(v)})</title>
        </circle>`).join('');

      const todayLine = todayIdx >= 0
        ? `<line class="b-today" x1="${xLeft(todayIdx).toFixed(1)}" x2="${xLeft(todayIdx).toFixed(1)}" y1="${padT}" y2="${padT + innerH}" />
           <text class="b-today-lbl" x="${(xLeft(todayIdx) + 4).toFixed(1)}" y="${padT + 10}">today</text>`
        : '';

      const months = weeks.map((w, i) => {
        if (w.start.getDate() > 7) return '';
        return `<text class="b-tick" x="${xMid(i).toFixed(1)}" y="${H - 12}" text-anchor="middle">${w.start.toLocaleDateString(undefined, { month: 'short', year: '2-digit' })}</text>`;
      }).join('');

      const yTicks = [0, maxY / 2, maxY].map((v) =>
        `<g class="b-ytick">
          <line x1="${padL - 3}" x2="${padL}" y1="${yFor(v).toFixed(1)}" y2="${yFor(v).toFixed(1)}" />
          <text x="${padL - 6}" y="${(yFor(v) + 3).toFixed(1)}" text-anchor="end">${fmtV(v)}</text>
        </g>`).join('');

      svg.innerHTML = `
        ${yTicks}
        ${areas}
        <path class="b-budget-line" d="${linePath}" />
        ${todayLine}
        <line class="b-hover-line" x1="0" x2="0" y1="${padT}" y2="${padT + innerH}" style="display:none" />
        ${points}
        ${months}`;

      // Hover tracking — vertical crosshair + side panel listing components & budget
      let hoverPanel = svg.parentElement.querySelector('.b-hover-panel');
      if (!hoverPanel) {
        hoverPanel = document.createElement('div');
        hoverPanel.className = 'b-hover-panel';
        svg.parentElement.appendChild(hoverPanel);
      }
      const hoverLine = svg.querySelector('.b-hover-line');
      svg.addEventListener('mousemove', (e) => {
        const r = svg.getBoundingClientRect();
        const localX = (e.clientX - r.left) * (W / r.width);
        if (localX < padL || localX > padL + innerW) {
          hoverLine.style.display = 'none';
          hoverPanel.style.display = 'none';
          return;
        }
        let bestI = 0, bestD = Infinity;
        for (let i = 0; i < weeks.length; i++) {
          const d = Math.abs(xMid(i) - localX);
          if (d < bestD) { bestD = d; bestI = i; }
        }
        const lineX = xMid(bestI);
        hoverLine.setAttribute('x1', lineX.toFixed(1));
        hoverLine.setAttribute('x2', lineX.toFixed(1));
        hoverLine.style.display = '';
        const w = weeks[bestI];
        const budget = evm.budgets[bestI];
        const totalActual = isCostMode
          ? actualCostForCCWeek(cc, w.isoStart)
          : actualHoursForCCWeek(cc, w.isoStart);
        const compRows = layerEntities.map((ent) => {
          const v = ent.series[bestI];
          if (v === 0) return '';
          return `<div class="b-h-row"><span class="b-h-dot" style="background:rgba(${ent.color.rgb},.95)"></span><b>${escapeHTML(ent.name)}</b><span class="b-h-val">${fmtV(v)}</span></div>`;
        }).filter(Boolean).join('');
        const overBudget = totalActual > budget && budget > 0;
        hoverPanel.innerHTML = `
          <div class="b-h-head">Week of ${w.isoStart}</div>
          ${compRows || '<div class="b-h-empty">No activity</div>'}
          <div class="b-h-row b-h-tot"><b>Total actual</b><span class="b-h-val">${fmtV(totalActual)}</span></div>
          <div class="b-h-row"><span class="b-h-dot b-h-dot-budget"></span><b>Budget</b><span class="b-h-val ${overBudget ? 'bad' : 'ok'}">${fmtV(budget)}</span></div>
          ${budget > 0 ? `<div class="b-h-delta ${overBudget ? 'bad' : 'ok'}">${overBudget ? 'Over' : 'Under'} by ${fmtV(Math.abs(totalActual - budget))}</div>` : ''}`;
        hoverPanel.style.display = 'block';
        const pr = hoverPanel.getBoundingClientRect();
        let px = e.clientX + 14;
        let py = e.clientY + 14;
        if (px + pr.width  > innerWidth  - 8) px = e.clientX - pr.width - 14;
        if (py + pr.height > innerHeight - 8) py = e.clientY - pr.height - 14;
        hoverPanel.style.left = Math.max(8, px) + 'px';
        hoverPanel.style.top  = Math.max(8, py) + 'px';
      });
      svg.addEventListener('mouseleave', () => {
        hoverLine.style.display = 'none';
        hoverPanel.style.display = 'none';
      });

      svg.querySelectorAll('.b-point').forEach((pt) => {
        pt.addEventListener('mousedown', (e) => {
          e.preventDefault();
          e.stopPropagation();
          const idx = parseInt(pt.dataset.idx, 10);
          const iso = pt.dataset.iso;
          const liveBudgets = evm.budgets.slice();
          state.budgets[cc] = state.budgets[cc] || {};
          const isCost = mode === 'cost';
          const avg = avgHourlyRate();
          const step = isCost ? 100 : 1;
          let displayed = liveBudgets[idx] || 0;

          // Floating editable input — shown near the point during drag, focusable on release
          let editor = document.getElementById('budgetEditor');
          if (!editor) {
            editor = document.createElement('input');
            editor.id = 'budgetEditor';
            editor.type = 'text';
            editor.className = 'budget-edit';
            document.body.appendChild(editor);
          }
          editor._cc = cc; editor._iso = iso; editor._idx = idx;
          editor._isCost = isCost; editor._avg = avg; editor._step = step;
          editor._chartCtx = { cc, weeks, redraw: () => drawChart(cc, weeks) };

          function applyDisplayedValue(v) {
            displayed = Math.max(0, Math.round(v / step) * step);
            const hours = isCost ? displayed / avg : displayed;
            state.budgets[cc][iso] = Math.round(hours * 10) / 10;
            liveBudgets[idx] = displayed;
            pt.setAttribute('cy', yFor(displayed).toFixed(1));
            const newPath = liveBudgets.map((vv, ii) =>
              `${ii === 0 ? 'M' : 'L'} ${xMid(ii).toFixed(1)} ${yFor(vv).toFixed(1)}`).join(' ');
            svg.querySelector('.b-budget-line').setAttribute('d', newPath);
            editor.value = isCost ? `${displayed} €` : `${displayed} h`;
          }
          function placeEditor(clientX, clientY) {
            editor.style.display = 'block';
            editor.style.left = (clientX + 14) + 'px';
            editor.style.top  = (clientY - 28) + 'px';
          }
          function onMove(em) {
            const r = svg.getBoundingClientRect();
            const localY = (em.clientY - r.top) * (H / r.height);
            let v = ((padT + innerH) - localY) / innerH * maxY;
            // Clamp to the chart's visible Y range so the dot and the line never
            // disappear above the chart top during drag. Higher values can still
            // be set via the floating editor input — drawChart on blur rescales.
            v = Math.max(0, Math.min(maxY, v));
            applyDisplayedValue(v);
            placeEditor(em.clientX, em.clientY);
          }
          function commitDrag() {
            saveState();
            drawChart(cc, weeks);
            toast(`${cc}: ${fmtV(displayed)} for ${iso}`);
          }
          function onUp(em) {
            window.removeEventListener('mousemove', onMove);
            window.removeEventListener('mouseup', onUp);
            // Don't redraw chart yet — let the user fine-tune via the input
            saveState();
            placeEditor(em.clientX, em.clientY);
            editor.focus();
            editor.select();
          }

          // Bind editor handlers fresh each drag (closure-scoped state)
          editor.oninput = () => {
            const v = parseFloat(editor.value.replace(/[^0-9.\-]/g, '')) || 0;
            applyDisplayedValue(v);
          };
          editor.onkeydown = (ev) => {
            if (ev.key === 'Enter') { ev.preventDefault(); editor.blur(); }
            if (ev.key === 'Escape') { ev.preventDefault(); editor.blur(); }
          };
          editor.onblur = () => {
            editor.style.display = 'none';
            saveState();
            drawChart(cc, weeks);
          };

          // Initial position + value
          applyDisplayedValue(displayed);
          placeEditor(e.clientX, e.clientY);

          window.addEventListener('mousemove', onMove);
          window.addEventListener('mouseup', onUp);
        });
      });
    }

    drawList();
  }

  /* -------------------------- Portfolio / People --------------------- */

  function renderPortfolio(root) {
    const view = document.createElement('div');
    view.className = 'view';
    view.innerHTML = `
      <div class="page-head">
        <div><div class="page-title">Portfolio</div><div class="page-sub">${state.projects.length} projects</div></div>
        <div class="page-actions"><button class="ghost" id="btnNewProj2">+ Project</button></div>
      </div>
      <div class="dashboard" id="portfolioGrid">
        ${state.projects.map((p) => {
          const acts = p.actions || [];
          const total = acts.length;
          const done = acts.filter((a) => a.status === 'done').length;
          const late = acts.filter((a) => a.status !== 'done' && a.due && dayDiff(a.due, todayISO()) < 0).length;
          const pct = total ? Math.round((done / total) * 100) : 0;
          return `
            <div class="kpi clickable" data-pid="${p.id}" style="grid-column: span 2; cursor:pointer;">
              ${ROW_GRIP_HTML}
              <div class="kpi-label">${escapeHTML(p.name)}</div>
              <div class="kpi-value">${pct}%</div>
              <div class="kpi-sub">${total} actions • ${late} late • ${done} done</div>
            </div>`;
        }).join('')}
      </div>`;
    root.appendChild(view);
    wireListReorder(view.querySelector('#portfolioGrid'), {
      rowSelector: '.kpi.clickable[data-pid]',
      idAttr: 'pid',
      getArray: () => state.projects,
      setOrder: (ids) => { state.projects = ids.map((id) => state.projects.find((x) => x.id === id)).filter(Boolean); },
      commitName: 'portfolio-reorder',
    });
    $$('.kpi.clickable', view).forEach((el) => {
      el.addEventListener('click', () => {
        state.currentProjectId = el.dataset.pid;
        state.currentView = 'board';
        saveState(); render();
      });
      el.addEventListener('contextmenu', (e) => {
        e.preventDefault();
        const pid = el.dataset.pid;
        const p = state.projects.find((x) => x.id === pid);
        if (!p) return;
        showContextMenu(e.clientX, e.clientY, [
          { icon: '✎', label: 'Edit project…', onClick: () => openProjectEditor(pid) },
          { icon: '⌕', label: 'Open project',  onClick: () => {
            state.currentProjectId = pid;
            state.currentView = 'board';
            saveState(); render();
          }},
          { icon: '⊕', label: 'Save as template…', onClick: () => saveProjectAsTemplate(pid) },
          { divider: true },
          { icon: '×', label: 'Delete project', danger: true, onClick: () => {
            if (state.projects.length <= 1) { toast('Cannot delete the only project'); return; }
            const counts = {
              actions: (p.actions || []).length,
              deliverables: (p.deliverables || []).length,
              milestones: (p.milestones || []).length,
              risks: (p.risks || []).length,
            };
            const summary = Object.entries(counts).filter(([, n]) => n > 0).map(([k, n]) => `${n} ${k}`).join(', ');
            const msg = `Delete project "${p.name}"?` + (summary ? `\n\nThis will permanently remove ${summary}.\n\nThis cannot be undone (except via Undo).` : '');
            if (!confirm(msg)) return;
            const typed = prompt(`Type DELETE to confirm permanently removing "${p.name}":`);
            if ((typed || '').trim() !== 'DELETE') { toast('Cancelled — nothing was changed'); return; }
            state.projects = state.projects.filter((x) => x.id !== pid);
            if (state.currentProjectId === pid) state.currentProjectId = state.projects[0]?.id || null;
            commit('project-delete');
            toast('Deleted');
          }},
        ]);
      });
    });
    $('#btnNewProj2').addEventListener('click', () => openQuickAdd('project'));
  }

  /* --------------------- Phase J: project templates -------------------- */
  // Build a structural skeleton from a project — strip ids, action data,
  // decisions, notes, history, comments. Keep components, deliverable
  // stubs, milestone stubs, risk skeletons (title + inherent matrix).
  function projectToTemplate(p, name) {
    const t = {
      id: uid('tpl'),
      name: (name || p.name + ' (template)'),
      createdAt: todayISO(),
      sourceProjectId: p.id,
      shape: {
        description: p.description || '',
        components: (p.components || []).map((c) => ({
          name: c.name, color: c.color, costCenter: c.costCenter || null,
        })),
        deliverables: (p.deliverables || []).map((d) => ({
          name: d.name, status: 'todo',
          // intentionally drop dates — instantiation should re-plan
        })),
        milestones: (p.milestones || []).map((m) => ({
          name: m.name, status: 'planned',
        })),
        risks: (p.risks || []).map((r) => ({
          title: r.title,
          kind: r.kind || 'risk',
          inherent: r.inherent ? { ...r.inherent } : { probability: 0, impact: 0 },
          mitigation: r.mitigation || '',
        })),
        tags: (p.tags || []).map((tg) => ({ name: tg.name, rgb: tg.rgb })),
        // Intentionally NOT cloned: actions, decisions, change requests,
        // notes, links, meetings, history. The shape is the skeleton, not
        // the in-flight work.
      },
    };
    return t;
  }
  function applyTemplateToProject(np, tpl) {
    const shape = tpl.shape || {};
    if (shape.description && !np.description) np.description = shape.description;
    np.components   = (shape.components   || []).map((c) => ({ id: uid('cm'), ...c }));
    np.deliverables = (shape.deliverables || []).map((d) => ({ id: uid('d'), ...d }));
    np.milestones   = (shape.milestones   || []).map((m) => ({ id: uid('m'), ...m }));
    np.risks        = (shape.risks        || []).map((r) => ({
      id: uid('rk'), ...r,
      residual: r.inherent ? { ...r.inherent } : { probability: 0, impact: 0 },
      owner: null,
      actionId: null,
    }));
    np.tags         = (shape.tags         || []).map((t) => ({ id: uid('tg'), ...t }));
  }
  function saveProjectAsTemplate(projectId) {
    const p = state.projects.find((x) => x.id === projectId);
    if (!p) return;
    const suggested = p.name + ' template';
    const name = prompt('Template name:', suggested);
    if (!name) return;
    const tpl = projectToTemplate(p, name.trim());
    state.templates = state.templates || [];
    state.templates.push(tpl);
    commit('template-create');
    toast(`Saved as template: ${tpl.name}`);
  }

  function openProjectEditor(projectId) {
    const p = state.projects.find((x) => x.id === projectId);
    if (!p) return;
    const overlay = document.createElement('div');
    overlay.className = 'overlay desc-overlay';
    overlay.innerHTML = `
      <div class="desc-modal" style="width:520px;">
        <div class="desc-head">
          <div class="desc-title">Edit project</div>
          <button class="icon-btn" id="prClose" title="Close">×</button>
        </div>
        <div style="padding:14px 16px; display:flex; flex-direction:column; gap:10px;">
          <div class="field"><label>Name</label><input id="prName" value="${escapeHTML(p.name)}" /></div>
          <div class="field"><label>Description</label><textarea id="prDesc" style="min-height:100px;">${escapeHTML(p.description || '')}</textarea></div>
        </div>
        <div class="desc-foot">
          <button class="ghost" id="prCancel">Cancel</button>
          <button class="primary" id="prSave">Save</button>
        </div>
      </div>`;
    document.body.appendChild(overlay);
    setTimeout(() => document.getElementById('prName').focus(), 30);
    const close = () => overlay.remove();
    overlay.querySelector('#prClose').addEventListener('click', close);
    overlay.querySelector('#prCancel').addEventListener('click', close);
    // Modal closes ONLY via the explicit × / Cancel — no backdrop or Escape.
    overlay.querySelector('#prSave').addEventListener('click', () => {
      const name = document.getElementById('prName').value.trim();
      if (!name) return toast('Name required');
      p.name = name;
      p.description = document.getElementById('prDesc').value;
      commit('project-edit');
      close();
      toast('Saved');
    });
  }

  // Compute weekly workload for a person across the next `weeks` weeks.
  // Returns array of { weekStart: Date, count: number, items: action[] }.
  function weeklyLoad(personId, weeks = 12) {
    const today = new Date(); today.setHours(0, 0, 0, 0);
    let weekStart = new Date(today);
    while (weekStart.getDay() !== 1) weekStart = new Date(weekStart.getTime() - dayMs);
    const out = [];
    for (let w = 0; w < weeks; w++) {
      const wStart = new Date(weekStart.getTime() + w * 7 * dayMs);
      const wEnd = new Date(wStart.getTime() + 6 * dayMs);
      const items = [];
      let commitmentSum = 0;
      state.projects.forEach((proj) => {
        (proj.actions || []).forEach((a) => {
          if (a.deletedAt) return;
          // Closed actions (done OR cancelled) don't consume capacity
          if (a.owner !== personId || isClosedStatus(a.status)) return;
          if (!a.due) return;
          const due = parseDate(a.due);
          const start = a.startDate ? parseDate(a.startDate) :
            new Date(due.getTime() - 2 * dayMs);
          if (start <= wEnd && due >= wStart) {
            items.push({ a, proj });
            commitmentSum += (typeof a.commitment === 'number') ? a.commitment : 100;
          }
        });
      });
      // `count` is now in % of FTE (commitment sum), not action headcount.
      out.push({ weekStart: wStart, count: commitmentSum, items });
    }
    return out;
  }

  function workloadSparkSVG(person, series) {
    const cap = person.capacity || 5;
    const W = 480, H = 70;
    const padL = 0, padR = 0, padT = 6, padB = 14;
    const innerW = W - padL - padR;
    const innerH = H - padT - padB;
    const n = series.length;
    const barW = innerW / n;
    const maxVal = Math.max(cap * 1.5, ...series.map((s) => s.count), 4);
    // Capacity baseline y
    const yFor = (v) => padT + innerH - (v / maxVal) * innerH;
    const capY = yFor(cap);

    const bars = series.map((s, i) => {
      const x = padL + i * barW;
      const y = yFor(s.count);
      const h = padT + innerH - y;
      const cls = s.count > cap ? 'over' : s.count > cap * 0.8 ? 'warn' : 'ok';
      const label = s.weekStart.toLocaleDateString(undefined, { month: 'short', day: 'numeric' });
      const hours = Math.round((s.count / 100) * HOURS_PER_WEEK);
      return `<rect class="spark-bar ${cls}" x="${x + 2}" y="${y}" width="${Math.max(2, barW - 4)}" height="${Math.max(0, h)}" rx="2">
        <title>Week of ${label} — ${s.count}% FTE (${hours} h) · cap ${cap}%</title>
      </rect>`;
    }).join('');

    const ticks = series.map((s, i) => {
      // Show monthly tick + label at month start
      const isMonthStart = s.weekStart.getDate() <= 7;
      if (!isMonthStart) return '';
      const x = padL + i * barW + barW / 2;
      const lbl = s.weekStart.toLocaleDateString(undefined, { month: 'short' });
      return `<text class="spark-tick" x="${x}" y="${H - 2}" text-anchor="middle">${lbl}</text>`;
    }).join('');

    return `
      <svg class="spark" viewBox="0 0 ${W} ${H}" preserveAspectRatio="none" role="img" aria-label="Workload over time">
        <line class="spark-cap" x1="${padL}" x2="${W - padR}" y1="${capY}" y2="${capY}" stroke-dasharray="4 3" />
        <text class="spark-cap-label" x="${W - padR - 4}" y="${Math.max(10, capY - 3)}" text-anchor="end">cap ${cap}%</text>
        ${bars}
        ${ticks}
      </svg>`;
  }

  function renderPeople(root) {
    const view = document.createElement('div');
    view.className = 'view';
    const wl = state.people.map((p) => {
      const open = state.projects.flatMap((pr) => pr.actions || []).filter((a) => a.owner === p.id && a.status !== 'done').length;
      const series = weeklyLoad(p.id, 12);
      const peakWeek = series.reduce((mx, s) => s.count > mx.count ? s : mx, series[0] || { count: 0 });
      return { p, open, series, peakWeek };
    });
    view.innerHTML = `
      <div class="page-head">
        <div><div class="page-title">People</div><div class="page-sub">${state.people.length} members • capacity in % of FTE (1 FTE = 8h/day × 5 days/week, 212 working days/year) • workload across all projects, projected over the next 12 weeks</div></div>
        <div class="page-actions"><button class="ghost" id="btnNewPerson">+ Person</button></div>
      </div>
      <div class="panel">
        <div class="panel-title">
          Workload <span class="panel-sub">— bars are % of FTE (1 FTE ≈ ${HOURS_PER_WEEK} h/week)</span>
          <span class="legend">
            <span class="legend-item"><span class="dot ok"></span>≤80% cap</span>
            <span class="legend-item"><span class="dot warn"></span>≤100% cap</span>
            <span class="legend-item"><span class="dot bad"></span>over cap</span>
          </span>
        </div>
        <div id="peopleWl">
          ${wl.map(({ p, open, series, peakWeek }) => {
            const cap = p.capacity || 100;
            // open is action count (headcount); compute current commitment % across active actions
            const openCmt = state.projects.flatMap((pr) => pr.actions || [])
              .filter((a) => a.owner === p.id && a.status !== 'done' && !a.deletedAt)
              .reduce((s, a) => s + ((typeof a.commitment === 'number') ? a.commitment : 100), 0);
            const pct = clamp(Math.round((openCmt / cap) * 100), 0, 200);
            const cls = pct > 100 ? 'over' : pct > 80 ? 'warn' : 'ok';
            const peakCls = peakWeek.count > cap ? 'over' : peakWeek.count > cap * 0.8 ? 'warn' : 'ok';
            return `
              <div class="person-row clickable" data-owner-id="${p.id}" title="Click to filter Register to ${escapeHTML(p.name)}">
                ${ROW_GRIP_HTML}
                <div class="name-cell">
                  <span class="avatar">${initials(p.name)}</span>
                  <span class="who">
                    <b>${escapeHTML(p.name)}</b>
                    <span>${escapeHTML(p.role || '')}</span>
                  </span>
                </div>
                <div class="now-load">
                  <div class="bar-track"><div class="bar-fill ${cls}" style="width:${Math.min(100, pct)}%"></div></div>
                  <div class="bar-val">${openCmt}% / ${cap}%</div>
                </div>
                <div class="spark-wrap">
                  ${workloadSparkSVG(p, series)}
                  <div class="spark-meta"><span class="${peakCls}">peak ${peakWeek.count}% (${Math.round((peakWeek.count / 100) * HOURS_PER_WEEK)} h)</span></div>
                </div>
              </div>`;
          }).join('')}
        </div>
      </div>`;
    root.appendChild(view);
    $('#btnNewPerson').addEventListener('click', () => openQuickAdd('person'));
    wireListReorder(view.querySelector('#peopleWl'), {
      rowSelector: '.person-row[data-owner-id]',
      idAttr: 'ownerId',
      getArray: () => state.people,
      setOrder: (ids) => { state.people = ids.map((id) => state.people.find((x) => x.id === id)).filter(Boolean); },
      commitName: 'people-reorder',
    });
    $$('.person-row.clickable', view).forEach((row) => {
      row.addEventListener('click', () => {
        openPersonDashboard(row.dataset.ownerId);
      });
      row.addEventListener('contextmenu', (e) => {
        e.preventDefault();
        const id = row.dataset.ownerId;
        const p = state.people.find((x) => x.id === id);
        showContextMenu(e.clientX, e.clientY, [
          { icon: '◰', label: 'Open dashboard',  onClick: () => openPersonDashboard(id) },
          { icon: '✎', label: 'Edit person…',    onClick: () => openPersonEditor(id) },
          { icon: '⌕', label: `Filter Register to ${p?.name || ''}`, onClick: () => applyTopbarFilter({ owner: id, view: 'register' }) },
          { divider: true },
          { icon: '×', label: 'Delete person…', danger: true, onClick: () => {
            const open = state.projects.flatMap((pr) => pr.actions || []).filter((a) => a.owner === id && a.status !== 'done').length;
            if (!confirm(`Delete ${p?.name}?` + (open ? ` (${open} open action${open === 1 ? '' : 's'} will be unassigned).` : ''))) return;
            state.projects.forEach((pr) => (pr.actions || []).forEach((a) => { if (a.owner === id) a.owner = null; }));
            state.people = state.people.filter((x) => x.id !== id);
            commit('person-delete');
            toast('Deleted');
          }},
        ]);
      });
    });
  }

  /* ---------------------- Phase I: Person dashboard -------------------- */
  // Opens the existing drawer with an aggregated read-only view of one
  // person: open / late actions, weekly load + spark, originated CRs,
  // decisions made, owned risks with non-trivial residual. Each list item
  // routes to its underlying drawer / editor.
  function openPersonDashboard(personId) {
    const p = state.people.find((x) => x.id === personId);
    if (!p) return;
    $('#drawerTitle').textContent = 'Person dashboard';
    const today = todayISO();
    const allActs = state.projects.flatMap((pr) =>
      (pr.actions || []).filter((a) => !a.deletedAt && a.owner === p.id).map((a) => ({ a, pr })));
    const open  = allActs.filter(({ a }) => !isClosedStatus(a.status));
    const late  = open.filter(({ a }) => a.due && dayDiff(a.due, today) < 0);
    const blocked = open.filter(({ a }) => a.status === 'blocked');
    const allCRs       = state.projects.flatMap((pr) => (pr.changes || []).map((c) => ({ c, pr })));
    const myCRs        = allCRs.filter(({ c }) => c.originator === p.id);
    const myDecisions  = state.projects.flatMap((pr) => (pr.decisions || []).filter((d) => d.owner === p.id).map((d) => ({ d, pr })));
    const myRisks      = state.projects.flatMap((pr) =>
      (pr.risks || []).filter((r) => r.owner === p.id).map((r) => {
        const res = r.residual || r.inherent || { probability: 0, impact: 0 };
        return { r, pr, _score: (res.probability || 0) * (res.impact || 0) };
      }))
      .filter((x) => x._score >= 6)
      .sort((a, b) => b._score - a._score);

    const series = weeklyLoad(p.id, 12);
    const cap = p.capacity || 100;
    const peakWeek = series.reduce((mx, s) => s.count > mx.count ? s : mx, series[0] || { count: 0 });
    const openCmt = open.reduce((s, { a }) => s + ((typeof a.commitment === 'number') ? a.commitment : 100), 0);
    const pct = clamp(Math.round((openCmt / cap) * 100), 0, 200);
    const cls = pct > 100 ? 'over' : pct > 80 ? 'warn' : 'ok';

    const listRow = (title, sub, run) => `
      <div class="dash-row${run ? ' clickable' : ''}" data-run="1">
        <div class="dash-row-title">${escapeHTML(title)}</div>
        <div class="dash-row-sub">${escapeHTML(sub)}</div>
      </div>`;
    const listOrEmpty = (title, count, html) => `
      <div class="dash-section">
        <div class="dash-section-title">${escapeHTML(title)}<span class="dash-section-count">${count}</span></div>
        ${count ? html : '<div class="empty">— none</div>'}
      </div>`;

    $('#drawerBody').innerHTML = `
      <div class="dash-head">
        <span class="avatar lg">${initials(p.name)}</span>
        <div class="dash-head-body">
          <div class="dash-name">${escapeHTML(p.name)}</div>
          <div class="dash-role">${escapeHTML(p.role || '—')}</div>
          <div class="dash-head-actions">
            <button class="ghost" id="dashEdit">Edit person…</button>
            <button class="ghost" id="dashFilter">Filter Register to ${escapeHTML(p.name)}</button>
          </div>
        </div>
      </div>
      <div class="dash-kpis">
        <div class="dash-kpi"><div class="dash-kpi-num">${open.length}</div><div class="dash-kpi-lbl">Open</div></div>
        <div class="dash-kpi ${late.length ? 'bad' : ''}"><div class="dash-kpi-num">${late.length}</div><div class="dash-kpi-lbl">Late</div></div>
        <div class="dash-kpi ${blocked.length ? 'warn' : ''}"><div class="dash-kpi-num">${blocked.length}</div><div class="dash-kpi-lbl">Blocked</div></div>
        <div class="dash-kpi ${cls === 'over' ? 'bad' : cls === 'warn' ? 'warn' : ''}"><div class="dash-kpi-num">${pct}%</div><div class="dash-kpi-lbl">Load</div></div>
      </div>
      <div class="dash-section">
        <div class="dash-section-title">Workload — next 12 weeks<span class="dash-section-count">peak ${peakWeek.count}%</span></div>
        ${workloadSparkSVG(p, series)}
      </div>
      ${listOrEmpty('Late actions', late.length,
        late.slice(0, 12).map(({ a, pr }) => `
          <div class="dash-row clickable" data-action-id="${a.id}">
            <div class="dash-row-title">${escapeHTML(a.title)}</div>
            <div class="dash-row-sub">${escapeHTML(pr.name)} · ${Math.abs(dayDiff(a.due, today))}d late</div>
          </div>`).join(''))}
      ${listOrEmpty('Open actions', open.length,
        open.slice(0, 20).map(({ a, pr }) => `
          <div class="dash-row clickable" data-action-id="${a.id}">
            <div class="dash-row-title">${escapeHTML(a.title)}</div>
            <div class="dash-row-sub">${escapeHTML(pr.name)} · ${escapeHTML(a.status)}${a.due ? ' · due ' + a.due : ''}</div>
          </div>`).join(''))}
      ${listOrEmpty('Originated change requests', myCRs.length,
        myCRs.slice(0, 12).map(({ c, pr }) => `
          <div class="dash-row clickable" data-cr-id="${c.id}" data-proj-id="${pr.id}">
            <div class="dash-row-title">${escapeHTML(c.title)}</div>
            <div class="dash-row-sub">${escapeHTML(pr.name)} · ${escapeHTML(c.status)}${c.decisionDate ? ' · ' + c.decisionDate : ''}</div>
          </div>`).join(''))}
      ${listOrEmpty('Decisions made', myDecisions.length,
        myDecisions.slice(0, 12).map(({ d, pr }) => `
          <div class="dash-row" data-dec-id="${d.id}">
            <div class="dash-row-title">${escapeHTML(d.title)}</div>
            <div class="dash-row-sub">${escapeHTML(pr.name)}${d.date ? ' · ' + d.date : ''}</div>
          </div>`).join(''))}
      ${listOrEmpty('Owned risks (residual ≥ 6)', myRisks.length,
        myRisks.slice(0, 12).map(({ r, pr, _score }) => `
          <div class="dash-row clickable" data-risk-id="${r.id}" data-proj-id="${pr.id}">
            <div class="dash-row-title">${escapeHTML(r.title)}</div>
            <div class="dash-row-sub">${escapeHTML(pr.name)} · residual ${_score}</div>
          </div>`).join(''))}`;
    $('#drawer').hidden = false;
    $('#dashEdit').addEventListener('click', () => openPersonEditor(personId));
    $('#dashFilter').addEventListener('click', () => {
      closeDrawer();
      applyTopbarFilter({ owner: personId, view: 'register' });
    });
    $('#drawerBody').querySelectorAll('[data-action-id]').forEach((row) => {
      row.addEventListener('click', () => openDrawer(row.dataset.actionId));
    });
    $('#drawerBody').querySelectorAll('[data-cr-id]').forEach((row) => {
      row.addEventListener('click', () => {
        state.currentProjectId = row.dataset.projId;
        openChangeRequestEditor(row.dataset.crId);
      });
    });
    $('#drawerBody').querySelectorAll('[data-risk-id]').forEach((row) => {
      row.addEventListener('click', () => {
        state.currentProjectId = row.dataset.projId;
        openRiskEditor(row.dataset.riskId);
      });
    });
  }

  /* -------------------------- Detail drawer -------------------------- */

  function openDrawer(actionId) {
    const proj = curProject();
    const a = proj.actions.find((x) => x.id === actionId);
    if (!a) return;
    $('#drawerTitle').textContent = 'Action details';
    const body = $('#drawerBody');
    body.innerHTML = `
      <div class="field"><label>Title</label><input id="dTitle" value="${escapeHTML(a.title)}" /></div>
      <div style="display:grid; grid-template-columns:1fr 1fr 1fr; gap:10px;">
        <div class="field"><label>Owner</label>
          <select id="dOwner">${state.people.map((p) => `<option value="${p.id}" ${p.id === a.owner ? 'selected' : ''}>${escapeHTML(p.name)}</option>`).join('')}</select>
        </div>
        <div class="field"><label>Originator</label>
          <select id="dOriginator"><option value="">— same as owner</option>${state.people.map((p) => `<option value="${p.id}" ${p.id === a.originator ? 'selected' : ''}>${escapeHTML(p.name)}</option>`).join('')}</select>
        </div>
        <div class="field"><label>Originator date</label>
          <input id="dOriginatorDate" type="date" value="${a.originatorDate || a.createdAt || ''}" title="When this action was originated" />
        </div>
      </div>
      <div style="display:grid; grid-template-columns:1fr 1fr 1fr; gap:10px;">
        <div class="field"><label>Status</label>
          <select id="dStatus">${STATUSES.map((s) => `<option value="${s.id}" ${s.id === a.status ? 'selected' : ''}>${s.name}</option>`).join('')}</select>
        </div>
        <div class="field"><label>Priority</label>
          <select id="dPriorityLevel">${PRIORITY_LEVELS.map((p) => `<option value="${p.id}" ${p.id === (a.priorityLevel || 'med') ? 'selected' : ''}>${p.label}</option>`).join('')}</select>
        </div>
        <div class="field"><label>Due</label><input id="dDue" type="date" value="${a.due || ''}" /></div>
      </div>
      <div class="field">
        <label>Commitment <span class="muted" id="dCmtVal">${typeof a.commitment === 'number' ? a.commitment : 100}%</span></label>
        <input id="dCmt" type="range" min="5" max="100" step="5" value="${typeof a.commitment === 'number' ? a.commitment : 100}" />
      </div>
      <div class="field"><label>Component</label>
        <select id="dComponent"><option value="">—</option>${(proj.components || []).map((pt) => `<option value="${pt.id}" ${pt.id === a.component ? 'selected' : ''}>${escapeHTML(pt.name)}</option>`).join('')}</select>
      </div>
      <div style="display:grid; grid-template-columns:1fr 1fr; gap:10px;">
        <div class="field"><label>Deliverable</label>
          <select id="dDel"><option value="">—</option>${(proj.deliverables || []).map((d) => `<option value="${d.id}" ${d.id === a.deliverable ? 'selected' : ''}>${escapeHTML(d.name)}</option>`).join('')}</select>
        </div>
        <div class="field"><label>Milestone</label>
          <select id="dMile"><option value="">—</option>${(proj.milestones || []).map((m) => `<option value="${m.id}" ${m.id === a.milestone ? 'selected' : ''}>${escapeHTML(m.name)}</option>`).join('')}</select>
        </div>
      </div>
      <div class="field"><label>Depends on <span class="muted">— actions that must finish before this one</span></label>
        <div class="depends-input" id="dDependsWrap"></div>
      </div>
      <div class="field"><label>Tags</label>
        <div class="tags-input" id="dTagsWrap"></div>
      </div>
      <div class="field"><label>Notes / justification</label><textarea id="dNotes">${escapeHTML(a.notes || '')}</textarea></div>
      <div class="field"><label>Comments</label>
        <div class="comments" id="dComments"></div>
      </div>
      <div class="field"><label>History</label>
        <div class="history">${(a.history || []).slice(-10).reverse().map((h) => `<div class="history-item"><b>${h.at}</b> — ${escapeHTML(h.what)}</div>`).join('')}</div>
      </div>
      <div style="display:flex; gap:8px; margin-top:6px;">
        <button class="primary" id="dSave">Save</button>
        <button class="ghost" id="dDelete" style="margin-left:auto; color:var(--bad);">Delete</button>
      </div>`;
    $('#drawer').hidden = false;
    $('#dCmt').addEventListener('input', (e) => { $('#dCmtVal').textContent = e.target.value + '%'; });

    // Depends-on multi-select. Local working list mirrored back to a.dependsOn
    // on save. Self-references and would-be cycles are filtered from the
    // candidate dropdown so users can't create them in the first place.
    const drawerDeps = Array.isArray(a.dependsOn) ? a.dependsOn.slice() : [];
    function reverseDepGraph(actsList) {
      // For cycle prevention: collect every action that already (transitively)
      // depends on the action we're editing — those can't be added as deps.
      const reverse = new Map();
      actsList.forEach((x) => (x.dependsOn || []).forEach((depId) => {
        if (!reverse.has(depId)) reverse.set(depId, new Set());
        reverse.get(depId).add(x.id);
      }));
      const blocked = new Set();
      const stack = [a.id];
      while (stack.length) {
        const id = stack.pop();
        if (blocked.has(id)) continue;
        blocked.add(id);
        (reverse.get(id) || []).forEach((parent) => stack.push(parent));
      }
      return blocked;
    }
    function renderDepsUI() {
      const wrap = $('#dDependsWrap');
      const acts = (proj.actions || []).filter((x) => !x.deletedAt);
      const byId = new Map(acts.map((x) => [x.id, x]));
      const blocked = reverseDepGraph(acts); // includes a.id itself
      wrap.innerHTML = `
        <div class="depends-chips">
          ${drawerDeps.map((id) => {
            const dep = byId.get(id);
            if (!dep) return '';
            return `<span class="dep-chip" data-id="${escapeHTML(id)}">
              <span class="dep-chip-text">${escapeHTML(dep.title)}</span>
              <button type="button" class="dep-chip-x" data-action="remove" title="Remove">×</button>
            </span>`;
          }).join('')}
        </div>
        <div class="depends-search" data-empty="Add a dependency…">
          <input type="text" id="dDepSearch" placeholder="Type to search actions…" autocomplete="off" />
          <div class="depends-results" id="dDepResults" hidden></div>
        </div>`;
      wrap.querySelectorAll('[data-action="remove"]').forEach((btn) => {
        btn.addEventListener('click', () => {
          const id = btn.closest('.dep-chip')?.dataset.id;
          const idx = drawerDeps.indexOf(id);
          if (idx >= 0) drawerDeps.splice(idx, 1);
          renderDepsUI();
        });
      });
      const inp = $('#dDepSearch');
      const results = $('#dDepResults');
      const updateResults = () => {
        const q = inp.value.trim().toLowerCase();
        const candidates = acts.filter((x) =>
          x.id !== a.id &&
          !drawerDeps.includes(x.id) &&
          !blocked.has(x.id));
        const ranked = q
          ? candidates.map((x) => ({ x, s: fuzzyScore(q, (x.title + ' ' + personName(x.owner))) }))
              .filter((r) => r.s > 0)
              .sort((p, q2) => q2.s - p.s)
          : candidates.slice(0, 8).map((x) => ({ x, s: 0 }));
        const list = ranked.slice(0, 8);
        if (!list.length) {
          results.innerHTML = '<div class="depends-empty">No matches</div>';
        } else {
          results.innerHTML = list.map((r) => `
            <div class="depends-result" data-id="${escapeHTML(r.x.id)}">
              <span class="depends-result-title">${escapeHTML(r.x.title)}</span>
              <span class="depends-result-meta">${escapeHTML(personName(r.x.owner))}${r.x.due ? ' · ' + r.x.due : ''}</span>
            </div>`).join('');
        }
        results.hidden = false;
        results.querySelectorAll('.depends-result').forEach((row) => {
          row.addEventListener('mousedown', (e) => {
            e.preventDefault();
            const id = row.dataset.id;
            if (!drawerDeps.includes(id)) drawerDeps.push(id);
            renderDepsUI();
            setTimeout(() => $('#dDepSearch')?.focus(), 0);
          });
        });
      };
      inp.addEventListener('focus', updateResults);
      inp.addEventListener('input', updateResults);
      inp.addEventListener('blur', () => setTimeout(() => { results.hidden = true; }, 120));
    }
    renderDepsUI();

    // Phase G — tags multi-select. Tags are project-scoped; new tags are
    // appended to proj.tags. Removing a chip just removes the id from the
    // record's tags array (does not delete the project tag).
    const drawerTags = Array.isArray(a.tags) ? a.tags.slice() : [];
    function renderTagsUI() {
      const wrap = $('#dTagsWrap');
      const projTags = proj.tags || [];
      const byId = new Map(projTags.map((t) => [t.id, t]));
      const remaining = projTags.filter((t) => !drawerTags.includes(t.id));
      wrap.innerHTML = `
        <div class="tags-chips">
          ${drawerTags.map((id) => {
            const t = byId.get(id);
            if (!t) return '';
            const rgb = t.rgb || '120, 120, 140';
            return `<span class="dep-chip tag-chip-edit" data-id="${escapeHTML(id)}" style="background:rgba(${rgb},.18);color:rgb(${rgb});border:1px solid rgba(${rgb},.40)">
              <span class="dep-chip-text">${escapeHTML(t.name)}</span>
              <button type="button" class="dep-chip-x" data-action="remove-tag" title="Remove">×</button>
            </span>`;
          }).join('')}
        </div>
        <div class="depends-search">
          <input type="text" id="dTagSearch" placeholder="${remaining.length || drawerTags.length ? 'Type to search or create…' : 'Type a name to create your first tag…'}" autocomplete="off" />
          <div class="depends-results" id="dTagResults" hidden></div>
        </div>`;
      wrap.querySelectorAll('[data-action="remove-tag"]').forEach((btn) => {
        btn.addEventListener('click', () => {
          const id = btn.closest('.tag-chip-edit')?.dataset.id;
          const idx = drawerTags.indexOf(id);
          if (idx >= 0) drawerTags.splice(idx, 1);
          renderTagsUI();
        });
      });
      const inp = $('#dTagSearch');
      const results = $('#dTagResults');
      const updateResults = () => {
        const q = inp.value.trim();
        const candidates = (proj.tags || []).filter((t) => !drawerTags.includes(t.id));
        const ranked = q
          ? candidates.map((t) => ({ t, s: fuzzyScore(q, t.name) })).filter((r) => r.s > 0).sort((p, q2) => q2.s - p.s)
          : candidates.slice(0, 8).map((t) => ({ t, s: 0 }));
        const list = ranked.slice(0, 8);
        const exact = q && (proj.tags || []).find((t) => t.name.toLowerCase() === q.toLowerCase());
        const canCreate = q && !exact;
        let html = list.map((r) => {
          const rgb = r.t.rgb || '120, 120, 140';
          return `<div class="depends-result tag-result" data-id="${escapeHTML(r.t.id)}">
            <span class="tag-color-dot" style="background:rgb(${rgb})"></span>
            <span class="depends-result-title">${escapeHTML(r.t.name)}</span>
          </div>`;
        }).join('');
        if (canCreate) {
          html += `<div class="depends-result tag-create" data-create="${escapeHTML(q)}">
            <span class="tag-color-dot" style="background:rgb(${TAG_PALETTE[(proj.tags || []).length % TAG_PALETTE.length]})"></span>
            <span class="depends-result-title">+ Create "<b>${escapeHTML(q)}</b>"</span>
          </div>`;
        }
        if (!list.length && !canCreate) {
          html = '<div class="depends-empty">Type a name to create a new tag.</div>';
        }
        results.innerHTML = html;
        results.hidden = false;
        results.querySelectorAll('.tag-result').forEach((row) => {
          row.addEventListener('mousedown', (e) => {
            e.preventDefault();
            const id = row.dataset.id;
            if (!drawerTags.includes(id)) drawerTags.push(id);
            inp.value = '';
            renderTagsUI();
            setTimeout(() => $('#dTagSearch')?.focus(), 0);
          });
        });
        results.querySelector('.tag-create')?.addEventListener('mousedown', (e) => {
          e.preventDefault();
          const name = e.currentTarget.dataset.create;
          const t = { id: uid('tg'), name, rgb: TAG_PALETTE[(proj.tags || []).length % TAG_PALETTE.length] };
          proj.tags = proj.tags || [];
          proj.tags.push(t);
          drawerTags.push(t.id);
          inp.value = '';
          renderTagsUI();
          setTimeout(() => $('#dTagSearch')?.focus(), 0);
        });
      };
      inp.addEventListener('focus', updateResults);
      inp.addEventListener('input', updateResults);
      inp.addEventListener('blur', () => setTimeout(() => { results.hidden = true; }, 120));
    }
    renderTagsUI();

    // Phase G — comments thread. Local working list; mutations are persisted
    // immediately on Post (independent of Save) so a user typing a long
    // comment doesn't lose it if they hit Cancel on the rest.
    function renderCommentsUI() {
      const wrap = $('#dComments');
      const items = (a.comments || []).slice().sort((x, y) => (x.at || '').localeCompare(y.at || ''));
      wrap.innerHTML = `
        <div class="comments-list">
          ${items.length ? items.map((c) => {
            const author = personName(c.by) || (c.by ? c.by : 'someone');
            const when = c.at ? c.at.replace('T', ' ').slice(0, 16) : '';
            return `<div class="comment-row" data-comment-id="${escapeHTML(c.id)}">
              <div class="comment-head">
                <span class="comment-author">${escapeHTML(author)}</span>
                <span class="comment-when">${escapeHTML(when)}</span>
                <button class="comment-del" data-action="del-comment" title="Delete comment">×</button>
              </div>
              <div class="comment-text">${escapeHTML(c.text)}</div>
            </div>`;
          }).join('') : '<div class="comments-empty">No comments yet.</div>'}
        </div>
        <div class="comment-compose">
          <textarea id="dCommentNew" placeholder="Add a comment…" rows="2"></textarea>
          <button class="ghost" id="dCommentPost" disabled>Post</button>
        </div>`;
      wrap.querySelectorAll('[data-action="del-comment"]').forEach((btn) => {
        btn.addEventListener('click', () => {
          const id = btn.closest('[data-comment-id]')?.dataset.commentId;
          if (!id) return;
          if (!confirm('Delete this comment?')) return;
          a.comments = (a.comments || []).filter((c) => c.id !== id);
          a.updatedAt = todayISO();
          commit('comment-delete');
          // Re-render the drawer so the comment thread updates immediately
          // without re-opening, but other drawer fields keep their state.
          renderCommentsUI();
        });
      });
      const ta = $('#dCommentNew');
      const post = $('#dCommentPost');
      ta?.addEventListener('input', () => { post.disabled = !ta.value.trim(); });
      post?.addEventListener('click', () => {
        const text = ta.value.trim();
        if (!text) return;
        a.comments = a.comments || [];
        a.comments.push({
          id: uid('cm'),
          by: state.settings?.localUser || null,
          at: new Date().toISOString(),
          text,
        });
        a.updatedAt = todayISO();
        commit('comment-add');
        renderCommentsUI();
      });
    }
    renderCommentsUI();

    $('#dSave').addEventListener('click', () => {
      const oldStatus = a.status;
      const oldOwner = a.owner;
      const oldDue = a.due;
      const oldCmt = typeof a.commitment === 'number' ? a.commitment : 100;
      a.title = $('#dTitle').value.trim() || a.title;
      a.owner = $('#dOwner').value;
      a.status = $('#dStatus').value;
      a.due = $('#dDue').value || null;
      a.deliverable = $('#dDel').value || null;
      a.milestone = $('#dMile').value || null;
      a.component = $('#dComponent').value || null;
      a.originator = $('#dOriginator')?.value || null;
      const oldOrigDate = a.originatorDate;
      a.originatorDate = $('#dOriginatorDate')?.value || a.originatorDate || a.createdAt || todayISO();
      if (oldOrigDate !== a.originatorDate) a.history.push({ at: todayISO(), what: `Originator date: ${oldOrigDate || '—'} → ${a.originatorDate || '—'}` });
      const oldPriorityLevel = a.priorityLevel || 'med';
      a.priorityLevel = $('#dPriorityLevel')?.value || a.priorityLevel || 'med';
      if (oldPriorityLevel !== a.priorityLevel) {
        a.history.push({ at: todayISO(), what: `Priority: ${priorityLevel(oldPriorityLevel).label} → ${priorityLevel(a.priorityLevel).label}` });
      }
      a.commitment = clamp(parseInt($('#dCmt').value, 10) || 100, 5, 100);
      if (oldCmt !== a.commitment) a.history.push({ at: todayISO(), what: `Commitment: ${oldCmt}% → ${a.commitment}%` });
      const oldDeps = Array.isArray(a.dependsOn) ? a.dependsOn.slice() : [];
      a.dependsOn = drawerDeps.slice();
      const depsChanged = oldDeps.length !== a.dependsOn.length || oldDeps.some((d, i) => d !== a.dependsOn[i]);
      if (depsChanged) a.history.push({ at: todayISO(), what: `Dependencies: ${oldDeps.length} → ${a.dependsOn.length}` });
      const oldTags = Array.isArray(a.tags) ? a.tags.slice() : [];
      a.tags = drawerTags.slice();
      const tagsChanged = oldTags.length !== a.tags.length || oldTags.some((t, i) => t !== a.tags[i]);
      if (tagsChanged) a.history.push({ at: todayISO(), what: `Tags: ${oldTags.length} → ${a.tags.length}` });
      a.notes = $('#dNotes').value;
      a.updatedAt = todayISO();
      if (oldStatus !== a.status) a.history.push({ at: todayISO(), what: `Status: ${oldStatus} → ${a.status}` });
      if (oldOwner !== a.owner) a.history.push({ at: todayISO(), what: `Owner: ${personName(oldOwner)} → ${personName(a.owner)}` });
      if (oldDue !== a.due) a.history.push({ at: todayISO(), what: `Due: ${oldDue || '—'} → ${a.due || '—'}` });
      commit('edit');
      closeDrawer();
      toast('Saved');
    });
    $('#dDelete').addEventListener('click', () => {
      if (!confirm('Move this action to Archive? You can restore it later.')) return;
      const sourceProj = projectOfAction(a.id) || proj;
      const target = (sourceProj.actions || []).find((x) => x.id === a.id);
      if (target) {
        target.deletedAt = todayISO();
        target.history.push({ at: todayISO(), what: 'Moved to Archive' });
        target.updatedAt = todayISO();
      }
      commit('delete');
      closeDrawer();
      toast('Moved to Archive');
    });
  }
  function closeDrawer() { $('#drawer').hidden = true; }

  /* --------------------------- Quick add ----------------------------- */

  let qaType = 'action';
  let qaInit = {};
  let qaSaveCallback = null;
  function openQuickAdd(type = 'action', init = {}, onSave = null) {
    // Change requests use the full editor instead of the simplified Quick Add
    // form, so the create flow has the same fields as the edit flow.
    if (type === 'change') {
      if (curProjectIsMerged()) { toast('Pick a single project to add items.'); return; }
      openChangeRequestEditor(newChangeRequestDraft());
      return;
    }
    qaType = type;
    qaInit = init || {};
    qaSaveCallback = onSave || null;
    $$('.qa-tab').forEach((t) => t.classList.toggle('active', t.dataset.qa === type));
    drawQA();
    $('#quickAdd').hidden = false;
    setTimeout(() => $('#qaBody input, #qaBody select, #qaBody textarea')?.focus(), 30);
  }
  function closeQuickAdd() { $('#quickAdd').hidden = true; }

  function drawQA() {
    const body = $('#qaBody');
    const proj = curProject();
    if (qaType === 'action') {
      const initTitle = qaInit.title || '';
      const initNotes = qaInit.notes || '';
      body.innerHTML = `
        <div class="field"><label>Title</label><input id="qTitle" placeholder="What needs to be done?" value="${escapeHTML(initTitle)}" /></div>
        <div class="qa-row">
          <div class="field"><label>Owner</label>
            <select id="qOwner">${state.people.map((p) => `<option value="${p.id}">${escapeHTML(p.name)}</option>`).join('')}</select>
          </div>
          <div class="field"><label>Originator</label>
            <select id="qOriginator"><option value="">— same as owner</option>${state.people.map((p) => `<option value="${p.id}">${escapeHTML(p.name)}</option>`).join('')}</select>
          </div>
          <div class="field"><label>Originator date</label>
            <input id="qOriginatorDate" type="date" value="${todayISO()}" title="Auto-set to today; editable" />
          </div>
        </div>
        <div class="qa-row">
          <div class="field"><label>Status</label>
            <select id="qStatus">${STATUSES.map((s) => `<option value="${s.id}">${s.name}</option>`).join('')}</select>
          </div>
          <div class="field"><label>Priority</label>
            <select id="qPriorityLevel">${PRIORITY_LEVELS.map((p) => `<option value="${p.id}" ${p.id === (qaInit.priorityLevel || 'med') ? 'selected' : ''}>${p.label}</option>`).join('')}</select>
          </div>
          <div class="field"><label>Due</label><input id="qDue" type="date" value="${todayISO()}" /></div>
        </div>
        <div class="qa-row">
          <div class="field"><label>Component (optional)</label>
            <select id="qComponent"><option value="">—</option>${(proj.components || []).map((pt) => `<option value="${pt.id}" ${pt.id === qaInit.component ? 'selected' : ''}>${escapeHTML(pt.name)}</option>`).join('')}</select>
          </div>
          <div class="field"><label>Deliverable (optional)</label>
            <select id="qDel"><option value="">—</option>${(proj.deliverables || []).map((d) => `<option value="${d.id}">${escapeHTML(d.name)}</option>`).join('')}</select>
          </div>
        </div>
        <div class="field">
          <label>Commitment <span class="muted" id="qCmtVal">100%</span></label>
          <input id="qCmt" type="range" min="5" max="100" step="5" value="100" oninput="document.getElementById('qCmtVal').textContent = this.value + '%';" />
        </div>
        <div class="field"><label>Notes</label><textarea id="qNotes" placeholder="Optional context">${escapeHTML(initNotes)}</textarea></div>`;
    } else if (qaType === "component") {
      const knownCCs = getCostCentres();
      body.innerHTML = `
        <div class="field"><label>Name</label><input id="qName" placeholder="e.g. Power, AOCS, Backend…" /></div>
        <div class="field"><label>Cost centre (optional)</label>
          <select id="qCC">
            <option value="">— none —</option>
            ${knownCCs.map((c) => `<option value="${escapeHTML(c)}">${escapeHTML(c)}</option>`).join('')}
            <option value="__new__">+ New cost centre…</option>
          </select>
        </div>
        <div class="field"><label>Color</label>
          <div id="qComponentColors" class="color-grid">
            ${COMPONENT_COLORS.map((c, i) => `
              <label class="color-swatch" title="${c.name}">
                <input type="radio" name="qComponentColor" value="${c.id}" ${i === 0 ? 'checked' : ''} />
                <span style="background: rgba(${c.rgb},.9);"></span>
                <em>${c.name}</em>
              </label>`).join('')}
          </div>
        </div>`;
    } else if (qaType === 'deliverable') {
      body.innerHTML = `
        <div class="field"><label>Name</label><input id="qName" /></div>
        <div class="qa-row">
          <div class="field"><label>Due</label><input id="qDue" type="date" /></div>
          <div class="field"><label>Status</label>
            <select id="qStatus"><option value="todo">Not started</option><option value="doing">In progress</option><option value="done">Done</option></select>
          </div>
        </div>`;
    } else if (qaType === 'milestone') {
      body.innerHTML = `
        <div class="field"><label>Name</label><input id="qName" /></div>
        <div class="field"><label>Date</label><input id="qDate" type="date" /></div>`;
    } else if (qaType === 'risk') {
      const kind = qaInit.kind === 'opportunity' ? 'opportunity' : 'risk';
      const actionsList = (proj.actions || []).slice().sort((a, b) => a.title.localeCompare(b.title));
      body.innerHTML = `
        <div class="field"><label>Type</label>
          <div class="seg" role="tablist" aria-label="Kind">
            <button type="button" class="seg-btn ${kind === 'risk' ? 'active' : ''}" data-qa-kind="risk">▲ Risk</button>
            <button type="button" class="seg-btn ${kind === 'opportunity' ? 'active' : ''}" data-qa-kind="opportunity">▽ Opportunity</button>
          </div>
        </div>
        <div class="field"><label>Title</label><input id="qTitle" placeholder="${kind === 'opportunity' ? 'Upside event to chase' : 'Downside event to mitigate'}" /></div>
        <div class="qa-row">
          <div class="field"><label>Inherent P (1-5)</label><input id="qProb" type="number" min="1" max="5" value="3" /></div>
          <div class="field"><label>Inherent I (1-5)</label><input id="qImp" type="number" min="1" max="5" value="3" /></div>
        </div>
        <div class="qa-row">
          <div class="field"><label>Residual P (post-action)</label><input id="qProbR" type="number" min="1" max="5" value="2" /></div>
          <div class="field"><label>Residual I (post-action)</label><input id="qImpR" type="number" min="1" max="5" value="2" /></div>
        </div>
        <div class="field"><label>Owner</label>
          <select id="qOwner">${state.people.map((p) => `<option value="${p.id}">${escapeHTML(p.name)}</option>`).join('')}</select>
        </div>
        <div class="field"><label id="qMitLbl">${kind === 'opportunity' ? 'Capture plan' : 'Mitigation'}</label><textarea id="qMit" placeholder="Brief description of the response"></textarea></div>
        <div class="field"><label>Linked action (optional)</label>
          <select id="qActionLink">
            <option value="">— none —</option>
            ${actionsList.map((a) => `<option value="${a.id}">${escapeHTML(a.title)} — ${escapeHTML(personName(a.owner))}</option>`).join('')}
          </select>
        </div>`;
      $$('.seg-btn[data-qa-kind]', body).forEach((b) => {
        b.addEventListener('click', () => {
          qaInit.kind = b.dataset.qaKind;
          $$('.seg-btn[data-qa-kind]', body).forEach((x) => x.classList.toggle('active', x.dataset.qaKind === qaInit.kind));
          $('#qMitLbl').textContent = qaInit.kind === 'opportunity' ? 'Capture plan' : 'Mitigation';
          $('#qTitle').placeholder = qaInit.kind === 'opportunity' ? 'Upside event to chase' : 'Downside event to mitigate';
        });
      });
    } else if (qaType === 'decision') {
      body.innerHTML = `
        <div class="field"><label>Title</label><input id="qTitle" /></div>
        <div class="field"><label>Rationale</label><textarea id="qRat"></textarea></div>
        <div class="qa-row">
          <div class="field"><label>Owner</label>
            <select id="qOwner">${state.people.map((p) => `<option value="${p.id}">${escapeHTML(p.name)}</option>`).join('')}</select>
          </div>
          <div class="field"><label>Date</label><input id="qDate" type="date" value="${todayISO()}" /></div>
        </div>`;
    } else if (qaType === 'meeting') {
      // Progressive disclosure: title / date / time / repeating-toggle are
      // always visible. When repeating is on, reveal a single periodicity
      // row ("Every [n] [unit]") plus optional end-date — and only show the
      // day-of-week selector when the unit is "Week(s)". When off, none
      // of those controls render.
      const repeating = qaInit.mtKind === 'recurring';
      const initUnit  = qaInit.mtUnit || 'week';
      body.innerHTML = `
        <div class="field"><label>Title</label><input id="qTitle" placeholder="e.g. Weekly standup, PDR walkthrough" /></div>
        <div class="qa-row">
          <div class="field"><label id="qDateLbl">${repeating ? 'Start date' : 'Date'}</label><input id="qDate" type="date" value="${todayISO()}" /></div>
          <div class="field"><label>Time <span class="muted">— optional</span></label><input id="qTime" type="time" /></div>
        </div>
        <div class="field">
          <label class="qa-toggle">
            <input type="checkbox" id="qMtRepeats" ${repeating ? 'checked' : ''} />
            <span>Repeating meeting</span>
          </label>
        </div>
        <div id="qMtRecurWrap" ${repeating ? '' : 'hidden'}>
          <div class="qa-row qa-row-tight">
            <div class="field" style="flex: 0 0 auto;">
              <label>Every</label>
              <input id="qInterval" type="number" min="1" max="99" value="1" style="width:64px;" />
            </div>
            <div class="field" style="flex: 1;">
              <label>&nbsp;</label>
              <select id="qUnit">
                <option value="day"   ${initUnit === 'day'   ? 'selected' : ''}>Day(s)</option>
                <option value="week"  ${initUnit === 'week'  ? 'selected' : ''}>Week(s)</option>
                <option value="month" ${initUnit === 'month' ? 'selected' : ''}>Month(s)</option>
              </select>
            </div>
            <div class="field" id="qDowField" ${initUnit === 'week' ? '' : 'hidden'}>
              <label>On</label>
              <select id="qDow">
                <option value="1">Monday</option>
                <option value="2">Tuesday</option>
                <option value="3">Wednesday</option>
                <option value="4">Thursday</option>
                <option value="5">Friday</option>
                <option value="6">Saturday</option>
                <option value="0">Sunday</option>
              </select>
            </div>
          </div>
          <div class="field">
            <label>Ends <span class="muted">— optional</span></label>
            <input id="qEndDate" type="date" />
          </div>
        </div>
        <div class="field"><label>Component <span class="muted">— optional</span></label>
          <select id="qComp">
            <option value="">— None</option>
            ${(proj.components || []).map((cmp) => `<option value="${cmp.id}">${escapeHTML(cmp.name)}</option>`).join('')}
          </select>
        </div>`;
      const repeatChk = body.querySelector('#qMtRepeats');
      const recurWrap = body.querySelector('#qMtRecurWrap');
      const dateInput = body.querySelector('#qDate');
      const unitSel   = body.querySelector('#qUnit');
      const dowField  = body.querySelector('#qDowField');
      const dowSel    = body.querySelector('#qDow');
      const dateLbl   = body.querySelector('#qDateLbl');
      function syncDowFromDate() {
        const v = dateInput.value;
        if (v) dowSel.value = String(parseDate(v).getDay());
      }
      if (repeating && initUnit === 'week') syncDowFromDate();
      repeatChk.addEventListener('change', () => {
        recurWrap.hidden = !repeatChk.checked;
        dateLbl.textContent = repeatChk.checked ? 'Start date' : 'Date';
        qaInit.mtKind = repeatChk.checked ? 'recurring' : 'oneoff';
        if (repeatChk.checked && unitSel.value === 'week') syncDowFromDate();
      });
      unitSel.addEventListener('change', () => {
        qaInit.mtUnit = unitSel.value;
        // Day-of-week selector is only relevant for the 'week' unit.
        dowField.hidden = unitSel.value !== 'week';
        if (unitSel.value === 'week') syncDowFromDate();
      });
      dateInput.addEventListener('change', () => {
        if (repeatChk.checked && unitSel.value === 'week') syncDowFromDate();
      });
    } else if (qaType === 'person') {
      body.innerHTML = `
        <div class="field"><label>Name</label><input id="qName" /></div>
        <div class="qa-row">
          <div class="field"><label>Role</label><input id="qRole" /></div>
          <div class="field"><label>Capacity</label><input id="qCap" type="number" min="1" max="20" value="5" /></div>
        </div>`;
    } else if (qaType === 'project') {
      const tpls = state.templates || [];
      body.innerHTML = `
        <div class="field"><label>Name</label><input id="qName" /></div>
        <div class="field"><label>Description</label><textarea id="qDesc"></textarea></div>
        ${tpls.length ? `
        <div class="field"><label>Start from <span class="muted">— template, or empty</span></label>
          <select id="qFrom">
            <option value="">Empty project</option>
            ${tpls.map((t) => `<option value="${escapeHTML(t.id)}">From template — ${escapeHTML(t.name)}</option>`).join('')}
          </select>
        </div>` : ''}`;
    } else if (qaType === 'link') {
      body.innerHTML = `
        <div class="field"><label>Title</label><input id="qTitle" placeholder="Display name" /></div>
        <div class="field"><label>URL or path</label><input id="qUrl" placeholder="https://… or file:///…" /></div>
        <div class="field"><label>Description</label><textarea id="qDesc" placeholder="What this is and when to use it"></textarea></div>
        <div class="qa-row">
          <div class="field"><label>Component (optional)</label>
            <select id="qComp">
              <option value="">— none —</option>
              ${(proj.components || []).map((c) => `<option value="${c.id}">${escapeHTML(c.name)}</option>`).join('')}
            </select>
          </div>
          <div class="field"><label>Folder</label>
            <select id="qFolder">
              <option value="">— Loose —</option>
              ${(proj.linkFolders || []).map((f) => `<option value="${f.id}" ${f.id === qaInit.folderId ? 'selected' : ''}>${escapeHTML(f.name)}</option>`).join('')}
            </select>
          </div>
        </div>`;
    }
  }

  function saveQA() {
    if (curProjectIsMerged()) {
      toast('Pick a single project to add items.');
      closeQuickAdd();
      return;
    }
    const proj = curProject();
    if (qaType === 'action') {
      const title = $('#qTitle').value.trim();
      if (!title) return toast('Title required');
      const a = {
        id: uid('a'), title,
        owner: $('#qOwner').value,
        originator: $('#qOriginator')?.value || null,
        originatorDate: $('#qOriginatorDate')?.value || todayISO(),
        due: $('#qDue').value || null,
        status: $('#qStatus').value,
        priority: typeof qaInit.priority === 'number' ? qaInit.priority : 0,
        priorityLevel: $('#qPriorityLevel')?.value || qaInit.priorityLevel || 'med',
        commitment: clamp(parseInt($('#qCmt')?.value, 10) || 100, 5, 100),
        component: $('#qComponent')?.value || null,
        deliverable: $('#qDel').value || null,
        milestone: null,
        description: qaInit.description || null,
        notes: $('#qNotes').value || '',
        createdAt: todayISO(), updatedAt: todayISO(),
        history: [{ at: todayISO(), what: 'Created' }],
      };
      proj.actions.push(a);
    } else if (qaType === 'component') {
      const name = $('#qName').value.trim();
      if (!name) return toast('Name required');
      const color = (document.querySelector('input[name="qComponentColor"]:checked')?.value) || COMPONENT_COLORS[0].id;
      let ccRaw = $('#qCC')?.value || '';
      if (ccRaw === '__new__') {
        const code = prompt('New cost-centre name:');
        const trimmed = (code || '').trim();
        if (!trimmed) ccRaw = '';
        else {
          ccRaw = trimmed;
          state.budgets = state.budgets || {};
          state.budgets[ccRaw] = state.budgets[ccRaw] || {};
        }
      }
      proj.components = proj.components || [];
      proj.components.push({ id: uid('cm'), name, color, costCenter: ccRaw || null });
    } else if (qaType === 'deliverable') {
      const name = $('#qName').value.trim();
      if (!name) return toast('Name required');
      proj.deliverables = proj.deliverables || [];
      proj.deliverables.push({ id: uid('d'), name, dueDate: $('#qDue').value || null, status: $('#qStatus').value });
    } else if (qaType === 'milestone') {
      const name = $('#qName').value.trim();
      if (!name) return toast('Name required');
      proj.milestones = proj.milestones || [];
      proj.milestones.push({ id: uid('m'), name, date: $('#qDate').value || null, status: 'todo' });
    } else if (qaType === 'risk') {
      const title = $('#qTitle').value.trim();
      if (!title) return toast('Title required');
      const inhP = clamp(parseInt($('#qProb').value, 10) || 3, 1, 5);
      const inhI = clamp(parseInt($('#qImp').value, 10) || 3, 1, 5);
      const resP = clamp(parseInt($('#qProbR').value, 10) || inhP, 1, 5);
      const resI = clamp(parseInt($('#qImpR').value, 10) || inhI, 1, 5);
      const isOpp = qaInit.kind === 'opportunity';
      proj.risks = proj.risks || [];
      proj.risks.push({
        id: uid(isOpp ? 'o' : 'r'),
        kind: isOpp ? 'opportunity' : 'risk',
        title,
        inherent: { probability: inhP, impact: inhI },
        residual: { probability: resP, impact: resI },
        mitigation: $('#qMit').value || '',
        actionId: $('#qActionLink').value || null,
        owner: $('#qOwner').value,
      });
      // If the R&O page is filtered in a way that would hide this new item,
      // open the filter so the user actually sees what they just added.
      if (roState.kind === (isOpp ? 'risk' : 'opportunity')) roState.kind = 'all';
    } else if (qaType === 'decision') {
      const title = $('#qTitle').value.trim();
      if (!title) return toast('Title required');
      proj.decisions = proj.decisions || [];
      proj.decisions.push({
        id: uid('dec'), title,
        rationale: $('#qRat').value || '',
        owner: $('#qOwner').value,
        date: $('#qDate').value || todayISO(),
      });
    } else if (qaType === 'meeting') {
      const title = $('#qTitle').value.trim();
      if (!title) return toast('Title required');
      const repeating = !!$('#qMtRepeats')?.checked;
      const date = $('#qDate').value || todayISO();
      const time = $('#qTime').value || null;
      const component = $('#qComp')?.value || null;
      const m = { id: uid('mtg'), title, time, endDate: null, component };
      if (!repeating) {
        m.kind = 'oneoff';
        m.date = date;
      } else {
        const interval = Math.max(1, parseInt($('#qInterval').value, 10) || 1);
        const unit = $('#qUnit').value === 'day' ? 'day'
                   : $('#qUnit').value === 'month' ? 'month'
                   : 'week';
        const ed = $('#qEndDate')?.value || '';
        if (ed && ed < date) { toast('End date can\'t be before the start date'); return; }
        m.kind      = 'recurring';
        m.recurUnit = unit;
        m.interval  = interval;
        m.startDate = date;
        m.endDate   = ed || null;
        if (unit === 'week') m.dayOfWeek = parseInt($('#qDow').value, 10);
      }
      proj.meetings = proj.meetings || [];
      proj.meetings.push(m);
    } else if (qaType === 'person') {
      const name = $('#qName').value.trim();
      if (!name) return toast('Name required');
      state.people.push({
        id: uid('p'), name,
        role: $('#qRole').value || '',
        capacity: clamp(parseInt($('#qCap').value, 10) || 5, 1, 20),
      });
    } else if (qaType === 'project') {
      const name = $('#qName').value.trim();
      if (!name) return toast('Name required');
      const np = {
        id: uid('pr'), name,
        description: $('#qDesc').value || '',
        actions: [], deliverables: [], milestones: [], risks: [], decisions: [], changes: [], components: [], links: [],
      };
      const fromId = $('#qFrom')?.value;
      if (fromId) {
        const tpl = (state.templates || []).find((t) => t.id === fromId);
        if (tpl) {
          applyTemplateToProject(np, tpl);
          np.templateOf = tpl.id;
        }
      }
      state.projects.push(np);
      state.currentProjectId = np.id;
    } else if (qaType === 'link') {
      const url = $('#qUrl').value.trim();
      if (!url) return toast('URL required');
      const title = $('#qTitle').value.trim() || url;
      proj.links = proj.links || [];
      proj.links.push({
        id: uid('lk'), title, url,
        description: $('#qDesc').value || '',
        component: $('#qComp').value || null,
        folderId: $('#qFolder')?.value || qaInit.folderId || null,
      });
    }
    // Capture the just-pushed action for callback consumers (e.g. notes panel).
    let createdAction = null;
    if (qaType === 'action') createdAction = proj.actions[proj.actions.length - 1];
    const cb = qaSaveCallback;
    qaSaveCallback = null;
    commit('add');
    closeQuickAdd();
    const addedLabel = qaType === 'risk'
      ? (qaInit.kind === 'opportunity' ? 'Opportunity added' : 'Risk added')
      : 'Added';
    toast(addedLabel);
    if (cb && createdAction) cb(createdAction);
  }

  /* -------------------------- Context menu --------------------------- */

  let _ctxOutsideHandler = null;
  // Small floating picker anchored under a chip. Used by open-point
  // criticality + priority chips so they replace the old <select>s
  // without losing the all-options-visible affordance.
  let _levelPopOutsideHandler = null;
  function closeLevelPopover() {
    document.querySelectorAll('.level-pop').forEach((m) => m.remove());
    if (_levelPopOutsideHandler) {
      document.removeEventListener('mousedown', _levelPopOutsideHandler);
      document.removeEventListener('keydown', _levelPopOutsideHandler);
      _levelPopOutsideHandler = null;
    }
  }
  function showLevelPopover(anchorEl, kind, currentValue, onPick) {
    closeLevelPopover();
    closeContextMenu();
    const opts = (kind === 'criticality')
      ? ['low', 'med', 'high', 'critical'].map((id) => ({ id, label: CRITICALITY_LABEL[id], rgb: CRITICALITY_RGB[id] }))
      : PRIORITY_LEVELS.map((p) => ({ id: p.id, label: p.label, rgb: p.rgb }));
    const subtitle = kind === 'criticality'
      ? 'Severity if not addressed'
      : 'Urgency to act';
    const heading = kind === 'criticality' ? 'Criticality' : 'Priority';

    const pop = document.createElement('div');
    pop.className = 'level-pop';
    pop.innerHTML = `
      <div class="level-pop-head">
        <div class="level-pop-title">${escapeHTML(heading)}</div>
        <div class="level-pop-sub">${escapeHTML(subtitle)}</div>
      </div>
      <div class="level-pop-list">
        ${opts.map((o) => `
          <button type="button" class="level-pop-item ${o.id === currentValue ? 'sel' : ''}" data-val="${escapeHTML(o.id)}">
            <span class="level-pop-dot" style="background:rgb(${o.rgb})"></span>
            <span class="level-pop-label">${escapeHTML(o.label)}</span>
            ${o.id === currentValue ? '<span class="level-pop-check">✓</span>' : ''}
          </button>`).join('')}
      </div>`;
    document.body.appendChild(pop);

    // Anchor below the chip; clamp to viewport so right-edge chips don't
    // push the popover off-screen
    const r = anchorEl.getBoundingClientRect();
    const w = pop.getBoundingClientRect().width;
    const h = pop.getBoundingClientRect().height;
    let x = r.left;
    let y = r.bottom + 6;
    if (x + w > innerWidth - 8)  x = innerWidth - w - 8;
    if (y + h > innerHeight - 8) y = r.top - h - 6; // flip above
    pop.style.left = Math.max(8, x) + 'px';
    pop.style.top  = Math.max(8, y) + 'px';

    pop.querySelectorAll('.level-pop-item').forEach((btn) => {
      btn.addEventListener('click', (e) => {
        e.stopPropagation();
        const val = btn.dataset.val;
        closeLevelPopover();
        try { onPick(val); } catch (err) { console.error(err); }
      });
    });

    // Outside-click + Escape close (popover is transient — same rule as
    // the existing ctx-menu / palette).
    _levelPopOutsideHandler = (e) => {
      if (e.type === 'keydown') {
        if (e.key === 'Escape') closeLevelPopover();
        return;
      }
      if (pop.contains(e.target)) return;
      closeLevelPopover();
    };
    setTimeout(() => {
      document.addEventListener('mousedown', _levelPopOutsideHandler);
      document.addEventListener('keydown', _levelPopOutsideHandler);
    }, 0);
  }

  function showContextMenu(x, y, items) {
    closeContextMenu();
    const menu = document.createElement('div');
    menu.className = 'ctx-menu';
    menu.style.left = x + 'px';
    menu.style.top = y + 'px';
    items.forEach((it) => {
      if (it.divider) {
        const d = document.createElement('div');
        d.className = 'ctx-divider';
        menu.appendChild(d);
        return;
      }
      const b = document.createElement('button');
      b.className = 'ctx-item' + (it.danger ? ' danger' : '');
      b.innerHTML = `<span class="ctx-icon">${it.icon || ''}</span><span>${escapeHTML(it.label)}</span>`;
      b.addEventListener('click', (e) => {
        e.stopPropagation();
        closeContextMenu();
        try { it.onClick?.(); } catch (err) { console.error(err); }
      });
      menu.appendChild(b);
    });
    document.body.appendChild(menu);
    const r = menu.getBoundingClientRect();
    if (r.right > innerWidth)  menu.style.left = (innerWidth - r.width - 6) + 'px';
    if (r.bottom > innerHeight) menu.style.top  = (innerHeight - r.height - 6) + 'px';
    // Outside-click closer: only fires when click is OUTSIDE the menu, so
    // item button clicks aren't pre-empted by removing the menu mid-event.
    _ctxOutsideHandler = (e) => {
      if (menu.contains(e.target)) return;
      closeContextMenu();
    };
    setTimeout(() => document.addEventListener('mousedown', _ctxOutsideHandler), 0);
  }
  function closeContextMenu() {
    document.querySelectorAll('.ctx-menu').forEach((m) => m.remove());
    if (_ctxOutsideHandler) {
      document.removeEventListener('mousedown', _ctxOutsideHandler);
      _ctxOutsideHandler = null;
    }
  }

  function openPersonEditor(personId) {
    const p = state.people.find((x) => x.id === personId);
    if (!p) return;
    $('#drawerTitle').textContent = 'Edit person';
    $('#drawerBody').innerHTML = `
      <div class="field"><label>Name</label><input id="pEdName" value="${escapeHTML(p.name)}" /></div>
      <div class="field"><label>Role</label><input id="pEdRole" value="${escapeHTML(p.role || '')}" placeholder="Job title" /></div>
      <div class="field"><label>Expertise / skills</label><textarea id="pEdSkills" placeholder="e.g. Avionics design, EMC testing">${escapeHTML(p.expertise || '')}</textarea></div>
      <div style="display:grid; grid-template-columns:1fr 1fr; gap:10px;">
        <div class="field"><label>Capacity (% of FTE)</label><input id="pEdCap" type="number" min="0" max="200" value="${p.capacity || 100}" title="100% = full-time. 1 FTE = 8h/day × 5 days/week, 212 working days/year." /></div>
        <div class="field"><label>Hourly rate (€/h)</label><input id="pEdRate" type="number" min="0" step="5" value="${p.hourlyRate || 100}" /></div>
      </div>
      <div style="display:flex; gap:8px; margin-top:6px;">
        <button class="primary" id="pEdSave">Save</button>
        <button class="ghost" id="pEdDelete" style="margin-left:auto; color:var(--bad);">Delete</button>
      </div>`;
    $('#drawer').hidden = false;
    $('#pEdSave').addEventListener('click', () => {
      const oldName = p.name;
      p.name = $('#pEdName').value.trim() || p.name;
      p.role = $('#pEdRole').value.trim();
      p.expertise = $('#pEdSkills').value.trim();
      p.capacity = clamp(parseInt($('#pEdCap').value, 10) || 100, 0, 200);
      p.hourlyRate = Math.max(0, parseFloat($('#pEdRate').value) || 100);
      commit('person-edit');
      closeDrawer();
      toast(oldName !== p.name ? 'Renamed to ' + p.name : 'Saved');
    });
    $('#pEdDelete').addEventListener('click', () => {
      const open = state.projects.flatMap((pr) => pr.actions || []).filter((a) => a.owner === p.id && a.status !== 'done').length;
      if (!confirm(`Delete ${p.name}?` + (open ? ` (${open} open action${open === 1 ? '' : 's'} will be unassigned).` : ''))) return;
      state.projects.forEach((pr) => (pr.actions || []).forEach((a) => { if (a.owner === p.id) a.owner = null; }));
      state.people = state.people.filter((x) => x.id !== p.id);
      commit('person-delete');
      closeDrawer();
      toast('Deleted');
    });
  }

  /* ---------------------------- Notes panel -------------------------- */

  let notesSaveTimer = null;
  let savedRange = null; // selection range in the notes body, captured before opening modals

  function notesIsOpen() { return !$('#notesPanel').hidden; }

  // Notes panel width (px) — clamped to keep the panel usable on small screens
  // and to leave room for the main view.
  const NOTES_W_MIN = 240;
  const NOTES_W_MAX = 720;
  const NOTES_W_DEFAULT = 340;
  /**
   * Generic drag-to-reorder for any list. Each row needs a `.row-grip`
   * element (the drag handle) and a data attribute carrying the item's id.
   * @param {Element} listEl       container that holds the rows directly
   * @param {object}  opts
   *   - rowSelector  CSS selector that matches the rows (e.g. '.row[data-component-id]')
   *   - idAttr       dataset key on each row (e.g. 'componentId')
   *   - getArray     () => array of items in current order (live reference)
   *   - setOrder     (idsInOrder: string[]) => void; mutate the source array
   *   - commitName   action name passed to commit() after a real reorder
   */
  function wireListReorder(listEl, opts) {
    if (!listEl) return;
    const rows = listEl.querySelectorAll(opts.rowSelector);
    rows.forEach((row) => {
      const grip = row.querySelector('.row-grip');
      if (!grip) return;
      grip.addEventListener('mousedown', (e) => {
        e.preventDefault();
        e.stopPropagation();
        row.classList.add('dragging');
        document.body.classList.add('is-row-dragging');
        const onMove = (em) => {
          const siblings = [...listEl.querySelectorAll(opts.rowSelector + ':not(.dragging)')];
          const after = siblings.find((sib) => {
            const r = sib.getBoundingClientRect();
            return em.clientY < r.top + r.height / 2;
          });
          if (after) listEl.insertBefore(row, after);
          else listEl.appendChild(row);
        };
        const onUp = () => {
          row.classList.remove('dragging');
          document.body.classList.remove('is-row-dragging');
          window.removeEventListener('mousemove', onMove);
          window.removeEventListener('mouseup', onUp);
          const newOrder = [...listEl.querySelectorAll(opts.rowSelector)].map((r) => r.dataset[opts.idAttr]);
          const arr = opts.getArray();
          const before = arr.map((x) => x.id).join(',');
          if (newOrder.join(',') === before) return; // no-op
          opts.setOrder(newOrder);
          commit(opts.commitName || 'reorder');
        };
        window.addEventListener('mousemove', onMove);
        window.addEventListener('mouseup', onUp);
      });
    });
  }

  // Tiny inline grip glyph used by every reorderable list row
  const ROW_GRIP_HTML = '<span class="row-grip" title="Drag to reorder" aria-hidden="true">⋮⋮</span>';

  function clampNotesWidth(w) {
    const max = Math.min(NOTES_W_MAX, Math.max(NOTES_W_MIN, innerWidth - 360));
    return Math.max(NOTES_W_MIN, Math.min(max, w));
  }
  function applyNotesWidth(w) {
    const px = clampNotesWidth(w);
    document.documentElement.style.setProperty('--notes-w', px + 'px');
    const app = $('#app');
    if (app) app.style.setProperty('--notes-w', px + 'px');
    return px;
  }

  function applyNotesPanel() {
    const open = state.notesOpen === true;
    const panel = $('#notesPanel');
    const app = $('#app');
    if (!panel || !app) return;
    state.settings = state.settings || {};
    applyNotesWidth(state.settings.notesWidth || NOTES_W_DEFAULT);
    if (open) {
      panel.hidden = false;
      app.classList.add('notes-open');
      loadNotesForCurrentProject();
    } else {
      panel.hidden = true;
      app.classList.remove('notes-open');
    }
  }

  function loadNotesForCurrentProject() {
    const proj = curProject();
    if (!proj) return;
    state.notes = state.notes || {};
    const body = $('#notesBody');
    const html = state.notes[proj.id];
    body.innerHTML = html || `<p><i>Notes for <b>${escapeHTML(proj.name)}</b> — type freely. Use the toolbar to format and to insert actions assigned to people.</i></p><p></p>`;
    $('#notesMeta').textContent = proj.name;
    buildNotesToc();
  }

  function saveNotesNow() {
    const proj = curProject();
    if (!proj) return;
    state.notes = state.notes || {};
    state.notes[proj.id] = $('#notesBody').innerHTML;
    saveState();
    const s = $('#notesSaved');
    if (s) { s.textContent = 'Saved'; s.classList.remove('saving'); }
    buildNotesToc();
  }

  // Auto-built table of contents — scans the notes for H1–H6 elements,
  // assigns stable ids, and renders a compact clickable list above the
  // editable body. Hidden when there are no headings (keeps the panel calm).
  function buildNotesToc() {
    const body = $('#notesBody');
    const toc  = $('#notesToc');
    if (!body || !toc) return;
    const headings = body.querySelectorAll('h1, h2, h3, h4, h5, h6');
    const items = [];
    headings.forEach((h, i) => {
      const text = (h.textContent || '').trim();
      if (!text) return;
      // Stable id per heading position; preserve any author-set id.
      if (!h.id) {
        const slug = text.toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/^-|-$/g, '').slice(0, 32) || 'h';
        h.id = `nt-h-${i}-${slug}`;
      }
      items.push({ id: h.id, text, lvl: parseInt(h.tagName[1], 10) || 3 });
    });
    if (!items.length) { toc.hidden = true; toc.innerHTML = ''; return; }
    toc.innerHTML = `
      <div class="notes-toc-head">Contents</div>
      <ol class="notes-toc-list">
        ${items.map((it) => `
          <li class="notes-toc-item lvl-${it.lvl}">
            <a href="#${escapeHTML(it.id)}" data-toc-target="${escapeHTML(it.id)}">${escapeHTML(it.text)}</a>
          </li>`).join('')}
      </ol>`;
    toc.hidden = false;
  }
  function scheduleNotesSave() {
    clearTimeout(notesSaveTimer);
    const s = $('#notesSaved');
    if (s) { s.textContent = 'Saving…'; s.classList.add('saving'); }
    notesSaveTimer = setTimeout(saveNotesNow, 350);
  }

  function snapshotSelection() {
    const sel = window.getSelection();
    const body = $('#notesBody');
    if (sel.rangeCount && body.contains(sel.anchorNode)) {
      savedRange = sel.getRangeAt(0).cloneRange();
    }
  }
  function restoreSelection() {
    if (!savedRange) {
      const body = $('#notesBody');
      body.focus();
      const sel = window.getSelection();
      const range = document.createRange();
      range.selectNodeContents(body);
      range.collapse(false);
      sel.removeAllRanges(); sel.addRange(range);
      return;
    }
    const sel = window.getSelection();
    sel.removeAllRanges();
    sel.addRange(savedRange);
    $('#notesBody').focus();
  }

  function insertActionChip(action) {
    restoreSelection();
    const due = action.due ? fmtDate(action.due) : '—';
    const owner = personName(action.owner);
    const safe = (s) => String(s).replace(/[<>"&]/g, (c) =>
      ({ '<': '&lt;', '>': '&gt;', '"': '&quot;', '&': '&amp;' }[c]));
    const html = ` <span class="note-chip" contenteditable="false" data-action-id="${action.id}">` +
      `<span class="chip-mark">✓</span>` +
      `<b>${safe(action.title)}</b>` +
      `<span class="chip-meta">${safe(owner)} · ${safe(due)}</span>` +
      `</span>&nbsp;`;
    document.execCommand('insertHTML', false, html);
    scheduleNotesSave();
  }

  function insertPersonChip(person) {
    restoreSelection();
    const safe = (s) => String(s).replace(/[<>"&]/g, (c) =>
      ({ '<': '&lt;', '>': '&gt;', '"': '&quot;', '&': '&amp;' }[c]));
    const html = ` <span class="note-chip person-chip" contenteditable="false" data-person-id="${person.id}">` +
      `<span class="chip-mark">@</span><b>${safe(person.name)}</b></span>&nbsp;`;
    document.execCommand('insertHTML', false, html);
    scheduleNotesSave();
  }

  /* --- Inline @/# autocomplete inside the meeting-notes editor ---
   * Typing `@` opens a people picker; typing `#` opens an action picker.
   * The popup filters as the user types; ArrowUp / ArrowDown navigate,
   * Enter / Tab insert the chip, Escape dismisses. */
  let notesAcState = null;

  function getMentionContext(body) {
    const sel = window.getSelection();
    if (!sel.rangeCount) return null;
    const range = sel.getRangeAt(0);
    if (!range.collapsed) return null;
    if (!body.contains(range.startContainer)) return null;
    const node = range.startContainer;
    if (node.nodeType !== Node.TEXT_NODE) return null;
    const text = node.textContent.slice(0, range.startOffset);
    // Trigger at start-of-line or after whitespace; query allows letters,
    // digits, hyphen, underscore, and a single space so "Sofia R" works.
    const m = text.match(/(?:^|[\s ])([@#])([\w\-]*(?: [\w\-]*)?)$/);
    if (!m) return null;
    const kind = m[1];
    const query = m[2];
    const triggerStart = range.startOffset - query.length - 1;
    if (triggerStart < 0) return null;
    const triggerRange = document.createRange();
    triggerRange.setStart(node, triggerStart);
    triggerRange.setEnd(node, range.startOffset);
    return { kind, query, triggerRange };
  }

  function showNotesAutocomplete(ctx) {
    const popup = $('#notesAutocomplete');
    if (!popup) return;
    const q = ctx.query.toLowerCase();
    let items;
    if (ctx.kind === '@') {
      items = state.people
        .filter((p) => p.name.toLowerCase().includes(q))
        .slice(0, 8);
    } else {
      items = state.projects
        .flatMap((proj) => (proj.actions || []).filter((a) => !a.deletedAt).map((a) => ({ ...a, _proj: proj })))
        .filter((a) => a.title.toLowerCase().includes(q))
        .slice(0, 8);
    }
    if (!items.length) { hideNotesAutocomplete(); return; }
    notesAcState = { kind: ctx.kind, query: ctx.query, items, idx: 0, triggerRange: ctx.triggerRange };
    renderNotesAutocomplete();
    positionNotesAutocomplete();
  }

  function hideNotesAutocomplete() {
    notesAcState = null;
    const popup = $('#notesAutocomplete');
    if (popup) { popup.hidden = true; popup.innerHTML = ''; }
  }

  function renderNotesAutocomplete() {
    const popup = $('#notesAutocomplete');
    if (!popup || !notesAcState) return;
    const { kind, items, idx } = notesAcState;
    popup.innerHTML = items.map((it, i) => {
      const sel = i === idx ? 'active' : '';
      if (kind === '@') {
        return `<button type="button" class="ac-item ${sel}" data-ac-idx="${i}" role="option">
          <span class="avatar">${initials(it.name)}</span>
          <span class="ac-text"><span class="ac-name">${escapeHTML(it.name)}</span><span class="ac-meta">${escapeHTML(it.role || '')}</span></span>
        </button>`;
      }
      const due = it.due ? fmtDate(it.due) : '—';
      return `<button type="button" class="ac-item ${sel}" data-ac-idx="${i}" role="option">
        <span class="ac-icon">▤</span>
        <span class="ac-text"><span class="ac-name">${escapeHTML(it.title)}</span><span class="ac-meta">${escapeHTML(personName(it.owner))} · ${due}</span></span>
      </button>`;
    }).join('');
    popup.hidden = false;
  }

  function positionNotesAutocomplete() {
    const popup = $('#notesAutocomplete');
    if (!popup || !notesAcState) return;
    const r = notesAcState.triggerRange.getBoundingClientRect();
    let left = r.left;
    let top  = r.bottom + 4;
    // Keep on-screen — flip above if there's no room below
    const popupRect = popup.getBoundingClientRect();
    if (top + popupRect.height > innerHeight - 8) {
      top = Math.max(8, r.top - popupRect.height - 4);
    }
    if (left + popupRect.width > innerWidth - 8) {
      left = Math.max(8, innerWidth - popupRect.width - 8);
    }
    popup.style.left = left + 'px';
    popup.style.top  = top  + 'px';
  }

  function selectAutocompleteItem(idx) {
    if (!notesAcState) return;
    const it = notesAcState.items[idx];
    if (!it) return;
    const range = notesAcState.triggerRange;
    range.deleteContents();
    const sel = window.getSelection();
    sel.removeAllRanges();
    sel.addRange(range);
    savedRange = range.cloneRange();
    if (notesAcState.kind === '@') insertPersonChip(it);
    else insertActionChip(it);
    hideNotesAutocomplete();
  }

  function refreshNoteChips() {
    const body = $('#notesBody');
    if (!body) return;
    body.querySelectorAll('.note-chip').forEach((chip) => {
      const id = chip.dataset.actionId;
      const a = state.projects.flatMap((p) => p.actions || []).find((x) => x.id === id);
      if (!a) {
        chip.classList.add('chip-stale');
        return;
      }
      chip.classList.remove('chip-stale');
      const mark = a.status === 'done' ? '✓' : a.status === 'blocked' ? '⨯' : a.status === 'doing' ? '◐' : '○';
      const markEl = chip.querySelector('.chip-mark');
      if (markEl) markEl.textContent = mark;
      chip.classList.toggle('done', a.status === 'done');
      chip.classList.toggle('blocked', a.status === 'blocked');
    });
  }

  function wireNotesPanel() {
    const panel = $('#notesPanel');
    if (!panel || panel.dataset.wired === '1') return;
    panel.dataset.wired = '1';

    $('#btnNotesClose').addEventListener('click', () => {
      state.notesOpen = false;
      saveState();
      applyNotesPanel();
    });

    // Drag-to-resize on the panel's left edge
    const handle = $('#notesResize');
    if (handle) {
      handle.addEventListener('mousedown', (e) => {
        e.preventDefault();
        const startX = e.clientX;
        const startW = parseFloat(getComputedStyle(document.documentElement).getPropertyValue('--notes-w')) || NOTES_W_DEFAULT;
        handle.classList.add('dragging');
        document.body.classList.add('is-notes-resizing');
        const onMove = (em) => {
          // Dragging right shrinks the panel (cursor moves toward main content);
          // dragging left expands it.
          const delta = startX - em.clientX;
          applyNotesWidth(startW + delta);
        };
        const onUp = () => {
          window.removeEventListener('mousemove', onMove);
          window.removeEventListener('mouseup', onUp);
          handle.classList.remove('dragging');
          document.body.classList.remove('is-notes-resizing');
          // Persist the final width
          const px = parseFloat(getComputedStyle(document.documentElement).getPropertyValue('--notes-w')) || NOTES_W_DEFAULT;
          state.settings = state.settings || {};
          state.settings.notesWidth = Math.round(px);
          saveState();
        };
        window.addEventListener('mousemove', onMove);
        window.addEventListener('mouseup', onUp);
      });
      // Double-click to reset to the default width
      handle.addEventListener('dblclick', () => {
        applyNotesWidth(NOTES_W_DEFAULT);
        state.settings = state.settings || {};
        state.settings.notesWidth = NOTES_W_DEFAULT;
        saveState();
      });
    }
    // Re-clamp when the window resizes (keeps the panel usable if the user
    // shrinks the viewport while the panel is wide).
    window.addEventListener('resize', () => {
      if (state.notesOpen) {
        applyNotesWidth(state.settings?.notesWidth || NOTES_W_DEFAULT);
      }
    });

    // Toolbar formatting
    panel.querySelectorAll('.notes-toolbar [data-cmd]').forEach((btn) => {
      // mousedown to preserve selection in contenteditable
      btn.addEventListener('mousedown', (e) => {
        e.preventDefault();
        document.execCommand(btn.dataset.cmd, false, btn.dataset.arg || null);
        $('#notesBody').focus();
        scheduleNotesSave();
      });
    });
    // Font-colour picker — applies foreColor to current selection. The bar
    // beneath the "A" glyph previews the most recently chosen colour.
    const ntColor    = $('#ntColorInput');
    const ntColorBar = $('#ntColorBar');
    if (ntColor) {
      // Restore last-used colour from settings
      const last = (state.settings && state.settings.notesColor) || '#6ea8ff';
      ntColor.value = last;
      if (ntColorBar) ntColorBar.style.background = last;
      ntColor.addEventListener('input', (e) => {
        const c = e.target.value;
        $('#notesBody').focus();
        document.execCommand('foreColor', false, c);
        if (ntColorBar) ntColorBar.style.background = c;
        state.settings = state.settings || {};
        state.settings.notesColor = c;
        scheduleNotesSave();
      });
      // Keep the swatch from stealing focus from the editor on click
      ntColor.parentElement.addEventListener('mousedown', (e) => { if (e.target === ntColor.parentElement) e.preventDefault(); });
    }

    $('#btnNotesAction').addEventListener('mousedown', (e) => {
      e.preventDefault();
      snapshotSelection();
      openQuickAdd('action', {}, (action) => {
        insertActionChip(action);
      });
    });

    const body = $('#notesBody');
    body.addEventListener('input', () => {
      scheduleNotesSave();
      // Update the @/# autocomplete based on what's now under the caret
      const ctx = getMentionContext(body);
      if (ctx) showNotesAutocomplete(ctx);
      else hideNotesAutocomplete();
    });
    body.addEventListener('keydown', (e) => {
      // Autocomplete navigation takes priority when the popup is open
      if (notesAcState) {
        if (e.key === 'ArrowDown') {
          e.preventDefault();
          notesAcState.idx = (notesAcState.idx + 1) % notesAcState.items.length;
          renderNotesAutocomplete();
          return;
        }
        if (e.key === 'ArrowUp') {
          e.preventDefault();
          notesAcState.idx = (notesAcState.idx - 1 + notesAcState.items.length) % notesAcState.items.length;
          renderNotesAutocomplete();
          return;
        }
        if (e.key === 'Enter' || e.key === 'Tab') {
          e.preventDefault();
          selectAutocompleteItem(notesAcState.idx);
          return;
        }
        if (e.key === 'Escape') {
          e.preventDefault();
          hideNotesAutocomplete();
          return;
        }
      }
      // Shift+Cmd/Ctrl+A → insert (new) action
      if ((e.metaKey || e.ctrlKey) && e.shiftKey && e.key.toLowerCase() === 'a') {
        e.preventDefault();
        snapshotSelection();
        openQuickAdd('action', {}, (action) => insertActionChip(action));
      }
    });
    body.addEventListener('click', (e) => {
      const chip = e.target.closest('.note-chip');
      if (!chip) return;
      e.preventDefault();
      // Person chips → filter Register to that person; action chips → drawer
      if (chip.classList.contains('person-chip') && chip.dataset.personId) {
        applyTopbarFilter({ owner: chip.dataset.personId, view: 'register' });
      } else if (chip.dataset.actionId) {
        openDrawer(chip.dataset.actionId);
      }
    });
    // Close the autocomplete when the body loses focus or scrolls
    body.addEventListener('blur',   () => setTimeout(hideNotesAutocomplete, 150));
    body.addEventListener('scroll', hideNotesAutocomplete);
    // Click an autocomplete item → insert
    const popup = $('#notesAutocomplete');
    if (popup) {
      popup.addEventListener('mousedown', (e) => {
        e.preventDefault(); // keep selection in body
        const btn = e.target.closest('.ac-item[data-ac-idx]');
        if (!btn) return;
        selectAutocompleteItem(parseInt(btn.dataset.acIdx, 10));
      });
      popup.addEventListener('mouseover', (e) => {
        const btn = e.target.closest('.ac-item[data-ac-idx]');
        if (!btn || !notesAcState) return;
        const idx = parseInt(btn.dataset.acIdx, 10);
        if (idx !== notesAcState.idx) {
          notesAcState.idx = idx;
          renderNotesAutocomplete();
        }
      });
    }

    // TOC navigation — click a contents entry → scroll the heading into view
    const toc = $('#notesToc');
    if (toc) {
      toc.addEventListener('click', (e) => {
        const a = e.target.closest('a[data-toc-target]');
        if (!a) return;
        e.preventDefault();
        const id = a.dataset.tocTarget;
        const target = body.querySelector('#' + (window.CSS && CSS.escape ? CSS.escape(id) : id));
        if (target) {
          target.scrollIntoView({ behavior: 'smooth', block: 'start' });
          // Subtle pulse to confirm where we landed
          target.classList.add('nt-flash');
          setTimeout(() => target.classList.remove('nt-flash'), 900);
        }
      });
    }
  }

  /* --------------------------- Import/Export ------------------------- */

  // Current export schema version. Bump on a breaking schema change.
  const EXPORT_SCHEMA_VERSION = 2;

  // Walk every collection and backfill defaults so that any state — produced by
  // an older app version, hand-edited JSON, or imported from elsewhere — has all
  // fields the rest of the app expects. Idempotent.
  function normalizeState(s) {
    if (!s || typeof s !== 'object') throw new Error('Invalid state');
    s.people = Array.isArray(s.people) ? s.people : [];
    s.projects = Array.isArray(s.projects) ? s.projects : [];
    s.settings = s.settings || {};
    s.settings.theme = s.settings.theme || 'dark';
    s.settings.holidayCountries = s.settings.holidayCountries || [];
    if (s.settings.budgetView !== 'hours' && s.settings.budgetView !== 'cost') s.settings.budgetView = 'cost';
    if (s.settings.budgetGroupBy !== 'component' && s.settings.budgetGroupBy !== 'person') s.settings.budgetGroupBy = 'component';
    // Notes panel width in pixels — added later; default keeps full-screen layouts breathing.
    if (typeof s.settings.notesWidth !== 'number' || s.settings.notesWidth < 200 || s.settings.notesWidth > 800) {
      s.settings.notesWidth = 340;
    }
    s.budgets = s.budgets || {};
    s.notes = s.notes || {};
    s.notesOpen = !!s.notesOpen;
    // Personal to-do list — project-independent, lives at the state level.
    // Each item: { id, text, done }. Defaulted to [] for retrocompat.
    s.todos = Array.isArray(s.todos) ? s.todos : [];
    s.todos.forEach((t) => {
      t.id = t.id || uid('td');
      t.text = t.text || '';
      t.done = !!t.done;
    });
    // Auto-backup configuration (folder handle lives in IndexedDB).
    s.settings.autoBackup = s.settings.autoBackup || {};
    const ab = s.settings.autoBackup;
    if (typeof ab.enabled !== 'boolean') ab.enabled = false;
    const allowedMins = [5, 15, 30, 60, 360, 1440];
    if (!allowedMins.includes(ab.intervalMinutes)) ab.intervalMinutes = 60;
    if (typeof ab.dirName !== 'string') ab.dirName = '';
    if (typeof ab.lastBackupAt !== 'string') ab.lastBackupAt = null;
    // Phase A — added flags / structures used by later phases. All retrocompat.
    if (typeof s.settings.tourSeen     !== 'boolean') s.settings.tourSeen = false;
    if (typeof s.settings.notifyEnabled !== 'boolean') s.settings.notifyEnabled = false;
    if (typeof s.settings.sidebarGroups !== 'object' || !s.settings.sidebarGroups) {
      s.settings.sidebarGroups = { workspace: true, work: true, insight: true };
    }
    // Migration: legacy 'project' / 'engineering' group keys were renamed
    // to 'work' / 'insight' when the sidebar was re-grouped by intent.
    // Carry over the user's collapse state so a tidy sidebar stays tidy.
    if (s.settings.sidebarGroups.project !== undefined && s.settings.sidebarGroups.work === undefined) {
      s.settings.sidebarGroups.work = s.settings.sidebarGroups.project;
      delete s.settings.sidebarGroups.project;
    }
    if (s.settings.sidebarGroups.engineering !== undefined && s.settings.sidebarGroups.insight === undefined) {
      s.settings.sidebarGroups.insight = s.settings.sidebarGroups.engineering;
      delete s.settings.sidebarGroups.engineering;
    }
    // Ensure the three current keys exist so first-render doesn't fall
    // through to undefined (which the helper treats as 'expanded').
    ['workspace', 'work', 'insight'].forEach((k) => {
      if (s.settings.sidebarGroups[k] === undefined) s.settings.sidebarGroups[k] = true;
    });
    s.inbox = s.inbox || {};
    s.inbox.dismissed = Array.isArray(s.inbox.dismissed) ? s.inbox.dismissed : [];
    s.templates = Array.isArray(s.templates) ? s.templates : [];
    s.templates.forEach((t) => {
      t.id = t.id || uid('tpl');
      t.name = t.name || 'Template';
      t.createdAt = t.createdAt || todayISO();
      t.shape = t.shape || {};
    });
    s.currentView = s.currentView || 'board';
    if (s.currentView === 'teams') s.currentView = 'people';
    if (!s.currentProjectId || (s.currentProjectId !== '__all__' && !s.projects.some((p) => p.id === s.currentProjectId))) {
      s.currentProjectId = s.projects[0]?.id || null;
    }

    // People — capacity is % of FTE; old data using small numbers is migrated.
    s.people.forEach((p) => {
      if (!p.id) p.id = uid('p');
      if (typeof p.capacity !== 'number') p.capacity = 100;
      if (p.capacity < 30) p.capacity = Math.round(p.capacity * 20); // legacy unit
      if (typeof p.hourlyRate !== 'number') p.hourlyRate = 100;
      p.role = p.role || '';
    });

    s.projects.forEach((p) => {
      p.id = p.id || uid('pr');
      p.name = p.name || 'Untitled project';
      p.description = p.description || '';
      // Per-collection arrays (presence is enough for renderers to short-circuit)
      p.actions      = Array.isArray(p.actions) ? p.actions : [];
      p.deliverables = Array.isArray(p.deliverables) ? p.deliverables : [];
      p.milestones   = Array.isArray(p.milestones) ? p.milestones : [];
      p.risks        = Array.isArray(p.risks) ? p.risks : [];
      p.decisions    = Array.isArray(p.decisions) ? p.decisions : [];
      p.changes      = Array.isArray(p.changes) ? p.changes : [];
      p.components   = Array.isArray(p.components) ? p.components : [];
      p.meetings     = Array.isArray(p.meetings) ? p.meetings : [];
      p.openPoints   = Array.isArray(p.openPoints) ? p.openPoints : [];
      p.links        = Array.isArray(p.links) ? p.links : [];
      p.costCenters  = Array.isArray(p.costCenters) ? p.costCenters : [];
      p.archive      = Array.isArray(p.archive) ? p.archive : [];
      // Phase A: project-scoped tags. { id, name, rgb } — used as labels on
      // actions / open points / CRs. Default empty for retrocompat.
      p.tags         = Array.isArray(p.tags) ? p.tags : [];
      p.tags.forEach((t) => {
        t.id = t.id || uid('tg');
        t.name = t.name || 'Tag';
        t.rgb = t.rgb || '148,163,184';
      });
      // Optional: track template lineage so future merges can detect drift
      if (p.templateOf === undefined) p.templateOf = null;

      p.actions.forEach((a) => {
        a.id = a.id || uid('a');
        a.title = a.title || '';
        a.status = a.status || 'todo';
        a.priority = (typeof a.priority === 'number') ? a.priority : 0;
        // priorityLevel — added later; older exports default to 'med' so the
        // existing rank-by-priority sort behaviour stays unchanged. Retrocompat.
        a.priorityLevel = (a.priorityLevel && PRIORITY_LEVELS.some((p) => p.id === a.priorityLevel)) ? a.priorityLevel : 'med';
        a.commitment = (typeof a.commitment === 'number') ? a.commitment : 100;
        a.owner = a.owner || null;
        a.originator = a.originator || null;
        a.due = a.due || null;
        a.startDate = a.startDate || null;
        a.component = a.component || null;
        a.deliverable = a.deliverable || null;
        a.milestone = a.milestone || null;
        a.description = a.description || null;
        a.notes = a.notes || '';
        a.createdAt = a.createdAt || todayISO();
        a.updatedAt = a.updatedAt || a.createdAt || todayISO();
        // originatorDate — when the action was originated. Added in a later
        // schema version; for older exports without this field, default to the
        // existing createdAt so legacy data behaves as if it were already set.
        a.originatorDate = a.originatorDate || a.createdAt || todayISO();
        a.history = Array.isArray(a.history) ? a.history : [];
        // Phase A — dependencies, comments, tags, signed metadata. All defaulted.
        a.dependsOn = Array.isArray(a.dependsOn) ? a.dependsOn : [];
        a.tags      = Array.isArray(a.tags) ? a.tags : [];
        a.comments  = Array.isArray(a.comments) ? a.comments : [];
        a.comments.forEach((c) => {
          c.id  = c.id  || uid('cm');
          c.by  = c.by  || null;
          c.at  = c.at  || todayISO();
          c.text = c.text || '';
        });
        if (a.__lastEditor === undefined) a.__lastEditor = null;
        if (a.__lastEditAt === undefined) a.__lastEditAt = null;
        // a.deletedAt is null for live, ISO string for archived — preserve as-is
      });

      p.deliverables.forEach((d) => {
        d.id = d.id || uid('d');
        d.name = d.name || '';
        d.dueDate = d.dueDate || null;
        d.status = d.status || 'todo';
        // Optional component link, surfaced as a colour stripe on calendar
        // chips so the user can scan-by-component across kinds.
        if (d.component === undefined) d.component = null;
      });

      p.milestones.forEach((m) => {
        m.id = m.id || uid('m');
        m.name = m.name || '';
        m.date = m.date || null;
        // Optional range end-date. When null, the milestone is a single-day
        // event (legacy behaviour); when set, the milestone spans
        // [date … endDate] inclusive in the calendar.
        if (m.endDate === undefined) m.endDate = null;
        if (m.component === undefined) m.component = null;
        m.status = m.status || 'todo';
      });

      p.risks.forEach((r) => {
        r.id = r.id || uid('r');
        r.kind = r.kind || 'risk';
        r.title = r.title || '';
        r.owner = r.owner || null;
        r.mitigation = r.mitigation || '';
        r.actionId = r.actionId || null;
        // ensureRiskShape backfills inherent/residual from legacy probability/impact
        if (!r.inherent) r.inherent = { probability: r.probability || 3, impact: r.impact || 3 };
        if (!r.residual) r.residual = { ...r.inherent };
      });

      p.decisions.forEach((d) => {
        d.id = d.id || uid('dec');
        d.title = d.title || '';
        d.rationale = d.rationale || '';
        d.owner = d.owner || null;
        d.date = d.date || todayISO();
      });

      p.changes.forEach((c) => {
        c.id = c.id || uid('cr');
        c.title = c.title || '';
        c.rationale = c.rationale || '';
        c.analysis = c.analysis || '';
        c.description = c.description || '';
        c.status = c.status || 'proposed';
        c.originator = c.originator || null;
        c.originatedDate = c.originatedDate || todayISO();
        c.decisionBy = c.decisionBy || null;
        c.decisionDate = c.decisionDate || null;
        c.impact = c.impact || {};
        c.impact.schedule = (typeof c.impact.schedule === 'number') ? c.impact.schedule : 0;
        c.impact.cost     = (typeof c.impact.cost     === 'number') ? c.impact.cost     : 0;
        c.impact.scope    = c.impact.scope || '';
        c.impact.risk     = c.impact.risk  || '';
        c.component = c.component || null;
        c.linkUrl = c.linkUrl || null;
        // priorityLevel — added later; default to 'med' for retrocompat
        c.priorityLevel = (c.priorityLevel && PRIORITY_LEVELS.some((p) => p.id === c.priorityLevel)) ? c.priorityLevel : 'med';
        delete c.linkTitle; // dropped from schema
        // Phase A — comments + tags + signed metadata
        c.tags = Array.isArray(c.tags) ? c.tags : [];
        c.comments = Array.isArray(c.comments) ? c.comments : [];
        c.comments.forEach((cm) => {
          cm.id   = cm.id   || uid('cm');
          cm.by   = cm.by   || null;
          cm.at   = cm.at   || todayISO();
          cm.text = cm.text || '';
        });
        if (c.__lastEditor === undefined) c.__lastEditor = null;
        if (c.__lastEditAt === undefined) c.__lastEditAt = null;
      });

      p.components.forEach((cmp) => {
        cmp.id = cmp.id || uid('cm');
        cmp.name = cmp.name || '';
        cmp.color = cmp.color || (COMPONENT_COLORS[0] && COMPONENT_COLORS[0].id) || 'sky';
        cmp.costCenter = cmp.costCenter || null;
      });

      p.meetings.forEach((m) => {
        m.id = m.id || uid('mtg');
        m.kind = m.kind || 'oneoff';
        m.title = m.title || '';
        m.time = m.time || null;
        // Migration: legacy 'weekly' kind → unified 'recurring' with
        // recurUnit/interval. Old data continues to work because all
        // consumers go through expandMeetingDates() which only sees the
        // new shape after normalizeState has run.
        if (m.kind === 'weekly') {
          m.kind = 'recurring';
          m.recurUnit = 'week';
          m.interval = m.interval || 1;
        }
        if (m.kind === 'oneoff') {
          m.date = m.date || todayISO();
        } else if (m.kind === 'recurring') {
          m.recurUnit = (m.recurUnit === 'day' || m.recurUnit === 'week' || m.recurUnit === 'month') ? m.recurUnit : 'week';
          m.interval  = (typeof m.interval === 'number' && m.interval >= 1) ? Math.floor(m.interval) : 1;
          m.startDate = m.startDate || todayISO();
          if (m.recurUnit === 'week') {
            m.dayOfWeek = (typeof m.dayOfWeek === 'number') ? m.dayOfWeek : parseDate(m.startDate).getDay();
          }
        }
        if (m.endDate === undefined) m.endDate = null;
        if (m.component === undefined) m.component = null;
      });

      p.openPoints.forEach((op) => {
        op.id = op.id || uid('op');
        op.title = op.title || '';
        op.notes = op.notes || '';
        op.component = op.component || null;
        op.criticality = op.criticality || 'med';
        op.createdAt = op.createdAt || todayISO();
        // priorityLevel — added later; default to 'med' for retrocompat
        op.priorityLevel = (op.priorityLevel && PRIORITY_LEVELS.some((p) => p.id === op.priorityLevel)) ? op.priorityLevel : 'med';
        // Resolution steps — added in a later schema version; default to []
        // for full retrocompat with older exports that never had this field.
        op.steps = Array.isArray(op.steps) ? op.steps : [];
        op.steps.forEach((s) => {
          s.id = s.id || uid('st');
          s.text = s.text || '';
          s.done = !!s.done;
        });
        // Phase A — comments + tags + signed metadata
        op.tags = Array.isArray(op.tags) ? op.tags : [];
        op.comments = Array.isArray(op.comments) ? op.comments : [];
        op.comments.forEach((cm) => {
          cm.id = cm.id || uid('cm');
          cm.by = cm.by || null;
          cm.at = cm.at || todayISO();
          cm.text = cm.text || '';
        });
        if (op.__lastEditor === undefined) op.__lastEditor = null;
        if (op.__lastEditAt === undefined) op.__lastEditAt = null;
      });

      p.links.forEach((l) => {
        l.id = l.id || uid('lk');
        l.title = l.title || '';
        l.url = l.url || '';
        l.description = l.description || '';
        l.component = l.component || null;
        // folderId — added later for the folder-organized links view. Default
        // to null (= "Loose links" / no folder) for retrocompat.
        if (l.folderId === undefined) l.folderId = null;
      });
      // Link folders — added later. Each: { id, name, collapsed }
      p.linkFolders = Array.isArray(p.linkFolders) ? p.linkFolders : [];
      p.linkFolders.forEach((f) => {
        f.id = f.id || uid('lf');
        f.name = f.name || 'Folder';
        f.collapsed = !!f.collapsed;
      });
      // Drop any link.folderId that no longer points to an existing folder
      const folderIds = new Set(p.linkFolders.map((f) => f.id));
      p.links.forEach((l) => { if (l.folderId && !folderIds.has(l.folderId)) l.folderId = null; });
    });
    return s;
  }

  function exportJSON() {
    flushPendingSaves();
    // Wrap in a metadata envelope so consumers can detect the version. The
    // actual state is preserved at the top level too for backward compat —
    // older importers that just `JSON.parse(file).projects` still work.
    const payload = {
      __schemaVersion: EXPORT_SCHEMA_VERSION,
      __exportedAt: new Date().toISOString(),
      __app: 'cockpit',
      ...state,
    };
    const blob = new Blob([JSON.stringify(payload, null, 2)], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url; a.download = `cockpit-${todayISO()}.json`;
    a.click();
    URL.revokeObjectURL(url);
    toast('Exported');
  }

  // Flush any debounced notes timer + force a saveState immediately. Called by
  // unload / visibility handlers and before export so nothing pending is lost.
  function flushPendingSaves() {
    if (notesSaveTimer) {
      clearTimeout(notesSaveTimer);
      notesSaveTimer = null;
      try { saveNotesNow(); } catch (e) { /* notes panel may not be open */ }
    }
    try { saveState(); } catch (e) { /* quota */ }
  }

  /* --------------------------- Auto-backup ---------------------------- */
  // Persists a FileSystemDirectoryHandle in IndexedDB so the user grants
  // folder access once and the app keeps writing backup JSONs there at the
  // chosen cadence. Browsers without showDirectoryPicker (Safari/Firefox)
  // fall back to a "Download backup now" button only.
  const IDB_NAME = 'cockpit-meta';
  const IDB_STORE = 'kv';
  function idbOpen() {
    return new Promise((resolve, reject) => {
      const req = indexedDB.open(IDB_NAME, 1);
      req.onupgradeneeded = () => req.result.createObjectStore(IDB_STORE);
      req.onsuccess = () => resolve(req.result);
      req.onerror = () => reject(req.error);
    });
  }
  async function idbSet(key, value) {
    const db = await idbOpen();
    return new Promise((resolve, reject) => {
      const tx = db.transaction(IDB_STORE, 'readwrite');
      tx.objectStore(IDB_STORE).put(value, key);
      tx.oncomplete = () => { db.close(); resolve(); };
      tx.onerror = () => { db.close(); reject(tx.error); };
    });
  }
  async function idbGet(key) {
    const db = await idbOpen();
    return new Promise((resolve, reject) => {
      const tx = db.transaction(IDB_STORE, 'readonly');
      const req = tx.objectStore(IDB_STORE).get(key);
      req.onsuccess = () => { db.close(); resolve(req.result); };
      req.onerror = () => { db.close(); reject(req.error); };
    });
  }
  async function idbDelete(key) {
    const db = await idbOpen();
    return new Promise((resolve, reject) => {
      const tx = db.transaction(IDB_STORE, 'readwrite');
      tx.objectStore(IDB_STORE).delete(key);
      tx.oncomplete = () => { db.close(); resolve(); };
      tx.onerror = () => { db.close(); reject(tx.error); };
    });
  }

  let _backupDirHandle = null;       // FileSystemDirectoryHandle once granted
  let _autoBackupTimer = null;       // setInterval id
  const supportsDirPicker = typeof window.showDirectoryPicker === 'function';

  async function loadBackupDirHandle() {
    if (!supportsDirPicker) return null;
    try {
      const h = await idbGet('autoBackupDir');
      if (!h) return null;
      // Verify we still have permission without prompting
      const perm = await h.queryPermission?.({ mode: 'readwrite' });
      if (perm === 'granted') {
        _backupDirHandle = h;
        return h;
      }
      // Need user gesture to re-prompt; leave for user to re-pick
      return null;
    } catch (e) { return null; }
  }

  async function requestBackupDir() {
    if (!supportsDirPicker) {
      toast('Folder backup requires Chrome / Edge — use Export instead.');
      return null;
    }
    try {
      const h = await window.showDirectoryPicker({ mode: 'readwrite', id: 'cockpit-backup' });
      const perm = await h.requestPermission({ mode: 'readwrite' });
      if (perm !== 'granted') { toast('Permission denied'); return null; }
      _backupDirHandle = h;
      await idbSet('autoBackupDir', h);
      state.settings.autoBackup.dirName = h.name || 'Selected folder';
      saveState();
      return h;
    } catch (e) {
      // User cancelled the picker
      return null;
    }
  }

  async function writeBackupNow() {
    if (!_backupDirHandle) {
      toast('Pick a backup folder first');
      return false;
    }
    try {
      // Ensure permission still granted (some browsers revoke on idle)
      const perm = await _backupDirHandle.queryPermission?.({ mode: 'readwrite' });
      if (perm !== 'granted') {
        const next = await _backupDirHandle.requestPermission?.({ mode: 'readwrite' });
        if (next !== 'granted') { toast('Folder permission expired — pick again'); return false; }
      }
      flushPendingSaves();
      const stamp = new Date().toISOString().replace(/[:.]/g, '-').slice(0, 19);
      const fname = `cockpit-${stamp}.json`;
      const payload = {
        __schemaVersion: EXPORT_SCHEMA_VERSION,
        __exportedAt: new Date().toISOString(),
        __app: 'cockpit',
        ...state,
      };
      const fh = await _backupDirHandle.getFileHandle(fname, { create: true });
      const w  = await fh.createWritable();
      await w.write(JSON.stringify(payload, null, 2));
      await w.close();
      state.settings.autoBackup.lastBackupAt = new Date().toISOString();
      saveState();
      refreshAutoBackupUI();
      return true;
    } catch (e) {
      toast('Backup failed: ' + (e?.message || e));
      return false;
    }
  }

  function scheduleAutoBackup() {
    // Always start clean
    if (_autoBackupTimer) { clearInterval(_autoBackupTimer); _autoBackupTimer = null; }
    const ab = state.settings.autoBackup;
    if (!ab.enabled || !_backupDirHandle) return;
    const ms = ab.intervalMinutes * 60 * 1000;
    _autoBackupTimer = setInterval(() => { writeBackupNow(); }, ms);
  }

  function refreshAutoBackupUI() {
    const panel = $('#autoBackupPanel');
    if (!panel || panel.hidden) return;
    renderAutoBackupPanel();
  }

  function renderAutoBackupPanel() {
    const panel = $('#autoBackupPanel');
    if (!panel) return;
    const ab = state.settings.autoBackup;
    const supported = supportsDirPicker;
    const hasHandle = !!_backupDirHandle;
    const lastTxt = ab.lastBackupAt ? new Date(ab.lastBackupAt).toLocaleString() : 'never';
    panel.innerHTML = `
      <div class="ab-head">
        <span class="ab-title">Auto-backup</span>
        <button class="icon-btn" id="abClose" title="Close" aria-label="Close">×</button>
      </div>
      <div class="ab-body">
        ${supported ? '' : '<div class="ab-warn">Folder backup requires Chrome / Edge. The "Download now" button below works in any browser.</div>'}
        <div class="ab-row">
          <span class="ab-lbl">Folder</span>
          <span class="ab-val">${ab.dirName ? escapeHTML(ab.dirName) + (hasHandle ? '' : ' <span class="ab-muted">(reconnect needed)</span>') : '<span class="ab-muted">— none —</span>'}</span>
          ${supported ? '<button class="ghost" id="abPick">Choose folder…</button>' : ''}
        </div>
        <div class="ab-row">
          <span class="ab-lbl">Frequency</span>
          <select id="abInterval">
            <option value="5"    ${ab.intervalMinutes === 5    ? 'selected' : ''}>Every 5 minutes</option>
            <option value="15"   ${ab.intervalMinutes === 15   ? 'selected' : ''}>Every 15 minutes</option>
            <option value="30"   ${ab.intervalMinutes === 30   ? 'selected' : ''}>Every 30 minutes</option>
            <option value="60"   ${ab.intervalMinutes === 60   ? 'selected' : ''}>Every hour</option>
            <option value="360"  ${ab.intervalMinutes === 360  ? 'selected' : ''}>Every 6 hours</option>
            <option value="1440" ${ab.intervalMinutes === 1440 ? 'selected' : ''}>Once a day</option>
          </select>
        </div>
        <div class="ab-row">
          <label class="ab-toggle">
            <input type="checkbox" id="abEnabled" ${ab.enabled && hasHandle ? 'checked' : ''} ${(!supported || !hasHandle) ? 'disabled' : ''} />
            <span>Enable auto-backup</span>
          </label>
        </div>
        <div class="ab-row ab-status">
          <span class="ab-lbl">Last backup</span>
          <span class="ab-val">${lastTxt}</span>
        </div>
        <div class="ab-actions">
          <button class="ghost" id="abDownload" title="Download a backup JSON to your default Downloads folder">Download now</button>
          ${supported ? `<button class="primary" id="abBackupNow" ${hasHandle ? '' : 'disabled'}>Back up to folder now</button>` : ''}
        </div>
      </div>`;

    // Wire
    panel.querySelector('#abClose')?.addEventListener('click', closeAutoBackupPanel);
    panel.querySelector('#abPick')?.addEventListener('click', async () => {
      await requestBackupDir();
      // After picking, refresh UI; auto-backup stays as the user set it
      renderAutoBackupPanel();
      scheduleAutoBackup();
    });
    panel.querySelector('#abInterval')?.addEventListener('change', (e) => {
      const v = parseInt(e.target.value, 10);
      state.settings.autoBackup.intervalMinutes = v;
      saveState();
      scheduleAutoBackup();
    });
    panel.querySelector('#abEnabled')?.addEventListener('change', (e) => {
      state.settings.autoBackup.enabled = !!e.target.checked;
      saveState();
      scheduleAutoBackup();
      if (e.target.checked) toast('Auto-backup enabled');
      else toast('Auto-backup disabled');
    });
    panel.querySelector('#abBackupNow')?.addEventListener('click', async () => {
      const ok = await writeBackupNow();
      if (ok) toast('Backed up');
    });
    panel.querySelector('#abDownload')?.addEventListener('click', () => {
      exportJSON();
    });
  }

  function openAutoBackupPanel() {
    const panel = $('#autoBackupPanel');
    if (!panel) return;
    panel.hidden = false;
    renderAutoBackupPanel();
  }
  function closeAutoBackupPanel() {
    const panel = $('#autoBackupPanel');
    if (panel) panel.hidden = true;
  }
  function wireAutoBackup() {
    const btn = $('#btnAutoBackup');
    const panel = $('#autoBackupPanel');
    if (!btn || !panel) return;
    btn.addEventListener('click', () => {
      panel.hidden ? openAutoBackupPanel() : closeAutoBackupPanel();
    });
    // Click outside the panel and the button → close
    document.addEventListener('mousedown', (e) => {
      if (panel.hidden) return;
      if (e.target.closest('#autoBackupPanel') || e.target.closest('#btnAutoBackup')) return;
      closeAutoBackupPanel();
    });
  }
  async function initAutoBackup() {
    if (!supportsDirPicker) return;
    await loadBackupDirHandle();
    scheduleAutoBackup();
  }
  function importJSON() {
    const input = document.createElement('input');
    input.type = 'file'; input.accept = 'application/json';
    input.addEventListener('change', () => {
      const file = input.files?.[0];
      if (!file) return;
      const reader = new FileReader();
      reader.onload = () => {
        try {
          const obj = JSON.parse(reader.result);
          if (!obj || typeof obj !== 'object') throw new Error('Invalid JSON');
          if (!obj.projects || !obj.people) throw new Error('Missing projects/people');
          // Strip metadata fields if present (envelope) — they're not state
          delete obj.__schemaVersion;
          delete obj.__exportedAt;
          delete obj.__app;
          undoStack.push(JSON.stringify(state));
          state = normalizeState(obj);
          saveState();
          render();
          toast('Imported');
        } catch (e) { toast('Import failed: ' + e.message); }
      };
      reader.readAsText(file);
    });
    input.click();
  }

  /* ----------------------- Phase K: light merge ------------------------ */
  // 2-way merge between local state ("mine") and a teammate's JSON
  // ("theirs"). The diff is computed per record-kind by id; conflicts
  // (same id, differing payload) get a per-row resolution UI. Records
  // that exist on only one side default to 'include' so nothing is lost
  // unless the user explicitly drops it.
  function openMergeOverlay() {
    const input = document.createElement('input');
    input.type = 'file'; input.accept = 'application/json';
    input.addEventListener('change', () => {
      const file = input.files?.[0];
      if (!file) return;
      const reader = new FileReader();
      reader.onload = () => {
        try {
          const obj = JSON.parse(reader.result);
          if (!obj || !obj.projects || !obj.people) throw new Error('Not a Cockpit JSON');
          delete obj.__schemaVersion; delete obj.__exportedAt; delete obj.__app;
          const theirs = normalizeState(JSON.parse(JSON.stringify(obj)));
          showMergeUI(state, theirs);
        } catch (e) { toast('Merge failed: ' + e.message); }
      };
      reader.readAsText(file);
    });
    input.click();
  }
  function recordKey(rec) { return rec.id || ''; }
  function recordSig(rec) {
    // Signature for change detection: stringify everything except edit
    // metadata (so a re-stamp without value change doesn't count as a diff).
    const { __lastEditor, __lastEditAt, ...rest } = rec || {};
    return JSON.stringify(rest);
  }
  function diffArr(mineArr, theirsArr) {
    const mineMap = new Map((mineArr || []).map((r) => [recordKey(r), r]));
    const theirsMap = new Map((theirsArr || []).map((r) => [recordKey(r), r]));
    const onlyMine = [], onlyTheirs = [], conflicts = [], same = [];
    mineMap.forEach((m, id) => {
      if (!theirsMap.has(id)) { onlyMine.push(m); return; }
      const t = theirsMap.get(id);
      if (recordSig(m) === recordSig(t)) same.push(m);
      else conflicts.push({ id, mine: m, theirs: t });
    });
    theirsMap.forEach((t, id) => { if (!mineMap.has(id)) onlyTheirs.push(t); });
    return { onlyMine, onlyTheirs, conflicts, same };
  }
  function showMergeUI(mine, theirs) {
    // Build per-project diffs across actions / openPoints / changes /
    // risks / decisions / deliverables / milestones / components / links.
    const projDiffs = [];
    const mineProj = new Map(mine.projects.map((p) => [p.id, p]));
    const theirsProj = new Map(theirs.projects.map((p) => [p.id, p]));
    const projIds = new Set([...mineProj.keys(), ...theirsProj.keys()]);
    projIds.forEach((pid) => {
      const m = mineProj.get(pid);
      const t = theirsProj.get(pid);
      if (!m && t) {
        projDiffs.push({ pid, name: t.name, only: 'theirs', proj: t });
        return;
      }
      if (m && !t) {
        projDiffs.push({ pid, name: m.name, only: 'mine', proj: m });
        return;
      }
      const arrays = ['actions', 'openPoints', 'changes', 'risks', 'decisions', 'deliverables', 'milestones', 'components', 'links', 'meetings'];
      const perKind = {};
      arrays.forEach((k) => { perKind[k] = diffArr(m[k] || [], t[k] || []); });
      projDiffs.push({ pid, name: m.name, perKind });
    });
    // Resolution state: choices per record-key
    const choices = {};
    // Default: include only-theirs (auto-add new from teammate), keep
    // only-mine (don't drop), conflicts default to 'mine'.
    projDiffs.forEach((pd) => {
      if (!pd.perKind) return;
      Object.entries(pd.perKind).forEach(([kind, d]) => {
        d.onlyTheirs.forEach((r) => { choices[`${pd.pid}:${kind}:${r.id}`] = 'include'; });
        d.onlyMine.forEach((r)   => { choices[`${pd.pid}:${kind}:${r.id}`] = 'keep'; });
        d.conflicts.forEach((c)  => {
          // Tiebreaker: take whichever has the most-recent __lastEditAt
          const mAt = c.mine.__lastEditAt || c.mine.updatedAt || '';
          const tAt = c.theirs.__lastEditAt || c.theirs.updatedAt || '';
          choices[`${pd.pid}:${kind}:${c.id}`] = (tAt > mAt) ? 'theirs' : 'mine';
        });
      });
    });

    const overlay = document.createElement('div');
    overlay.className = 'overlay merge-overlay';
    overlay.innerHTML = `
      <div class="merge-modal">
        <div class="merge-head">
          <div class="merge-title">Merge teammate's JSON</div>
          <button class="icon-btn" id="mClose" title="Close">×</button>
        </div>
        <div class="merge-body" id="mBody"></div>
        <div class="merge-foot">
          <span class="merge-summary" id="mSummary"></span>
          <button class="ghost" id="mCancel">Cancel</button>
          <button class="primary" id="mApply">Apply merge</button>
        </div>
      </div>`;
    document.body.appendChild(overlay);
    const close = () => overlay.remove();
    $('#mClose', overlay).addEventListener('click', close);
    $('#mCancel', overlay).addEventListener('click', close);

    function renderMergeBody() {
      const body = $('#mBody', overlay);
      let totalConflicts = 0, totalAdds = 0;
      body.innerHTML = projDiffs.map((pd) => {
        if (pd.only === 'theirs') {
          totalAdds++;
          return `<div class="merge-section">
            <div class="merge-section-title">${escapeHTML(pd.name)} <span class="merge-tag merge-tag-add">new project</span></div>
            <div class="merge-row">
              <label class="merge-choice">
                <input type="radio" name="proj-${escapeHTML(pd.pid)}" value="include" checked />
                <span>Include teammate's project</span>
              </label>
              <label class="merge-choice">
                <input type="radio" name="proj-${escapeHTML(pd.pid)}" value="skip" />
                <span>Skip</span>
              </label>
            </div>
          </div>`;
        }
        if (pd.only === 'mine') {
          return `<div class="merge-section">
            <div class="merge-section-title">${escapeHTML(pd.name)} <span class="merge-tag merge-tag-keep">only in yours</span></div>
            <div class="merge-row muted">Project will remain as-is.</div>
          </div>`;
        }
        const blocks = Object.entries(pd.perKind).map(([kind, d]) => {
          const adds = d.onlyTheirs.length;
          const conflicts = d.conflicts.length;
          const onlyMineN = d.onlyMine.length;
          totalAdds += adds;
          totalConflicts += conflicts;
          if (!adds && !conflicts && !onlyMineN) return '';
          const summary = [
            adds      ? `<span class="merge-tag merge-tag-add">+${adds} new</span>`     : '',
            conflicts ? `<span class="merge-tag merge-tag-conflict">${conflicts} differ</span>` : '',
            onlyMineN ? `<span class="merge-tag merge-tag-keep">${onlyMineN} only-yours</span>` : '',
          ].join(' ');
          const conflictRows = d.conflicts.map((c) => {
            const key = `${pd.pid}:${kind}:${c.id}`;
            const cur = choices[key] || 'mine';
            const label = c.mine.title || c.mine.name || c.theirs.title || c.theirs.name || c.id;
            const mAt = c.mine.__lastEditAt || c.mine.updatedAt || '—';
            const tAt = c.theirs.__lastEditAt || c.theirs.updatedAt || '—';
            return `<div class="merge-conflict" data-key="${escapeHTML(key)}">
              <div class="merge-conflict-title">${escapeHTML(label)}</div>
              <div class="merge-conflict-versions">
                <label class="merge-version ${cur === 'mine' ? 'sel' : ''}">
                  <input type="radio" name="${escapeHTML(key)}" value="mine" ${cur === 'mine' ? 'checked' : ''} />
                  <span class="merge-version-label">Yours <span class="merge-version-when">${escapeHTML(mAt)}</span></span>
                </label>
                <label class="merge-version ${cur === 'theirs' ? 'sel' : ''}">
                  <input type="radio" name="${escapeHTML(key)}" value="theirs" ${cur === 'theirs' ? 'checked' : ''} />
                  <span class="merge-version-label">Theirs <span class="merge-version-when">${escapeHTML(tAt)}</span></span>
                </label>
                <label class="merge-version ${cur === 'both' ? 'sel' : ''}">
                  <input type="radio" name="${escapeHTML(key)}" value="both" ${cur === 'both' ? 'checked' : ''} />
                  <span class="merge-version-label">Keep both</span>
                </label>
              </div>
            </div>`;
          }).join('');
          return `<div class="merge-block">
            <div class="merge-block-head">${kind} ${summary}</div>
            ${conflicts ? `<div class="merge-conflicts">${conflictRows}</div>` : ''}
          </div>`;
        }).filter(Boolean).join('');
        return `<div class="merge-section">
          <div class="merge-section-title">${escapeHTML(pd.name)}</div>
          ${blocks || '<div class="merge-row muted">No differences.</div>'}
        </div>`;
      }).join('') || '<div class="empty">No differences detected.</div>';

      $('#mSummary', overlay).textContent =
        `${totalAdds} additions · ${totalConflicts} conflicts`;

      // Wire change events for conflict + project radios
      body.querySelectorAll('input[type=radio]').forEach((rd) => {
        rd.addEventListener('change', () => {
          const name = rd.name;
          if (name.startsWith('proj-')) {
            const pid = name.slice(5);
            choices[`__proj__:${pid}`] = rd.value;
          } else {
            choices[name] = rd.value;
          }
        });
      });
    }
    renderMergeBody();

    $('#mApply', overlay).addEventListener('click', () => {
      // Apply choices to mine
      undoStack.push(JSON.stringify(state));
      const result = JSON.parse(JSON.stringify(mine));
      const resultProjMap = new Map(result.projects.map((p, i) => [p.id, i]));
      projDiffs.forEach((pd) => {
        if (pd.only === 'theirs') {
          const choice = choices[`__proj__:${pd.pid}`] || 'include';
          if (choice === 'include') {
            result.projects.push(JSON.parse(JSON.stringify(pd.proj)));
          }
          return;
        }
        if (pd.only === 'mine') return;
        const targetIdx = resultProjMap.get(pd.pid);
        if (targetIdx == null) return;
        const target = result.projects[targetIdx];
        Object.entries(pd.perKind).forEach(([kind, d]) => {
          // 1. Apply only-theirs additions for which the user kept default 'include'
          d.onlyTheirs.forEach((r) => {
            const key = `${pd.pid}:${kind}:${r.id}`;
            if ((choices[key] || 'include') !== 'include') return;
            target[kind] = target[kind] || [];
            target[kind].push(JSON.parse(JSON.stringify(r)));
          });
          // 2. Conflicts
          d.conflicts.forEach((c) => {
            const key = `${pd.pid}:${kind}:${c.id}`;
            const choice = choices[key] || 'mine';
            const list = target[kind] || (target[kind] = []);
            const idx = list.findIndex((x) => x.id === c.id);
            if (choice === 'theirs') {
              if (idx >= 0) list[idx] = JSON.parse(JSON.stringify(c.theirs));
            } else if (choice === 'both') {
              const dup = JSON.parse(JSON.stringify(c.theirs));
              const prefix = (c.theirs.id || 'id').split('_')[0] || 'id';
              dup.id = uid(prefix);
              if (dup.title) dup.title = dup.title + ' (theirs)';
              else if (dup.name) dup.name = dup.name + ' (theirs)';
              list.push(dup);
            }
            // 'mine' = no-op
          });
        });
      });
      state = normalizeState(result);
      saveState();
      render();
      close();
      toast('Merged');
    });
  }

  /* ------------------- Personal to-do (floating widget) ------------------ */

  // The widget lives at the state level (project-independent), so changes only
  // need a saveState() — no commit() (we don't want personal todos showing up
  // in the global undo stack, and we want quick local-only redraws).
  function refreshTodoFab() {
    const fab = $('#todoFab');
    const badge = $('#todoFabBadge');
    if (!fab) return;
    const todos = state.todos || [];
    const open = todos.filter((t) => !t.done).length;
    const total = todos.length;
    fab.classList.toggle('empty',  total === 0);
    fab.classList.toggle('all-done', total > 0 && open === 0);
    fab.classList.toggle('has-open', open > 0);
    if (badge) {
      if (open > 0) { badge.textContent = String(open); badge.hidden = false; }
      else { badge.hidden = true; }
    }
    fab.title = total === 0
      ? 'My to-do — empty'
      : (open === 0 ? `My to-do — ${total} done` : `My to-do — ${open} open of ${total}`);
  }

  function renderTodoList() {
    const list = $('#todoList');
    const meta = $('#todoMeta');
    if (!list) return;
    const todos = state.todos || [];
    const open = todos.filter((t) => !t.done).length;
    if (meta) meta.textContent = todos.length ? `${open}/${todos.length}` : '';
    if (!todos.length) {
      list.innerHTML = '<div class="todo-empty">No personal to-dos yet — add one above.</div>';
      return;
    }
    list.innerHTML = todos.map((t) => `
      <div class="todo-item ${t.done ? 'done' : ''}" data-todo-id="${t.id}">
        <span class="todo-grip" title="Drag to reorder" aria-hidden="true">⋮⋮</span>
        <button type="button" class="todo-check ${t.done ? 'on' : ''}" aria-label="${t.done ? 'Mark not done' : 'Mark done'}"></button>
        <span class="todo-text" contenteditable="true" data-placeholder="To-do…">${escapeHTML(t.text)}</span>
        <button type="button" class="todo-del" aria-label="Delete" title="Delete">×</button>
      </div>`).join('');

    // Wire each row
    list.querySelectorAll('.todo-item').forEach((row) => {
      const id = row.dataset.todoId;
      row.querySelector('.todo-check').addEventListener('click', () => {
        const t = (state.todos || []).find((x) => x.id === id);
        if (!t) return;
        t.done = !t.done;
        saveState();
        renderTodoList(); refreshTodoFab();
      });
      const text = row.querySelector('.todo-text');
      text.addEventListener('keydown', (e) => {
        if (e.key === 'Enter' && !e.shiftKey) { e.preventDefault(); text.blur(); }
        else if (e.key === 'Escape') { e.preventDefault(); text.blur(); }
        else if (e.key === 'Backspace' && text.textContent === '') {
          e.preventDefault();
          state.todos = (state.todos || []).filter((x) => x.id !== id);
          saveState(); renderTodoList(); refreshTodoFab();
        }
      });
      text.addEventListener('blur', () => {
        const t = (state.todos || []).find((x) => x.id === id);
        if (!t) return;
        const v = text.textContent.trim();
        if (t.text !== v) { t.text = v; saveState(); refreshTodoFab(); }
      });
      row.querySelector('.todo-del').addEventListener('click', (e) => {
        e.stopPropagation();
        state.todos = (state.todos || []).filter((x) => x.id !== id);
        saveState(); renderTodoList(); refreshTodoFab();
      });
      // Drag-to-reorder via the grip — same custom-mouse pattern as op-steps
      row.querySelector('.todo-grip').addEventListener('mousedown', (e) => {
        e.preventDefault();
        const listEl = $('#todoList');
        row.classList.add('dragging');
        document.body.classList.add('is-todo-dragging');
        const onMove = (em) => {
          const sibs = [...listEl.querySelectorAll('.todo-item:not(.dragging)')];
          const after = sibs.find((sib) => {
            const r = sib.getBoundingClientRect();
            return em.clientY < r.top + r.height / 2;
          });
          if (after) listEl.insertBefore(row, after);
          else listEl.appendChild(row);
        };
        const onUp = () => {
          row.classList.remove('dragging');
          document.body.classList.remove('is-todo-dragging');
          window.removeEventListener('mousemove', onMove);
          window.removeEventListener('mouseup', onUp);
          const newOrder = [...listEl.querySelectorAll('.todo-item')].map((r) => r.dataset.todoId);
          const before = (state.todos || []).map((t) => t.id).join(',');
          state.todos = newOrder.map((tid) => (state.todos || []).find((x) => x.id === tid)).filter(Boolean);
          if (state.todos.map((t) => t.id).join(',') !== before) { saveState(); }
        };
        window.addEventListener('mousemove', onMove);
        window.addEventListener('mouseup', onUp);
      });
    });
  }

  function wireTodoWidget() {
    const fab = $('#todoFab');
    const panel = $('#todoPanel');
    const input = $('#todoInput');
    if (!fab || !panel) return;
    const open = () => {
      panel.hidden = false;
      renderTodoList();
      setTimeout(() => input?.focus(), 30);
    };
    const close = () => { panel.hidden = true; };
    const toggle = () => panel.hidden ? open() : close();
    fab.addEventListener('click', toggle);
    $('#todoClose').addEventListener('click', close);
    // Add via Enter in the input or click of +
    const addItem = () => {
      const v = (input.value || '').trim();
      if (!v) return;
      state.todos = state.todos || [];
      state.todos.push({ id: uid('td'), text: v, done: false });
      input.value = '';
      saveState();
      renderTodoList(); refreshTodoFab();
    };
    $('#todoAdd').addEventListener('click', addItem);
    input.addEventListener('keydown', (e) => {
      if (e.key === 'Enter') { e.preventDefault(); addItem(); }
      else if (e.key === 'Escape') { e.preventDefault(); close(); }
    });
    // Click outside to close — but ignore clicks inside the panel or the FAB
    document.addEventListener('mousedown', (e) => {
      if (panel.hidden) return;
      if (e.target.closest('#todoPanel') || e.target.closest('#todoFab')) return;
      close();
    });
    refreshTodoFab();
  }

  /* ----------------------- Phase D: Inbox + bell ---------------------- */
  // Aggregate items that need user attention. Each item has a stable id so
  // the user can dismiss it and the dismissal sticks across reloads.
  const STALE_DAYS = 14;
  const SOON_DAYS = 3;
  function inboxItems() {
    const today = todayISO();
    const dismissed = new Set(state.inbox?.dismissed || []);
    const out = [];
    state.projects.forEach((proj) => {
      // Late actions
      (proj.actions || []).forEach((a) => {
        if (a.deletedAt) return;
        if (isClosedStatus(a.status)) return;
        if (a.due && dayDiff(a.due, today) < 0) {
          out.push({
            id: 'late-action:' + a.id,
            kind: 'late', icon: '⏰', tone: 'bad',
            title: a.title,
            sub: `${proj.name} · ${personName(a.owner)} · ${Math.abs(dayDiff(a.due, today))}d late`,
            run: () => { state.currentProjectId = proj.id; state.currentView = 'board'; render(); setTimeout(() => openDrawer(a.id), 30); },
          });
        } else if (a.due && dayDiff(a.due, today) <= SOON_DAYS) {
          out.push({
            id: 'due-soon:' + a.id,
            kind: 'soon', icon: '⏳', tone: 'warn',
            title: a.title,
            sub: `${proj.name} · ${personName(a.owner)} · due in ${dayDiff(a.due, today)}d`,
            run: () => { state.currentProjectId = proj.id; state.currentView = 'board'; render(); setTimeout(() => openDrawer(a.id), 30); },
          });
        }
        // Stale (open + not updated in N days)
        if (!isClosedStatus(a.status) && a.updatedAt && dayDiff(today, a.updatedAt) >= STALE_DAYS) {
          out.push({
            id: 'stale-action:' + a.id,
            kind: 'stale', icon: '·', tone: 'muted',
            title: a.title,
            sub: `${proj.name} · ${personName(a.owner)} · untouched ${dayDiff(today, a.updatedAt)}d`,
            run: () => { state.currentProjectId = proj.id; state.currentView = 'board'; render(); setTimeout(() => openDrawer(a.id), 30); },
          });
        }
      });
      // Pending CRs > 14d
      (proj.changes || []).forEach((c) => {
        if (c.status !== 'proposed' && c.status !== 'under_review') return;
        if (c.originatedDate && dayDiff(today, c.originatedDate) >= STALE_DAYS) {
          out.push({
            id: 'cr-pending:' + c.id,
            kind: 'cr', icon: '⇆', tone: 'warn',
            title: c.title,
            sub: `${proj.name} · ${c.status.replace('_', ' ')} · ${dayDiff(today, c.originatedDate)}d ago`,
            run: () => { state.currentProjectId = proj.id; state.currentView = 'changes'; render(); setTimeout(() => openChangeRequestEditor(c.id), 30); },
          });
        }
      });
      // High-criticality unmitigated risks
      (proj.risks || []).forEach((r) => {
        if ((r.kind || 'risk') === 'opportunity') return;
        const inh = r.inherent || { probability: 0, impact: 0 };
        const res = r.residual || inh;
        const score = (res.probability || 0) * (res.impact || 0);
        if (score >= 12 && !r.actionId) {
          out.push({
            id: 'risk-unmit:' + r.id,
            kind: 'risk', icon: '△', tone: 'bad',
            title: r.title,
            sub: `${proj.name} · residual ${score} · no linked action`,
            run: () => { state.currentProjectId = proj.id; state.currentView = 'risks'; render(); setTimeout(() => openRiskEditor(r.id), 30); },
          });
        }
      });
    });
    // Personal todos due / open
    (state.todos || []).forEach((t) => {
      if (t.done) return;
      out.push({
        id: 'todo:' + t.id,
        kind: 'todo', icon: '✓', tone: 'muted',
        title: t.text || '(untitled)',
        sub: 'Personal to-do',
        run: () => $('#todoFab')?.click(),
      });
    });
    return out.filter((it) => !dismissed.has(it.id));
  }
  function inboxCount() { return inboxItems().length; }
  function refreshBell() {
    const bell = $('#bellBadge');
    if (!bell) return;
    const n = inboxCount();
    if (n > 0) { bell.textContent = String(n); bell.hidden = false; }
    else bell.hidden = true;
  }
  function dismissInbox(id) {
    state.inbox = state.inbox || { dismissed: [] };
    state.inbox.dismissed = state.inbox.dismissed || [];
    if (!state.inbox.dismissed.includes(id)) state.inbox.dismissed.push(id);
    saveState();
  }
  function renderInbox(root) {
    const view = document.createElement('div');
    view.className = 'view';
    const items = inboxItems();
    const grouped = { late: [], soon: [], cr: [], risk: [], stale: [], todo: [] };
    items.forEach((it) => { (grouped[it.kind] || (grouped[it.kind] = [])).push(it); });
    const sectionTitle = { late: 'Late', soon: 'Due soon', cr: 'CR aging', risk: 'Unmitigated risks', stale: 'Stale', todo: 'Personal to-do' };
    const order = ['late', 'soon', 'cr', 'risk', 'stale', 'todo'];
    view.innerHTML = `
      <div class="page-head">
        <div>
          <div class="page-title">Inbox</div>
          <div class="page-sub">${items.length ? `${items.length} item${items.length === 1 ? '' : 's'} need your attention` : 'Nothing requires action right now — nice work.'}</div>
        </div>
        <div class="page-actions">
          ${state.settings.notifyEnabled ? '' : '<button class="ghost" id="btnEnableNotify" title="Browser notifications when new items arrive">Enable notifications</button>'}
          ${items.length ? '<button class="ghost" id="btnDismissAll" title="Dismiss all">Dismiss all</button>' : ''}
        </div>
      </div>
      <div class="inbox">
        ${items.length === 0 ? '<div class="empty">Your inbox is clear.</div>' : ''}
        ${order.filter((k) => (grouped[k] || []).length).map((k) => `
          <div class="inbox-section">
            <div class="inbox-section-title">${escapeHTML(sectionTitle[k] || k)}<span class="inbox-section-count">${grouped[k].length}</span></div>
            <div class="inbox-list">
              ${grouped[k].map((it) => `
                <div class="inbox-item tone-${it.tone}" data-id="${escapeHTML(it.id)}">
                  <span class="inbox-icon">${escapeHTML(it.icon)}</span>
                  <div class="inbox-text">
                    <div class="inbox-title">${escapeHTML(it.title)}</div>
                    <div class="inbox-sub">${escapeHTML(it.sub)}</div>
                  </div>
                  <button class="inbox-dismiss" data-action="dismiss" title="Dismiss" aria-label="Dismiss">×</button>
                </div>
              `).join('')}
            </div>
          </div>
        `).join('')}
      </div>`;
    root.appendChild(view);

    view.querySelectorAll('.inbox-item').forEach((row) => {
      row.addEventListener('click', (e) => {
        if (e.target.closest('[data-action="dismiss"]')) return;
        const id = row.dataset.id;
        const it = inboxItems().find((x) => x.id === id);
        if (it) it.run();
      });
      row.querySelector('[data-action="dismiss"]')?.addEventListener('click', (e) => {
        e.stopPropagation();
        const id = row.dataset.id;
        dismissInbox(id);
        render();
      });
    });
    $('#btnDismissAll')?.addEventListener('click', () => {
      if (!confirm('Dismiss all inbox items? They\'ll re-appear if their condition still holds tomorrow.')) return;
      items.forEach((it) => dismissInbox(it.id));
      render();
    });
    $('#btnEnableNotify')?.addEventListener('click', async () => {
      try {
        const perm = await Notification.requestPermission();
        if (perm === 'granted') {
          state.settings.notifyEnabled = true;
          saveState(); render();
          toast('Notifications enabled');
        } else {
          toast('Notifications denied');
        }
      } catch (e) { toast('Notifications not supported'); }
    });
  }
  function maybeNotifyInbox() {
    if (!state.settings.notifyEnabled) return;
    if (typeof Notification === 'undefined' || Notification.permission !== 'granted') return;
    const today = todayISO();
    if (state.inbox?.lastNotifyDate === today) return;
    const items = inboxItems();
    if (!items.length) return;
    state.inbox = state.inbox || {};
    state.inbox.lastNotifyDate = today;
    saveState();
    try {
      new Notification('Cockpit — ' + items.length + ' item' + (items.length === 1 ? '' : 's') + ' need your attention', {
        body: items.slice(0, 3).map((it) => '• ' + it.title).join('\n'),
        silent: false,
      });
    } catch (e) { /* ignore */ }
  }

  /* ----------------------- Phase E: Status Report ---------------------- */
  // Compose a period-bounded status report from existing data, render as
  // a live HTML preview, then export as Markdown / Markdown file / Print.
  // Period state lives outside the rendered view so re-renders preserve it.
  const reportState = {
    period: '7d',          // '7d' | '30d' | '90d' | 'custom'
    customSince: '',
    customUntil: '',
  };
  function reportPeriodRange() {
    const today = todayISO();
    if (reportState.period === 'custom' && reportState.customSince && reportState.customUntil) {
      return { since: reportState.customSince, until: reportState.customUntil };
    }
    const days = reportState.period === '30d' ? 30 : reportState.period === '90d' ? 90 : 7;
    const since = fmtISO(new Date(Date.now() - days * dayMs));
    return { since, until: today };
  }
  function buildReportData(proj, since, until) {
    const acts = (proj.actions || []).filter((a) => !a.deletedAt);
    const today = todayISO();
    const horizon = fmtISO(new Date(parseDate(today).getTime() + 14 * dayMs));
    const inRange = (d) => d && d >= since && d <= until;

    const changed = acts.filter((a) => a.updatedAt && inRange(a.updatedAt))
      .sort((a, b) => (b.updatedAt || '').localeCompare(a.updatedAt || ''));
    const lateOrBlocked = acts.filter((a) =>
      !isClosedStatus(a.status) && ((a.due && dayDiff(a.due, today) < 0) || a.status === 'blocked'));
    const doneInPeriod = acts.filter((a) => a.status === 'done' && inRange(a.updatedAt));
    const decisions = (proj.decisions || []).filter((d) => inRange(d.date))
      .sort((a, b) => (b.date || '').localeCompare(a.date || ''));
    const crsDecided = (proj.changes || []).filter((c) =>
      ['approved', 'rejected', 'implemented', 'cancelled'].includes(c.status) && inRange(c.decisionDate))
      .sort((a, b) => (b.decisionDate || '').localeCompare(a.decisionDate || ''));
    const topRisks = (proj.risks || [])
      .filter((r) => (r.kind || 'risk') === 'risk')
      .map((r) => {
        const res = r.residual || r.inherent || { probability: 0, impact: 0 };
        return { ...r, _score: (res.probability || 0) * (res.impact || 0) };
      })
      .filter((r) => r._score > 0)
      .sort((a, b) => b._score - a._score)
      .slice(0, 5);
    const upcomingMilestones = (proj.milestones || []).filter((m) =>
      !m.done && m.date && m.date >= today && m.date <= horizon)
      .sort((a, b) => (a.date || '').localeCompare(b.date || ''));
    const upcomingDeliverables = (proj.deliverables || []).filter((d) =>
      !d.done && d.date && d.date >= today && d.date <= horizon)
      .sort((a, b) => (a.date || '').localeCompare(b.date || ''));
    const upcomingDue = acts.filter((a) =>
      !isClosedStatus(a.status) && a.due && a.due >= today && a.due <= horizon)
      .sort((a, b) => (a.due || '').localeCompare(b.due || ''));

    const k = {
      done: doneInPeriod.length,
      changed: changed.length,
      late: lateOrBlocked.filter((a) => a.due && dayDiff(a.due, today) < 0).length,
      blocked: lateOrBlocked.filter((a) => a.status === 'blocked').length,
      decisions: decisions.length,
      crs: crsDecided.length,
    };
    return {
      proj, since, until, today,
      kpis: k, changed, lateOrBlocked, doneInPeriod,
      decisions, crsDecided, topRisks,
      next: { milestones: upcomingMilestones, deliverables: upcomingDeliverables, actions: upcomingDue },
    };
  }
  function reportPeriodLabel() {
    const r = reportPeriodRange();
    return `${fmtFull(r.since)} – ${fmtFull(r.until)}`;
  }
  function reportToMarkdown(data) {
    const lines = [];
    const proj = data.proj;
    lines.push(`# ${proj.name} — Status report`);
    lines.push('');
    lines.push(`_${fmtFull(data.since)} – ${fmtFull(data.until)}_`);
    lines.push('');
    lines.push('## KPIs');
    lines.push('');
    lines.push(`| Done | Changed | Late | Blocked | Decisions | CRs decided |`);
    lines.push(`|---:|---:|---:|---:|---:|---:|`);
    lines.push(`| ${data.kpis.done} | ${data.kpis.changed} | ${data.kpis.late} | ${data.kpis.blocked} | ${data.kpis.decisions} | ${data.kpis.crs} |`);
    lines.push('');
    const section = (title, items, fmt) => {
      lines.push(`## ${title}`);
      lines.push('');
      if (!items.length) { lines.push('_None._'); lines.push(''); return; }
      items.forEach((it) => lines.push('- ' + fmt(it)));
      lines.push('');
    };
    section('What changed', data.changed.slice(0, 30), (a) =>
      `**${mdEscape(a.title)}** — ${mdEscape(personName(a.owner))} · ${a.status}${a.updatedAt ? ' · ' + a.updatedAt : ''}`);
    section('Late & blocked', data.lateOrBlocked, (a) => {
      const reason = a.status === 'blocked'
        ? 'blocked'
        : `${Math.abs(dayDiff(a.due, data.today))}d late`;
      return `**${mdEscape(a.title)}** — ${mdEscape(personName(a.owner))} · ${reason}`;
    });
    section('Decisions made', data.decisions, (d) =>
      `**${mdEscape(d.title)}** — ${mdEscape(personName(d.owner))} · ${d.date || '—'}${d.rationale ? '\n  > ' + mdEscape(d.rationale) : ''}`);
    section('Change requests decided', data.crsDecided, (c) =>
      `**${mdEscape(c.title)}** — ${c.status}${c.decisionDate ? ' · ' + c.decisionDate : ''}${c.decisionBy ? ' · ' + mdEscape(personName(c.decisionBy)) : ''}`);
    section('Top risks (by residual)', data.topRisks, (r) =>
      `**${mdEscape(r.title)}** — residual ${r._score}${r.mitigation ? ' · ' + mdEscape(r.mitigation) : ''}`);
    lines.push(`## What's next (next 14 days)`);
    lines.push('');
    if (!data.next.milestones.length && !data.next.deliverables.length && !data.next.actions.length) {
      lines.push('_Nothing scheduled in the next 14 days._');
    } else {
      if (data.next.milestones.length) {
        lines.push('**Milestones**');
        data.next.milestones.forEach((m) => lines.push(`- ${mdEscape(m.name || m.title || '')} · ${m.date}`));
        lines.push('');
      }
      if (data.next.deliverables.length) {
        lines.push('**Deliverables**');
        data.next.deliverables.forEach((d) => lines.push(`- ${mdEscape(d.name || d.title || '')} · ${d.date}`));
        lines.push('');
      }
      if (data.next.actions.length) {
        lines.push('**Due actions**');
        data.next.actions.slice(0, 30).forEach((a) =>
          lines.push(`- ${mdEscape(a.title)} — ${mdEscape(personName(a.owner))} · ${a.due}`));
      }
    }
    return lines.join('\n');
  }
  function reportToPrintHTML(data) {
    const k = data.kpis;
    const proj = data.proj;
    const list = (arr, fmt, empty) => arr.length
      ? '<ul>' + arr.map(fmt).join('') + '</ul>'
      : `<p class="empty">${empty}</p>`;
    const li = (a) => `<li><b>${escapeHTML(a.title)}</b> — ${escapeHTML(personName(a.owner))} · ${escapeHTML(a.status)}${a.updatedAt ? ' · ' + a.updatedAt : ''}</li>`;
    const liDue = (a) => {
      const reason = a.status === 'blocked' ? 'blocked' : `${Math.abs(dayDiff(a.due, data.today))}d late`;
      return `<li><b>${escapeHTML(a.title)}</b> — ${escapeHTML(personName(a.owner))} · ${reason}</li>`;
    };
    const liDec = (d) => `<li><b>${escapeHTML(d.title)}</b> — ${escapeHTML(personName(d.owner))} · ${d.date || '—'}${d.rationale ? '<br/><i>' + escapeHTML(d.rationale) + '</i>' : ''}</li>`;
    const liCR = (c) => `<li><b>${escapeHTML(c.title)}</b> — ${escapeHTML(c.status)}${c.decisionDate ? ' · ' + c.decisionDate : ''}${c.decisionBy ? ' · ' + escapeHTML(personName(c.decisionBy)) : ''}</li>`;
    const liRisk = (r) => `<li><b>${escapeHTML(r.title)}</b> — residual ${r._score}${r.mitigation ? '<br/><i>' + escapeHTML(r.mitigation) + '</i>' : ''}</li>`;
    const liNext = (it, dateField) => `<li><b>${escapeHTML(it.name || it.title || '')}</b> · ${escapeHTML(it[dateField] || '—')}</li>`;
    const liActDue = (a) => `<li><b>${escapeHTML(a.title)}</b> — ${escapeHTML(personName(a.owner))} · due ${a.due}</li>`;

    return `<!doctype html><html><head><meta charset="utf-8"><title>${escapeHTML(proj.name)} — Status report</title>
<style>
  body{font:14px -apple-system,Segoe UI,Helvetica,Arial,sans-serif;color:#1a1c24;max-width:820px;margin:30px auto;padding:0 20px;}
  h1{font-size:22px;margin:0 0 4px;} h2{font-size:15px;margin:22px 0 8px;border-bottom:1px solid #ddd;padding-bottom:4px;letter-spacing:.04em;text-transform:uppercase;color:#444;}
  .meta{color:#666;font-size:12px;margin-bottom:18px;}
  .kpis{display:grid;grid-template-columns:repeat(6,1fr);gap:10px;margin:14px 0;}
  .kpi{background:#f6f7fb;border:1px solid #e5e7ee;border-radius:8px;padding:10px;text-align:center;}
  .kpi b{font-size:22px;display:block;font-weight:700;}
  .kpi span{font-size:11px;color:#666;text-transform:uppercase;letter-spacing:.06em;}
  ul{padding-left:18px;margin:6px 0;} li{margin:5px 0;} .empty{color:#888;font-style:italic;margin:6px 0;}
  @media print { body { margin: 0 auto; } }
</style></head><body>
<h1>${escapeHTML(proj.name)} — Status report</h1>
<div class="meta">${fmtFull(data.since)} – ${fmtFull(data.until)} · generated ${fmtFull(data.today)}</div>
<div class="kpis">
  <div class="kpi"><b>${k.done}</b><span>Done</span></div>
  <div class="kpi"><b>${k.changed}</b><span>Changed</span></div>
  <div class="kpi"><b>${k.late}</b><span>Late</span></div>
  <div class="kpi"><b>${k.blocked}</b><span>Blocked</span></div>
  <div class="kpi"><b>${k.decisions}</b><span>Decisions</span></div>
  <div class="kpi"><b>${k.crs}</b><span>CRs decided</span></div>
</div>
<h2>What changed</h2>${list(data.changed.slice(0, 30), li, 'No updates in this period.')}
<h2>Late &amp; blocked</h2>${list(data.lateOrBlocked, liDue, 'Nothing late or blocked — nice work.')}
<h2>Decisions made</h2>${list(data.decisions, liDec, 'No decisions logged in this period.')}
<h2>Change requests decided</h2>${list(data.crsDecided, liCR, 'No CRs decided in this period.')}
<h2>Top risks (by residual)</h2>${list(data.topRisks, liRisk, 'No risks logged.')}
<h2>What's next (next 14 days)</h2>
${data.next.milestones.length ? '<p><b>Milestones</b></p>' + list(data.next.milestones, (m) => liNext(m, 'date'), '') : ''}
${data.next.deliverables.length ? '<p><b>Deliverables</b></p>' + list(data.next.deliverables, (d) => liNext(d, 'date'), '') : ''}
${data.next.actions.length ? '<p><b>Due actions</b></p>' + list(data.next.actions.slice(0, 30), liActDue, '') : ''}
${(!data.next.milestones.length && !data.next.deliverables.length && !data.next.actions.length) ? '<p class="empty">Nothing scheduled in the next 14 days.</p>' : ''}
</body></html>`;
  }
  // Reports is merged into Review — keep the function name for any
  // remaining callers (palette, stale state.currentView), but route to
  // the merged Review in 'full' mode so the old entry-point gives the
  // single-page snapshot users expected.
  function renderReports(root) {
    reviewModeState.mode = 'full';
    return renderReview(root);
  }
  // The full implementation below is dead-code post-merge but kept for
  // reference and so any direct callers continue to compile.
  function renderReports_legacy_unused(root) {
    const proj = curProject();
    const view = document.createElement('div');
    view.className = 'view';
    const isAll = state.currentProjectId === '__all__';
    if (isAll || !proj) {
      view.innerHTML = `
        <div class="page-head">
          <div>
            <div class="page-title">Status report</div>
            <div class="page-sub">Pick a project to generate a report.</div>
          </div>
        </div>
        <div class="empty">Reports are project-scoped. Choose a project from the topbar.</div>`;
      root.appendChild(view);
      return;
    }
    const range = reportPeriodRange();
    const data = buildReportData(proj, range.since, range.until);
    const k = data.kpis;
    const optSel = (val) => reportState.period === val ? 'selected' : '';
    view.innerHTML = `
      <div class="page-head">
        <div>
          <div class="page-title">${escapeHTML(proj.name)} — Status report</div>
          <div class="page-sub">${fmtFull(data.since)} – ${fmtFull(data.until)}</div>
        </div>
        <div class="page-actions">
          <select id="reportPeriod" class="report-period">
            <option value="7d"  ${optSel('7d')}>Last 7 days</option>
            <option value="30d" ${optSel('30d')}>Last 30 days</option>
            <option value="90d" ${optSel('90d')}>Last 90 days</option>
            <option value="custom" ${optSel('custom')}>Custom…</option>
          </select>
          <span class="report-custom" id="reportCustom" ${reportState.period === 'custom' ? '' : 'hidden'}>
            <input type="date" id="reportSince" value="${reportState.customSince || data.since}" />
            <span class="report-dash">–</span>
            <input type="date" id="reportUntil" value="${reportState.customUntil || data.until}" />
          </span>
          <button class="ghost" id="btnReportCopyMd"  title="Copy as Markdown">Copy Markdown</button>
          <button class="ghost" id="btnReportDownload" title="Download .md file">Download .md</button>
          <button class="ghost" id="btnReportPrint"   title="Open print-ready HTML in a new tab">Print → PDF</button>
        </div>
      </div>
      <div class="report">
        <div class="report-kpis">
          <div class="report-kpi"><div class="report-kpi-num">${k.done}</div><div class="report-kpi-lbl">Done</div></div>
          <div class="report-kpi"><div class="report-kpi-num">${k.changed}</div><div class="report-kpi-lbl">Changed</div></div>
          <div class="report-kpi ${k.late > 0 ? 'bad' : ''}"><div class="report-kpi-num">${k.late}</div><div class="report-kpi-lbl">Late</div></div>
          <div class="report-kpi ${k.blocked > 0 ? 'warn' : ''}"><div class="report-kpi-num">${k.blocked}</div><div class="report-kpi-lbl">Blocked</div></div>
          <div class="report-kpi"><div class="report-kpi-num">${k.decisions}</div><div class="report-kpi-lbl">Decisions</div></div>
          <div class="report-kpi"><div class="report-kpi-num">${k.crs}</div><div class="report-kpi-lbl">CRs decided</div></div>
        </div>

        ${reportSection('What changed', data.changed.slice(0, 30), (a) =>
          `<div class="report-row clickable" data-open-action="${a.id}"><span>${escapeHTML(a.title)}</span><span class="report-meta">${escapeHTML(personName(a.owner))} · ${escapeHTML(a.status)}${a.updatedAt ? ' · ' + a.updatedAt : ''}</span></div>`,
          'No updates in this period.')}

        ${reportSection('Late & blocked', data.lateOrBlocked, (a) => {
          const reason = a.status === 'blocked' ? 'blocked' : `${Math.abs(dayDiff(a.due, data.today))}d late`;
          return `<div class="report-row clickable bad" data-open-action="${a.id}"><span>${escapeHTML(a.title)}</span><span class="report-meta">${escapeHTML(personName(a.owner))} · ${reason}</span></div>`;
        }, 'Nothing late or blocked — nice work.')}

        ${reportSection('Decisions made', data.decisions, (d) =>
          `<div class="report-row"><span>${escapeHTML(d.title)}</span><span class="report-meta">${escapeHTML(personName(d.owner))} · ${escapeHTML(d.date || '—')}</span></div>${d.rationale ? `<div class="report-rationale">${escapeHTML(d.rationale)}</div>` : ''}`,
          'No decisions logged in this period.')}

        ${reportSection('Change requests decided', data.crsDecided, (c) =>
          `<div class="report-row clickable" data-open-cr="${c.id}"><span>${escapeHTML(c.title)}</span><span class="report-meta">${escapeHTML(c.status)}${c.decisionDate ? ' · ' + c.decisionDate : ''}${c.decisionBy ? ' · ' + escapeHTML(personName(c.decisionBy)) : ''}</span></div>`,
          'No CRs decided in this period.')}

        ${reportSection('Top risks (by residual)', data.topRisks, (r) =>
          `<div class="report-row clickable" data-open-risk="${r.id}"><span>${escapeHTML(r.title)}</span><span class="report-meta">residual ${r._score}</span></div>${r.mitigation ? `<div class="report-rationale">${escapeHTML(r.mitigation)}</div>` : ''}`,
          'No risks logged.')}

        <div class="report-section">
          <div class="report-section-title">What's next (next 14 days)</div>
          ${(data.next.milestones.length || data.next.deliverables.length || data.next.actions.length)
            ? `${data.next.milestones.length ? '<div class="report-sub-title">Milestones</div>' + data.next.milestones.map((m) => `<div class="report-row"><span>${escapeHTML(m.name || m.title || '')}</span><span class="report-meta">${escapeHTML(m.date || '')}</span></div>`).join('') : ''}
               ${data.next.deliverables.length ? '<div class="report-sub-title">Deliverables</div>' + data.next.deliverables.map((d) => `<div class="report-row"><span>${escapeHTML(d.name || d.title || '')}</span><span class="report-meta">${escapeHTML(d.date || '')}</span></div>`).join('') : ''}
               ${data.next.actions.length ? '<div class="report-sub-title">Due actions</div>' + data.next.actions.slice(0, 30).map((a) => `<div class="report-row clickable" data-open-action="${a.id}"><span>${escapeHTML(a.title)}</span><span class="report-meta">${escapeHTML(personName(a.owner))} · ${escapeHTML(a.due || '')}</span></div>`).join('') : ''}`
            : '<div class="empty">Nothing scheduled in the next 14 days.</div>'}
        </div>
      </div>`;
    root.appendChild(view);

    // Interactivity
    $('#reportPeriod').addEventListener('change', (e) => {
      reportState.period = e.target.value;
      render();
    });
    $('#reportSince')?.addEventListener('change', (e) => {
      reportState.customSince = e.target.value; reportState.period = 'custom'; render();
    });
    $('#reportUntil')?.addEventListener('change', (e) => {
      reportState.customUntil = e.target.value; reportState.period = 'custom'; render();
    });
    $('#btnReportCopyMd').addEventListener('click', async () => {
      const md = reportToMarkdown(buildReportData(curProject(), range.since, range.until));
      try { await navigator.clipboard.writeText(md); toast('Copied as Markdown'); }
      catch (e) { toast('Copy failed — clipboard unavailable'); }
    });
    $('#btnReportDownload').addEventListener('click', () => {
      const md = reportToMarkdown(buildReportData(curProject(), range.since, range.until));
      const blob = new Blob([md], { type: 'text/markdown' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url; a.download = `report-${proj.id}-${range.until}.md`;
      a.click();
      URL.revokeObjectURL(url);
      toast('Report downloaded');
    });
    $('#btnReportPrint').addEventListener('click', () => {
      const html = reportToPrintHTML(buildReportData(curProject(), range.since, range.until));
      const w = window.open('', '_blank');
      if (!w) { toast('Pop-up blocked'); return; }
      w.document.write(html); w.document.close();
      setTimeout(() => { try { w.print(); } catch (e) { /* ignore */ } }, 250);
    });

    // Row drilldowns
    view.querySelectorAll('[data-open-action]').forEach((el) => {
      el.addEventListener('click', () => openDrawer(el.dataset.openAction));
    });
    view.querySelectorAll('[data-open-cr]').forEach((el) => {
      el.addEventListener('click', () => openChangeRequestEditor(el.dataset.openCr));
    });
    view.querySelectorAll('[data-open-risk]').forEach((el) => {
      el.addEventListener('click', () => openRiskEditor(el.dataset.openRisk));
    });
  }
  function reportSection(title, items, fmt, empty) {
    return `
      <div class="report-section">
        <div class="report-section-title">${escapeHTML(title)}<span class="report-section-count">${items.length}</span></div>
        ${items.length ? items.map(fmt).join('') : `<div class="empty">${escapeHTML(empty)}</div>`}
      </div>`;
  }

  /* ----------------------- Phase H: Calendar view ---------------------- */
  // Month grid (7 cols × N rows) of every dated item. Each cell shows
  // colour-tinted chips per kind: action due, milestone, deliverable,
  // CR decided, meeting. Arrow keys navigate months. Clicking a chip
  // opens the relevant drawer / editor.
  const calState = {
    monthOffset: 0, // 0 = current month, -1 = previous, +1 = next
    keyWired: false,
    // 'month' = vertically-scrolling weekly grid with continuous reveal
    // (default); 'timeline' = horizontal swim-lane view.
    format: 'month',
    // Month-mode window state. firstWeekStart = Monday of the first
    // rendered week. weekCount = how many weeks are currently rendered.
    // windowAnchorOffset tracks the monthOffset the window was last
    // built for; mismatch means re-center.
    firstWeekStart: null,
    weekCount: 12,
    windowAnchorOffset: null,
    // Per-kind visibility filters. The legend doubles as the filter
    // strip: clicking a pill toggles its kind on/off. Persists across
    // month navigation but resets on reload (intentional — fresh
    // sessions start with everything shown).
    visible: { action: true, deliverable: true, milestone: true, cr: true, meeting: true },
  };
  // ISO 8601 week number — Mon-anchored, week 1 is the one containing
  // the first Thursday of the year.
  function isoWeekNumber(d) {
    const dt = new Date(d.getFullYear(), d.getMonth(), d.getDate());
    dt.setDate(dt.getDate() + 3 - ((dt.getDay() + 6) % 7));
    const week1 = new Date(dt.getFullYear(), 0, 4);
    return 1 + Math.round(((dt.getTime() - week1.getTime()) / 86400000 - 3 + ((week1.getDay() + 6) % 7)) / 7);
  }
  function ensureCalendarWindow() {
    const today = new Date();
    if (calState.windowAnchorOffset !== calState.monthOffset || !calState.firstWeekStart) {
      const anchor = new Date(today.getFullYear(), today.getMonth() + calState.monthOffset, 1);
      const dow = (anchor.getDay() + 6) % 7; // Mon=0
      const monthGridStart = new Date(anchor.getFullYear(), anchor.getMonth(), 1 - dow);
      // Lead with 4 weeks before the anchor month start, 12 weeks total
      calState.firstWeekStart = new Date(monthGridStart.getTime() - 4 * 7 * dayMs);
      calState.weekCount = 12;
      calState.windowAnchorOffset = calState.monthOffset;
    }
  }
  function calendarMonthBounds() {
    ensureCalendarWindow();
    const today = new Date();
    const anchor = new Date(today.getFullYear(), today.getMonth() + calState.monthOffset, 1);
    const year = anchor.getFullYear();
    const month = anchor.getMonth();
    const monthLabel = anchor.toLocaleDateString(undefined, { year: 'numeric', month: 'long' });
    const cells = [];
    for (let i = 0; i < calState.weekCount * 7; i++) {
      const d = new Date(calState.firstWeekStart);
      d.setDate(calState.firstWeekStart.getDate() + i);
      cells.push(d);
    }
    return { year, month, monthLabel, cells };
  }
  function buildCalendarItems(proj, year, month, gridStartISO, gridEndISO) {
    const inRange = (d) => d && d >= gridStartISO && d <= gridEndISO;
    const items = []; // { date, kind, label, tone, run, icon, rangePos }
    const acts = (proj.actions || []).filter((a) => !a.deletedAt);
    acts.forEach((a) => {
      if (!a.due || !inRange(a.due)) return;
      const cmp = a.component ? findComponent(proj, a.component) : null;
      const cmpRgb = cmp ? componentColor(cmp.color)?.rgb : null;
      const subBase = personName(a.owner);
      items.push({
        date: a.due,
        kind: 'action',
        tone: a.status === 'done' ? 'muted' : (a.status === 'blocked' ? 'bad' : 'accent'),
        label: a.title,
        sub: cmp ? `${subBase} · ${cmp.name}` : subBase,
        // Checkbox glyph reflects status — filled box for done, empty
        // outline otherwise. Reads at-a-glance like a to-do checklist.
        icon: a.status === 'done' ? '☑' : '☐',
        tint: cmpRgb,
        run: () => openDrawer(a.id),
      });
    });
    (proj.milestones || []).forEach((m) => {
      if (!m.date) return;
      // Ranged milestones expand into one chip per day in [date … endDate],
      // each tagged with its position (start / middle / end / single) so the
      // calendar can paint a connected band across cells.
      const start = m.date;
      const end = m.endDate || m.date;
      if (!inRange(start) && !inRange(end) && !(start <= gridStartISO && end >= gridEndISO)) return;
      const startT = parseDate(start).getTime();
      const endT   = parseDate(end).getTime();
      const isRange = endT > startT;
      const cmp = m.component ? findComponent(proj, m.component) : null;
      const cmpRgb = cmp ? componentColor(cmp.color)?.rgb : null;
      for (let t = startT; t <= endT; t += dayMs) {
        const iso = fmtISO(new Date(t));
        if (!inRange(iso)) continue;
        const isStart = iso === start;
        const isEnd   = iso === end;
        const rangePos = !isRange ? 'single' : isStart ? 'start' : isEnd ? 'end' : 'middle';
        const icon = !isRange ? '◇' : isStart ? '▷' : isEnd ? '▭' : '─';
        const baseSub = isRange
          ? (isStart ? `Milestone start · ends ${end}` : isEnd ? `Milestone end · started ${start}` : 'Milestone in progress')
          : 'Milestone';
        items.push({
          date: iso,
          kind: 'milestone',
          tone: 'milestone',
          label: m.name,
          sub: cmp ? `${baseSub} · ${cmp.name}` : baseSub,
          icon,
          rangePos,
          tint: cmpRgb,
          run: () => openMilestoneEditor(m.id),
        });
      }
    });
    (proj.deliverables || []).forEach((d) => {
      const dt = d.dueDate || d.date;
      if (!dt || !inRange(dt)) return;
      const cmp = d.component ? findComponent(proj, d.component) : null;
      const cmpRgb = cmp ? componentColor(cmp.color)?.rgb : null;
      items.push({
        date: dt,
        kind: 'deliverable',
        tone: 'deliverable',
        label: d.name,
        sub: cmp ? `Deliverable · ${cmp.name}` : 'Deliverable',
        icon: '◆',
        tint: cmpRgb,
        run: () => openDeliverableEditor(d.id),
      });
    });
    (proj.changes || []).forEach((c) => {
      if (!c.decisionDate || !inRange(c.decisionDate)) return;
      const decided = ['approved', 'rejected', 'implemented', 'cancelled'].includes(c.status);
      if (!decided) return;
      items.push({
        date: c.decisionDate,
        kind: 'cr',
        tone: c.status === 'rejected' || c.status === 'cancelled' ? 'bad' : 'good',
        label: c.title,
        sub: 'CR ' + c.status,
        icon: '⇆',
        run: () => openChangeRequestEditor(c.id),
      });
    });
    (proj.meetings || []).forEach((m) => {
      const dates = expandMeetingDates(m, gridStartISO, gridEndISO);
      if (!dates.length) return;
      const baseLabel = meetingRecurrenceLabel(m);
      const cmp = m.component ? findComponent(proj, m.component) : null;
      const cmpRgb = cmp ? componentColor(cmp.color)?.rgb : null;
      const subBase = m.time ? `${baseLabel} · ${m.time}` : baseLabel;
      const subWithComponent = cmp ? `${subBase} · ${cmp.name}` : subBase;
      dates.forEach((iso) => {
        items.push({
          date: iso,
          kind: 'meeting',
          tone: 'meeting',
          label: m.title,
          sub: subWithComponent,
          icon: '⊕',
          // When a component is set, the chip's icon + label pick up the
          // component's colour so meetings cluster visually with their
          // related work area. Stored on the item so the chip renderer
          // can apply it inline alongside the existing tone class.
          tint: cmpRgb,
          run: () => openMeetingEditor(m.id),
        });
      });
    });
    return items;
  }
  // Horizontal swim-lane variant of the calendar — same window, same data,
  // same chips — but laid out with one lane per kind, with chips placed by
  // date offset and ranged milestones stretching across days as a bar.
  function renderCalendarTimeline(cells, items, byDate, todayISO_) {
    const startMs = cells[0].getTime();
    const totalDays = cells.length;
    const dayPx = 30; // column width per day
    const totalPx = totalDays * dayPx;

    // Build per-kind lanes from the (already-filter-aware) flat items list,
    // skipping milestone middle/end positions — they're absorbed into the
    // start chip's wider bar so the range reads as one continuous element
    // rather than 4 separate chips.
    const laneOrder = ['milestone', 'deliverable', 'action', 'meeting', 'cr'];
    const laneLabels = { milestone: 'Milestones', deliverable: 'Deliverables', action: 'Actions', meeting: 'Meetings', cr: 'CRs' };
    const laneIcons  = { milestone: '◇', deliverable: '◆', action: '☐', meeting: '⊕', cr: '⇆' };
    const lanes = laneOrder
      .filter((k) => calState.visible[k] !== false)
      .map((k) => ({ kind: k, items: [] }));
    const laneByKind = Object.fromEntries(lanes.map((l) => [l.kind, l]));

    // Track which (item.key) we've already added so a milestone band's
    // start chip is the only one we keep (its width covers the range).
    const seenMilestoneIds = new Set();
    items.forEach((it) => {
      const lane = laneByKind[it.kind];
      if (!lane) return;
      // Range-aware milestone packing: only the 'start' / 'single' positions
      // create a chip; middle/end positions are skipped because the start
      // chip's width spans the range.
      if (it.kind === 'milestone' && it.rangePos && (it.rangePos === 'middle' || it.rangePos === 'end')) {
        return;
      }
      lane.items.push(it);
    });

    // Index helper (offset days from the grid start)
    const offsetDays = (iso) => Math.round((parseDate(iso).getTime() - startMs) / dayMs);

    // Today indicator x-position
    const todayOffset = todayISO_ >= fmtISO(cells[0]) && todayISO_ <= fmtISO(cells[cells.length - 1])
      ? offsetDays(todayISO_) * dayPx + dayPx / 2 : null;

    // Axis: render a tick + label every 7 days, plus highlight Mondays.
    // Labels show 'Mon Apr 27' / 'May 4' etc; abbreviated to fit 30px column.
    const axisCells = cells.map((d, i) => {
      const iso = fmtISO(d);
      const isToday = iso === todayISO_;
      const isMon = d.getDay() === 1;
      const showLabel = isMon || i === 0;
      const lbl = showLabel
        ? (d.getDate() === 1 || i === 0
            ? d.toLocaleDateString(undefined, { month: 'short', day: 'numeric' })
            : String(d.getDate()))
        : '';
      return `<div class="tl-axis-cell ${isToday ? 'is-today' : ''} ${isMon ? 'is-mon' : ''}" data-iso="${iso}" style="width:${dayPx}px;">${lbl}</div>`;
    }).join('');

    // For each item we need to know its index in byDate.get(iso) so chip
    // clicks can resolve back to the run() callback. We re-derive that here.
    function indexInDateBucket(it) {
      const bucket = byDate.get(it.date) || [];
      return bucket.indexOf(it);
    }

    // Render lanes. Each chip is absolutely-positioned by left/width.
    // Range milestones get width = (endOffset - startOffset + 1) * dayPx;
    // everything else gets width = dayPx.
    function renderLane(lane) {
      const chipsHTML = lane.items.map((it) => {
        const idx = indexInDateBucket(it);
        const x = offsetDays(it.date) * dayPx;
        // Compute width for milestone ranges; use the underlying record so
        // we can read endDate (item itself only carries rangePos+date).
        let w = dayPx - 2;
        if (it.kind === 'milestone' && it.rangePos === 'start') {
          // Find the matching 'end' chip to derive width
          const startD = it.date;
          // The day-bucketed items include 'end' positions; locate the same
          // milestone label across the whole `items` array.
          const same = items.filter((x) => x.kind === 'milestone' && x.label === it.label);
          const lastIso = same.reduce((mx, x) => x.date > mx ? x.date : mx, startD);
          const span = offsetDays(lastIso) - offsetDays(startD) + 1;
          w = Math.max(dayPx - 2, span * dayPx - 2);
        }
        const tintStyle = it.tint ? `--cal-chip-tint: ${it.tint};` : '';
        const tintCls = it.tint ? ' has-tint' : '';
        const rangeCls = it.rangePos ? ` is-${it.rangePos}` : '';
        return `
          <button class="tl-chip cal-chip cal-chip-${it.tone} kind-${it.kind}${rangeCls}${tintCls}"
                  style="left:${x}px;width:${w}px;${tintStyle}"
                  data-tl-item-key="${escapeHTML(it.date)}|${idx}"
                  title="${escapeHTML(it.label)} — ${escapeHTML(it.sub || '')}">
            <span class="cal-chip-icon" aria-hidden="true">${escapeHTML(it.icon || '·')}</span>
            <span class="cal-chip-label">${escapeHTML(it.label)}</span>
          </button>`;
      }).join('');
      const empty = !lane.items.length ? '<div class="tl-lane-empty">— none</div>' : '';
      return `
        <div class="tl-lane">
          <div class="tl-lane-label">
            <span class="cal-chip-icon kind-${lane.kind}">${laneIcons[lane.kind]}</span>
            ${escapeHTML(laneLabels[lane.kind])}
          </div>
          <div class="tl-lane-track" style="width:${totalPx}px;">
            ${chipsHTML}
            ${empty}
          </div>
        </div>`;
    }

    return `
      <div class="cal-timeline">
        <div class="tl-scroll">
          <div class="tl-axis-row">
            <div class="tl-lane-label tl-axis-spacer"></div>
            <div class="tl-axis-track" style="width:${totalPx}px;">
              ${axisCells}
              ${todayOffset != null ? `<div class="tl-today-line" style="left:${todayOffset}px;"></div>` : ''}
            </div>
          </div>
          ${lanes.map(renderLane).join('')}
        </div>
      </div>`;
  }

  function renderCalendar(root) {
    const proj = curProject();
    const view = document.createElement('div');
    view.className = 'view';
    const isAll = state.currentProjectId === '__all__';
    if (isAll || !proj) {
      view.innerHTML = `
        <div class="page-head">
          <div>
            <div class="page-title">Calendar</div>
            <div class="page-sub">Pick a project to see its month grid.</div>
          </div>
        </div>
        <div class="empty">Calendar is project-scoped. Choose a project from the topbar.</div>`;
      root.appendChild(view);
      return;
    }
    const { year, month, monthLabel, cells } = calendarMonthBounds();
    const gridStartISO = fmtISO(cells[0]);
    const gridEndISO   = fmtISO(cells[cells.length - 1]);
    const allItems = buildCalendarItems(proj, year, month, gridStartISO, gridEndISO);
    // Apply per-kind visibility filters (driven by the interactive legend
    // below). buildCalendarItems still returns everything so the toggle
    // state can flip without rebuilding the source data.
    const items = allItems.filter((it) => calState.visible[it.kind] !== false);
    const byDate = new Map();
    items.forEach((it) => {
      if (!byDate.has(it.date)) byDate.set(it.date, []);
      byDate.get(it.date).push(it);
    });
    // Per-kind counts for the legend (using the unfiltered set so users
    // see the underlying total, not the post-filter count).
    const kindCounts = { milestone: 0, deliverable: 0, action: 0, cr: 0, meeting: 0 };
    allItems.forEach((it) => { if (it.kind in kindCounts) kindCounts[it.kind]++; });

    const todayISO_ = todayISO();
    const dayHeaders = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun'];
    // Build week-row HTML. Each row is one week-number cell + 7 day cells.
    // Cells stay addressable via the original linear cell index so chip
    // wiring (data-cell-idx) keeps working without changes elsewhere.
    function dayCellHTML(d, i) {
      const iso = fmtISO(d);
      const inMonth = d.getMonth() === month;
      const isToday = iso === todayISO_;
      const isWeekend = d.getDay() === 0 || d.getDay() === 6;
      const isFirstOfMonth = d.getDate() === 1;
      const cellItems = byDate.get(iso) || [];
      const VISIBLE = 4;
      const shown = cellItems.slice(0, VISIBLE);
      const hidden = cellItems.length - shown.length;
      const chipsHTML = shown.map((it, idx) => {
        const icon = it.icon || '·';
        const rangeCls = it.kind === 'milestone' && it.rangePos
          ? ` is-${it.rangePos}` : '';
        const tintStyle = it.tint ? ` style="--cal-chip-tint: ${it.tint};"` : '';
        const tintCls = it.tint ? ' has-tint' : '';
        return `
          <button class="cal-chip cal-chip-${it.tone} kind-${it.kind}${rangeCls}${tintCls}" data-cell-idx="${i}" data-item-idx="${idx}" title="${escapeHTML(it.label)} — ${escapeHTML(it.sub || '')}"${tintStyle}>
            <span class="cal-chip-icon" aria-hidden="true">${escapeHTML(icon)}</span>
            <span class="cal-chip-label">${escapeHTML(it.label)}</span>
          </button>`;
      }).join('');
      const moreHTML = hidden > 0 ? `<button class="cal-more" data-cell-idx="${i}">+ ${hidden} more</button>` : '';
      // First-of-month gets a small inline label so the user can find
      // month boundaries while scrolling.
      const dayLabel = isFirstOfMonth
        ? `<span class="cal-day-month-tag">${d.toLocaleDateString(undefined, { month: 'short' })}</span> ${d.getDate()}`
        : `${d.getDate()}`;
      return `
        <div class="cal-cell ${inMonth ? '' : 'out-of-month'} ${isToday ? 'today' : ''} ${isWeekend ? 'weekend' : ''}" data-cell-idx="${i}" data-iso="${iso}">
          <div class="cal-day-num">${dayLabel}</div>
          <div class="cal-chips">${chipsHTML}${moreHTML}</div>
        </div>`;
    }
    const weekRowsHTML = [];
    for (let w = 0; w < calState.weekCount; w++) {
      const startIdx = w * 7;
      const weekStart = cells[startIdx];
      const wn = isoWeekNumber(weekStart);
      const cellsInWeek = [];
      for (let j = 0; j < 7; j++) cellsInWeek.push(dayCellHTML(cells[startIdx + j], startIdx + j));
      weekRowsHTML.push(`
        <div class="cal-week-num" title="ISO week ${wn} (starts ${fmtISO(weekStart)})">${wn}</div>
        ${cellsInWeek.join('')}`);
    }
    const cellHTML = weekRowsHTML.join('');

    view.innerHTML = `
      <div class="page-head">
        <div>
          <div class="page-title">${escapeHTML(proj.name)} — Calendar</div>
          <div class="page-sub" id="calPageSub">${escapeHTML(monthLabel)} · scroll vertically for more weeks · ← / → for months</div>
        </div>
        <div class="page-actions">
          <div class="seg" role="tablist" aria-label="Calendar format">
            <button type="button" class="seg-btn ${calState.format === 'month'    ? 'active' : ''}" data-cal-format="month"    title="Month grid view">Month</button>
            <button type="button" class="seg-btn ${calState.format === 'timeline' ? 'active' : ''}" data-cal-format="timeline" title="Horizontal timeline view">Timeline</button>
          </div>
          <button class="ghost" id="calAddMile" title="Add a milestone">+ Milestone</button>
          <button class="ghost" id="calAddDel"  title="Add a deliverable">+ Deliverable</button>
          <button class="ghost" id="calAddMtg"  title="Add a meeting (one-off or recurring)">+ Meeting</button>
          <button class="icon-btn" id="calPrev" title="Previous month (←)">‹</button>
          <button class="ghost"   id="calToday" title="Jump to current month">Today</button>
          <button class="icon-btn" id="calNext" title="Next month (→)">›</button>
        </div>
      </div>
      <div class="cal-legend" role="group" aria-label="Show / hide on calendar">
        <button type="button" class="cal-legend-item ${calState.visible.milestone   ? 'on' : 'off'}" data-toggle-kind="milestone"   title="Show / hide milestones">
          <span class="cal-chip-icon kind-milestone">◇</span> Milestones <span class="cal-legend-count">${kindCounts.milestone}</span>
        </button>
        <button type="button" class="cal-legend-item ${calState.visible.deliverable ? 'on' : 'off'}" data-toggle-kind="deliverable" title="Show / hide deliverables">
          <span class="cal-chip-icon kind-deliverable">◆</span> Deliverables <span class="cal-legend-count">${kindCounts.deliverable}</span>
        </button>
        <button type="button" class="cal-legend-item ${calState.visible.action      ? 'on' : 'off'}" data-toggle-kind="action"      title="Show / hide actions">
          <span class="cal-chip-icon kind-action">☐</span> Actions <span class="cal-legend-count">${kindCounts.action}</span>
        </button>
        <button type="button" class="cal-legend-item ${calState.visible.meeting     ? 'on' : 'off'}" data-toggle-kind="meeting"     title="Show / hide meetings">
          <span class="cal-chip-icon kind-meeting">⊕</span> Meetings <span class="cal-legend-count">${kindCounts.meeting}</span>
        </button>
        <button type="button" class="cal-legend-item ${calState.visible.cr          ? 'on' : 'off'}" data-toggle-kind="cr"          title="Show / hide CR decisions">
          <span class="cal-chip-icon kind-cr">⇆</span> CRs <span class="cal-legend-count">${kindCounts.cr}</span>
        </button>
      </div>
      ${calState.format === 'month' ? `
        <div class="calendar">
          <div class="cal-head">
            <div class="cal-head-cell cal-head-wk" title="ISO week number">Wk</div>
            ${dayHeaders.map((h) => `<div class="cal-head-cell">${h}</div>`).join('')}
          </div>
          <div class="cal-grid" id="calGrid">${cellHTML}</div>
        </div>
      ` : renderCalendarTimeline(cells, items, byDate, todayISO_)}`;
    root.appendChild(view);

    $('#calPrev').addEventListener('click', () => { calState.monthOffset -= 1; render(); });
    $('#calNext').addEventListener('click', () => { calState.monthOffset += 1; render(); });
    $('#calToday').addEventListener('click', () => { calState.monthOffset = 0; render(); });
    $('#calAddMile').addEventListener('click', () => openQuickAdd('milestone'));
    $('#calAddDel').addEventListener('click', () => openQuickAdd('deliverable'));
    $('#calAddMtg').addEventListener('click', () => openQuickAdd('meeting'));

    // Format toggle (Month / Timeline)
    view.querySelectorAll('[data-cal-format]').forEach((btn) => {
      btn.addEventListener('click', () => {
        const f = btn.dataset.calFormat;
        if (calState.format === f) return;
        calState.format = f;
        render();
      });
    });

    // Legend pills double as visibility filters
    view.querySelectorAll('[data-toggle-kind]').forEach((btn) => {
      btn.addEventListener('click', () => {
        const k = btn.dataset.toggleKind;
        calState.visible[k] = !calState.visible[k];
        render();
      });
    });

    // Timeline-mode chip + bar wiring (mirrors the month-grid wiring)
    if (calState.format === 'timeline') {
      view.querySelectorAll('[data-tl-item-key]').forEach((el) => {
        el.addEventListener('click', (e) => {
          e.stopPropagation();
          const key = el.dataset.tlItemKey;
          // key encodes "iso:idx" — find the matching item
          const [iso, idxStr] = key.split('|');
          const idx = parseInt(idxStr, 10);
          const it = (byDate.get(iso) || [])[idx];
          if (it && typeof it.run === 'function') it.run();
        });
      });
    }

    // Per-chip click: dispatch to the item's run() callback. We store
    // (cell, item) indices on the button so the data closure stays simple.
    view.querySelectorAll('.cal-chip').forEach((btn) => {
      btn.addEventListener('click', (e) => {
        e.stopPropagation();
        const cellIdx = +btn.dataset.cellIdx;
        const itemIdx = +btn.dataset.itemIdx;
        const iso = fmtISO(cells[cellIdx]);
        const it = (byDate.get(iso) || [])[itemIdx];
        if (it && typeof it.run === 'function') it.run();
      });
    });
    // "+ N more" expands into a quick popover listing every chip for the day
    view.querySelectorAll('.cal-more').forEach((btn) => {
      btn.addEventListener('click', (e) => {
        e.stopPropagation();
        const cellIdx = +btn.dataset.cellIdx;
        const iso = fmtISO(cells[cellIdx]);
        const list = byDate.get(iso) || [];
        const r = btn.getBoundingClientRect();
        showContextMenu(r.left, r.bottom + 4, list.map((it) => ({
          icon: it.icon || '·',
          label: it.label,
          onClick: it.run || (() => {}),
        })));
      });
    });

    // Wire arrow-key navigation once globally (re-renders share state)
    if (!calState.keyWired) {
      calState.keyWired = true;
      document.addEventListener('keydown', (e) => {
        if (state.currentView !== 'calendar') return;
        const inField = ['INPUT', 'TEXTAREA', 'SELECT'].includes(e.target.tagName) || e.target.isContentEditable;
        if (inField) return;
        if (e.key === 'ArrowLeft')  { e.preventDefault(); calState.monthOffset -= 1; render(); }
        if (e.key === 'ArrowRight') { e.preventDefault(); calState.monthOffset += 1; render(); }
      });
    }

    // Drag-to-pan time. Both formats accept click-and-drag on empty
    // body space (chips and buttons opt out). Month view lives-translates
    // the grid then snaps to whole months on release; Timeline view
    // drag-scrolls within the existing horizontally-scrollable container,
    // and any over-pull at the edges accumulates into a month-shift on
    // release so the user can pan continuously past the visible window.
    if (calState.format === 'month') wireMonthGridDrag(view);
    else                              wireTimelineDragPan(view);

    // Month view: vertical scroll-to-reveal-more-weeks. When the user
    // scrolls within ~200 px of the top or bottom edge, prepend / append
    // 4 more weeks and adjust scrollTop so the visual position holds
    // steady — no jump under the cursor. Also keep the page-sub label
    // tracking the dominant month in the viewport.
    if (calState.format === 'month') wireMonthGridScroll(view);
  }
  function wireMonthGridScroll(viewEl) {
    const grid = viewEl.querySelector('#calGrid');
    if (!grid) return;
    const sub = viewEl.querySelector('#calPageSub');
    let extending = false; // re-entry guard during prepend/append
    function rowHeight() {
      const first = grid.querySelector('.cal-cell');
      return first ? first.getBoundingClientRect().height : 110;
    }
    function visibleCenterDate() {
      // Use the row at the vertical center of the viewport.
      const rh = rowHeight();
      if (!rh) return null;
      const rowsFromTop = Math.floor((grid.scrollTop + grid.clientHeight / 2) / rh);
      const idx = clamp(rowsFromTop * 7, 0, calState.weekCount * 7 - 1);
      const d = new Date(calState.firstWeekStart);
      d.setDate(calState.firstWeekStart.getDate() + idx);
      return d;
    }
    function refreshSub() {
      const d = visibleCenterDate();
      if (!d || !sub) return;
      const lbl = d.toLocaleDateString(undefined, { year: 'numeric', month: 'long' });
      sub.textContent = `${lbl} · scroll vertically for more weeks · ← / → for months`;
    }
    function extendTop() {
      if (extending) return;
      extending = true;
      const ADD = 4;
      const beforeH = grid.scrollHeight;
      const newStart = new Date(calState.firstWeekStart);
      newStart.setDate(newStart.getDate() - ADD * 7);
      calState.firstWeekStart = newStart;
      calState.weekCount += ADD;
      // Re-render and restore scroll position so the user's view doesn't
      // jump by the height of the prepended rows.
      const prevTop = grid.scrollTop;
      render();
      const newGrid = $('#calGrid');
      if (newGrid) {
        const afterH = newGrid.scrollHeight;
        newGrid.scrollTop = prevTop + (afterH - beforeH);
      }
      extending = false;
    }
    function extendBottom() {
      if (extending) return;
      extending = true;
      const ADD = 4;
      calState.weekCount += ADD;
      const prevTop = grid.scrollTop;
      render();
      const newGrid = $('#calGrid');
      if (newGrid) newGrid.scrollTop = prevTop;
      extending = false;
    }
    grid.addEventListener('scroll', () => {
      const max = grid.scrollHeight - grid.clientHeight;
      if (grid.scrollTop < 200) extendTop();
      else if (grid.scrollTop > max - 200) extendBottom();
      refreshSub();
    });
    // First paint: scroll today's row into the upper third of the viewport
    // (so a few weeks of context above are visible without scrolling).
    const todayCell = grid.querySelector('.cal-cell.today');
    if (todayCell) {
      const rh = rowHeight();
      const todayIdx = +todayCell.dataset.cellIdx;
      const todayRow = Math.floor(todayIdx / 7);
      grid.scrollTop = Math.max(0, todayRow * rh - 80);
    }
    refreshSub();
  }

  // — Drag-pan helpers —
  // Common rule: ignore drags that started on a chip / button / link /
  // form control so click-throughs still work.
  function _dragIgnoreTarget(t) {
    return !!t.closest('.cal-chip, .tl-chip, .cal-more, button, a, input, select, textarea, [contenteditable="true"]');
  }
  function wireMonthGridDrag(viewEl) {
    const grid = viewEl.querySelector('.cal-grid');
    if (!grid) return;
    let active = false;
    let startX = 0;
    let dx = 0;
    const PX_PER_MONTH = 200; // sensitivity — ~200 px of horizontal drag = 1 month
    grid.style.cursor = 'grab';
    grid.addEventListener('mousedown', (e) => {
      if (e.button !== 0) return;
      if (_dragIgnoreTarget(e.target)) return;
      active = true; startX = e.clientX; dx = 0;
      grid.style.cursor = 'grabbing';
      document.body.style.userSelect = 'none';
      e.preventDefault();
    });
    // Attach move/up at render time and leave them attached for the
    // lifetime of this view's render. They no-op when `active` is false,
    // which is the rest state. (Earlier we used { once: true } for mouseup
    // — that broke after the first drag because the listener was removed.)
    function onMove(e) {
      if (!active) return;
      dx = e.clientX - startX;
      grid.style.transform = `translateX(${dx}px)`;
    }
    function onUp() {
      if (!active) return;
      active = false;
      grid.style.cursor = 'grab';
      grid.style.transform = '';
      document.body.style.userSelect = '';
      const months = Math.round(-dx / PX_PER_MONTH);
      if (months !== 0) {
        calState.monthOffset += months;
        render();
      }
    }
    document.addEventListener('mousemove', onMove);
    document.addEventListener('mouseup', onUp);
  }
  function wireTimelineDragPan(viewEl) {
    const scroll = viewEl.querySelector('.tl-scroll');
    if (!scroll) return;
    let active = false;
    let startX = 0;
    let startScrollLeft = 0;
    let overflowPx = 0; // accumulated pull past the edges, used for month-shift on release
    const PX_PER_MONTH = 220;
    scroll.style.cursor = 'grab';
    scroll.addEventListener('mousedown', (e) => {
      if (e.button !== 0) return;
      if (_dragIgnoreTarget(e.target)) return;
      active = true;
      startX = e.clientX;
      startScrollLeft = scroll.scrollLeft;
      overflowPx = 0;
      scroll.style.cursor = 'grabbing';
      document.body.style.userSelect = 'none';
      e.preventDefault();
    });
    function onMove(e) {
      if (!active) return;
      const dx = e.clientX - startX;
      const desired = startScrollLeft - dx;
      const max = Math.max(0, scroll.scrollWidth - scroll.clientWidth);
      // Clamp scroll to [0, max] and remember any overflow so we can
      // convert it to a month-shift on release.
      if (desired < 0)        { scroll.scrollLeft = 0;   overflowPx = desired; }
      else if (desired > max) { scroll.scrollLeft = max; overflowPx = desired - max; }
      else                    { scroll.scrollLeft = desired; overflowPx = 0; }
    }
    function onUp() {
      if (!active) return;
      active = false;
      scroll.style.cursor = 'grab';
      document.body.style.userSelect = '';
      // If the user pulled past either edge, convert the overflow to
      // a month shift so panning feels continuous past the visible window.
      const months = Math.round(overflowPx / PX_PER_MONTH);
      if (months !== 0) {
        calState.monthOffset += months;
        render();
      }
    }
    document.addEventListener('mousemove', onMove);
    document.addEventListener('mouseup', onUp);
  }

  /* ---------------------- Phase C: command palette --------------------- */
  // Universal Cmd+K palette. Indexes everything searchable + a small menu of
  // "slash commands" (action you can take). Up/Down + Enter to navigate +
  // open. Type `/` to filter to commands only.
  const paletteState = { items: [], idx: 0, query: '' };

  // Build a flat searchable index from current state. Cheap to rebuild on
  // each open since project data is in-memory.
  function buildPaletteIndex() {
    const out = [];
    // Slash-commands always available
    const sc = (label, hint, run) => out.push({ kind: 'cmd', label, hint, run, sortBoost: 5 });
    sc('+ New action',          'Open Quick Add',    () => openQuickAdd('action'));
    sc('+ New deliverable',     'Open Quick Add',    () => openQuickAdd('deliverable'));
    sc('+ New milestone',       'Open Quick Add',    () => openQuickAdd('milestone'));
    sc('+ New risk',            'Open Quick Add',    () => openQuickAdd('risk', { kind: 'risk' }));
    sc('+ New opportunity',     'Open Quick Add',    () => openQuickAdd('risk', { kind: 'opportunity' }));
    sc('+ New decision',        'Open Quick Add',    () => openQuickAdd('decision'));
    sc('+ New change request',  'Open editor',       () => openQuickAdd('change'));
    sc('+ New link',            'Open Quick Add',    () => openQuickAdd('link'));
    sc('+ New meeting',         'Open Quick Add',    () => openQuickAdd('meeting'));
    sc('+ New person',          'Open Quick Add',    () => openQuickAdd('person'));
    sc('+ New project',         'Open Quick Add',    () => openQuickAdd('project'));
    sc('Save current project as template', 'Project skeleton',
      () => { const p = curProject(); if (!p || curProjectIsMerged()) { toast('Pick a single project'); return; } saveProjectAsTemplate(p.id); });
    sc('Open Inbox',            'Reminders + alerts',() => { state.currentView = 'inbox'; render(); });
    sc('Open Calendar',         'Month view',        () => { state.currentView = 'calendar'; render(); });
    sc('Open Review (walk-through)', 'Live review wizard', () => { reviewModeState.mode = 'walkthrough'; state.currentView = 'review'; render(); });
    sc('Open Status report',         'Single-page snapshot', () => { reviewModeState.mode = 'full';        state.currentView = 'review'; render(); });
    sc('Toggle theme',          'Light / dark',      () => $('#btnTheme')?.click());
    sc('Toggle notes',          'Side panel',        () => $('#btnNotesToggle')?.click());
    sc('Run tour',              '5-step intro',      () => runTour(0));

    state.projects.forEach((proj) => {
      // Project itself
      out.push({ kind: 'project', label: proj.name, hint: 'Project', run: () => {
        state.currentProjectId = proj.id;
        state.currentView = 'board';
        saveState(); render();
      }});
      (proj.actions || []).forEach((a) => {
        if (a.deletedAt) return;
        const haystack = [a.title, personName(a.owner), a.notes || '', a.description || ''].join(' ');
        out.push({ kind: 'action', label: a.title, hint: `${proj.name} · ${personName(a.owner)} · ${a.due ? fmtDate(a.due) : 'no date'}`, hay: haystack, run: () => {
          state.currentProjectId = proj.id;
          state.currentView = 'board';
          render();
          setTimeout(() => openDrawer(a.id), 30);
        }});
      });
      (proj.openPoints || []).forEach((op) => {
        const hay = [op.title, op.notes || '', (op.steps || []).map((s) => s.text).join(' ')].join(' ');
        out.push({ kind: 'open-point', label: op.title, hint: `${proj.name} · open point`, hay, run: () => {
          state.currentProjectId = proj.id;
          state.currentView = 'openpoints';
          render();
        }});
      });
      (proj.changes || []).forEach((c) => {
        const hay = [c.title, c.rationale, c.analysis, c.description].join(' ');
        out.push({ kind: 'change', label: c.title, hint: `${proj.name} · ${c.status}`, hay, run: () => {
          state.currentProjectId = proj.id;
          state.currentView = 'changes';
          render();
          setTimeout(() => openChangeRequestEditor(c.id), 30);
        }});
      });
      (proj.decisions || []).forEach((d) => {
        out.push({ kind: 'decision', label: d.title, hint: `${proj.name} · decision · ${d.date || ''}`, hay: d.title + ' ' + (d.rationale || ''), run: () => {
          state.currentProjectId = proj.id;
          state.currentView = 'decisions';
          render();
        }});
      });
      (proj.risks || []).forEach((r) => {
        out.push({ kind: 'risk', label: r.title, hint: `${proj.name} · ${r.kind || 'risk'}`, hay: r.title + ' ' + (r.mitigation || ''), run: () => {
          state.currentProjectId = proj.id;
          state.currentView = 'risks';
          render();
          setTimeout(() => openRiskEditor(r.id), 30);
        }});
      });
      (proj.deliverables || []).forEach((d) => {
        out.push({ kind: 'deliverable', label: d.name, hint: `${proj.name} · deliverable · ${d.dueDate ? fmtDate(d.dueDate) : 'no date'}`, run: () => {
          state.currentProjectId = proj.id;
          state.currentView = 'deliverables';
          render();
        }});
      });
      (proj.milestones || []).forEach((m) => {
        out.push({ kind: 'milestone', label: m.name, hint: `${proj.name} · milestone · ${m.date ? fmtDate(m.date) : 'no date'}`, run: () => {
          state.currentProjectId = proj.id;
          state.currentView = 'milestones';
          render();
        }});
      });
      (proj.components || []).forEach((cmp) => {
        out.push({ kind: 'component', label: cmp.name, hint: `${proj.name} · component`, run: () => {
          state.currentProjectId = proj.id;
          state.currentView = 'components';
          render();
        }});
      });
      (proj.links || []).forEach((l) => {
        const hay = [l.title, l.description, l.url].join(' ');
        out.push({ kind: 'link', label: l.title || l.url, hint: `${proj.name} · link`, hay, run: () => {
          window.open(l.url, '_blank', 'noopener,noreferrer');
        }});
      });
      // Project notes — searchable as one big chunk
      const notesHTML = state.notes?.[proj.id] || '';
      if (notesHTML) {
        const text = notesHTML.replace(/<[^>]+>/g, ' ').replace(/&nbsp;/g, ' ').trim();
        if (text) out.push({ kind: 'notes', label: 'Meeting notes — ' + proj.name, hint: 'Project notes', hay: text, run: () => {
          state.currentProjectId = proj.id;
          if (!state.notesOpen) { state.notesOpen = true; saveState(); applyNotesPanel(); loadNotesForCurrentProject(); }
          render();
        }});
      }
    });
    state.people.forEach((p) => {
      out.push({ kind: 'person', label: p.name, hint: `${p.role || 'Person'} · ${p.capacity || 100}% FTE`, run: () => {
        state.currentView = 'people';
        render();
      }});
    });
    return out;
  }

  function paletteRank(items, query) {
    if (!query) {
      // Default — slash-commands first, then alphabetical
      return items.slice().sort((a, b) => {
        const sa = (a.kind === 'cmd' ? 1 : 0);
        const sb = (b.kind === 'cmd' ? 1 : 0);
        if (sa !== sb) return sb - sa;
        return a.label.localeCompare(b.label);
      }).slice(0, 60);
    }
    const isCmdOnly = query.startsWith('/');
    const q = (isCmdOnly ? query.slice(1) : query).trim();
    const scored = items
      .filter((it) => !isCmdOnly || it.kind === 'cmd')
      .map((it) => {
        const labelScore = fuzzyScore(q, it.label) * 3;
        const hayScore   = fuzzyScore(q, it.hay || '') * 1;
        const hintScore  = fuzzyScore(q, it.hint || '') * 0.5;
        const total = labelScore + hayScore + hintScore + (it.sortBoost || 0);
        return { it, total };
      })
      .filter((x) => x.total > 0)
      .sort((a, b) => b.total - a.total)
      .slice(0, 40)
      .map((x) => x.it);
    return scored;
  }

  function renderPalette() {
    const list = $('#paletteResults');
    if (!list) return;
    const { items, idx } = paletteState;
    if (!items.length) {
      list.innerHTML = '<div class="palette-empty">No matches. Try fewer characters or remove the leading <kbd>/</kbd>.</div>';
      return;
    }
    const kindIcon = {
      cmd: '⌘', action: '✓', 'open-point': '⚐', change: '⇆', decision: '⬡',
      risk: '△', deliverable: '◆', milestone: '◇', component: '▣',
      link: '↗', notes: '✎', project: '▦', person: '◔',
    };
    list.innerHTML = items.map((it, i) => `
      <button class="palette-item ${i === idx ? 'active' : ''}" data-idx="${i}" role="option">
        <span class="palette-kind">${kindIcon[it.kind] || '·'}</span>
        <span class="palette-label">${escapeHTML(it.label)}</span>
        <span class="palette-hint">${escapeHTML(it.hint || '')}</span>
      </button>`).join('');
    // Scroll active into view
    list.querySelector('.palette-item.active')?.scrollIntoView({ block: 'nearest' });
  }

  function openPalette() {
    const overlay = $('#paletteOverlay');
    if (!overlay) return;
    overlay.hidden = false;
    paletteState.items = paletteRank(buildPaletteIndex(), '');
    paletteState.idx = 0;
    paletteState.query = '';
    const inp = $('#paletteInput');
    inp.value = '';
    renderPalette();
    setTimeout(() => inp.focus(), 30);
  }
  function closePalette() {
    const overlay = $('#paletteOverlay');
    if (overlay) overlay.hidden = true;
    paletteState.items = [];
  }
  function paletteRunSelected() {
    const it = paletteState.items[paletteState.idx];
    if (!it) return;
    closePalette();
    setTimeout(() => { try { it.run(); } catch (e) { /* ignore */ } }, 0);
  }

  function wirePalette() {
    const overlay = $('#paletteOverlay');
    const inp = $('#paletteInput');
    const list = $('#paletteResults');
    if (!overlay || !inp) return;
    inp.addEventListener('input', () => {
      paletteState.query = inp.value;
      paletteState.items = paletteRank(buildPaletteIndex(), paletteState.query);
      paletteState.idx = 0;
      renderPalette();
    });
    inp.addEventListener('keydown', (e) => {
      if (e.key === 'ArrowDown') { e.preventDefault(); paletteState.idx = (paletteState.idx + 1) % Math.max(1, paletteState.items.length); renderPalette(); }
      else if (e.key === 'ArrowUp')   { e.preventDefault(); paletteState.idx = (paletteState.idx - 1 + paletteState.items.length) % Math.max(1, paletteState.items.length); renderPalette(); }
      else if (e.key === 'Enter')     { e.preventDefault(); paletteRunSelected(); }
      else if (e.key === 'Escape')    { e.preventDefault(); closePalette(); }
    });
    list.addEventListener('mousedown', (e) => {
      const btn = e.target.closest('.palette-item[data-idx]');
      if (!btn) return;
      e.preventDefault();
      paletteState.idx = parseInt(btn.dataset.idx, 10);
      paletteRunSelected();
    });
    list.addEventListener('mouseover', (e) => {
      const btn = e.target.closest('.palette-item[data-idx]');
      if (!btn) return;
      const i = parseInt(btn.dataset.idx, 10);
      if (i !== paletteState.idx) { paletteState.idx = i; renderPalette(); }
    });
    overlay.addEventListener('click', (e) => { if (e.target === overlay) closePalette(); });
    $('#paletteClose')?.addEventListener('click', closePalette);
  }

  /* --------------------- Phase B: sidebar groups + help ---------------- */
  function wireSidebarGroups() {
    document.querySelectorAll('.nav-section-toggle').forEach((toggle) => {
      const group = toggle.dataset.group;
      const body = document.querySelector(`[data-group-body="${group}"]`);
      if (!body) return;
      // Apply persisted state
      const open = state.settings.sidebarGroups?.[group] !== false;
      body.hidden = !open;
      toggle.setAttribute('aria-expanded', String(open));
      toggle.classList.toggle('collapsed', !open);
      toggle.addEventListener('click', () => {
        const isOpen = toggle.getAttribute('aria-expanded') === 'true';
        const next = !isOpen;
        body.hidden = !next;
        toggle.setAttribute('aria-expanded', String(next));
        toggle.classList.toggle('collapsed', !next);
        state.settings.sidebarGroups = state.settings.sidebarGroups || {};
        state.settings.sidebarGroups[group] = next;
        saveState();
      });
    });
  }

  // Keyboard cheatsheet
  function wireHelpModal() {
    const overlay = $('#helpOverlay');
    const btn = $('#btnHelp');
    if (!overlay || !btn) return;
    const close = () => { overlay.hidden = true; };
    btn.addEventListener('click', () => openHelp());
    $('#helpClose')?.addEventListener('click', close);
    $('#helpDone')?.addEventListener('click', close);
    // ? to open from anywhere outside an input
    document.addEventListener('keydown', (e) => {
      if (e.key !== '?' && !(e.shiftKey && e.key === '/')) return;
      const inField = ['INPUT','TEXTAREA','SELECT'].includes(e.target.tagName) || e.target.isContentEditable;
      if (inField) return;
      e.preventDefault();
      overlay.hidden ? openHelp() : close();
    });
  }
  function openHelp() {
    const overlay = $('#helpOverlay');
    const body = $('#helpBody');
    if (!overlay || !body) return;
    const isMac = navigator.platform.toLowerCase().includes('mac');
    const mod = isMac ? '⌘' : 'Ctrl';
    const sections = [
      ['Global', [
        [`${mod}+K`,             'Open command palette'],
        [`${mod}+Z`,             'Undo'],
        [`${mod}+Shift+Z`,       'Redo'],
        [`${mod}+\\`,            'Toggle meeting notes'],
        [`/`,                    'Focus search'],
        [`?`,                    'Open this cheatsheet'],
        [`← / →`,                'Calendar: previous / next month'],
      ]],
      ['Editing', [
        ['Right-click any row',  'Edit / status / delete menu'],
        ['Double-click any row', 'Open editor'],
        ['Drag the ⋮⋮ grip',     'Reorder rows'],
        ['Backspace on empty',   'Delete the current step / todo / row'],
        ['×',                    'Always closes the modal — click-outside is intentionally disabled'],
      ]],
      ['Notes', [
        ['@ then type',          'Mention a person'],
        ['# then type',          'Insert an existing action'],
        [`${mod}+Shift+A`,       'Create + insert a new action'],
        [`${mod}+B / I / U`,     'Bold / italic / underline'],
      ]],
      ['Lists & boards', [
        ['Drag card / row',      'Move within or across columns / folders'],
        ['Drop on Done bin',     'Mark done from the board'],
        ['Drop on Archive bin',  'Soft-delete'],
      ]],
    ];
    body.innerHTML = sections.map(([title, rows]) => `
      <div class="help-section">
        <div class="help-section-title">${escapeHTML(title)}</div>
        <div class="help-rows">
          ${rows.map(([k, v]) => `
            <div class="help-row"><kbd>${escapeHTML(k)}</kbd><span>${escapeHTML(v)}</span></div>
          `).join('')}
        </div>
      </div>`).join('') + `
      <div class="help-section">
        <div class="help-section-title">Re-run intro</div>
        <div class="help-row"><button class="ghost" id="btnHelpRunTour" style="padding:5px 10px;font-size:12px;">Run 5-step tour</button><span>Get the guided introduction again.</span></div>
      </div>`;
    overlay.hidden = false;
    $('#btnHelpRunTour')?.addEventListener('click', () => {
      $('#helpOverlay').hidden = true;
      runTour(0);
    });
  }

  /* ----------------------------- wire-up ----------------------------- */

  /* ----------------------- Phase L: first-run tour --------------------- */
  // 5-step Shepherd-style overlay. Pinned to the most stable selectors
  // available; if a target isn't on screen (rare), we still show the
  // step's body in a centered card. Sets state.settings.tourSeen on
  // skip / finish so it never auto-fires twice.
  const TOUR_STEPS = [
    {
      target: '#sidebar .nav, .sidebar nav',
      side: 'right',
      title: 'Navigate the project',
      body: 'Switch views from here — Board, Register, Timeline, Calendar, Reports, plus engineering side-views like Risks and Change Requests.',
    },
    {
      target: '#btnQuickAdd',
      side: 'bottom',
      title: 'One palette for everything',
      body: 'Press <b>⌘K</b> (or click here) to open the universal palette. Type to search across actions, links, decisions, risks — or run quick commands like <i>+ action</i>, <i>report</i>, <i>today</i>.',
    },
    {
      target: '.card[data-id]',
      side: 'right',
      title: 'Edit anywhere',
      body: 'Click any card to open its drawer. Drag between columns to change status. Hover to reveal a <b>⋯</b> menu with quick actions like Mark blocked / Add note / Archive.',
    },
    {
      target: '#btnInbox',
      side: 'bottom',
      title: 'Stay on top of what\'s slipping',
      body: 'The bell aggregates late actions, due-soon items, aging change requests and uncovered risks. Click an item to jump to it; dismissals stick.',
    },
    {
      target: '#btnHelp',
      side: 'bottom',
      title: 'Keyboard shortcuts',
      body: 'Press <b>?</b> any time for the full shortcut cheatsheet. You can re-run this tour from there.',
    },
  ];
  function maybeRunFirstRunTour() {
    if (state.settings?.tourSeen) return;
    setTimeout(() => runTour(0), 600); // wait for first render
  }
  function runTour(stepIdx) {
    closeTour();
    if (stepIdx < 0 || stepIdx >= TOUR_STEPS.length) {
      finishTour();
      return;
    }
    const step = TOUR_STEPS[stepIdx];
    const target = document.querySelector(step.target);
    const overlay = document.createElement('div');
    overlay.className = 'tour-overlay';
    overlay.innerHTML = `
      <div class="tour-mask" id="tourMask"></div>
      <div class="tour-card" id="tourCard">
        <div class="tour-step">Step ${stepIdx + 1} of ${TOUR_STEPS.length}</div>
        <div class="tour-title">${escapeHTML(step.title)}</div>
        <div class="tour-body">${step.body}</div>
        <div class="tour-foot">
          <button class="ghost" id="tourSkip">Skip tour</button>
          <button class="primary" id="tourNext">${stepIdx === TOUR_STEPS.length - 1 ? 'Finish' : 'Next →'}</button>
        </div>
      </div>`;
    document.body.appendChild(overlay);
    const card = overlay.querySelector('#tourCard');
    const mask = overlay.querySelector('#tourMask');
    if (target) {
      target.classList.add('tour-highlight');
      const r = target.getBoundingClientRect();
      // Soft-spot the highlighted region with a ring on the mask
      mask.style.setProperty('--spot-x', (r.left + r.width / 2) + 'px');
      mask.style.setProperty('--spot-y', (r.top + r.height / 2) + 'px');
      mask.style.setProperty('--spot-r', (Math.max(r.width, r.height) / 2 + 12) + 'px');
      // Place the card to the side requested by the step (clamp to viewport)
      const W = card.getBoundingClientRect().width || 320;
      const H = card.getBoundingClientRect().height || 160;
      const margin = 16;
      let x = r.right + margin, y = r.top;
      if (step.side === 'bottom') { x = Math.max(margin, r.left); y = r.bottom + margin; }
      if (step.side === 'left')   { x = Math.max(margin, r.left - W - margin); y = r.top; }
      if (step.side === 'right')  { x = r.right + margin; y = Math.max(margin, r.top); }
      x = clamp(x, margin, innerWidth - W - margin);
      y = clamp(y, margin, innerHeight - H - margin);
      card.style.left = x + 'px';
      card.style.top = y + 'px';
    } else {
      // No anchor — center the card and fade the mask uniformly
      mask.classList.add('full');
      card.style.left = '50%';
      card.style.top = '50%';
      card.style.transform = 'translate(-50%, -50%)';
    }
    overlay.querySelector('#tourSkip').addEventListener('click', () => { closeTour(); finishTour(); });
    overlay.querySelector('#tourNext').addEventListener('click', () => runTour(stepIdx + 1));
  }
  function closeTour() {
    document.querySelectorAll('.tour-highlight').forEach((el) => el.classList.remove('tour-highlight'));
    document.querySelectorAll('.tour-overlay').forEach((el) => el.remove());
  }
  function finishTour() {
    state.settings = state.settings || {};
    state.settings.tourSeen = true;
    saveState();
    toast('Tour finished — press ? for shortcuts');
  }

  function init() {
    state = loadState();
    // Run a single comprehensive backfill so every renderer gets the fields it
    // expects, regardless of whether the saved state predates the latest schema
    // additions (criticality, costCenters, links, changes/analysis, etc).
    try { normalizeState(state); } catch (e) { state = seedState(); }
    saveState();
    applyTheme(state.settings.theme || 'dark');

    $('#btnSidebarToggle').addEventListener('click', () => {
      $('#app').classList.toggle('sidebar-collapsed');
    });
    $$('.nav-item').forEach((b) => {
      b.addEventListener('click', () => {
        state.currentView = b.dataset.view;
        saveState(); render();
      });
    });
    $('#projectSelect').addEventListener('change', (e) => {
      state.currentProjectId = e.target.value;
      saveState(); render();
      if (state.notesOpen) loadNotesForCurrentProject();
    });
    $('#btnNewProject').addEventListener('click', () => openQuickAdd('project'));

    ['#search', '#filterOwner', '#filterComponent', '#filterStatus', '#filterDue'].forEach((sel) => {
      $(sel).addEventListener('input', () => render());
      $(sel).addEventListener('change', () => render());
    });

    $('#btnUndo').addEventListener('click', undo);
    $('#btnRedo').addEventListener('click', redo);
    $('#btnQuickAdd').addEventListener('click', () => openPalette());

    $('#btnNotesToggle').addEventListener('click', () => {
      state.notesOpen = !state.notesOpen;
      saveState();
      applyNotesPanel();
    });
    $('#btnInbox').addEventListener('click', () => {
      state.currentView = 'inbox';
      saveState(); render();
    });
    wireNotesPanel();
    applyNotesPanel();
    wireHoverDescOnce();
    wireEvmTooltipsOnce();
    wireTodoWidget();
    wireAutoBackup();
    initAutoBackup();
    wireSidebarGroups();
    wireHelpModal();
    wirePalette();

    $('#btnExport').addEventListener('click', exportJSON);
    $('#btnImport').addEventListener('click', importJSON);
    $('#btnMerge').addEventListener('click', openMergeOverlay);
    $('#btnReset').addEventListener('click', () => {
      if (!confirm('Replace the current data with the sample dataset?\n\nYou can Export first if you want to keep what\'s here.')) return;
      undoStack.push(JSON.stringify(state));
      state = seedState();
      saveState();
      render();
      toast('Sample data restored');
    });

    $('#btnEmpty').addEventListener('click', () => {
      if (!confirm('Wipe everything and start from zero?\n\nThis removes every project, person, action, deliverable, milestone, and note. You can Export first if you want to keep what\'s here.\n\nThis cannot be undone (except via Undo).')) return;
      const typed = prompt('Type EMPTY (in caps) to confirm wiping all data:');
      if ((typed || '').trim() !== 'EMPTY') { toast('Cancelled — nothing was changed'); return; }
      undoStack.push(JSON.stringify(state));
      state = emptyState();
      saveState();
      render();
      toast('Started fresh');
    });

    $('#btnTheme').addEventListener('click', () => {
      const next = (state.settings.theme === 'light') ? 'dark' : 'light';
      state.settings.theme = next;
      applyTheme(next);
      saveState();
    });

    // Quick add controls
    $$('.qa-tab').forEach((t) => t.addEventListener('click', () => openQuickAdd(t.dataset.qa)));
    $('#qaCancel').addEventListener('click', closeQuickAdd);
    // Quick Add closes only via Cancel / Save / × — backdrop clicks ignored.
    $('#qaSave').addEventListener('click', saveQA);
    // Backdrop click does NOT close Quick Add — user must use Cancel / Save / ×.

    // Drawer
    $('#drawerClose').addEventListener('click', closeDrawer);
    // Backdrop click does NOT close the action drawer — user must use × so
    // mid-edit fields aren't dismissed by an accidental click outside.

    // Inline edits in review
    document.addEventListener('change', (e) => {
      const t = e.target;
      if (!t.classList?.contains('inline')) return;
      const id = t.dataset.id;
      const action = t.dataset.action;
      const proj = curProject();
      const a = proj.actions.find((x) => x.id === id);
      if (!a) return;
      if (action === 'status') {
        a.history.push({ at: todayISO(), what: `Status: ${a.status} → ${t.value}` });
        a.status = t.value;
      } else if (action === 'owner') {
        a.history.push({ at: todayISO(), what: `Owner: ${personName(a.owner)} → ${personName(t.value)}` });
        a.owner = t.value;
      } else if (action === 'due') {
        a.history.push({ at: todayISO(), what: `Due: ${a.due || '—'} → ${t.value || '—'}` });
        a.due = t.value || null;
      }
      a.updatedAt = todayISO();
      commit('inline');
    });

    // Click into action title from review
    document.addEventListener('click', (e) => {
      const span = e.target.closest('.row .clickable');
      if (span && span.dataset.id) openDrawer(span.dataset.id);
    });

    // Close popovers on outside click (registered once)
    document.addEventListener('click', (e) => {
      const pop = $('#holidayPopover');
      if (!pop || pop.hidden) return;
      if (pop.contains(e.target)) return;
      if (e.target.closest('#btnTLHolidays')) return;
      pop.hidden = true;
    });

    // Keyboard shortcuts
    document.addEventListener('keydown', (e) => {
      // `inField` should also be true for contenteditable regions (notes
      // panel, open-point context, CR rich fields, action description editor)
      // — otherwise typing characters like `/` is swallowed by the global
      // search-focus shortcut, and Cmd+Z would undo app state instead of text.
      const inField = ['INPUT', 'TEXTAREA', 'SELECT'].includes(e.target.tagName) || e.target.isContentEditable;
      if ((e.metaKey || e.ctrlKey) && e.key.toLowerCase() === 'k') {
        e.preventDefault(); openPalette();
      } else if ((e.metaKey || e.ctrlKey) && e.key.toLowerCase() === 'z' && !e.shiftKey) {
        if (!inField) { e.preventDefault(); undo(); }
      } else if ((e.metaKey || e.ctrlKey) && e.key.toLowerCase() === 'z' && e.shiftKey) {
        if (!inField) { e.preventDefault(); redo(); }
      // Escape no longer closes Quick Add / drawer / editor modals — losing
      // an in-progress edit to a stray Escape was too easy. Use × or Cancel.
      } else if (e.key === '/' && !inField) {
        e.preventDefault(); $('#search').focus();
      } else if ((e.metaKey || e.ctrlKey) && e.key === '\\') {
        e.preventDefault();
        state.notesOpen = !state.notesOpen;
        saveState();
        applyNotesPanel();
      }
    });

    // Persistence safety net — flush on tab close / hide / refresh, and run a
    // low-cost periodic save in case a crash skips the unload handlers.
    window.addEventListener('beforeunload', flushPendingSaves);
    window.addEventListener('pagehide',     flushPendingSaves);
    document.addEventListener('visibilitychange', () => {
      if (document.visibilityState === 'hidden') flushPendingSaves();
    });
    setInterval(flushPendingSaves, 10000);

    // Phase D — refresh the bell once per minute so day-rollover late-actions
    // surface without needing a manual reload, and try the daily notification
    // once per app load.
    setInterval(refreshBell, 60000);
    setTimeout(maybeNotifyInbox, 1500);

    render();
    maybeRunFirstRunTour();
  }

  document.addEventListener('DOMContentLoaded', init);
})();
