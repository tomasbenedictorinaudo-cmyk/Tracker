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
    { id: 'todo', name: 'Not started', dot: 'todo' },
    { id: 'doing', name: 'In progress', dot: 'doing' },
    { id: 'blocked', name: 'Blocked', dot: 'blocked' },
    { id: 'done', name: 'Done', dot: 'done' },
  ];

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

    const people = [
      { id: 'p_sofia', name: 'Sofia Reyes',     role: 'Project Manager',     capacity: 6 },
      { id: 'p_marie', name: 'Marie Laurent',   role: 'Systems Engineer',    capacity: 5 },
      { id: 'p_arjun', name: 'Arjun Patel',     role: 'Avionics Lead',       capacity: 5 },
      { id: 'p_jonas', name: 'Jonas Becker',    role: 'Mechanical',          capacity: 4 },
      { id: 'p_kira',  name: 'Kira Nakamura',   role: 'Software Architect',  capacity: 5 },
      { id: 'p_omar',  name: 'Omar El-Sayed',   role: 'Power Systems',       capacity: 4 },
      { id: 'p_lena',  name: 'Lena Holmberg',   role: 'Thermal Engineer',    capacity: 4 },
      { id: 'p_diego', name: 'Diego Ferreira',  role: 'AOCS',                capacity: 5 },
      { id: 'p_yuki',  name: 'Yuki Tanaka',     role: 'Software Developer',  capacity: 5 },
      { id: 'p_nadia', name: 'Nadia Rahman',    role: 'Test Engineer',       capacity: 5 },
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
        { id: 'r_supply', title: 'Reaction wheel lead time slip',  probability: 3, impact: 4, mitigation: 'Dual-source supplier engaged.', owner: 'p_arjun' },
        { id: 'r_mass',   title: 'Mass margin trending under 5%',  probability: 4, impact: 3, mitigation: 'Lightweighting study + panel optimisation.', owner: 'p_jonas' },
        { id: 'r_power',  title: 'EOL power margin tight',         probability: 3, impact: 4, mitigation: 'Trade study on cell vendor.', owner: 'p_omar' },
        { id: 'r_thermal',title: 'Hot-case radiator under-sized',  probability: 2, impact: 4, mitigation: 'Adding louvres to baseline.', owner: 'p_lena' },
        { id: 'r_sw',     title: 'FSW timeline at risk',           probability: 3, impact: 3, mitigation: 'Early integration build, MIL-STD scrum.', owner: 'p_kira' },
        { id: 'r_test',   title: 'TVAC chamber availability',      probability: 4, impact: 3, mitigation: 'Booked alternate facility on standby.', owner: 'p_nadia' },
      ],
      decisions: [
        { id: 'dec_bus',    title: 'Down-select to BusFrame v3',   rationale: 'Best mass and thermal envelope after trade study.',    date: d(-220), owner: 'p_sofia' },
        { id: 'dec_rw',     title: 'Reaction wheel: Vendor Bravo', rationale: 'Lifetime + lead time vs Vendor Alpha.',                date: d(-160), owner: 'p_arjun' },
        { id: 'dec_battery',title: 'Li-ion 18650 cell — Vendor C', rationale: 'Heritage in similar LEO mission, 15-yr vendor support.', date: d(-95),  owner: 'p_omar' },
        { id: 'dec_optic',  title: 'Single-aperture optical bench', rationale: 'Mass and integration win vs dual-aperture option.',    date: d(-60),  owner: 'p_marie' },
        { id: 'dec_pdr',    title: 'PDR slipped 2 weeks',           rationale: 'Customer requested additional FDIR work; risk register updated.', date: d(-12), owner: 'p_sofia' },
      ],
      changes: [],
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
        { id: 'r_h_perf', title: 'Telemetry decoder throughput',  probability: 3, impact: 3, mitigation: 'Profile hot path, add backpressure.', owner: 'p_kira' },
        { id: 'r_h_ux',   title: 'Procedure editor UX scope',     probability: 3, impact: 2, mitigation: 'Two design rounds with ops users.',     owner: 'p_yuki' },
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
        { id: 'r_f_battery', title: 'Battery thermal runaway during fast charge', probability: 2, impact: 5, mitigation: 'Cell-level temperature monitoring; conservative charge profile.', owner: 'p_omar' },
        { id: 'r_f_field',   title: 'Outdoor test weather window',                probability: 3, impact: 2, mitigation: 'Two backup test windows scheduled.', owner: 'p_nadia' },
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

  /* ----------------------- selectors / helpers ----------------------- */

  function curProject() {
    return state.projects.find((p) => p.id === state.currentProjectId) || state.projects[0];
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
  function actionMatchesFilters(a) {
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
    const acts = proj.actions || [];
    const today = todayISO();
    const total = acts.length;
    const done = acts.filter((a) => a.status === 'done').length;
    const blocked = acts.filter((a) => a.status === 'blocked').length;
    const doing = acts.filter((a) => a.status === 'doing').length;
    const late = acts.filter((a) => a.due && a.status !== 'done' && dayDiff(a.due, today) < 0).length;
    const upcoming = acts.filter((a) => a.due && a.status !== 'done' && dayDiff(a.due, today) >= 0 && dayDiff(a.due, today) <= 7).length;
    const completionRate = total ? Math.round((done / total) * 100) : 0;
    const lateRate = total ? Math.round((late / total) * 100) : 0;
    const blockedRatio = total ? Math.round((blocked / total) * 100) : 0;
    // Throughput: items completed in last 14 days (using updatedAt for done)
    const since = fmtISO(new Date(Date.now() - 14 * dayMs));
    const throughput = acts.filter((a) => a.status === 'done' && a.updatedAt >= since).length;

    // Workload by person
    const workload = state.people.map((p) => {
      const open = acts.filter((a) => a.owner === p.id && a.status !== 'done').length;
      return { id: p.id, name: p.name, open, capacity: p.capacity || 5 };
    });

    return { total, done, doing, blocked, late, upcoming, completionRate, lateRate, blockedRatio, throughput, workload };
  }

  /* ---------------------------- rendering ---------------------------- */

  function render() {
    renderTopbar();
    renderSidebar();
    renderView();
  }

  function renderTopbar() {
    const sel = $('#projectSelect');
    sel.innerHTML = state.projects
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
      timeline: renderTimeline,
      dashboard: renderDashboard,
      charts: renderCharts,
      review: renderReview,
      components: renderComponents,
      deliverables: renderDeliverables,
      milestones: renderMilestones,
      risks: renderRisks,
      decisions: renderDecisions,
    };
    (fns[view] || renderBoard)(main);
  }

  /* ----------------------------- Board ------------------------------- */

  function renderBoard(root) {
    const proj = curProject();
    const view = document.createElement('div');
    view.className = 'view';

    const head = document.createElement('div');
    head.className = 'page-head';
    head.innerHTML = `
      <div>
        <div class="page-title">${escapeHTML(proj.name)}</div>
        <div class="page-sub">${(proj.actions || []).length} actions • ${proj.deliverables?.length || 0} deliverables • ${proj.milestones?.length || 0} milestones</div>
      </div>
      <div class="page-actions">
        <button class="ghost" id="btnAddAction">+ Action</button>
      </div>`;
    view.appendChild(head);

    const board = document.createElement('div');
    board.className = 'board';
    STATUSES.forEach((s) => {
      const items = (proj.actions || [])
        .filter((a) => a.status === s.id && actionMatchesFilters(a))
        .sort((a, b) => (a.priority || 0) - (b.priority || 0));
      const col = document.createElement('div');
      col.className = 'column';
      col.dataset.status = s.id;
      col.innerHTML = `
        <div class="col-head">
          <span class="col-dot ${s.dot}"></span>
          <span class="col-name">${s.name}</span>
          <span class="col-count">${items.length}</span>
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
      attachColumnDND(body);
      board.appendChild(col);
    });
    view.appendChild(board);
    root.appendChild(view);

    $('#btnAddAction').addEventListener('click', () => openQuickAdd('action'));
  }

  function makeCard(a) {
    const due = a.due;
    const dueClass = statusOfDue(due, a.status);
    const card = document.createElement('div');
    card.className = `card ${a.status === 'doing' ? 'doing' : ''} ${dueClass}`;
    card.draggable = true;
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
    card.innerHTML = `
      ${component ? `<div class="component-chip">${escapeHTML(component.name)}</div>` : ''}
      <div class="card-title">${escapeHTML(a.title)}</div>
      <div class="card-meta">
        <span class="avatar" title="${escapeHTML(owner?.name || 'Unassigned')}">${initials(owner?.name)}</span>
        <span class="due ${dueClass}">${due ? fmtDate(due) : 'no date'}</span>
        ${a.notes ? '<span class="tag">note</span>' : ''}
      </div>`;
    card.addEventListener('click', (e) => {
      if (e.detail === 1 && !card.classList.contains('dragging')) openDrawer(a.id);
    });
    card.addEventListener('dragstart', (e) => {
      card.classList.add('dragging');
      e.dataTransfer.setData('text/cockpit-action', a.id);
      e.dataTransfer.effectAllowed = 'move';
    });
    card.addEventListener('dragend', () => card.classList.remove('dragging'));
    return card;
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

  /* ---------------------------- Register ----------------------------- */

  // Persistent sort state (so it survives navigation away and back)
  const regState = { sortBy: 'due', sortDir: 'asc' };

  function regSortValue(a, col, proj) {
    switch (col) {
      case 'title':     return (a.title || '').toLowerCase();
      case 'component': return (findComponent(proj, a.component)?.name || 'zzz').toLowerCase();
      case 'owner':     return personName(a.owner).toLowerCase();
      case 'status':    return ['todo','doing','blocked','done'].indexOf(a.status);
      case 'due':       return a.due || '9999-99-99';
      case 'updatedAt': return a.updatedAt || '0000-00-00';
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
          <div class="page-sub">Flat view of all actions. Click a column to sort, click a row to open details.</div>
        </div>
        <div class="page-actions"><button class="ghost" id="btnAddAction">+ Action</button></div>
      </div>
      <div class="register">
        <div class="reg-head">
          <button class="reg-col" data-col="title">Title</button>
          <button class="reg-col" data-col="component">Component</button>
          <button class="reg-col" data-col="owner">Owner</button>
          <button class="reg-col" data-col="status">Status</button>
          <button class="reg-col" data-col="due">Due</button>
          <button class="reg-col" data-col="updatedAt">Updated</button>
        </div>
        <div class="reg-body" id="regBody"></div>
      </div>`;
    root.appendChild(view);

    function draw() {
      const acts = (proj.actions || []).filter(actionMatchesFilters).slice();
      acts.sort((a, b) => {
        const av = regSortValue(a, regState.sortBy, proj);
        const bv = regSortValue(b, regState.sortBy, proj);
        if (av < bv) return regState.sortDir === 'asc' ? -1 : 1;
        if (av > bv) return regState.sortDir === 'asc' ? 1 : -1;
        return 0;
      });
      const body = $('#regBody');
      if (!acts.length) {
        body.innerHTML = '<div class="empty">No actions match the current filters.</div>';
      } else {
        body.innerHTML = acts.map((a) => {
          const cmp = findComponent(proj, a.component);
          const c = cmp ? componentColor(cmp.color) : null;
          const dueCls = statusOfDue(a.due, a.status);
          const stat = STATUSES.find((s) => s.id === a.status);
          return `
            <div class="reg-row" data-id="${a.id}">
              <div class="reg-cell title">${escapeHTML(a.title)}${a.notes ? ' <span class="tag">note</span>' : ''}</div>
              <div class="reg-cell">${cmp ? `<span class="component-chip" style="background:rgba(${c.rgb},.2);color:rgb(${c.rgb})">${escapeHTML(cmp.name)}</span>` : '<span class="muted">—</span>'}</div>
              <div class="reg-cell"><span class="avatar">${initials(personName(a.owner))}</span><span class="ow-name">${escapeHTML(personName(a.owner))}</span></div>
              <div class="reg-cell"><span class="col-dot ${stat?.dot}"></span> ${stat?.name}</div>
              <div class="reg-cell due ${dueCls}">${a.due ? fmtDate(a.due) : '—'}</div>
              <div class="reg-cell muted">${a.updatedAt ? fmtDate(a.updatedAt) : '—'}</div>
            </div>`;
        }).join('');
        $$('.reg-row', body).forEach((row) =>
          row.addEventListener('click', () => openDrawer(row.dataset.id)));
      }
      $$('.reg-col', view).forEach((b) => b.classList.remove('asc', 'desc'));
      const active = view.querySelector(`.reg-col[data-col="${regState.sortBy}"]`);
      if (active) active.classList.add(regState.sortDir);
    }

    $$('.reg-col', view).forEach((btn) => {
      btn.addEventListener('click', () => {
        const col = btn.dataset.col;
        if (regState.sortBy === col) regState.sortDir = regState.sortDir === 'asc' ? 'desc' : 'asc';
        else { regState.sortBy = col; regState.sortDir = 'asc'; }
        draw();
      });
    });
    $('#btnAddAction').addEventListener('click', () => openQuickAdd('action'));
    draw();
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

    // Build lanes by person
    const people = state.people;
    lanesEl.innerHTML = `<div class="tl-lane header">Owner</div>` +
      people.map((p) => `<div class="tl-lane" data-owner="${p.id}">${escapeHTML(p.name)}</div>`).join('');

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

    // Milestones as diamonds spanning all rows
    const proj = curProject();
    (proj.milestones || []).forEach((m) => {
      if (!m.date) return;
      const offset = Math.round((parseDate(m.date) - start) / dayMs);
      if (offset < 0 || offset > totalDays) return;
      const ms = document.createElement('div');
      ms.className = 'tl-milestone';
      ms.style.left = (offset * dw) + 'px';
      ms.title = `${m.name} — ${fmtFull(m.date)}`;
      ms.innerHTML = `<span>◇ ${escapeHTML(m.name)}</span>`;
      gridEl.appendChild(ms);
    });

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
      bar.className = `tl-bar ${a.status} ${dueCls === 'late' && a.status !== 'done' ? 'late' : ''} ${onHoliday ? 'holiday-conflict' : ''} ${overload ? 'over-allocated' : ''}`;
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
          history: [{ at: todayISO(), what: 'Created from timeline' }],
        };
        proj.actions.push(a);
        commit('create');
        toast('Action created — double-click bar to edit');
      });
    });
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

  function renderDashboard(root) {
    const proj = curProject();
    const k = kpis();
    const view = document.createElement('div');
    view.className = 'view';
    view.innerHTML = `
      <div class="page-head">
        <div>
          <div class="page-title">${escapeHTML(proj.name)} — Dashboard</div>
          <div class="page-sub">Health summary based on current data</div>
        </div>
      </div>
      <div class="dashboard">
        <div class="kpi">
          <div class="kpi-label">Late items</div>
          <div class="kpi-value ${k.late > 0 ? 'bad' : 'ok'}">${k.late}</div>
          <div class="kpi-sub">${k.lateRate}% of all actions</div>
        </div>
        <div class="kpi">
          <div class="kpi-label">Blocked</div>
          <div class="kpi-value ${k.blocked > 0 ? 'warn' : 'ok'}">${k.blocked}</div>
          <div class="kpi-sub">${k.blockedRatio}% of all actions</div>
        </div>
        <div class="kpi">
          <div class="kpi-label">Due ≤ 7 days</div>
          <div class="kpi-value ${k.upcoming > 4 ? 'warn' : ''}">${k.upcoming}</div>
          <div class="kpi-sub">Upcoming workload</div>
        </div>
        <div class="kpi">
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
      </div>`;
    root.appendChild(view);

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

    const capLine = visibleCap > 0
      ? `<line class="chart-cap" x1="${padL}" x2="${W - padR}" y1="${yFor(visibleCap)}" y2="${yFor(visibleCap)}" />
         <text class="chart-label" x="${W - padR - 4}" y="${Math.max(11, yFor(visibleCap) - 3)}" text-anchor="end">cap ${visibleCap}</text>`
      : '';

    return `
      <svg viewBox="0 0 ${W} ${H}" class="chart-svg" preserveAspectRatio="xMidYMid meet">
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
    const candidates = acts.filter(({ a }) => a.due).map(({ proj, a }) => ({
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

  function renderReview(root) {
    const proj = curProject();
    const view = document.createElement('div');
    view.className = 'view';
    view.innerHTML = `
      <div class="review">
        <div class="page-head" style="margin-bottom:0;">
          <div>
            <div class="page-title">${escapeHTML(proj.name)} — Review</div>
            <div class="page-sub">${fmtFull(todayISO())}</div>
          </div>
          <div class="page-actions">
            <button class="ghost" id="btnReviewExport">Export HTML</button>
          </div>
        </div>
        <div class="review-stepper" id="reviewStepper"></div>
        <div class="review-card" id="reviewBody"></div>
        <div class="review-foot">
          <button class="ghost" id="btnReviewPrev">← Previous</button>
          <button class="primary" id="btnReviewNext">Next →</button>
        </div>
      </div>`;
    root.appendChild(view);

    drawReviewStep();
    $('#btnReviewPrev').addEventListener('click', () => {
      reviewStep = clamp(reviewStep - 1, 0, REVIEW_STEPS.length - 1);
      drawReviewStep();
    });
    $('#btnReviewNext').addEventListener('click', () => {
      reviewStep = clamp(reviewStep + 1, 0, REVIEW_STEPS.length - 1);
      drawReviewStep();
    });
    $('#btnReviewExport').addEventListener('click', exportReviewHTML);
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

  function renderDeliverables(root) {
    const proj = curProject();
    const view = document.createElement('div');
    view.className = 'view';
    view.innerHTML = `
      <div class="page-head">
        <div><div class="page-title">Deliverables</div><div class="page-sub">Optional — group actions under a deliverable.</div></div>
        <div class="page-actions"><button class="ghost" id="btnAddDel">+ Deliverable</button></div>
      </div>
      <div class="row-list" id="delList"></div>`;
    root.appendChild(view);
    const list = $('#delList');
    if (!proj.deliverables?.length) list.innerHTML = '<div class="empty">No deliverables yet.</div>';
    else {
      list.innerHTML = proj.deliverables.map((d) => `
        <div class="row">
          <span>◆ ${escapeHTML(d.name)}</span>
          <span class="row-meta">${d.dueDate || '—'} • ${escapeHTML(d.status || 'todo')}</span>
        </div>`).join('');
    }
    $('#btnAddDel').addEventListener('click', () => openQuickAdd('deliverable'));
  }

  function renderMilestones(root) {
    const proj = curProject();
    const view = document.createElement('div');
    view.className = 'view';
    view.innerHTML = `
      <div class="page-head">
        <div><div class="page-title">Milestones</div><div class="page-sub">Optional — anchor key dates.</div></div>
        <div class="page-actions"><button class="ghost" id="btnAddMile">+ Milestone</button></div>
      </div>
      <div class="row-list" id="mileList"></div>`;
    root.appendChild(view);
    const list = $('#mileList');
    if (!proj.milestones?.length) list.innerHTML = '<div class="empty">No milestones yet.</div>';
    else {
      list.innerHTML = proj.milestones.map((m) => `
        <div class="row">
          <span>◇ ${escapeHTML(m.name)}</span>
          <span class="row-meta">${m.date || '—'}</span>
        </div>`).join('');
    }
    $('#btnAddMile').addEventListener('click', () => openQuickAdd('milestone'));
  }

  function renderRisks(root) {
    const proj = curProject();
    const view = document.createElement('div');
    view.className = 'view';
    view.innerHTML = `
      <div class="page-head">
        <div><div class="page-title">Risks</div><div class="page-sub">Probability × Impact — both 1 (low) to 5 (high).</div></div>
        <div class="page-actions"><button class="ghost" id="btnAddRisk">+ Risk</button></div>
      </div>
      <div class="row-list" id="riskList"></div>`;
    root.appendChild(view);
    const list = $('#riskList');
    if (!proj.risks?.length) list.innerHTML = '<div class="empty">No risks logged.</div>';
    else {
      list.innerHTML = proj.risks.map((r) => {
        const score = (r.probability || 0) * (r.impact || 0);
        const cls = score >= 12 ? 'late' : score >= 6 ? 'soon' : '';
        return `
          <div class="row ${cls}">
            <span>△ ${escapeHTML(r.title)}</span>
            <span class="row-meta">P${r.probability}×I${r.impact} = ${score} • ${escapeHTML(personName(r.owner))}</span>
          </div>`;
      }).join('');
    }
    $('#btnAddRisk').addEventListener('click', () => openQuickAdd('risk'));
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
        <div class="row">
          <span>⬡ ${escapeHTML(d.title)}</span>
          <span class="row-meta">${escapeHTML(personName(d.owner))} • ${d.date || ''}</span>
        </div>`).join('');
    }
    $('#btnAddDec').addEventListener('click', () => openQuickAdd('decision'));
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
      list.innerHTML = proj.components.map((pt) => {
        const c = componentColor(pt.color);
        const count = (proj.actions || []).filter((a) => a.component === pt.id).length;
        return `
          <div class="row" data-component-id="${pt.id}">
            <span class="component-swatch" style="background: rgba(${c.rgb},.9);"></span>
            <input class="inline component-name" value="${escapeHTML(pt.name)}" />
            <select class="inline component-color">
              ${COMPONENT_COLORS.map((co) => `<option value="${co.id}" ${co.id === pt.color ? 'selected' : ''}>${co.name}</option>`).join('')}
            </select>
            <span class="row-meta">${count} action${count === 1 ? '' : 's'}</span>
            <button class="icon-btn component-del" title="Delete">×</button>
          </div>`;
      }).join('');
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

  /* -------------------------- Portfolio / People --------------------- */

  function renderPortfolio(root) {
    const view = document.createElement('div');
    view.className = 'view';
    view.innerHTML = `
      <div class="page-head">
        <div><div class="page-title">Portfolio</div><div class="page-sub">${state.projects.length} projects</div></div>
        <div class="page-actions"><button class="ghost" id="btnNewProj2">+ Project</button></div>
      </div>
      <div class="dashboard">
        ${state.projects.map((p) => {
          const acts = p.actions || [];
          const total = acts.length;
          const done = acts.filter((a) => a.status === 'done').length;
          const late = acts.filter((a) => a.status !== 'done' && a.due && dayDiff(a.due, todayISO()) < 0).length;
          const pct = total ? Math.round((done / total) * 100) : 0;
          return `
            <div class="kpi clickable" data-pid="${p.id}" style="grid-column: span 2; cursor:pointer;">
              <div class="kpi-label">${escapeHTML(p.name)}</div>
              <div class="kpi-value">${pct}%</div>
              <div class="kpi-sub">${total} actions • ${late} late • ${done} done</div>
            </div>`;
        }).join('')}
      </div>`;
    root.appendChild(view);
    $$('.kpi.clickable', view).forEach((el) => {
      el.addEventListener('click', () => {
        state.currentProjectId = el.dataset.pid;
        state.currentView = 'board';
        saveState(); render();
      });
    });
    $('#btnNewProj2').addEventListener('click', () => openQuickAdd('project'));
  }

  // Compute weekly workload for a person across the next `weeks` weeks.
  // Returns array of { weekStart: Date, count: number, items: action[] }.
  function weeklyLoad(personId, weeks = 12) {
    const today = new Date(); today.setHours(0, 0, 0, 0);
    let weekStart = new Date(today);
    while (weekStart.getDay() !== 1) weekStart = new Date(weekStart.getTime() - dayMs); // back to Monday
    const out = [];
    for (let w = 0; w < weeks; w++) {
      const wStart = new Date(weekStart.getTime() + w * 7 * dayMs);
      const wEnd = new Date(wStart.getTime() + 6 * dayMs);
      const items = [];
      state.projects.forEach((proj) => {
        (proj.actions || []).forEach((a) => {
          if (a.owner !== personId || a.status === 'done') return;
          if (!a.due) return;
          const due = parseDate(a.due);
          const start = a.startDate ? parseDate(a.startDate) :
            new Date(due.getTime() - 2 * dayMs); // default 3-day window
          // Overlap test: action [start..due] intersects [wStart..wEnd]
          if (start <= wEnd && due >= wStart) items.push({ a, proj });
        });
      });
      out.push({ weekStart: wStart, count: items.length, items });
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
      return `<rect class="spark-bar ${cls}" x="${x + 2}" y="${y}" width="${Math.max(2, barW - 4)}" height="${Math.max(0, h)}" rx="2">
        <title>Week of ${label} — ${s.count} action${s.count === 1 ? '' : 's'} (cap ${cap})</title>
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
        <text class="spark-cap-label" x="${W - padR - 4}" y="${Math.max(10, capY - 3)}" text-anchor="end">cap ${cap}</text>
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
        <div><div class="page-title">People</div><div class="page-sub">${state.people.length} members • workload across all projects, projected over the next 12 weeks</div></div>
        <div class="page-actions"><button class="ghost" id="btnNewPerson">+ Person</button></div>
      </div>
      <div class="panel">
        <div class="panel-title">
          Workload
          <span class="legend">
            <span class="legend-item"><span class="dot ok"></span>≤80% cap</span>
            <span class="legend-item"><span class="dot warn"></span>≤100% cap</span>
            <span class="legend-item"><span class="dot bad"></span>over cap</span>
          </span>
        </div>
        <div id="peopleWl">
          ${wl.map(({ p, open, series, peakWeek }) => {
            const cap = p.capacity || 5;
            const pct = clamp(Math.round((open / cap) * 100), 0, 200);
            const cls = pct > 100 ? 'over' : pct > 80 ? 'warn' : 'ok';
            const peakCls = peakWeek.count > cap ? 'over' : peakWeek.count > cap * 0.8 ? 'warn' : 'ok';
            return `
              <div class="person-row">
                <div class="name-cell">
                  <span class="avatar">${initials(p.name)}</span>
                  <span class="who">
                    <b>${escapeHTML(p.name)}</b>
                    <span>${escapeHTML(p.role || '')}</span>
                  </span>
                </div>
                <div class="now-load">
                  <div class="bar-track"><div class="bar-fill ${cls}" style="width:${Math.min(100, pct)}%"></div></div>
                  <div class="bar-val">${open}/${cap}</div>
                </div>
                <div class="spark-wrap">
                  ${workloadSparkSVG(p, series)}
                  <div class="spark-meta"><span class="${peakCls}">peak ${peakWeek.count}</span></div>
                </div>
              </div>`;
          }).join('')}
        </div>
      </div>`;
    root.appendChild(view);
    $('#btnNewPerson').addEventListener('click', () => openQuickAdd('person'));
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
      <div class="field"><label>Owner</label>
        <select id="dOwner">${state.people.map((p) => `<option value="${p.id}" ${p.id === a.owner ? 'selected' : ''}>${escapeHTML(p.name)}</option>`).join('')}</select>
      </div>
      <div style="display:grid; grid-template-columns:1fr 1fr; gap:10px;">
        <div class="field"><label>Status</label>
          <select id="dStatus">${STATUSES.map((s) => `<option value="${s.id}" ${s.id === a.status ? 'selected' : ''}>${s.name}</option>`).join('')}</select>
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
      <div class="field"><label>Notes / justification</label><textarea id="dNotes">${escapeHTML(a.notes || '')}</textarea></div>
      <div class="field"><label>History</label>
        <div class="history">${(a.history || []).slice(-10).reverse().map((h) => `<div class="history-item"><b>${h.at}</b> — ${escapeHTML(h.what)}</div>`).join('')}</div>
      </div>
      <div style="display:flex; gap:8px; margin-top:6px;">
        <button class="primary" id="dSave">Save</button>
        <button class="ghost" id="dDelete" style="margin-left:auto; color:var(--bad);">Delete</button>
      </div>`;
    $('#drawer').hidden = false;
    $('#dCmt').addEventListener('input', (e) => { $('#dCmtVal').textContent = e.target.value + '%'; });
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
      a.commitment = clamp(parseInt($('#dCmt').value, 10) || 100, 5, 100);
      if (oldCmt !== a.commitment) a.history.push({ at: todayISO(), what: `Commitment: ${oldCmt}% → ${a.commitment}%` });
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
      if (!confirm('Delete this action?')) return;
      proj.actions = proj.actions.filter((x) => x.id !== a.id);
      commit('delete');
      closeDrawer();
      toast('Deleted');
    });
  }
  function closeDrawer() { $('#drawer').hidden = true; }

  /* --------------------------- Quick add ----------------------------- */

  let qaType = 'action';
  function openQuickAdd(type = 'action') {
    qaType = type;
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
      body.innerHTML = `
        <div class="field"><label>Title</label><input id="qTitle" placeholder="What needs to be done?" /></div>
        <div class="qa-row">
          <div class="field"><label>Owner</label>
            <select id="qOwner">${state.people.map((p) => `<option value="${p.id}">${escapeHTML(p.name)}</option>`).join('')}</select>
          </div>
          <div class="field"><label>Due</label><input id="qDue" type="date" value="${todayISO()}" /></div>
        </div>
        <div class="qa-row">
          <div class="field"><label>Status</label>
            <select id="qStatus">${STATUSES.map((s) => `<option value="${s.id}">${s.name}</option>`).join('')}</select>
          </div>
          <div class="field"><label>Component (optional)</label>
            <select id="qComponent"><option value="">—</option>${(proj.components || []).map((pt) => `<option value="${pt.id}">${escapeHTML(pt.name)}</option>`).join('')}</select>
          </div>
        </div>
        <div class="field">
          <label>Commitment <span class="muted" id="qCmtVal">100%</span></label>
          <input id="qCmt" type="range" min="5" max="100" step="5" value="100" oninput="document.getElementById('qCmtVal').textContent = this.value + '%';" />
        </div>
        <div class="field"><label>Deliverable (optional)</label>
          <select id="qDel"><option value="">—</option>${(proj.deliverables || []).map((d) => `<option value="${d.id}">${escapeHTML(d.name)}</option>`).join('')}</select>
        </div>
        <div class="field"><label>Notes</label><textarea id="qNotes" placeholder="Optional context"></textarea></div>`;
    } else if (qaType === "component") {
      body.innerHTML = `
        <div class="field"><label>Name</label><input id="qName" placeholder="e.g. Power, AOCS, Backend…" /></div>
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
      body.innerHTML = `
        <div class="field"><label>Title</label><input id="qTitle" /></div>
        <div class="qa-row">
          <div class="field"><label>Probability (1-5)</label><input id="qProb" type="number" min="1" max="5" value="3" /></div>
          <div class="field"><label>Impact (1-5)</label><input id="qImp" type="number" min="1" max="5" value="3" /></div>
        </div>
        <div class="field"><label>Owner</label>
          <select id="qOwner">${state.people.map((p) => `<option value="${p.id}">${escapeHTML(p.name)}</option>`).join('')}</select>
        </div>
        <div class="field"><label>Mitigation</label><textarea id="qMit"></textarea></div>`;
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
    } else if (qaType === 'person') {
      body.innerHTML = `
        <div class="field"><label>Name</label><input id="qName" /></div>
        <div class="qa-row">
          <div class="field"><label>Role</label><input id="qRole" /></div>
          <div class="field"><label>Capacity</label><input id="qCap" type="number" min="1" max="20" value="5" /></div>
        </div>`;
    } else if (qaType === 'project') {
      body.innerHTML = `
        <div class="field"><label>Name</label><input id="qName" /></div>
        <div class="field"><label>Description</label><textarea id="qDesc"></textarea></div>`;
    }
  }

  function saveQA() {
    const proj = curProject();
    if (qaType === 'action') {
      const title = $('#qTitle').value.trim();
      if (!title) return toast('Title required');
      const a = {
        id: uid('a'), title,
        owner: $('#qOwner').value,
        due: $('#qDue').value || null,
        status: $('#qStatus').value,
        priority: 0,
        commitment: clamp(parseInt($('#qCmt')?.value, 10) || 100, 5, 100),
        component: $('#qComponent')?.value || null,
        deliverable: $('#qDel').value || null,
        milestone: null,
        notes: $('#qNotes').value || '',
        createdAt: todayISO(), updatedAt: todayISO(),
        history: [{ at: todayISO(), what: 'Created' }],
      };
      proj.actions.push(a);
    } else if (qaType === 'component') {
      const name = $('#qName').value.trim();
      if (!name) return toast('Name required');
      const color = (document.querySelector('input[name="qComponentColor"]:checked')?.value) || COMPONENT_COLORS[0].id;
      proj.components = proj.components || [];
      proj.components.push({ id: uid('cm'), name, color });
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
      proj.risks = proj.risks || [];
      proj.risks.push({
        id: uid('r'), title,
        probability: clamp(parseInt($('#qProb').value, 10) || 3, 1, 5),
        impact: clamp(parseInt($('#qImp').value, 10) || 3, 1, 5),
        mitigation: $('#qMit').value || '',
        owner: $('#qOwner').value,
      });
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
        actions: [], deliverables: [], milestones: [], risks: [], decisions: [], changes: [], components: [],
      };
      state.projects.push(np);
      state.currentProjectId = np.id;
    }
    commit('add');
    closeQuickAdd();
    toast('Added');
  }

  /* --------------------------- Import/Export ------------------------- */

  function exportJSON() {
    const blob = new Blob([JSON.stringify(state, null, 2)], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url; a.download = `cockpit-${todayISO()}.json`;
    a.click();
    URL.revokeObjectURL(url);
    toast('Exported');
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
          if (!obj.projects || !obj.people) throw new Error('Invalid file');
          // Tolerate missing fields
          obj.projects.forEach((p) => {
            p.actions = p.actions || [];
            p.deliverables = p.deliverables || [];
            p.milestones = p.milestones || [];
            p.risks = p.risks || [];
            p.decisions = p.decisions || [];
            p.components = p.components || [];
            p.actions.forEach((a) => {
              a.history = a.history || [];
              a.createdAt = a.createdAt || todayISO();
              a.updatedAt = a.updatedAt || todayISO();
              a.priority = a.priority || 0;
            });
          });
          undoStack.push(JSON.stringify(state));
          state = obj;
          if (!state.currentProjectId) state.currentProjectId = state.projects[0]?.id;
          state.currentView = state.currentView || 'board';
          saveState(); render();
          toast('Imported');
        } catch (e) { toast('Import failed: ' + e.message); }
      };
      reader.readAsText(file);
    });
    input.click();
  }

  /* ----------------------------- wire-up ----------------------------- */

  function init() {
    state = loadState();
    if (state.currentView === 'teams') state.currentView = 'people';
    state.settings = state.settings || {};
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
    });
    $('#btnNewProject').addEventListener('click', () => openQuickAdd('project'));

    ['#search', '#filterOwner', '#filterComponent', '#filterStatus', '#filterDue'].forEach((sel) => {
      $(sel).addEventListener('input', () => render());
      $(sel).addEventListener('change', () => render());
    });

    $('#btnUndo').addEventListener('click', undo);
    $('#btnRedo').addEventListener('click', redo);
    $('#btnQuickAdd').addEventListener('click', () => openQuickAdd('action'));

    $('#btnExport').addEventListener('click', exportJSON);
    $('#btnImport').addEventListener('click', importJSON);
    $('#btnReset').addEventListener('click', () => {
      if (!confirm('Replace the current data with the sample dataset?\n\nYou can Export first if you want to keep what\'s here.')) return;
      undoStack.push(JSON.stringify(state));
      state = seedState();
      saveState();
      render();
      toast('Sample data restored');
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
    $('#qaSave').addEventListener('click', saveQA);
    $('#quickAdd').addEventListener('click', (e) => { if (e.target.id === 'quickAdd') closeQuickAdd(); });

    // Drawer
    $('#drawerClose').addEventListener('click', closeDrawer);
    $('#drawer').addEventListener('click', (e) => { if (e.target.id === 'drawer') closeDrawer(); });

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
      const inField = ['INPUT', 'TEXTAREA', 'SELECT'].includes(e.target.tagName);
      if ((e.metaKey || e.ctrlKey) && e.key.toLowerCase() === 'k') {
        e.preventDefault(); openQuickAdd('action');
      } else if ((e.metaKey || e.ctrlKey) && e.key.toLowerCase() === 'z' && !e.shiftKey) {
        if (!inField) { e.preventDefault(); undo(); }
      } else if ((e.metaKey || e.ctrlKey) && e.key.toLowerCase() === 'z' && e.shiftKey) {
        if (!inField) { e.preventDefault(); redo(); }
      } else if (e.key === 'Escape') {
        if (!$('#quickAdd').hidden) closeQuickAdd();
        else if (!$('#drawer').hidden) closeDrawer();
      } else if (e.key === '/' && !inField) {
        e.preventDefault(); $('#search').focus();
      }
    });

    render();
  }

  document.addEventListener('DOMContentLoaded', init);
})();
