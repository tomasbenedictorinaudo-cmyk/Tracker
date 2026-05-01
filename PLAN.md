# Cockpit `next-major` — coordinated roadmap

This branch carries the full set of recommendations from the three-lens
assessment. The branch must remain a working app at every commit and never
break a previously exported JSON. Schema additions are **always** retrocompatible
via `normalizeState`.

## Constraints (carry through every phase)

- **Stay airy.** New surfaces hide their controls until hover; chips show only
  when their value differs from the default. No information without a question.
- **Stay intuitive.** Every new affordance must have a visible cue (icon,
  hover hint, or tooltip). Keyboard shortcut must be paired with a click path.
- **Stay retrocompat.** Every schema addition gets a default in
  `normalizeState`. Older JSON exports import unchanged.
- **One commit per phase** (multiple if a phase grows large), each commit
  individually testable.

## Phases

### A · Foundation (schema + helpers)
Everything subsequent depends on these defaults landing first.
- New per-action fields: `dependsOn[]`, `comments[]`, `tags[]`, `authorBy`,
  `authorAt` (signed metadata for team-mode merges later).
- New per-open-point and per-CR fields: `tags[]`, `comments[]`.
- New per-project fields: `templateOf` (id of source template, optional).
- New state-level: `state.templates[]`, `state.inbox.dismissed[]`,
  `state.settings.tourSeen`, `state.settings.notifyEnabled`.
- Helpers: `mutate(scope, fn, commitName)`, `fuzzyScore(query, hay)`,
  `markdown.escape(s)` (for report export).

### B · UX standardization
Discoverability without crowding.
- Single shared drag-grip style (always-visible at .35 opacity).
- `…` overflow button per row that opens the same menu as right-click. Hover-
  visible only — keeps the rest density-neutral.
- Sidebar grouping: collapsible Workspace / Project / Engineering sections.
- Modal consistency rule: transactional (data-loss risky) close only via
  ✕ / Cancel; transient popups close on outside-click.
- Floating `?` button → keyboard shortcut cheatsheet.

### C · Universal Cmd+K palette
- Replace topbar quick-search with a proper palette.
- Indexes title + rich text across actions, open points, CRs, decisions,
  links, risks, components, deliverables, milestones, projects, people, notes
  (project-keyed).
- Up/Down navigates, Enter opens / drawer / file. First few groups are
  "actions you can take" (slash commands) — `+ action`, `+ link`, `+ folder`,
  `report`, `today`, `inbox`.
- Cmd+K rebound from "open Quick Add" to "open palette". Quick Add stays
  reachable via `+ Quick add` button and as a palette command.

### D · Reminders + Inbox
- Topbar bell icon with a count badge.
- Inbox view aggregates: my late actions, my actions due ≤ 3 d, CRs pending
  > 14 d, stale actions ≥ 14 d, todos due, R&O residual ≥ threshold.
- `Notification.requestPermission()` — daily check on first load of the day,
  user can dismiss items, dismissals persist in `state.inbox.dismissed`.
- Bell shows red when items present, neutral when zero.

### E · Status Report export
- New "Report" view in the sidebar. Period selector (last 7 d / last 30 d /
  custom). Renders an HTML preview using existing data:
  - KPIs row (mirrors decision-KPI strip)
  - What changed (review-wizard "changes" data)
  - Late + blocked
  - Decisions made in period
  - CRs decided in period
  - Top risks (by residual)
  - What's next (next 14 d milestones / due actions)
- Export buttons: Copy Markdown · Download .md · Print → PDF.

### F · Action dependencies + critical path
- `a.dependsOn = [actionId, …]`. Drawer UI: searchable multi-select.
- Gantt renders thin connector arrows from each `dependsOn` source's right
  edge to the dependent's left edge.
- Critical-path detection: longest unbroken chain to each `done = false`
  milestone. Highlight participating bars.

### G · Comments + tags
- `a.comments = [{ id, by, at, text }]` shown in drawer; reuse the rich-text
  editor.
- `proj.tags = [{ id, name, color }]` (project-scoped to keep noise low).
- Action / open-point / CR rows render visible tags as small chips.
- Tag filter in topbar (next to existing component filter).

### H · Calendar view
- New nav item under Project.
- Month grid (5 × 7 cells with overflow handling), arrow keys to navigate
  months. Each cell shows due actions + milestones + deliverables + CR
  decision dates + meetings as small color-tinted dots / chips.
- Click a chip → drawer / editor of that item.

### I · Person dashboard
- People list rows already clickable (filter Register). Replace with: click
  → person drawer view aggregating their open actions, week load, late
  items, originated CRs, decisions made, criticality of unmitigated risks
  they own, mini-spark.

### J · Project templates
- "Save current project as template" → strips ids and stores
  `state.templates.push({ id, name, projectShape })`.
- "+ Project" → if templates exist, offer "from scratch" / "from template <X>".
- Instantiating clones the structural skeleton (components, deliverable
  stubs, milestone stubs, risk skeleton) without actions / decisions /
  notes.

### K · Light team-mode
- On every commit, stamp `__lastEditor` and `__lastEditAt` on the affected
  record (action / CR / open-point / risk / decision).
- New "Merge…" entry in the sidebar foot. Picks a second JSON, runs a 3-way
  diff (current, theirs, common ancestor where possible). Per-record UI:
  Keep mine / Take theirs / Edit both.

### L · First-run tour
- 5-step Shepherd-style overlay highlighting: sidebar nav, Quick Add, board
  card, Cmd+K palette, bell. Skippable. Sets `state.settings.tourSeen = true`.

### M · Final polish
- Sweep modal consistency.
- Verify every new feature against the four constraints.
- Update the README / preview screenshots.
- Tag the commit (e.g. `v2.0.0-rc1`) before merging back to `main`.

## Order of execution

A → B → C → D → E → F → G → H → I → J → K → L → M.
A and B can be partially overlapped because B touches no schema.
Each phase ends with a commit on `next-major`. Final phase tags the release.
Merge to `main` only after a full-app smoke pass in the preview.
