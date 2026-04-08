// ── RENDER.JS ─────────────────────────────────────────────────────────────────
// All DOM rendering — summary banner, bento-grid cards, history view, and
// the meetings display strip in the summary banner.
//
// Dependencies: config.js (data, coord, currentFilter, historyOpen),
//               utils.js (esc, parseDepsIn, getWeekLabel),
//               ui.js (getRosterTeams) — note circular dep; both are global.
//
// None of these functions fetch data. They read the shared `data` and `coord`
// globals and write to the DOM. Call renderAll() after any data change.
// ─────────────────────────────────────────────────────────────────────────────

// Re-render everything: summary banner counts + top items, then the grid.
function renderAll() {
  updateSummary()
  renderGrid()
}

// ── SUMMARY BANNER ────────────────────────────────────────────────────────────
// Update the counts pill, overall-status pill, top-highlights, top-blockers,
// news text, and meetings display in the summary banner at the top of the page.
//
// Overall status logic (when not overridden by coord.status):
//   - ≥3 red teams OR weighted score < 45% → Red
//   - Weighted score ≥ 75%                 → Green
//   - Otherwise                            → Yellow
//   Weighted score = (greens×2 + yellows) / (total×2) × 100
function updateSummary() {
  const teams = Object.values(data)
  const counts = { green: 0, yellow: 0, red: 0 }
  teams.forEach((t) => {
    if (counts[t.status] !== undefined) counts[t.status]++
  })
  document.getElementById('g-count').textContent = counts.green
  document.getElementById('y-count').textContent = counts.yellow
  document.getElementById('r-count').textContent = counts.red

  // Determine overall status — coordinator can override with an explicit value
  let status = coord.status || ''
  if (!status) {
    const total = counts.green + counts.yellow + counts.red
    if (total === 0) {
      status = 'yellow'
    } else {
      const pct = ((counts.green * 2 + counts.yellow) / (total * 2)) * 100
      if (counts.red >= 3 || pct < 45) status = 'red'
      else if (pct >= 75) status = 'green'
      else status = 'yellow'
    }
  }
  const pill = document.getElementById('overall-pill')
  pill.className = 'overall-pill ' + status
  pill.textContent =
    status === 'green'
      ? '🟢 On Track'
      : status === 'yellow'
        ? '🟡 At Risk'
        : '🔴 Off Track'

  // Top highlights — coordinator-featured teams shown first, then by urgency
  const featHL = coord.featuredHighlights || []
  const hlTeams = [...teams]
    .sort((a, b) => {
      const af = featHL.includes(a.team) ? 0 : 1,
        bf = featHL.includes(b.team) ? 0 : 1
      return (
        af - bf ||
        ['red', 'yellow', 'green'].indexOf(a.status) -
          ['red', 'yellow', 'green'].indexOf(b.status)
      )
    })
    .filter((t) => t.highlight)
    .slice(0, 3)
  document.getElementById('top-highlights').innerHTML = hlTeams.length
    ? hlTeams
        .map(
          (t) =>
            `<div class="summary-item"><span class="summary-item-dot" aria-hidden="true"></span><span class="summary-item-team" title="${esc(t.team)}">${esc(t.team)}</span><span class="summary-item-text">${esc(t.highlight)}</span></div>`,
        )
        .join('')
    : `<div class="summary-item"><span style="color:var(--text3);font-style:italic">No highlights yet</span></div>`

  // Top blockers — same sort order as highlights
  const featBL = coord.featuredBlockers || []
  const blTeams = [...teams]
    .sort((a, b) => {
      const af = featBL.includes(a.team) ? 0 : 1,
        bf = featBL.includes(b.team) ? 0 : 1
      return (
        af - bf ||
        ['red', 'yellow', 'green'].indexOf(a.status) -
          ['red', 'yellow', 'green'].indexOf(b.status)
      )
    })
    .filter((t) => t.blocker)
    .slice(0, 3)
  document.getElementById('top-blockers').innerHTML = blTeams.length
    ? blTeams
        .map(
          (t) =>
            `<div class="summary-item"><span class="summary-item-dot" aria-hidden="true"></span><span class="summary-item-team" title="${esc(t.team)}">${esc(t.team)}</span><span class="summary-item-text">${esc(t.blocker)}</span></div>`,
        )
        .join('')
    : `<div class="summary-item"><span style="color:var(--text3);font-style:italic">No blockers this week</span></div>`

  const news = coord.news || ''
  const newsEl = document.getElementById('news-el')
  newsEl.className = news ? 'news-text' : 'news-empty'
  newsEl.textContent = news || 'No announcements this week.'

  renderMeetingsDisplay()
}

// ── BENTO GRID ────────────────────────────────────────────────────────────────
// Render team status cards filtered and sorted by the current filter setting.
//
// Sort order (within each filter):
//   1. Teams WITH data: red → yellow → green
//   2. Teams WITHOUT data: sorted by their position in the roster
//
// "Deps out" (teams depending on this team) are computed by inverting
// the depsIn arrays across all teams.
function renderGrid() {
  // Build reverse-dependency map: depTeam → [teams waiting on depTeam]
  const depsOut = {}
  Object.values(data).forEach((t) => {
    ;(t.depsIn || []).forEach((dep) => {
      if (!depsOut[dep]) depsOut[dep] = []
      if (!depsOut[dep].includes(t.team)) depsOut[dep].push(t.team)
    })
  })

  const order = ['red', 'yellow', 'green']
  const teamsToShowAll = getRosterTeams()
  let teamsToShow = teamsToShowAll

  if (currentFilter === 'empty') {
    teamsToShow = teamsToShowAll.filter((n) => !data[n])
  } else if (currentFilter !== 'all') {
    teamsToShow = teamsToShowAll.filter(
      (n) => data[n] && data[n].status === currentFilter,
    )
  }

  if (teamsToShow.length === 0) {
    document.getElementById('grid').innerHTML =
      `<div style="grid-column:1/-1;text-align:center;padding:48px 0;color:var(--text3)" role="status"><div style="font-size:24px;margin-bottom:8px" aria-hidden="true">✅</div><div style="font-size:14px;font-weight:600;color:var(--text)">All teams have submitted!</div></div>`
    return
  }

  const rosterOrder = getRosterTeams()
  const sorted = [...teamsToShow].sort((a, b) => {
    const ta = data[a],
      tb = data[b]
    if (!ta && !tb) {
      // Both empty: preserve roster order, then alpha
      const ai = rosterOrder.indexOf(a),
        bi = rosterOrder.indexOf(b)
      if (ai !== bi) return ai - bi
      return a.localeCompare(b, undefined, { sensitivity: 'base' })
    }
    if (!ta) return 1  // empty cards sink to the bottom
    if (!tb) return -1
    return order.indexOf(ta.status) - order.indexOf(tb.status)
  })

  document.getElementById('grid').innerHTML = sorted
    .map((teamName, i) => {
      const t = data[teamName]

      // Empty card — team has no submission this week
      if (!t) {
        return `<article class="card empty" style="animation-delay:${i * 0.03}s" role="listitem" aria-label="${esc(teamName)}: No data submitted">
      <div class="card-header"><div class="card-team" style="color:var(--text3)">${esc(teamName)}</div></div>
      <div class="card-empty-body">
        <div class="card-empty-icon" aria-hidden="true">📋</div>
        <div class="card-empty-text">No data submitted yet</div>
        <button class="card-add-btn" onclick="openModal('${esc(teamName)}')" aria-label="Add data for ${esc(teamName)}">+ Add data</button>
      </div>
    </article>`
      }

      const s = t.status || 'yellow'
      const label =
        s === 'green' ? '🟢 Green' : s === 'yellow' ? '🟡 Yellow' : '🔴 Red'
      const di = (t.depsIn || [])
        .map((d) => d.replace(/[\[\]"]/g, '').trim())
        .filter(Boolean)
      const dout = depsOut[teamName] || []

      return `<article class="card ${s}" style="animation-delay:${i * 0.03}s" role="listitem" aria-label="${esc(teamName)}: ${s}">
      <div class="card-header"><div class="card-team">${esc(teamName)}</div><div class="status-pill ${s}" aria-label="Status: ${s}">${label}</div></div>
      ${t.highlight ? `<div class="card-row"><div class="card-lbl">Highlight</div><div class="card-val">${esc(t.highlight)}${t.initiativeNum ? ` <span style="font-size:10px;color:var(--text3)" aria-label="Initiative ${esc(t.initiativeNum)}">#${esc(t.initiativeNum)}</span>` : ''}</div></div>` : ''}
      ${t.blocker ? `<div class="card-row"><div class="card-lbl">Blocker</div><div class="card-val">${esc(t.blocker)}${t.escalatorNum ? ` <span style="font-size:10px;color:var(--red)" aria-label="Escalator ${esc(t.escalatorNum)}"> ↑${esc(t.escalatorNum)}</span>` : ''}</div></div>` : ''}
      ${di.length > 0 || dout.length > 0 ? '<div class="card-divider" role="separator"></div>' : ''}
      ${di.length > 0 ? `<div class="card-row"><div class="card-lbl">Deps in</div><div class="deps-row">${di.map((d) => `<span class="dep-tag">${esc(d)}</span>`).join('')}</div></div>` : ''}
      ${dout.length > 0 ? `<div class="card-row"><div class="card-lbl">Deps out</div><div class="deps-row">${dout.map((d) => `<span class="dep-tag out">${esc(d)}</span>`).join('')}</div></div>` : ''}
      ${t.summary ? `<div class="card-divider" role="separator"></div><div class="card-summary">${esc(t.summary)}</div>` : ''}
      <button class="card-edit-btn" onclick="openModal('${esc(teamName)}')" aria-label="Edit ${esc(teamName)} data" title="Edit">✏️</button>
    </article>`
    })
    .join('')
}

// ── HISTORY VIEW ──────────────────────────────────────────────────────────────
// Render the week-picker row and show the most recent submission per team for
// the initially selected week. All data comes from allHistoryItems[].

// Given a Date, return the Monday of that week at local midnight.
// Shared by renderHistory and showHistoryWeek to keep grouping consistent.
function getWeekMonday(date) {
  const d = new Date(date),
    day = d.getDay()
  const diff = d.getDate() - day + (day === 0 ? -6 : 1)
  d.setDate(diff)
  d.setHours(0, 0, 0, 0)
  return d
}

// Toggle the history view panel on/off.
function toggleHistory() {
  historyOpen = !historyOpen
  document.getElementById('dashboard-view').style.display = historyOpen
    ? 'none'
    : 'block'
  document.getElementById('history-view').classList.toggle('show', historyOpen)
  const btn = document.getElementById('history-btn')
  btn.setAttribute('aria-pressed', historyOpen.toString())
  if (historyOpen) renderHistory()
}

// Build the week-selector button row from allHistoryItems, then show the
// most recent week. Items created before 2026-01-01 are excluded (pre-project).
function renderHistory() {
  if (!allHistoryItems.length) {
    document.getElementById('week-selector').innerHTML =
      '<span style="color:var(--text3);font-size:13px">No history available yet</span>'
    document.getElementById('history-grid').innerHTML = ''
    return
  }

  // Group items by their week's Monday label (en-CA: YYYY-MM-DD)
  const byWeek = {}
  allHistoryItems.forEach((item) => {
    const f = item.fields
    if (!f.TeamName) return
    const created = new Date(f.Created || item.createdDateTime)
    if (created < new Date('2026-01-01')) return

    const monday = getWeekMonday(created)
    const weekLabel = monday.toLocaleDateString('en-CA', {
      day: '2-digit',
      month: '2-digit',
      year: 'numeric',
    })
    if (!byWeek[weekLabel]) byWeek[weekLabel] = []
    byWeek[weekLabel].push({ ...f, _created: created, _id: item.id })
  })

  // Sort weeks newest-first (en-CA labels are YYYY-MM-DD so lexicographic works)
  const weeks = Object.keys(byWeek).sort((a, b) => b.localeCompare(a))

  const currentWeek = getWeekLabel()
  document.getElementById('week-selector').innerHTML = weeks
    .map(
      (w) =>
        `<button class="week-btn${w === currentWeek ? ' active' : ''}" onclick="showHistoryWeek('${w}',this)" aria-pressed="${w === currentWeek}">${w}</button>`,
    )
    .join('')

  // Show the most recent week by default
  const firstBtn = document.querySelector('.week-btn')
  if (weeks.length && firstBtn) showHistoryWeek(weeks[0], firstBtn)
}

// Render history cards for a specific week label.
// Shows the most recent submission per team for that week.
//
// @param {string} weekLabel - en-CA date string matching a week-selector button
// @param {HTMLElement} btn  - The button element to mark active
function showHistoryWeek(weekLabel, btn) {
  document.querySelectorAll('.week-btn').forEach((b) => {
    b.classList.remove('active')
    b.setAttribute('aria-pressed', 'false')
  })
  if (btn) {
    btn.classList.add('active')
    btn.setAttribute('aria-pressed', 'true')
  }

  // Collect items that belong to this week
  const allItems = []
  allHistoryItems.forEach((item) => {
    const f = item.fields
    if (!f.TeamName) return
    const created = new Date(f.Created || item.createdDateTime)
    if (created < new Date('2026-01-01')) return

    const monday = getWeekMonday(created)
    const wl = monday.toLocaleDateString('en-CA', {
      day: '2-digit',
      month: '2-digit',
      year: 'numeric',
    })
    if (wl === weekLabel)
      allItems.push({ ...f, _created: created, _id: item.id })
  })

  // Deduplicate: keep only the most recent entry per team
  const byTeam = {}
  allItems.forEach((f) => {
    if (!byTeam[f.TeamName] || f._created > byTeam[f.TeamName]._created)
      byTeam[f.TeamName] = f
  })

  const teams = Object.values(byTeam)
  if (!teams.length) {
    document.getElementById('history-grid').innerHTML =
      `<div style="grid-column:1/-1;color:var(--text3);font-size:13px;padding:20px 0">No data for this week</div>`
    return
  }

  const order = ['red', 'yellow', 'green']
  teams.sort(
    (a, b) =>
      order.indexOf((a.OverallStatus || 'yellow').toLowerCase()) -
      order.indexOf((b.OverallStatus || 'yellow').toLowerCase()),
  )

  document.getElementById('history-grid').innerHTML = teams
    .map((f) => {
      const s = (f.OverallStatus || 'yellow').toLowerCase()
      const label =
        s === 'green' ? '🟢 Green' : s === 'yellow' ? '🟡 Yellow' : '🔴 Red'
      const depsIn = parseDepsIn(f.DependenciesIn)
      return `<article class="history-card ${s}" aria-label="${esc(f.TeamName)}: ${s}">
      <div class="card-header"><div class="card-team">${esc(f.TeamName)}</div><div class="status-pill ${s}">${label}</div></div>
      ${f.Highlight ? `<div class="card-row"><div class="card-lbl">Highlight</div><div class="card-val">${esc(f.Highlight)}${f.InitiativeNumber ? ` <span style="font-size:10px;color:var(--text3)">#${esc(f.InitiativeNumber)}</span>` : ''}</div></div>` : ''}
      ${f.Blocker ? `<div class="card-row"><div class="card-lbl">Blocker</div><div class="card-val">${esc(f.Blocker)}</div></div>` : ''}
      ${depsIn.length > 0 ? `<div class="card-row"><div class="card-lbl">Deps in</div><div class="deps-row">${depsIn.map((d) => `<span class="dep-tag">${esc(d)}</span>`).join('')}</div></div>` : ''}
      ${f.WeekSummary ? `<div class="card-divider" role="separator"></div><div class="card-summary">${esc(f.WeekSummary)}</div>` : ''}
    </article>`
    })
    .join('')
}

// ── MEETINGS DISPLAY ──────────────────────────────────────────────────────────
// Render the meetings chip strip in the summary banner (read-only display).
// The editable meeting list in the coordinator panel is rendered by
// renderMeetingsList() in ui.js.
function renderMeetingsDisplay() {
  const meetings = coord.meetings || [],
    el = document.getElementById('meetings-display')
  if (!el) return
  el.innerHTML = meetings.length
    ? meetings
        .map(
          (m) => `<div class="meeting-chip">
        <span aria-hidden="true" style="font-size:14px">📅</span>
        <div style="display:flex;flex-direction:column;gap:1px;">
          <span class="meeting-chip-title">${esc(m.title)}</span>
          ${m.time ? `<span class="meeting-chip-time">${esc(m.time)}</span>` : ''}
        </div>
        ${m.link ? `<a href="${esc(m.link)}" target="_blank" class="meeting-chip-link" aria-label="Join ${esc(m.title)}">Join →</a>` : ''}
      </div>`,
        )
        .join('')
    : '<span class="meeting-empty">No meetings added yet.</span>'
}
