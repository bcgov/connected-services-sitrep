// ── RENDER ───────────────────────────────────────────────────────────────────

function renderAll() {
  updateSummary()
  renderGrid()
}

// ── SUMMARY BANNER ────────────────────────────────────────────────────────────
function updateSummary() {
  const teams = Object.values(data)
  const counts = { green: 0, yellow: 0, red: 0 }
  teams.forEach(t => { if (counts[t.status] !== undefined) counts[t.status]++ })
  document.getElementById('g-count').textContent = counts.green
  document.getElementById('y-count').textContent = counts.yellow
  document.getElementById('r-count').textContent = counts.red

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
  pill.textContent = status === 'green' ? '🟢 On Track' : status === 'yellow' ? '🟡 At Risk' : '🔴 Off Track'

  const featHL = coord.featuredHighlights || []
  const hlTeams = [...teams].sort((a, b) => {
    const af = featHL.includes(a.team) ? 0 : 1, bf = featHL.includes(b.team) ? 0 : 1
    return af - bf || ['red','yellow','green'].indexOf(a.status) - ['red','yellow','green'].indexOf(b.status)
  }).filter(t => t.highlight).slice(0, 3)
  document.getElementById('top-highlights').innerHTML = hlTeams.length
    ? hlTeams.map(t => `<div class="summary-item"><span class="summary-item-dot" aria-hidden="true"></span><span class="summary-item-team" title="${esc(t.team)}">${esc(t.team)}</span><span class="summary-item-text">${esc(t.highlight)}</span></div>`).join('')
    : `<div class="summary-item"><span style="color:var(--text3);font-style:italic">No highlights yet</span></div>`

  const featBL = coord.featuredBlockers || []
  const blTeams = [...teams].sort((a, b) => {
    const af = featBL.includes(a.team) ? 0 : 1, bf = featBL.includes(b.team) ? 0 : 1
    return af - bf || ['red','yellow','green'].indexOf(a.status) - ['red','yellow','green'].indexOf(b.status)
  }).filter(t => t.blocker).slice(0, 3)
  document.getElementById('top-blockers').innerHTML = blTeams.length
    ? blTeams.map(t => `<div class="summary-item"><span class="summary-item-dot" aria-hidden="true"></span><span class="summary-item-team" title="${esc(t.team)}">${esc(t.team)}</span><span class="summary-item-text">${esc(t.blocker)}</span></div>`).join('')
    : `<div class="summary-item"><span style="color:var(--text3);font-style:italic">No blockers this week</span></div>`

  const news = coord.news || ''
  const newsEl = document.getElementById('news-el')
  newsEl.className = news ? 'news-text' : 'news-empty'
  newsEl.textContent = news || 'No announcements this week.'
  renderMeetingsDisplay()
}

// ── BENTO GRID ────────────────────────────────────────────────────────────────
function renderGrid() {
  const depsOut = {}
  Object.values(data).forEach(t => {
    ;(t.depsIn || []).forEach(dep => {
      if (!depsOut[dep]) depsOut[dep] = []
      if (!depsOut[dep].includes(t.team)) depsOut[dep].push(t.team)
    })
  })
  const order = ['red', 'yellow', 'green']
  let teamsToShow = TEAMS
  if (currentFilter === 'empty') teamsToShow = TEAMS.filter(n => !data[n])
  else if (currentFilter !== 'all') teamsToShow = TEAMS.filter(n => data[n] && data[n].status === currentFilter)

  if (teamsToShow.length === 0) {
    document.getElementById('grid').innerHTML = `<div style="grid-column:1/-1;text-align:center;padding:48px 0;color:var(--text3)" role="status"><div style="font-size:24px;margin-bottom:8px" aria-hidden="true">✅</div><div style="font-size:14px;font-weight:600;color:var(--text)">All teams have submitted!</div></div>`
    return
  }
  const sorted = [...teamsToShow].sort((a, b) => {
    const ta = data[a], tb = data[b]
    if (!ta && !tb) return TEAMS.indexOf(a) - TEAMS.indexOf(b)
    if (!ta) return 1
    if (!tb) return -1
    return order.indexOf(ta.status) - order.indexOf(tb.status)
  })
  document.getElementById('grid').innerHTML = sorted.map((teamName, i) => {
    const t = data[teamName]
    if (!t) return `<article class="card empty" style="animation-delay:${i * 0.03}s" role="listitem" aria-label="${esc(teamName)}: No data submitted">
      <div class="card-header"><div class="card-team" style="color:var(--text3)">${esc(teamName)}</div></div>
      <div class="card-empty-body">
        <div class="card-empty-icon" aria-hidden="true">📋</div>
        <div class="card-empty-text">No data submitted yet</div>
        <button class="card-add-btn" onclick="openModal('${esc(teamName)}')" aria-label="Add data for ${esc(teamName)}">+ Add data</button>
      </div>
    </article>`
    const s = t.status || 'yellow'
    const label = s === 'green' ? '🟢 Green' : s === 'yellow' ? '🟡 Yellow' : '🔴 Red'
    const di = (t.depsIn || []).map(d => d.replace(/[\[\]"]/g, '').trim()).filter(Boolean)
    const dout = depsOut[teamName] || []
    return `<article class="card ${s}" style="animation-delay:${i * 0.03}s" role="listitem" aria-label="${esc(teamName)}: ${s}">
      <div class="card-header"><div class="card-team">${esc(teamName)}</div><div class="status-pill ${s}" aria-label="Status: ${s}">${label}</div></div>
      ${t.highlight ? `<div class="card-row"><div class="card-lbl">Highlight</div><div class="card-val">${esc(t.highlight)}${t.initiativeNum ? ` <span style="font-size:10px;color:var(--text3)" aria-label="Initiative ${esc(t.initiativeNum)}">#${esc(t.initiativeNum)}</span>` : ''}</div></div>` : ''}
      ${t.blocker ? `<div class="card-row"><div class="card-lbl">Blocker</div><div class="card-val">${esc(t.blocker)}${t.escalatorNum ? ` <span style="font-size:10px;color:var(--red)" aria-label="Escalator ${esc(t.escalatorNum)}"> ↑${esc(t.escalatorNum)}</span>` : ''}</div></div>` : ''}
      ${di.length > 0 || dout.length > 0 ? '<div class="card-divider" role="separator"></div>' : ''}
      ${di.length > 0 ? `<div class="card-row"><div class="card-lbl">Deps in</div><div class="deps-row">${di.map(d => `<span class="dep-tag">${esc(d)}</span>`).join('')}</div></div>` : ''}
      ${dout.length > 0 ? `<div class="card-row"><div class="card-lbl">Deps out</div><div class="deps-row">${dout.map(d => `<span class="dep-tag out">${esc(d)}</span>`).join('')}</div></div>` : ''}
      ${t.summary ? `<div class="card-divider" role="separator"></div><div class="card-summary">${esc(t.summary)}</div>` : ''}
      <button class="card-edit-btn" onclick="openModal('${esc(teamName)}')" aria-label="Edit ${esc(teamName)} data" title="Edit">✏️</button>
    </article>`
  }).join('')
}

// ── HISTORY VIEW ──────────────────────────────────────────────────────────────
function toggleHistory() {
  historyOpen = !historyOpen
  document.getElementById('dashboard-view').style.display = historyOpen ? 'none' : 'block'
  document.getElementById('history-view').classList.toggle('show', historyOpen)
  const btn = document.getElementById('history-btn')
  btn.setAttribute('aria-pressed', historyOpen.toString())
  if (historyOpen) renderHistory()
}

function renderHistory() {
  if (!allHistoryItems.length) {
    document.getElementById('week-selector').innerHTML = '<span style="color:var(--text3);font-size:13px">No history available yet</span>'
    document.getElementById('history-grid').innerHTML = ''
    return
  }

  const byWeek = {}
  allHistoryItems.forEach(item => {
    const f = item.fields
    if (!f.TeamName) return
    const created = new Date(f.Created || item.createdDateTime)
    if (created < new Date('2026-01-01')) return
    const d = new Date(created), day = d.getDay()
    const diff = d.getDate() - day + (day === 0 ? -6 : 1)
    d.setDate(diff)
    const weekLabel = d.toLocaleDateString('en-CA', { day: '2-digit', month: '2-digit', year: 'numeric' })
    if (!byWeek[weekLabel]) byWeek[weekLabel] = []
    byWeek[weekLabel].push({ ...f, _created: created, _id: item.id })
  })

  const weeks = Object.keys(byWeek).sort((a, b) => {
    const pa = a.split('-').reverse().join('-'), pb = b.split('-').reverse().join('-')
    return pb.localeCompare(pa)
  })

  const currentWeek = getWeekLabel()
  document.getElementById('week-selector').innerHTML = weeks.map(w =>
    `<button class="week-btn${w === currentWeek ? ' active' : ''}" onclick="showHistoryWeek('${w}',this)" aria-pressed="${w === currentWeek}">${w}</button>`
  ).join('')

  showHistoryWeek(weeks[0], document.querySelector('.week-btn'))
}

function showHistoryWeek(weekLabel, btn) {
  document.querySelectorAll('.week-btn').forEach(b => { b.classList.remove('active'); b.setAttribute('aria-pressed', 'false') })
  if (btn) { btn.classList.add('active'); btn.setAttribute('aria-pressed', 'true') }

  const allItems = []
  allHistoryItems.forEach(item => {
    const f = item.fields
    if (!f.TeamName) return
    const created = new Date(f.Created || item.createdDateTime)
    if (created < new Date('2026-01-01')) return
    const d = new Date(created), day = d.getDay()
    const diff = d.getDate() - day + (day === 0 ? -6 : 1)
    d.setDate(diff)
    const wl = d.toLocaleDateString('en-CA', { day: '2-digit', month: '2-digit', year: 'numeric' })
    if (wl === weekLabel) allItems.push({ ...f, _created: created, _id: item.id })
  })

  const byTeam = {}
  allItems.forEach(f => {
    if (!byTeam[f.TeamName] || f._created > byTeam[f.TeamName]._created) byTeam[f.TeamName] = f
  })

  const teams = Object.values(byTeam)
  if (!teams.length) {
    document.getElementById('history-grid').innerHTML = `<div style="grid-column:1/-1;color:var(--text3);font-size:13px;padding:20px 0">No data for this week</div>`
    return
  }

  const order = ['red', 'yellow', 'green']
  teams.sort((a, b) =>
    order.indexOf((a.OverallStatus || 'yellow').toLowerCase()) -
    order.indexOf((b.OverallStatus || 'yellow').toLowerCase())
  )

  document.getElementById('history-grid').innerHTML = teams.map(f => {
    const s = (f.OverallStatus || 'yellow').toLowerCase()
    const label = s === 'green' ? '🟢 Green' : s === 'yellow' ? '🟡 Yellow' : '🔴 Red'
    const depsIn = parseDepsIn(f.DependenciesIn)
    return `<article class="history-card ${s}" aria-label="${esc(f.TeamName)}: ${s}">
      <div class="card-header"><div class="card-team">${esc(f.TeamName)}</div><div class="status-pill ${s}">${label}</div></div>
      ${f.Highlight ? `<div class="card-row"><div class="card-lbl">Highlight</div><div class="card-val">${esc(f.Highlight)}${f.InitiativeNumber ? ` <span style="font-size:10px;color:var(--text3)">#${esc(f.InitiativeNumber)}</span>` : ''}</div></div>` : ''}
      ${f.Blocker ? `<div class="card-row"><div class="card-lbl">Blocker</div><div class="card-val">${esc(f.Blocker)}</div></div>` : ''}
      ${depsIn.length > 0 ? `<div class="card-row"><div class="card-lbl">Deps in</div><div class="deps-row">${depsIn.map(d => `<span class="dep-tag">${esc(d)}</span>`).join('')}</div></div>` : ''}
      ${f.WeekSummary ? `<div class="card-divider" role="separator"></div><div class="card-summary">${esc(f.WeekSummary)}</div>` : ''}
    </article>`
  }).join('')
}

// ── MEETINGS DISPLAY (summary banner) ────────────────────────────────────────
function renderMeetingsDisplay() {
  const meetings = coord.meetings || [], el = document.getElementById('meetings-display')
  if (!el) return
  el.innerHTML = meetings.length
    ? meetings.map(m => `<div class="meeting-chip">
        <span aria-hidden="true" style="font-size:14px">📅</span>
        <div style="display:flex;flex-direction:column;gap:1px;">
          <span class="meeting-chip-title">${esc(m.title)}</span>
          ${m.time ? `<span class="meeting-chip-time">${esc(m.time)}</span>` : ''}
        </div>
        ${m.link ? `<a href="${esc(m.link)}" target="_blank" class="meeting-chip-link" aria-label="Join ${esc(m.title)}">Join →</a>` : ''}
      </div>`).join('')
    : '<span class="meeting-empty">No meetings added yet.</span>'
}
