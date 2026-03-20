// ── UI ───────────────────────────────────────────────────────────────────────

// ── INIT ──────────────────────────────────────────────────────────────────────
function init() {
  document.getElementById('week-chip').textContent = 'Week of ' + getWeekLabel()
  buildDepsPicker()
  if (CONFIG.useSharePoint) {
    document.getElementById('csv-controls').style.display = 'none'
    if (window.msal) { loadFromSharePoint() }
    else {
      const c = setInterval(() => { if (window.msal) { clearInterval(c); loadFromSharePoint() } }, 100)
    }
  } else {
    data = JSON.parse(localStorage.getItem('sitrep_v2') || '{}')
    coord = JSON.parse(localStorage.getItem('sitrep_coord') || '{}')
    renderAll()
  }
}

// ── FILTER ────────────────────────────────────────────────────────────────────
function setFilter(f, btn) {
  currentFilter = f
  document.querySelectorAll('.filter-btn').forEach(b => b.setAttribute('aria-pressed', 'false'))
  btn.setAttribute('aria-pressed', 'true')
  renderGrid()
}

// ── MODAL ─────────────────────────────────────────────────────────────────────
function openModal(teamName) {
  const t = data[teamName]
  document.getElementById('modal-title').textContent = t ? `Edit — ${teamName}` : `Add data — ${teamName}`
  document.getElementById('f-team').value = teamName
  document.getElementById('f-highlight').value = t?.highlight || ''
  document.getElementById('f-blocker').value = t?.blocker || ''
  document.getElementById('f-init').value = t?.initiativeNum || ''
  document.getElementById('f-esc').value = t?.escalatorNum || ''
  document.getElementById('f-summary').value = t?.summary || ''
  document.getElementById('modal-save-status').textContent = ''
  selectedRYG = t?.status || null
  ;['green', 'yellow', 'red'].forEach(x => {
    const btn = document.getElementById('r' + x[0])
    btn.className = 'ryg-opt' + (x === selectedRYG ? ' sel-' + x : '')
    btn.setAttribute('aria-pressed', (x === selectedRYG).toString())
  })
  selectedDeps = [...(t?.depsIn || [])]
  TEAMS.forEach(tm => {
    const el = document.getElementById('dep-' + tm.replace(/\//g, '-'))
    if (el) {
      el.classList.toggle('on', selectedDeps.includes(tm))
      el.setAttribute('aria-pressed', selectedDeps.includes(tm).toString())
    }
  })
  // Populate week picker
  const [prev, curr, next] = getWeekOptions()
  const weekSel = document.getElementById('f-week')
  weekSel.innerHTML = `
    <option value="${prev}">${prev} (last week)</option>
    <option value="${curr}" selected>${curr} (this week)</option>
    <option value="${next}">${next} (next week)</option>
  `
  // If editing an existing entry keep its week, otherwise default to this week
  weekSel.value = t?._weekOf || curr

  document.getElementById('modal-overlay').classList.add('show')
  document.body.style.overflow = 'hidden'
  setTimeout(() => document.getElementById('f-team').focus(), 50)
}

function closeModal() {
  document.getElementById('modal-overlay').classList.remove('show')
  document.body.style.overflow = ''
  selectedRYG = null
  selectedDeps = []
}

function closeModalOutside(e) {
  if (e.target === document.getElementById('modal-overlay')) closeModal()
}

function saveTeam() {
  const team = document.getElementById('f-team').value
  if (!team || !selectedRYG) { showToast('Please select a status'); return }
  const teamData = {
    team,
    status: selectedRYG,
    highlight: document.getElementById('f-highlight').value.trim(),
    blocker: document.getElementById('f-blocker').value.trim(),
    initiativeNum: document.getElementById('f-init').value.trim(),
    escalatorNum: document.getElementById('f-esc').value.trim(),
    depsIn: [...selectedDeps],
    summary: document.getElementById('f-summary').value.trim(),
    _weekOf: document.getElementById('f-week').value,
    _spId: data[team]?._spId || null,
  }
  data[team] = teamData
  if (!CONFIG.useSharePoint) localStorage.setItem('sitrep_v2', JSON.stringify(data))
  else saveTeamToSharePoint(team, teamData)
  renderAll()
  renderFeaturedChips()
  showToast('✓ ' + team + ' saved!')
  setTimeout(() => closeModal(), 1200)
}

// Focus trap in modal
document.addEventListener('keydown', e => {
  if (e.key === 'Escape') { closeModal(); return }
  if (e.key === 'Tab' && document.getElementById('modal-overlay').classList.contains('show')) {
    const modal = document.querySelector('.modal')
    const focusable = modal.querySelectorAll('button,input,select,textarea,[tabindex]:not([tabindex="-1"])')
    const first = focusable[0], last = focusable[focusable.length - 1]
    if (e.shiftKey) { if (document.activeElement === first) { e.preventDefault(); last.focus() } }
    else { if (document.activeElement === last) { e.preventDefault(); first.focus() } }
  }
})

// ── RYG PICKER ────────────────────────────────────────────────────────────────
function pickRYG(c) {
  selectedRYG = c
  ;['green', 'yellow', 'red'].forEach(x => {
    const btn = document.getElementById('r' + x[0])
    btn.className = 'ryg-opt' + (x === c ? ' sel-' + c : '')
    btn.setAttribute('aria-pressed', (x === c).toString())
  })
}

// ── DEPS PICKER ───────────────────────────────────────────────────────────────
function buildDepsPicker() {
  document.getElementById('deps-picker').innerHTML = TEAMS.map(t =>
    `<button type="button" class="featured-chip" id="dep-${t.replace(/\//g, '-')}" onclick="toggleDep('${esc(t)}')" aria-pressed="false">${esc(t)}</button>`
  ).join('')
}

function toggleDep(team) {
  const idx = selectedDeps.indexOf(team)
  if (idx >= 0) selectedDeps.splice(idx, 1)
  else selectedDeps.push(team)
  const el = document.getElementById('dep-' + team.replace(/\//g, '-'))
  if (el) {
    el.classList.toggle('on', selectedDeps.includes(team))
    el.setAttribute('aria-pressed', selectedDeps.includes(team).toString())
  }
}

// ── COORDINATOR ───────────────────────────────────────────────────────────────
function toggleCoord() {
  coordOpen = !coordOpen
  document.getElementById('coord-bar').classList.toggle('show', coordOpen)
  document.getElementById('coord-bar').setAttribute('aria-hidden', (!coordOpen).toString())
  const btn = document.getElementById('coord-btn')
  btn.setAttribute('aria-pressed', coordOpen.toString())
  btn.setAttribute('aria-expanded', coordOpen.toString())
  if (coordOpen) {
    document.getElementById('coord-status').value = coord.status || ''
    document.getElementById('coord-news').value = coord.news || ''
    renderFeaturedChips()
    renderMeetingsList()
    document.getElementById('coord-status').focus()
  }
}

function saveCoord() {
  coord.status = document.getElementById('coord-status').value
  coord.news = document.getElementById('coord-news').value
  debounceSave()
  updateSummary()
}

function renderFeaturedChips() {
  const teams = Object.values(data)
  const featHL = coord.featuredHighlights || [], featBL = coord.featuredBlockers || []
  document.getElementById('feat-hl-chips').innerHTML = teams.filter(t => t.highlight).map(t =>
    `<button type="button" class="featured-chip ${featHL.includes(t.team) ? 'on' : ''}" onclick="toggleFeat('hl','${esc(t.team)}')" aria-pressed="${featHL.includes(t.team)}">${esc(t.team)}</button>`
  ).join('') || '<span style="font-size:12px;color:var(--text3)">No highlights yet</span>'
  document.getElementById('feat-bl-chips').innerHTML = teams.filter(t => t.blocker).map(t =>
    `<button type="button" class="featured-chip ${featBL.includes(t.team) ? 'on' : ''}" onclick="toggleFeat('bl','${esc(t.team)}')" aria-pressed="${featBL.includes(t.team)}">${esc(t.team)}</button>`
  ).join('') || '<span style="font-size:12px;color:var(--text3)">No blockers yet</span>'
}

function toggleFeat(type, team) {
  const key = type === 'hl' ? 'featuredHighlights' : 'featuredBlockers'
  if (!coord[key]) coord[key] = []
  const idx = coord[key].indexOf(team)
  if (idx >= 0) coord[key].splice(idx, 1)
  else coord[key].push(team)
  debounceSave()
  renderFeaturedChips()
  updateSummary()
}

// ── MEETINGS ──────────────────────────────────────────────────────────────────
function addMeeting() {
  const title = document.getElementById('mtg-title').value.trim()
  const time = document.getElementById('mtg-time').value.trim()
  const link = document.getElementById('mtg-link').value.trim()
  if (!title) { showToast('Please enter a meeting title'); return }
  if (!coord.meetings) coord.meetings = []
  coord.meetings.push({ title, time, link, id: Date.now() })
  debounceSave()
  document.getElementById('mtg-title').value = ''
  document.getElementById('mtg-time').value = ''
  document.getElementById('mtg-link').value = ''
  renderMeetingsList()
  updateSummary()
  showToast('✓ Meeting added')
}

function removeMeeting(id) {
  coord.meetings = (coord.meetings || []).filter(m => m.id !== id)
  debounceSave()
  renderMeetingsList()
  updateSummary()
}

function renderMeetingsList() {
  const meetings = coord.meetings || [], el = document.getElementById('meetings-list')
  if (!el) return
  el.innerHTML = meetings.length
    ? meetings.map(m => `<div style="display:flex;align-items:center;gap:8px;background:white;border:1px solid #bfdbfe;border-radius:var(--radius-sm);padding:6px 10px;font-size:12px;">
        <span style="flex:1;font-weight:600">${esc(m.title)}</span>
        ${m.time ? `<span style="color:var(--text3);font-size:11px">${esc(m.time)}</span>` : ''}
        ${m.link ? `<a href="${esc(m.link)}" target="_blank" style="font-size:11px;color:var(--link)">🔗 Link</a>` : ''}
        <button type="button" onclick="removeMeeting(${m.id})" style="background:none;border:none;color:var(--text3);cursor:pointer;font-size:14px;line-height:1;padding:2px 4px;" aria-label="Remove meeting ${esc(m.title)}">✕</button>
      </div>`).join('')
    : '<span style="font-size:12px;color:var(--text3)">No meetings added yet</span>'
}

// ── CSV UPLOAD ────────────────────────────────────────────────────────────────
function uploadCSV(event) {
  const file = event.target.files[0]
  if (!file) return
  const reader = new FileReader()
  reader.onload = function (e) {
    const text = e.target.result.replace(/^\uFEFF/, '')
    const lines = text.split(/\r?\n/).filter(l => l.trim())
    const headers = lines[0].split(',').map(h => h.trim().replace(/"/g, ''))
    let count = 0
    for (let i = 1; i < lines.length; i++) {
      const cols = []; let current = '', inQuotes = false
      for (let c = 0; c < lines[i].length; c++) {
        const ch = lines[i][c]
        if (ch === '"') { inQuotes = !inQuotes }
        else if (ch === ',' && !inQuotes) { cols.push(current.trim()); current = '' }
        else { current += ch }
      }
      cols.push(current.trim())
      const row = {}
      headers.forEach((h, j) => row[h] = (cols[j] || '').trim())
      const team = row.TeamName
      if (!team) continue
      data[team] = {
        team,
        status: (row.OverallStatus || 'yellow').toLowerCase(),
        highlight: row.Highlight || '',
        blocker: row.Blocker || '',
        initiativeNum: row.InitiativeNumber || '',
        escalatorNum: row.EscalatorNumber || '',
        depsIn: parseDepsIn(row.DependenciesIn),
        summary: row.WeekSummary || '',
        _spId: null,
      }
      count++
    }
    localStorage.setItem('sitrep_v2', JSON.stringify(data))
    renderAll()
    showToast(`✓ ${count} teams loaded from CSV`)
    event.target.value = ''
  }
  reader.readAsText(file)
}

function clearAll() {
  if (!confirm('Clear all team data for this week? This cannot be undone.')) return
  data = {}
  localStorage.removeItem('sitrep_v2')
  renderAll()
  showToast('Week cleared')
}
