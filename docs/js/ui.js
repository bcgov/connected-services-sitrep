// ── UI.JS ─────────────────────────────────────────────────────────────────────
// User interaction handlers: init, modal open/save/close, coordinator panel,
// meetings editor, filter bar, CSV upload, team roster helpers, and the
// smart auto-refresh loop.
//
// Dependencies: config.js (state), utils.js (helpers), auth.js (getToken),
//               sharepoint.js (load/save), render.js (renderAll, renderGrid)
//
// Entry point: init() is called by the MSAL <script> tag's onload attribute
// in sitrepdash.html once the MSAL library has downloaded.
// ─────────────────────────────────────────────────────────────────────────────

// ── INIT ──────────────────────────────────────────────────────────────────────
// Called once when the MSAL browser library finishes loading.
// Sets up the week chip, builds the dependency picker (rebuilt again on each
// modal open once SP data is available), then kicks off authentication and load.
function init() {
  document.getElementById('week-chip').textContent = 'Week of ' + getWeekLabel()
  buildDepsPicker()

  if (CONFIG.useSharePoint) {
    // Hide CSV controls — they are only useful in non-SP mode
    document.getElementById('csv-controls').style.display = 'none'

    if (window.msal) {
      loadFromSharePoint()
    } else {
      // MSAL CDN sometimes fires onload before window.msal is defined;
      // poll until it is available. Give up after 10 s to avoid an
      // infinite loop if the script fails to initialise.
      let waited = 0
      const c = setInterval(() => {
        waited += 100
        if (window.msal) {
          clearInterval(c)
          loadFromSharePoint()
        } else if (waited >= 10000) {
          clearInterval(c)
          console.error('[INIT] MSAL did not initialise after 10 s')
          showErrorModal(
            'Authentication Library Failed',
            'The Microsoft authentication library did not load. Check your network connection and reload the page.',
            '',
          )
        }
      }, 100)
    }

    startAutoRefresh()
  } else {
    // Non-SP mode: load from localStorage and CSV only
    loadLocalDataAndSync()
    renderAll()
  }
}

// ── AUTO-REFRESH ──────────────────────────────────────────────────────────────
// Polls for changes every 60 seconds using a lightweight "last modified"
// check rather than a full data fetch. Only triggers a full reload when
// SharePoint data has actually changed. Paused when the tab is hidden.

// Timestamps of the last-seen modifications; used to detect changes without
// a full data fetch.
let lastDataHash = null
let lastCoordHash = null

// Start (or restart) the 60-second polling interval.
function startAutoRefresh() {
  if (autoRefreshTimer) clearInterval(autoRefreshTimer)
  autoRefreshTimer = setInterval(async () => {
    if (CONFIG.useSharePoint && document.visibilityState === 'visible') {
      try {
        await checkForUpdates()
      } catch (e) {
        console.warn('[AUTO-REFRESH] Check failed:', e.message)
      }
    }
  }, 60000)
  console.log('[AUTO-REFRESH] Started (60 s interval, change detection)')
}

// Quick check: fetch only the most recently modified item from each list.
// If the lastModifiedDateTime is different from what we saw last time,
// trigger a full loadFromSharePoint(). This avoids unnecessary full fetches
// when nothing has changed.
async function checkForUpdates() {
  try {
    const token = await getToken()
    if (!token || !_siteId || !_teamListId) return

    const itemsResp = await fetch(
      `https://graph.microsoft.com/v1.0/sites/${_siteId}/lists/${_teamListId}/items?$top=1&$orderby=lastModifiedDateTime desc&$select=lastModifiedDateTime,id`,
      { headers: { Authorization: `Bearer ${token}` } },
    )
    if (!itemsResp.ok) return

    const items = await itemsResp.json()
    const latestModified = items.value?.[0]?.lastModifiedDateTime

    let coordModified = null
    if (_coordListId) {
      const coordResp = await fetch(
        `https://graph.microsoft.com/v1.0/sites/${_siteId}/lists/${_coordListId}/items?$top=1&$orderby=lastModifiedDateTime desc&$select=lastModifiedDateTime,id`,
        { headers: { Authorization: `Bearer ${token}` } },
      )
      if (coordResp.ok) {
        const coordItems = await coordResp.json()
        coordModified = coordItems.value?.[0]?.lastModifiedDateTime
      }
    }

    const currentHash = `${latestModified || ''}|${coordModified || ''}`
    const previousHash = `${lastDataHash || ''}|${lastCoordHash || ''}`

    if (currentHash !== previousHash) {
      console.log('[AUTO-REFRESH] Changes detected, refreshing data...')
      lastDataHash = latestModified
      lastCoordHash = coordModified
      await loadFromSharePoint()
    }
  } catch (e) {
    // Silent failure is acceptable here — the next interval will retry.
    console.warn('[AUTO-REFRESH] checkForUpdates error:', e.message)
  }
}

// Stop the polling interval (called when the tab is hidden).
function stopAutoRefresh() {
  if (autoRefreshTimer) {
    clearInterval(autoRefreshTimer)
    autoRefreshTimer = null
    console.log('[AUTO-REFRESH] Stopped')
  }
}

// Pause auto-refresh when the tab is hidden, resume when it becomes visible.
// This avoids making background API calls when the user has switched tabs.
document.addEventListener('visibilitychange', () => {
  if (CONFIG.useSharePoint) {
    if (document.hidden) stopAutoRefresh()
    else startAutoRefresh()
  }
})

// ── FILTER BAR ────────────────────────────────────────────────────────────────
// Updates currentFilter and re-renders the grid.
// @param {string} f      - Filter key: 'all' | 'green' | 'yellow' | 'red' | 'empty'
// @param {HTMLElement} btn - The clicked filter button (to update aria-pressed)
function setFilter(f, btn) {
  currentFilter = f
  document
    .querySelectorAll('.filter-btn')
    .forEach((b) => b.setAttribute('aria-pressed', 'false'))
  btn.setAttribute('aria-pressed', 'true')
  renderGrid()
}

// ── MODAL ─────────────────────────────────────────────────────────────────────
// Open the team edit/add modal pre-populated with the team's current data.
// The team and deps selects are always rebuilt here so new teams from SP
// are available without requiring a page reload.
//
// @param {string} teamName - The team to edit; may or may not have data yet.
function openModal(teamName) {
  // Rebuild both selects so newly loaded SP teams are present
  buildTeamSelect()
  buildDepsPicker()

  const t = data[teamName]
  document.getElementById('modal-title').textContent = t
    ? `Edit — ${teamName}`
    : `Add data — ${teamName}`
  document.getElementById('f-team').value = teamName
  document.getElementById('f-highlight').value = t?.highlight || ''
  document.getElementById('f-blocker').value = t?.blocker || ''
  document.getElementById('f-init').value = t?.initiativeNum || ''
  document.getElementById('f-esc').value = t?.escalatorNum || ''
  document.getElementById('f-summary').value = t?.summary || ''
  document.getElementById('modal-save-status').textContent = ''

  // RYG picker state
  selectedRYG = t?.status || null
  ;['green', 'yellow', 'red'].forEach((x) => {
    const btn = document.getElementById('r' + x[0])
    btn.className = 'ryg-opt' + (x === selectedRYG ? ' sel-' + x : '')
    btn.setAttribute('aria-pressed', (x === selectedRYG).toString())
  })

  // Deps picker state — apply after buildDepsPicker() recreates the buttons
  selectedDeps = [...(t?.depsIn || [])]
  getRosterTeams().forEach((tm) => {
    const el = document.getElementById('dep-' + tm.replace(/\//g, '-'))
    if (el) {
      el.classList.toggle('on', selectedDeps.includes(tm))
      el.setAttribute('aria-pressed', selectedDeps.includes(tm).toString())
    }
  })

  // Week picker
  const [prev, curr, next] = getWeekOptions()
  const weekSel = document.getElementById('f-week')
  weekSel.innerHTML = `
    <option value="${prev}">${prev} (last week)</option>
    <option value="${curr}" selected>${curr} (this week)</option>
    <option value="${next}">${next} (next week)</option>
  `
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

// Close the modal when the user clicks the dark overlay outside it.
function closeModalOutside(e) {
  if (e.target === document.getElementById('modal-overlay')) closeModal()
}

// Validate required fields, build the teamData object, and save to SP (or
// localStorage if SP is disabled). Keeps the modal open on failure so the
// user can retry or see the error detail.
async function saveTeam() {
  const team = document.getElementById('f-team').value
  const highlight = document.getElementById('f-highlight').value.trim()

  const missing = []
  if (!team) missing.push('Team name')
  if (!selectedRYG) missing.push('Status (Red / Yellow / Green)')
  if (!highlight) missing.push('Key Highlight')

  if (missing.length > 0) {
    showErrorModal(
      'Required Fields Missing',
      `Please fill in all required fields before saving:<br><br>${missing.map((f) => `• ${f}`).join('<br>')}`,
      '',
    )
    return
  }

  const teamData = {
    team,
    status: selectedRYG,
    highlight,
    blocker: document.getElementById('f-blocker').value.trim(),
    initiativeNum: document.getElementById('f-init').value.trim(),
    escalatorNum: document.getElementById('f-esc').value.trim(),
    depsIn: [...selectedDeps],
    summary: document.getElementById('f-summary').value.trim(),
    _weekOf: document.getElementById('f-week').value,
    _spId: data[team]?._spId || null,
  }
  data[team] = teamData
  console.log('[SAVE] Saving team:', team, teamData)

  if (!CONFIG.useSharePoint) {
    localStorage.setItem('sitrep_v2', JSON.stringify(data))
    renderAll()
    renderFeaturedChips()
    showToast('✓ ' + team + ' saved locally!')
    setTimeout(() => closeModal(), 1200)
    return
  }

  try {
    await saveTeamToSharePoint(team, teamData)
    // Only reach here on success — saveTeamToSharePoint re-throws on failure
    renderAll()
    renderFeaturedChips()
    showToast('✓ ' + team + ' saved to SharePoint!')
    setTimeout(() => closeModal(), 1200)
  } catch (e) {
    // Error modal already shown by saveTeamToSharePoint.
    // Keep the modal open so the user can retry or copy the error detail.
    console.error('[SAVE] Save failed:', e.message)
  }
}

// Focus trap: Tab cycles within the modal; Escape closes it.
document.addEventListener('keydown', (e) => {
  if (e.key === 'Escape') {
    closeModal()
    return
  }
  if (
    e.key === 'Tab' &&
    document.getElementById('modal-overlay').classList.contains('show')
  ) {
    const modal = document.querySelector('.modal')
    const focusable = modal.querySelectorAll(
      'button,input,select,textarea,[tabindex]:not([tabindex="-1"])',
    )
    const first = focusable[0],
      last = focusable[focusable.length - 1]
    if (e.shiftKey) {
      if (document.activeElement === first) {
        e.preventDefault()
        last.focus()
      }
    } else {
      if (document.activeElement === last) {
        e.preventDefault()
        first.focus()
      }
    }
  }
})

// ── RYG PICKER ────────────────────────────────────────────────────────────────
// Update the selected RYG button and reflect the choice visually.
function pickRYG(c) {
  selectedRYG = c
  ;['green', 'yellow', 'red'].forEach((x) => {
    const btn = document.getElementById('r' + x[0])
    btn.className = 'ryg-opt' + (x === c ? ' sel-' + c : '')
    btn.setAttribute('aria-pressed', (x === c).toString())
  })
}

// ── DEPS PICKER ───────────────────────────────────────────────────────────────
// Build (or rebuild) the chip buttons for the "Dependencies in" picker.
// Called at init and again every time the modal is opened, so the list always
// reflects the latest team roster from SharePoint.
function buildDepsPicker() {
  document.getElementById('deps-picker').innerHTML = getRosterTeams()
    .map(
      (t) =>
        `<button type="button" class="featured-chip" id="dep-${t.replace(/\//g, '-')}" onclick="toggleDep('${esc(t)}')" aria-pressed="false">${esc(t)}</button>`,
    )
    .join('')
}

// Toggle a team in the selectedDeps array and update its chip button state.
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

// ── COORDINATOR PANEL ─────────────────────────────────────────────────────────
// Toggle the coordinator controls panel below the summary banner.
function toggleCoord() {
  coordOpen = !coordOpen
  document.getElementById('coord-bar').classList.toggle('show', coordOpen)
  document
    .getElementById('coord-bar')
    .setAttribute('aria-hidden', (!coordOpen).toString())
  const btn = document.getElementById('coord-btn')
  btn.setAttribute('aria-pressed', coordOpen.toString())
  btn.setAttribute('aria-expanded', coordOpen.toString())
  if (coordOpen) {
    document.getElementById('coord-status').value = coord.status || ''
    document.getElementById('coord-news').value = coord.news || ''
    renderTeamsList()
    renderFeaturedChips()
    renderMeetingsList()
    document.getElementById('coord-status').focus()
  }
}

// Capture coordinator field changes and trigger a debounced save.
function saveCoord() {
  coord.status = document.getElementById('coord-status').value
  coord.news = document.getElementById('coord-news').value
  debounceSave()
  updateSummary()
}

// Render the featured highlight/blocker chip pickers in the coordinator panel.
// Chips are built from teams that have submitted data; toggled chips are saved
// to coord.featuredHighlights / coord.featuredBlockers.
function renderFeaturedChips() {
  const teams = Object.values(data)
  const featHL = coord.featuredHighlights || [],
    featBL = coord.featuredBlockers || []

  document.getElementById('feat-hl-chips').innerHTML =
    teams
      .filter((t) => t.highlight)
      .map(
        (t) =>
          `<button type="button" class="featured-chip ${featHL.includes(t.team) ? 'on' : ''}" onclick="toggleFeat('hl','${esc(t.team)}')" aria-pressed="${featHL.includes(t.team)}">${esc(t.team)}</button>`,
      )
      .join('') ||
    '<span style="font-size:12px;color:var(--text3)">No highlights yet</span>'

  document.getElementById('feat-bl-chips').innerHTML =
    teams
      .filter((t) => t.blocker)
      .map(
        (t) =>
          `<button type="button" class="featured-chip ${featBL.includes(t.team) ? 'on' : ''}" onclick="toggleFeat('bl','${esc(t.team)}')" aria-pressed="${featBL.includes(t.team)}">${esc(t.team)}</button>`,
      )
      .join('') ||
    '<span style="font-size:12px;color:var(--text3)">No blockers yet</span>'
}

// Toggle a team in the featured highlights or blockers list.
// @param {string} type - 'hl' for highlights, 'bl' for blockers
// @param {string} team - Team name
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

// ── MEETINGS EDITOR ───────────────────────────────────────────────────────────
// Adds a meeting to coord.meetings and triggers a save.
function addMeeting() {
  const title = document.getElementById('mtg-title').value.trim()
  const time = document.getElementById('mtg-time').value.trim()
  const link = document.getElementById('mtg-link').value.trim()
  if (!title) {
    showToast('Please enter a meeting title')
    return
  }
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

// Remove a meeting by its id (a Date.now() timestamp).
function removeMeeting(id) {
  coord.meetings = (coord.meetings || []).filter((m) => m.id !== id)
  debounceSave()
  renderMeetingsList()
  updateSummary()
}

// Render the editable meeting list inside the coordinator panel.
// The read-only display in the summary banner is rendered by
// renderMeetingsDisplay() in render.js.
function renderMeetingsList() {
  const meetings = coord.meetings || [],
    el = document.getElementById('meetings-list')
  if (!el) return
  el.innerHTML = meetings.length
    ? meetings
        .map(
          (m) =>
            `<div style="display:flex;align-items:center;gap:8px;background:white;border:1px solid #bfdbfe;border-radius:var(--radius-sm);padding:6px 10px;font-size:12px;">
        <span style="flex:1;font-weight:600">${esc(m.title)}</span>
        ${m.time ? `<span style="color:var(--text3);font-size:11px">${esc(m.time)}</span>` : ''}
        ${m.link ? `<a href="${esc(m.link)}" target="_blank" style="font-size:11px;color:var(--link)">🔗 Link</a>` : ''}
        <button type="button" onclick="removeMeeting(${m.id})" style="background:none;border:none;color:var(--text3);cursor:pointer;font-size:14px;line-height:1;padding:2px 4px;" aria-label="Remove meeting ${esc(m.title)}">✕</button>
      </div>`,
        )
        .join('')
    : '<span style="font-size:12px;color:var(--text3)">No meetings added yet</span>'
}

// ── CSV UPLOAD (non-SP mode only) ─────────────────────────────────────────────
// Parse a CSV file exported from SharePoint or Excel and load it into `data`.
// CSV controls are hidden when CONFIG.useSharePoint is true.
function uploadCSV(event) {
  const file = event.target.files[0]
  if (!file) return
  const reader = new FileReader()
  reader.onload = function (e) {
    const text = e.target.result.replace(/^\uFEFF/, '') // strip BOM
    const lines = text.split(/\r?\n/).filter((l) => l.trim())
    const headers = lines[0].split(',').map((h) => h.trim().replace(/"/g, ''))
    let count = 0
    for (let i = 1; i < lines.length; i++) {
      // Minimal CSV parser that respects quoted fields
      const cols = []
      let current = '',
        inQuotes = false
      for (let c = 0; c < lines[i].length; c++) {
        const ch = lines[i][c]
        if (ch === '"') {
          inQuotes = !inQuotes
        } else if (ch === ',' && !inQuotes) {
          cols.push(current.trim())
          current = ''
        } else {
          current += ch
        }
      }
      cols.push(current.trim())

      const row = {}
      headers.forEach((h, j) => (row[h] = (cols[j] || '').trim()))
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

// Clear all team data for the current week (non-SP mode only).
function clearAll() {
  if (!confirm('Clear all team data for this week? This cannot be undone.'))
    return
  data = {}
  localStorage.removeItem('sitrep_v2')
  renderAll()
  showToast('Week cleared')
}

// ── TEAM ROSTER ───────────────────────────────────────────────────────────────
// Return the current team roster: DEFAULT_TEAMS + teams with data + all teams
// from SharePoint history and column choices (allTeamNames). Deduplicated and
// sorted case-insensitively. This is the single source of truth for which
// teams appear in the grid, modal dropdowns, and deps picker.
function getRosterTeams() {
  return [
    ...new Set([...DEFAULT_TEAMS, ...Object.keys(data), ...allTeamNames]),
  ].sort((a, b) => a.localeCompare(b, undefined, { sensitivity: 'base' }))
}

// Render the read-only team list inside the coordinator panel (info only).
function renderTeamsList() {
  const listHtml = getRosterTeams()
    .map(
      (team) =>
        `<div style="display:inline-flex;align-items:center;gap:6px;padding:6px 10px;background:#eef4ff;border-radius:999px;font-size:13px;white-space:nowrap;">
          <span>${esc(team)}</span>
        </div>`,
    )
    .join('')
  document.getElementById('teams-list').innerHTML =
    listHtml ||
    '<div style="font-size:12px;color:var(--text3);">No teams available</div>'
}

// Populate the team <select> inside the edit modal with the current roster.
// Called every time the modal opens.
function buildTeamSelect() {
  const select = document.getElementById('f-team')
  if (!select) return
  select.innerHTML =
    '<option value="">Select team...</option>' +
    getRosterTeams()
      .map((team) => `<option value="${esc(team)}">${esc(team)}</option>`)
      .join('')
}

// ── TEAM MANAGEMENT ──────────────────────────────────────────────────────────
// Team additions and removals are managed manually through SharePoint and the
// MS Form — the dashboard has no write access to list schema or form questions.
// The "Add Team" button in the coordinator panel shows these instructions.
function openAddTeamModal() {
  showManualTeamInstructions()
}

function showManualTeamInstructions() {
  showErrorModal(
    'Team Management is Manual',
    `Team additions, renames and removals are managed directly in SharePoint and the MS Form.<br><br>` +
      `<ul style="padding-left:18px;margin:8px 0;line-height:1.8;">` +
      `<li>Update the MS Form team selection question: <a href="https://forms.office.com/r/5qzNa4JpH9" target="_blank">Edit the form</a></li>` +
      `<li>Add the new team to the SharePoint <strong>TeamName</strong> choice list: <a href="https://bcgov.sharepoint.com/teams/12320-ConnectedServicesStrategicPriority/Lists/Weekly%20SitRep%20Data/AllItems.aspx" target="_blank">Open SharePoint list</a></li>` +
      `<li>Add the new team to the SharePoint <strong>DependenciesIn</strong> choice list</li>` +
      `<li>After those updates, refresh the dashboard — the new team will appear automatically</li></ul>`,
    '',
  )
}
