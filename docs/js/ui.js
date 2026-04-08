// ── UI ───────────────────────────────────────────────────────────────────────

// ── INIT ──────────────────────────────────────────────────────────────────────
function init() {
  document.getElementById('week-chip').textContent = 'Week of ' + getWeekLabel()
  buildDepsPicker()
  if (CONFIG.useSharePoint) {
    document.getElementById('csv-controls').style.display = 'none'
    if (window.msal) {
      loadFromSharePoint()
    } else {
      const c = setInterval(() => {
        if (window.msal) {
          clearInterval(c)
          loadFromSharePoint()
        }
      }, 100)
    }
    // Start smart auto-refresh: check for changes every minute
    startAutoRefresh()
  } else {
    loadLocalDataAndSync()
    renderAll()
  }
}

// ── AUTO-REFRESH ──────────────────────────────────────────────────────────────
let lastDataHash = null
let lastCoordHash = null

function startAutoRefresh() {
  if (autoRefreshTimer) clearInterval(autoRefreshTimer)
  autoRefreshTimer = setInterval(async () => {
    if (CONFIG.useSharePoint && document.visibilityState === 'visible') {
      try {
        await checkForUpdates()
      } catch (e) {
        console.warn('Auto-refresh check failed:', e)
      }
    }
  }, 60000) // Check every minute instead of 30 seconds
  console.log('[AUTO-REFRESH] Started (60s interval, change detection)')
}

async function checkForUpdates() {
  try {
    const token = await getToken()
    if (!token || !_siteId || !_teamListId) return

    // Quick check: get latest items count and modified date
    const itemsResp = await fetch(
      `https://graph.microsoft.com/v1.0/sites/${_siteId}/lists/${_teamListId}/items?$top=1&$orderby=lastModifiedDateTime desc&$select=lastModifiedDateTime,id`,
      { headers: { Authorization: `Bearer ${token}` } },
    )

    if (!itemsResp.ok) return // Skip if we can't check

    const items = await itemsResp.json()
    const latestModified = items.value?.[0]?.lastModifiedDateTime

    // Check coordinator list too
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

    // If we have new data, do a full refresh
    const currentHash = `${latestModified || ''}|${coordModified || ''}`
    const previousHash = `${lastDataHash || ''}|${lastCoordHash || ''}`

    if (currentHash !== previousHash) {
      console.log('[AUTO-REFRESH] Changes detected, refreshing...')
      lastDataHash = latestModified
      lastCoordHash = coordModified
      await loadFromSharePoint()
    }
  } catch (e) {
    // Silent fail for auto-refresh checks
  }
}

function stopAutoRefresh() {
  if (autoRefreshTimer) {
    clearInterval(autoRefreshTimer)
    autoRefreshTimer = null
    console.log('[AUTO-REFRESH] Stopped')
  }
}

// Stop auto-refresh if page becomes hidden, restart if visible
document.addEventListener('visibilitychange', () => {
  if (CONFIG.useSharePoint) {
    if (document.hidden) stopAutoRefresh()
    else startAutoRefresh()
  }
})

// ── FILTER ────────────────────────────────────────────────────────────────────
function setFilter(f, btn) {
  currentFilter = f
  document
    .querySelectorAll('.filter-btn')
    .forEach((b) => b.setAttribute('aria-pressed', 'false'))
  btn.setAttribute('aria-pressed', 'true')
  renderGrid()
}

// ── MODAL ─────────────────────────────────────────────────────────────────────
function openModal(teamName) {
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
  selectedRYG = t?.status || null
  ;['green', 'yellow', 'red'].forEach((x) => {
    const btn = document.getElementById('r' + x[0])
    btn.className = 'ryg-opt' + (x === selectedRYG ? ' sel-' + x : '')
    btn.setAttribute('aria-pressed', (x === selectedRYG).toString())
  })
  selectedDeps = [...(t?.depsIn || [])]
  TEAMS.forEach((tm) => {
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

async function saveTeam() {
  const team = document.getElementById('f-team').value
  const highlight = document.getElementById('f-highlight').value.trim()

  // Validate required fields
  const missing = []
  if (!team) missing.push('Team name')
  if (!selectedRYG) missing.push('Status (Red/Yellow/Green)')
  if (!highlight) missing.push('Key Highlight')

  if (missing.length > 0) {
    showErrorModal(
      'Required Fields Missing',
      `Please fill in all required fields before saving:\n\n${missing.map((f) => '• ' + f).join('\n')}`,
      '',
    )
    return
  }
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
  console.log('[SAVE] Attempting to save:', team, teamData)

  if (!CONFIG.useSharePoint) {
    console.log('[SAVE] Using localStorage (SharePoint disabled)')
    localStorage.setItem('sitrep_v2', JSON.stringify(data))
    renderAll()
    renderFeaturedChips()
    showToast('✓ ' + team + ' saved locally!')
    setTimeout(() => closeModal(), 1200)
  } else {
    console.log('[SAVE] Calling saveTeamToSharePoint for SharePoint...')
    try {
      await saveTeamToSharePoint(team, teamData)
      console.log('[SAVE] Save succeeded')
      // Only update UI and close on success
      renderAll()
      renderFeaturedChips()
      showToast('✓ ' + team + ' saved to SharePoint!')
      setTimeout(() => closeModal(), 1200)
    } catch (e) {
      console.error('[SAVE] Save failed:', e)
      // Error modal is shown inside saveTeamToSharePoint
      // Keep modal open so user can see error and try again
    }
  }
}

// Focus trap in modal
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
function pickRYG(c) {
  selectedRYG = c
  ;['green', 'yellow', 'red'].forEach((x) => {
    const btn = document.getElementById('r' + x[0])
    btn.className = 'ryg-opt' + (x === c ? ' sel-' + c : '')
    btn.setAttribute('aria-pressed', (x === c).toString())
  })
}

// ── DEPS PICKER ───────────────────────────────────────────────────────────────
function buildDepsPicker() {
  document.getElementById('deps-picker').innerHTML = TEAMS.map(
    (t) =>
      `<button type="button" class="featured-chip" id="dep-${t.replace(/\//g, '-')}" onclick="toggleDep('${esc(t)}')" aria-pressed="false">${esc(t)}</button>`,
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

function saveCoord() {
  coord.status = document.getElementById('coord-status').value
  coord.news = document.getElementById('coord-news').value
  debounceSave()
  updateSummary()
}

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

function removeMeeting(id) {
  coord.meetings = (coord.meetings || []).filter((m) => m.id !== id)
  debounceSave()
  renderMeetingsList()
  updateSummary()
}

function renderMeetingsList() {
  const meetings = coord.meetings || [],
    el = document.getElementById('meetings-list')
  if (!el) return
  el.innerHTML = meetings.length
    ? meetings
        .map(
          (
            m,
          ) => `<div style="display:flex;align-items:center;gap:8px;background:white;border:1px solid #bfdbfe;border-radius:var(--radius-sm);padding:6px 10px;font-size:12px;">
        <span style="flex:1;font-weight:600">${esc(m.title)}</span>
        ${m.time ? `<span style="color:var(--text3);font-size:11px">${esc(m.time)}</span>` : ''}
        ${m.link ? `<a href="${esc(m.link)}" target="_blank" style="font-size:11px;color:var(--link)">🔗 Link</a>` : ''}
        <button type="button" onclick="removeMeeting(${m.id})" style="background:none;border:none;color:var(--text3);cursor:pointer;font-size:14px;line-height:1;padding:2px 4px;" aria-label="Remove meeting ${esc(m.title)}">✕</button>
      </div>`,
        )
        .join('')
    : '<span style="font-size:12px;color:var(--text3)">No meetings added yet</span>'
}

// ── CSV UPLOAD ────────────────────────────────────────────────────────────────
function uploadCSV(event) {
  const file = event.target.files[0]
  if (!file) return
  const reader = new FileReader()
  reader.onload = function (e) {
    const text = e.target.result.replace(/^\uFEFF/, '')
    const lines = text.split(/\r?\n/).filter((l) => l.trim())
    const headers = lines[0].split(',').map((h) => h.trim().replace(/"/g, ''))
    let count = 0
    for (let i = 1; i < lines.length; i++) {
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

function clearAll() {
  if (!confirm('Clear all team data for this week? This cannot be undone.'))
    return
  data = {}
  localStorage.removeItem('sitrep_v2')
  renderAll()
  showToast('Week cleared')
}

// ── TEAM MANAGEMENT (COORDINATOR PANEL) ──────────────────────────────────────
function renderTeamsList() {
  const listHtml = TEAMS.map(
    (team) =>
      `<div style="display: flex; gap: 8px; align-items: center; padding: 6px 0; border-bottom: 1px solid #ddd;">
        <span style="flex: 1; font-size: 13px;">${esc(team)}</span>
        <button type="button" style="padding: 4px 8px; font-size: 11px; background: #f3e5e5; color: #c41c3b; border: none; border-radius: 4px; cursor: pointer;" onclick="removeTeamFromPanel('${esc(team)}')">Remove</button>
      </div>`,
  ).join('')
  document.getElementById('teams-list').innerHTML =
    listHtml ||
    '<div style="font-size: 12px; color: var(--text3); padding: 10px 0;">No teams to manage</div>'
}

function openAddTeamModal() {
  const newTeamName = prompt('Enter new team name:', '')
  if (!newTeamName || !newTeamName.trim()) return
  addTeamToSharePoint(newTeamName.trim())
}

async function addTeamToSharePoint(teamName) {
  try {
    console.log('[TEAM-MGMT] Adding team:', teamName)
    const token = await getToken()
    if (!token || !_siteId) throw new Error('Not authenticated')

    // Check if team already exists
    if (TEAMS.includes(teamName)) {
      showErrorModal(
        'Team Already Exists',
        `The team "${teamName}" already exists in the system.`,
        '',
      )
      return
    }

    // Add to TEAMS array
    TEAMS.push(teamName)
    TEAMS.sort() // Keep alphabetical

    // Force UI rebuild
    buildDepsPicker()
    renderTeamsList()
    renderFeaturedChips()
    renderGrid()

    // Sync to SharePoint
    await saveTeamsRegistry(token, TEAMS)

    showErrorModal(
      '✅ Team Added Successfully',
      `"${teamName}" has been added and is now available to all users.\n\n⚠️ Important Next Steps:\n\n• Update the MS Form to include "${teamName}" in the team selection question\n• Add "${teamName}" to dependency fields if applicable\n• The change will sync to all connected users automatically`,
      '',
    )

    console.log('[TEAM-MGMT] Team added and synced:', teamName)
  } catch (e) {
    console.error('[TEAM-MGMT] Error adding team:', e)
    showErrorModal(
      'Failed to Add Team',
      `Could not add "${teamName}": ${e.message}`,
      '',
    )
  }
}

async function removeTeamFromPanel(teamName) {
  if (
    !confirm(
      `Are you sure you want to remove "${teamName}"?\n\nThis will hide it from future submissions but existing data will remain.`,
    )
  )
    return

  try {
    console.log('[TEAM-MGMT] Removing team:', teamName)
    const token = await getToken()
    if (!token || !_siteId) throw new Error('Not authenticated')

    // Remove from TEAMS array
    const idx = TEAMS.indexOf(teamName)
    if (idx >= 0) TEAMS.splice(idx, 1)

    // Force UI rebuild
    buildDepsPicker()
    renderTeamsList()
    renderFeaturedChips()
    renderGrid()

    // Sync to SharePoint
    await saveTeamsRegistry(token, TEAMS)

    showErrorModal(
      '✅ Team Removed',
      `"${teamName}" has been removed from the active team list.\n\n⚠️ Important: Update the MS Form to remove "${teamName}" from team selection.`,
      '',
    )

    console.log('[TEAM-MGMT] Team removed and synced:', teamName)
  } catch (e) {
    console.error('[TEAM-MGMT] Error removing team:', e)
    showErrorModal(
      'Failed to Remove Team',
      `Could not remove "${teamName}": ${e.message}`,
      '',
    )
  }
}

async function saveTeamsRegistry(token, teamsList) {
  // Create or update a registry item in SharePoint that tracks all teams
  // This allows other users/sessions to sync the team list
  if (!_siteId || !_teamListId) return

  const fields = {
    TeamName: '__TEAMS_REGISTRY__',
    OverallStatus: 'Yellow', // Placeholder
    Highlight: JSON.stringify(teamsList), // Store team list as JSON
    Blocker: new Date().toISOString(), // Track last update
  }

  try {
    // Try to find existing registry item
    const searchResp = await fetch(
      `https://graph.microsoft.com/v1.0/sites/${_siteId}/lists/${_teamListId}/items?$filter=fields/TeamName eq '__TEAMS_REGISTRY__'`,
      { headers: { Authorization: `Bearer ${token}` } },
    )
    const search = await searchResp.json()

    if (search.value?.length > 0) {
      // Update existing
      const itemId = search.value[0].id
      const updateResp = await fetch(
        `https://graph.microsoft.com/v1.0/sites/${_siteId}/lists/${_teamListId}/items/${itemId}/fields`,
        {
          method: 'PATCH',
          headers: {
            Authorization: `Bearer ${token}`,
            'Content-Type': 'application/json',
          },
          body: JSON.stringify(fields),
        },
      )
      if (!updateResp.ok) throw new Error(`PATCH failed ${updateResp.status}`)
    } else {
      // Create new
      const createResp = await fetch(
        `https://graph.microsoft.com/v1.0/sites/${_siteId}/lists/${_teamListId}/items`,
        {
          method: 'POST',
          headers: {
            Authorization: `Bearer ${token}`,
            'Content-Type': 'application/json',
          },
          body: JSON.stringify({ fields }),
        },
      )
      if (!createResp.ok) throw new Error(`POST failed ${createResp.status}`)
    }
  } catch (e) {
    console.warn('[TEAM-MGMT] Could not save teams registry:', e.message)
    // Non-critical, continue anyway
  }
}

async function syncTeamsFromSharePoint(token) {
  // On load, check if there's an updated team list in SharePoint
  if (!_siteId || !_teamListId) return

  try {
    const searchResp = await fetch(
      `https://graph.microsoft.com/v1.0/sites/${_siteId}/lists/${_teamListId}/items?$filter=fields/TeamName eq '__TEAMS_REGISTRY__'&$expand=fields`,
      { headers: { Authorization: `Bearer ${token}` } },
    )
    const search = await searchResp.json()

    if (search.value?.length > 0) {
      const registryItem = search.value[0]
      const highlight = registryItem.fields?.Highlight
      if (highlight) {
        try {
          const syncedTeams = JSON.parse(highlight)
          if (Array.isArray(syncedTeams)) {
            TEAMS = syncedTeams
            console.log('[TEAM-MGMT] Synced teams from SharePoint:', TEAMS)
            buildDepsPicker()
          }
        } catch (e) {
          console.warn('[TEAM-MGMT] Could not parse synced teams:', e)
        }
      }
    }
  } catch (e) {
    console.warn('[TEAM-MGMT] Could not sync teams from SharePoint:', e.message)
  }
}
