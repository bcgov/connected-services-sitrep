// ── SHAREPOINT ───────────────────────────────────────────────────────────────

// Show/hide dashboard chrome depending on auth state
function showFullPageState(state) {
  const summary = document.querySelector('.summary')
  const coordBar = document.getElementById('coord-bar')
  const filterBar = document.querySelector('.filter-bar')
  const grid = document.getElementById('grid')
  const authenticated = state === 'loaded'

  if (summary) summary.style.display = authenticated ? '' : 'none'
  if (coordBar) coordBar.style.display = authenticated ? '' : 'none'
  if (filterBar) filterBar.style.display = authenticated ? '' : 'none'
  document
    .querySelectorAll('.section-heading')
    .forEach((el) => (el.style.display = authenticated ? '' : 'none'))

  if (state === 'loading') {
    grid.innerHTML = `<div style="grid-column:1/-1;text-align:center;padding:80px 20px;" role="status" aria-label="Loading">
      <div style="width:32px;height:32px;border:3px solid var(--border);border-top-color:var(--primary);border-radius:50%;animation:spin 0.8s linear infinite;margin:0 auto 16px" aria-hidden="true"></div>
      <div style="font-size:14px;color:var(--text2)">Connecting to SharePoint...</div>
    </div>`
  } else if (state === 'signin') {
    grid.innerHTML = `<div style="grid-column:1/-1;display:flex;flex-direction:column;align-items:center;justify-content:center;min-height:60vh;text-align:center;padding:40px 20px;">
      <img src="images/BCID_H_RGB_pos.png" alt="" style="height:48px;margin-bottom:32px;opacity:0.8" aria-hidden="true"/>
      <h2 style="font-size:22px;font-weight:700;color:var(--text);margin-bottom:8px;letter-spacing:-0.02em;">Connected Services SitRep</h2>
      <p style="font-size:14px;color:var(--text2);margin-bottom:32px;max-width:380px;line-height:1.6;">Sign in with your BC Gov Microsoft 365 account to view and contribute to the weekly status dashboard.</p>
      <button onclick="signIn()" style="font-family:'BCSans',sans-serif;font-size:14px;font-weight:700;padding:14px 32px;border-radius:var(--radius-pill);border:none;background:var(--primary);color:white;cursor:pointer;box-shadow:0 2px 8px rgba(1,51,102,0.3);" aria-label="Sign in with BC Gov account">
        Sign in with BC Gov account →
      </button>
      <p style="font-size:11px;color:var(--text3);margin-top:20px;">Your data is stored securely in SharePoint</p>
    </div>`
  }
}

function setGridLoading() {
  showFullPageState('loading')
}

// Load local data and attempt to sync unsaved changes
async function loadLocalDataAndSync() {
  // Load any locally saved team data from localStorage keys
  Object.keys(localStorage)
    .filter((key) => key.startsWith('sitrep_team_'))
    .forEach((key) => {
      const team = key.replace(/^sitrep_team_/, '')
      const localData = localStorage.getItem(key)
      if (!localData) return
      try {
        const parsed = JSON.parse(localData)
        if (parsed._localOnly) {
          data[team] = parsed
          console.log(`[LOCAL] Loaded unsaved data for ${team}`)
        }
      } catch (e) {
        console.warn(`Failed to parse local data for ${team}:`, e)
      }
    })

  // Load coordinator data
  const coordData = localStorage.getItem('sitrep_coord')
  if (coordData) {
    try {
      coord = JSON.parse(coordData)
      console.log('[LOCAL] Loaded coordinator data')
    } catch (e) {
      console.warn('Failed to parse local coordinator data:', e)
    }
  }

  // Try to sync local data if we have SharePoint access
  if (CONFIG.useSharePoint && _siteId && _teamListId) {
    await syncLocalData()
  }
}

// Attempt to sync locally saved data to SharePoint
async function syncLocalData() {
  const teamsToSync = Object.keys(data).filter((team) => data[team]._localOnly)

  for (const team of teamsToSync) {
    try {
      console.log(`[SYNC] Attempting to sync ${team}...`)
      await saveTeamToSharePoint(team, data[team])
      // If successful, remove local flag and localStorage
      delete data[team]._localOnly
      localStorage.removeItem('sitrep_team_' + team)
      console.log(`[SYNC] Successfully synced ${team}`)
    } catch (e) {
      console.warn(`[SYNC] Failed to sync ${team}, keeping local:`, e.message)
    }
  }
}

// ── LOAD TEAM DATA ────────────────────────────────────────────────────────────
async function loadFromSharePoint() {
  setGridLoading()
  try {
    const initResult = await initMsal()
    if (!initResult) {
      await loadLocalDataAndSync()
      renderAll()
      return
    }

    // Not signed in → show sign in screen
    const accounts = msalInstance.getAllAccounts()
    if (accounts.length === 0) {
      showFullPageState('signin')
      return
    }

    let token = typeof initResult === 'string' ? initResult : await getToken()
    if (!token) return

    // Resolve site
    const siteResp = await fetch(
      `https://graph.microsoft.com/v1.0/sites/bcgov.sharepoint.com:/teams/12320-ConnectedServicesStrategicPriority`,
      { headers: { Authorization: `Bearer ${token}` } },
    )
    if (!siteResp.ok) throw new Error('Cannot access SharePoint site')
    const site = await siteResp.json()
    _siteId = site.id

    // Resolve team data list
    const listsResp = await fetch(
      `https://graph.microsoft.com/v1.0/sites/${_siteId}/lists?$filter=displayName eq '${CONFIG.listName}'`,
      { headers: { Authorization: `Bearer ${token}` } },
    )
    const lists = await listsResp.json()
    if (!lists.value?.length)
      throw new Error(`List "${CONFIG.listName}" not found`)
    _teamListId = lists.value[0].id

    // Fetch all items
    const itemsResp = await fetch(
      `https://graph.microsoft.com/v1.0/sites/${_siteId}/lists/${_teamListId}/items?$expand=fields&$top=500`,
      { headers: { Authorization: `Bearer ${token}` } },
    )
    const items = await itemsResp.json()
    allHistoryItems = items.value || []

    // Extract all unique team names from the entire history
    const teamNamesSet = new Set()
    allHistoryItems.forEach((item) => {
      const teamName = item.fields?.TeamName
      if (teamName) teamNamesSet.add(teamName)
    })
    allTeamNames = Array.from(teamNamesSet).sort((a, b) =>
      a.localeCompare(b, undefined, { sensitivity: 'base' }),
    )
    console.log('[SP-LOAD] Found teams in SharePoint history:', allTeamNames)

    // Latest item per team for THIS week only.

    // Belongs to this week if:
    //   (a) WeekOf field matches this week label (dashboard edits / week-picker), OR
    //   (b) no WeekOf and created on or after this Monday 00:00 local time (form submissions)
    const currentWeek = getWeekLabel()
    const weekStart = getWeekStart()
    const byTeam = {}

    // Parse WeekOf to a Date for comparison (handles both YYYY-MM-DD and DD-MM-YYYY formats)
    function parseWeekOf(weekOfStr) {
      if (!weekOfStr) return null
      const parts = weekOfStr.split('-')
      if (parts.length !== 3) return null
      // Try YYYY-MM-DD first
      if (parts[0].length === 4) {
        return new Date(parts[0], parseInt(parts[1]) - 1, parts[2])
      }
      // Otherwise assume DD-MM-YYYY
      return new Date(parts[2], parseInt(parts[1]) - 1, parts[0])
    }

    allHistoryItems.forEach((item) => {
      const f = item.fields,
        team = f.TeamName
      if (!team) return
      const created = new Date(f.Created || item.createdDateTime)
      const weekOfDate = parseWeekOf(f.WeekOf)
      const currentWeekDate = parseWeekOf(currentWeek)

      // Include if: WeekOf matches this week, OR entry was created this week
      const belongsToThisWeek =
        (weekOfDate &&
          currentWeekDate &&
          weekOfDate.getTime() === currentWeekDate.getTime()) ||
        created >= weekStart

      if (!belongsToThisWeek) return
      if (!byTeam[team]) {
        byTeam[team] = { fields: f, created, id: item.id }
      } else if (created > byTeam[team].created) {
        byTeam[team] = { fields: f, created, id: item.id }
      }
    })

    data = {}
    Object.entries(byTeam).forEach(([team, { fields: f, id }]) => {
      data[team] = {
        team,
        status: (f.OverallStatus || 'yellow').toLowerCase(),
        highlight: f.Highlight || '',
        blocker: f.Blocker || '',
        initiativeNum: f.InitiativeNumber || '',
        escalatorNum: f.EscalatorNumber || '',
        depsIn: parseDepsIn(f.DependenciesIn),
        summary: f.WeekSummary || '',
        _spId: id,
      }
    })

    renderAll()
    showToast(`✓ Loaded ${Object.keys(data).length} teams from SharePoint`)
    await loadCoordFromSharePoint(token)
    showFullPageState('loaded')
  } catch (err) {
    showFullPageState('signin')
  }
}

// ── SAVE TEAM TO SHAREPOINT ───────────────────────────────────────────────────
async function saveTeamToSharePoint(team, teamData) {
  const statusEl = document.getElementById('modal-save-status')
  if (statusEl) statusEl.textContent = 'Saving...'

  console.log('[SP-SAVE] Starting save for team:', team)
  console.log('[SP-SAVE] CONFIG.useSharePoint:', CONFIG.useSharePoint)
  console.log('[SP-SAVE] _siteId:', _siteId ? 'present' : 'MISSING')
  console.log('[SP-SAVE] _teamListId:', _teamListId ? 'present' : 'MISSING')

  if (!CONFIG.useSharePoint) {
    console.warn('[SP-SAVE] SharePoint disabled, using localStorage')
    throw new Error('SharePoint disabled in CONFIG')
  }
  if (!_siteId || !_teamListId) {
    console.error(
      '[SP-SAVE] SharePoint IDs not initialized. _siteId:',
      _siteId,
      '_teamListId:',
      _teamListId,
    )
    throw new Error(
      'SharePoint not initialized - _siteId or _teamListId missing. Try refreshing the page.',
    )
  }
  let token = null
  try {
    token = await getToken()
    const fields = {
      TeamName: team,
      // Note: WeekOf is read-only in SharePoint, cannot be set directly
      // It's managed by the form submission process
      OverallStatus:
        teamData.status.charAt(0).toUpperCase() + teamData.status.slice(1),
      Highlight: teamData.highlight,
      Blocker: teamData.blocker,
      InitiativeNumber: teamData.initiativeNum,
      EscalatorNumber: teamData.escalatorNum,
      DependenciesIn: JSON.stringify(teamData.depsIn || []),
      WeekSummary: teamData.summary,
    }
    if (teamData._spId) {
      const patchResp = await fetch(
        `https://graph.microsoft.com/v1.0/sites/${_siteId}/lists/${_teamListId}/items/${teamData._spId}/fields`,
        {
          method: 'PATCH',
          headers: {
            Authorization: `Bearer ${token}`,
            'Content-Type': 'application/json',
          },
          body: JSON.stringify(fields),
        },
      )
      if (!patchResp.ok) {
        const errDetail = await patchResp.text()
        throw new Error(`PATCH failed ${patchResp.status}: ${errDetail}`)
      }
    } else {
      const postResp = await fetch(
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
      if (!postResp.ok) {
        const errDetail = await postResp.text()
        throw new Error(`POST failed ${postResp.status}: ${errDetail}`)
      }
      const created = await postResp.json()
      if (created.id) {
        data[team]._spId = created.id
      } else {
        throw new Error('No item ID returned from SharePoint')
      }
    }
    console.log('[SP-SAVE] Success for team:', team)
    if (statusEl) statusEl.textContent = '✓ Saved to SharePoint'
    setTimeout(() => {
      if (statusEl) statusEl.textContent = ''
    }, 2500)
  } catch (e) {
    console.error('[SP-SAVE] ERROR for team:', team, e)
    const userMsg = e.message.includes('429')
      ? 'Too many requests (server busy). Try again in a moment.'
      : e.message.includes('read-only')
        ? 'SharePoint field configuration issue (WeekOf is read-only). Contact IT support.'
        : e.message.includes('401')
          ? 'Authentication expired. Please refresh the page.'
          : e.message.includes('403')
            ? 'Permission denied. Check your access or contact support.'
            : `Save failed: ${e.message}`

    // Fallback to localStorage
    const localData = {
      team,
      status: teamData.status,
      highlight: teamData.highlight,
      blocker: teamData.blocker,
      initiativeNum: teamData.initiativeNum,
      escalatorNum: teamData.escalatorNum,
      depsIn: teamData.depsIn || [],
      summary: teamData.summary,
      _weekOf: teamData._weekOf || getWeekLabel(),
      _localOnly: true, // Flag to indicate this is local-only
      _savedAt: new Date().toISOString(),
    }
    data[team] = localData
    localStorage.setItem('sitrep_team_' + team, JSON.stringify(localData))
    renderGrid() // Update UI immediately

    const authInfo = debugAuth()
    const details = `Error: ${e.message}
Token: ${token ? 'Present' : 'Missing'}
User: ${authInfo.account || 'Unknown'}
Signed In: ${authInfo.signedIn}
Scopes: ${authInfo.scopes.join(', ')}
Team: ${team}
Time: ${new Date().toISOString()}

Troubleshooting:
• If you can see the dashboard but can't save, you may have read-only permissions
• The 'WeekOf' field appears to be read-only - this is a SharePoint list configuration issue
• Try refreshing the page to re-authenticate
• Contact support with this error details`
    showErrorModal(
      'Team Save Failed',
      `${userMsg} Data saved locally and will sync when connection is restored.`,
      details,
    )

    if (statusEl) statusEl.textContent = '⚠ Saved locally'
    setTimeout(() => {
      if (statusEl) statusEl.textContent = ''
    }, 3000)
    throw e // re-throw so saveTeam() knows the SP save failed
  }
}

// ── COORDINATOR SHAREPOINT ────────────────────────────────────────────────────
async function loadCoordFromSharePoint(token) {
  try {
    if (!token) token = await getToken()
    if (!token || !_siteId) return

    const listsResp = await fetch(
      `https://graph.microsoft.com/v1.0/sites/${_siteId}/lists?$filter=displayName eq '${CONFIG.coordListName}'`,
      { headers: { Authorization: `Bearer ${token}` } },
    )
    const lists = await listsResp.json()
    if (!lists.value?.length) {
      console.warn('SitRep Coordinator list not found')
      return
    }
    _coordListId = lists.value[0].id

    const weekLabel = getWeekLabel()
    const itemsResp = await fetch(
      `https://graph.microsoft.com/v1.0/sites/${_siteId}/lists/${_coordListId}/items?$expand=fields&$top=50`,
      { headers: { Authorization: `Bearer ${token}` } },
    )
    const items = await itemsResp.json()
    const match = (items.value || []).find((i) => i.fields.WeekOf === weekLabel)

    if (match) {
      coordItemId = match.id
      const f = match.fields
      coord = {
        status: f.OverallStatus || '',
        news: f.News || '',
        meetings: safeJSON(f.Meetings, []),
        featuredHighlights: safeJSON(f.FeaturedHighlights, []),
        featuredBlockers: safeJSON(f.FeaturedBlockers, []),
      }
    } else {
      coordItemId = null
      coord = {}
    }
    updateSummary()
    if (coordOpen) {
      document.getElementById('coord-status').value = coord.status || ''
      document.getElementById('coord-news').value = coord.news || ''
      renderFeaturedChips()
      renderMeetingsList()
    }
  } catch (e) {
    console.warn('Could not load coordinator data:', e)
  }
}

async function saveCoordToSharePoint() {
  const savingEl = document.getElementById('coord-saving')
  if (savingEl) savingEl.textContent = 'Saving...'
  try {
    const token = await getToken()
    if (!token || !_siteId || !_coordListId) {
      localStorage.setItem('sitrep_coord', JSON.stringify(coord))
      if (savingEl) savingEl.textContent = ''
      return
    }
    const fields = {
      WeekOf: getWeekLabel(),
      OverallStatus: coord.status || '',
      News: coord.news || '',
      Meetings: JSON.stringify(coord.meetings || []),
      FeaturedHighlights: JSON.stringify(coord.featuredHighlights || []),
      FeaturedBlockers: JSON.stringify(coord.featuredBlockers || []),
    }
    if (coordItemId) {
      const patchResp = await fetch(
        `https://graph.microsoft.com/v1.0/sites/${_siteId}/lists/${_coordListId}/items/${coordItemId}/fields`,
        {
          method: 'PATCH',
          headers: {
            Authorization: `Bearer ${token}`,
            'Content-Type': 'application/json',
          },
          body: JSON.stringify(fields),
        },
      )
      if (!patchResp.ok) {
        const errDetail = await patchResp.text()
        throw new Error(`PATCH failed ${patchResp.status}: ${errDetail}`)
      }
    } else {
      const postResp = await fetch(
        `https://graph.microsoft.com/v1.0/sites/${_siteId}/lists/${_coordListId}/items`,
        {
          method: 'POST',
          headers: {
            Authorization: `Bearer ${token}`,
            'Content-Type': 'application/json',
          },
          body: JSON.stringify({ fields }),
        },
      )
      if (!postResp.ok) {
        const errDetail = await postResp.text()
        throw new Error(`POST failed ${postResp.status}: ${errDetail}`)
      }
      const created = await postResp.json()
      if (created.id) {
        coordItemId = created.id
      } else {
        throw new Error('No item ID returned from SharePoint')
      }
    }
    if (savingEl) savingEl.textContent = '✓ Saved'
    setTimeout(() => {
      if (savingEl) savingEl.textContent = ''
    }, 2000)
  } catch (e) {
    console.error('ERROR: Could not save coordinator data:', e.message, e)
    localStorage.setItem('sitrep_coord', JSON.stringify(coord))
    const userMsg = e.message.includes('429')
      ? 'Server busy (rate limit). Saved locally—will sync when available.'
      : e.message.includes('401')
        ? 'Authentication expired. Saved locally. Please refresh.'
        : e.message.includes('403')
          ? 'Permission denied. Saved locally. Check your access or contact support.'
          : `Could not reach SharePoint. Saved locally.`

    const authInfo = debugAuth()
    const details = `Error: ${e.message}
User: ${authInfo.account || 'Unknown'}
Signed In: ${authInfo.signedIn}
Scopes: ${authInfo.scopes.join(', ')}
Time: ${new Date().toISOString()}

Troubleshooting:
• If you can see the dashboard but can't save, you may have read-only permissions
• Try refreshing the page to re-authenticate
• Contact support with this error details`
    showErrorModal('Coordinator Save Failed', userMsg, details)

    if (savingEl) savingEl.textContent = `⚠ Saved locally`
    setTimeout(() => {
      if (savingEl) savingEl.textContent = ''
    }, 3000)
  }
}

function debounceSave() {
  clearTimeout(saveTimer)
  saveTimer = setTimeout(() => {
    CONFIG.useSharePoint
      ? saveCoordToSharePoint()
      : localStorage.setItem('sitrep_coord', JSON.stringify(coord))
  }, 800)
}
