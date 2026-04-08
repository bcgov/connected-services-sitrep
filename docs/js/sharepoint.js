// ── SHAREPOINT.JS ────────────────────────────────────────────────────────────
// All Microsoft Graph API calls: loading and saving team data, loading and
// saving coordinator data, local-storage fallback, and page-state management.
//
// Dependencies: config.js (state), utils.js (helpers), auth.js (getToken)
//
// Key data flows:
//   Load  → loadFromSharePoint() → data{}, allHistoryItems[], allTeamNames[]
//   Save  → saveTeamToSharePoint(team, teamData)   (PATCH or POST)
//   Coord → loadCoordFromSharePoint(token)          (reads coord{})
//           saveCoordToSharePoint()                 (debounced via debounceSave)
//
// Local-storage fallback:
//   If a save fails the data is written to localStorage with _localOnly: true.
//   On the next successful load, syncLocalData() attempts to push those
//   changes back to SharePoint and clears the local copies.
// ─────────────────────────────────────────────────────────────────────────────

// ── PAGE STATE ────────────────────────────────────────────────────────────────
// Controls which "screen" is visible:
//   'loading' — spinner only, all chrome hidden
//   'signin'  — full-page sign-in card, all chrome hidden
//   'loaded'  — full dashboard visible
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

// Convenience alias used before the full load completes.
function setGridLoading() {
  showFullPageState('loading')
}

// ── LOCAL STORAGE FALLBACK ────────────────────────────────────────────────────
// On load, merge any localStorage-only team entries into `data` so the UI
// reflects unsaved changes even before SharePoint responds.
// If SP IDs are already resolved, immediately try to sync those entries back.
async function loadLocalDataAndSync() {
  // Restore locally-saved team entries
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
        console.warn(`[LOCAL] Failed to parse local data for ${team}:`, e)
      }
    })

  // Restore locally-saved coordinator data
  const coordData = localStorage.getItem('sitrep_coord')
  if (coordData) {
    try {
      coord = JSON.parse(coordData)
      console.log('[LOCAL] Loaded coordinator data from localStorage')
    } catch (e) {
      console.warn('[LOCAL] Failed to parse local coordinator data:', e)
    }
  }

  // If SP is reachable, attempt to push any local-only entries back
  if (CONFIG.useSharePoint && _siteId && _teamListId) {
    await syncLocalData()
  }
}

// Push any team entries flagged _localOnly back to SharePoint.
// On success, removes the local flag and the localStorage key.
// On failure, leaves the local entry intact for the next retry.
async function syncLocalData() {
  const teamsToSync = Object.keys(data).filter((team) => data[team]._localOnly)

  for (const team of teamsToSync) {
    try {
      console.log(`[SYNC] Attempting to sync ${team}...`)
      await saveTeamToSharePoint(team, data[team])
      delete data[team]._localOnly
      localStorage.removeItem('sitrep_team_' + team)
      console.log(`[SYNC] Successfully synced ${team}`)
    } catch (e) {
      console.warn(`[SYNC] Failed to sync ${team}, keeping local:`, e.message)
    }
  }
}

// ── LOAD TEAM DATA ────────────────────────────────────────────────────────────
// Main entry point. Authenticates, resolves SP list IDs, fetches all items,
// builds the team roster from column choices + history, filters to this week,
// and renders the dashboard.
//
// Week-matching logic (an item belongs to the current week if either):
//   (a) Its WeekOf field parses to the same Monday date as getWeekLabel(), OR
//   (b) It has no WeekOf field and was created on or after this Monday midnight.
// Condition (a) covers dashboard edits; condition (b) covers MS Form submissions
// which Power Automate may leave without a WeekOf value.
async function loadFromSharePoint() {
  setGridLoading()
  try {
    const initResult = await initMsal()
    if (!initResult) {
      // MSAL failed to load — show whatever local data we have
      await loadLocalDataAndSync()
      renderAll()
      return
    }

    // Not signed in → show the sign-in landing page
    const accounts = msalInstance.getAllAccounts()
    if (accounts.length === 0) {
      showFullPageState('signin')
      return
    }

    let token = typeof initResult === 'string' ? initResult : await getToken()
    if (!token) return

    // ── Resolve site ID ──────────────────────────────────────────────────────
    const siteResp = await fetch(
      `https://graph.microsoft.com/v1.0/sites/bcgov.sharepoint.com:/teams/12320-ConnectedServicesStrategicPriority`,
      { headers: { Authorization: `Bearer ${token}` } },
    )
    if (!siteResp.ok) throw new Error(`Cannot access SharePoint site (${siteResp.status})`)
    const site = await siteResp.json()
    _siteId = site.id

    // ── Resolve team-data list ID ─────────────────────────────────────────────
    const listsResp = await fetch(
      `https://graph.microsoft.com/v1.0/sites/${_siteId}/lists?$filter=displayName eq '${CONFIG.listName}'`,
      { headers: { Authorization: `Bearer ${token}` } },
    )
    if (!listsResp.ok) throw new Error(`Cannot query SP lists (${listsResp.status})`)
    const lists = await listsResp.json()
    if (!lists.value?.length)
      throw new Error(`List "${CONFIG.listName}" not found — check CONFIG.listName`)
    _teamListId = lists.value[0].id

    // ── Build team roster from TeamName column choices ────────────────────────
    // This is the authoritative source: teams appear on the dashboard as soon
    // as they are added to the SP column, even before their first submission.
    const colsResp = await fetch(
      `https://graph.microsoft.com/v1.0/sites/${_siteId}/lists/${_teamListId}/columns`,
      { headers: { Authorization: `Bearer ${token}` } },
    )
    if (colsResp.ok) {
      const cols = await colsResp.json()
      const teamCol = (cols.value || []).find((c) => c.name === 'TeamName')
      const choices = teamCol?.choice?.choices || []
      if (choices.length) {
        allTeamNames = [...choices].sort((a, b) =>
          a.localeCompare(b, undefined, { sensitivity: 'base' }),
        )
        console.log('[SP-LOAD] Team roster from column choices:', allTeamNames)
      }
    } else {
      console.warn('[SP-LOAD] Could not fetch column choices — falling back to history')
    }

    // ── Fetch all list items ─────────────────────────────────────────────────
    // $top=500 covers ~2 years of weekly submissions for 15 teams.
    // If the list ever exceeds 500 items, add @odata.nextLink pagination here.
    const itemsResp = await fetch(
      `https://graph.microsoft.com/v1.0/sites/${_siteId}/lists/${_teamListId}/items?$expand=fields&$top=500`,
      { headers: { Authorization: `Bearer ${token}` } },
    )
    if (!itemsResp.ok) throw new Error(`Cannot fetch list items (${itemsResp.status})`)
    const items = await itemsResp.json()
    allHistoryItems = items.value || []

    // Supplement roster with any team names found in history but missing from
    // column choices (e.g. teams that submitted before being formally added).
    const teamNamesSet = new Set(allTeamNames)
    allHistoryItems.forEach((item) => {
      const teamName = item.fields?.TeamName
      if (teamName) teamNamesSet.add(teamName)
    })
    allTeamNames = Array.from(teamNamesSet).sort((a, b) =>
      a.localeCompare(b, undefined, { sensitivity: 'base' }),
    )
    console.log('[SP-LOAD] Final team roster:', allTeamNames)

    // ── Filter items to the current week ─────────────────────────────────────
    const currentWeek = getWeekLabel()
    const weekStart = getWeekStart()
    const byTeam = {}

    allHistoryItems.forEach((item) => {
      const f = item.fields,
        team = f.TeamName
      if (!team) return

      const created = new Date(f.Created || item.createdDateTime)
      const weekOfDate = parseWeekOf(f.WeekOf)
      const currentWeekDate = parseWeekOf(currentWeek)

      // See function-level comment above for the two-condition logic
      const belongsToThisWeek =
        (weekOfDate &&
          currentWeekDate &&
          weekOfDate.getTime() === currentWeekDate.getTime()) ||
        created >= weekStart

      if (!belongsToThisWeek) return

      // Keep the most recent entry per team if multiple exist for this week
      if (!byTeam[team]) {
        byTeam[team] = { fields: f, created, id: item.id }
      } else if (created > byTeam[team].created) {
        byTeam[team] = { fields: f, created, id: item.id }
      }
    })

    // Build the `data` object from this week's entries
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
        _weekOf: f.WeekOf || currentWeek,
        _spId: id,
      }
    })

    renderAll()
    showToast(`✓ Loaded ${Object.keys(data).length} teams from SharePoint`)
    await loadCoordFromSharePoint(token)
    showFullPageState('loaded')
  } catch (err) {
    console.error('[SP-LOAD] Failed to load from SharePoint:', err.message, err)
    // Don't leave the user on a blank loading screen — show sign-in which at
    // least lets them retry. A more specific error message could be shown here
    // by inspecting err.message for status codes.
    showFullPageState('signin')
  }
}

// ── SAVE TEAM DATA ────────────────────────────────────────────────────────────
// PATCH an existing item (teamData._spId set) or POST a new one.
// On failure, falls back to localStorage, shows an error modal, and RE-THROWS
// so the calling saveTeam() in ui.js knows not to close the modal or show a
// false success toast.
//
// @param {string} team     - Team name (used as the SP TeamName field value)
// @param {object} teamData - Team data object (see shape in config.js)
async function saveTeamToSharePoint(team, teamData) {
  const statusEl = document.getElementById('modal-save-status')
  if (statusEl) statusEl.textContent = 'Saving...'

  console.log('[SP-SAVE] Starting save for team:', team)

  if (!CONFIG.useSharePoint) {
    throw new Error('SharePoint disabled in CONFIG')
  }
  if (!_siteId || !_teamListId) {
    console.error('[SP-SAVE] SharePoint IDs not initialised. Reload required.')
    throw new Error(
      'SharePoint not initialised — try refreshing the page.',
    )
  }

  // Declare token outside try so it is accessible in the catch block for
  // the error report.
  let token = null
  try {
    token = await getToken()
    const fields = {
      TeamName: team,
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
      // Update existing item
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
      // Create new item
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
        throw new Error('No item ID returned from SharePoint after POST')
      }
    }

    console.log('[SP-SAVE] Success for team:', team)
    if (statusEl) statusEl.textContent = '✓ Saved to SharePoint'
    setTimeout(() => {
      if (statusEl) statusEl.textContent = ''
    }, 2500)
  } catch (e) {
    console.error('[SP-SAVE] Error saving team:', team, e.message, e)

    // Translate HTTP status codes to user-friendly messages
    const userMsg = e.message.includes('429')
      ? 'Too many requests (server busy). Try again in a moment.'
      : e.message.includes('401')
        ? 'Authentication expired. Please refresh the page.'
        : e.message.includes('403')
          ? 'Permission denied. Check your access or contact support.'
          : `Save failed: ${e.message}`

    // Persist locally so data is not lost
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
      _localOnly: true,
      _savedAt: new Date().toISOString(),
    }
    data[team] = localData
    localStorage.setItem('sitrep_team_' + team, JSON.stringify(localData))
    renderGrid() // reflect local save immediately

    const authInfo = debugAuth()
    const details = `Error: ${e.message}
Token: ${token ? 'Present' : 'Missing'}
User: ${authInfo.account || 'Unknown'}
Signed In: ${authInfo.signedIn}
Scopes: ${authInfo.scopes.join(', ')}
Team: ${team}
Time: ${new Date().toISOString()}

Troubleshooting:
• If you can see the dashboard but cannot save, you may have read-only permissions
• Try refreshing the page to re-authenticate
• Contact support with this error detail`

    showErrorModal(
      'Team Save Failed',
      `${userMsg} Data saved locally and will sync when connection is restored.`,
      details,
    )

    if (statusEl) statusEl.textContent = '⚠ Saved locally'
    setTimeout(() => {
      if (statusEl) statusEl.textContent = ''
    }, 3000)

    // Re-throw so ui.js › saveTeam() knows the SP save failed and does not
    // show a false success toast or close the modal.
    throw e
  }
}

// ── COORDINATOR DATA ──────────────────────────────────────────────────────────
// Load this week's coordinator entry from the SitRep Coordinator SP list.
// Populates coord{} and coordItemId, then re-renders the summary banner.
// Called at the end of every loadFromSharePoint().
//
// @param {string} [token] - Optional pre-fetched token; acquires one if absent.
async function loadCoordFromSharePoint(token) {
  try {
    if (!token) token = await getToken()
    if (!token || !_siteId) return

    // Resolve the coordinator list ID (cached in _coordListId after first load)
    const listsResp = await fetch(
      `https://graph.microsoft.com/v1.0/sites/${_siteId}/lists?$filter=displayName eq '${CONFIG.coordListName}'`,
      { headers: { Authorization: `Bearer ${token}` } },
    )
    if (!listsResp.ok) {
      console.warn('[COORD-LOAD] Cannot query SP lists:', listsResp.status)
      return
    }
    const lists = await listsResp.json()
    if (!lists.value?.length) {
      console.warn('[COORD-LOAD] Coordinator list not found:', CONFIG.coordListName)
      return
    }
    _coordListId = lists.value[0].id

    const weekLabel = getWeekLabel()
    const itemsResp = await fetch(
      `https://graph.microsoft.com/v1.0/sites/${_siteId}/lists/${_coordListId}/items?$expand=fields&$top=50`,
      { headers: { Authorization: `Bearer ${token}` } },
    )
    if (!itemsResp.ok) {
      console.warn('[COORD-LOAD] Cannot fetch coordinator items:', itemsResp.status)
      return
    }
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

    // If the coordinator panel is open, refresh its fields live
    if (coordOpen) {
      document.getElementById('coord-status').value = coord.status || ''
      document.getElementById('coord-news').value = coord.news || ''
      renderFeaturedChips()
      renderMeetingsList()
    }
  } catch (e) {
    console.warn('[COORD-LOAD] Could not load coordinator data:', e.message, e)
  }
}

// Save the current coord{} object to the SitRep Coordinator SP list.
// PATCH if coordItemId exists (this week's entry already created), POST if not.
// Falls back to localStorage on any error and shows an error modal.
async function saveCoordToSharePoint() {
  const savingEl = document.getElementById('coord-saving')
  if (savingEl) savingEl.textContent = 'Saving...'

  try {
    const token = await getToken()

    // If SP infrastructure is not ready (e.g. called before load completes),
    // save locally and bail out quietly.
    if (!token || !_siteId || !_coordListId) {
      console.warn('[COORD-SAVE] SP not ready — saving locally')
      localStorage.setItem('sitrep_coord', JSON.stringify(coord))
      if (savingEl) savingEl.textContent = '⚠ Saved locally (SP not ready)'
      setTimeout(() => {
        if (savingEl) savingEl.textContent = ''
      }, 3000)
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
        throw new Error('No item ID returned from SharePoint after POST')
      }
    }

    if (savingEl) savingEl.textContent = '✓ Saved'
    setTimeout(() => {
      if (savingEl) savingEl.textContent = ''
    }, 2000)
  } catch (e) {
    console.error('[COORD-SAVE] Error:', e.message, e)

    // Always save locally so coordinator data is not lost
    localStorage.setItem('sitrep_coord', JSON.stringify(coord))

    const userMsg = e.message.includes('429')
      ? 'Server busy (rate limit). Saved locally — will sync when available.'
      : e.message.includes('401')
        ? 'Authentication expired. Saved locally. Please refresh.'
        : e.message.includes('403')
          ? 'Permission denied. Saved locally. Check your access or contact support.'
          : 'Could not reach SharePoint. Saved locally.'

    const authInfo = debugAuth()
    const details = `Error: ${e.message}
User: ${authInfo.account || 'Unknown'}
Signed In: ${authInfo.signedIn}
Scopes: ${authInfo.scopes.join(', ')}
Time: ${new Date().toISOString()}

Troubleshooting:
• If you can see the dashboard but cannot save, you may have read-only permissions
• Try refreshing the page to re-authenticate
• Contact support with this error detail`

    showErrorModal('Coordinator Save Failed', userMsg, details)

    if (savingEl) savingEl.textContent = '⚠ Saved locally'
    setTimeout(() => {
      if (savingEl) savingEl.textContent = ''
    }, 3000)
  }
}

// Debounce coordinator saves to avoid hammering the Graph API on every
// keystroke. Fires 800 ms after the last change.
function debounceSave() {
  clearTimeout(saveTimer)
  saveTimer = setTimeout(() => {
    CONFIG.useSharePoint
      ? saveCoordToSharePoint()
      : localStorage.setItem('sitrep_coord', JSON.stringify(coord))
  }, 800)
}
