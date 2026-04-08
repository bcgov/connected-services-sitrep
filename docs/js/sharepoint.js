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

// ── LOAD TEAM DATA ────────────────────────────────────────────────────────────
async function loadFromSharePoint() {
  setGridLoading()
  try {
    const initResult = await initMsal()
    if (!initResult) {
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
  if (!CONFIG.useSharePoint || !_siteId || !_teamListId) return
  const statusEl = document.getElementById('modal-save-status')
  if (statusEl) statusEl.textContent = 'Saving...'
  try {
    const token = await getToken()
    const fields = {
      TeamName: team,
      WeekOf: teamData._weekOf || getWeekLabel(),
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
    if (statusEl) statusEl.textContent = '✓ Saved to SharePoint'
    setTimeout(() => {
      if (statusEl) statusEl.textContent = ''
    }, 2500)
  } catch (e) {
    console.error('ERROR: Could not save team to SharePoint:', e.message, e)
    const userMsg = e.message.includes('429')
      ? 'Save failed: Too many requests (server busy). Try again in a moment.'
      : e.message.includes('401')
        ? 'Save failed: Authentication expired. Please refresh the page.'
        : e.message.includes('403')
          ? 'Save failed: Permission denied. Check your access.'
          : `Save failed: ${e.message}`
    if (statusEl) statusEl.textContent = `⚠ ${userMsg}`
    showToast(`⚠ Team save failed: ${userMsg}`)
    setTimeout(() => {
      if (statusEl) statusEl.textContent = ''
    }, 5000)
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
          ? 'Permission denied. Saved locally. Check your access.'
          : `Could not reach SharePoint. Saved locally.`
    if (savingEl) savingEl.textContent = `⚠ ${userMsg}`
    showToast(`⚠ ${userMsg}`)
    setTimeout(() => {
      if (savingEl) savingEl.textContent = ''
    }, 5000)
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
