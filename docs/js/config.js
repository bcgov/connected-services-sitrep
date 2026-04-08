// ── CONFIG.JS ────────────────────────────────────────────────────────────────
// Global configuration and shared mutable state for the SitRep dashboard.
//
// Load order: this file must be first — all other scripts depend on CONFIG,
// DEFAULT_TEAMS, and the shared `let` variables declared here.
//
// There is no module system (plain <script> tags), so everything declared
// here lives on the global scope and is accessible by all other files.
// ─────────────────────────────────────────────────────────────────────────────

// ── APP CONFIG ────────────────────────────────────────────────────────────────
// Static settings that rarely change. To disable SharePoint integration for
// local/CSV-only testing, set useSharePoint: false.
const CONFIG = {
  useSharePoint: true,
  clientId: 'f5642799-c284-4221-92a5-06abf8795f97', // Azure AD app registration
  sharePointSite:
    'https://bcgov.sharepoint.com/teams/12320-ConnectedServicesStrategicPriority',
  listName: 'Weekly SitRep Data',      // SharePoint list for team submissions
  coordListName: 'SitRep Coordinator', // SharePoint list for coordinator notes
}

// ── DEFAULT TEAMS ─────────────────────────────────────────────────────────────
// Fallback roster shown before SharePoint loads, and safety net if the SP
// column-choices API call fails. Keep in sync with the TeamName choices in
// the SharePoint list and the MS Form question.
//
// To add a new team permanently: update the SharePoint TeamName choices column,
// update the MS Form, then add the name here so it appears before the first
// submission. See CLAUDE.md › "Pending" for the full team-management workflow.
const DEFAULT_TEAMS = [
  'BCSC',
  'EFV',
  'SDG',
  'ADR',
  'CHEFS',
  'Notify',
  'Workflow',
  'DDS',
  'SDX',
  'DevX',
  'CSTAR',
  'Disability Services',
  'SSI',
  'SDPR/SDD',
  'ACT',
]

// ── SHARED STATE ──────────────────────────────────────────────────────────────
// All mutable dashboard state lives here so every script can read/write it
// without passing arguments across the global function calls.

// Current-week team data. Keyed by team name string.
// Shape of each value — see sharepoint.js loadFromSharePoint for full details:
//   { team, status, highlight, blocker, initiativeNum, escalatorNum,
//     depsIn[], summary, _weekOf, _spId, _localOnly?, _stale? }
let data = {}

// Coordinator panel data for the current week.
// Shape: { status, news, meetings[], featuredHighlights[], featuredBlockers[] }
// Loaded from the SitRep Coordinator SP list; falls back to localStorage.
let coord = {}

// All raw SharePoint items ever fetched (used by the history view).
// Each element is a Graph API list item with .fields and .id.
let allHistoryItems = []

// Authoritative team roster derived from the SharePoint TeamName column choices
// (supplemented by any team names found in history). Populated after the first
// successful SharePoint load. ui.js › getRosterTeams() merges this with
// DEFAULT_TEAMS and Object.keys(data) for the final rendered list.
let allTeamNames = []

// ── UI STATE ──────────────────────────────────────────────────────────────────
let currentFilter = 'all'  // active filter-bar selection
let selectedRYG = null     // RYG button selection inside the edit modal
let selectedDeps = []      // dependency chip selections inside the edit modal
let coordOpen = false      // whether the coordinator panel is expanded
let historyOpen = false    // whether the history view is shown

// ── SHAREPOINT IDS ────────────────────────────────────────────────────────────
// Resolved once on first load and reused for all subsequent Graph API calls.
// null until sharepoint.js › loadFromSharePoint() succeeds.
let coordItemId = null  // SP item ID for this week's coordinator entry (PATCH vs POST)
let _siteId = null
let _coordListId = null
let _teamListId = null

// ── TIMERS ────────────────────────────────────────────────────────────────────
let saveTimer = null        // debounce handle for coordinator auto-save (800 ms)
let autoRefreshTimer = null // interval handle for smart auto-refresh (60 s)
