// ── CONFIG ──────────────────────────────────────────────────────────────────
const CONFIG = {
  useSharePoint: true,
  clientId: 'f5642799-c284-4221-92a5-06abf8795f97',
  sharePointSite:
    'https://bcgov.sharepoint.com/teams/12320-ConnectedServicesStrategicPriority',
  listName: 'Weekly SitRep Data',
  coordListName: 'SitRep Coordinator',
}

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

// ── SHARED STATE ─────────────────────────────────────────────────────────────
let data = {} // current week team data keyed by team name
let coord = {} // coordinator data (news, meetings, featured etc)
let allHistoryItems = [] // raw SharePoint items for history view

let currentFilter = 'all'
let selectedRYG = null
let selectedDeps = []
let coordOpen = false
let historyOpen = false

// SharePoint IDs cached after first load
let coordItemId = null
let _siteId = null
let _coordListId = null
let _teamListId = null

let saveTimer = null
let autoRefreshTimer = null // debounce timer for coordinator saves
