// ── AUTH.JS ──────────────────────────────────────────────────────────────────
// MSAL (Microsoft Authentication Library) setup and token management.
//
// Auth flow:
//   1. MSAL browser library loads via <script> tag in sitrepdash.html.
//   2. The script's `onload` fires init() (in ui.js).
//   3. init() calls loadFromSharePoint() (in sharepoint.js).
//   4. loadFromSharePoint() calls initMsal() here to handle any in-progress
//      redirect and check for existing sessions.
//   5. If no session exists, showFullPageState('signin') shows the sign-in
//      screen. The user clicks "Sign in" → signIn() → loginPopup().
//   6. After sign-in, getToken() is called before every Graph API request.
//      It tries acquireTokenSilent first, falling back to loginPopup.
//
// Azure AD app: CSBC-CITZ-SitRep (client ID in CONFIG)
// Tenant: BC Gov (tenant ID hardcoded in msalConfig below)
// Scopes: Sites.Read.All, Sites.ReadWrite.All (delegated)
// ─────────────────────────────────────────────────────────────────────────────

const msalConfig = {
  auth: {
    clientId: CONFIG.clientId,
    authority:
      'https://login.microsoftonline.com/6fdb5200-3d0d-4a8a-b036-d3685e359adc',
    redirectUri: window.location.origin,
  },
  cache: { cacheLocation: 'sessionStorage' },
}

// Shared MSAL instance — initialised once by initMsal(), reused by getToken().
let msalInstance = null

// Initialise the MSAL PublicClientApplication and handle any in-progress
// redirect flow. Must be awaited before any other MSAL call.
//
// Returns:
//   false         — MSAL library not loaded (script tag failed)
//   true          — initialised, no active redirect
//   string (JWT)  — redirect completed; value is the access token
async function initMsal() {
  const { PublicClientApplication } = window.msal || {}
  if (!PublicClientApplication) {
    showToast('MSAL library not loaded — check network connectivity')
    return false
  }
  msalInstance = new PublicClientApplication(msalConfig)
  await msalInstance.initialize()
  try {
    const r = await msalInstance.handleRedirectPromise()
    if (r) return r.accessToken
  } catch (e) {
    // Redirect errors (e.g. cancelled login, misconfigured redirect URI) are
    // logged but not fatal — fall through to the normal silent-token path.
    console.error('[AUTH] handleRedirectPromise error:', e.message, e)
  }
  return true
}

// Acquire a Graph API access token for the current user.
// Tries the MSAL cache/refresh first (silent), then falls back to a popup.
//
// Throws if the popup is blocked or the user cancels — callers should catch
// and surface a user-friendly message.
async function getToken() {
  const scopes = [
    'https://graph.microsoft.com/Sites.Read.All',
    'https://graph.microsoft.com/Sites.ReadWrite.All',
  ]
  const accounts = msalInstance.getAllAccounts()
  if (accounts.length > 0) {
    try {
      const r = await msalInstance.acquireTokenSilent({
        scopes,
        account: accounts[0],
      })
      console.log('[AUTH] Token acquired silently for:', accounts[0].username)
      return r.accessToken
    } catch (e) {
      // Silent acquisition can fail when the refresh token expires or consent
      // is revoked. Fall through to popup.
      console.warn('[AUTH] Silent token acquisition failed:', e.message)
    }
  }
  try {
    console.log('[AUTH] Acquiring token via popup...')
    const r = await msalInstance.loginPopup({ scopes })
    console.log('[AUTH] Token acquired via popup for:', r.account.username)
    return r.accessToken
  } catch (e) {
    throw new Error(
      'Sign in failed — please allow popups for this site and try again',
    )
  }
}

// Return a snapshot of the current auth state for inclusion in error reports.
// Called by saveTeamToSharePoint and saveCoordToSharePoint when building the
// "Technical details" block shown in error modals.
function debugAuth() {
  const accounts = msalInstance?.getAllAccounts() || []
  console.log('[AUTH DEBUG]', {
    accounts: accounts.map((a) => ({
      username: a.username,
      localAccountId: a.localAccountId,
    })),
    hasMsal: !!window.msal,
    msalInstance: !!msalInstance,
    scopes: ['Sites.Read.All', 'Sites.ReadWrite.All'],
  })
  return {
    signedIn: accounts.length > 0,
    account: accounts[0]?.username,
    scopes: ['Sites.Read.All', 'Sites.ReadWrite.All'],
  }
}

// Triggered by the "Sign in with BC Gov account" button on the landing screen.
// Re-uses the existing msalInstance if already initialised.
async function signIn() {
  try {
    const initResult = await initMsal()
    if (!initResult) {
      showToast('MSAL not loaded — cannot sign in')
      return
    }
    const scopes = [
      'https://graph.microsoft.com/Sites.Read.All',
      'https://graph.microsoft.com/Sites.ReadWrite.All',
    ]
    await msalInstance.loginPopup({ scopes })
    loadFromSharePoint()
  } catch (e) {
    showToast('Sign in failed — please allow popups and try again')
  }
}
