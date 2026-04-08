// ── AUTH (MSAL) ──────────────────────────────────────────────────────────────

const msalConfig = {
  auth: {
    clientId: CONFIG.clientId,
    authority: 'https://login.microsoftonline.com/6fdb5200-3d0d-4a8a-b036-d3685e359adc',
    redirectUri: window.location.origin,
  },
  cache: { cacheLocation: 'sessionStorage' },
}

let msalInstance = null

async function initMsal() {
  const { PublicClientApplication } = window.msal || {}
  if (!PublicClientApplication) {
    showToast('MSAL library not loaded')
    return false
  }
  msalInstance = new PublicClientApplication(msalConfig)
  await msalInstance.initialize()
  try {
    const r = await msalInstance.handleRedirectPromise()
    if (r) return r.accessToken
  } catch (e) {}
  return true
}

async function getToken() {
  const scopes = [
    'https://graph.microsoft.com/Sites.Read.All',
    'https://graph.microsoft.com/Sites.ReadWrite.All',
  ]
  const accounts = msalInstance.getAllAccounts()
  if (accounts.length > 0) {
    try {
      const r = await msalInstance.acquireTokenSilent({ scopes, account: accounts[0] })
      console.log('[AUTH] Token acquired silently for:', accounts[0].username)
      return r.accessToken
    } catch (e) {
      console.warn('[AUTH] Silent token acquisition failed:', e.message)
    }
  }
  try {
    console.log('[AUTH] Acquiring token via popup...')
    const r = await msalInstance.loginPopup({ scopes })
    console.log('[AUTH] Token acquired via popup for:', r.account.username)
    return r.accessToken
  } catch (e) {
    throw new Error('Sign in failed — please allow popups for this site and try again')
  }
}

// Debug function to check current authentication state
function debugAuth() {
  const accounts = msalInstance?.getAllAccounts() || []
  console.log('[AUTH DEBUG]', {
    accounts: accounts.map(a => ({ username: a.username, localAccountId: a.localAccountId })),
    hasMsal: !!window.msal,
    msalInstance: !!msalInstance,
    scopes: ['Sites.Read.All', 'Sites.ReadWrite.All']
  })
  return {
    signedIn: accounts.length > 0,
    account: accounts[0]?.username,
    scopes: ['Sites.Read.All', 'Sites.ReadWrite.All']
  }
}

// Called by the sign in button on the unauthenticated landing screen
async function signIn() {
  try {
    const initResult = await initMsal()
    if (!initResult) { showToast('MSAL not loaded'); return }
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
