// ── UTILS.JS ─────────────────────────────────────────────────────────────────
// Pure utility functions with no side-effects and no dependencies on other
// dashboard modules. Safe to call from any file.
//
// Exports (globals): esc, safeJSON, getWeekLabel, getWeekStart, getWeekOptions,
//                    parseWeekOf, parseDepsIn, showToast, showErrorModal
// ─────────────────────────────────────────────────────────────────────────────

// Escape a value for safe insertion into HTML attribute or text content.
// Always call this on user-supplied or SharePoint-sourced strings before
// injecting into innerHTML to prevent XSS.
function esc(str) {
  return String(str || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
}

// Safely parse a JSON string, returning `fallback` on any error.
// Used for SP fields that store arrays as JSON strings (Meetings, DependenciesIn, etc.).
function safeJSON(str, fallback) {
  if (!str) return fallback
  try {
    return JSON.parse(str)
  } catch {
    return fallback
  }
}

// Return the ISO-style week label (en-CA locale: YYYY-MM-DD) for the Monday
// of the current week. This is the canonical "week key" used throughout the
// dashboard to match SP items to the current week.
//
// Example: called on Wednesday 2026-04-08 → returns '2026-04-06'
function getWeekLabel() {
  const d = new Date(),
    day = d.getDay()
  const diff = d.getDate() - day + (day === 0 ? -6 : 1) // roll back to Monday
  d.setDate(diff)
  return d.toLocaleDateString('en-CA', {
    day: '2-digit',
    month: '2-digit',
    year: 'numeric',
  })
}

// Return a Date set to local midnight at the start of the current week (Monday).
// Used to test whether an item was created this week when its WeekOf field is absent.
function getWeekStart() {
  const d = new Date(),
    day = d.getDay()
  const diff = d.getDate() - day + (day === 0 ? -6 : 1)
  d.setDate(diff)
  d.setHours(0, 0, 0, 0)
  return d
}

// Return [lastWeekLabel, thisWeekLabel, nextWeekLabel] as en-CA date strings.
// Used to populate the week-picker dropdown inside the edit modal.
function getWeekOptions() {
  return [-1, 0, 1].map((offset) => {
    const d = new Date(),
      day = d.getDay()
    const diff = d.getDate() - day + (day === 0 ? -6 : 1) + offset * 7
    d.setDate(diff)
    return d.toLocaleDateString('en-CA', {
      day: '2-digit',
      month: '2-digit',
      year: 'numeric',
    })
  })
}

// Parse a week-label string into a local-midnight Date object, or null if the
// string cannot be interpreted.
//
// Handles two formats produced by this dashboard and by Power Automate:
//   YYYY-MM-DD  (en-CA locale, used by getWeekLabel and dashboard edits)
//   DD-MM-YYYY  (legacy / Power Automate output in some tenants)
//
// Note: both parts are split on '-', so ISO datetime strings (e.g. from SP)
// should be trimmed to the date portion before being passed in.
function parseWeekOf(weekOfStr) {
  if (!weekOfStr) return null
  const parts = weekOfStr.split('-')
  if (parts.length !== 3) return null
  if (parts[0].length === 4) {
    // YYYY-MM-DD
    return new Date(parseInt(parts[0]), parseInt(parts[1]) - 1, parseInt(parts[2]))
  }
  // DD-MM-YYYY
  return new Date(parseInt(parts[2]), parseInt(parts[1]) - 1, parseInt(parts[0]))
}

// Parse the DependenciesIn field from a SharePoint item into a clean string[].
// The field is stored as a JSON array string ("[\"BCSC\",\"Notify\"]") but may
// also arrive as a raw comma/semicolon-separated string from the MS Form.
function parseDepsIn(raw) {
  if (!raw) return []
  try {
    const p = JSON.parse(raw.replace(/""/g, '"'))
    return Array.isArray(p) ? p : [p]
  } catch {
    return raw
      .replace(/[\[\]"]/g, '')
      .split(/[;,]/)
      .map((s) => s.trim())
      .filter(Boolean)
  }
}

// Show a brief auto-dismissing toast notification (2.5 s).
// Also updates the ARIA live region so screen readers announce the message.
function showToast(msg) {
  const t = document.getElementById('toast')
  t.textContent = msg
  t.classList.add('show')
  document.getElementById('live-region').textContent = msg
  setTimeout(() => {
    t.classList.remove('show')
    document.getElementById('live-region').textContent = ''
  }, 2500)
}

// Show a persistent error modal that the user must manually dismiss.
// Prefer this over showToast for save failures and auth errors so users
// have time to read and screenshot the error for support.
//
// @param {string} title   - Short error heading (plain text).
// @param {string} message - Human-readable explanation. May contain HTML
//                           (e.g. links, lists) — it is injected via innerHTML
//                           so must only contain controlled/trusted markup.
// @param {string} details - Optional technical detail string shown in a
//                           collapsible <details> element (plain text).
function showErrorModal(title, message, details = '') {
  // Remove any existing error modal before adding a new one
  document.getElementById('error-modal')?.remove()

  const modal = document.createElement('div')
  modal.id = 'error-modal'
  modal.setAttribute('role', 'alertdialog')
  modal.setAttribute('aria-modal', 'true')
  modal.setAttribute('aria-labelledby', 'error-modal-title')
  modal.style.cssText = `
    position: fixed; top: 0; left: 0; width: 100%; height: 100%;
    background: rgba(0,0,0,0.5); z-index: 1000; display: flex;
    align-items: center; justify-content: center; font-family: 'BCSans', sans-serif;
  `
  modal.innerHTML = `
    <div style="background: white; padding: 24px; border-radius: 8px; max-width: 500px; width: 90%; box-shadow: 0 4px 20px rgba(0,0,0,0.3);">
      <h3 id="error-modal-title" style="margin: 0 0 16px 0; color: #d32f2f; font-size: 18px; font-weight: 700;">${esc(title)}</h3>
      <div style="margin: 0 0 16px 0; color: #333; line-height: 1.5;">${message}</div>
      ${details ? `<details style="margin-bottom: 20px;"><summary style="cursor: pointer; color: #666;">Technical details</summary><pre style="background: #f5f5f5; padding: 8px; border-radius: 4px; font-size: 12px; margin-top: 8px; overflow-x: auto; white-space: pre-wrap;">${esc(details)}</pre></details>` : ''}
      <button onclick="this.closest('#error-modal').remove()" style="background: #d32f2f; color: white; border: none; padding: 10px 20px; border-radius: 4px; cursor: pointer; font-weight: 600; font-family: 'BCSans', sans-serif;">Dismiss</button>
    </div>
  `
  document.body.appendChild(modal)
  // Move focus into the modal so keyboard users are aware of it
  modal.querySelector('button').focus()
}
