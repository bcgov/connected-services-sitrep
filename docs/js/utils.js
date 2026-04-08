// ── UTILS ────────────────────────────────────────────────────────────────────

function esc(str) {
  return String(str || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
}

function safeJSON(str, fallback) {
  if (!str) return fallback
  try { return JSON.parse(str) } catch { return fallback }
}

function getWeekLabel() {
  const d = new Date(), day = d.getDay()
  const diff = d.getDate() - day + (day === 0 ? -6 : 1)
  d.setDate(diff)
  return d.toLocaleDateString('en-CA', { day: '2-digit', month: '2-digit', year: 'numeric' })
}

function getWeekStart() {
  const d = new Date(), day = d.getDay()
  const diff = d.getDate() - day + (day === 0 ? -6 : 1)
  d.setDate(diff)
  d.setHours(0, 0, 0, 0)
  return d
}

function parseDepsIn(raw) {
  if (!raw) return []
  try {
    const p = JSON.parse(raw.replace(/""/g, '"'))
    return Array.isArray(p) ? p : [p]
  } catch {
    return raw.replace(/[\[\]"]/g, '').split(/[;,]/).map(s => s.trim()).filter(Boolean)
  }
}

function getWeekOptions() {
  // Returns [lastWeek, thisWeek, nextWeek] as en-CA date strings
  return [-1, 0, 1].map(offset => {
    const d = new Date(), day = d.getDay()
    const diff = d.getDate() - day + (day === 0 ? -6 : 1) + (offset * 7)
    d.setDate(diff)
    return d.toLocaleDateString('en-CA', { day: '2-digit', month: '2-digit', year: 'numeric' })
  })
}

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

function showErrorModal(title, message, details = '') {
  const modal = document.createElement('div')
  modal.id = 'error-modal'
  modal.style.cssText = `
    position: fixed; top: 0; left: 0; width: 100%; height: 100%;
    background: rgba(0,0,0,0.5); z-index: 1000; display: flex;
    align-items: center; justify-content: center; font-family: 'BCSans', sans-serif;
  `
  modal.innerHTML = `
    <div style="background: white; padding: 24px; border-radius: 8px; max-width: 500px; width: 90%; box-shadow: 0 4px 20px rgba(0,0,0,0.3);">
      <h3 style="margin: 0 0 16px 0; color: #d32f2f; font-size: 18px; font-weight: 700;">${title}</h3>
      <p style="margin: 0 0 16px 0; color: #333; line-height: 1.5;">${message}</p>
      ${details ? `<details style="margin-bottom: 20px;"><summary style="cursor: pointer; color: #666;">Technical Details</summary><pre style="background: #f5f5f5; padding: 8px; border-radius: 4px; font-size: 12px; margin-top: 8px; overflow-x: auto;">${details}</pre></details>` : ''}
      <button onclick="this.closest('#error-modal').remove()" style="background: #d32f2f; color: white; border: none; padding: 10px 20px; border-radius: 4px; cursor: pointer; font-weight: 600;">Dismiss</button>
    </div>
  `
  document.body.appendChild(modal)
  modal.focus()
}
