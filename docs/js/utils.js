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
