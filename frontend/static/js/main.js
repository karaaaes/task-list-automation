// Sidebar toggle on mobile
document.addEventListener('DOMContentLoaded', () => {
  const tg = document.getElementById('toggleSidebar');
  const sb = document.getElementById('sidebar');
  if (tg && sb) tg.addEventListener('click', () => sb.classList.toggle('open'));
});

// Toast helper (dark mode)
function showToast(message, type = 'success') {
  const box = document.getElementById('toastBox');
  if (!box) return;
  const colors = {
    success: 'background:#1F3225; color:#7BB389; border:1px solid #7BB389;',
    error:   'background:#3A1F1B; color:#C9776B; border:1px solid #C9776B;',
    info:    'background:#1F2A38; color:#6B89B8; border:1px solid #6B89B8;',
  };
  const icon = { success: 'check-circle-fill', error: 'exclamation-circle-fill', info: 'info-circle-fill' }[type] || 'info-circle-fill';
  const el = document.createElement('div');
  el.className = 'toast show';
  el.style.cssText = `border-radius:10px; padding:12px 16px; font-weight:600; box-shadow:0 8px 24px rgba(0,0,0,.4); min-width:240px; display:flex; align-items:center; gap:10px; ${colors[type]}`;
  el.innerHTML = `<i class="bi bi-${icon}"></i><span>${message}</span>`;
  box.appendChild(el);
  setTimeout(() => { el.style.opacity = '0'; el.style.transition = 'opacity .3s'; }, 2400);
  setTimeout(() => el.remove(), 2800);
}

function debounce(fn, ms = 300) {
  let t; return function (...args) {
    clearTimeout(t); t = setTimeout(() => fn.apply(this, args), ms);
  };
}

function fmtDate(s) {
  if (!s) return '-';
  const d = new Date(s);
  if (isNaN(d)) return s;
  return d.toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric' });
}
function fmtDateTime(s) {
  if (!s) return '-';
  const d = new Date(s);
  if (isNaN(d)) return s;
  return d.toLocaleDateString('en-GB', { day:'2-digit', month:'short', year:'numeric' }) + ' ' + d.toLocaleTimeString('en-GB', { hour:'2-digit', minute:'2-digit'});
}
