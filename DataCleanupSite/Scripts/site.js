/* ============================================================
   DataCleanup Tool — Shared JavaScript
   ============================================================ */

'use strict';

// ── Toast notification ───────────────────────────────
window.DataCleanup = window.DataCleanup || {};

DataCleanup.showToast = function (msg, type) {
    let t = document.getElementById('globalToast');
    if (!t) {
        t = document.createElement('div');
        t.id = 'globalToast';
        t.style.cssText = [
            'position:fixed', 'bottom:28px', 'right:28px',
            'background:var(--surface)', 'border:1px solid var(--accent)',
            'border-radius:8px', 'padding:13px 18px',
            'font-family:var(--mono)', 'font-size:13px', 'color:var(--accent)',
            'box-shadow:0 8px 32px rgba(0,0,0,.45)',
            'transition:all .3s', 'z-index:9999',
            'transform:translateY(20px)', 'opacity:0',
            'max-width:340px', 'line-height:1.4'
        ].join(';');
        document.body.appendChild(t);
    }
    t.textContent = msg;
    t.style.borderColor = type === 'error' ? 'var(--danger)' : 'var(--accent)';
    t.style.color       = type === 'error' ? 'var(--danger)' : 'var(--accent)';
    t.style.transform   = 'translateY(0)';
    t.style.opacity     = '1';
    clearTimeout(t._timer);
    t._timer = setTimeout(function () {
        t.style.transform = 'translateY(20px)';
        t.style.opacity   = '0';
    }, 4200);
};

// ── Highlight active nav link ────────────────────────
document.addEventListener('DOMContentLoaded', function () {
    var path = window.location.pathname.toLowerCase();
    document.querySelectorAll('.nav-link').forEach(function (a) {
        var href = (a.getAttribute('href') || '').toLowerCase().replace(/^~\//, '/');
        if (href && path.endsWith(href.replace(/^\//, ''))) {
            a.style.color = 'var(--accent)';
        }
    });
});
