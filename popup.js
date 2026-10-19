/**
 * popup.js — Shared popup window + drag support for every popup in the dashboard.
 *
 * All popups (existing *-modal overlays and the ones opened via openPopup) follow
 * the same shape: a fixed full-screen overlay whose direct child is the panel.
 * initDraggablePopups() installs one delegated handler so any such panel can be
 * moved by its header, clamped to stay inside the browser window.
 * @module popup
 */

const esc = (s) => String(s ?? '').replace(/[&<>"']/g, c =>
    ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));

const fmtUsd = (v) => '$' + Math.round(v || 0).toLocaleString('en-US');
const fmtDate = (d) => d instanceof Date && !isNaN(d)
    ? d.toLocaleDateString('en-US', { year: 'numeric', month: 'short', day: 'numeric' }) : '-';

/* ═══════════════════════════════════════════════════════════════
   Drag support
   ═══════════════════════════════════════════════════════════════ */

const OVERLAY_SELECTOR = '[id$="-modal"], .dash-popup';
const HANDLE_HEIGHT = 64;   // px from the panel top that act as the drag handle
const NO_DRAG = 'button, a, input, select, textarea, label, canvas, [contenteditable="true"]';

/** Clamp so the panel stays inside the window (or pinned to its top-left when larger). */
const clamp = (v, limit) => Math.min(Math.max(v, Math.min(0, limit)), Math.max(0, limit));

let installed = false;

export function initDraggablePopups() {
    if (installed) return;
    installed = true;

    document.addEventListener('pointerdown', (e) => {
        if (e.button !== 0 || e.target.closest(NO_DRAG)) return;
        const overlay = e.target.closest(OVERLAY_SELECTOR);
        if (!overlay || getComputedStyle(overlay).position !== 'fixed') return;
        const panel = [...overlay.children].find(c => c.contains(e.target));
        if (!panel) return;

        const rect = panel.getBoundingClientRect();
        if (e.clientY - rect.top > HANDLE_HEIGHT) return;

        const [curX, curY] = (panel.style.translate || '0px 0px').split(' ').map(v => parseFloat(v) || 0);
        const baseLeft = rect.left - curX;
        const baseTop = rect.top - curY;
        const startX = e.clientX, startY = e.clientY;
        let moved = false;

        const onMove = (ev) => {
            const dx = ev.clientX - startX, dy = ev.clientY - startY;
            if (!moved && Math.hypot(dx, dy) < 4) return;
            moved = true;
            const left = clamp(rect.left + dx, window.innerWidth - rect.width);
            const top = clamp(rect.top + dy, window.innerHeight - rect.height);
            panel.style.translate = `${left - baseLeft}px ${top - baseTop}px`;
        };
        const onUp = () => {
            document.removeEventListener('pointermove', onMove);
            document.removeEventListener('pointerup', onUp);
            document.body.classList.remove('popup-dragging');
            // A drag that ends over the backdrop must not count as a click-to-close.
            // The click (if any) is dispatched in the same task as pointerup.
            if (moved) {
                const swallow = ev => ev.stopPropagation();
                window.addEventListener('click', swallow, { capture: true, once: true });
                setTimeout(() => window.removeEventListener('click', swallow, { capture: true }), 0);
            }
        };

        document.body.classList.add('popup-dragging');
        document.addEventListener('pointermove', onMove);
        document.addEventListener('pointerup', onUp);
        e.preventDefault();
    });
}

/* ═══════════════════════════════════════════════════════════════
   Generic popup
   ═══════════════════════════════════════════════════════════════ */

/**
 * Open a popup window. Only one generic popup is open at a time.
 * @param {Object} o
 * @param {string} o.title
 * @param {string} [o.eyebrow]  small label above the title
 * @param {string} [o.icon]     Font Awesome icon name, e.g. 'fa-circle-xmark'
 * @param {string} [o.color]    accent color
 * @param {string} [o.badge]    short text shown at the top right
 * @param {string} o.bodyHtml
 * @param {number} [o.width]
 */
export function openPopup({ title, eyebrow = '', icon = 'fa-circle-info', color = '#007AFF', badge = '', bodyHtml, width = 880 }) {
    closePopup();

    const overlay = document.createElement('div');
    overlay.className = 'dash-popup';
    overlay.addEventListener('click', (e) => { if (e.target === overlay) closePopup(); });
    overlay.innerHTML = `
        <div class="dash-popup-panel" style="width:min(${width}px, 100%);">
            <div class="dash-popup-header">
                <div style="display:flex; align-items:center; gap:10px; min-width:0;">
                    <div class="dash-popup-icon" style="background:${color}1F; color:${color};"><i class="fa-solid ${icon}"></i></div>
                    <div style="min-width:0;">
                        ${eyebrow ? `<div class="dash-popup-eyebrow" style="color:${color};">${esc(eyebrow)}</div>` : ''}
                        <div class="dash-popup-title">${esc(title)}</div>
                    </div>
                </div>
                <div style="display:flex; align-items:center; gap:8px; flex-shrink:0;">
                    ${badge ? `<div class="dash-popup-badge" style="color:${color}; background:${color}14; border-color:${color}40;">${esc(badge)}</div>` : ''}
                    <button class="dash-popup-close" title="Close (Esc)"><i class="fa-solid fa-xmark"></i></button>
                </div>
            </div>
            <div class="dash-popup-body">${bodyHtml}</div>
        </div>`;
    overlay.querySelector('.dash-popup-close').addEventListener('click', closePopup);
    document.body.appendChild(overlay);
    document.addEventListener('keydown', onEsc);
    return overlay;
}

function onEsc(e) { if (e.key === 'Escape') closePopup(); }

export function closePopup() {
    document.querySelectorAll('.dash-popup').forEach(el => el.remove());
    document.removeEventListener('keydown', onEsc);
}

/* ═══════════════════════════════════════════════════════════════
   Building blocks for popup bodies
   ═══════════════════════════════════════════════════════════════ */

/** Row of summary chips: [{ label, value, color?, sub? }] */
export function statChipsHTML(chips) {
    return `<div class="dash-popup-chips">${chips.map(c => `
        <div class="dash-popup-chip">
            <div class="dash-popup-chip-label">${esc(c.label)}</div>
            <div class="dash-popup-chip-value" style="color:${c.color || '#0F172A'};">${esc(c.value)}</div>
            ${c.sub ? `<div class="dash-popup-chip-sub">${esc(c.sub)}</div>` : ''}
        </div>`).join('')}</div>`;
}

/**
 * Table of rows. columns: [{ label, key | render(row), align?, money?, date? }]
 * @param {Object[]} rows
 * @param {Object[]} columns
 * @param {string} [emptyText]
 */
export function tableHTML(rows, columns, emptyText = 'No records.') {
    if (!rows.length) return `<div class="dash-popup-empty">${esc(emptyText)}</div>`;
    const cell = (row, c) => {
        if (c.render) return c.render(row);
        const v = row[c.key];
        if (c.money) return v ? fmtUsd(v) : '-';
        if (c.date) return fmtDate(v);
        return esc(v) || '-';
    };
    return `
        <div class="dash-popup-table-wrap">
            <table class="dash-popup-table">
                <thead><tr><th style="width:36px;">#</th>${columns.map(c => `<th style="text-align:${c.align || (c.money ? 'right' : 'left')};">${esc(c.label)}</th>`).join('')}</tr></thead>
                <tbody>${rows.map((r, i) => `
                    <tr><td class="dash-popup-rownum">${i + 1}</td>${columns.map(c => `<td style="text-align:${c.align || (c.money ? 'right' : 'left')};">${cell(r, c)}</td>`).join('')}</tr>`).join('')}
                </tbody>
            </table>
        </div>`;
}

export function sectionHTML(title, inner) {
    return `<div class="dash-popup-section"><div class="dash-popup-section-title">${esc(title)}</div>${inner}</div>`;
}

export { esc as escapeHtml, fmtUsd, fmtDate };
