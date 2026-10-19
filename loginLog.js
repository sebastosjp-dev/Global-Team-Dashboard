/**
 * loginLog.js — Records each sign-in to the Apps Script login log and shows
 * recent visitors in the sidebar under "Last Updated".
 * Backend: apps-script/login-log.gs (verifies the Google ID token server-side).
 * @module loginLog
 */
import { AUTH } from './config.js';

const MAX_VISITORS = 8;

/** POST as text/plain so the browser sends a simple request (no CORS preflight). */
async function callLog(action, idToken) {
    const res = await fetch(AUTH.LOGIN_LOG_URL, {
        method: 'POST',
        headers: { 'Content-Type': 'text/plain;charset=utf-8' },
        body: JSON.stringify({ action, idToken, userAgent: navigator.userAgent }),
    });
    const data = await res.json();
    if (!data.ok) throw new Error(data.error || 'login log request failed');
    return data;
}

/**
 * Record a fresh sign-in.
 * @param {string} idToken
 */
export async function recordLogin(idToken) {
    if (!AUTH.LOGIN_LOG_URL || !idToken) return;
    try { await callLog('log', idToken); }
    catch (e) { console.warn('[LoginLog] record failed:', e); }
}

/** "5m ago", "3h ago", "2d ago", else a date. */
function timeAgo(date) {
    const mins = Math.floor((Date.now() - date.getTime()) / 60000);
    if (mins < 1) return 'just now';
    if (mins < 60) return `${mins}m ago`;
    const hours = Math.floor(mins / 60);
    if (hours < 24) return `${hours}h ago`;
    const days = Math.floor(hours / 24);
    if (days < 7) return `${days}d ago`;
    return `${String(date.getDate()).padStart(2, '0')}.${String(date.getMonth() + 1).padStart(2, '0')}`;
}

const escapeHtml = s => String(s).replace(/[&<>"']/g, c =>
    ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));

/**
 * Fetch the log and render the latest visit per person into #recent-visitors.
 * @param {string} idToken
 */
export async function renderRecentVisitors(idToken) {
    const box = document.getElementById('recent-visitors');
    if (!box || !AUTH.LOGIN_LOG_URL || !idToken) return;
    try {
        const { entries } = await callLog('list', idToken);   // newest first

        const latest = new Map();          // email → { time, name } of their most recent visit
        const startOfToday = new Date(); startOfToday.setHours(0, 0, 0, 0);
        entries.forEach(e => {
            const time = new Date(e.time);
            if (isNaN(time)) return;
            if (!latest.has(e.email)) latest.set(e.email, { time, name: e.name });
        });
        const todayCount = [...latest.values()].filter(v => v.time >= startOfToday).length;

        const rows = [...latest.entries()].slice(0, MAX_VISITORS).map(([email, v]) => `
            <li class="visitor-row" title="${escapeHtml(v.name ? `${v.name} — ` : '')}${escapeHtml(v.time.toLocaleString())}">
                <span class="visitor-email">${escapeHtml(email)}</span>
                <span class="visitor-time">${timeAgo(v.time)}</span>
            </li>`).join('');

        box.innerHTML = `
            <span class="last-updated-label">Recent Visitors · ${todayCount} today</span>
            <ul class="visitor-list">${rows || '<li class="visitor-row"><span class="visitor-time">No visits yet</span></li>'}</ul>`;
        box.classList.remove('hidden');
    } catch (e) {
        console.warn('[LoginLog] list failed:', e);
    }
}
