/**
 * weeklyReport.js — Weekly Report generation, PDF download, and notes management.
 * Auto-finalizes on Friday 18:00 KST. PDF via html2pdf.js.
 */
import { parseCurrency, normalizeCountry } from './utils.js';
import {
    findCountryKey, findKorTcvKey, findArrKey,
    findContractStartKey, findStatusKey, findDealNameKey,
    findPocStartKey, findEstimatedValueKey,
    findKey, parseExcelDateSafe
} from './columnFinder.js';

/* ═══════════════════════════════════════════════════════════════
   Country Config — Teams & Flags
   ═══════════════════════════════════════════════════════════════ */

const COUNTRY_TEAMS = {
    'Indonesia': [
        { role: 'Country Director' },
        { role: 'Channel Manager' },
        { role: 'Sales Admin' },
        { role: 'Solution Engineer' },
    ],
    'Thailand': [
        { role: 'Solution Engineer' },
    ],
};

const COUNTRY_FLAGS = {
    'Indonesia':    '🇮🇩',
    'Thailand':     '🇹🇭',
    'Singapore':    '🇸🇬',
    'Malaysia':     '🇲🇾',
    'Vietnam':      '🇻🇳',
    'Philippines':  '🇵🇭',
    'Japan':        '🇯🇵',
    'India':        '🇮🇳',
    'Australia':    '🇦🇺',
    'USA':          '🇺🇸',
    'Taiwan':       '🇹🇼',
    'Hong Kong':    '🇭🇰',
};

/* ═══════════════════════════════════════════════════════════════
   Date Utilities (KST-aware)
   ═══════════════════════════════════════════════════════════════ */

function getKSTNow() {
    return new Date(Date.now() + 9 * 3600 * 1000);
}

function getWeekBounds() {
    const kst = getKSTNow();
    const day = kst.getUTCDay(); // 0=Sun,1=Mon,...,5=Fri,6=Sat
    const diffMon = day === 0 ? -6 : 1 - day;

    const monKST = new Date(kst);
    monKST.setUTCDate(kst.getUTCDate() + diffMon);
    monKST.setUTCHours(0, 0, 0, 0);

    const friKST = new Date(monKST);
    friKST.setUTCDate(monKST.getUTCDate() + 4);
    friKST.setUTCHours(18, 0, 0, 0); // 18:00 KST

    // UTC equivalents (subtract 9h offset added above)
    const monUTC = new Date(monKST.getTime() - 9 * 3600 * 1000);
    const friUTC = new Date(friKST.getTime() - 9 * 3600 * 1000);

    return { monKST, friKST, monUTC, friUTC };
}

function inThisWeek(dateVal) {
    const d = parseExcelDateSafe(dateVal);
    if (!d) return false;
    const { monUTC, friUTC } = getWeekBounds();
    const ts = d.getTime();
    return ts >= monUTC.getTime() && ts <= friUTC.getTime() + 86400000;
}

function getNextWeekBounds() {
    const kst = getKSTNow();
    const day = kst.getUTCDay();
    const diffMon = day === 0 ? 1 : 8 - day;
    const monKST = new Date(kst);
    monKST.setUTCDate(kst.getUTCDate() + diffMon);
    monKST.setUTCHours(0, 0, 0, 0);
    const friKST = new Date(monKST);
    friKST.setUTCDate(monKST.getUTCDate() + 4);
    friKST.setUTCHours(23, 59, 59, 0);
    const monUTC = new Date(monKST.getTime() - 9 * 3600 * 1000);
    const friUTC = new Date(friKST.getTime() - 9 * 3600 * 1000);
    return { monKST, friKST, monUTC, friUTC };
}

function inNextWeek(dateVal) {
    const d = parseExcelDateSafe(dateVal);
    if (!d) return false;
    const { monUTC, friUTC } = getNextWeekBounds();
    const ts = d.getTime();
    return ts >= monUTC.getTime() && ts <= friUTC.getTime();
}

function isReportFinal() {
    const kst = getKSTNow();
    return kst.getUTCDay() === 5 && kst.getUTCHours() >= 18;
}

function fmtKST(d) {
    if (!d || isNaN(d.getTime())) return '—';
    return d.toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric' });
}

function fmtKSTDate(dateVal) {
    const d = parseExcelDateSafe(dateVal);
    return d ? fmtKST(d) : '—';
}

function daysSince(dateVal) {
    const d = parseExcelDateSafe(dateVal);
    if (!d) return null;
    const diff = Math.floor((Date.now() - d.getTime()) / 86400000);
    return diff >= 0 ? diff : null;
}

function fmtMoney(v) {
    if (!v || isNaN(v)) return '$0';
    if (Math.abs(v) >= 1e6) return `$${(v / 1e6).toFixed(1)}M`;
    if (Math.abs(v) >= 1000) return `$${(v / 1000).toFixed(0)}K`;
    return `$${Math.round(v).toLocaleString()}`;
}

function getWeekKey() {
    const { monUTC } = getWeekBounds();
    return monUTC.toISOString().slice(0, 10);
}

function getISOWeekNumber(kstDate) {
    // kstDate: Date where UTC values represent KST date parts
    const d = new Date(Date.UTC(kstDate.getUTCFullYear(), kstDate.getUTCMonth(), kstDate.getUTCDate()));
    const dayNum = d.getUTCDay() || 7; // Mon=1 … Sun=7
    d.setUTCDate(d.getUTCDate() + 4 - dayNum); // nearest Thursday
    const yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
    return Math.ceil((((d - yearStart) / 86400000) + 1) / 7);
}

function getISOWeeksInYear(year) {
    // Dec 28 always falls in the last ISO week of its year
    const dec28 = new Date(Date.UTC(year, 11, 28));
    const dayNum = dec28.getUTCDay() || 7;
    dec28.setUTCDate(dec28.getUTCDate() + 4 - dayNum);
    const yearStart = new Date(Date.UTC(dec28.getUTCFullYear(), 0, 1));
    return Math.ceil((((dec28 - yearStart) / 86400000) + 1) / 7);
}

/* ═══════════════════════════════════════════════════════════════
   Data Aggregation
   ═══════════════════════════════════════════════════════════════ */

function aggregateWeeklyStats(workbookData) {
    const { monKST, friKST } = getWeekBounds();
    const stats = {
        weekStart: new Date(monKST.getTime() - 9 * 3600 * 1000),
        weekEnd: new Date(friKST.getTime() - 9 * 3600 * 1000),
        newDeals: [], totalKtcv: 0, totalArr: 0, dealsByCountry: {},
        activePocs: [], newPocs: [], pocDecisions: [], pocWonThisWeek: [],
        pipelineNew: [], pipelineByCountry: {}, pipelineByCountryByQuarter: {}, pipelineTotalWeighted: 0,
        collectionSummary: [], collectionByDistributor: {},
        events: [],
        nextWeekEvents: [],
        partnerDealsThisWeek: {},
    };

    /* ── ORDER SHEET: new deals this week ── */
    const orderData = workbookData['ORDER SHEET'] || [];
    if (orderData.length > 0) {
        const keys = Object.keys(orderData[0]);
        const startKey = findContractStartKey(keys);
        const ktcvKey  = findKorTcvKey(keys);
        const arrKey   = findArrKey(keys);
        const nameKey  = findDealNameKey(keys);
        const cKey     = findCountryKey(keys);
        const partnerKey = findKey(keys, k => k.toLowerCase() === 'partner', k => k.toLowerCase().includes('partner'));
        const serviceKey = findKey(keys, k => k.toLowerCase().includes('service'), k => k.toLowerCase().includes('product'));

        orderData.forEach(row => {
            if (!inThisWeek(row[startKey])) return;
            const ktcv = parseCurrency(row[ktcvKey]) || 0;
            const arr  = parseCurrency(row[arrKey])  || 0;
            const country = normalizeCountry(row[cKey]) || 'Unknown';
            const partner = String(row[partnerKey] || '').trim() || 'Direct';
            stats.newDeals.push({
                name:    String(row[nameKey] || row[cKey] || 'N/A').trim(),
                country, partner,
                service: String(row[serviceKey] || '').trim(),
                ktcv, arr,
                dateVal: row[startKey]
            });
            stats.totalKtcv += ktcv;
            stats.totalArr  += arr;
            if (!stats.dealsByCountry[country]) stats.dealsByCountry[country] = { count: 0, ktcv: 0 };
            stats.dealsByCountry[country].count++;
            stats.dealsByCountry[country].ktcv += ktcv;
            if (partner && partner !== 'Direct') {
                stats.partnerDealsThisWeek[partner] = (stats.partnerDealsThisWeek[partner] || 0) + 1;
            }
        });
    }

    /* ── POC: active + new this week + decisions ── */
    const pocData = workbookData['POC'] || [];
    if (pocData.length > 0) {
        pocData.forEach(row => {
            const keys       = Object.keys(row);
            const cKey       = findCountryKey(keys);
            const sKey       = findStatusKey(keys);
            const pKey       = findKey(keys, k => k.toLowerCase() === 'partner');
            const startKey   = findPocStartKey(keys);
            const nameKey    = findKey(keys, k => k.toLowerCase().replace(/[^a-z]/g,'') === 'crmopportunityname' || k.toLowerCase().includes('end user') || k.toLowerCase().includes('company'));
            const valKey     = findEstimatedValueKey(keys);
            const decKey     = findKey(keys, k => k.toLowerCase().includes('decision'));

            const status  = String(row[sKey] || '').trim().toLowerCase();
            const country = normalizeCountry(row[cKey]) || '';
            const partner = String(row[pKey] || '').trim() || '—';
            const name    = String(row[nameKey] || '').trim() || 'N/A';
            const estVal  = valKey ? parseCurrency(row[valKey]) : 0;
            const decVal  = String(row[decKey] || '').trim();

            const isActive = status.includes('running') || status.includes('progress') || status.includes('active') || status.includes('ing');
            const isWon    = status.includes('won');

            if (isActive) {
                const isNewThisWeek = inThisWeek(row[startKey]);
                stats.activePocs.push({ name, country, partner, status: row[sKey], estVal, startVal: row[startKey], isNewThisWeek });
                if (isNewThisWeek) {
                    stats.newPocs.push({ name, country, partner, dateVal: row[startKey] });
                }
                if (decVal && decVal.toLowerCase() !== 'no' && decVal !== '') {
                    stats.pocDecisions.push({ name, country, partner, decision: decVal });
                }
            }
            if (isWon && inThisWeek(row[startKey])) {
                stats.pocWonThisWeek.push({ name, country, partner, estVal });
            }
        });
    }

    /* ── PIPELINE: new entries + full summary by country ── */
    const pipeData = workbookData['PIPELINE'] || [];
    if (pipeData.length > 0) {
        pipeData.forEach(row => {
            const keys     = Object.keys(row);
            const cKey     = findCountryKey(keys);
            const dateKey  = findContractStartKey(keys) || findKey(keys, k => k.toUpperCase().includes('DATE'));
            const wKey     = findKey(keys,
                k => k.toUpperCase().includes('WEIGHTED') && k.toUpperCase().includes('KOR TCV'),
                k => k.toUpperCase().includes('WEIGHTED') && k.toUpperCase().includes('USD'),
                k => k.toLowerCase().includes('weighted')
            );
            const vKey     = findKey(keys,
                k => k.toUpperCase().includes('KOR TCV') && !k.toUpperCase().includes('WEIGHTED'),
                k => k.toUpperCase().includes('TCV') && k.toUpperCase().includes('USD') && !k.toUpperCase().includes('WEIGHTED')
            ) || findEstimatedValueKey(keys);
            const nameKey  = findDealNameKey(keys) || findKey(keys, k => k.toLowerCase().includes('name'), k => k.toLowerCase().includes('end user'));
            const qKey     = findKey(keys, k => k.toLowerCase() === 'quarter', k => k.toLowerCase().includes('qtr'));

            const country  = normalizeCountry(row[cKey]) || 'Unknown';
            const weighted = parseCurrency(row[wKey]) || 0;
            const total    = parseCurrency(row[vKey]) || 0;

            const quarter = String(row[qKey] || '').trim() || 'Unknown';

            stats.pipelineTotalWeighted += weighted;
            if (!stats.pipelineByCountry[country]) stats.pipelineByCountry[country] = { count: 0, weighted: 0, total: 0, hasNewThisWeek: false };
            stats.pipelineByCountry[country].count++;
            stats.pipelineByCountry[country].weighted += weighted;
            stats.pipelineByCountry[country].total += total;

            if (!stats.pipelineByCountryByQuarter[country]) stats.pipelineByCountryByQuarter[country] = {};
            if (!stats.pipelineByCountryByQuarter[country][quarter]) stats.pipelineByCountryByQuarter[country][quarter] = { count: 0, weighted: 0, total: 0 };
            stats.pipelineByCountryByQuarter[country][quarter].count++;
            stats.pipelineByCountryByQuarter[country][quarter].weighted += weighted;
            stats.pipelineByCountryByQuarter[country][quarter].total += total;

            if (inThisWeek(row[dateKey])) {
                stats.pipelineByCountry[country].hasNewThisWeek = true;
                stats.pipelineNew.push({
                    name:    String(row[nameKey] || '').trim() || 'N/A',
                    country, weighted, total,
                    quarter: String(row[qKey] || '').trim()
                });
            }
        });
    }

    /* ── COLLECTION: unpaid / outstanding summary ── */
    const colData = workbookData['COLLECTION'] || [];
    if (colData.length > 0) {
        colData.forEach(row => {
            const keys        = Object.keys(row);
            const ktcvKey     = findKorTcvKey(keys);
            const distributor = String(row['Distributor'] || row['Partner'] || '').trim() || 'Unknown';
            const endUser     = String(row['End User'] || '').trim() || '—';
            const ktcv        = parseCurrency(row[ktcvKey]) || 0;
            let received = 0;
            keys.forEach(k => { if (k.toUpperCase().includes('RECEIVED')) received += parseCurrency(row[k]) || 0; });
            const outstanding = ktcv - received;
            if (outstanding > 0) {
                stats.collectionSummary.push({ endUser, distributor, ktcv, received, outstanding });
                if (!stats.collectionByDistributor[distributor]) stats.collectionByDistributor[distributor] = 0;
                stats.collectionByDistributor[distributor] += outstanding;
            }
        });
        stats.collectionSummary.sort((a, b) => b.outstanding - a.outstanding);
    }

    /* ── EVENT: events this week ── */
    const eventData = workbookData['EVENT'] || [];
    if (eventData.length > 0) {
        eventData.forEach(row => {
            const keys    = Object.keys(row);
            const dateKey = findKey(keys, k => k.toLowerCase().includes('date')) || keys[0];
            const nKey    = findKey(keys, k => k.toLowerCase() === 'event name', k => k.toLowerCase().includes('name'));
            const cKey    = findCountryKey(keys);
            const tKey    = findKey(keys, k => k.toLowerCase().includes('type') || k.toLowerCase().includes('category'));
            const entry = {
                name:    String(row[nKey] || '').trim() || 'Event',
                country: String(row[cKey] || '').trim(),
                type:    String(row[tKey] || '').trim(),
                dateVal: row[dateKey]
            };
            if (inThisWeek(row[dateKey]))  stats.events.push(entry);
            if (inNextWeek(row[dateKey]))  stats.nextWeekEvents.push(entry);
        });
    }

    return stats;
}

/* ═══════════════════════════════════════════════════════════════
   Auto-generated Next Week Suggestions
   ═══════════════════════════════════════════════════════════════ */

function buildSuggestions(stats) {
    const items = [];

    if (stats.pocDecisions.length > 0) {
        const names = stats.pocDecisions.slice(0, 3).map(p => p.name).join(', ');
        items.push(`Follow up on <strong>${stats.pocDecisions.length} POC decision(s)</strong> pending: ${names}${stats.pocDecisions.length > 3 ? '…' : ''}`);
    }
    if (stats.newPocs.length > 0) {
        items.push(`Ensure kickoff completion for <strong>${stats.newPocs.length} new POC(s)</strong> started this week`);
    }
    if (stats.pipelineNew.length > 0) {
        items.push(`Advance <strong>${stats.pipelineNew.length} new pipeline deal(s)</strong> — schedule discovery calls`);
    }
    const topCountries = Object.entries(stats.pipelineByCountry)
        .sort((a, b) => b[1].weighted - a[1].weighted).slice(0, 2).map(([c]) => c);
    if (topCountries.length > 0) {
        items.push(`Prioritize pipeline in <strong>${topCountries.join(' & ')}</strong> — highest weighted value`);
    }
    if (stats.newDeals.length > 0) {
        items.push(`Initiate onboarding for <strong>${stats.newDeals.length} deal(s)</strong> signed this week`);
    }
    const totalOutstanding = stats.collectionSummary.reduce((s, r) => s + r.outstanding, 0);
    if (totalOutstanding > 0) {
        items.push(`Chase outstanding collections: <strong>${fmtMoney(totalOutstanding)}</strong> receivable`);
    }
    if (stats.activePocs.length > 0) {
        items.push(`Review status of <strong>${stats.activePocs.length} active POC(s)</strong> — identify blockers`);
    }
    if (items.length === 0) {
        items.push('Review overall pipeline health and update deal stages in CRM');
        items.push('Check in with regional partners on active opportunities');
    }
    return items;
}

/* ═══════════════════════════════════════════════════════════════
   Next Week Reference Auto-items
   ═══════════════════════════════════════════════════════════════ */

function pocLabel(p) {
    if (p.name && p.name !== 'N/A' && p.name !== '') return p.name;
    if (p.partner && p.partner !== '—' && p.partner !== '') return p.partner;
    return p.country || 'Unknown';
}

function buildNextWeekReferences(stats) {
    const items = [];

    if (stats.nextWeekEvents.length > 0) {
        stats.nextWeekEvents.forEach(e => {
            items.push({
                icon: 'fa-calendar-alt',
                tag: 'Event',
                tagClass: 'ref-tag-event',
                text: `<strong>${e.name}</strong>${e.country ? ` — ${e.country}` : ''}${e.type ? ` <em>(${e.type})</em>` : ''} · ${fmtKSTDate(e.dateVal)}`
            });
        });
    }

    if (stats.pocDecisions.length > 0) {
        stats.pocDecisions.forEach(p => {
            const label = pocLabel(p);
            items.push({
                icon: 'fa-gavel',
                tag: 'Decision',
                tagClass: 'ref-tag-decision',
                text: `<strong>${label}</strong> · ${p.country} · Partner: ${p.partner} — Go/No-Go decision required`
            });
        });
    }

    const longPocs = stats.activePocs.filter(p => {
        const d = daysSince(p.startVal);
        return d !== null && d >= 30;
    }).sort((a, b) => (daysSince(b.startVal) || 0) - (daysSince(a.startVal) || 0));
    longPocs.slice(0, 5).forEach(p => {
        const d = daysSince(p.startVal);
        const label = pocLabel(p);
        items.push({
            icon: 'fa-hourglass-half',
            tag: 'Long POC',
            tagClass: 'ref-tag-poc',
            text: `<strong>${label}</strong> · ${p.country} · Partner: ${p.partner} — Running ${d} days, status check needed`
        });
    });

    const topCollections = stats.collectionSummary.slice(0, 3);
    topCollections.forEach(c => {
        items.push({
            icon: 'fa-file-invoice-dollar',
            tag: 'Collection',
            tagClass: 'ref-tag-col',
            text: `<strong>${c.endUser}</strong> · ${c.distributor} — Outstanding: <strong>${fmtMoney(c.outstanding)}</strong> — follow up required`
        });
    });

    const topPipe = Object.entries(stats.pipelineByCountry)
        .sort((a, b) => b[1].weighted - a[1].weighted).slice(0, 2);
    topPipe.forEach(([country, d]) => {
        items.push({
            icon: 'fa-funnel-dollar',
            tag: 'Pipeline',
            tagClass: 'ref-tag-pipe',
            text: `<strong>${country}</strong> — ${d.count} deal(s) · Weighted ${fmtMoney(d.weighted)} — advance pipeline`
        });
    });

    return items;
}

/* ═══════════════════════════════════════════════════════════════
   Country Strategic Insights
   ═══════════════════════════════════════════════════════════════ */

function buildCountryStrategies(stats) {
    const map = {};
    const ensure = c => {
        if (!map[c]) map[c] = {
            pipelineCount: 0, pipelineWeighted: 0, pipelineTotal: 0, pipelineNewThisWeek: false,
            dealsCount: 0, dealsKtcv: 0,
            activePocCount: 0, longPocs: [], decisionPocs: [], newPocCount: 0,
        };
        return map[c];
    };

    Object.entries(stats.pipelineByCountry).forEach(([c, d]) => {
        const s = ensure(c);
        s.pipelineCount     = d.count;
        s.pipelineWeighted  = d.weighted;
        s.pipelineTotal     = d.total;
        s.pipelineNewThisWeek = d.hasNewThisWeek;
    });

    Object.entries(stats.dealsByCountry).forEach(([c, d]) => {
        const s = ensure(c);
        s.dealsCount = d.count;
        s.dealsKtcv  = d.ktcv;
    });

    stats.activePocs.forEach(p => {
        if (!p.country) return;
        const s = ensure(p.country);
        s.activePocCount++;
        if (p.isNewThisWeek) s.newPocCount++;
        const d = daysSince(p.startVal);
        if (d !== null && d >= 30) s.longPocs.push({ ...p, days: d });
    });

    stats.pocDecisions.forEach(p => {
        if (!p.country) return;
        ensure(p.country).decisionPocs.push(p);
    });

    return Object.entries(map)
        .filter(([c]) => c && c !== 'Unknown' && c !== '')
        .map(([country, d]) => {
            const recs = [];
            let urgencyScore = 0;

            if (d.decisionPocs.length > 0) {
                urgencyScore += d.decisionPocs.length * 3;
                const names = d.decisionPocs.slice(0, 2).map(p => pocLabel(p)).join(', ');
                recs.push({ level: 'high', text: `<strong>${d.decisionPocs.length} POC(s) awaiting go/no-go decision</strong> — ${names}${d.decisionPocs.length > 2 ? '…' : ''} · Confirm customer readiness and escalate immediately` });
            }

            if (d.longPocs.length > 0) {
                urgencyScore += d.longPocs.length * 2;
                d.longPocs.sort((a, b) => b.days - a.days);
                const worst = d.longPocs[0];
                recs.push({ level: 'high', text: `<strong>${d.longPocs.length} POC(s) running 30+ days</strong> (longest: ${worst.days}d — ${pocLabel(worst)}) · Identify blockers and push for resolution` });
            }

            if (d.newPocCount > 0) {
                recs.push({ level: 'medium', text: `<strong>${d.newPocCount} new POC(s) kicked off</strong> this week — confirm environment setup and stakeholder alignment` });
            }

            if (d.pipelineNewThisWeek) {
                recs.push({ level: 'medium', text: `New pipeline entry added this week — assign owner, schedule discovery call, and qualify within 5 business days` });
            }

            if (d.pipelineCount > 0 && !d.pipelineNewThisWeek) {
                urgencyScore += 1;
                recs.push({ level: 'medium', text: `<strong>${d.pipelineCount} active deal(s)</strong> · Weighted ${fmtMoney(d.pipelineWeighted)} — review stage progression and update CRM` });
            }

            if (d.dealsCount > 0) {
                recs.push({ level: 'low', text: `<strong>${d.dealsCount} deal(s) signed</strong> this week (${fmtMoney(d.dealsKtcv)} KTCV) — initiate onboarding and issue invoice promptly` });
            }

            if (d.activePocCount === 0 && d.pipelineCount === 0 && d.dealsCount === 0) {
                urgencyScore += 3;
                recs.push({ level: 'high', text: `<strong>No active POCs or pipeline</strong> — review market strategy, identify top 3 target accounts, and launch outreach this week` });
            }

            const urgency = urgencyScore >= 5 ? 'high' : urgencyScore >= 2 ? 'medium' : 'low';

            return { country, flag: COUNTRY_FLAGS[country] || '', team: COUNTRY_TEAMS[country] || [], data: d, urgency, urgencyScore, recs };
        })
        .sort((a, b) => b.urgencyScore - a.urgencyScore);
}

/* ═══════════════════════════════════════════════════════════════
   HTML Template
   ═══════════════════════════════════════════════════════════════ */

function buildReportHTML(stats, savedNotes, savedRefs) {
    const { monKST, friKST } = getWeekBounds();
    const weekLabel = `${fmtKST(new Date(monKST.getTime() - 9*3600*1000))} – ${fmtKST(new Date(friKST.getTime() - 9*3600*1000))}`;
    const kstNow = getKSTNow();
    const currentWeek = getISOWeekNumber(kstNow);
    const totalWeeks = getISOWeeksInYear(kstNow.getUTCFullYear());
    const { monKST: nxMon, friKST: nxFri } = getNextWeekBounds();
    const nextWeekLabel = `${fmtKST(new Date(nxMon.getTime() - 9*3600*1000))} – ${fmtKST(new Date(nxFri.getTime() - 9*3600*1000))}`;
    const isFinal  = isReportFinal();
    const suggestions = buildSuggestions(stats);
    const refItems = buildNextWeekReferences(stats);
    const countryInsights = buildCountryStrategies(stats);
    const totalOutstanding = stats.collectionSummary.reduce((s, r) => s + r.outstanding, 0);
    const sortedActivePocs = [...stats.activePocs].sort((a, b) => (b.isNewThisWeek ? 1 : 0) - (a.isNewThisWeek ? 1 : 0));

    return `
<div class="wr-wrap" id="wr-printable">

  <!-- ── Header ── -->
  <div class="wr-header">
    <div class="wr-header-left">
      <img src="whatap-logo.png" alt="WhaTap" class="wr-logo">
      <div>
        <div class="wr-title">Weekly Business Report</div>
        <div class="wr-week">
          <span class="wr-week-num">Week ${currentWeek} / ${totalWeeks}</span>
          <span class="wr-week-range">${weekLabel}</span>
        </div>
      </div>
    </div>
    <div class="wr-header-right">
      <span class="wr-badge ${isFinal ? 'wr-final' : 'wr-draft'}">${isFinal ? 'FINAL' : 'DRAFT'}</span>
      <span class="wr-gen">Generated ${new Date().toLocaleString('ko-KR', {timeZone:'Asia/Seoul'}).replace(/\./g,'.').trim()} KST</span>
      <button class="wr-pdf-btn no-print" onclick="window.downloadWeeklyReportPDF()">
        <i class="fas fa-file-pdf"></i> Download PDF
      </button>
    </div>
  </div>

  <!-- ── KPI Strip ── -->
  <div class="wr-kpi-strip">
    <div class="wr-kpi-card wr-kpi-deals">
      <div class="wr-kpi-icon"><i class="fas fa-handshake"></i></div>
      <div class="wr-kpi-val">${stats.newDeals.length}</div>
      <div class="wr-kpi-lbl">Deals Signed</div>
      <div class="wr-kpi-sub">${fmtMoney(stats.totalKtcv)} KTCV</div>
    </div>
    <div class="wr-kpi-card wr-kpi-poc">
      <div class="wr-kpi-icon"><i class="fas fa-flask"></i></div>
      <div class="wr-kpi-val">${stats.newPocs.length}</div>
      <div class="wr-kpi-lbl">New POCs This Week</div>
      <div class="wr-kpi-sub">${stats.activePocs.length} active total</div>
    </div>
    <div class="wr-kpi-card wr-kpi-pipe">
      <div class="wr-kpi-icon"><i class="fas fa-chart-line"></i></div>
      <div class="wr-kpi-val">${Object.values(stats.pipelineByCountry).reduce((s,c)=>s+c.count,0)}</div>
      <div class="wr-kpi-lbl">Pipeline Deals</div>
      <div class="wr-kpi-sub">${fmtMoney(stats.pipelineTotalWeighted)} weighted</div>
    </div>
    <div class="wr-kpi-card wr-kpi-col">
      <div class="wr-kpi-icon"><i class="fas fa-coins"></i></div>
      <div class="wr-kpi-val">${stats.collectionSummary.length}</div>
      <div class="wr-kpi-lbl">Pending Collections</div>
      <div class="wr-kpi-sub">${fmtMoney(totalOutstanding)} outstanding</div>
    </div>
    <div class="wr-kpi-card wr-kpi-evt">
      <div class="wr-kpi-icon"><i class="fas fa-calendar-check"></i></div>
      <div class="wr-kpi-val">${stats.events.length}</div>
      <div class="wr-kpi-lbl">Events</div>
      <div class="wr-kpi-sub">this week</div>
    </div>
  </div>

  <!-- ── Main Grid ── -->
  <div class="wr-grid">

    <!-- Deals Signed -->
    <div class="wr-section wr-full">
      <div class="wr-sec-hdr">
        <i class="fas fa-handshake"></i>
        <span>Deals Signed This Week</span>
        <span class="wr-count">${stats.newDeals.length}</span>
      </div>
      ${stats.newDeals.length > 0 ? `
      <table class="wr-table">
        <thead><tr><th>Company / Deal</th><th>Country</th><th>Partner</th><th>Service</th><th class="num">KTCV (USD)</th><th class="num">ARR (USD)</th><th>Date</th></tr></thead>
        <tbody>${stats.newDeals.map(d=>`
          <tr>
            <td><strong>${d.name}</strong></td>
            <td>${d.country}</td>
            <td>${d.partner}</td>
            <td>${d.service || '—'}</td>
            <td class="num">${fmtMoney(d.ktcv)}</td>
            <td class="num">${fmtMoney(d.arr)}</td>
            <td>${fmtKSTDate(d.dateVal)}</td>
          </tr>`).join('')}
        </tbody>
        <tfoot><tr><td colspan="4"><strong>Total</strong></td><td class="num"><strong>${fmtMoney(stats.totalKtcv)}</strong></td><td class="num"><strong>${fmtMoney(stats.totalArr)}</strong></td><td></td></tr></tfoot>
      </table>` : `<div class="wr-empty">No deals signed this week.</div>`}
    </div>

    <!-- POC Status -->
    <div class="wr-section wr-half">
      <div class="wr-sec-hdr">
        <i class="fas fa-flask"></i>
        <span>POC Status</span>
      </div>
      <div class="wr-poc-summary">
        <div class="wr-poc-row"><span>Active POCs</span><strong>${stats.activePocs.length}</strong></div>
        <div class="wr-poc-row"><span>New this week</span><strong class="green">${stats.newPocs.length}</strong></div>
        <div class="wr-poc-row"><span>Pending decisions</span><strong class="amber">${stats.pocDecisions.length}</strong></div>
        <div class="wr-poc-row"><span>Won this week</span><strong class="blue">${stats.pocWonThisWeek.length}</strong></div>
      </div>
      ${stats.pocDecisions.length > 0 ? `
      <div class="wr-alert"><i class="fas fa-exclamation-triangle"></i> ${stats.pocDecisions.length} POC(s) require a go/no-go decision</div>` : ''}
      ${sortedActivePocs.length > 0 ? `
      <table class="wr-table wr-compact">
        <thead><tr><th>End User</th><th>Country</th><th>Partner</th><th>Status</th><th class="num">Days</th></tr></thead>
        <tbody>${sortedActivePocs.slice(0,10).map(p => {
          const days = daysSince(p.startVal);
          return `
          <tr${p.isNewThisWeek ? ' style="background:#f0fdf4;"' : ''}>
            <td><strong>${p.name}</strong>${p.isNewThisWeek ? ' <span class="wr-new-badge">NEW</span>' : ''}</td>
            <td>${p.country}</td>
            <td>${p.partner}</td>
            <td><span class="wr-pill active">${p.status||'Active'}</span></td>
            <td class="num">${days !== null ? `${days}d` : '—'}</td>
          </tr>`;
        }).join('')}
          ${sortedActivePocs.length>10?`<tr><td colspan="5" class="wr-more">+${sortedActivePocs.length-10} more active POCs</td></tr>`:''}
        </tbody>
      </table>` : `<div class="wr-empty">No active POCs.</div>`}
    </div>

    <!-- Pipeline Update -->
    <div class="wr-section wr-half">
      <div class="wr-sec-hdr">
        <i class="fas fa-chart-line"></i>
        <span>Pipeline Update</span>
      </div>
      <div class="wr-subsec">Pipeline by Country (all active)</div>
      ${Object.keys(stats.pipelineByCountry).length > 0 ? `
      <table class="wr-table wr-compact">
        <thead><tr><th>Country</th><th class="num">Deals</th><th class="num">Total Value</th><th class="num">Weighted Value</th></tr></thead>
        <tbody>${Object.entries(stats.pipelineByCountry)
          .sort((a,b) => {
            if (a[1].hasNewThisWeek && !b[1].hasNewThisWeek) return -1;
            if (!a[1].hasNewThisWeek && b[1].hasNewThisWeek) return 1;
            return b[1].weighted - a[1].weighted;
          }).slice(0,8).map(([c,d])=>`
          <tr${d.hasNewThisWeek ? ' class="wr-pipe-new"' : ''}>
            <td>${d.hasNewThisWeek ? '<span class="wr-star">★</span> ' : ''}${c}</td>
            <td class="num">${d.count}</td>
            <td class="num">${fmtMoney(d.total)}</td>
            <td class="num">${fmtMoney(d.weighted)}</td>
          </tr>`).join('')}
        </tbody>
      </table>` : `<div class="wr-empty">No pipeline data.</div>`}
      ${stats.pipelineNew.length > 0 ? `
      <div class="wr-subsec" style="margin-top:12px">New Entries This Week (${stats.pipelineNew.length})</div>
      <table class="wr-table wr-compact">
        <thead><tr><th>Deal</th><th>Country</th><th class="num">Total</th><th class="num">Weighted</th><th>Quarter</th></tr></thead>
        <tbody>${stats.pipelineNew.slice(0,5).map(p=>`
          <tr><td><strong>★ ${p.name}</strong></td><td>${p.country}</td><td class="num">${fmtMoney(p.total)}</td><td class="num">${fmtMoney(p.weighted)}</td><td>${p.quarter||'—'}</td></tr>`).join('')}
        </tbody>
      </table>` : ''}
    </div>

    <!-- Pipeline by Country by Quarter -->
    ${(() => {
      const cbq = stats.pipelineByCountryByQuarter;
      if (!cbq || Object.keys(cbq).length === 0) return '';
      const allQuarters = [...new Set(
        Object.values(cbq).flatMap(q => Object.keys(q))
      )].sort();
      const countries = Object.keys(cbq).sort((a, b) => {
        const wa = Object.values(cbq[a]).reduce((s,d)=>s+d.weighted,0);
        const wb = Object.values(cbq[b]).reduce((s,d)=>s+d.weighted,0);
        return wb - wa;
      });
      return `
    <div class="wr-section wr-full">
      <div class="wr-sec-hdr">
        <i class="fas fa-table"></i>
        <span>Pipeline by Country &amp; Quarter</span>
        <span class="wr-week-sub">Weighted value (USD)</span>
      </div>
      <table class="wr-table wr-compact wr-pipe-qtr-tbl">
        <thead>
          <tr>
            <th>Country</th>
            ${allQuarters.map(q=>`<th class="num">${q||'—'}</th>`).join('')}
            <th class="num">Total</th>
          </tr>
        </thead>
        <tbody>
          ${countries.map(country => {
            const qMap = cbq[country];
            const rowTotal = Object.values(qMap).reduce((s,d)=>s+d.weighted,0);
            return `<tr>
              <td><strong>${country}</strong></td>
              ${allQuarters.map(q => {
                const d = qMap[q];
                return `<td class="num">${d ? `<span title="${d.count} deal(s)">${fmtMoney(d.weighted)}</span>` : '<span class="wr-dim">—</span>'}</td>`;
              }).join('')}
              <td class="num"><strong>${fmtMoney(rowTotal)}</strong></td>
            </tr>`;
          }).join('')}
          <tr class="wr-tfoot-row">
            <td><strong>Total</strong></td>
            ${allQuarters.map(q => {
              const s = countries.reduce((acc,c) => acc + (cbq[c][q]?.weighted||0), 0);
              return `<td class="num"><strong>${fmtMoney(s)}</strong></td>`;
            }).join('')}
            <td class="num"><strong>${fmtMoney(countries.reduce((acc,c)=>acc+Object.values(cbq[c]).reduce((s,d)=>s+d.weighted,0),0))}</strong></td>
          </tr>
        </tbody>
      </table>
    </div>`;
    })()}

    <!-- Collections -->
    <div class="wr-section wr-half">
      <div class="wr-sec-hdr">
        <i class="fas fa-coins"></i>
        <span>Collection Status</span>
      </div>
      ${stats.collectionSummary.length > 0 ? `
      <div class="wr-poc-row"><span>Outstanding balance</span><strong class="amber">${fmtMoney(totalOutstanding)}</strong></div>
      <table class="wr-table wr-compact" style="margin-top:10px">
        <thead><tr><th>End User</th><th>Distributor</th><th class="num">KTCV</th><th class="num">Received</th><th class="num">Outstanding</th></tr></thead>
        <tbody>${stats.collectionSummary.slice(0,6).map(c=>`
          <tr>
            <td><strong>${c.endUser}</strong></td>
            <td>${c.distributor}</td>
            <td class="num">${fmtMoney(c.ktcv)}</td>
            <td class="num">${fmtMoney(c.received)}</td>
            <td class="num amber"><strong>${fmtMoney(c.outstanding)}</strong></td>
          </tr>`).join('')}
          ${stats.collectionSummary.length>6?`<tr><td colspan="5" class="wr-more">+${stats.collectionSummary.length-6} more rows</td></tr>`:''}
        </tbody>
      </table>` : `<div class="wr-empty">No outstanding collections.</div>`}
    </div>

    <!-- Events -->
    <div class="wr-section wr-half">
      <div class="wr-sec-hdr">
        <i class="fas fa-calendar-check"></i>
        <span>Events This Week</span>
        <span class="wr-count">${stats.events.length}</span>
      </div>
      ${stats.events.length > 0 ? `
      <table class="wr-table wr-compact">
        <thead><tr><th>Event</th><th>Country</th><th>Type</th><th>Date</th></tr></thead>
        <tbody>${stats.events.map(e=>`
          <tr>
            <td><strong>${e.name}</strong></td>
            <td>${e.country||'—'}</td>
            <td>${e.type||'—'}</td>
            <td>${fmtKSTDate(e.dateVal)}</td>
          </tr>`).join('')}
        </tbody>
      </table>` : `<div class="wr-empty">No events this week.</div>`}
    </div>

    <!-- Next Week Action Plan -->
    <div class="wr-section wr-full wr-actions-section">
      <div class="wr-sec-hdr">
        <i class="fas fa-tasks"></i>
        <span>Next Week Action Plan</span>
        <button class="wr-edit-btn no-print" onclick="window._wrToggleEdit()">
          <i class="fas fa-pen"></i> Edit notes
        </button>
      </div>

      <div class="wr-actions-grid">
        <div class="wr-auto-block">
          <div class="wr-block-lbl"><i class="fas fa-robot"></i> Suggested Actions</div>
          <ul class="wr-suggestions">${suggestions.map(s=>`<li>${s}</li>`).join('')}</ul>
        </div>
        <div class="wr-notes-block">
          <div class="wr-block-lbl"><i class="fas fa-pen-nib"></i> Team Notes</div>
          <div id="wr-notes-view" class="wr-notes-view">${savedNotes
            ? savedNotes.replace(/\n/g,'<br>')
            : '<span class="wr-empty">Click "Edit notes" to add next-week action items…</span>'}</div>
          <div id="wr-notes-edit" style="display:none">
            <textarea id="wr-notes-ta" class="wr-notes-ta" placeholder="Enter next-week action items, goals, follow-ups…">${savedNotes||''}</textarea>
            <div class="wr-notes-btns">
              <button class="wr-save-btn" onclick="window._wrSaveNotes()"><i class="fas fa-save"></i> Save</button>
              <button class="wr-cancel-btn" onclick="window._wrCancelEdit()"><i class="fas fa-times"></i> Cancel</button>
            </div>
          </div>
        </div>
      </div>
    </div>

    <!-- Next Week References -->
    <div class="wr-section wr-full wr-refs-section">
      <div class="wr-sec-hdr">
        <i class="fas fa-bookmark"></i>
        <span>Next Week References</span>
        <span class="wr-week-sub">${nextWeekLabel}</span>
        <button class="wr-edit-btn no-print" onclick="window._wrToggleRefs()">
          <i class="fas fa-pen"></i> Add Reference
        </button>
      </div>

      <div class="wr-actions-grid">
        <div class="wr-auto-block">
          <div class="wr-block-lbl"><i class="fas fa-robot"></i> Auto-detected Items</div>
          ${refItems.length > 0 ? `
          <ul class="wr-ref-list">${refItems.map(r => `
            <li class="wr-ref-item">
              <span class="wr-ref-tag ${r.tagClass}">${r.tag}</span>
              <i class="fas ${r.icon} wr-ref-icon"></i>
              <span>${r.text}</span>
            </li>`).join('')}
          </ul>` : `<div class="wr-empty">No auto-detected items for next week.</div>`}
        </div>
        <div class="wr-notes-block">
          <div class="wr-block-lbl"><i class="fas fa-link"></i> Reference Links / Docs / Notes</div>
          <div id="wr-refs-view" class="wr-notes-view">${savedRefs
            ? savedRefs.replace(/\n/g, '<br>')
            : '<span class="wr-empty">Click "Add Reference" to add links, documents, or notes for next week…</span>'}</div>
          <div id="wr-refs-edit" style="display:none">
            <textarea id="wr-refs-ta" class="wr-notes-ta" placeholder="e.g. QBR deck: https://…&#10;Partner meeting agenda: …&#10;Renewal contracts due: …">${savedRefs || ''}</textarea>
            <div class="wr-notes-btns">
              <button class="wr-save-btn" onclick="window._wrSaveRefs()"><i class="fas fa-save"></i> Save</button>
              <button class="wr-cancel-btn" onclick="window._wrCancelRefs()"><i class="fas fa-times"></i> Cancel</button>
            </div>
          </div>
        </div>
      </div>
    </div>

    <!-- Country Strategic Insights -->
    <div class="wr-section wr-full wr-strategy-section">
      <div class="wr-sec-hdr">
        <i class="fas fa-chess-king"></i>
        <span>Country Strategic Insights</span>
        <span class="wr-week-sub">Priority-ranked by urgency</span>
      </div>
      ${countryInsights.length > 0 ? `
      <div class="wr-country-grid">
        ${countryInsights.map(ins => `
        <div class="wr-country-card urgency-${ins.urgency}">
          <div class="wr-country-card-hdr">
            <span class="wr-country-flag-lg">${ins.flag}</span>
            <strong class="wr-country-name">${ins.country}</strong>
            <span class="wr-urgency-chip chip-${ins.urgency}">${ins.urgency === 'high' ? '⚠ High Priority' : ins.urgency === 'medium' ? '◉ Medium' : '✓ Stable'}</span>
          </div>
          <div class="wr-country-mini-stats">
            <span><i class="fas fa-chart-line"></i> ${ins.data.pipelineCount} pipeline</span>
            <span><i class="fas fa-flask"></i> ${ins.data.activePocCount} POC${ins.data.activePocCount !== 1 ? 's' : ''}</span>
            <span><i class="fas fa-handshake"></i> ${ins.data.dealsCount} deal${ins.data.dealsCount !== 1 ? 's' : ''} signed</span>
            ${ins.data.pipelineWeighted > 0 ? `<span><i class="fas fa-dollar-sign"></i> ${fmtMoney(ins.data.pipelineWeighted)} wtd</span>` : ''}
          </div>
          <ul class="wr-country-recs">
            ${ins.recs.map(r => `<li class="rec-${r.level}"><i class="fas ${r.level === 'high' ? 'fa-exclamation-circle' : r.level === 'medium' ? 'fa-arrow-circle-right' : 'fa-check-circle'}"></i> ${r.text}</li>`).join('')}
          </ul>
        </div>`).join('')}
      </div>` : `<div class="wr-empty">No country data available.</div>`}
    </div>

  </div><!-- /wr-grid -->
</div><!-- /wr-wrap -->`;
}

/* ═══════════════════════════════════════════════════════════════
   Notes & References Persistence
   ═══════════════════════════════════════════════════════════════ */

function loadNotes() {
    return localStorage.getItem(`wr-notes-${getWeekKey()}`) || '';
}

function loadRefs() {
    return localStorage.getItem(`wr-refs-${getWeekKey()}`) || '';
}

window._wrToggleEdit = function () {
    document.getElementById('wr-notes-view').style.display = 'none';
    document.getElementById('wr-notes-edit').style.display = 'block';
    document.getElementById('wr-notes-ta').focus();
};

window._wrCancelEdit = function () {
    document.getElementById('wr-notes-view').style.display = 'block';
    document.getElementById('wr-notes-edit').style.display = 'none';
};

window._wrSaveNotes = function () {
    const text = document.getElementById('wr-notes-ta').value;
    localStorage.setItem(`wr-notes-${getWeekKey()}`, text);
    document.getElementById('wr-notes-view').innerHTML = text
        ? text.replace(/\n/g, '<br>')
        : '<span class="wr-empty">Click "Edit notes" to add next-week action items…</span>';
    window._wrCancelEdit();
};

window._wrToggleRefs = function () {
    document.getElementById('wr-refs-view').style.display = 'none';
    document.getElementById('wr-refs-edit').style.display = 'block';
    document.getElementById('wr-refs-ta').focus();
};

window._wrCancelRefs = function () {
    document.getElementById('wr-refs-view').style.display = 'block';
    document.getElementById('wr-refs-edit').style.display = 'none';
};

window._wrSaveRefs = function () {
    const text = document.getElementById('wr-refs-ta').value;
    localStorage.setItem(`wr-refs-${getWeekKey()}`, text);
    document.getElementById('wr-refs-view').innerHTML = text
        ? text.replace(/\n/g, '<br>')
        : '<span class="wr-empty">차주 필요한 자료 링크나 메모를 입력하세요…</span>';
    window._wrCancelRefs();
};

/* ═══════════════════════════════════════════════════════════════
   PDF Download
   ═══════════════════════════════════════════════════════════════ */

window.downloadWeeklyReportPDF = function () {
    const el = document.getElementById('wr-printable');
    if (!el) return;
    if (typeof html2pdf === 'undefined') { window.print(); return; }

    const { monUTC, friUTC } = getWeekBounds();
    const filename = `WhaTap_Weekly_Report_${monUTC.toISOString().slice(0,10)}_to_${friUTC.toISOString().slice(0,10)}.pdf`;

    // Loading overlay
    const overlay = document.createElement('div');
    overlay.style.cssText = 'position:fixed;inset:0;background:rgba(17,24,39,0.55);z-index:99999;display:flex;align-items:center;justify-content:center;';
    overlay.innerHTML = '<div style="background:#fff;padding:20px 32px;border-radius:12px;font-size:0.95rem;font-weight:600;color:#374151;display:flex;align-items:center;gap:10px;"><i class="fas fa-spinner fa-spin" style="color:#007AFF"></i> Generating PDF…</div>';
    document.body.appendChild(overlay);

    // Only hide interactive buttons — do NOT touch layout at all
    const noPrint = el.querySelectorAll('.no-print');
    noPrint.forEach(n => n.style.visibility = 'hidden');

    requestAnimationFrame(() => {
        html2pdf().set({
            margin:      [8, 8, 8, 8],
            filename,
            image:       { type: 'jpeg', quality: 0.97 },
            html2canvas: {
                scale:       2,
                useCORS:     true,
                logging:     false,
                windowWidth: window.innerWidth,   // match actual screen — no layout change
                scrollX:     0,
                scrollY:     -window.scrollY      // correct for page scroll
            },
            jsPDF:       { unit: 'mm', format: 'a4', orientation: 'landscape' },
            pagebreak:   { mode: ['css', 'legacy'], avoid: '.wr-section, .wr-country-card, .wr-kpi-strip' }
        }).from(el).save().finally(() => {
            noPrint.forEach(n => n.style.visibility = '');
            document.body.removeChild(overlay);
        });
    });
};

/* ═══════════════════════════════════════════════════════════════
   Sidebar Notification Badge
   ═══════════════════════════════════════════════════════════════ */

export function checkWeeklyReportBadge() {
    if (!isReportFinal()) return;
    const seenKey = `wr-seen-${getWeekKey()}`;
    if (localStorage.getItem(seenKey)) return;
    // Delay slightly so sidebar is rendered
    setTimeout(() => {
        const btn = document.querySelector('.weekly-report-tab');
        if (btn && !btn.querySelector('.wr-badge-dot')) {
            const dot = document.createElement('span');
            dot.className = 'wr-badge-dot';
            dot.title = 'Weekly report ready';
            btn.appendChild(dot);
        }
    }, 800);
}

/* ═══════════════════════════════════════════════════════════════
   Main Entry Point
   ═══════════════════════════════════════════════════════════════ */

export function selectWeeklyReportView(setCurrentTab, workbookData) {
    setCurrentTab('WEEKLY_REPORT');

    // Mark notification as seen
    localStorage.setItem(`wr-seen-${getWeekKey()}`, '1');
    const dot = document.querySelector('.weekly-report-tab .wr-badge-dot');
    if (dot) dot.remove();

    // Update header
    const titleEl = document.getElementById('current-tab-title');
    if (titleEl) titleEl.innerText = 'Weekly Report';

    // Hide table, clear active nav state
    document.querySelectorAll('.nav-item').forEach(i => i.classList.remove('active'));
    const dataSection = document.querySelector('.data-section');
    if (dataSection) dataSection.classList.add('hidden');

    const grid = document.getElementById('tab-metrics-grid');
    if (!grid) return;

    grid.classList.remove('hidden');
    grid.style.display = 'block';
    grid.innerHTML = '<div style="padding:40px;text-align:center;color:#888"><i class="fas fa-spinner fa-spin fa-2x"></i><p style="margin-top:12px">Building weekly report…</p></div>';

    // Defer so the spinner shows before heavy computation
    setTimeout(() => {
        const stats = aggregateWeeklyStats(workbookData);
        const notes = loadNotes();
        const refs  = loadRefs();
        grid.innerHTML = buildReportHTML(stats, notes, refs);
    }, 50);
}
