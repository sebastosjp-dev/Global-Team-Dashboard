/**
 * services.js — Data processing and statistics functions.
 * Uses columnFinder.js for DRY column key resolution.
 * @module services
 */
import { CONFIG } from './config.js';
import { parseCurrency, formatCurrency, normalizeCountry, isCountryMatch, sortCountriesByCount } from './utils.js';
import {
    findCountryKey, findKorTcvKey, findArrKey, findMrrKey,
    findContractStartKey, findContractEndKey, findStatusKey,
    findDealNameKey, findPocNameKey, findPocStartKey,
    findEstimatedValueKey, findWeightedValueKey,
    findKey, parseExcelDateSafe
} from './columnFinder.js';

/* ═══════════════════════════════════════════════════════════════
   ORDER SHEET
   ═══════════════════════════════════════════════════════════════ */

/**
 * Compute aggregated stats for the ORDER SHEET tab.
 * @param {Object[]} data
 * @param {string|null} filterCountry
 * @param {string} tabName
 * @param {Object} workbookData
 * @returns {Object}
 */
export function getOrderSheetStats(data, filterCountry, tabName, workbookData) {
    const orderData = tabName === 'ORDER SHEET' ? data : (workbookData['ORDER SHEET'] || []);
    if (orderData.length === 0) return {
        sumLocalTcv: 0, sumKorTcv: 0, sumArr: 0, sumMrr: 0, dealCount: 0,
        yearlyTcv: {}, qSums: { Q1: 0, Q2: 0, Q3: 0, Q4: 0 },
        qDeals: { Q1: [], Q2: [], Q3: [], Q4: [] }
    };

    const keys = Object.keys(orderData[0]);
    const korTcvKey = findKorTcvKey(keys);
    const arrKey = findArrKey(keys);
    const mrrKey = findMrrKey(keys);
    const startDateKey = findContractStartKey(keys);
    const dealNameKey = findDealNameKey(keys);
    const countryKey = findCountryKey(keys);

    let sumLocalTcv = 0, sumKorTcv = 0, sumArr = 0, sumMrr = 0, dealCount = 0;
    const yearlyTcv = {};
    const yearlyArr = {};
    const yearlyMrr = {};
    const qSums = { Q1: 0, Q2: 0, Q3: 0, Q4: 0 };
    const qDeals = { Q1: [], Q2: [], Q3: [], Q4: [] };
    const tcvByCountry = {};
    const currentYear = new Date().getFullYear();

    /**
     * Collect per-contract start years and ARR/MRR values first,
     * then build cumulative yearly totals in a second pass.
     * Premise: contracts renew each year unless explicitly terminated.
     */
    const contractEntries = [];

    orderData.filter(r => isCountryMatch(r, filterCountry)).forEach(row => {
        const lTcv = parseCurrency(row['Local TCV'] || row['Local TCV Amount']);
        const kTcv = korTcvKey ? parseCurrency(row[korTcvKey]) : 0;
        const arrVal = arrKey ? parseCurrency(row[arrKey]) : 0;
        const mrrVal = mrrKey ? parseCurrency(row[mrrKey]) : 0;

        sumLocalTcv += lTcv;
        sumKorTcv += kTcv;
        sumArr += arrVal;
        sumMrr += mrrVal;
        dealCount++;

        const countryName = countryKey ? (normalizeCountry(row[countryKey]) || 'Other') : 'Other';
        tcvByCountry[countryName] = (tcvByCountry[countryName] || 0) + kTcv;

        const d = parseExcelDateSafe(row[startDateKey]);
        if (d) {
            const startYear = d.getFullYear();
            if (!yearlyTcv[startYear]) yearlyTcv[startYear] = { local: 0, korea: 0 };
            yearlyTcv[startYear].local += lTcv;
            yearlyTcv[startYear].korea += kTcv;

            // Store for cumulative ARR/MRR calculation
            contractEntries.push({ startYear, arrVal, mrrVal });

            if (startYear === currentYear) {
                const qId = `Q${Math.floor(d.getMonth() / 3) + 1}`;
                qSums[qId] += kTcv;

                // Collect deal info for tooltip
                const name = dealNameKey ? String(row[dealNameKey] || 'N/A').trim() : 'N/A';
                qDeals[qId].push({ name, tcv: kTcv });
            }
        }
    });

    /* ── Build cumulative ARR / MRR ──
     * Each contract's ARR accumulates from its start year through
     * the current year (no termination flag = assumed renewal). */
    if (contractEntries.length > 0) {
        const minYear = contractEntries.reduce((m, e) => Math.min(m, e.startYear), Infinity);
        for (let y = minYear; y <= currentYear; y++) {
            yearlyArr[y] = 0;
            yearlyMrr[y] = 0;
        }
        contractEntries.forEach(({ startYear, arrVal, mrrVal }) => {
            for (let y = startYear; y <= currentYear; y++) {
                yearlyArr[y] += arrVal;
                yearlyMrr[y] += mrrVal;
            }
        });
        // Override headline sumArr/sumMrr with cumulative current-year total
        sumArr = yearlyArr[currentYear] || sumArr;
        sumMrr = yearlyMrr[currentYear] || sumMrr;
    }

    return {
        sumLocalTcv, sumKorTcv, sumArr, sumMrr, dealCount,
        yearlyTcv, yearlyArr, yearlyMrr,
        qSums, qDeals, tcvByCountry
    };
}

/* ═══════════════════════════════════════════════════════════════
   PIPELINE
   ═══════════════════════════════════════════════════════════════ */

/**
 * Compute aggregated stats for the PIPELINE tab.
 * @param {Object[]} pData
 * @param {Object[]} [orderData=[]]
 * @returns {Object}
 */
export function getPipelineStats(pData, orderData = []) {
    let pipelineByCountry = {};
    let pipelineByQuarter = {
        Q1: { countries: {}, deals: [] },
        Q2: { countries: {}, deals: [] },
        Q3: { countries: {}, deals: [] },
        Q4: { countries: {}, deals: [] }
    };

    // --- TCV FROM ORDER SHEET ---
    const orderTcvByQuarterCountry = { Q1: {}, Q2: {}, Q3: {}, Q4: {} };
    const currentYear = new Date().getFullYear();

    if (orderData.length > 0) {
        const oKeys = Object.keys(orderData[0]);
        const oCountryKey = findCountryKey(oKeys);
        const oTcvKey = findKorTcvKey(oKeys);
        const oStartKey = findContractStartKey(oKeys);

        orderData.forEach(row => {
            const country = normalizeCountry(row[oCountryKey]) || 'Other';
            const tcv = parseCurrency(row[oTcvKey]);
            const d = parseExcelDateSafe(row[oStartKey]);

            if (tcv > 0 && d && d.getFullYear() === currentYear) {
                const qId = `Q${Math.floor(d.getMonth() / 3) + 1}`;
                if (!pipelineByCountry[country]) pipelineByCountry[country] = { amount: 0, weighted: 0, tcv: 0, count: 0 };
                pipelineByCountry[country].tcv = (pipelineByCountry[country].tcv || 0) + tcv;
                pipelineByCountry[country].count++;

                if (!pipelineByQuarter[qId].countries[country]) pipelineByQuarter[qId].countries[country] = { amount: 0, weighted: 0, tcv: 0, count: 0 };
                pipelineByQuarter[qId].countries[country].tcv = (pipelineByQuarter[qId].countries[country].tcv || 0) + tcv;
                pipelineByQuarter[qId].countries[country].count++;
            }
        });
    }

    let pipelineByYearCountry = {};
    const pipelineInfluxData = Array(12).fill(0).map(() => ({ count: 0, amount: 0, weighted: 0, value: 0, accounts: [] }));
    const currentYearStr = new Date().getFullYear().toString();

    pData.forEach(r => {
        const keys = Object.keys(r);
        const c = normalizeCountry(r[findCountryKey(keys)]) || 'Other';

        const amtRaw = findKey(keys, k => (k.toUpperCase().includes('KOR TCV') && k.toUpperCase().includes('USD')) || k === 'Amount') || 'Amount';
        const wAmtRaw = findKey(keys, k => (k.toUpperCase().includes('WEIGHTED') && k.toUpperCase().includes('KOR TCV')) || k === 'Weighted Amount') || 'Weighted Amount';
        const nameKey = findKey(keys, k => k.toLowerCase().includes('deal name'), k => k.toLowerCase().includes('crm deal name'), k => k.toLowerCase().includes('customer'), k => k.toLowerCase().includes('end user'));

        const amt = parseCurrency(r[amtRaw] || r['Amount']);
        const wAmt = parseCurrency(r[wAmtRaw] || r['Weighted Amount']);
        const dealName = r[nameKey] || 'N/A';

        if (!pipelineByCountry[c]) pipelineByCountry[c] = { amount: 0, weighted: 0, tcv: 0, count: 0 };
        pipelineByCountry[c].amount += amt;
        pipelineByCountry[c].weighted += wAmt;
        pipelineByCountry[c].count++;

        const dKey = findKey(keys, k => k.toLowerCase().includes('date'), k => k.toLowerCase().includes('start'));
        const d = dKey ? parseExcelDateSafe(r[dKey]) : null;
        const year = d ? d.getFullYear().toString() : 'Unknown';

        const qKey = findKey(keys, k => k.toLowerCase() === 'quarter', k => k.toLowerCase().includes('qtr'), k => k.toLowerCase() === 'q');
        if (qKey && r[qKey]) {
            const qRaw = String(r[qKey]).toUpperCase().trim();
            let qMatch = '';
            if (qRaw.includes('Q1')) qMatch = 'Q1';
            else if (qRaw.includes('Q2')) qMatch = 'Q2';
            else if (qRaw.includes('Q3')) qMatch = 'Q3';
            else if (qRaw.includes('Q4')) qMatch = 'Q4';

            if (qMatch) {
                if (!pipelineByQuarter[qMatch].countries[c]) {
                    pipelineByQuarter[qMatch].countries[c] = { amount: 0, weighted: 0, tcv: 0, count: 0 };
                }
                pipelineByQuarter[qMatch].countries[c].amount += amt;
                pipelineByQuarter[qMatch].countries[c].weighted += wAmt;
                pipelineByQuarter[qMatch].countries[c].count++;
                pipelineByQuarter[qMatch].deals.push({ name: dealName, amount: amt, weighted: wAmt, country: c, year });
            }
        }

        if (year === currentYearStr && d) {
            const m = d.getMonth();
            pipelineInfluxData[m].count++;
            pipelineInfluxData[m].amount += amt;
            pipelineInfluxData[m].weighted += wAmt;
            pipelineInfluxData[m].value += amt;
            const nKey = findKey(keys, k => k.toLowerCase().includes('name'), k => k.toLowerCase().includes('customer'), k => k.toLowerCase().includes('end user'));
            if (nKey) pipelineInfluxData[m].accounts.push(String(r[nKey]).trim());
        }

        if (!pipelineByYearCountry[year]) pipelineByYearCountry[year] = {};
        if (!pipelineByYearCountry[year][c]) pipelineByYearCountry[year][c] = { amount: 0, weighted: 0 };
        pipelineByYearCountry[year][c].amount += amt;
        pipelineByYearCountry[year][c].weighted += wAmt;
    });

    return {
        pipelineByCountry,
        pipelineByQuarter,
        pipelineInfluxData,
        pipelineByYearCountry,
        sortedPipeline: Object.entries(pipelineByCountry).sort((a, b) => b[1].amount - a[1].amount),
        sortedQuarterly: Object.entries(pipelineByQuarter).sort((a, b) => a[0].localeCompare(b[0])),
        globalTotalAmount: Object.values(pipelineByCountry).reduce((acc, curr) => acc + curr.amount, 0),
        globalTotalWeighted: Object.values(pipelineByCountry).reduce((acc, curr) => acc + curr.weighted, 0),
        globalTotalTcv: Object.values(pipelineByCountry).reduce((acc, curr) => acc + (curr.tcv || 0), 0),
        globalTotalCount: Object.values(pipelineByCountry).reduce((acc, curr) => acc + (curr.count || 0), 0)
    };
}

/* ═══════════════════════════════════════════════════════════════
   PARTNER
   ═══════════════════════════════════════════════════════════════ */

/**
 * @param {Object[]} data
 * @param {string|null} filterCountry
 * @param {Object} workbookData
 * @returns {Object}
 */
export function getPartnerStats(data, filterCountry, workbookData) {
    const pKeys = Object.keys(data[0]);
    const pCountryKey = findCountryKey(pKeys);
    const pNameKey = findKey(pKeys, k => k.toLowerCase().includes('partner'), k => k.toLowerCase().includes('name')) || pKeys[0];

    const counts = {};
    const partnerGroups = {};
    data.forEach(r => {
        const country = normalizeCountry(r[pCountryKey]);
        if (country) {
            if (!partnerGroups[country]) partnerGroups[country] = [];
            partnerGroups[country].push(r);
            counts[country] = (counts[country] || 0) + 1;
        }
    });

    const pocDataForPartner = workbookData['POC'] || [];
    const currentPartnerNames = new Set(data.map(p => String(p[pNameKey] || '').trim().toLowerCase()));

    let pRankingData = {};
    pocDataForPartner.forEach(r => {
        const sKey = findStatusKey(Object.keys(r));
        const statusStr = sKey ? String(r[sKey]).trim().toLowerCase() : '';

        if (statusStr.includes('running') || statusStr.includes('progress') || statusStr.includes('ing')) {
            const pKey = findKey(Object.keys(r), k => k.toLowerCase() === 'partner');
            let pName = pKey ? String(r[pKey]).trim() : 'Unknown';
            if (!pName || pName.toLowerCase() === 'unknown' || pName === '') pName = 'Direct/Unknown';

            if (filterCountry && pName !== 'Direct/Unknown' && !currentPartnerNames.has(pName.toLowerCase())) return;

            const vKey = findEstimatedValueKey(Object.keys(r));
            const estValTotal = vKey ? parseCurrency(r[vKey]) : 0;

            if (!pRankingData[pName]) pRankingData[pName] = { count: 0, sumValue: 0 };
            pRankingData[pName].count += 1;
            pRankingData[pName].sumValue += estValTotal;
        }
    });

    return {
        counts,
        partnerGroups,
        sortedCountries: Object.keys(partnerGroups).sort((a, b) => partnerGroups[b].length - partnerGroups[a].length),
        sortedP: Object.entries(pRankingData)
            .map(([name, stats]) => ({ name, ...stats }))
            .sort((a, b) => b.count - a.count || b.sumValue - a.sumValue),
        pNameKey
    };
}

/* ═══════════════════════════════════════════════════════════════
   GENERIC COUNTRY STATS
   ═══════════════════════════════════════════════════════════════ */

/**
 * @param {Object[]} data
 * @param {string|null} filterCountry
 * @returns {Object|null}
 */
export function getGenericCountryStats(data, filterCountry) {
    if (!data || data.length === 0) return null;
    const keys = Object.keys(data[0]);
    const countryKey = findCountryKey(keys);
    if (!countryKey) return null;

    const countryCounts = {};
    const yearlyCounts = {};

    const getYear = (row) => {
        const dKey = findContractStartKey(keys) || findKey(keys, k => k.toUpperCase().includes('YEAR'));
        const d = parseExcelDateSafe(row[dKey]);
        if (d) return d.getFullYear().toString();
        const v = row[dKey];
        if (v) { const m = String(v).match(/(20\d{2})/); if (m) return m[1]; }
        return 'Unknown';
    };

    data.forEach(row => {
        const c = normalizeCountry(row[countryKey]);
        if (c) {
            countryCounts[c] = (countryCounts[c] || 0) + 1;
            const y = getYear(row);
            if (!yearlyCounts[y]) yearlyCounts[y] = {};
            yearlyCounts[y][c] = (yearlyCounts[y][c] || 0) + 1;
        }
    });

    return {
        sortedTotal: Object.entries(countryCounts).sort(sortCountriesByCount),
        yearlyCounts,
        sortedYears: Object.keys(yearlyCounts).sort((a, b) => (a === 'Unknown' ? 1 : b === 'Unknown' ? -1 : b - a))
    };
}

/* ═══════════════════════════════════════════════════════════════
   EXPIRING CONTRACTS
   ═══════════════════════════════════════════════════════════════ */

/**
 * @param {Object[]} data
 * @returns {Object[]|null}
 */
export function getExpiringContractsStats(data) {
    if (!data || data.length === 0) return null;
    const keys = Object.keys(data[0]);
    const endKey = findContractEndKey(keys);
    if (!endKey) return null;

    const dealNameKey = findDealNameKey(keys);
    const clientKey = findKey(keys, k => k.toUpperCase().includes('CLIENT'), k => k.toUpperCase().includes('CUSTOMER'), k => k.toUpperCase().includes('ACCOUNT'));
    const contractYrKey = findKey(keys, k => k.toUpperCase().replace(/\s/g, '').includes('CONTRACTYR'), k => k.toUpperCase().includes('YR'), k => k.toUpperCase().includes('YEAR'));

    const now = new Date();
    now.setHours(0, 0, 0, 0);
    const threeMonthsLater = new Date(now);
    threeMonthsLater.setMonth(now.getMonth() + 3);

    const expiringDeals = data.filter(row => {
        const expDate = parseExcelDateSafe(row[endKey]);
        return expDate && expDate >= now && expDate <= threeMonthsLater;
    }).sort((a, b) => (parseExcelDateSafe(a[endKey]) || 0) - (parseExcelDateSafe(b[endKey]) || 0));

    if (expiringDeals.length === 0) return null;

    return expiringDeals.map(deal => {
        const expDate = parseExcelDateSafe(deal[endKey]);
        let name = 'Unknown Deal';
        if (dealNameKey && deal[dealNameKey]) name = deal[dealNameKey];
        else if (clientKey && deal[clientKey]) name = deal[clientKey];
        return { name, date: expDate ? expDate.toISOString().split('T')[0] : 'N/A', year: contractYrKey ? deal[contractYrKey] : '' };
    });
}

/* ═══════════════════════════════════════════════════════════════
   PARTNER PERFORMANCE
   ═══════════════════════════════════════════════════════════════ */

/**
 * @param {Object[]} data
 * @returns {Object[]|null}
 */
export function getPartnerPerformanceStats(data) {
    if (!data || data.length === 0) return null;
    const keys = Object.keys(data[0]);
    const companyKey = findKey(keys, k => k.toLowerCase().replace(/\s/g, '') === 'companyname');
    const countryKey = findCountryKey(keys);
    const tcvKey = findKey(keys, k => k.toLowerCase().replace(/\s/g, '') === 'accumulatedtcv');

    if (!companyKey || !tcvKey) return null;

    const stats = data.map(row => ({
        name: String(row[companyKey] || 'Unknown').trim(),
        country: String(row[countryKey] || 'N/A').trim(),
        tcv: parseCurrency(row[tcvKey])
    })).filter(p => p.tcv > 0).sort((a, b) => b.tcv - a.tcv).slice(0, 10);

    return stats.length > 0 ? stats : null;
}

/* ═══════════════════════════════════════════════════════════════
   POC
   ═══════════════════════════════════════════════════════════════ */

/**
 * @param {Object[]} data
 * @param {Object} filters
 * @param {Object} workbookData
 * @returns {{ stats: Object, uniqueValues: Object }}
 */
export function getPocStats(data, filters, workbookData) {
    const sheet9Data = workbookData['Sheet9'] || [];
    const mergedData = data.map(pRow => {
        const nameKey = findPocNameKey(Object.keys(pRow));
        const pName = nameKey ? String(pRow[nameKey]).trim().toLowerCase() : '';
        const matchedS9 = sheet9Data.find(sRow => {
            const sNameKey = findPocNameKey(Object.keys(sRow));
            return sNameKey && String(sRow[sNameKey]).trim().toLowerCase() === pName;
        });
        return { ...pRow, _s9: matchedS9 || {} };
    });

    const uniqueValues = { countries: new Set(['All']), industries: new Set(['All']), partners: new Set(['All']) };
    mergedData.forEach(r => {
        const cKey = findCountryKey(Object.keys(r));
        const iKey = findKey(Object.keys(r), k => k.toLowerCase().includes('industry'));
        const pKey = findKey(Object.keys(r), k => k.toLowerCase() === 'partner');
        if (cKey && r[cKey]) uniqueValues.countries.add(normalizeCountry(r[cKey]));
        if (iKey && r[iKey]) uniqueValues.industries.add(String(r[iKey]).trim());
        if (pKey && r[pKey]) uniqueValues.partners.add(String(r[pKey]).trim());
    });

    const displayData = mergedData.filter(r => {
        const cKey = findCountryKey(Object.keys(r));
        const rCountry = cKey ? normalizeCountry(r[cKey]) : '';
        const rIndustry = String(r[findKey(Object.keys(r), k => k.toLowerCase().includes('industry'))] || '').trim();
        const rPartner = String(r[findKey(Object.keys(r), k => k.toLowerCase() === 'partner')] || '').trim();
        return (filters.country === 'All' || rCountry === filters.country) &&
            (filters.industry === 'All' || rIndustry === filters.industry) &&
            (filters.partner === 'All' || rPartner === filters.partner);
    });

    let s = {
        longTermCount: 0, midTermCount: 0, normalCount: 0,
        runningList: [], staledRunningList: [], runningNames: [],
        partnerRunDays: {},
        statusStats: { won: 0, drop: 0, running: 0, hold: 0, others: 0 },
        totalWonDays: 0, wonCountForStats: 0,
        totalHold: 0, holdNames: [],
        influxData: Array(12).fill(0).map(() => ({ count: 0, estimated: 0, weighted: 0, accounts: [] })),
        industryVal: {}, partnerRankingData: {},
        monthlyActive: Array(12).fill(0).map(() => ({ count: 0, tcv: 0 }))
    };

    displayData.forEach(r => {
        const allKeys = Object.keys(r);
        const s9Keys = Object.keys(r._s9 || {});

        const wdKey = findKey(allKeys, k => k.toLowerCase().includes('working days'), k => k.toLowerCase().includes('workingdays'), k => k.toLowerCase().includes('running days'));
        const s9WdKey = findKey(s9Keys, k => k.toLowerCase().includes('working days'), k => k.toLowerCase().includes('workingdays'), k => k.toLowerCase().includes('running days'));
        let runningDays = Number(r[wdKey] || (r._s9 && r._s9[s9WdKey]) || 0);

        if (runningDays >= 100) s.longTermCount++; else if (runningDays >= 60) s.midTermCount++; else s.normalCount++;

        const statusKey = findStatusKey(allKeys);
        const curStatus = String(r[statusKey] || '').trim().toLowerCase();

        // New Won Criteria
        const pnKey = findKey(allKeys, k => k.toLowerCase().replace(/\s/g, '') === 'pocnotes');
        const peKey = findKey(allKeys, k => k.toLowerCase().replace(/\s/g, '') === 'pocend');
        const notesStr = String(r[pnKey] || '').trim().toLowerCase();
        const pEndDate = parseExcelDateSafe(r[peKey]);
        const analysisYear = new Date().getFullYear();
        const now = new Date();
        const isActuallyWon = notesStr.includes('won') && pEndDate && pEndDate.getFullYear() === analysisYear && pEndDate <= now;

        if (isActuallyWon) {
            s.statusStats.won++;
        } else if (curStatus.includes('hold') || curStatus.includes('pause')) {
            s.totalHold++;
            const name = r[findPocNameKey(allKeys)] || 'Unknown';
            s.holdNames.push(name);
            s.statusStats.hold++;
        } else if (curStatus.includes('running') || curStatus.includes('progress') || curStatus.includes('ing')) {
            s.statusStats.running++;
            const name = r[findPocNameKey(allKeys)] || 'Unknown';
            s.runningNames.push(name);
        } else if (curStatus.includes('drop') || curStatus.includes('cancel') || curStatus.includes('fail')) {
            s.statusStats.drop++;
        } else if (curStatus !== '') {
            s.statusStats.others++;
        }

        const isWon = curStatus.includes('won') || curStatus.includes('complete') || isActuallyWon;

        if (!isWon && (curStatus.includes('running') || curStatus.includes('progress') || runningDays >= 100)) {
            let statusLabel = 'Running';
            let statusColor = '#007AFF';
            if (curStatus.includes('hold') || curStatus.includes('pause')) { statusLabel = 'Hold'; statusColor = '#FF9500'; }
            else if (curStatus.includes('drop') || curStatus.includes('cancel') || curStatus.includes('fail')) { statusLabel = 'Drop'; statusColor = '#FF3B30'; }
            else if (curStatus.includes('won') || curStatus.includes('complete')) { statusLabel = 'Won'; statusColor = '#34C759'; }
            else if (curStatus === '') { statusLabel = 'Others'; statusColor = '#9CA3AF'; }

            const entry = {
                name: r[findPocNameKey(allKeys)] || 'Unknown',
                partner: r[findKey(allKeys, k => k.toLowerCase() === 'partner')] || 'Unknown',
                country: r[findCountryKey(allKeys)] || 'Unknown',
                industry: r[findKey(allKeys, k => k.toLowerCase().includes('industry'))] || 'Unknown',
                days: runningDays,
                status: statusLabel,
                statusColor,
                notes: r._s9 && r._s9[findKey(s9Keys, k => k.toLowerCase().includes('notes'))] || '',
                techComm: r[findKey(allKeys, k => k.toLowerCase().includes('technical comment'), k => k.toLowerCase().includes('report'))] || '',
                isStalled: runningDays >= 100
            };
            s.runningList.push(entry);
            if (runningDays >= 100) s.staledRunningList.push(entry);
        }

        const startValKey = findPocStartKey(allKeys);
        const startVal = r[startValKey];
        const estValKey = findEstimatedValueKey(allKeys);
        const wValKey = findWeightedValueKey(allKeys);
        const estVal = parseCurrency(r[estValKey] || 0);
        const wVal = parseCurrency(r[wValKey] || 0);

        if (startVal) {
            const d = parseExcelDateSafe(startVal);
            const pocAnalysisYear = new Date().getFullYear();
            if (d && d.getFullYear() === pocAnalysisYear) {
                const m = d.getMonth();
                s.influxData[m].count++;
                s.influxData[m].estimated += estVal;
                s.influxData[m].weighted += wVal;
                const pocNameKey = findPocNameKey(allKeys);
                if (pocNameKey && r[pocNameKey]) s.influxData[m].accounts.push(String(r[pocNameKey]).trim());
            }

            if (d) {
                let pocEndDate = new Date(8640000000000000);
                if (curStatus.includes('won') || curStatus.includes('drop') || curStatus.includes('fail') || curStatus.includes('complete')) {
                    pocEndDate = new Date(d.getTime() + (runningDays * 86400 * 1000));
                }
                for (let m = 0; m < 12; m++) {
                    const monthStart = new Date(pocAnalysisYear, m, 1);
                    const monthEnd = new Date(pocAnalysisYear, m + 1, 0);
                    if (d <= monthEnd && pocEndDate >= monthStart) {
                        s.monthlyActive[m].tcv += estVal;
                        s.monthlyActive[m].count++;
                    }
                }
            }
        }

        const pName = String(r[findKey(allKeys, k => k.toLowerCase() === 'partner')] || 'Unknown').trim();
        if (pName && pName !== 'Unknown' && pName !== '') {
            if (!s.partnerRunDays[pName]) s.partnerRunDays[pName] = { sum: 0, count: 0 };
            s.partnerRunDays[pName].sum += runningDays;
            s.partnerRunDays[pName].count++;
        }

        if (isActuallyWon) {
            const wd = parseCurrency(r[findKey(allKeys, k => k.toLowerCase().includes('working days'))] || 0);
            if (wd > 0) { s.totalWonDays += wd; s.wonCountForStats++; }
        }

        if (curStatus.includes('running') || curStatus.includes('progress') || curStatus.includes('ing')) {
            const rawInd = String(r[findKey(allKeys, k => k.toLowerCase().includes('industry'))] || 'Other').trim();
            const ind = rawInd.split('|').pop().replace(/[ㄱ-ㅎㅏ-ㅣ가-힣]/g, '').trim() || 'Other';
            s.industryVal[ind] = (s.industryVal[ind] || 0) + estVal;
            if (!s.partnerRankingData[pName]) s.partnerRankingData[pName] = { count: 0, sumValue: 0 };
            s.partnerRankingData[pName].count++;
            s.partnerRankingData[pName].sumValue += estVal;
        }
    });

    s.runningList.sort((a, b) => b.days - a.days);
    s.partnerAvg = Object.keys(s.partnerRunDays).map(p => ({ partner: p, avg: Math.round(s.partnerRunDays[p].sum / s.partnerRunDays[p].count) })).sort((a, b) => b.avg - a.avg);
    s.sortedIndustry = Object.entries(s.industryVal).map(([name, val]) => ({ name, val })).sort((a, b) => b.val - a.val);
    s.sortedPartnersRanking = Object.entries(s.partnerRankingData).map(([name, st]) => ({ name, ...st })).sort((a, b) => b.count - a.count || b.sumValue - a.sumValue);

    return { stats: s, uniqueValues };
}

/* ═══════════════════════════════════════════════════════════════
   EVENT
   ═══════════════════════════════════════════════════════════════ */

/**
 * @param {Object[]} eventData
 * @param {string|null} filterCountry
 * @returns {Object|null}
 */
export function getEventStats(eventData, filterCountry) {
    const filteredEvents = eventData.filter(row => {
        // 국가 필터가 적용된 경우에만 필터링하고, 빈 행은 제외합니다.
        if (filterCountry && !isCountryMatch(row, filterCountry)) return false;
        if (Object.values(row).every(v => v === null || v === "")) return false;
        return true;
    });

    if (filteredEvents.length === 0) return null;

    const sampleKeys = Object.keys(filteredEvents[0]);
    const sKey = findKey(sampleKeys, k => k.toLowerCase().includes('spending') && k.toLowerCase().includes('usd'), k => k.toLowerCase().includes('spending')) || sampleKeys.find(k => k.toLowerCase().includes('usd'));
    const pKey = findKey(sampleKeys, k => k.toLowerCase().includes('poc') && k.toLowerCase().includes('generated'));
    const dKey = findKey(sampleKeys, k => k.toLowerCase().includes('deal') && (k.toLowerCase().includes('converted') || k.toLowerCase().includes('closed')));

    const dateKey = findKey(sampleKeys, k => k.toLowerCase().includes('date')) || sampleKeys[0];
    const nKey = findKey(sampleKeys, k => k.toLowerCase() === 'event name', k => k.toLowerCase().includes('name')) || findKey(sampleKeys, k => k.toLowerCase().includes('event')) || sampleKeys[1];

    let totalSpending = 0, totalPOC = 0, totalDeals = 0;
    const comparisonData = filteredEvents.reduce((acc, row) => {
        const spend = sKey ? parseCurrency(row[sKey]) : 0;
        if (spend <= 0) return acc;

        const pocs = pKey ? (parseInt(row[pKey]) || 0) : 0;
        const deals = dKey ? (parseInt(row[dKey]) || 0) : 0;

        let eventName = nKey ? String(row[nKey] || 'Unnamed').trim() : 'Unnamed';
        let dateVal = row[dateKey];
        let dateObj = parseExcelDateSafe(dateVal);
        let dateStr = dateObj ? dateObj.toISOString().split('T')[0].substring(0, 7) : String(dateVal || 'Unknown Date');

        totalSpending += spend;
        totalPOC += pocs;
        totalDeals += deals;

        acc.push({ name: dateStr, eventName, spend, pocs, deals });
        return acc;
    }, []);

    return {
        totalSpending, totalPOC, totalDeals,
        eventCount: filteredEvents.length,
        comparisonData,
        costPerPOC: totalPOC > 0 ? (totalSpending / totalPOC) : 0,
        costPerDeal: totalDeals > 0 ? (totalSpending / totalDeals) : 0
    };
}

/* ═══════════════════════════════════════════════════════════════
   COUNTRY SPECIFIC
   ═══════════════════════════════════════════════════════════════ */

/**
 * @param {Object[]} data
 * @returns {Object}
 */
export function getCountrySpecificStats(data) {
    const sortedYears = ['2026', '2025', '2024', '2023'];
    const summary = {};
    sortedYears.forEach(y => summary[y] = { kTcv: 0, lArr: 0 });

    data.forEach(row => {
        const dKey = findContractStartKey(Object.keys(row)) || findKey(Object.keys(row), k => k.toLowerCase().includes('date'));
        const d = dKey ? parseExcelDateSafe(row[dKey]) : null;
        if (d) {
            const y = d.getFullYear().toString();
            if (summary[y]) {
                summary[y].kTcv += parseCurrency(row['KOR TCV(USD)'] || row['KOR TCV']);
                summary[y].lArr += parseCurrency(row['Local ARR'] || row['ARR']);
            }
        }
    });
    return { summary, sortedYears };
}

/* ═══════════════════════════════════════════════════════════════
   SERVICE ANALYSIS
   ═══════════════════════════════════════════════════════════════ */

/**
 * Parse End License Date from CSM row.
 * @param {Object} row
 * @returns {Date|null}
 */
function _parseLicenseDate(row, key) {
    const val = row[key];
    if (!val) return null;
    if (val instanceof Date) return val;
    if (typeof val === 'number') return new Date(Math.round((val - 25569) * 86400 * 1000));
    return parseExcelDateSafe(val);
}

/**
 * Compute stats for END USER (CSM) dashboard.
 * Includes total customer count, expiring licenses, and service analysis.
 * @param {Object[]} data - All CSM rows
 * @param {string|null} filterCountry
 * @returns {Object|null}
 */
export function getServiceAnalysisStats(data, filterCountry = null) {
    const countryFiltered = data.filter(r => isCountryMatch(r, filterCountry));

    /* ── Total end user count (rows with No. column) ── */
    const totalEndUsers = countryFiltered.filter(r => {
        const no = r['No.'] ?? r['No'] ?? r['NO.'] ?? r['NO'];
        return no !== '' && no !== null && no !== undefined;
    }).length;

    /* ── License expiration analysis (6 months) ── */
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    const sixMonthsLater = new Date(today);
    sixMonthsLater.setMonth(today.getMonth() + 6);

    const expiringCustomers = [];
    countryFiltered.forEach(row => {
        const endDate = _parseLicenseDate(row, 'End License Date');
        if (!endDate) return;
        if (endDate >= today && endDate <= sixMonthsLater) {
            const diffDays = Math.ceil((endDate - today) / (1000 * 60 * 60 * 24));
            let urgency = 'normal';
            if (diffDays <= 30) urgency = 'critical';
            else if (diffDays <= 90) urgency = 'warning';

            expiringCustomers.push({
                name: row['End User'] || 'Unknown',
                country: row['Country'] || 'N/A',
                status: row['Status'] || 'N/A',
                endDate,
                endDateStr: endDate.toISOString().split('T')[0],
                startDateStr: (() => {
                    const sd = _parseLicenseDate(row, 'Active License Date');
                    return sd ? sd.toISOString().split('T')[0] : 'N/A';
                })(),
                dDay: diffDays === 0 ? 'D-Day' : `D-${diffDays}`,
                diffDays,
                urgency,
                tcv: parseCurrency(row['TCV Amount']),
                arr: parseCurrency(row['ARR Amount']),
                services: row['Services'] || 'N/A'
            });
        }
    });
    expiringCustomers.sort((a, b) => a.diffDays - b.diffDays);

    /* ── Service analysis (Active only) ── */
    const activeData = countryFiltered.filter(r => r['Status'] === 'Active' && r['Services']);
    if (activeData.length === 0 && totalEndUsers === 0) return null;

    const comboCounts = {};
    const upscaleTargetsArr = [];
    activeData.forEach(row => {
        const raw = String(row['Services'] || '').trim();
        if (!raw) return;
        const services = raw.split(',').map(s => s.trim()).filter(Boolean);
        const combo = [...services].sort().join(' + ');
        comboCounts[combo] = (comboCounts[combo] || 0) + 1;
        if (services.length === 1) {
            upscaleTargetsArr.push({ name: row['End User'], country: row['Country'], service: raw, tcv: parseCurrency(row['TCV Amount']) });
        }
    });

    /* ── Count by urgency ── */
    const criticalCount = expiringCustomers.filter(c => c.urgency === 'critical').length;
    const warningCount = expiringCustomers.filter(c => c.urgency === 'warning').length;
    const normalCount = expiringCustomers.filter(c => c.urgency === 'normal').length;
    const activeCount = countryFiltered.filter(r => r['Status'] === 'Active').length;
    const inactiveCount = countryFiltered.filter(r => r['Status'] === 'Inactive').length;

    return {
        totalEndUsers,
        activeCount,
        inactiveCount,
        totalCustomers: activeData.length,
        singleServiceCustomers: upscaleTargetsArr.length,
        multiServiceCustomers: activeData.length - upscaleTargetsArr.length,
        sortedCombos: Object.entries(comboCounts).sort((a, b) => b[1] - a[1]),
        upsellTargets: upscaleTargetsArr.sort((a, b) => b.tcv - a.tcv),
        palette: CONFIG.COLORS,
        expiringCustomers,
        expiringCount: expiringCustomers.length,
        criticalCount,
        warningCount,
        normalExpCount: normalCount
    };
}

/* ═══════════════════════════════════════════════════════════════
   COLLECTION
   ═══════════════════════════════════════════════════════════════ */

/**
 * Compute stats for the COLLECTION dashboard.
 * @param {Object[]} data
 * @returns {Object}
 */
export function getCollectionStats(data) {
    const currentYear = new Date().getFullYear();
    let globalTotalTcv = 0;
    let globalTotalReceived = 0;
    const yearlyStats = {}; // { 2024: { target: 0, actual: 0 }, ... }
    const yearlyDistributorTargets = {}; // { 2024: { 'Distributor A': 100, ... }, ... }
    const unpaidList = [];
    const distributorEndYears = {};

    // Initialize data structure for 2024 to 2030
    for (let y = 2024; y <= 2030; y++) {
        yearlyStats[y] = { target: 0, actual: 0 };
    }

    data.forEach(row => {
        const keys = Object.keys(row);
        const tcv = parseCurrency(row[findKorTcvKey(keys)] || row['KOR TCV']);
        globalTotalTcv += tcv;

        // Accumulate all received columns dynamically (2023, 2024, 2025...)
        keys.forEach(key => {
            if (key.toUpperCase().includes('RECEIVED')) {
                globalTotalReceived += parseCurrency(row[key]);
            }
        });

        const distributor = row['Distributor'] || row['Partner'] || 'Unknown';
        const endUser = row['End User'] || 'N/A';
        const contractStart = parseExcelDateSafe(row[findContractStartKey(keys)]);
        const contractEnd = parseExcelDateSafe(row[findContractEndKey(keys)] || row['Contract End']);

        // Calculate Monthly Amortized TCV (MRR)
        let monthlyAmortized = 0;
        if (contractStart && contractEnd && tcv > 0) {
            const totalMonths = ((contractEnd.getFullYear() - contractStart.getFullYear()) * 12) + (contractEnd.getMonth() - contractStart.getMonth()) + 1;
            monthlyAmortized = totalMonths > 0 ? tcv / totalMonths : 0;
        }

        let lastArrYear = null;

        for (let y = 2024; y <= 2030; y++) {
            // Calculate Target for this specific year based on active months within the year
            let targetForYear = 0;
            if (monthlyAmortized > 0 && y >= contractStart.getFullYear() && y <= contractEnd.getFullYear()) {
                const mStart = (y === contractStart.getFullYear()) ? contractStart.getMonth() : 0;
                const mEnd = (y === contractEnd.getFullYear()) ? contractEnd.getMonth() : 11;
                const activeMonths = mEnd - mStart + 1;
                targetForYear = monthlyAmortized * activeMonths;
            }

            const recVal = parseCurrency(row[`${y} RECEIVED`]);

            yearlyStats[y].target += targetForYear;
            yearlyStats[y].actual += recVal;

            if (targetForYear > 0) {
                if (!yearlyDistributorTargets[y]) yearlyDistributorTargets[y] = {};
                yearlyDistributorTargets[y][distributor] = (yearlyDistributorTargets[y][distributor] || 0) + targetForYear;

                lastArrYear = y;

                // Unpaid logic: current year or earlier, and target > 0 but RECEIVED is 0 or less
                if (y <= currentYear && recVal <= 0) {
                    unpaidList.push({
                        distributor,
                        endUser,
                        year: y,
                        unpaidAmount: targetForYear,
                        contractEnd: contractEnd,
                        contractEndDateStr: contractEnd ? contractEnd.toISOString().split('T')[0] : 'N/A'
                    });
                }
            }
        }

        if (lastArrYear) {
            if (!distributorEndYears[distributor] || lastArrYear > distributorEndYears[distributor]) {
                distributorEndYears[distributor] = lastArrYear;
            }
        }
    });

    // Calculate dunning priority: Unpaid amount DESC, Contract End ASC
    unpaidList.sort((a, b) => {
        if (b.unpaidAmount !== a.unpaidAmount) return b.unpaidAmount - a.unpaidAmount;
        if (!a.contractEnd) return 1;
        if (!b.contractEnd) return -1;
        return a.contractEnd - b.contractEnd;
    });

    return {
        globalTotalTcv,
        globalTotalReceived,
        yearlyStats,
        yearlyDistributorTargets,
        unpaidList,
        distributorEndYears,
        currentYearTarget: yearlyStats[currentYear].target,
        currentYearActual: yearlyStats[currentYear].actual
    };
}

/**
 * Detailed collection analysis for the new UI table.
 * @param {Object[]} data
 * @returns {Object}
 */
export function getDetailedCollectionAnalysis(data) {
    const keys = data.length > 0 ? Object.keys(data[0]) : [];
    const korTcvKey = findKorTcvKey(keys);
    const startKey = findContractStartKey(keys);
    const yrKey = findKey(keys, k => k.toUpperCase().includes('CONTRACT YR'), k => k.toUpperCase() === 'YR');

    const resultRows = data.map(row => {
        const distributor = row['Distributor'] || row['Partner'] || 'Unknown';
        const endUser = row['End User'] || 'N/A';
        const contractStart = parseExcelDateSafe(row[startKey]);
        const contractYrRaw = row[yrKey];
        const korTcv = parseCurrency(row[korTcvKey]);

        let startYear = contractStart ? contractStart.getFullYear() : null;
        let numYears = 0;
        let isPerpetual = false;

        if (String(contractYrRaw).toLowerCase().includes('perpetual')) {
            isPerpetual = true;
            numYears = 1; // Or handle as ongoing
        } else {
            numYears = parseInt(contractYrRaw) || 0;
        }

        const cy = new Date().getFullYear();
        const yearsToTrack = [cy - 2, cy - 1, cy];
        const status = {};
        let totalReceived = 0;

        // Sum 2023-2028 RECEIVED
        for (let y = 2023; y <= 2028; y++) {
            let recVal = parseCurrency(row[`${y} RECEIVED`]);

            // For Perpetual, we assume 100% collection in the start year
            if (isPerpetual && y === startYear) {
                recVal = korTcv;
            }

            totalReceived += recVal;
            if (yearsToTrack.includes(y)) {
                // Check if Y is within contract period
                let inPeriod = false;
                if (startYear && !isNaN(numYears)) {
                    if (isPerpetual) {
                        inPeriod = (y === startYear); // Only the start year is a collection year for Perpetual
                    } else {
                        inPeriod = (y >= startYear && y < (startYear + numYears));
                    }
                }

                if (inPeriod) {
                    let displayRecVal = recVal;
                    // Double check display value for Perpetual start year
                    if (isPerpetual && y === startYear) {
                        displayRecVal = korTcv;
                    }
                    const formattedVal = `$${formatCurrency(displayRecVal)}`;
                    status[y] = displayRecVal > 0 ? `${formattedVal} ✅` : `${formattedVal} ❌`;
                } else {
                    status[y] = '-';
                }
            }
        }

        const balance = isPerpetual ? 0 : Math.round((korTcv - totalReceived) * 100) / 100;

        const statusKeys = {};
        yearsToTrack.forEach(y => { statusKeys[`status${y}`] = status[y]; });

        return {
            contractStart,
            contractStartDateStr: contractStart ? contractStart.toISOString().split('T')[0] : 'N/A',
            distributor,
            endUser,
            contractYr: contractYrRaw,
            ...statusKeys,
            totalTcv: korTcv,
            balance: balance
        };
    });

    // Sort by Contract Start ASC
    resultRows.sort((a, b) => {
        if (!a.contractStart) return 1;
        if (!b.contractStart) return -1;
        return a.contractStart - b.contractStart;
    });

    // Summary by Distributor
    const distributorSummary = {};
    resultRows.forEach(r => {
        if (!distributorSummary[r.distributor]) distributorSummary[r.distributor] = 0;
        distributorSummary[r.distributor] += r.balance;
    });

    const sortedSummary = Object.entries(distributorSummary)
        .map(([name, balance]) => ({ name, balance }))
        .sort((a, b) => b.balance - a.balance);

    return {
        rows: resultRows,
        summary: sortedSummary
    };
}

/* ═══════════════════════════════════════════════════════════════
   TCV vs ARR (Revenue Mix Analysis)
   ═══════════════════════════════════════════════════════════════ */

/**
 * Compute aggregated TCV vs ARR stats grouped by End User.
 * Rows with KOR TCV <= 0 or empty are excluded.
 * Empty KOR ARR is treated as 0 (Perpetual License).
 * @param {Object[]} data - ORDER SHEET rows
 * @param {{ country: string, contractYr: string }} filters
 * @returns {Object}
 */
export function getTcvArrStats(data, filters = {}) {
    if (!data || data.length === 0) return null;

    const keys = Object.keys(data[0]);
    const korTcvKey = findKorTcvKey(keys);
    const arrKey = findArrKey(keys);
    const countryKey = findCountryKey(keys);
    const endUserKey = findKey(keys,
        k => k.toUpperCase().replace(/\s/g, '') === 'ENDUSER',
        k => k.toUpperCase().includes('END USER'),
        k => k.toUpperCase().includes('CUSTOMER')
    );
    const contractYrKey = findKey(keys,
        k => k.toUpperCase().replace(/\s/g, '').includes('CONTRACTYR'),
        k => k.toUpperCase().includes('YR'),
        k => k.toUpperCase().includes('YEAR')
    );

    if (!endUserKey || !korTcvKey) return null;

    const uniqueCountries = new Set();
    const uniqueYears = new Set();

    /* ── Phase 1: Collect unique filter values ── */
    data.forEach(row => {
        const country = countryKey ? normalizeCountry(row[countryKey]) : null;
        if (country) uniqueCountries.add(country);
        const yr = contractYrKey ? String(row[contractYrKey]).trim() : '';
        if (yr && yr !== '' && yr !== 'undefined') uniqueYears.add(yr);
    });

    /* ── Phase 2: Filter + Aggregate by End User ── */
    const aggregated = {};
    let totalTcv = 0;
    let totalArr = 0;
    let totalRecurringTcv = 0;
    let totalPureGap = 0;
    let perpetualCount = 0;
    let recurringCount = 0;

    data.forEach(row => {
        const tcv = korTcvKey ? parseCurrency(row[korTcvKey]) : 0;
        if (tcv <= 0) return;

        const country = countryKey ? normalizeCountry(row[countryKey]) : null;
        if (filters.country && filters.country !== 'All' && country !== filters.country) return;

        const yr = contractYrKey ? String(row[contractYrKey]).trim() : '';
        if (filters.contractYr && filters.contractYr !== 'All' && yr !== filters.contractYr) return;

        const endUser = String(row[endUserKey] || 'Unknown').trim();
        const arr = arrKey ? parseCurrency(row[arrKey]) : 0;

        let numYears = parseInt(yr) || 1;
        if (yr.toLowerCase().includes('perpetual')) {
            numYears = 1;
        }
        if (numYears <= 0) numYears = 1;

        const expectedTcvFromArr = arr * numYears;
        const dealRecurringTcv = Math.min(tcv, expectedTcvFromArr);
        const dealPureGap = Math.max(0, tcv - dealRecurringTcv);

        if (!aggregated[endUser]) {
            aggregated[endUser] = { tcv: 0, arr: 0, recurringTcv: 0, pureGap: 0, country: country || 'N/A', deals: 0 };
        }
        aggregated[endUser].tcv += tcv;
        aggregated[endUser].arr += arr;
        aggregated[endUser].recurringTcv += dealRecurringTcv;
        aggregated[endUser].pureGap += dealPureGap;
        aggregated[endUser].deals += 1;

        totalTcv += tcv;
        totalArr += arr;
        totalRecurringTcv += dealRecurringTcv;
        totalPureGap += dealPureGap;
    });

    /* ── Phase 3: Build sorted list & classify ── */
    const sorted = Object.entries(aggregated)
        .map(([name, vals]) => {
            const gap = vals.pureGap;
            const gapPct = vals.tcv > 0 ? (gap / vals.tcv) * 100 : 0;
            const recurringPct = vals.tcv > 0 ? (vals.recurringTcv / vals.tcv) * 100 : 0;
            const isPerpetual = vals.arr === 0;
            if (isPerpetual) perpetualCount++;
            else recurringCount++;
            return { name, ...vals, gap, gapPct, recurringPct, isPerpetual };
        })
        .sort((a, b) => b.tcv - a.tcv);

    return {
        items: sorted,
        totalTcv,
        totalArr,
        totalRecurringTcv,
        totalGap: totalPureGap,
        accountCount: sorted.length,
        perpetualCount,
        recurringCount,
        uniqueCountries: ['All', ...Array.from(uniqueCountries).sort()],
        uniqueYears: ['All', ...Array.from(uniqueYears).sort()]
    };
}
