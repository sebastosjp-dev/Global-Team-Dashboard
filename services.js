/**
 * services.js - Data processing and statistics functions
 */
import { CONFIG } from './config.js';
import { parseCurrency, formatCurrency, normalizeCountry, isCountryMatch, sortCountriesByCount } from './utils.js';

export function getOrderSheetStats(data, filterCountry, tabName, workbookData) {
    const orderData = tabName === 'ORDER SHEET' ? data : (workbookData['ORDER SHEET'] || []);
    const keys = orderData.length > 0 ? Object.keys(orderData[0]) : [];
    const korTcvKey = keys.find(k => k.toUpperCase().replace(/\s/g, '') === 'KORTCV') || keys.find(k => k.toUpperCase().includes('TCV') && (k.toUpperCase().includes('KOR') || k.toUpperCase().includes('KRW')));
    const arrKey = keys.find(k => k.toUpperCase().replace(/\s/g, '') === 'KORARR') || keys.find(k => k.toUpperCase().includes('ARR') && (k.toUpperCase().includes('KOR') || k.toUpperCase().includes('KRW')));
    const mrrKey = keys.find(k => k.toUpperCase().replace(/\s/g, '') === 'KORMRR') || keys.find(k => k.toUpperCase().includes('MRR') && (k.toUpperCase().includes('KOR') || k.toUpperCase().includes('KRW')));
    const startDateKey = keys.find(k => k.toUpperCase().replace(/\s/g, '').includes('CONTRACTSTART')) || keys.find(k => k.toUpperCase().includes('START') && k.toUpperCase().includes('DATE'));

    let sumLocalTcv = 0, sumKorTcv = 0, sumArr = 0, sumMrr = 0, dealCount = 0;
    const yearlyTcv = {};
    const qSums = { 'Q1': 0, 'Q2': 0, 'Q3': 0, 'Q4': 0 };

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

        const cStart = row[startDateKey];
        if (cStart) {
            let d = (cStart instanceof Date) ? cStart : (typeof cStart === 'number' && cStart > 40000) ? new Date(Math.round((cStart - 25569) * 86400 * 1000)) : new Date(cStart);
            if (d && !isNaN(d.getTime())) {
                const y = d.getFullYear().toString();
                if (!yearlyTcv[y]) yearlyTcv[y] = { local: 0, korea: 0 };
                yearlyTcv[y].local += lTcv;
                yearlyTcv[y].korea += kTcv;
                if (d.getFullYear() === 2026) {
                    qSums[`Q${Math.floor(d.getMonth() / 3) + 1}`] += kTcv;
                }
            }
        }
    });

    return { sumLocalTcv, sumKorTcv, sumArr, sumMrr, dealCount, yearlyTcv, qSums };
}

export function getPipelineStats(pData) {
    let pipelineByCountry = {};
    let pipelineByQuarter = { 
        'Q1': { countries: {}, deals: [] }, 
        'Q2': { countries: {}, deals: [] }, 
        'Q3': { countries: {}, deals: [] }, 
        'Q4': { countries: {}, deals: [] } 
    };
    let pipelineByYearCountry = {};
    const pipelineInfluxData = Array(12).fill(0).map(() => ({ count: 0, amount: 0, weighted: 0, value: 0, accounts: [] }));

    pData.forEach(r => {
        const keys = Object.keys(r);
        const c = normalizeCountry(r[Object.keys(r).find(k => k.toLowerCase().includes('country'))]) || 'Other';

        const amtRaw = keys.find(k => (k.toUpperCase().includes('KOR TCV') && k.toUpperCase().includes('USD')) || k === 'Amount') || 'Amount';
        const wAmtRaw = keys.find(k => (k.toUpperCase().includes('WEIGHTED') && k.toUpperCase().includes('KOR TCV')) || k === 'Weighted Amount') || 'Weighted Amount';
        const nameKey = keys.find(k => k.toLowerCase().includes('deal name') || k.toLowerCase().includes('crm deal name') || k.toLowerCase().includes('customer') || k.toLowerCase().includes('end user'));

        const amt = parseCurrency(r[amtRaw] || r['Amount']);
        const wAmt = parseCurrency(r[wAmtRaw] || r['Weighted Amount']);
        const dealName = r[nameKey] || 'N/A';

        if (!pipelineByCountry[c]) pipelineByCountry[c] = { amount: 0, weighted: 0 };
        pipelineByCountry[c].amount += amt;
        pipelineByCountry[c].weighted += wAmt;

        const qKey = keys.find(k => k.toLowerCase() === 'quarter' || k.toLowerCase().includes('qtr') || k.toLowerCase() === 'q');
        if (qKey && r[qKey]) {
            const qRaw = String(r[qKey]).toUpperCase().trim();
            let qMatch = '';
            if (qRaw.includes('Q1')) qMatch = 'Q1';
            else if (qRaw.includes('Q2')) qMatch = 'Q2';
            else if (qRaw.includes('Q3')) qMatch = 'Q3';
            else if (qRaw.includes('Q4')) qMatch = 'Q4';

            if (qMatch) {
                if (!pipelineByQuarter[qMatch].countries[c]) pipelineByQuarter[qMatch].countries[c] = { amount: 0, weighted: 0 };
                pipelineByQuarter[qMatch].countries[c].amount += amt;
                pipelineByQuarter[qMatch].countries[c].weighted += wAmt;
                pipelineByQuarter[qMatch].deals.push({ name: dealName, amount: amt, weighted: wAmt, country: c });
            }
        }

        const dKey = keys.find(k => k.toLowerCase().includes('date') || k.toLowerCase().includes('start'));
        let year = 'Unknown';
        if (dKey && r[dKey]) {
            let d = null;
            if (r[dKey] instanceof Date) d = r[dKey];
            else if (typeof r[dKey] === 'number' && r[dKey] > 40000) d = new Date(Math.round((r[dKey] - 25569) * 86400 * 1000));
            else {
                const parsed = new Date(r[dKey]);
                if (!isNaN(parsed.getTime())) d = parsed;
            }
            if (d) {
                year = d.getFullYear().toString();
                if (year === '2026') {
                    const m = d.getMonth();
                    pipelineInfluxData[m].count++;
                    pipelineInfluxData[m].amount += amt;
                    pipelineInfluxData[m].weighted += wAmt;
                    pipelineInfluxData[m].value += amt;
                    const nKey = keys.find(k => k.toLowerCase().includes('name') || k.toLowerCase().includes('customer') || k.toLowerCase().includes('end user'));
                    if (nKey) pipelineInfluxData[m].accounts.push(String(r[nKey]).trim());
                }
            }
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
        globalTotalWeighted: Object.values(pipelineByCountry).reduce((acc, curr) => acc + curr.weighted, 0)
    };
}


export function getPartnerStats(data, filterCountry, workbookData) {
    const pKeys = Object.keys(data[0]);
    const pCountryKey = pKeys.find(k => k.toLowerCase().includes('country') || k.toLowerCase().includes('region') || k.toLowerCase() === 'nation');
    const pNameKey = pKeys.find(k => k.toLowerCase().includes('partner') || k.toLowerCase().includes('name') || k.toLowerCase().includes('???텢')) || pKeys[0];

    const counts = {};
    const partnerGroups = {};
    data.forEach(r => {
        let country = normalizeCountry(r[pCountryKey]);
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
        const sKey = Object.keys(r).find(k => k.toLowerCase().includes('status'));
        const statusStr = sKey ? String(r[sKey]).trim().toLowerCase() : '';

        if (statusStr.includes('running') || statusStr.includes('progress') || statusStr.includes('筌욊쑵六') || statusStr.includes('ing')) {
            const pKey = Object.keys(r).find(k => k.toLowerCase() === 'partner');
            let pName = pKey ? String(r[pKey]).trim() : 'Unknown';
            if (!pName || pName.toLowerCase() === 'unknown' || pName === '') pName = 'Direct/Unknown';

            if (filterCountry && pName !== 'Direct/Unknown' && !currentPartnerNames.has(pName.toLowerCase())) return;

            const vKey = Object.keys(r).find(k => k.toUpperCase().includes('ESTIMATED VALUE') && k.toUpperCase().includes('KOR') && k.toUpperCase().includes('USD'));
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

export function getGenericCountryStats(data, filterCountry) {
    if (!data || data.length === 0) return null;
    const keys = Object.keys(data[0]);
    const countryKey = keys.find(k => k.toLowerCase().includes('country') || k.toLowerCase().includes('region') || k.toLowerCase() === 'nation');
    if (!countryKey) return null;

    const countryCounts = {};
    const yearlyCounts = {};

    const getYear = (row) => {
        const dKey = keys.find(k => k.toUpperCase().replace(/\s/g, '') === 'CONTRACTSTART') ||
            keys.find(k => k.toUpperCase().includes('START') && k.toUpperCase().includes('DATE')) ||
            keys.find(k => k.toUpperCase().includes('YEAR'));
        const v = row[dKey];
        if (v instanceof Date) return v.getFullYear().toString();
        if (typeof v === 'number' && v > 40000) return new Date(Math.round((v - 25569) * 86400 * 1000)).getFullYear().toString();
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

    const sortedTotal = Object.entries(countryCounts).sort(sortCountriesByCount);
    const sortedYears = Object.keys(yearlyCounts).sort((a, b) => (a === 'Unknown' ? 1 : b === 'Unknown' ? -1 : b - a));

    return { sortedTotal, yearlyCounts, sortedYears };
}

export function getExpiringContractsStats(data) {
    if (!data || data.length === 0) return null;
    const keys = Object.keys(data[0]);
    const endKey = keys.find(k => k.toUpperCase().replace(/\s/g, '').includes('CONTRACTEND')) || keys.find(k => k.toUpperCase().includes('END'));
    if (!endKey) return null;

    const dealNameKey = keys.find(k => k.toUpperCase().replace(/\s/g, '').includes('CRMDEALNAME')) || keys.find(k => k.toUpperCase().includes('DEAL') && k.toUpperCase().includes('NAME')) || keys.find(k => k.toUpperCase().includes('DEAL'));
    const clientKey = keys.find(k => k.toUpperCase().includes('CLIENT') || k.toUpperCase().includes('CUSTOMER') || k.toUpperCase().includes('ACCOUNT'));
    const contractYrKey = keys.find(k => k.toUpperCase().replace(/\s/g, '').includes('CONTRACTYR')) || keys.find(k => k.toUpperCase().includes('YR')) || keys.find(k => k.toUpperCase().includes('YEAR'));

    const now = new Date('2026-03-18');
    const threeMonthsLater = new Date(now);
    threeMonthsLater.setMonth(now.getMonth() + 3);

    const getExpDate = (row) => {
        let d = row[endKey];
        if (d instanceof Date) return d;
        if (typeof d === 'number' && d > 40000) return new Date(Math.round((d - 25569) * 86400 * 1000));
        if (d) { const parsed = new Date(d); return isNaN(parsed.getTime()) ? null : parsed; }
        return null;
    };

    const expiringDeals = data.filter(row => {
        const expDate = getExpDate(row);
        return expDate && expDate >= now && expDate <= threeMonthsLater;
    }).sort((a, b) => (getExpDate(a) || 0) - (getExpDate(b) || 0));

    if (expiringDeals.length === 0) return null;

    return expiringDeals.map(deal => {
        const expDate = getExpDate(deal);
        let name = 'Unknown Deal';
        if (dealNameKey && deal[dealNameKey]) name = deal[dealNameKey];
        else if (clientKey && deal[clientKey]) name = deal[clientKey];
        const yr = contractYrKey ? deal[contractYrKey] : '';
        return { name, date: expDate ? expDate.toISOString().split('T')[0] : 'N/A', year: yr };
    });
}

export function getPartnerPerformanceStats(data) {
    if (!data || data.length === 0) return null;
    const keys = Object.keys(data[0]);
    const companyKey = keys.find(k => k.toLowerCase().replace(/\s/g, '') === 'companyname');
    const countryKey = keys.find(k => k.toLowerCase().includes('country'));
    const tcvKey = keys.find(k => k.toLowerCase().replace(/\s/g, '') === 'accumulatedtcv');

    if (!companyKey || !tcvKey) return null;

    const stats = data.map(row => ({
        name: String(row[companyKey] || 'Unknown').trim(),
        country: String(row[countryKey] || 'N/A').trim(),
        tcv: parseCurrency(row[tcvKey])
    })).filter(p => p.tcv > 0).sort((a, b) => b.tcv - a.tcv).slice(0, 10);

    return stats.length > 0 ? stats : null;
}

export function getPocStats(data, filters, workbookData) {
    const sheet9Data = workbookData['Sheet9'] || [];
    const mergedData = data.map(pRow => {
        const nameKey = Object.keys(pRow).find(k => k.toLowerCase().replace(/[^a-z]/g, '') === 'crmpocname');
        const pName = nameKey ? String(pRow[nameKey]).trim().toLowerCase() : '';
        let matchedS9 = sheet9Data.find(sRow => {
            const sNameKey = Object.keys(sRow).find(k => k.toLowerCase().replace(/[^a-z]/g, '') === 'crmpocname');
            return sNameKey && String(sRow[sNameKey]).trim().toLowerCase() === pName;
        });
        return { ...pRow, _s9: matchedS9 || {} };
    });

    const uniqueValues = { countries: new Set(['All']), industries: new Set(['All']), partners: new Set(['All']) };
    mergedData.forEach(r => {
        const cKey = Object.keys(r).find(k => k.toLowerCase().includes('country'));
        const iKey = Object.keys(r).find(k => k.toLowerCase().includes('industry'));
        const pKey = Object.keys(r).find(k => k.toLowerCase() === 'partner');
        if (cKey && r[cKey]) uniqueValues.countries.add(normalizeCountry(r[cKey]));
        if (iKey && r[iKey]) uniqueValues.industries.add(String(r[iKey]).trim());
        if (pKey && r[pKey]) uniqueValues.partners.add(String(r[pKey]).trim());
    });

    const displayData = mergedData.filter(r => {
        const cKey = Object.keys(r).find(k => k.toLowerCase().includes('country'));
        const rCountry = cKey ? normalizeCountry(r[cKey]) : '';
        const rIndustry = String(r[Object.keys(r).find(k => k.toLowerCase().includes('industry'))] || '').trim();
        const rPartner = String(r[Object.keys(r).find(k => k.toLowerCase() === 'partner')] || '').trim();
        return (filters.country === 'All' || rCountry === filters.country) &&
            (filters.industry === 'All' || rIndustry === filters.industry) &&
            (filters.partner === 'All' || rPartner === filters.partner);
    });

    let s = {
        longTermCount: 0, midTermCount: 0, normalCount: 0, runningList: [], staledRunningList: [], runningNames: [], partnerRunDays: {},
        statusStats: { won: 0, drop: 0, running: 0, hold: 0, others: 0 }, totalWonDays: 0, wonCountForStats: 0,
        totalHold: 0, holdNames: [],
        influxData: Array(12).fill(0).map(() => ({ count: 0, estimated: 0, weighted: 0, accounts: [] })), industryVal: {}, partnerRankingData: {},
        monthlyActive: Array(12).fill(0).map(() => ({ count: 0, tcv: 0 }))
    };

    displayData.forEach(r => {
        const allKeys = Object.keys(r);
        const s9Keys = Object.keys(r._s9 || {});

        const wdKey = allKeys.find(k => k.toLowerCase().includes('working days')) || allKeys.find(k => k.toLowerCase().includes('workingdays')) || allKeys.find(k => k.toLowerCase().includes('running days'));
        const s9WdKey = s9Keys.find(k => k.toLowerCase().includes('working days')) || s9Keys.find(k => k.toLowerCase().includes('workingdays')) || s9Keys.find(k => k.toLowerCase().includes('running days'));
        let runningDays = Number(r[wdKey] || (r._s9 && r._s9[s9WdKey]) || 0);

        if (runningDays >= 100) s.longTermCount++; else if (runningDays >= 60) s.midTermCount++; else s.normalCount++;

        const statusKey = allKeys.find(k => k.toLowerCase().includes('current status')) || allKeys.find(k => k.toLowerCase().includes('currentstatus')) || allKeys.find(k => k.toLowerCase().includes('status'));
        const curStatus = String(r[statusKey] || '').trim().toLowerCase();
        const startValForStatus = r[allKeys.find(k => k.toLowerCase().replace(/\s/g, '').includes('pocstart') || k.toLowerCase().replace(/\s/g, '').includes('poc??뽰삂'))];
        let dForStatus = startValForStatus ? (startValForStatus instanceof Date ? startValForStatus : new Date(startValForStatus)) : null;
        if (typeof startValForStatus === 'number' && startValForStatus > 30000) dForStatus = new Date(Math.round((startValForStatus - 25569) * 86400 * 1000));

        if (curStatus.includes('hold') || curStatus.includes('pause')) {
            s.totalHold++;
            const name = r[allKeys.find(k => k.toLowerCase().replace(/[^a-z]/g, '') === 'crmpocname')] || 'Unknown';
            s.holdNames.push(name);
        }

        if (curStatus.includes('running') || curStatus.includes('progress') || curStatus.includes('ing')) {
            s.statusStats.running++;
            const name = r[allKeys.find(k => k.toLowerCase().replace(/[^a-z]/g, '') === 'crmpocname')] || 'Unknown';
            s.runningNames.push(name);
        } else if (dForStatus && !isNaN(dForStatus.getTime()) && dForStatus.getFullYear() === 2026) {
            if (curStatus.includes('won') || curStatus.includes('complete') || curStatus.includes('success')) s.statusStats.won++;
            else if (curStatus.includes('drop') || curStatus.includes('cancel') || curStatus.includes('fail')) s.statusStats.drop++;
            else if (curStatus.includes('hold') || curStatus.includes('pause')) s.statusStats.hold++;
            else if (curStatus !== '') s.statusStats.others++;
        }

        if (curStatus.includes('running') || curStatus.includes('progress') || runningDays >= 100) {
            const entry = {
                name: r[allKeys.find(k => k.toLowerCase().replace(/[^a-z]/g, '') === 'crmpocname')] || 'Unknown',
                partner: r[allKeys.find(k => k.toLowerCase() === 'partner')] || 'Unknown',
                country: r[allKeys.find(k => k.toLowerCase().includes('country'))] || 'Unknown',
                industry: r[allKeys.find(k => k.toLowerCase().includes('industry'))] || 'Unknown',
                days: runningDays,
                notes: r._s9[s9Keys.find(k => k.toLowerCase().includes('notes'))] || '',
                techComm: r[allKeys.find(k => k.toLowerCase().includes('technical comment') || k.toLowerCase().includes('report'))] || '',
                isStalled: runningDays >= 100
            };
            s.runningList.push(entry);
            if (runningDays >= 100) s.staledRunningList.push(entry);
        }

        // Prioritize "POC Start" over "License Start" as per user request
        const startValKey = allKeys.find(k => {
            const n = k.toLowerCase().replace(/[^a-z0-9??뽰삂]/g, '');
            return n === 'pocstart' || (n.includes('poc') && n.includes('start')) || n.includes('poc??뽰삂') || n === 'startdate';
        }) || allKeys.find(k => {
            const n = k.toLowerCase().replace(/[^a-z0-9]/g, '');
            return n.includes('license') && n.includes('start');
        });
        const startVal = r[startValKey];

        const estValKey = allKeys.find(k => k.toUpperCase().includes('ESTIMATED VALUE') && k.toUpperCase().includes('KOR') && k.toUpperCase().includes('USD'));
        const wValKey = allKeys.find(k => k.toUpperCase().includes('WEIGHTED VALUE') && k.toUpperCase().includes('KOR') && k.toUpperCase().includes('USD')) ||
                        allKeys.find(k => k.toUpperCase().includes('WEIGHTED') && k.toUpperCase().includes('USD'));
        
        const estVal = parseCurrency(r[estValKey] || 0);
        const wVal = parseCurrency(r[wValKey] || 0);

        if (startVal) {
            let d = (startVal instanceof Date) ? startVal : new Date(startVal);
            // Robust check for Excel serial date numbers
            if (typeof startVal === 'number' && startVal > 30000) {
                d = new Date(Math.round((startVal - 25569) * 86400 * 1000));
            } else if (typeof startVal === 'string' && !isNaN(startVal) && startVal.length > 4) {
                d = new Date(Math.round((parseFloat(startVal) - 25569) * 86400 * 1000));
            }

            if (d && !isNaN(d.getTime()) && d.getFullYear() === 2026) {
                const m = d.getMonth();
                s.influxData[m].count++; 
                s.influxData[m].estimated += estVal;
                s.influxData[m].weighted += wVal;
                const pocNameKey = allKeys.find(k => k.toLowerCase().replace(/[^a-z]/g, '') === 'crmpocname');
                if (pocNameKey && r[pocNameKey]) s.influxData[m].accounts.push(String(r[pocNameKey]).trim());
            }

            // Monthly Active Analysis (2026)
            if (d && !isNaN(d.getTime())) {
                const year2026 = 2026;
                let pocEndDate = new Date(8640000000000000); // Future date for running/hold
                if (curStatus.includes('won') || curStatus.includes('drop') || curStatus.includes('fail') || curStatus.includes('complete')) {
                    pocEndDate = new Date(d.getTime() + (runningDays * 86400 * 1000));
                }

                for (let m = 0; m < 12; m++) {
                    const monthStart = new Date(year2026, m, 1);
                    const monthEnd = new Date(year2026, m + 1, 0);
                    if (d <= monthEnd && pocEndDate >= monthStart) {
                        s.monthlyActive[m].tcv += estVal;
                        s.monthlyActive[m].count++;
                    }
                }
            }
        }

        const pName = String(r[allKeys.find(k => k.toLowerCase() === 'partner')] || 'Unknown').trim();
        if (pName && pName !== 'Unknown' && pName !== '') {
            if (!s.partnerRunDays[pName]) s.partnerRunDays[pName] = { sum: 0, count: 0 };
            s.partnerRunDays[pName].sum += runningDays; s.partnerRunDays[pName].count++;
        }

        if (curStatus.includes('won') || curStatus.includes('complete') || curStatus.includes('success')) {
            const wd = parseCurrency(r[allKeys.find(k => k.toLowerCase().includes('working days'))] || 0);
            if (wd > 0) { s.totalWonDays += wd; s.wonCountForStats++; }
        }

        if (curStatus.includes('running') || curStatus.includes('progress') || curStatus.includes('ing')) {
            const ind = String(r[allKeys.find(k => k.toLowerCase().includes('industry'))] || 'Other').trim();
            s.industryVal[ind] = (s.industryVal[ind] || 0) + estVal;
            if (!s.partnerRankingData[pName]) s.partnerRankingData[pName] = { count: 0, sumValue: 0 };
            s.partnerRankingData[pName].count++; s.partnerRankingData[pName].sumValue += estVal;
        }
    });

    s.runningList.sort((a, b) => b.days - a.days);
    s.partnerAvg = Object.keys(s.partnerRunDays).map(p => ({ partner: p, avg: Math.round(s.partnerRunDays[p].sum / s.partnerRunDays[p].count) })).sort((a, b) => b.avg - a.avg);
    s.sortedIndustry = Object.entries(s.industryVal).map(([name, val]) => ({ name, val })).sort((a, b) => b.val - a.val);
    s.sortedPartnersRanking = Object.entries(s.partnerRankingData).map(([name, st]) => ({ name, ...st })).sort((a, b) => b.count - a.count || b.sumValue - a.sumValue);

    return { stats: s, uniqueValues };
}

export function getEventStats(eventData, filterCountry) {
    const filteredEvents = eventData.filter(row => {
        if (filterCountry && !isCountryMatch(row, filterCountry)) return false;
        const dateKey = Object.keys(row).find(k => k.toLowerCase().includes('date'));
        const dateVal = dateKey ? row[dateKey] : null;
        let year = 0;
        if (dateVal instanceof Date) year = dateVal.getFullYear();
        else if (typeof dateVal === 'number' && dateVal > 40000) year = new Date(Math.round((dateVal - 25569) * 86400 * 1000)).getFullYear();
        else { const match = String(dateVal).match(/(20\d{2})/); if (match) year = parseInt(match[1]); }
        return year === 2026;
    });

    if (filteredEvents.length === 0) return null;

    let totalSpending = 0, totalPOC = 0, totalDeals = 0;
    const comparisonData = filteredEvents.map(row => {
        const sKey = Object.keys(row).find(k => k.toLowerCase().includes('spending'));
        const pKey = Object.keys(row).find(k => k.toLowerCase().includes('poc') && k.toLowerCase().includes('generated'));
        const dKey = Object.keys(row).find(k => k.toLowerCase().includes('deal') && (k.toLowerCase().includes('converted') || k.toLowerCase().includes('closed')));
        const nKey = Object.keys(row).find(k => k.toLowerCase().includes('event') || k.toLowerCase().includes('name') || k.toLowerCase().includes('??梨')) || Object.keys(row)[0];

        const spend = parseCurrency(row[sKey]);
        const pocs = parseInt(row[pKey]) || 0;
        const deals = parseInt(row[dKey]) || 0;

        totalSpending += spend;
        totalPOC += pocs;
        totalDeals += deals;
        return { name: String(row[nKey] || 'Unnamed').trim(), spend, pocs, deals };
    });

    return { totalSpending, totalPOC, totalDeals, eventCount: filteredEvents.length, comparisonData, costPerPOC: totalPOC > 0 ? (totalSpending / totalPOC) : 0, costPerDeal: totalDeals > 0 ? (totalSpending / totalDeals) : 0 };
}

export function getCountrySpecificStats(data) {
    const sortedYears = ['2026', '2025', '2024', '2023'];
    const summary = {};
    sortedYears.forEach(y => summary[y] = { kTcv: 0, lArr: 0 });

    data.forEach(row => {
        const dKey = Object.keys(row).find(k => k.toLowerCase().includes('contractstart')) || Object.keys(row).find(k => k.toLowerCase().includes('date'));
        if (!row[dKey]) return;
        let d = (row[dKey] instanceof Date) ? row[dKey] : (typeof row[dKey] === 'number' && row[dKey] > 40000) ? new Date(Math.round((row[dKey] - 25569) * 86400 * 1000)) : new Date(row[dKey]);
        if (d && !isNaN(d.getTime())) {
            const y = d.getFullYear().toString();
            if (summary[y]) {
                summary[y].kTcv += parseCurrency(row['KOR TCV(USD)'] || row['KOR TCV']);
                summary[y].lArr += parseCurrency(row['Local ARR'] || row['ARR']);
            }
        }
    });
    return { summary, sortedYears };
}

export function getServiceAnalysisStats(data) {
    const activeData = data.filter(r => r['Status'] === 'Active' && r['Services']);
    if (activeData.length === 0) return null;
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
    return {
        totalCustomers: activeData.length,
        singleServiceCustomers: upscaleTargetsArr.length,
        multiServiceCustomers: activeData.length - upscaleTargetsArr.length,
        sortedCombos: Object.entries(comboCounts).sort((a, b) => b[1] - a[1]),
        upsellTargets: upscaleTargetsArr.sort((a, b) => b.tcv - a.tcv),
        palette: CONFIG.COLORS
    };
}

