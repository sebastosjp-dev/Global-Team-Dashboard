/**
 * utils.js — Shared pure utility functions (no DOM dependencies)
 */

/**
 * Parse a currency string/number into a numeric value.
 * @param {*} val - Raw value
 * @returns {number}
 */
export function parseCurrency(val) {
    if (!val) return 0;
    if (typeof val === 'number') return val;
    const clean = String(val).replace(/[^0-9.-]+/g, '');
    return parseFloat(clean) || 0;
}

/**
 * Format a number with locale-appropriate grouping.
 * @param {number} val
 * @param {boolean} [isKRW=false]
 * @returns {string}
 */
export function formatCurrency(val, isKRW = false) {
    const rounded = Math.round(val || 0);
    if (rounded === 0) return '0';
    const locale = isKRW ? 'ko-KR' : 'en-US';
    return new Intl.NumberFormat(locale, {
        minimumFractionDigits: 0,
        maximumFractionDigits: 0
    }).format(rounded);
}

/**
 * Parse an Excel cell value into a JavaScript Date.
 * Handles Date objects, Excel serial numbers, and date strings.
 * @param {*} val - Raw cell value
 * @returns {Date|null}
 */
export function parseExcelDate(val) {
    if (!val) return null;
    if (val instanceof Date) return val;
    if (typeof val === 'number' && val > 30000) {
        return new Date(Math.round((val - 25569) * 86400 * 1000));
    }
    if (typeof val === 'string' && !isNaN(val) && val.length > 4) {
        return new Date(Math.round((parseFloat(val) - 25569) * 86400 * 1000));
    }
    const parsed = new Date(val);
    return isNaN(parsed.getTime()) ? null : parsed;
}

/**
 * Find a column key from an array using multiple match strategies.
 * @param {string[]} keys
 * @param {...function(string): boolean} patterns
 * @returns {string|undefined}
 */
export function findKey(keys, ...patterns) {
    for (const pattern of patterns) {
        const found = keys.find(pattern);
        if (found) return found;
    }
    return undefined;
}

/**
 * Check if a status string matches any of the given terms.
 * @param {string} status
 * @param {...string} terms
 * @returns {boolean}
 */
export function matchStatus(status, ...terms) {
    const lower = String(status || '').trim().toLowerCase();
    return terms.some(t => lower.includes(t));
}

/**
 * Normalize country name variations into a canonical form.
 * @param {string} raw
 * @returns {string|null}
 */
export function normalizeCountry(raw) {
    if (!raw) return null;
    const c = String(raw).trim();
    const up = c.toUpperCase();
    if (up === 'IDN' || up.includes('INDONESIA')) return 'Indonesia';
    if (up === 'US' || up === 'USA' || up.includes('UNITED STATES') || up.includes('미국')) return 'USA';
    if (up === 'MA' || up === 'MAL' || up.includes('MALAYSIA')) return 'Malaysia';
    if (up === 'TH' || up === 'THA' || up.includes('THAILAND')) return 'Thailand';
    if (up === 'PH' || up === 'PHI' || up === 'PHL' || up.includes('PHILIPPINES')) return 'Philippines';
    if (up === 'VN' || up === 'VNM' || up.includes('VIETNAM')) return 'Vietnam';
    if (up === 'TUR' || up.includes('TURKEY') || up === 'TR') return 'Turkey';
    if (up === 'SIN' || up.includes('SINGAPORE') || up === 'SGP') return 'Singapore';
    return c;
}

/**
 * Check if a row's country matches the filter.
 * @param {Object} row
 * @param {string|null} filterCountry
 * @returns {boolean}
 */
export function isCountryMatch(row, filterCountry) {
    if (!filterCountry || filterCountry === 'All') return true;
    const k = Object.keys(row).find(k => k.toLowerCase().includes('country') || k.toLowerCase().includes('region') || k.toLowerCase().includes('nation'));
    if (!k) return false;
    return normalizeCountry(row[k]) === filterCountry;
}

/** Get sort priority for a country (lower = higher priority) */
export const getCountrySortOrder = (c) => {
    if (c === 'Indonesia') return 1;
    if (c === 'Malaysia') return 2;
    if (c === 'Thailand') return 3;
    return 4;
};

export const sortCountriesByAmount = (a, b) => {
    const oA = getCountrySortOrder(a[0]), oB = getCountrySortOrder(b[0]);
    if (oA !== oB) return oA - oB;
    return b[1].amount - a[1].amount;
};

export const sortCountriesByCount = (a, b) => {
    const oA = getCountrySortOrder(a[0]), oB = getCountrySortOrder(b[0]);
    if (oA !== oB) return oA - oB;
    return b[1] - a[1];
};
