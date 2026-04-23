/**
 * columnFinder.js — Shared column-key resolution and date-parsing utilities.
 * Eliminates repeated `keys.find(k => k.toUpperCase()...)` patterns across services.js.
 * @module columnFinder
 */

/**
 * Normalize a header key for comparison (uppercase, no whitespace).
 * @param {string} key
 * @returns {string}
 */
function _norm(key) {
    return key.toUpperCase().replace(/\s/g, '');
}

/**
 * Find the first key matching any of the given normalized patterns.
 * @param {string[]} keys - All column header keys
 * @param {...function(string):boolean} predicates - Matcher functions (tested in order)
 * @returns {string|undefined}
 */
export function findKey(keys, ...predicates) {
    for (const pred of predicates) {
        const found = keys.find(pred);
        if (found) return found;
    }
    return undefined;
}

/**
 * Find a country/region column key.
 * @param {string[]} keys
 * @returns {string|undefined}
 */
export function findCountryKey(keys) {
    return findKey(keys,
        k => k.toLowerCase().includes('country'),
        k => k.toLowerCase().includes('region'),
        k => k.toLowerCase() === 'nation'
    );
}

/**
 * Find a KOR TCV (USD) column key.
 * @param {string[]} keys
 * @returns {string|undefined}
 */
export function findKorTcvKey(keys) {
    return findKey(keys,
        k => _norm(k) === 'KORTCV',
        k => k.toUpperCase().includes('TCV') && (k.toUpperCase().includes('KOR') || k.toUpperCase().includes('KRW') || k.toUpperCase().includes('USD'))
    );
}

/**
 * Find a KOR ARR column key.
 * @param {string[]} keys
 * @returns {string|undefined}
 */
export function findArrKey(keys) {
    return findKey(keys,
        k => _norm(k) === 'KORARR',
        k => k.toUpperCase().includes('ARR') && (k.toUpperCase().includes('KOR') || k.toUpperCase().includes('KRW'))
    );
}

/**
 * Find a KOR MRR column key.
 * @param {string[]} keys
 * @returns {string|undefined}
 */
export function findMrrKey(keys) {
    return findKey(keys,
        k => _norm(k) === 'KORMRR',
        k => k.toUpperCase().includes('MRR') && (k.toUpperCase().includes('KOR') || k.toUpperCase().includes('KRW'))
    );
}

/**
 * Find a Contract Start date column key.
 * @param {string[]} keys
 * @returns {string|undefined}
 */
export function findContractStartKey(keys) {
    return findKey(keys,
        k => _norm(k).includes('CONTRACTSTART'),
        k => k.toUpperCase().includes('START') && k.toUpperCase().includes('DATE')
    );
}

/**
 * Find a Contract End date column key.
 * @param {string[]} keys
 * @returns {string|undefined}
 */
export function findContractEndKey(keys) {
    return findKey(keys,
        k => _norm(k).includes('CONTRACTEND'),
        k => k.toUpperCase().includes('END')
    );
}

/**
 * Find a status column key.
 * @param {string[]} keys
 * @returns {string|undefined}
 */
export function findStatusKey(keys) {
    return findKey(keys,
        k => k.toLowerCase().includes('current status'),
        k => k.toLowerCase().includes('currentstatus'),
        k => k.toLowerCase().includes('status')
    );
}

/**
 * Find a deal/POC name column key.
 * @param {string[]} keys
 * @returns {string|undefined}
 */
export function findDealNameKey(keys) {
    return findKey(keys,
        k => _norm(k).includes('CRMDEALNAME'),
        k => k.toUpperCase().includes('DEAL') && k.toUpperCase().includes('NAME'),
        k => k.toUpperCase().includes('DEAL')
    );
}

/**
 * Find a CRM POC name column key.
 * @param {string[]} keys
 * @returns {string|undefined}
 */
export function findPocNameKey(keys) {
    return findKey(keys,
        k => k.toLowerCase().replace(/[^a-z]/g, '') === 'crmpocname'
    );
}

/**
 * Find a POC Start date column key (prioritizing exact match over license start).
 * @param {string[]} keys
 * @returns {string|undefined}
 */
export function findPocStartKey(keys) {
    return findKey(keys,
        k => k.toLowerCase().replace(/[^a-z0-9]/g, '') === 'pocstart',
        k => {
            const n = k.toLowerCase().replace(/[^a-z0-9시작]/g, '');
            return (n.includes('poc') && n.includes('start') && !n.includes('license')) || n.includes('poc시작') || n === 'startdate';
        },
        k => {
            const n = k.toLowerCase().replace(/[^a-z0-9]/g, '');
            return n.includes('license') && n.includes('start');
        }
    );
}

/**
 * Find an estimated value (KOR USD) column key.
 * @param {string[]} keys
 * @returns {string|undefined}
 */
export function findEstimatedValueKey(keys) {
    return findKey(keys,
        k => k.toUpperCase().includes('ESTIMATED VALUE') && k.toUpperCase().includes('KOR') && k.toUpperCase().includes('USD')
    );
}

/**
 * Find a weighted value (KOR USD) column key.
 * @param {string[]} keys
 * @returns {string|undefined}
 */
export function findWeightedValueKey(keys) {
    return findKey(keys,
        k => k.toUpperCase().includes('WEIGHTED VALUE') && k.toUpperCase().includes('KOR') && k.toUpperCase().includes('USD'),
        k => k.toUpperCase().includes('WEIGHTED') && k.toUpperCase().includes('USD')
    );
}

/**
 * Parse any Excel cell value into a Date safely.
 * Handles Date objects, serial numbers (>30000), numeric strings, and date strings.
 * @param {*} val
 * @returns {Date|null}
 */
export function parseExcelDateSafe(val) {
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
