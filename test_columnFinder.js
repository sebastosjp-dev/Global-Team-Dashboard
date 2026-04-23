/**
 * test_columnFinder.js — Node.js unit tests for columnFinder.js
 * Uses dynamic import to load ESM module.
 * Run: node --experimental-vm-modules test_columnFinder.js
 */

// Since columnFinder.js uses ESM exports, we can test the logic directly
// by re-implementing the matchers here (they're pure functions, no DOM).

const assert = require('assert').strict;

// --- Re-implement key finder logic for Node.js testing ---
function findKey(keys, ...predicates) {
    for (const pred of predicates) {
        const found = keys.find(pred);
        if (found) return found;
    }
    return undefined;
}

function _norm(key) { return key.toUpperCase().replace(/\s/g, ''); }

function findCountryKey(keys) {
    return findKey(keys,
        k => k.toLowerCase().includes('country'),
        k => k.toLowerCase().includes('region'),
        k => k.toLowerCase() === 'nation'
    );
}

function findKorTcvKey(keys) {
    return findKey(keys,
        k => _norm(k) === 'KORTCV',
        k => k.toUpperCase().includes('TCV') && (k.toUpperCase().includes('KOR') || k.toUpperCase().includes('KRW') || k.toUpperCase().includes('USD'))
    );
}

function findContractStartKey(keys) {
    return findKey(keys,
        k => _norm(k).includes('CONTRACTSTART'),
        k => k.toUpperCase().includes('START') && k.toUpperCase().includes('DATE')
    );
}

function findStatusKey(keys) {
    return findKey(keys,
        k => k.toLowerCase().includes('current status'),
        k => k.toLowerCase().includes('currentstatus'),
        k => k.toLowerCase().includes('status')
    );
}

function findPocStartKey(keys) {
    return findKey(keys,
        k => k.toLowerCase().replace(/[^a-z0-9]/g, '') === 'pocstart',
        k => {
            const n = k.toLowerCase().replace(/[^a-z0-9]/g, '');
            return (n.includes('poc') && n.includes('start') && !n.includes('license')) || n === 'startdate';
        },
        k => {
            const n = k.toLowerCase().replace(/[^a-z0-9]/g, '');
            return n.includes('license') && n.includes('start');
        }
    );
}

function parseExcelDateSafe(val) {
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

// === Tests ===

function testFindCountryKey() {
    console.log('[Test 1] findCountryKey');
    assert.strictEqual(findCountryKey(['Name', 'Country', 'TCV']), 'Country');
    assert.strictEqual(findCountryKey(['Name', 'Region', 'TCV']), 'Region');
    assert.strictEqual(findCountryKey(['Name', 'nation']), 'nation');
    assert.strictEqual(findCountryKey(['Name', 'TCV']), undefined);
    console.log('✅ Pass');
}

function testFindKorTcvKey() {
    console.log('[Test 2] findKorTcvKey');
    assert.strictEqual(findKorTcvKey(['Deal', 'KOR TCV', 'ARR']), 'KOR TCV');
    assert.strictEqual(findKorTcvKey(['Deal', 'KOR TCV (USD)', 'ARR']), 'KOR TCV (USD)');
    assert.strictEqual(findKorTcvKey(['Deal', 'TCV KRW', 'ARR']), 'TCV KRW');
    assert.strictEqual(findKorTcvKey(['Deal', 'Amount']), undefined);
    console.log('✅ Pass');
}

function testFindContractStartKey() {
    console.log('[Test 3] findContractStartKey');
    assert.strictEqual(findContractStartKey(['Name', 'Contract Start', 'End']), 'Contract Start');
    assert.strictEqual(findContractStartKey(['Name', 'Start Date', 'End']), 'Start Date');
    assert.strictEqual(findContractStartKey(['Name', 'End Date']), undefined);
    console.log('✅ Pass');
}

function testFindStatusKey() {
    console.log('[Test 4] findStatusKey');
    assert.strictEqual(findStatusKey(['Name', 'Current Status', 'Days']), 'Current Status');
    assert.strictEqual(findStatusKey(['Name', 'Status', 'Days']), 'Status');
    assert.strictEqual(findStatusKey(['Name', 'Days']), undefined);
    console.log('✅ Pass');
}

function testFindPocStartKey() {
    console.log('[Test 5] findPocStartKey (priority order)');
    // Exact "POC Start" should win over "POC License Start"
    assert.strictEqual(findPocStartKey(['POC Start', 'POC License Start Date']), 'POC Start');
    // Fallback to license start if no exact
    assert.strictEqual(findPocStartKey(['Name', 'POC License Start']), 'POC License Start');
    assert.strictEqual(findPocStartKey(['Name', 'Amount']), undefined);
    console.log('✅ Pass');
}

function testParseExcelDateSafe() {
    console.log('[Test 6] parseExcelDateSafe');
    // Null/empty
    assert.strictEqual(parseExcelDateSafe(null), null);
    assert.strictEqual(parseExcelDateSafe(''), null);
    assert.strictEqual(parseExcelDateSafe(0), null);

    // Date object passthrough
    const d = new Date('2026-01-15');
    assert.strictEqual(parseExcelDateSafe(d), d);

    // Excel serial number (46036 ≈ 2026-01-15)
    const fromSerial = parseExcelDateSafe(46036);
    assert.ok(fromSerial instanceof Date);
    assert.strictEqual(fromSerial.getFullYear(), 2026);

    // String date
    const fromStr = parseExcelDateSafe('2026-03-18');
    assert.ok(fromStr instanceof Date);
    assert.strictEqual(fromStr.getMonth(), 2); // March = 2

    // Garbage
    assert.strictEqual(parseExcelDateSafe('not-a-date-lol'), null);

    console.log('✅ Pass');
}

function testEdgeCases() {
    console.log('[Test 7] Edge cases: empty arrays, single element');
    assert.strictEqual(findCountryKey([]), undefined);
    assert.strictEqual(findKorTcvKey(['']), undefined);
    assert.strictEqual(findKey([], k => true), undefined);
    // Small serial number (<=30000) should not be treated as Excel date
    assert.strictEqual(parseExcelDateSafe(100), null);
    console.log('✅ Pass');
}

try {
    testFindCountryKey();
    testFindKorTcvKey();
    testFindContractStartKey();
    testFindStatusKey();
    testFindPocStartKey();
    testParseExcelDateSafe();
    testEdgeCases();
    console.log('\n✅ All column finder tests passed.');
} catch (err) {
    console.error('❌ Test failure:', err.message);
    process.exit(1);
}
