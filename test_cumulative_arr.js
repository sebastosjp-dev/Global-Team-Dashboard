/**
 * Verify cumulative ARR logic offline.
 * Simulates the new contractEntries approach.
 */
const assert = require('assert');

function testCumulativeArr() {
    const currentYear = 2026;
    const contracts = [
        { startYear: 2023, arrVal: 10000 },
        { startYear: 2024, arrVal: 50000 },
        { startYear: 2024, arrVal: 30000 },
        { startYear: 2025, arrVal: 20000 },
        { startYear: 2026, arrVal: 5000  },
    ];

    const yearlyArr = {};
    const minYear = contracts.reduce((m, e) => Math.min(m, e.startYear), Infinity);
    for (let y = minYear; y <= currentYear; y++) {
        yearlyArr[y] = 0;
    }
    contracts.forEach(({ startYear, arrVal }) => {
        for (let y = startYear; y <= currentYear; y++) {
            yearlyArr[y] += arrVal;
        }
    });

    console.log('yearlyArr:', yearlyArr);

    // 2023: only contract#1 = 10000
    assert.strictEqual(yearlyArr[2023], 10000, '2023 should be 10000');
    // 2024: contract#1 + #2 + #3 = 10000+50000+30000 = 90000
    assert.strictEqual(yearlyArr[2024], 90000, '2024 should be 90000');
    // 2025: +#4 = 90000+20000 = 110000
    assert.strictEqual(yearlyArr[2025], 110000, '2025 should be 110000');
    // 2026: +#5 = 110000+5000 = 115000
    assert.strictEqual(yearlyArr[2026], 115000, '2026 should be 115000');

    // Graph should be monotonically non-decreasing
    const years = Object.keys(yearlyArr).sort();
    for (let i = 1; i < years.length; i++) {
        assert(yearlyArr[years[i]] >= yearlyArr[years[i - 1]],
            `yearlyArr should be non-decreasing: ${years[i-1]}=${yearlyArr[years[i-1]]} vs ${years[i]}=${yearlyArr[years[i]]}`);
    }

    // sumArr should be current year's value
    const sumArr = yearlyArr[currentYear];
    assert.strictEqual(sumArr, 115000, 'sumArr (headline) should be cumulative current year total');

    console.log('✅ All cumulative ARR tests passed.');
}

testCumulativeArr();
