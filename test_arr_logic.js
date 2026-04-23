function distributeAmountByYear(startDate, endDate, amount) {
    const allocations = {};
    const totalDuration = endDate.getTime() - startDate.getTime();
    if (totalDuration <= 0) {
        allocations[startDate.getFullYear()] = amount;
        return allocations;
    }

    const startYear = startDate.getFullYear();
    const endYear = endDate.getFullYear();

    for (let y = startYear; y <= endYear; y++) {
        // yearStart is Jan 1st of year Y
        const yearStart = new Date(y, 0, 1).getTime();
        // yearEnd is Jan 1st of year Y+1
        const yearEnd = new Date(y + 1, 0, 1).getTime();

        const overlapStart = Math.max(startDate.getTime(), yearStart);
        const overlapEnd = Math.min(endDate.getTime(), yearEnd);

        const durationInYear = Math.max(0, overlapEnd - overlapStart);
        const proportion = durationInYear / totalDuration;
        
        allocations[y] = amount * proportion;
    }
    return allocations;
}

const start = new Date('2024-04-15T00:00:00Z');
const end = new Date('2026-04-15T00:00:00Z'); // Adding 1 day to make exactly 24 months visually, but let's test with user exact 04-14.
const endUser = new Date('2026-04-14T23:59:59Z'); // User said 04-14. Often these are meant to be end of day.

console.log("=== EXACT MS CALCULATION using 2026-04-14 23:59:59 ===");
console.log(distributeAmountByYear(start, endUser, 900000));

// What if we do exactly month-based logic as user demonstrated, but fallback to days if needed?
console.log("\n=== USER EXPECTATION ===");
console.log("2024: ~318,750");
console.log("2025: 450,000");
console.log("2026: ~131,250");
