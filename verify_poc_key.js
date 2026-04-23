// Simulate the key finding logic with actual Excel headers
const headers = [
    'No.', 'Country', 'Industry', 'CRM POC Name', 'Current Status',
    'POC License Start', 'POC License End', 'Partner',
    'Estimated Value KOR (USD)', 'Weighted Value KOR (USD)',
    'POC Start', 'POC End', 'POC In Progress', 'Working Days',
    'POC Report', 'POC Notes', 'Technical Comment', 'Sales Comment'
];

// OLD logic (buggy)
const oldKey = headers.find(k => {
    const n = k.toLowerCase().replace(/[^a-z0-9]/g, '');
    return n === 'pocstart' || (n.includes('poc') && n.includes('start')) || n === 'startdate';
}) || headers.find(k => {
    const n = k.toLowerCase().replace(/[^a-z0-9]/g, '');
    return n.includes('license') && n.includes('start');
});

// NEW logic (fixed)
const newKey = headers.find(k => {
    const n = k.toLowerCase().replace(/[^a-z0-9]/g, '');
    return n === 'pocstart';
}) || headers.find(k => {
    const n = k.toLowerCase().replace(/[^a-z0-9]/g, '');
    return (n.includes('poc') && n.includes('start') && !n.includes('license')) || n === 'startdate';
}) || headers.find(k => {
    const n = k.toLowerCase().replace(/[^a-z0-9]/g, '');
    return n.includes('license') && n.includes('start');
});

console.log('=== Key Finding Verification ===');
console.log('OLD logic matched:', oldKey);
console.log('NEW logic matched:', newKey);
console.log('');
console.log('Expected: POC Start');
console.log('OLD correct:', oldKey === 'POC Start');
console.log('NEW correct:', newKey === 'POC Start');
