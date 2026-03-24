// --- PARTNER Summary & Analytics ---
function renderPartnerMetrics(data, tabName, filterCountry = null) {
    const metricsGrid = document.getElementById('tab-metrics-grid');
    const isCountryTab = tabName && filterCountry === null && !['ORDER SHEET', 'PIPELINE', 'PARTNER', 'POC', 'EVENT', 'END USER (CSM)'].includes(tabName);
    
    if (tabName === 'PARTNER' || isCountryTab) {
        // Partner count stat cards
        const pKeys = Object.keys(data[0]);
        const pCountryKey = pKeys.find(k => k.toLowerCase().includes('country'));
        if (pCountryKey) {
            const counts = {};
            data.forEach(r => {
                const c = String(r[pCountryKey] || '').trim();
                if (c) counts[c] = (counts[c] || 0) + 1;
            });
            // ... (Rest of Partner logic from app.js)
        }
    }
}
window.renderPartnerMetrics = renderPartnerMetrics;
