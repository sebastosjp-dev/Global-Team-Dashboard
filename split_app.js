/**
 * split_app.js — Script to extract services.js, ui.js from app_backup.js
 * and create a new app.js entry point.
 * 
 * Run: node split_app.js
 */
const fs = require('fs');

const src = fs.readFileSync('app_backup.js', 'utf8');

// Helper: extract lines range (1-indexed)
function lines(start, end) {
    return src.split('\n').slice(start - 1, end).join('\n');
}

// ═══ services.js: All getXxxStats functions ═══
const servicesContent = `/**
 * services.js — Data processing and statistics functions
 */
import { parseCurrency, formatCurrency, normalizeCountry, isCountryMatch, sortCountriesByCount } from './utils.js';

${lines(735, 775)}

${lines(882, 947)}

${lines(1202, 1251)}

${lines(1500, 1534)}

${lines(1579, 1616)}

${lines(1651, 1667)}

${lines(1739, 1899)}

${lines(2161, 2193)}

${lines(2274, 2292)}

${lines(435, 458)}
`;

// Add export to each function
const servicesWithExports = servicesContent
    .replace(/^function getOrderSheetStats/m, 'export function getOrderSheetStats')
    .replace(/^function getPipelineStats/m, 'export function getPipelineStats')
    .replace(/^function getPartnerStats/m, 'export function getPartnerStats')
    .replace(/^function getGenericCountryStats/m, 'export function getGenericCountryStats')
    .replace(/^function getExpiringContractsStats/m, 'export function getExpiringContractsStats')
    .replace(/^function getPartnerPerformanceStats/m, 'export function getPartnerPerformanceStats')
    .replace(/^function getPocStats/m, 'export function getPocStats')
    .replace(/^function getEventStats/m, 'export function getEventStats')
    .replace(/^function getCountrySpecificStats/m, 'export function getCountrySpecificStats')
    .replace(/^function getServiceAnalysisStats/m, 'export function getServiceAnalysisStats')
    // Remove workbookData references - they'll be passed as params
    ;

fs.writeFileSync('services.js', servicesWithExports, 'utf8');
console.log('✅ services.js created');

// ═══ ui.js: All getXxxHTML functions + style injection ═══
const uiContent = `/**
 * ui.js — HTML template generators for dashboard components
 */
import { formatCurrency, parseCurrency } from './utils.js';
import { CONFIG } from './config.js';

${lines(405, 433)}

${lines(460, 496)}

${lines(572, 608)}

${lines(777, 806)}

${lines(949, 1083)}

${lines(1253, 1402)}

${lines(1536, 1566)}

${lines(1618, 1639)}

${lines(1669, 1682)}

${lines(1901, 2034)}

${lines(2195, 2223)}

${lines(2294, 2314)}
`;

const uiWithExports = uiContent
    .replace(/^function injectServiceAnalysisStyles/m, 'export function injectServiceAnalysisStyles')
    .replace(/^function getServiceAnalysisHTML/m, 'export function getServiceAnalysisHTML')
    .replace(/^function getRenewalHTML/m, 'export function getRenewalHTML')
    .replace(/^function getOrderSheetHTML/m, 'export function getOrderSheetHTML')
    .replace(/^function getPipelineHTML/m, 'export function getPipelineHTML')
    .replace(/^function getPartnerHTML/m, 'export function getPartnerHTML')
    .replace(/^function getGenericCountryHTML/m, 'export function getGenericCountryHTML')
    .replace(/^function getExpiringContractsHTML/m, 'export function getExpiringContractsHTML')
    .replace(/^function getPartnerPerformanceHTML/m, 'export function getPartnerPerformanceHTML')
    .replace(/^function getPocHTML/m, 'export function getPocHTML')
    .replace(/^function getEventHTML/m, 'export function getEventHTML')
    .replace(/^function getCountrySpecificHTML/m, 'export function getCountrySpecificHTML')
    ;

fs.writeFileSync('ui.js', uiWithExports, 'utf8');
console.log('✅ ui.js created');

console.log('Done! Now create app.js manually as the entry point.');
