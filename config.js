/**
 * config.js — Global configuration constants
 */

// Google Drive file IDs — replace with your actual file IDs after uploading to Drive
// To get the ID: Share link looks like https://drive.google.com/file/d/FILE_ID/view
export const DATA_SOURCES = {
    MAIN_FILE_ID: 'YOUR_MAIN_FILE_ID_HERE',      // 2026 Global Rev.01.xlsx
    MRR_FILE_ID: 'YOUR_MRR_FILE_ID_HERE',         // Global MRR ARR.xlsx
};

export const CONFIG = {
    COUNTRIES: ['Indonesia', 'Thailand', 'Malaysia', 'USA', 'Philippines', 'Vietnam', 'Singapore', 'Turkey'],
    COLORS: [
        '#6366f1', '#0ea5e9', '#10b981', '#f59e0b', '#ef4444',
        '#a855f7', '#ec4899', '#14b8a6', '#f97316', '#84cc16'
    ],
    CHART_DEFAULTS: {
        font: "'Inter', sans-serif",
        gridColor: 'rgba(0,0,0,0.05)'
    }
};
