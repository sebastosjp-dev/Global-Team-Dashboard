/**
 * config.js — Global configuration constants
 */

export const AUTH = {
    PASSWORD_ENCODED: 'MDEyNA==',  // btoa('0124')
};

// Google Sheets spreadsheet IDs
export const DATA_SOURCES = {
    MAIN_SHEET_ID: '1GLisAXT8E8a4Pqrt_nbQHupb7QkXrztWEgK_8YwyfWg',  // 2026 Global Rev.01
    MRR_SHEET_ID:  '1CJpEY65WBfQoSCwfBVfC4KT7Q2IQaJyXXjcsoxGZFlM',  // Global MRR ARR
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
