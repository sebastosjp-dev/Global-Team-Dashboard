/**
 * sidebar.js — Sidebar navigation building and tab selection UI logic.
 * Extracted from app.js for single-responsibility.
 * @module sidebar
 */
import { CONFIG } from './config.js';

/**
 * Build the sidebar navigation from sheet names.
 * @param {string[]} sheetNames - Excel sheet names
 * @param {Object} callbacks - Event handlers
 * @param {function(string, string|null):void} callbacks.onSelectTab
 * @param {function():void} callbacks.onSelectKPI
 * @param {function():void} callbacks.onSelectTraining
 * @param {function():void} callbacks.onSelectTcvArr
 */
export function buildSidebar(sheetNames, callbacks) {
    const sidebarNav = document.getElementById('sidebar-nav');
    if (!sidebarNav) return;
    sidebarNav.innerHTML = '';

    sheetNames.forEach(name => {
        // Updated exclusion list to include "Staff Training" sheet while keeping the dashboard menu
        if (name.includes('Global(Contract Date)') || ['Sheet9', 'Sheet10', 'Sheet12', 'Sheet13', 'Sheet16', '2026 Q1 Review', 'Staff Training', 'FEEDBACK', 'Weekly draft'].includes(name)) return;

        if (name === 'ORDER SHEET') {
            _buildExpandableNav(sidebarNav, name, callbacks.onSelectTab);
        } else {
            _buildSimpleNav(sidebarNav, name, callbacks.onSelectTab);
        }
    });

    _buildKPINav(sidebarNav, callbacks.onSelectKPI);
    _buildTcvArrNav(sidebarNav, callbacks.onSelectTcvArr);
    _buildTrainingNav(sidebarNav, callbacks.onSelectTraining);
}

/* ═══════════════════════════════════════════════════════════════
   Private Builders
   ═══════════════════════════════════════════════════════════════ */

/**
 * Build an expandable nav item with country sub-items.
 * @param {HTMLElement} container
 * @param {string} name
 * @param {function(string, string|null):void} onSelect
 */
function _buildExpandableNav(container, name, onSelect) {
    const parentItem = document.createElement('div');
    parentItem.className = 'nav-item nav-item-parent';
    parentItem.innerHTML = `<div style="display:flex; align-items:center; gap:12px;"><i class="fa-solid fa-folder"></i> <span>${name}</span></div><i class="fa-solid fa-chevron-down toggle-icon" style="font-size: 0.8em; transition: transform 0.3s;"></i>`;

    const subList = document.createElement('div');
    subList.className = 'nav-sublist';

    const allItem = document.createElement('div');
    allItem.className = 'nav-item sub-item';
    allItem.innerHTML = `<i class="fa-solid fa-earth-americas"></i> <span>All</span>`;
    allItem.onclick = (e) => { e.stopPropagation(); onSelect(name, null); };
    subList.appendChild(allItem);

    CONFIG.COUNTRIES.forEach(country => {
        const subItem = document.createElement('div');
        subItem.className = 'nav-item sub-item';
        subItem.innerHTML = `<i class="fa-solid fa-location-dot"></i> <span>${country}</span>`;
        subItem.onclick = (e) => { e.stopPropagation(); onSelect(name, country); };
        subList.appendChild(subItem);
    });

    parentItem.onclick = () => {
        subList.classList.toggle('expanded');
        const icon = parentItem.querySelector('.toggle-icon');
        icon.style.transform = subList.classList.contains('expanded') ? 'rotate(180deg)' : 'rotate(0deg)';
        onSelect(name, null);
    };

    container.appendChild(parentItem);
    container.appendChild(subList);
}

/**
 * Build a simple (flat) nav item.
 * @param {HTMLElement} container
 * @param {string} name
 * @param {function(string, string|null):void} onSelect
 */
function _buildSimpleNav(container, name, onSelect) {
    const navItem = document.createElement('div');
    navItem.className = 'nav-item';
    const icon = name === 'EVENT' ? 'fa-calendar-check' : 'fa-folder';
    navItem.innerHTML = `<i class="fa-solid ${icon}"></i> <span>${name}</span>`;
    navItem.onclick = () => onSelect(name, null);
    container.appendChild(navItem);
}

/**
 * Build the KPI Goals nav item at the bottom.
 * @param {HTMLElement} container
 * @param {function():void} onSelectKPI
 */
function _buildKPINav(container, onSelectKPI) {
    const kpiItem = document.createElement('div');
    kpiItem.className = 'nav-item kpi-tab';
    kpiItem.style.marginTop = '10px';
    kpiItem.style.borderTop = '1px solid rgba(255,255,255,0.1)';
    kpiItem.style.paddingTop = '15px';
    kpiItem.innerHTML = `<i class="fa-solid fa-bullseye" style="color: #ef4444;"></i> <span style="font-weight: 700;">KPI GOALS</span>`;
    kpiItem.onclick = () => onSelectKPI();
    container.appendChild(kpiItem);
}

/**
 * Build the TCV vs ARR nav item.
 * @param {HTMLElement} container
 * @param {function():void} onSelectTcvArr
 */
function _buildTcvArrNav(container, onSelectTcvArr) {
    const tcvArrItem = document.createElement('div');
    tcvArrItem.className = 'nav-item tcv-arr-tab';
    tcvArrItem.innerHTML = `<i class="fa-solid fa-chart-column" style="color: #f59e0b;"></i> <span style="font-weight: 700;">TCV vs ARR</span>`;
    tcvArrItem.onclick = () => onSelectTcvArr();
    container.appendChild(tcvArrItem);
}

/**
 * Build the Staff Training nav item.
 * @param {HTMLElement} container
 * @param {function():void} onSelectTraining
 */
function _buildTrainingNav(container, onSelectTraining) {
    const trainingItem = document.createElement('div');
    trainingItem.className = 'nav-item training-tab';
    trainingItem.innerHTML = `<i class="fa-solid fa-graduation-cap" style="color: #0ea5e9;"></i> <span style="font-weight: 700;">STAFF TRAINING</span>`;
    trainingItem.onclick = () => onSelectTraining();
    container.appendChild(trainingItem);
}
