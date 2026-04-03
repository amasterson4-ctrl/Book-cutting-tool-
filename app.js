// ============================
// RevOps Book Balancer - Core App
// ============================

let rawData = [];
let parsedAccounts = [];
let analysisResult = null;
let rebalanceResult = null;
let charts = {};

// ---------- SEGMENT CLASSIFICATION ----------
const ENT_CS_SEGMENTS = ['CS - Enterprise', 'CS - Enterprise Plus', 'CS - Strategic'];
const MM_CS_SEGMENTS = ['CS - Mid-Market', 'CS - SMB', 'CS - New Verticals'];
const INTL_CS_SEGMENT = 'CS - International';
const PARTNERSHIPS_SEGMENT = 'Partnerships';

const ENT_AM_SEGMENTS = ['AM - Enterprise', 'AD - Strategic'];
const MM_AM_SEGMENTS = ['AM - MM', 'AM - SMB'];
const INTL_AM_SEGMENT = 'AM - International';

function getSegmentGroup(csSegment, amSegment) {
    if (ENT_CS_SEGMENTS.includes(csSegment) || ENT_AM_SEGMENTS.includes(amSegment)) return 'enterprise';
    if (MM_CS_SEGMENTS.includes(csSegment) || MM_AM_SEGMENTS.includes(amSegment)) return 'mm_smb';
    if (csSegment === INTL_CS_SEGMENT || amSegment === INTL_AM_SEGMENT) return 'international';
    if (csSegment === PARTNERSHIPS_SEGMENT) return 'partnerships';
    return 'other';
}

// ---------- PARSING ----------
function parseCurrency(val) {
    if (!val || val === '—' || val === '-') return 0;
    return parseFloat(String(val).replace(/[$,]/g, '')) || 0;
}

function parsePercent(val) {
    if (!val || val === '—' || val === '-') return null;
    return parseFloat(String(val).replace('%', '')) / 100;
}

function parseDate(val) {
    if (!val || val === '—' || val === '-') return null;
    const d = new Date(val);
    return isNaN(d) ? null : d;
}

function parseRow(row) {
    const products = [];
    if (String(row['Uses SMS']).toLowerCase() === 'yes') products.push('SMS');
    if (String(row['Uses Email']).toLowerCase() === 'yes') products.push('Email');
    if (String(row['Uses AI']).toLowerCase() === 'yes') products.push('AI');

    const csSegment = (row['CS Segment'] || '').trim();
    const amSegment = (row['Account Management Segment'] || '').trim();

    return {
        parentAccountId: (row['Parent Account ID'] || '').trim(),
        accountId: (row['Account ID'] || '').trim(),
        accountType: (row['Account Type'] || '').trim(),
        renewalDate: parseDate(row['Renewal Date']),
        customerTier: (row['Customer Tier'] || '').trim(),
        arr: parseCurrency(row['Total L12M Revenue']),
        whitespace: parseCurrency(row['Total Estimated Whitespace']),
        countryCode: (row['Country Code'] || '').trim(),
        products,
        adoptionDepth: products.length,
        csSegment,
        csmId: (row['CSM ID'] || '').trim(),
        salesOwnerType: (row['Sales Owner Type'] || '').trim(),
        amSegment,
        salesOwnerId: (row['Sales Owner ID'] || '').trim(),
        healthScore: parsePercent(row['Health Score']),
        segmentGroup: getSegmentGroup(csSegment, amSegment),
        isChild: (row['Parent Account ID'] || '').trim() !== (row['Account ID'] || '').trim(),
        // Will be filled during rebalance:
        newCsmId: null,
        newSalesOwnerId: null,
        newCsSegment: null,
        newAmSegment: null,
    };
}

// ---------- UPLOAD ----------
const dropZone = document.getElementById('drop-zone');
const fileInput = document.getElementById('file-input');

dropZone.addEventListener('click', () => fileInput.click());
dropZone.addEventListener('dragover', e => { e.preventDefault(); dropZone.classList.add('dragover'); });
dropZone.addEventListener('dragleave', () => dropZone.classList.remove('dragover'));
dropZone.addEventListener('drop', e => {
    e.preventDefault();
    dropZone.classList.remove('dragover');
    if (e.dataTransfer.files.length) handleFile(e.dataTransfer.files[0]);
});
fileInput.addEventListener('change', e => { if (e.target.files.length) handleFile(e.target.files[0]); });

function handleFile(file) {
    Papa.parse(file, {
        header: true,
        skipEmptyLines: true,
        complete(results) {
            rawData = results.data;
            parsedAccounts = rawData.map(parseRow).filter(a => a.accountId);
            showUploadSuccess(parsedAccounts.length);
            // Automatically run analysis since weights are already configured
            runAnalysis();
        },
        error(err) {
            showUploadError(err.message);
        }
    });
}

function showUploadSuccess(count) {
    const el = document.getElementById('upload-status');
    el.classList.remove('hidden');
    document.getElementById('upload-msg').textContent =
        `Successfully loaded ${count} accounts. Analysis running with current weights.`;
    el.querySelector('.status-bar').className = 'status-bar success';
}

function showUploadError(msg) {
    const el = document.getElementById('upload-status');
    el.classList.remove('hidden');
    document.getElementById('upload-msg').textContent = `Error: ${msg}`;
    el.querySelector('.status-bar').className = 'status-bar error';
}

// ---------- WEIGHT SLIDERS ----------
['arr', 'whitespace', 'renewal', 'adoption', 'parent', 'geo', 'health'].forEach(key => {
    const slider = document.getElementById(`w-${key}`);
    const valEl = document.getElementById(`w-${key}-val`);
    slider.addEventListener('input', () => { valEl.textContent = slider.value; });
});

// ---------- CSM:AM RATIO SLIDERS ----------
// Slider value 1-5: 1=3:1 CSM:AM, 3=1:1, 5=1:3 AM:CSM
const RATIO_MAP = {
    1: '3 : 1', 2: '2 : 1',
    3: '1 : 1',
    4: '1 : 2', 5: '1 : 3'
};
const RATIO_NUMERIC = {
    1: 3, 2: 2, 3: 1, 4: 0.5, 5: 1/3
};

function setupRatioSlider(prefix) {
    const slider = document.getElementById(`${prefix}-ratio`);
    const valEl = document.getElementById(`${prefix}-ratio-val`);
    if (!slider || !valEl) return;
    const update = () => { valEl.textContent = RATIO_MAP[slider.value] || '1 : 1'; };
    slider.addEventListener('input', update);
    update();
}
setupRatioSlider('ent');
setupRatioSlider('mm');

function getRatios() {
    return {
        ent: RATIO_NUMERIC[+document.getElementById('ent-ratio').value] || 1,
        mm: RATIO_NUMERIC[+document.getElementById('mm-ratio').value] || 1,
    };
}

function getWeights() {
    return {
        arr: +document.getElementById('w-arr').value / 100,
        whitespace: +document.getElementById('w-whitespace').value / 100,
        renewal: +document.getElementById('w-renewal').value / 100,
        adoption: +document.getElementById('w-adoption').value / 100,
        parent: +document.getElementById('w-parent').value / 100,
        geo: +document.getElementById('w-geo').value / 100,
        health: +document.getElementById('w-health').value / 100,
    };
}

function getBookSizeTargets() {
    return {
        entCsm: +document.getElementById('ent-csm-book').value,
        entAm: +document.getElementById('ent-am-book').value,
        mmCsm: +document.getElementById('mm-csm-book').value,
        mmAm: +document.getElementById('mm-am-book').value,
    };
}

// ---------- SCOPE CONTROLS ----------
function getSegmentScope() {
    return document.getElementById('segment-scope').value; // 'both', 'enterprise', 'mm_smb'
}

function getGroupingMode() {
    return document.getElementById('grouping-mode').value; // 'standard', 'parent-child'
}

function filterByScope(accounts) {
    const scope = getSegmentScope();
    if (scope === 'both') return accounts;
    return accounts.filter(a => a.segmentGroup === scope ||
        (scope === 'enterprise' && a.segmentGroup === 'international') // intl rolls into enterprise
    );
}

// Re-run analysis when scope or grouping changes (if data is loaded)
document.getElementById('segment-scope').addEventListener('change', () => {
    if (parsedAccounts.length > 0) {
        // Hide previous results since scope changed
        document.getElementById('results-section').classList.add('hidden');
        runAnalysis();
    }
});
document.getElementById('grouping-mode').addEventListener('change', () => {
    if (parsedAccounts.length > 0) {
        document.getElementById('results-section').classList.add('hidden');
        runAnalysis();
    }
});

// ---------- TABS ----------
document.querySelectorAll('.tab').forEach(tab => {
    tab.addEventListener('click', () => {
        const section = tab.closest('section');
        section.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
        section.querySelectorAll('.tab-content').forEach(tc => tc.classList.remove('active'));
        tab.classList.add('active');
        document.getElementById(`tab-${tab.dataset.tab}`).classList.add('active');
    });
});

// ---------- ANALYSIS ----------
document.getElementById('rebalance-btn').addEventListener('click', runRebalance);

function runAnalysis() {
    const scoped = filterByScope(parsedAccounts);
    analysisResult = analyzeBooks(scoped);
    renderAnalysis(analysisResult);
    document.getElementById('analysis-section').classList.remove('hidden');
    document.getElementById('analysis-section').scrollIntoView({ behavior: 'smooth' });
}

function analyzeBooks(accounts) {
    // Group by CSM and AM
    const csmBooks = {};
    const amBooks = {};
    const parentGroups = {};

    accounts.forEach(a => {
        // CSM books
        if (a.csmId) {
            if (!csmBooks[a.csmId]) csmBooks[a.csmId] = { id: a.csmId, segment: a.csSegment, accounts: [] };
            csmBooks[a.csmId].accounts.push(a);
        }
        // AM books
        if (a.salesOwnerId && a.salesOwnerType === 'Expansion AM/AD') {
            if (!amBooks[a.salesOwnerId]) amBooks[a.salesOwnerId] = { id: a.salesOwnerId, segment: a.amSegment, accounts: [] };
            amBooks[a.salesOwnerId].accounts.push(a);
        }
        // Parent groups
        if (a.parentAccountId) {
            if (!parentGroups[a.parentAccountId]) parentGroups[a.parentAccountId] = [];
            parentGroups[a.parentAccountId].push(a);
        }
    });

    // Compute book-level metrics
    const computeBookMetrics = (book) => {
        const accts = book.accounts;
        const totalArr = accts.reduce((s, a) => s + a.arr, 0);
        const totalWhitespace = accts.reduce((s, a) => s + a.whitespace, 0);
        const healthScores = accts.filter(a => a.healthScore !== null).map(a => a.healthScore);
        const avgHealth = healthScores.length ? healthScores.reduce((s, h) => s + h, 0) / healthScores.length : null;
        const lowHealthCount = healthScores.filter(h => h <= 0.25).length;
        const totalAdoption = accts.reduce((s, a) => s + a.adoptionDepth, 0);
        const avgAdoption = accts.length ? totalAdoption / accts.length : 0;

        // Renewal distribution by month
        const renewalsByMonth = {};
        accts.forEach(a => {
            if (a.renewalDate) {
                const key = `${a.renewalDate.getFullYear()}-${String(a.renewalDate.getMonth() + 1).padStart(2, '0')}`;
                if (!renewalsByMonth[key]) renewalsByMonth[key] = { count: 0, arr: 0 };
                renewalsByMonth[key].count++;
                renewalsByMonth[key].arr += a.arr;
            }
        });

        // Geo distribution
        const geos = {};
        accts.forEach(a => { geos[a.countryCode] = (geos[a.countryCode] || 0) + 1; });
        const intlCount = accts.filter(a => a.countryCode && a.countryCode !== 'US').length;

        // Parent-child integrity: check if any parent group is split across owners
        const parentIds = [...new Set(accts.filter(a => a.isChild).map(a => a.parentAccountId))];

        return {
            ...book,
            count: accts.length,
            totalArr,
            totalWhitespace,
            avgHealth,
            lowHealthCount,
            avgAdoption,
            renewalsByMonth,
            geos,
            intlCount,
            parentIds,
        };
    };

    const csmMetrics = Object.values(csmBooks).map(computeBookMetrics);
    const amMetrics = Object.values(amBooks).map(computeBookMetrics);

    // Find parent-child splits
    const parentChildSplits = [];
    Object.entries(parentGroups).forEach(([parentId, children]) => {
        if (children.length > 1) {
            const csmIds = [...new Set(children.map(c => c.csmId).filter(Boolean))];
            const amIds = [...new Set(children.map(c => c.salesOwnerId).filter(Boolean))];
            if (csmIds.length > 1 || amIds.length > 1) {
                parentChildSplits.push({ parentId, children, csmIds, amIds });
            }
        }
    });

    // Segment breakdown
    const segments = {};
    accounts.forEach(a => {
        if (!segments[a.segmentGroup]) segments[a.segmentGroup] = { count: 0, arr: 0 };
        segments[a.segmentGroup].count++;
        segments[a.segmentGroup].arr += a.arr;
    });

    return { csmMetrics, amMetrics, parentChildSplits, parentGroups, segments, accounts };
}

// ---------- RENDER ANALYSIS ----------
function fmt$(val) { return '$' + Math.round(val).toLocaleString(); }
function fmtPct(val) { return val !== null ? Math.round(val * 100) + '%' : 'N/A'; }

function stdDev(arr) {
    if (arr.length < 2) return 0;
    const mean = arr.reduce((s, v) => s + v, 0) / arr.length;
    return Math.sqrt(arr.reduce((s, v) => s + (v - mean) ** 2, 0) / arr.length);
}

function cv(arr) {
    const mean = arr.reduce((s, v) => s + v, 0) / arr.length;
    if (mean === 0) return 0;
    return stdDev(arr) / mean;
}

function getScopeLabel() {
    const scope = getSegmentScope();
    if (scope === 'enterprise') return 'Enterprise & Strategic';
    if (scope === 'mm_smb') return 'Mid-Market & SMB';
    return 'All Segments';
}

function renderAnalysis(result) {
    const { csmMetrics, amMetrics, parentChildSplits, segments } = result;
    const scoped = filterByScope(parsedAccounts);

    // Overview metrics
    const totalAccounts = scoped.length;
    const totalArr = scoped.reduce((s, a) => s + a.arr, 0);
    const totalWhitespace = scoped.reduce((s, a) => s + a.whitespace, 0);
    const csmArrValues = csmMetrics.map(b => b.totalArr);
    const amArrValues = amMetrics.map(b => b.totalArr);

    const scopeLabel = getScopeLabel();
    const groupLabel = getGroupingMode() === 'parent-child' ? ' | Parent-Child Priority' : '';
    const metricsHtml = `
        <div class="metric-card scope-badge-card"><div class="metric-label">Scope</div><div class="metric-value scope-badge-value">${scopeLabel}${groupLabel}</div></div>
        <div class="metric-card"><div class="metric-label">Total Accounts</div><div class="metric-value">${totalAccounts}</div></div>
        <div class="metric-card"><div class="metric-label">Total ARR</div><div class="metric-value">${fmt$(totalArr)}</div></div>
        <div class="metric-card"><div class="metric-label">Total Whitespace</div><div class="metric-value">${fmt$(totalWhitespace)}</div></div>
        <div class="metric-card"><div class="metric-label">CSM Count</div><div class="metric-value">${csmMetrics.length}</div><div class="metric-sub">ARR CV: ${(cv(csmArrValues) * 100).toFixed(1)}%</div></div>
        <div class="metric-card"><div class="metric-label">AM Count</div><div class="metric-value">${amMetrics.length}</div><div class="metric-sub">ARR CV: ${(cv(amArrValues) * 100).toFixed(1)}%</div></div>
        <div class="metric-card"><div class="metric-label">Parent-Child Splits</div><div class="metric-value ${parentChildSplits.length > 0 ? 'cell-bad' : 'cell-good'}">${parentChildSplits.length}</div><div class="metric-sub">families split across owners</div></div>
    `;
    document.getElementById('overview-metrics').innerHTML = metricsHtml;

    // Imbalance flags
    renderFlags(result);

    // Charts
    renderBookChart('chart-csm-arr', 'CSM Books - ARR Distribution', csmMetrics, 'totalArr', fmt$);
    renderBookChart('chart-am-arr', 'AM Books - ARR Distribution', amMetrics, 'totalArr', fmt$);
    renderBookChart('chart-csm-health', 'CSM Books - Avg Health Score', csmMetrics, 'avgHealth', v => fmtPct(v));
    renderBookChart('chart-am-whitespace', 'AM Books - Whitespace', amMetrics, 'totalWhitespace', fmt$);

    // CSM detail table
    renderBookTable('csm-table-wrap', csmMetrics, 'CSM');
    renderBookTable('am-table-wrap', amMetrics, 'AM');

    // Renewal view
    renderRenewalView(result);

    // Geo view
    renderGeoView(result);
}

function renderBookChart(canvasId, title, books, metric, formatter) {
    if (charts[canvasId]) charts[canvasId].destroy();

    const sorted = [...books].sort((a, b) => (b[metric] || 0) - (a[metric] || 0));
    const labels = sorted.map(b => b.id);
    const data = sorted.map(b => b[metric] || 0);
    const mean = data.length ? data.reduce((s, v) => s + v, 0) / data.length : 0;

    const colors = data.map(v => {
        const pctOff = mean > 0 ? Math.abs(v - mean) / mean : 0;
        if (pctOff > 0.5) return 'rgba(220, 38, 38, 0.7)';
        if (pctOff > 0.25) return 'rgba(217, 119, 6, 0.7)';
        return 'rgba(37, 99, 235, 0.7)';
    });

    charts[canvasId] = new Chart(document.getElementById(canvasId), {
        type: 'bar',
        data: {
            labels,
            datasets: [{
                label: title,
                data,
                backgroundColor: colors,
                borderRadius: 4,
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                title: { display: true, text: title, font: { size: 13, weight: '600' } },
                legend: { display: false },
                tooltip: {
                    callbacks: {
                        label: ctx => formatter(ctx.parsed.y)
                    }
                },
                annotation: undefined,
            },
            scales: {
                x: { ticks: { font: { size: 9 }, maxRotation: 45 } },
                y: { ticks: { callback: v => formatter(v) } }
            }
        }
    });
}

function renderFlags(result) {
    const { csmMetrics, amMetrics, parentChildSplits } = result;
    const weights = getWeights();
    const flags = [];

    // ARR imbalance
    const csmArrCv = cv(csmMetrics.map(b => b.totalArr));
    const amArrCv = cv(amMetrics.map(b => b.totalArr));
    if (csmArrCv > 0.4) flags.push({ type: 'danger', msg: `CSM ARR highly imbalanced (CV ${(csmArrCv * 100).toFixed(1)}%). Some CSMs have significantly more revenue than others.` });
    else if (csmArrCv > 0.25) flags.push({ type: 'warning', msg: `CSM ARR moderately imbalanced (CV ${(csmArrCv * 100).toFixed(1)}%).` });
    if (amArrCv > 0.4) flags.push({ type: 'danger', msg: `AM ARR highly imbalanced (CV ${(amArrCv * 100).toFixed(1)}%). Expansion opportunity is uneven.` });
    else if (amArrCv > 0.25) flags.push({ type: 'warning', msg: `AM ARR moderately imbalanced (CV ${(amArrCv * 100).toFixed(1)}%).` });

    // Whitespace imbalance
    const amWsCv = cv(amMetrics.map(b => b.totalWhitespace));
    if (amWsCv > 0.4) flags.push({ type: 'danger', msg: `AM Whitespace highly imbalanced (CV ${(amWsCv * 100).toFixed(1)}%). TAM opportunity is not equal.` });

    // Health imbalance
    const lowHealthBooks = csmMetrics.filter(b => b.lowHealthCount >= 3);
    if (lowHealthBooks.length) {
        flags.push({ type: 'warning', msg: `${lowHealthBooks.length} CSM book(s) have 3+ low-health accounts: ${lowHealthBooks.map(b => b.id).join(', ')}` });
    }

    // Parent-child splits
    if (parentChildSplits.length > 0) {
        flags.push({ type: 'danger', msg: `${parentChildSplits.length} parent account families are split across different CSMs/AMs. Pod structure is broken.` });
    }

    // Renewal concentration
    csmMetrics.forEach(b => {
        const months = Object.entries(b.renewalsByMonth);
        const maxMonth = months.reduce((max, [k, v]) => v.arr > (max?.arr || 0) ? { month: k, ...v } : max, null);
        if (maxMonth && b.totalArr > 0 && maxMonth.arr / b.totalArr > 0.4) {
            flags.push({ type: 'warning', msg: `${b.id}: ${Math.round(maxMonth.arr / b.totalArr * 100)}% of ARR renews in ${maxMonth.month}. Heavy renewal concentration.` });
        }
    });

    const container = document.getElementById('imbalance-flags');
    if (flags.length === 0) {
        container.innerHTML = '<div class="flag info"><span class="flag-icon">&#10003;</span> No major imbalances detected in the current book structure.</div>';
    } else {
        container.innerHTML = flags.map(f =>
            `<div class="flag ${f.type}"><span class="flag-icon">${f.type === 'danger' ? '!' : '&#9888;'}</span> ${f.msg}</div>`
        ).join('');
    }
}

function renderBookTable(wrapperId, books, type) {
    const sorted = [...books].sort((a, b) => b.totalArr - a.totalArr);
    const rows = sorted.map(b => `
        <tr>
            <td>${b.id}</td>
            <td>${b.segment}</td>
            <td class="cell-num">${b.count}</td>
            <td class="cell-num">${fmt$(b.totalArr)}</td>
            <td class="cell-num">${fmt$(b.totalWhitespace)}</td>
            <td class="cell-num ${b.avgHealth !== null && b.avgHealth < 0.4 ? 'cell-bad' : ''}">${fmtPct(b.avgHealth)}</td>
            <td class="cell-num">${b.lowHealthCount}</td>
            <td class="cell-num">${b.avgAdoption.toFixed(1)}</td>
            <td class="cell-num">${b.intlCount}</td>
            <td class="cell-num">${b.parentIds.length}</td>
        </tr>
    `).join('');

    document.getElementById(wrapperId).innerHTML = `
        <div class="data-table-wrap">
            <table class="data-table">
                <thead><tr>
                    <th>${type} ID</th><th>Segment</th><th>Accts</th><th>ARR</th><th>Whitespace</th>
                    <th>Avg Health</th><th>Low Health</th><th>Avg Adoption</th><th>Intl Accts</th><th>Parent Groups</th>
                </tr></thead>
                <tbody>${rows}</tbody>
            </table>
        </div>
    `;
}

function renderRenewalView(result) {
    // Aggregate renewal ARR by month across all CSM books
    const allMonths = {};
    result.csmMetrics.forEach(b => {
        Object.entries(b.renewalsByMonth).forEach(([month, data]) => {
            if (!allMonths[month]) allMonths[month] = {};
            allMonths[month][b.id] = data.arr;
        });
    });

    const months = Object.keys(allMonths).sort();
    const csmIds = result.csmMetrics.map(b => b.id);
    const palette = generatePalette(csmIds.length);

    const datasets = csmIds.map((id, i) => ({
        label: id,
        data: months.map(m => (allMonths[m] && allMonths[m][id]) || 0),
        backgroundColor: palette[i],
    }));

    if (charts['chart-renewal-heatmap']) charts['chart-renewal-heatmap'].destroy();
    charts['chart-renewal-heatmap'] = new Chart(document.getElementById('chart-renewal-heatmap'), {
        type: 'bar',
        data: { labels: months, datasets },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                title: { display: true, text: 'Renewal ARR by Month (Stacked by CSM)', font: { size: 13, weight: '600' } },
                legend: { position: 'bottom', labels: { font: { size: 9 } } },
                tooltip: { callbacks: { label: ctx => `${ctx.dataset.label}: ${fmt$(ctx.parsed.y)}` } },
            },
            scales: {
                x: { stacked: true },
                y: { stacked: true, ticks: { callback: v => fmt$(v) } }
            }
        }
    });

    // Renewal flags
    const renewalFlags = [];
    result.csmMetrics.forEach(b => {
        const quarters = {};
        Object.entries(b.renewalsByMonth).forEach(([m, data]) => {
            const [y, mo] = m.split('-').map(Number);
            const q = `${y}-Q${Math.ceil(mo / 3)}`;
            if (!quarters[q]) quarters[q] = 0;
            quarters[q] += data.arr;
        });
        Object.entries(quarters).forEach(([q, arr]) => {
            if (b.totalArr > 0 && arr / b.totalArr > 0.5) {
                renewalFlags.push({ type: 'warning', msg: `${b.id}: ${Math.round(arr / b.totalArr * 100)}% of book ARR renews in ${q}` });
            }
        });
    });

    document.getElementById('renewal-flags').innerHTML = renewalFlags.length
        ? renewalFlags.map(f => `<div class="flag ${f.type}"><span class="flag-icon">&#9888;</span> ${f.msg}</div>`).join('')
        : '<div class="flag info"><span class="flag-icon">&#10003;</span> No severe quarterly renewal concentration detected.</div>';
}

function renderGeoView(result) {
    const geoAccounts = {};
    parsedAccounts.forEach(a => {
        if (!geoAccounts[a.countryCode]) geoAccounts[a.countryCode] = [];
        geoAccounts[a.countryCode].push(a);
    });

    const geoCards = Object.entries(geoAccounts)
        .sort((a, b) => b[1].length - a[1].length)
        .map(([cc, accts]) => {
            const csms = [...new Set(accts.map(a => a.csmId).filter(Boolean))];
            const ams = [...new Set(accts.map(a => a.salesOwnerId).filter(Boolean))];
            const totalArr = accts.reduce((s, a) => s + a.arr, 0);
            return `
                <div class="geo-card">
                    <h4>${cc} (${accts.length} accounts)</h4>
                    <div class="geo-detail">ARR: ${fmt$(totalArr)}</div>
                    <div class="geo-detail">CSMs: ${csms.join(', ') || 'None'}</div>
                    <div class="geo-detail">AMs: ${ams.join(', ') || 'None'}</div>
                </div>
            `;
        }).join('');

    document.getElementById('geo-summary').innerHTML = `<div class="geo-grid">${geoCards}</div>`;
}

function generatePalette(n) {
    const colors = [];
    for (let i = 0; i < n; i++) {
        const h = (i * 360 / n) % 360;
        colors.push(`hsla(${h}, 60%, 55%, 0.75)`);
    }
    return colors;
}

// ============================
// REBALANCING ALGORITHM
// ============================
document.getElementById('rebalance-btn').addEventListener('click', runRebalance);

function runRebalance() {
    const weights = getWeights();
    const targets = getBookSizeTargets();
    const ratios = getRatios();
    const groupingMode = getGroupingMode();
    const scopedSource = filterByScope(parsedAccounts);
    const accounts = scopedSource.map(a => ({ ...a })); // shallow copy

    // Re-run analysis on current scope to ensure before/after comparison is accurate
    analysisResult = analyzeBooks(scopedSource);
    renderAnalysis(analysisResult);

    // In parent-child mode, boost parent weight dramatically
    if (groupingMode === 'parent-child') {
        weights.parent = 1.0; // max out parent-child affinity
    }

    // Step 1: Build parent groups (must stay together)
    const parentGroups = {};
    accounts.forEach(a => {
        const pid = a.parentAccountId || a.accountId;
        if (!parentGroups[pid]) parentGroups[pid] = [];
        parentGroups[pid].push(a);
    });

    // Step 2: Build atomic units (parent families)
    const units = Object.entries(parentGroups).map(([pid, members]) => {
        const totalArr = members.reduce((s, a) => s + a.arr, 0);
        const totalWhitespace = members.reduce((s, a) => s + a.whitespace, 0);
        const healthScores = members.filter(a => a.healthScore !== null).map(a => a.healthScore);
        const avgHealth = healthScores.length ? healthScores.reduce((s, h) => s + h, 0) / healthScores.length : 0.5;
        const avgAdoption = members.reduce((s, a) => s + a.adoptionDepth, 0) / members.length;
        const geos = [...new Set(members.map(a => a.countryCode))];
        const isIntl = geos.some(g => g && g !== 'US');
        const renewalMonths = {};
        members.forEach(a => {
            if (a.renewalDate) {
                const key = `${a.renewalDate.getFullYear()}-${String(a.renewalDate.getMonth() + 1).padStart(2, '0')}`;
                if (!renewalMonths[key]) renewalMonths[key] = 0;
                renewalMonths[key] += a.arr;
            }
        });

        // Determine segment from majority
        const segCounts = {};
        members.forEach(a => { segCounts[a.segmentGroup] = (segCounts[a.segmentGroup] || 0) + 1; });
        const segmentGroup = Object.entries(segCounts).sort((a, b) => b[1] - a[1])[0][0];

        return {
            parentId: pid,
            members,
            count: members.length,
            totalArr,
            totalWhitespace,
            avgHealth,
            avgAdoption,
            geos,
            isIntl,
            renewalMonths,
            segmentGroup,
            currentCsm: members[0].csmId,
            currentAm: members[0].salesOwnerId,
            currentCsSegment: members[0].csSegment,
            currentAmSegment: members[0].amSegment,
        };
    });

    // Step 3: Separate units by segment group for CSM assignment
    const entUnits = units.filter(u => u.segmentGroup === 'enterprise');
    const mmUnits = units.filter(u => u.segmentGroup === 'mm_smb');
    const intlUnits = units.filter(u => u.segmentGroup === 'international');
    const otherUnits = units.filter(u => !['enterprise', 'mm_smb', 'international'].includes(u.segmentGroup));

    // Get existing CSMs and AMs per segment (from scoped accounts)
    const entCsms = [...new Set(accounts.filter(a => ENT_CS_SEGMENTS.includes(a.csSegment)).map(a => a.csmId).filter(Boolean))];
    const mmCsms = [...new Set(accounts.filter(a => MM_CS_SEGMENTS.includes(a.csSegment)).map(a => a.csmId).filter(Boolean))];
    const intlCsms = [...new Set(accounts.filter(a => a.csSegment === INTL_CS_SEGMENT).map(a => a.csmId).filter(Boolean))];
    const otherCsms = [...new Set(accounts.filter(a => a.csSegment === PARTNERSHIPS_SEGMENT).map(a => a.csmId).filter(Boolean))];

    const entAms = [...new Set(accounts.filter(a => ENT_AM_SEGMENTS.includes(a.amSegment) && a.salesOwnerType === 'Expansion AM/AD').map(a => a.salesOwnerId).filter(Boolean))];
    const mmAms = [...new Set(accounts.filter(a => MM_AM_SEGMENTS.includes(a.amSegment) && a.salesOwnerType === 'Expansion AM/AD').map(a => a.salesOwnerId).filter(Boolean))];
    const intlAms = [...new Set(accounts.filter(a => a.amSegment === INTL_AM_SEGMENT && a.salesOwnerType === 'Expansion AM/AD').map(a => a.salesOwnerId).filter(Boolean))];

    // Step 4: Balance assignment using weighted scoring
    // csmAmRatio: CSM:AM ratio for this segment (e.g. 2 means 2 CSMs per 1 AM)
    // Used to scale target capacity: CSMs get fewer accounts when ratio > 1, AMs get fewer when ratio < 1
    function assignUnitsToOwners(unitList, ownerIds, role, csmAmRatio) {
        if (ownerIds.length === 0 || unitList.length === 0) return;

        // Compute capacity multiplier from ratio
        // ratio > 1 means more CSMs per AM → each CSM handles fewer accounts
        // ratio < 1 means more AMs per CSM → each AM handles fewer accounts
        const capacityMultiplier = role === 'CSM'
            ? 1 / Math.max(csmAmRatio, 0.33)  // Higher ratio → smaller CSM books
            : Math.max(csmAmRatio, 0.33);       // Higher ratio → larger AM books

        // Initialize buckets
        const buckets = {};
        ownerIds.forEach(id => {
            buckets[id] = { id, units: [], totalArr: 0, totalWhitespace: 0, healthSum: 0, healthCount: 0, renewalMonths: {}, adoptionSum: 0, count: 0, intlCount: 0, capacityMultiplier };
        });

        // Sort units: largest ARR first (greedy bin-packing)
        const sorted = [...unitList].sort((a, b) => b.totalArr - a.totalArr);

        sorted.forEach(unit => {
            // Score each bucket for this unit
            let bestOwner = null;
            let bestScore = Infinity;

            ownerIds.forEach(ownerId => {
                const bucket = buckets[ownerId];
                let score = 0;

                // Capacity from CSM:AM ratio — scale count penalty by multiplier
                const adjustedCount = bucket.count * (1 / bucket.capacityMultiplier);
                score += adjustedCount * 5000; // soft pressure to respect ratio-based capacity

                // ARR balance: prefer bucket with lowest ARR
                const arrAfter = bucket.totalArr + unit.totalArr;
                score += weights.arr * arrAfter;

                // Whitespace balance: prefer bucket with lowest whitespace
                const wsAfter = bucket.totalWhitespace + unit.totalWhitespace;
                score += weights.whitespace * wsAfter;

                // Health balance: penalize buckets already heavy with low-health
                if (unit.avgHealth < 0.4) {
                    const lowHealthRatio = bucket.healthCount > 0
                        ? (bucket.healthSum / bucket.healthCount) : 0.5;
                    score += weights.health * (1 - lowHealthRatio) * 100000;
                }

                // Renewal balance: penalize if this unit's renewal month already heavy
                Object.entries(unit.renewalMonths).forEach(([month, arr]) => {
                    const existing = bucket.renewalMonths[month] || 0;
                    score += weights.renewal * (existing + arr);
                });

                // Adoption balance
                score += weights.adoption * Math.abs(bucket.adoptionSum / Math.max(bucket.count, 1) - unit.avgAdoption) * 10000;

                // Geo: prefer keeping intl accounts with owners who already have intl
                if (unit.isIntl && bucket.intlCount === 0 && weights.geo > 0) {
                    score += weights.geo * 50000;
                }

                // Parent affinity: prefer current owner to minimize changes
                const currentOwner = role === 'CSM' ? unit.currentCsm : unit.currentAm;
                if (ownerId === currentOwner) {
                    score -= weights.parent * 100000; // strong bonus to keep in place
                }

                if (score < bestScore) {
                    bestScore = score;
                    bestOwner = ownerId;
                }
            });

            // Assign
            const bucket = buckets[bestOwner];
            bucket.units.push(unit);
            bucket.totalArr += unit.totalArr;
            bucket.totalWhitespace += unit.totalWhitespace;
            bucket.healthSum += unit.avgHealth * unit.count;
            bucket.healthCount += unit.count;
            bucket.adoptionSum += unit.avgAdoption * unit.count;
            bucket.count += unit.count;
            if (unit.isIntl) bucket.intlCount += unit.count;
            Object.entries(unit.renewalMonths).forEach(([m, arr]) => {
                bucket.renewalMonths[m] = (bucket.renewalMonths[m] || 0) + arr;
            });

            // Write back assignments
            unit.members.forEach(a => {
                if (role === 'CSM') {
                    a.newCsmId = bestOwner;
                    // Determine new CS segment based on segment group
                    if (unit.segmentGroup === 'enterprise') {
                        a.newCsSegment = a.csSegment; // keep existing enterprise sub-segment
                    } else {
                        a.newCsSegment = a.csSegment;
                    }
                } else {
                    a.newSalesOwnerId = bestOwner;
                    a.newAmSegment = a.amSegment;
                }
            });
        });

        return buckets;
    }

    // Run CSM assignment per segment (pass CSM:AM ratio)
    const csmEntBuckets = assignUnitsToOwners(entUnits, entCsms, 'CSM', ratios.ent);
    const csmMmBuckets = assignUnitsToOwners(mmUnits, mmCsms, 'CSM', ratios.mm);
    const csmIntlBuckets = assignUnitsToOwners(intlUnits, intlCsms.length ? intlCsms : entCsms, 'CSM', 1);
    const csmOtherBuckets = assignUnitsToOwners(otherUnits, otherCsms.length ? otherCsms : mmCsms, 'CSM', ratios.mm);

    // Run AM assignment per segment (pass CSM:AM ratio)
    const amEntBuckets = assignUnitsToOwners(entUnits, entAms, 'AM', ratios.ent);
    const amMmBuckets = assignUnitsToOwners(mmUnits, mmAms, 'AM', ratios.mm);
    const amIntlBuckets = assignUnitsToOwners(intlUnits, intlAms.length ? intlAms : entAms, 'AM', 1);

    // For accounts with New Business AE or Unassigned, keep as-is
    accounts.forEach(a => {
        if (!a.newCsmId) a.newCsmId = a.csmId;
        if (!a.newSalesOwnerId) a.newSalesOwnerId = a.salesOwnerId;
        if (!a.newCsSegment) a.newCsSegment = a.csSegment;
        if (!a.newAmSegment) a.newAmSegment = a.amSegment;
    });

    // Count changes
    const csmChanges = accounts.filter(a => a.newCsmId && a.newCsmId !== a.csmId);
    const amChanges = accounts.filter(a => a.newSalesOwnerId && a.newSalesOwnerId !== a.salesOwnerId && a.salesOwnerType === 'Expansion AM/AD');
    const totalChanges = accounts.filter(a =>
        (a.newCsmId && a.newCsmId !== a.csmId) ||
        (a.newSalesOwnerId && a.newSalesOwnerId !== a.salesOwnerId && a.salesOwnerType === 'Expansion AM/AD')
    );

    // Re-analyze with new assignments
    const newAccounts = accounts.map(a => ({
        ...a,
        csmId: a.newCsmId || a.csmId,
        salesOwnerId: a.newSalesOwnerId || a.salesOwnerId,
    }));
    const newAnalysis = analyzeBooks(newAccounts);

    rebalanceResult = {
        accounts,
        csmChanges,
        amChanges,
        totalChanges,
        beforeAnalysis: analysisResult,
        afterAnalysis: newAnalysis,
    };

    renderResults(rebalanceResult);
    document.getElementById('results-section').classList.remove('hidden');
    document.getElementById('results-section').scrollIntoView({ behavior: 'smooth' });
}

// ---------- RENDER RESULTS ----------
function renderResults(result) {
    const { accounts, csmChanges, amChanges, totalChanges, beforeAnalysis, afterAnalysis } = result;

    // Summary
    const beforeCsmCv = cv(beforeAnalysis.csmMetrics.map(b => b.totalArr));
    const afterCsmCv = cv(afterAnalysis.csmMetrics.map(b => b.totalArr));
    const beforeAmCv = cv(beforeAnalysis.amMetrics.map(b => b.totalArr));
    const afterAmCv = cv(afterAnalysis.amMetrics.map(b => b.totalArr));
    const beforeSplits = beforeAnalysis.parentChildSplits.length;
    const afterSplits = afterAnalysis.parentChildSplits.length;

    document.getElementById('results-summary').innerHTML = `
        <div class="metric-card">
            <div class="metric-label">Total Account Moves</div>
            <div class="metric-value">${totalChanges.length}</div>
            <div class="metric-sub">${csmChanges.length} CSM + ${amChanges.length} AM changes</div>
        </div>
        <div class="metric-card">
            <div class="metric-label">CSM ARR Balance (CV)</div>
            <div class="metric-value ${afterCsmCv < beforeCsmCv ? 'cell-good' : 'cell-warn'}">${(beforeCsmCv * 100).toFixed(1)}% &rarr; ${(afterCsmCv * 100).toFixed(1)}%</div>
            <div class="metric-sub">${afterCsmCv < beforeCsmCv ? 'Improved' : 'Check weights'}</div>
        </div>
        <div class="metric-card">
            <div class="metric-label">AM ARR Balance (CV)</div>
            <div class="metric-value ${afterAmCv < beforeAmCv ? 'cell-good' : 'cell-warn'}">${(beforeAmCv * 100).toFixed(1)}% &rarr; ${(afterAmCv * 100).toFixed(1)}%</div>
            <div class="metric-sub">${afterAmCv < beforeAmCv ? 'Improved' : 'Check weights'}</div>
        </div>
        <div class="metric-card">
            <div class="metric-label">Parent-Child Splits</div>
            <div class="metric-value ${afterSplits < beforeSplits ? 'cell-good' : ''}">${beforeSplits} &rarr; ${afterSplits}</div>
            <div class="metric-sub">${afterSplits < beforeSplits ? 'Improved' : afterSplits === 0 ? 'Clean' : 'Check'}</div>
        </div>
    `;

    // Before/After charts
    renderBeforeAfterChart('chart-before-after-arr', 'CSM ARR Distribution',
        beforeAnalysis.csmMetrics, afterAnalysis.csmMetrics, 'totalArr', fmt$);
    renderBeforeAfterChart('chart-before-after-health', 'CSM Health Distribution',
        beforeAnalysis.csmMetrics, afterAnalysis.csmMetrics, 'avgHealth', fmtPct);

    // Pod mapping table
    renderPodMapping(result);

    // Changes table
    renderChangesTable(result);
}

function renderBeforeAfterChart(canvasId, title, beforeBooks, afterBooks, metric, formatter) {
    if (charts[canvasId]) charts[canvasId].destroy();

    const allIds = [...new Set([...beforeBooks.map(b => b.id), ...afterBooks.map(b => b.id)])].sort();
    const beforeMap = {};
    beforeBooks.forEach(b => { beforeMap[b.id] = b[metric] || 0; });
    const afterMap = {};
    afterBooks.forEach(b => { afterMap[b.id] = b[metric] || 0; });

    charts[canvasId] = new Chart(document.getElementById(canvasId), {
        type: 'bar',
        data: {
            labels: allIds,
            datasets: [
                { label: 'Before', data: allIds.map(id => beforeMap[id] || 0), backgroundColor: 'rgba(156, 163, 175, 0.6)', borderRadius: 4 },
                { label: 'After', data: allIds.map(id => afterMap[id] || 0), backgroundColor: 'rgba(37, 99, 235, 0.6)', borderRadius: 4 },
            ]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                title: { display: true, text: title + ' (Before vs After)', font: { size: 13, weight: '600' } },
                tooltip: { callbacks: { label: ctx => `${ctx.dataset.label}: ${formatter(ctx.parsed.y)}` } },
            },
            scales: {
                x: { ticks: { font: { size: 9 }, maxRotation: 45 } },
                y: { ticks: { callback: v => formatter(v) } }
            }
        }
    });
}

function renderPodMapping(result) {
    const { accounts } = result;

    // Build pod pairings from the rebalanced assignments
    const pods = {};
    accounts.forEach(a => {
        const csm = a.newCsmId || a.csmId;
        const am = (a.newSalesOwnerId || a.salesOwnerId);
        if (!csm || !am || a.salesOwnerType !== 'Expansion AM/AD') return;

        const key = `${csm}||${am}`;
        if (!pods[key]) {
            pods[key] = {
                csm,
                am,
                segment: a.newCsSegment || a.csSegment,
                amSegment: a.newAmSegment || a.amSegment,
                sharedAccounts: 0,
                totalArr: 0,
                totalWhitespace: 0,
                accountIds: [],
            };
        }
        pods[key].sharedAccounts++;
        pods[key].totalArr += a.arr;
        pods[key].totalWhitespace += a.whitespace;
        pods[key].accountIds.push(a.accountId);
    });

    const podList = Object.values(pods).sort((a, b) => b.totalArr - a.totalArr);

    // Compute CSM:AM ratios per segment from the actual pairings
    const segRatios = {};
    podList.forEach(p => {
        const seg = p.segment;
        if (!segRatios[seg]) segRatios[seg] = { csms: new Set(), ams: new Set() };
        segRatios[seg].csms.add(p.csm);
        segRatios[seg].ams.add(p.am);
    });

    const ratioSummary = Object.entries(segRatios).map(([seg, r]) => {
        const csmCount = r.csms.size;
        const amCount = r.ams.size;
        const ratio = amCount > 0 ? (csmCount / amCount).toFixed(1) : 'N/A';
        return `<span class="pod-ratio-badge">${seg}: ${csmCount} CSMs / ${amCount} AMs = <strong>${ratio} : 1</strong></span>`;
    }).join(' ');

    const rows = podList.map(p => `
        <tr>
            <td>${p.csm}</td>
            <td>${p.am}</td>
            <td>${p.segment}</td>
            <td>${p.amSegment}</td>
            <td class="cell-num">${p.sharedAccounts}</td>
            <td class="cell-num">${fmt$(p.totalArr)}</td>
            <td class="cell-num">${fmt$(p.totalWhitespace)}</td>
        </tr>
    `).join('');

    document.getElementById('pod-mapping-wrap').innerHTML = `
        <div class="pod-section">
            <h3 class="pod-title">CSM : AM Pod Mapping</h3>
            <div class="pod-ratios">${ratioSummary}</div>
            <div class="data-table-wrap" style="max-height:400px;overflow-y:auto;">
                <table class="data-table">
                    <thead><tr>
                        <th>CSM</th><th>AM</th><th>CS Segment</th><th>AM Segment</th>
                        <th>Shared Accounts</th><th>Combined ARR</th><th>Combined Whitespace</th>
                    </tr></thead>
                    <tbody>${rows}</tbody>
                </table>
            </div>
        </div>
    `;
}

function renderChangesTable(result) {
    const changed = result.accounts.filter(a =>
        (a.newCsmId && a.newCsmId !== a.csmId) ||
        (a.newSalesOwnerId && a.newSalesOwnerId !== a.salesOwnerId && a.salesOwnerType === 'Expansion AM/AD')
    );

    if (changed.length === 0) {
        document.getElementById('changes-table-wrap').innerHTML =
            '<div class="flag info"><span class="flag-icon">&#10003;</span> No changes proposed - current structure is well-balanced for the given weights.</div>';
        return;
    }

    const rows = changed.map(a => {
        const csmChanged = a.newCsmId && a.newCsmId !== a.csmId;
        const amChanged = a.newSalesOwnerId && a.newSalesOwnerId !== a.salesOwnerId && a.salesOwnerType === 'Expansion AM/AD';
        return `
            <tr class="changed">
                <td>${a.accountId}</td>
                <td>${a.parentAccountId !== a.accountId ? a.parentAccountId : '—'}</td>
                <td>${a.csSegment}</td>
                <td class="cell-num">${fmt$(a.arr)}</td>
                <td>${csmChanged ? `<span class="cell-bad">${a.csmId}</span> &rarr; <span class="cell-good">${a.newCsmId}</span>` : a.csmId}</td>
                <td>${amChanged ? `<span class="cell-bad">${a.salesOwnerId}</span> &rarr; <span class="cell-good">${a.newSalesOwnerId}</span>` : a.salesOwnerId || '—'}</td>
                <td>${a.renewalDate ? a.renewalDate.toISOString().slice(0, 10) : '—'}</td>
                <td>${fmtPct(a.healthScore)}</td>
            </tr>
        `;
    }).join('');

    document.getElementById('changes-table-wrap').innerHTML = `
        <p style="margin-bottom:12px;font-size:0.85rem;color:var(--gray-600);">Showing ${changed.length} accounts with proposed ownership changes:</p>
        <div class="data-table-wrap" style="max-height:500px;overflow-y:auto;">
            <table class="data-table">
                <thead><tr>
                    <th>Account ID</th><th>Parent ID</th><th>Segment</th><th>ARR</th>
                    <th>CSM Change</th><th>AM Change</th><th>Renewal</th><th>Health</th>
                </tr></thead>
                <tbody>${rows}</tbody>
            </table>
        </div>
    `;
}

// ---------- CSV EXPORT ----------
document.getElementById('export-full-btn').addEventListener('click', exportFullCSV);
document.getElementById('export-changes-btn').addEventListener('click', exportChangesCSV);
document.getElementById('export-pods-btn').addEventListener('click', exportPodsCSV);

function buildExportRow(a) {
    const csmChanged = !!(a.newCsmId && a.newCsmId !== a.csmId);
    const amChanged = !!(a.newSalesOwnerId && a.newSalesOwnerId !== a.salesOwnerId && a.salesOwnerType === 'Expansion AM/AD');
    const csSegChanged = !!(a.newCsSegment && a.newCsSegment !== a.csSegment);
    const amSegChanged = !!(a.newAmSegment && a.newAmSegment !== a.amSegment);

    return {
        // Original columns preserved in input format
        'Parent Account ID': a.parentAccountId,
        'Account ID': a.accountId,
        'Account Type': a.accountType,
        'Renewal Date': a.renewalDate ? a.renewalDate.toISOString().slice(0, 10) : '',
        'Customer Tier': a.customerTier,
        'Total L12M Revenue': a.arr,
        'Total Estimated Whitespace': a.whitespace,
        'Country Code': a.countryCode,
        'Uses SMS': a.products.includes('SMS') ? 'Yes' : 'No',
        'Uses Email': a.products.includes('Email') ? 'Yes' : 'No',
        'Uses AI': a.products.includes('AI') ? 'Yes' : 'No',
        'CS Segment': a.csSegment,
        'CSM ID': a.csmId,
        'Sales Owner Type': a.salesOwnerType,
        'Account Management Segment': a.amSegment,
        'Sales Owner ID': a.salesOwnerId,
        'Health Score': a.healthScore !== null ? Math.round(a.healthScore * 100) + '%' : '',
        // New assignment columns
        'NEW CS Segment': a.newCsSegment || a.csSegment,
        'NEW CSM ID': a.newCsmId || a.csmId,
        'NEW AM Segment': a.newAmSegment || a.amSegment,
        'NEW Sales Owner ID': a.newSalesOwnerId || a.salesOwnerId,
        // Change tracking columns
        'CSM Changed': csmChanged ? 'Yes' : 'No',
        'OLD CSM ID': csmChanged ? a.csmId : '',
        'AM Changed': amChanged ? 'Yes' : 'No',
        'OLD Sales Owner ID': amChanged ? a.salesOwnerId : '',
        'CS Segment Changed': csSegChanged ? 'Yes' : 'No',
        'OLD CS Segment': csSegChanged ? a.csSegment : '',
        'AM Segment Changed': amSegChanged ? 'Yes' : 'No',
        'OLD AM Segment': amSegChanged ? a.amSegment : '',
    };
}

function downloadCSV(rows, filename) {
    const csv = Papa.unparse(rows);
    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    a.click();
    URL.revokeObjectURL(url);
}

function exportFullCSV() {
    if (!rebalanceResult) return;
    const rows = rebalanceResult.accounts.map(buildExportRow);
    downloadCSV(rows, `book_rebalance_full_${new Date().toISOString().slice(0, 10)}.csv`);
}

function exportChangesCSV() {
    if (!rebalanceResult) return;
    const changed = rebalanceResult.accounts.filter(a =>
        (a.newCsmId && a.newCsmId !== a.csmId) ||
        (a.newSalesOwnerId && a.newSalesOwnerId !== a.salesOwnerId && a.salesOwnerType === 'Expansion AM/AD')
    );
    const rows = changed.map(buildExportRow);
    downloadCSV(rows, `book_rebalance_changes_${new Date().toISOString().slice(0, 10)}.csv`);
}

function exportPodsCSV() {
    if (!rebalanceResult) return;
    const { accounts } = rebalanceResult;

    // Build pod pairings
    const pods = {};
    accounts.forEach(a => {
        const csm = a.newCsmId || a.csmId;
        const am = a.newSalesOwnerId || a.salesOwnerId;
        if (!csm || !am || a.salesOwnerType !== 'Expansion AM/AD') return;
        const key = `${csm}||${am}`;
        if (!pods[key]) {
            pods[key] = {
                csm, am,
                csSegment: a.newCsSegment || a.csSegment,
                amSegment: a.newAmSegment || a.amSegment,
                sharedAccounts: 0,
                totalArr: 0,
                totalWhitespace: 0,
                accountIds: [],
            };
        }
        pods[key].sharedAccounts++;
        pods[key].totalArr += a.arr;
        pods[key].totalWhitespace += a.whitespace;
        pods[key].accountIds.push(a.accountId);
    });

    const rows = Object.values(pods).sort((a, b) => b.totalArr - a.totalArr).map(p => ({
        'CSM ID': p.csm,
        'AM ID': p.am,
        'CS Segment': p.csSegment,
        'AM Segment': p.amSegment,
        'Shared Account Count': p.sharedAccounts,
        'Combined ARR': p.totalArr,
        'Combined Whitespace': p.totalWhitespace,
        'Account IDs': p.accountIds.join('; '),
    }));
    downloadCSV(rows, `pod_mapping_${new Date().toISOString().slice(0, 10)}.csv`);
}
