// ===== DATA FROM EXCEL =====
const months = ['24/01', '24/02', '24/03', '24/04', '24/05', '24/06', '24/07', '24/08', '24/09', '24/10', '24/11', '24/12', '25/01', '25/02', '25/03', '25/04', '25/09', '当前'];
const receivables = [800626, 793161, 1142292, 1714885, 2183231, 3507317, 5019747, 5853230, 5711282, 5436476, 5559407, 5348382, 5200521, 4586267, 3992471, 3647527, 2273437, 2056908];
const investment = [583143, 601811, 750122, 1113441, 1484822, 2204088, 3202145, 3707303, 3309735, 3014486, 2806621, 2634353, 2016459, 1412205, 826605, 437198, -614439, -856310];
const costs = [75330, 176946, 225010, 395219, 687467, 866871, 1262963, 1637156, 1857819, 2191338, 2335110, 2671996, 2671996, 2671996, 2671996, 2528825, 2580296, 2439236];
const overdue = [0, 49797, 0, 27201, 46607, 104688, 391368, 685629, 759430, 1000438, 1019315, 1451169, 1525081, 2172447, 2067337, 2142722, 2460085, 2056908];
const profit = [217484, 191350, 392171, 574243, 651803, 1198540, 1426234, 1460299, 1642117, 1421552, 1733471, 1262860, 1658981, 1001615, 1098529, 1067607, 427790, 856310];
const marginRate = [27.16, 24.13, 34.33, 33.49, 29.85, 34.17, 28.41, 24.95, 28.75, 26.15, 31.18, 23.61, 31.90, 21.84, 27.52, 29.27, 18.82, 41.63];
const overdueRate = receivables.map((r, i) => r > 0 ? +(overdue[i] / r * 100).toFixed(1) : 0);

// Sheet2 cumulative
const cumLabels = ['2月', '3月', '4月', '5月', '6月', '7月', '8月', '9月', '10月', '11月', '12月'];
const cumTotal = [168816, 211330, 360389, 645137, 808061, 1173793, 1531906, 1741889, 2067548, 2201180, 2562726];

// Sheet3 expense categories (aggregated from raw data)
const expenseCategories = {
    '固定支出': 933361, '提成': 617157, '一次性支出': 319748,
    '风控充值': 53500, '资金成本': 120247
};
// Sheet3 monthly expenses (aggregated)
const expMonths = ['24/01', '24/02', '24/03', '24/04', '24/05', '24/06', '24/07', '24/08', '24/09', '24/10', '24/11', '24/12'];
const expAmounts = [78813, 71994, 37494, 144888, 232944, 174506, 298870, 256570, 102050, 296443, 129457, 220984];

// ===== CHART DEFAULTS =====
Chart.defaults.color = '#94a3b8';
Chart.defaults.borderColor = 'rgba(99,102,241,0.08)';
Chart.defaults.font.family = "'Inter', sans-serif";
Chart.defaults.plugins.legend.labels.usePointStyle = true;
Chart.defaults.plugins.legend.labels.pointStyleWidth = 10;
Chart.defaults.plugins.legend.labels.padding = 14;

const tt = {
    backgroundColor: 'rgba(15,23,42,0.95)', titleColor: '#f1f5f9', bodyColor: '#cbd5e1',
    borderColor: 'rgba(99,102,241,0.3)', borderWidth: 1, padding: 12, cornerRadius: 8, displayColors: true,
    callbacks: {
        label: ctx => {
            let v = ctx.parsed.y ?? ctx.parsed;
            if (typeof v === 'number' && Math.abs(v) >= 10000) return ctx.dataset.label + ': ¥' + (v / 10000).toFixed(2) + '万';
            return ctx.dataset.label + ': ' + v;
        }
    }
};
function fmtY(v) { return Math.abs(v) >= 10000 ? (v / 10000).toFixed(0) + '万' : v; }
function mkGrad(ctx, c1, c2) { const g = ctx.chart.ctx.createLinearGradient(0, 0, 0, 260); g.addColorStop(0, c1); g.addColorStop(1, c2); return g; }

// ===== 1. MAIN TREND (3 lines) =====
new Chart(document.getElementById('chartMain'), {
    type: 'line',
    data: {
        labels: months, datasets: [
            { label: '待收金额', data: receivables, borderColor: '#6366f1', backgroundColor: 'rgba(99,102,241,0.08)', fill: true, tension: .4, pointRadius: 3, borderWidth: 2.5 },
            { label: '实际出资', data: investment, borderColor: '#06b6d4', backgroundColor: 'rgba(6,182,212,0.06)', fill: true, tension: .4, pointRadius: 3, borderWidth: 2.5 },
            { label: '成本支出', data: costs, borderColor: '#a855f7', borderDash: [6, 3], tension: .4, pointRadius: 2, borderWidth: 2, fill: false }
        ]
    },
    options: { responsive: true, maintainAspectRatio: false, plugins: { tooltip: tt }, scales: { y: { ticks: { callback: fmtY }, grid: { color: 'rgba(99,102,241,0.06)' } } } }
});

// ===== 2. PROFIT DUAL AXIS =====
new Chart(document.getElementById('chartProfitDual'), {
    type: 'bar',
    data: {
        labels: months, datasets: [
            { label: '利润金额', data: profit, backgroundColor: 'rgba(16,185,129,0.6)', borderRadius: 4, barPercentage: .6, yAxisID: 'y' },
            { label: '利润率%', data: marginRate, type: 'line', borderColor: '#f59e0b', backgroundColor: 'transparent', pointRadius: 4, pointBackgroundColor: '#f59e0b', borderWidth: 2.5, tension: .3, yAxisID: 'y1' }
        ]
    },
    options: {
        responsive: true, maintainAspectRatio: false, plugins: { tooltip: { ...tt, callbacks: { label: ctx => ctx.dataset.label + ': ' + (ctx.datasetIndex === 1 ? ctx.parsed.y.toFixed(1) + '%' : '¥' + (ctx.parsed.y / 10000).toFixed(2) + '万') } } },
        scales: { y: { position: 'left', ticks: { callback: fmtY } }, y1: { position: 'right', min: 0, max: 50, ticks: { callback: v => v + '%' }, grid: { drawOnChartArea: false } } }
    }
});

// ===== 3. INVESTMENT RECOVERY =====
new Chart(document.getElementById('chartRecovery'), {
    type: 'bar',
    data: {
        labels: months, datasets: [{
            label: '实际出资(负=已回本)', data: investment,
            backgroundColor: investment.map(v => v >= 0 ? 'rgba(244,63,94,0.6)' : 'rgba(16,185,129,0.7)'),
            borderRadius: 4, barPercentage: .7
        }]
    },
    options: {
        responsive: true, maintainAspectRatio: false, plugins: {
            tooltip: tt,
            annotation: { annotations: { zeroline: { type: 'line', yMin: 0, yMax: 0, borderColor: 'rgba(255,255,255,0.3)', borderWidth: 1, borderDash: [4, 4] } } }
        }, scales: { y: { ticks: { callback: fmtY } } }
    }
});

// ===== 4. OVERDUE AMOUNT =====
new Chart(document.getElementById('chartOverdue'), {
    type: 'line',
    data: {
        labels: months, datasets: [{
            label: '逾期金额', data: overdue, borderColor: '#f43f5e',
            backgroundColor: ctx => mkGrad(ctx, 'rgba(244,63,94,0.25)', 'rgba(244,63,94,0)'),
            fill: true, tension: .4, pointRadius: 3, borderWidth: 2.5
        }]
    },
    options: { responsive: true, maintainAspectRatio: false, plugins: { tooltip: tt }, scales: { y: { ticks: { callback: fmtY } } } }
});

// ===== 5. OVERDUE RATE =====
new Chart(document.getElementById('chartOverdueRate'), {
    type: 'line',
    data: {
        labels: months, datasets: [{
            label: '逾期率%', data: overdueRate, borderColor: '#f59e0b',
            backgroundColor: ctx => mkGrad(ctx, 'rgba(245,158,11,0.2)', 'rgba(245,158,11,0)'),
            fill: true, tension: .4, pointRadius: 4, pointBackgroundColor: overdueRate.map(v => v > 50 ? '#f43f5e' : v > 30 ? '#f59e0b' : '#10b981'),
            borderWidth: 2.5
        }]
    },
    options: {
        responsive: true, maintainAspectRatio: false, plugins: { tooltip: { ...tt, callbacks: { label: ctx => '逾期率: ' + ctx.parsed.y + '%' } } },
        scales: { y: { min: 0, max: 110, ticks: { callback: v => v + '%' } } }
    }
});

// ===== 6. REVENUE COMPOSITION =====
new Chart(document.getElementById('chartRevenueCompose'), {
    type: 'doughnut',
    data: {
        labels: ['租金收入', '尾款收入', '增值费', '延保服务', '首付款', '买断金'],
        datasets: [{
            data: [10251940, 2407173, 166612, 418858, 31206, 16],
            backgroundColor: ['rgba(99,102,241,0.8)', 'rgba(6,182,212,0.8)', 'rgba(168,85,247,0.8)', 'rgba(16,185,129,0.8)', 'rgba(245,158,11,0.8)', 'rgba(244,63,94,0.6)'],
            borderWidth: 0, hoverOffset: 12
        }]
    },
    options: { responsive: true, maintainAspectRatio: false, cutout: '60%', plugins: { tooltip: tt, legend: { position: 'bottom', labels: { font: { size: 11 } } } } }
});

// ===== 7. STORE RECEIVABLE CONCENTRATION =====
new Chart(document.getElementById('chartStoreReceivable'), {
    type: 'doughnut',
    data: {
        labels: ['涛涛好物', '刚刚好物', '太太租物'],
        datasets: [{
            data: [1612442, 414804, 29662],
            backgroundColor: ['rgba(244,63,94,0.75)', 'rgba(245,158,11,0.75)', 'rgba(16,185,129,0.75)'],
            borderWidth: 0, hoverOffset: 10
        }]
    },
    options: { responsive: true, maintainAspectRatio: false, cutout: '60%', plugins: { tooltip: tt, legend: { position: 'bottom', labels: { font: { size: 11 } } } } }
});

// ===== 8. STORE ORDER =====
new Chart(document.getElementById('chartStoreOrder'), {
    type: 'doughnut',
    data: {
        labels: ['涛涛好物', '刚刚好物', '太太租物'],
        datasets: [{
            data: [726, 232, 74],
            backgroundColor: ['rgba(99,102,241,0.8)', 'rgba(6,182,212,0.8)', 'rgba(168,85,247,0.8)'],
            borderWidth: 0, hoverOffset: 10
        }]
    },
    options: { responsive: true, maintainAspectRatio: false, cutout: '60%', plugins: { tooltip: tt, legend: { position: 'bottom', labels: { font: { size: 11 } } } } }
});

// ===== 9. LEASE vs ECOM =====
new Chart(document.getElementById('chartLeaseEcom'), {
    type: 'bar',
    data: {
        labels: ['涛涛好物', '刚刚好物', '太太租物', '总平台'],
        datasets: [
            { label: '租赁业绩', data: [3011805, 0, 105082, 3116887], backgroundColor: 'rgba(99,102,241,0.7)', borderRadius: 4 },
            { label: '电商业绩', data: [0, 1250200, 15797, 1265996], backgroundColor: 'rgba(6,182,212,0.7)', borderRadius: 4 }
        ]
    },
    options: { responsive: true, maintainAspectRatio: false, plugins: { tooltip: tt }, scales: { y: { ticks: { callback: fmtY } } } }
});

// ===== 10. RADAR =====
new Chart(document.getElementById('chartRadar'), {
    type: 'radar',
    data: {
        labels: ['订单数', '总待收', '租金', '放款', '业绩'],
        datasets: [
            { label: '涛涛好物', data: [100, 78, 92, 71, 69], borderColor: '#6366f1', backgroundColor: 'rgba(99,102,241,0.15)', borderWidth: 2, pointRadius: 3 },
            { label: '刚刚好物', data: [32, 20, 0.02, 23, 29], borderColor: '#06b6d4', backgroundColor: 'rgba(6,182,212,0.1)', borderWidth: 2, pointRadius: 3 },
            { label: '太太租物', data: [10, 1.4, 7.6, 6.3, 2.8], borderColor: '#a855f7', backgroundColor: 'rgba(168,85,247,0.1)', borderWidth: 2, pointRadius: 3 }
        ]
    },
    options: { responsive: true, maintainAspectRatio: false, plugins: { tooltip: tt }, scales: { r: { min: 0, max: 100, ticks: { display: false }, grid: { color: 'rgba(99,102,241,0.15)' }, pointLabels: { color: '#94a3b8', font: { size: 11 } } } } }
});

// ===== 11. CUMULATIVE COST =====
new Chart(document.getElementById('chartCumulative'), {
    type: 'line',
    data: {
        labels: cumLabels, datasets: [{
            label: '累计成本', data: cumTotal, borderColor: '#f59e0b',
            backgroundColor: ctx => mkGrad(ctx, 'rgba(245,158,11,0.15)', 'rgba(245,158,11,0)'),
            fill: true, tension: .4, pointRadius: 4, borderWidth: 2.5
        }]
    },
    options: { responsive: true, maintainAspectRatio: false, plugins: { tooltip: tt }, scales: { y: { ticks: { callback: fmtY } } } }
});

// ===== 12. MARGIN BAR =====
new Chart(document.getElementById('chartMarginBar'), {
    type: 'bar',
    data: {
        labels: months, datasets: [{
            label: '利润率%', data: marginRate,
            backgroundColor: marginRate.map(v => v >= 35 ? 'rgba(16,185,129,0.8)' : v >= 25 ? 'rgba(99,102,241,0.7)' : 'rgba(244,63,94,0.6)'),
            borderRadius: 6, barPercentage: .7
        }]
    },
    options: {
        responsive: true, maintainAspectRatio: false, plugins: { tooltip: { ...tt, callbacks: { label: ctx => '利润率: ' + ctx.parsed.y.toFixed(1) + '%' } } },
        scales: { y: { ticks: { callback: v => v + '%' }, suggestedMax: 50 } }
    }
});

// ===== 13. EXPENSE CATEGORIES =====
const catLabels = Object.keys(expenseCategories);
const catValues = Object.values(expenseCategories);
new Chart(document.getElementById('chartExpenseCat'), {
    type: 'doughnut',
    data: {
        labels: catLabels, datasets: [{
            data: catValues,
            backgroundColor: ['rgba(99,102,241,0.8)', 'rgba(244,63,94,0.7)', 'rgba(245,158,11,0.7)', 'rgba(6,182,212,0.7)', 'rgba(168,85,247,0.7)'],
            borderWidth: 0, hoverOffset: 10
        }]
    },
    options: { responsive: true, maintainAspectRatio: false, cutout: '60%', plugins: { tooltip: tt, legend: { position: 'bottom', labels: { font: { size: 11 } } } } }
});

// ===== 14. MONTHLY EXPENSE =====
new Chart(document.getElementById('chartExpenseMonth'), {
    type: 'bar',
    data: {
        labels: expMonths, datasets: [{
            label: '月度支出', data: expAmounts,
            backgroundColor: 'rgba(168,85,247,0.6)', borderRadius: 4, barPercentage: .7
        }]
    },
    options: { responsive: true, maintainAspectRatio: false, plugins: { tooltip: tt }, scales: { y: { ticks: { callback: fmtY } } } }
});
