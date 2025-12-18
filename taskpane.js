/**
 * DataGen Statistics - Office Add-in
 * JavaScript logic for Excel integration
 *
 * Based on DataGen by Dr. Russell Hurlburt
 * Office Add-in implementation by dep2025
 */

// Global state
let isOfficeInitialized = false;

// Initialize Office.js
Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        isOfficeInitialized = true;
        document.getElementById("statusMessage").textContent = "Ready";

        // Set up event listeners
        setupEventListeners();

        // Initial data load
        refreshDescriptives();

        // Set up worksheet change event handler for real-time updates
        setupChangeHandler();
    } else {
        document.getElementById("statusMessage").textContent = "This add-in requires Excel";
    }
});

/**
 * Set up all UI event listeners
 */
function setupEventListeners() {
    // Tab navigation
    document.querySelectorAll('.tab-btn').forEach(btn => {
        btn.addEventListener('click', (e) => {
            switchTab(e.target.dataset.tab);
        });
    });

    // Refresh buttons
    document.getElementById('refreshDescriptives').addEventListener('click', refreshDescriptives);
    document.getElementById('refreshStatistics').addEventListener('click', refreshStatistics);
    document.getElementById('generateData').addEventListener('click', generateData);

    // Distribution type change
    document.getElementById('distType').addEventListener('change', (e) => {
        updateDistributionParams(e.target.value);
    });
}

/**
 * Set up Excel change event handler for real-time updates
 */
async function setupChangeHandler() {
    try {
        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getActiveWorksheet();

            // Register for change events
            sheet.onChanged.add(async (event) => {
                // Debounce rapid changes
                clearTimeout(window.changeTimeout);
                window.changeTimeout = setTimeout(() => {
                    refreshDescriptives();
                }, 500);
            });

            await context.sync();
        });
    } catch (error) {
        console.log("Change handler setup error:", error);
    }
}

/**
 * Switch between tabs
 */
function switchTab(tabName) {
    // Update tab buttons
    document.querySelectorAll('.tab-btn').forEach(btn => {
        btn.classList.toggle('active', btn.dataset.tab === tabName);
    });

    // Update tab content
    document.querySelectorAll('.tab-content').forEach(content => {
        content.classList.toggle('active', content.id === tabName);
    });

    // Refresh data for the active tab
    if (tabName === 'descriptives') {
        refreshDescriptives();
    } else if (tabName === 'statistics') {
        refreshStatistics();
    }
}

/**
 * Refresh descriptive statistics from spreadsheet data
 */
async function refreshDescriptives() {
    if (!isOfficeInitialized) return;

    document.getElementById("statusMessage").textContent = "Calculating...";

    try {
        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getActiveWorksheet();

            // Define data ranges for 5 variables (columns B through F, rows 4-203)
            // This matches the VBA DataGen structure
            const dataRanges = [
                sheet.getRange("B4:B203"),
                sheet.getRange("C4:C203"),
                sheet.getRange("D4:D203"),
                sheet.getRange("E4:E203"),
                sheet.getRange("F4:F203")
            ];

            // Get variable names from row 3
            const headerRange = sheet.getRange("B3:F3");
            headerRange.load("values");

            // Load all data ranges
            dataRanges.forEach(range => range.load("values"));

            await context.sync();

            // Update variable headers
            const headers = headerRange.values[0];
            for (let i = 0; i < 5; i++) {
                const headerEl = document.getElementById(`var${i + 1}-header`);
                if (headerEl) {
                    headerEl.textContent = headers[i] || `Variable ${i + 1}`;
                }
            }

            // Calculate and display statistics for each variable
            for (let i = 0; i < 5; i++) {
                const values = dataRanges[i].values
                    .flat()
                    .filter(v => v !== null && v !== "" && !isNaN(v))
                    .map(Number);

                updateVariableStats(i + 1, values);
            }

            document.getElementById("statusMessage").textContent = "Ready";
        });
    } catch (error) {
        console.error("Error refreshing descriptives:", error);
        document.getElementById("statusMessage").textContent = "Error: " + error.message;
    }
}

/**
 * Update statistics display for a single variable
 */
function updateVariableStats(varNum, values) {
    const stats = calculateDescriptiveStats(values);

    document.getElementById(`n-${varNum}`).textContent = stats.n;
    document.getElementById(`mean-${varNum}`).textContent = formatNumber(stats.mean);
    document.getElementById(`median-${varNum}`).textContent = formatNumber(stats.median);
    document.getElementById(`mode-${varNum}`).textContent = stats.mode;
    document.getElementById(`stddev-${varNum}`).textContent = formatNumber(stats.stdDev);
    document.getElementById(`variance-${varNum}`).textContent = formatNumber(stats.variance);
    document.getElementById(`min-${varNum}`).textContent = formatNumber(stats.min);
    document.getElementById(`max-${varNum}`).textContent = formatNumber(stats.max);
    document.getElementById(`range-${varNum}`).textContent = formatNumber(stats.range);
    document.getElementById(`sum-${varNum}`).textContent = formatNumber(stats.sum);
    document.getElementById(`sumsq-${varNum}`).textContent = formatNumber(stats.sumOfSquares);
}

/**
 * Calculate descriptive statistics for an array of values
 */
function calculateDescriptiveStats(values) {
    if (!values || values.length === 0) {
        return {
            n: 0, mean: '-', median: '-', mode: '-', stdDev: '-',
            variance: '-', min: '-', max: '-', range: '-', sum: '-', sumOfSquares: '-'
        };
    }

    const n = values.length;
    const sum = values.reduce((a, b) => a + b, 0);
    const mean = sum / n;

    // Sorted values for median
    const sorted = [...values].sort((a, b) => a - b);
    const median = n % 2 === 0
        ? (sorted[n / 2 - 1] + sorted[n / 2]) / 2
        : sorted[Math.floor(n / 2)];

    // Mode
    const frequency = {};
    values.forEach(v => {
        const rounded = Math.round(v * 1000) / 1000;
        frequency[rounded] = (frequency[rounded] || 0) + 1;
    });
    const maxFreq = Math.max(...Object.values(frequency));
    const modes = Object.keys(frequency)
        .filter(k => frequency[k] === maxFreq)
        .map(Number);
    const mode = maxFreq === 1 ? "None" : (modes.length > 3 ? "Multiple" : modes.join(", "));

    // Variance and Standard Deviation (sample)
    const squaredDiffs = values.map(v => Math.pow(v - mean, 2));
    const variance = squaredDiffs.reduce((a, b) => a + b, 0) / (n - 1);
    const stdDev = Math.sqrt(variance);

    // Min, Max, Range
    const min = sorted[0];
    const max = sorted[n - 1];
    const range = max - min;

    // Sum of Squares
    const sumOfSquares = values.reduce((a, b) => a + b * b, 0);

    return {
        n, mean, median, mode, stdDev, variance,
        min, max, range, sum, sumOfSquares
    };
}

/**
 * Refresh inferential statistics
 */
async function refreshStatistics() {
    if (!isOfficeInitialized) return;

    document.getElementById("statusMessage").textContent = "Calculating...";

    try {
        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            const selection = context.workbook.getSelectedRange();
            selection.load("columnCount, rowCount, values, address");

            await context.sync();

            const colCount = selection.columnCount;

            // Hide all stat panels first
            document.getElementById('oneVarStats').style.display = 'none';
            document.getElementById('twoVarStats').style.display = 'none';
            document.getElementById('pairedStats').style.display = 'none';

            if (colCount === 1) {
                // One variable selected
                document.getElementById('oneVarStats').style.display = 'block';
            } else if (colCount === 2) {
                // Two variables selected
                document.getElementById('twoVarStats').style.display = 'block';

                // Get data from both columns
                const values = selection.values;
                const col1 = values.map(row => row[0]).filter(v => v !== null && v !== "" && !isNaN(v)).map(Number);
                const col2 = values.map(row => row[1]).filter(v => v !== null && v !== "" && !isNaN(v)).map(Number);

                if (col1.length > 0 && col2.length > 0) {
                    const twoVarStats = calculateTwoVariableStats(col1, col2);

                    document.getElementById('correlation').textContent = formatNumber(twoVarStats.correlation);
                    document.getElementById('r-squared').textContent = formatNumber(twoVarStats.rSquared);
                    document.getElementById('slope').textContent = formatNumber(twoVarStats.slope);
                    document.getElementById('intercept').textContent = formatNumber(twoVarStats.intercept);
                    document.getElementById('t-independent').textContent = formatNumber(twoVarStats.tIndependent);

                    // Check if paired (same N)
                    if (col1.length === col2.length) {
                        document.getElementById('pairedStats').style.display = 'block';
                        document.getElementById('t-paired').textContent = formatNumber(twoVarStats.tPaired);
                    }
                }
            }

            document.getElementById("statusMessage").textContent = "Ready";
        });
    } catch (error) {
        console.error("Error refreshing statistics:", error);
        document.getElementById("statusMessage").textContent = "Error: " + error.message;
    }
}

/**
 * Calculate two-variable statistics (correlation, regression, t-tests)
 */
function calculateTwoVariableStats(x, y) {
    // Use paired data (minimum length)
    const n = Math.min(x.length, y.length);
    const xData = x.slice(0, n);
    const yData = y.slice(0, n);

    // Means
    const xMean = xData.reduce((a, b) => a + b, 0) / n;
    const yMean = yData.reduce((a, b) => a + b, 0) / n;

    // Correlation
    let sumXY = 0, sumX2 = 0, sumY2 = 0;
    for (let i = 0; i < n; i++) {
        const dx = xData[i] - xMean;
        const dy = yData[i] - yMean;
        sumXY += dx * dy;
        sumX2 += dx * dx;
        sumY2 += dy * dy;
    }
    const correlation = sumXY / Math.sqrt(sumX2 * sumY2);
    const rSquared = correlation * correlation;

    // Regression (y = slope * x + intercept)
    const slope = sumXY / sumX2;
    const intercept = yMean - slope * xMean;

    // Independent samples t-test
    const xVar = sumX2 / (n - 1);
    const yVar = sumY2 / (n - 1);
    const pooledSE = Math.sqrt(xVar / n + yVar / n);
    const tIndependent = (xMean - yMean) / pooledSE;

    // Paired samples t-test
    const diffs = xData.map((v, i) => v - yData[i]);
    const diffMean = diffs.reduce((a, b) => a + b, 0) / n;
    const diffVar = diffs.reduce((a, b) => a + Math.pow(b - diffMean, 2), 0) / (n - 1);
    const diffSE = Math.sqrt(diffVar / n);
    const tPaired = diffMean / diffSE;

    return {
        correlation,
        rSquared,
        slope,
        intercept,
        tIndependent,
        tPaired
    };
}

/**
 * Generate random data based on user settings
 */
async function generateData() {
    if (!isOfficeInitialized) return;

    document.getElementById("statusMessage").textContent = "Generating...";

    try {
        await Excel.run(async (context) => {
            const selection = context.workbook.getSelectedRange();
            selection.load("address, columnIndex, rowIndex, rowCount");

            await context.sync();

            const distType = document.getElementById('distType').value;
            const decimals = parseInt(document.getElementById('decimals').value);
            const rowCount = parseInt(document.getElementById('rowCount').value);

            // Generate data array
            const data = [];
            for (let i = 0; i < rowCount; i++) {
                let value;

                switch (distType) {
                    case 'normal':
                        const mean = parseFloat(document.getElementById('mean').value);
                        const stdDev = parseFloat(document.getElementById('stdDev').value);
                        value = generateNormal(mean, stdDev);
                        break;
                    case 'uniform':
                        const minVal = parseFloat(document.getElementById('minVal').value);
                        const maxVal = parseFloat(document.getElementById('maxVal').value);
                        value = generateUniform(minVal, maxVal);
                        break;
                    case 'sequence':
                        const startVal = parseFloat(document.getElementById('startVal').value);
                        const increment = parseFloat(document.getElementById('increment').value);
                        value = startVal + (i * increment);
                        break;
                }

                data.push([roundTo(value, decimals)]);
            }

            // Write to Excel starting at row 4 (matching VBA DataGen structure)
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            const colIndex = selection.columnIndex;
            const targetRange = sheet.getRangeByIndexes(3, colIndex, rowCount, 1);
            targetRange.values = data;

            await context.sync();

            document.getElementById("statusMessage").textContent = "Data generated!";

            // Refresh statistics after generating
            setTimeout(refreshDescriptives, 100);
        });
    } catch (error) {
        console.error("Error generating data:", error);
        document.getElementById("statusMessage").textContent = "Error: " + error.message;
    }
}

/**
 * Generate normally distributed random number using Box-Muller transform
 * (Same algorithm as VBA version)
 */
function generateNormal(mean, stdDev) {
    const u1 = Math.random();
    const u2 = Math.random();
    const z = Math.sqrt(-2 * Math.log(u1)) * Math.cos(2 * Math.PI * u2);
    return mean + z * stdDev;
}

/**
 * Generate uniformly distributed random number
 */
function generateUniform(min, max) {
    return min + Math.random() * (max - min);
}

/**
 * Update distribution parameter visibility
 */
function updateDistributionParams(distType) {
    document.getElementById('normalParams').style.display = distType === 'normal' ? 'block' : 'none';
    document.getElementById('uniformParams').style.display = distType === 'uniform' ? 'block' : 'none';
    document.getElementById('sequenceParams').style.display = distType === 'sequence' ? 'block' : 'none';
}

/**
 * Format number for display
 */
function formatNumber(value, decimals = 4) {
    if (value === '-' || value === null || value === undefined || isNaN(value)) {
        return '-';
    }
    return Number(value).toFixed(decimals);
}

/**
 * Round to specified decimal places
 */
function roundTo(value, decimals) {
    const factor = Math.pow(10, decimals);
    return Math.round(value * factor) / factor;
}
