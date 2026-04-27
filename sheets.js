// Google Sheets data fetcher
// Uses the public CSV export endpoint — no OAuth needed for sheets
// shared with "Anyone with the link can view"

const SHEET_CONFIGS = {
  all-entites: {
    sheetId: '1bKBvKEUvLXdpi3Ft5KjrBp3Vs_Xq_chn8xZAmvdso3g',  // Replace with your sheet ID
    tabName: 'Master KPI - All Entities - by Mo.',
    gid: ''  // Fill in after checking URL of that tab
  },
  sagehouse: {
    sheetId: '1bKBvKEUvLXdpi3Ft5KjrBp3Vs_Xq_chn8xZAmvdso3g',  // Replace with your sheet ID
    tabName: 'Master KPI - TSH-Only Group - by Mo.',
    gid: ''  // Fill in after checking URL of that tab
  },
  spv1: {
    sheetId: '1bKBvKEUvLXdpi3Ft5KjrBp3Vs_Xq_chn8xZAmvdso3g',
    tabName: 'Master KPI - SHC SPV1 - by Mo',
    gid: ''
  },
  proof3: {
    sheetId: '1bKBvKEUvLXdpi3Ft5KjrBp3Vs_Xq_chn8xZAmvdso3g',
    tabName: 'Master KPI - p3 w/ SHC SPV1 - by Mo.',
    gid: ''
  }
};

async function fetchSheetData(newsletterType) {
  const config = SHEET_CONFIGS[newsletterType];
  if (!config) throw new Error('Unknown newsletter type: ' + newsletterType);

  // Google Sheets CSV export URL — works for any sheet shared as "view"
  const url = `https://docs.google.com/spreadsheets/d/${config.sheetId}/export?format=csv&gid=${config.gid}`;

  const response = await fetch(url);
  if (!response.ok) {
    throw new Error(
      'Could not read sheet. Make sure it is shared as "Anyone with the link can view".'
    );
  }

  const csv = await response.text();
  return parseAndCleanCsv(csv, config.tabName);
}

function parseAndCleanCsv(csv, tabName) {
  const lines = csv.split('\n').filter(l => l.trim());
  if (lines.length < 3) throw new Error('Sheet appears empty or not loaded yet.');

  // Parse CSV properly handling quoted fields
  const rows = lines.map(parseCsvLine);

  // Find header row (row with month dates)
  const headerRowIndex = rows.findIndex(row =>
    row.some(cell => /jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec/i.test(cell))
  );

  if (headerRowIndex === -1) {
    throw new Error('Could not find month headers in sheet. Check the tab structure.');
  }

  // Get last 13 columns (metric name + 12 months)
  const headers = rows[headerRowIndex];
  const totalCols = headers.length;
  const startCol = Math.max(1, totalCols - 12); // Last 12 months
  const monthHeaders = headers.slice(startCol);

  // Get data rows (skip empty rows and formula helper rows)
  const dataRows = rows
    .slice(headerRowIndex + 1)
    .filter(row => row[0] && row[0].trim() && !row[0].startsWith('='));

  // Build clean table string for Claude
  let table = `Tab: ${tabName}\nMetric\t${monthHeaders.join('\t')}\n`;
  table += '─'.repeat(80) + '\n';

  dataRows.forEach(row => {
    const metricName = row[0];
    const values = row.slice(startCol).map(v => {
      const num = parseFloat(v.replace(/[$,%]/g, ''));
      if (!isNaN(num)) {
        // Format intelligently
        if (v.includes('%')) return (num).toFixed(1) + '%';
        if (Math.abs(num) >= 1000000) return '$' + (num/1000000).toFixed(2) + 'M';
        if (Math.abs(num) >= 1000) return '$' + (num/1000).toFixed(1) + 'k';
        return v.trim();
      }
      return v.trim();
    });
    table += `${metricName}\t${values.join('\t')}\n`;
  });

  return {
    table,
    monthHeaders,
    rowCount: dataRows.length,
    colCount: monthHeaders.length
  };
}

function parseCsvLine(line) {
  const result = [];
  let current = '';
  let inQuotes = false;

  for (let i = 0; i < line.length; i++) {
    if (line[i] === '"') {
      inQuotes = !inQuotes;
    } else if (line[i] === ',' && !inQuotes) {
      result.push(current.trim());
      current = '';
    } else {
      current += line[i];
    }
  }
  result.push(current.trim());
  return result;
}

export { fetchSheetData, SHEET_CONFIGS };
