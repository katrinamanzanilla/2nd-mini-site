(() => {
  const STORAGE_KEY = 'projstat:lastSheetSource';
  const ALIASES = {
    system: ['system', 'project name', 'system project name'],
    milestone: ['milestone', 'next milestone'],
    developer: ['assigned developer', 'developer'],
    manager: ['assigned project manager', 'project manager']
  };

  const state = {
    headers: [],
    rows: [],
    filteredRows: [],
    indexes: {
      system: -1,
      milestone: -1,
      developer: -1,
      manager: -1
    }
  };

  const el = {};

  document.addEventListener('DOMContentLoaded', init);

  function init() {
    el.form = document.getElementById('sheet-source-form');
    el.input = document.getElementById('sheet-source-input');
    el.reset = document.getElementById('sheet-source-reset');
    el.feedback = document.getElementById('sheet-source-feedback');
    el.tableHead = document.getElementById('status-table-head');
    el.tableBody = document.getElementById('status-table-body');
    el.kpiProjects = document.getElementById('kpi-total-projects');
    el.kpiMilestones = document.getElementById('kpi-total-milestones');
    el.systemFilter = document.getElementById('system-filter');
    el.milestoneFilter = document.getElementById('milestone-filter');
    el.search = document.getElementById('sheet-search');

    el.form.addEventListener('submit', handleLoad);
    el.reset.addEventListener('click', handleReset);
    el.systemFilter.addEventListener('change', applyFilters);
    el.milestoneFilter.addEventListener('change', applyFilters);
    el.search.addEventListener('input', applyFilters);

    const saved = localStorage.getItem(STORAGE_KEY);
    if (saved) {
      el.input.value = saved;
      loadSheet(saved, { isAutoLoad: true });
    }
  }

  async function handleLoad(event) {
    event.preventDefault();
    const source = el.input.value.trim();
    if (!source) {
      setFeedback('Please enter a Google Sheets link or sheet ID.', 'error');
      return;
    }

    localStorage.setItem(STORAGE_KEY, source);
    await loadSheet(source, { isAutoLoad: false });
  }

  function handleReset() {
    localStorage.removeItem(STORAGE_KEY);
    el.input.value = '';
    el.systemFilter.value = '';
    el.milestoneFilter.value = '';
    el.search.value = '';

    state.headers = [];
    state.rows = [];
    state.filteredRows = [];
    state.indexes = { system: -1, milestone: -1, developer: -1, manager: -1 };

    renderTable();
    updateKpis();
    populateFilterOptions();
    setFeedback('Cleared saved source, filters, table, and scorecards.', 'ok');
  }

  async function loadSheet(source, { isAutoLoad }) {
    const parsed = parseSheetSource(source);
    if (!parsed.sheetId) {
      setFeedback('Could not extract a Google Sheet ID. Use a docs/drive link or a raw sheet ID.', 'error');
      return;
    }

    setFeedback('Loading sheet data…', 'ok');

    try {
      const loaded = await loadWithFallbackChain(parsed);
      if (!loaded.headers.length) {
        throw new Error('Sheet was loaded but no columns were found. Ensure the first row contains column headers.');
      }

      state.headers = loaded.headers;
      state.rows = loaded.rows;
      state.filteredRows = loaded.rows.slice();
      state.indexes = findColumnIndexes(state.headers);

      populateFilterOptions();
      applyFilters();

      const via = loaded.sourceLabel ? ` via ${loaded.sourceLabel}` : '';
      const auto = isAutoLoad ? ' from saved source' : '';
      setFeedback(`Loaded ${loaded.rows.length} rows${auto}${via}.`, 'ok');
    } catch (error) {
      state.headers = [];
      state.rows = [];
      state.filteredRows = [];
      renderTable();
      updateKpis();

      setFeedback(
        `Unable to load this sheet. Confirm the sheet is shared for public viewing and the link/ID is correct. Details: ${error.message}`,
        'error'
      );
    }
  }

  function parseSheetSource(input) {
    const raw = input.trim();
    const result = { sheetId: '', gid: '', sheet: '' };

    if (/^[a-zA-Z0-9-_]{20,}$/.test(raw)) {
      result.sheetId = raw;
      return result;
    }

    let url;
    try {
      url = new URL(raw);
    } catch {
      return result;
    }

    const docsMatch = url.pathname.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
    const fileMatch = url.pathname.match(/\/file\/d\/([a-zA-Z0-9-_]+)/);

    result.sheetId = docsMatch?.[1] || fileMatch?.[1] || url.searchParams.get('id') || url.searchParams.get('key') || '';
    result.gid = url.searchParams.get('gid') || '';
    result.sheet = url.searchParams.get('sheet') || url.searchParams.get('sheetName') || '';

    if (url.hash) {
      const hash = new URLSearchParams(url.hash.replace(/^#/, ''));
      result.gid = result.gid || hash.get('gid') || '';
      result.sheet = result.sheet || hash.get('sheet') || hash.get('sheetName') || '';
    }

    return result;
  }

  async function loadWithFallbackChain({ sheetId, gid, sheet }) {
    const attempts = [
      { label: 'GViz JSON', fn: () => loadFromGvizJson(sheetId, gid, sheet) },
      { label: 'GViz JSONP', fn: () => loadFromGvizJsonp(sheetId, gid, sheet) },
      { label: 'CSV export', fn: () => loadFromCsv(sheetId, gid, sheet) },
      { label: 'OpenSheet API', fn: () => loadFromOpenSheet(sheetId, sheet) }
    ];

    const errors = [];
    for (const attempt of attempts) {
      try {
        const data = await attempt.fn();
        return { ...data, sourceLabel: attempt.label };
      } catch (err) {
        errors.push(`${attempt.label}: ${err.message}`);
      }
    }

    throw new Error(errors.join(' | '));
  }

  async function loadFromGvizJson(sheetId, gid, sheet) {
    const base = `https://docs.google.com/spreadsheets/d/${encodeURIComponent(sheetId)}/gviz/tq`;
    const url = new URL(base);
    url.searchParams.set('tqx', 'out:json');
    if (gid) url.searchParams.set('gid', gid);
    if (sheet) url.searchParams.set('sheet', sheet);

    const response = await fetch(url.toString(), { mode: 'cors' });
    if (!response.ok) {
      throw new Error(`HTTP ${response.status}`);
    }

    const text = await response.text();
    const table = parseGvizPayload(text);
    return gvizTableToDataset(table);
  }

  function loadFromGvizJsonp(sheetId, gid, sheet) {
    return new Promise((resolve, reject) => {
      const callback = `projStatJsonp_${Date.now()}_${Math.floor(Math.random() * 1000)}`;
      const base = `https://docs.google.com/spreadsheets/d/${encodeURIComponent(sheetId)}/gviz/tq`;
      const url = new URL(base);
      url.searchParams.set('tqx', 'out:json');
      url.searchParams.set('responseHandler', callback);
      if (gid) url.searchParams.set('gid', gid);
      if (sheet) url.searchParams.set('sheet', sheet);

      const script = document.createElement('script');
      const timeout = window.setTimeout(() => cleanup(new Error('JSONP timeout after 10s')), 10000);

      function cleanup(error, payload) {
        window.clearTimeout(timeout);
        if (script.parentNode) {
          script.parentNode.removeChild(script);
        }
        delete window[callback];
        if (error) {
          reject(error);
        } else {
          resolve(payload);
        }
      }

      window[callback] = (json) => {
        try {
          const data = gvizTableToDataset(json.table);
          cleanup(null, data);
        } catch (err) {
          cleanup(err);
        }
      };

      script.onerror = () => cleanup(new Error('JSONP script failed to load'));
      script.src = url.toString();
      document.head.appendChild(script);
    });
  }

  async function loadFromCsv(sheetId, gid, sheet) {
    const base = `https://docs.google.com/spreadsheets/d/${encodeURIComponent(sheetId)}/export`;
    const url = new URL(base);
    url.searchParams.set('format', 'csv');
    if (gid) url.searchParams.set('gid', gid);
    if (sheet) url.searchParams.set('sheet', sheet);

    const response = await fetch(url.toString(), { mode: 'cors' });
    if (!response.ok) {
      throw new Error(`HTTP ${response.status}`);
    }

    const csvText = await response.text();
    const rows = parseCsv(csvText);
    if (!rows.length) {
      throw new Error('CSV response was empty');
    }

    const headers = rows[0].map((h) => String(h || '').trim());
    const dataRows = rows.slice(1).map((row) => normalizeRowLength(row.map(stringValue), headers.length));
    return { headers, rows: dataRows };
  }

  async function loadFromOpenSheet(sheetId, sheetName) {
    const targetSheet = sheetName || 'Sheet1';
    const url = `https://opensheet.elk.sh/${encodeURIComponent(sheetId)}/${encodeURIComponent(targetSheet)}`;
    const response = await fetch(url, { mode: 'cors' });
    if (!response.ok) {
      throw new Error(`HTTP ${response.status}`);
    }

    const json = await response.json();
    if (!Array.isArray(json) || !json.length) {
      throw new Error('OpenSheet response was empty');
    }

    const headers = Object.keys(json[0]);
    const rows = json.map((rowObj) => headers.map((h) => stringValue(rowObj[h])));
    return { headers, rows };
  }

  function parseGvizPayload(text) {
    const start = text.indexOf('{');
    const end = text.lastIndexOf('}');
    if (start === -1 || end === -1 || end <= start) {
      throw new Error('GViz payload was not recognized');
    }

    const rawObject = text.slice(start, end + 1);
    let data;
    try {
      data = Function(`"use strict"; return (${rawObject});`)();
    } catch {
      throw new Error('Could not parse GViz response');
    }

    if (!data?.table) {
      throw new Error('GViz response had no table data');
    }

    return data.table;
  }

  function gvizTableToDataset(table) {
    const headers = (table.cols || []).map((col, index) => {
      const label = stringValue(col?.label).trim();
      const id = stringValue(col?.id).trim();
      return label || id || `Column ${index + 1}`;
    });

    const rows = (table.rows || []).map((row) => {
      const cells = Array.isArray(row?.c) ? row.c : [];
      const values = cells.map((cell) => {
        if (!cell) return '';
        if (cell.f != null && String(cell.f).trim() !== '') return String(cell.f).trim();
        if (cell.v instanceof Date) return cell.v.toLocaleDateString();
        return stringValue(cell.v);
      });
      return normalizeRowLength(values, headers.length);
    });

    return { headers, rows };
  }

  function parseCsv(text) {
    const rows = [];
    let row = [];
    let current = '';
    let inQuotes = false;

    for (let i = 0; i < text.length; i += 1) {
      const char = text[i];
      const next = text[i + 1];

      if (char === '"') {
        if (inQuotes && next === '"') {
          current += '"';
          i += 1;
        } else {
          inQuotes = !inQuotes;
        }
      } else if (char === ',' && !inQuotes) {
        row.push(current);
        current = '';
      } else if ((char === '\n' || char === '\r') && !inQuotes) {
        if (char === '\r' && next === '\n') {
          i += 1;
        }
        row.push(current);
        rows.push(row);
        row = [];
        current = '';
      } else {
        current += char;
      }
    }

    row.push(current);
    rows.push(row);

    return rows.filter((r) => r.some((cell) => String(cell).trim() !== ''));
  }

  function findColumnIndexes(headers) {
    const normalizedHeaders = headers.map(normalizeHeader);

    const indexes = {
      system: findAliasIndex(normalizedHeaders, ALIASES.system),
      milestone: findAliasIndex(normalizedHeaders, ALIASES.milestone),
      developer: findAliasIndex(normalizedHeaders, ALIASES.developer),
      manager: findAliasIndex(normalizedHeaders, ALIASES.manager)
    };

    if (indexes.system < 0 && headers.length) {
      indexes.system = 0;
    }

    return indexes;
  }

  function findAliasIndex(normalizedHeaders, aliases) {
    const normalizedAliases = aliases.map(normalizeHeader);
    return normalizedHeaders.findIndex((header) => normalizedAliases.includes(header));
  }

  function normalizeHeader(value) {
    return String(value || '')
      .toLowerCase()
      .replace(/\([^)]*\)/g, ' ')
      .replace(/[^a-z0-9]+/g, ' ')
      .trim()
      .replace(/\s+/g, ' ');
  }

  function applyFilters() {
    if (!state.headers.length) {
      renderTable();
      updateKpis();
      return;
    }

    const systemValue = el.systemFilter.value;
    const milestoneValue = el.milestoneFilter.value;
    const query = el.search.value.trim().toLowerCase();

    state.filteredRows = state.rows.filter((row) => {
      const rowSystem = getCell(row, state.indexes.system);
      const rowMilestone = getCell(row, state.indexes.milestone);

      const matchesSystem = !systemValue || rowSystem === systemValue;
      const matchesMilestone = !milestoneValue || rowMilestone === milestoneValue;

      const searchHaystack = [
        rowSystem,
        rowMilestone,
        getCell(row, state.indexes.developer),
        getCell(row, state.indexes.manager)
      ].join(' ').toLowerCase();
      const matchesSearch = !query || searchHaystack.includes(query);

      return matchesSystem && matchesMilestone && matchesSearch;
    });

    renderTable();
    updateKpis();
  }

  function populateFilterOptions() {
    const systems = uniqueValues(state.rows, state.indexes.system);
    const milestones = uniqueValues(state.rows, state.indexes.milestone);

    replaceOptions(el.systemFilter, 'All Systems', systems);
    replaceOptions(el.milestoneFilter, 'All Milestones', milestones);
  }

  function renderTable() {
    el.tableHead.innerHTML = '';
    el.tableBody.innerHTML = '';

    if (!state.headers.length) {
      const row = document.createElement('tr');
      row.innerHTML = '<td class="empty-state" colspan="1">Paste a Google Sheet and click View Data.</td>';
      el.tableBody.appendChild(row);
      return;
    }

    const headerRow = document.createElement('tr');
    state.headers.forEach((header) => {
      const th = document.createElement('th');
      th.textContent = header;
      headerRow.appendChild(th);
    });
    el.tableHead.appendChild(headerRow);

    if (!state.filteredRows.length) {
      const row = document.createElement('tr');
      const td = document.createElement('td');
      td.colSpan = state.headers.length;
      td.className = 'empty-state';
      td.textContent = 'No results found';
      row.appendChild(td);
      el.tableBody.appendChild(row);
      return;
    }

    const fragment = document.createDocumentFragment();
    state.filteredRows.forEach((rowData) => {
      const row = document.createElement('tr');
      state.headers.forEach((_, index) => {
        const td = document.createElement('td');
        td.textContent = rowData[index] || '';
        row.appendChild(td);
      });
      fragment.appendChild(row);
    });
    el.tableBody.appendChild(fragment);
  }

  function updateKpis() {
    const visibleRows = state.filteredRows;
    const systems = new Set();

    visibleRows.forEach((row) => {
      const systemValue = getCell(row, state.indexes.system);
      if (systemValue) {
        systems.add(systemValue);
      }
    });

    el.kpiProjects.textContent = String(systems.size);
    el.kpiMilestones.textContent = String(visibleRows.length);
  }

  function uniqueValues(rows, index) {
    if (index < 0) return [];
    const values = new Set();
    rows.forEach((row) => {
      const value = getCell(row, index);
      if (value) {
        values.add(value);
      }
    });
    return Array.from(values).sort((a, b) => a.localeCompare(b));
  }

  function replaceOptions(select, allLabel, values) {
    select.innerHTML = '';

    const all = document.createElement('option');
    all.value = '';
    all.textContent = allLabel;
    select.appendChild(all);

    values.forEach((value) => {
      const option = document.createElement('option');
      option.value = value;
      option.textContent = value;
      select.appendChild(option);
    });
  }

  function setFeedback(message, type) {
    el.feedback.textContent = message;
    el.feedback.classList.remove('error', 'success');

    if (type === 'error') {
      el.feedback.classList.add('error');
    } else {
      el.feedback.classList.add('success');
    }
  }

  function normalizeRowLength(row, expectedLength) {
    const out = row.slice(0, expectedLength);
    while (out.length < expectedLength) {
      out.push('');
    }
    return out;
  }

  function getCell(row, index) {
    if (index < 0) return '';
    return stringValue(row[index]).trim();
  }

  function stringValue(value) {
    if (value == null) return '';
    return String(value);
  }
})();