// Format King - Main Application
class FormatKing {
    constructor() {
        this.data = [];
        this.headers = [];
        this.sortColumn = -1;
        this.sortDirection = 'asc';
        this.filteredData = [];
        this.tables = []; // For multiple tables (Databricks format)
        this.currentTableIndex = 0;

        this.initElements();
        this.initEventListeners();
    }

    initElements() {
        // Tabs
        this.tabButtons = document.querySelectorAll('.tab-btn');
        this.tabContents = document.querySelectorAll('.tab-content');

        // Input elements
        this.pasteInput = document.getElementById('paste-input');
        this.delimiterSelect = document.getElementById('delimiter-select');
        this.firstRowHeader = document.getElementById('first-row-header');
        this.formatBtn = document.getElementById('format-btn');

        // File upload elements
        this.dropZone = document.getElementById('drop-zone');
        this.fileInput = document.getElementById('file-input');
        this.fileNameDisplay = document.getElementById('file-name');

        // Output elements
        this.outputSection = document.getElementById('output-section');
        this.dataTable = document.getElementById('data-table');
        this.tableHead = document.getElementById('table-head');
        this.tableBody = document.getElementById('table-body');
        this.emptyState = document.getElementById('empty-state');
        this.rowCount = document.getElementById('row-count');
        this.colCount = document.getElementById('col-count');

        // Controls
        this.searchInput = document.getElementById('search-input');
        this.copyBtn = document.getElementById('copy-btn');
        this.copyRichBtn = document.getElementById('copy-rich-btn');
        this.exportCsvBtn = document.getElementById('export-csv-btn');
        this.exportJsonBtn = document.getElementById('export-json-btn');

        // Table selector for multiple tables
        this.tableSelector = document.getElementById('table-selector');
        this.tableSelectorContainer = document.getElementById('table-selector-container');
    }

    initEventListeners() {
        // Tab switching
        this.tabButtons.forEach(btn => {
            btn.addEventListener('click', () => this.switchTab(btn.dataset.tab));
        });

        // Format button
        this.formatBtn.addEventListener('click', () => this.formatFromText());

        // File upload
        this.fileInput.addEventListener('change', (e) => this.handleFileSelect(e));

        // Drag and drop
        this.dropZone.addEventListener('dragover', (e) => {
            e.preventDefault();
            this.dropZone.classList.add('dragover');
        });

        this.dropZone.addEventListener('dragleave', () => {
            this.dropZone.classList.remove('dragover');
        });

        this.dropZone.addEventListener('drop', (e) => {
            e.preventDefault();
            this.dropZone.classList.remove('dragover');
            const file = e.dataTransfer.files[0];
            if (file) this.processFile(file);
        });

        // Search
        this.searchInput.addEventListener('input', () => this.filterTable());

        // Export buttons
        this.copyBtn.addEventListener('click', () => this.copyToClipboard());
        this.copyRichBtn.addEventListener('click', () => this.copyRichTable());
        this.exportCsvBtn.addEventListener('click', () => this.exportCSV());
        this.exportJsonBtn.addEventListener('click', () => this.exportJSON());

        // Table selector for multiple tables
        if (this.tableSelector) {
            this.tableSelector.addEventListener('change', (e) => this.switchToTable(parseInt(e.target.value)));
        }

        // Keyboard shortcut for format
        this.pasteInput.addEventListener('keydown', (e) => {
            if (e.ctrlKey && e.key === 'Enter') {
                this.formatFromText();
            }
        });
    }

    switchTab(tabId) {
        this.tabButtons.forEach(btn => {
            btn.classList.toggle('active', btn.dataset.tab === tabId);
        });

        this.tabContents.forEach(content => {
            content.classList.toggle('active', content.id === `${tabId}-tab`);
        });
    }

    detectDelimiter(text) {
        const lines = text.trim().split('\n').slice(0, 5);
        const delimiters = [',', ';', '\t', '|'];
        const counts = {};

        delimiters.forEach(d => {
            counts[d] = lines.map(line => (line.match(new RegExp(d === '|' ? '\\|' : d, 'g')) || []).length);
        });

        // Find delimiter with most consistent count across lines
        let bestDelimiter = ',';
        let bestScore = 0;

        delimiters.forEach(d => {
            const avg = counts[d].reduce((a, b) => a + b, 0) / counts[d].length;
            const variance = counts[d].reduce((sum, val) => sum + Math.pow(val - avg, 2), 0) / counts[d].length;
            const score = avg > 0 ? avg / (variance + 1) : 0;

            if (score > bestScore) {
                bestScore = score;
                bestDelimiter = d;
            }
        });

        return bestDelimiter;
    }

    // Detect if text is Databricks/SQL ASCII table format or Unicode box table
    isDatabricksFormat(text) {
        // Look for patterns like +----+----+ and |value|value|
        // Use \s* to handle leading whitespace
        const hasBorderLines = /^\s*\+[-+]+\+/m.test(text);
        const hasDataLines = /^\s*\|.+\|/m.test(text);
        const hasTableHeader = /^=+\s*\n\s*TABLE:/m.test(text) || hasBorderLines;

        // Also check for Unicode box-drawing characters (Claude terminal output)
        // Use \s* to handle leading whitespace
        const hasUnicodeBorders = /[┌┐└┘├┤┬┴┼─│]/m.test(text);
        const hasUnicodeDataLines = /^\s*│.+│/m.test(text);

        return (hasBorderLines && hasDataLines) || hasTableHeader || (hasUnicodeBorders && hasUnicodeDataLines);
    }

    // Parse Databricks/SQL ASCII table format
    parseDatabricksFormat(text) {
        const tables = [];
        const lines = text.split('\n');

        let currentTableName = null;
        let currentRows = [];
        let inTable = false;
        let headerParsed = false;
        let currentHeaders = [];

        for (let i = 0; i < lines.length; i++) {
            const line = lines[i].trim();

            // Check for table name header (e.g., "TABLE: blasts")
            const tableNameMatch = line.match(/^TABLE:\s*(.+)$/);
            if (tableNameMatch) {
                // Save previous table if exists
                if (currentRows.length > 0 && currentHeaders.length > 0) {
                    tables.push({
                        name: currentTableName || `Table ${tables.length + 1}`,
                        headers: currentHeaders,
                        data: currentRows
                    });
                }
                currentTableName = tableNameMatch[1].trim();
                currentRows = [];
                currentHeaders = [];
                headerParsed = false;
                inTable = false;
                continue;
            }

            // Clean line - remove trailing commas that might be artifacts
            const cleanLine = line.replace(/,+$/, '').trim();

            // Skip separator lines (====== or +----+ or Unicode borders ┌─┬─┐ ├─┼─┤ └─┴─┘)
            // Also handle lines that might not have closing characters
            if (/^=+$/.test(cleanLine) || /^\+[-+]+\+?$/.test(cleanLine) || /^[┌├└][─┬┼┴]+[┐┤┘]?$/.test(cleanLine)) {
                if (inTable && headerParsed) {
                    // This might be end of table or just a separator
                }
                inTable = true;
                continue;
            }

            // Parse data lines (|value|value|value| or │value│value│value│)
            // Handle lines that might be missing closing | or │
            const isAsciiDataLine = cleanLine.startsWith('|');
            const isUnicodeDataLine = cleanLine.startsWith('│');

            if (isAsciiDataLine || isUnicodeDataLine) {
                const delimiter = isUnicodeDataLine ? '│' : '|';
                let content = cleanLine;

                // Remove leading delimiter
                content = content.slice(1);

                // Remove trailing delimiter if present
                if (content.endsWith(delimiter)) {
                    content = content.slice(0, -1);
                }

                const cells = content
                    .split(delimiter)
                    .map(cell => cell.trim());

                if (!headerParsed) {
                    currentHeaders = cells;
                    headerParsed = true;
                } else {
                    currentRows.push(cells);
                }
            }
        }

        // Don't forget the last table
        if (currentRows.length > 0 && currentHeaders.length > 0) {
            tables.push({
                name: currentTableName || `Table ${tables.length + 1}`,
                headers: currentHeaders,
                data: currentRows
            });
        }

        // If no named tables found, try parsing as single table
        if (tables.length === 0) {
            const singleTable = this.parseSingleAsciiTable(text);
            if (singleTable) {
                tables.push(singleTable);
            }
        }

        return tables;
    }

    // Parse a single ASCII table without TABLE: headers
    parseSingleAsciiTable(text) {
        const lines = text.split('\n');
        let headers = [];
        let data = [];
        let headerParsed = false;
        let tableName = null;

        // Check for a title line before the table (e.g., "Table Name Decoder" or "Summary")
        for (let i = 0; i < lines.length; i++) {
            const trimmed = lines[i].trim();
            // If we find a non-empty line that's not a border, it might be a title
            if (trimmed && !/^[┌├└\+]/.test(trimmed) && !/^[│|]/.test(trimmed) && !/^=+$/.test(trimmed)) {
                // Check if any following line (within next 3 lines) is a border (indicating this is a title)
                for (let j = 1; j <= 3 && i + j < lines.length; j++) {
                    const nextLine = lines[i + j]?.trim() || '';
                    if (/^[┌\+]/.test(nextLine)) {
                        tableName = trimmed;
                        break;
                    }
                    // Stop if we hit a data line
                    if (/^[│|]/.test(nextLine)) break;
                }
                if (tableName) break;
            }
            // Stop looking once we hit the table
            if (/^[┌\+│|]/.test(trimmed)) break;
        }

        for (const line of lines) {
            // Clean line - remove trailing commas that might be artifacts
            const trimmed = line.trim().replace(/,+$/, '').trim();

            // Skip separator lines (ASCII or Unicode) - handle missing closing chars
            if (/^\+[-+]+\+?$/.test(trimmed) || /^=+$/.test(trimmed) || /^[┌├└][─┬┼┴]+[┐┤┘]?$/.test(trimmed)) {
                continue;
            }

            // Parse data lines (ASCII | or Unicode │) - handle missing closing delimiter
            const isAsciiDataLine = trimmed.startsWith('|');
            const isUnicodeDataLine = trimmed.startsWith('│');

            if (isAsciiDataLine || isUnicodeDataLine) {
                const delimiter = isUnicodeDataLine ? '│' : '|';
                let content = trimmed;

                // Remove leading delimiter
                content = content.slice(1);

                // Remove trailing delimiter if present
                if (content.endsWith(delimiter)) {
                    content = content.slice(0, -1);
                }

                const cells = content
                    .split(delimiter)
                    .map(cell => cell.trim());

                if (!headerParsed) {
                    headers = cells;
                    headerParsed = true;
                } else {
                    data.push(cells);
                }
            }
        }

        if (headers.length > 0) {
            return { name: tableName || 'Table 1', headers, data };
        }
        return null;
    }

    // Detect Markdown table format
    isMarkdownTable(text) {
        // Look for separator row with dashes between pipes: | --- | --- |
        const separatorPattern = /^\|?\s*:?-{2,}:?\s*(\|\s*:?-{2,}:?\s*)+\|?\s*$/m;
        if (!separatorPattern.test(text)) return false;
        // Must also have at least one data row with | delimiters
        const lines = text.split('\n').filter(l => l.trim());
        const dataLines = lines.filter(l => l.includes('|') && !/^[\s|:*-]+$/.test(l.trim()));
        return dataLines.length >= 1;
    }

    // Parse Markdown table(s)
    parseMarkdownTable(text) {
        const tables = [];
        const blocks = text.split(/\n\s*\n/);

        for (const block of blocks) {
            const lines = block.split('\n').filter(l => l.trim());
            const sepPattern = /^\|?\s*:?-{2,}:?\s*(\|\s*:?-{2,}:?\s*)+\|?\s*$/;
            const sepIndex = lines.findIndex(l => sepPattern.test(l.trim()));
            if (sepIndex < 1) continue;

            const splitRow = (line) => {
                let s = line.trim();
                if (s.startsWith('|')) s = s.slice(1);
                if (s.endsWith('|')) s = s.slice(0, -1);
                return s.split('|').map(c => c.trim());
            };

            const headers = splitRow(lines[sepIndex - 1]);
            const data = [];
            for (let i = sepIndex + 1; i < lines.length; i++) {
                if (!lines[i].includes('|')) continue;
                data.push(splitRow(lines[i]));
            }

            if (headers.length > 0 && data.length > 0) {
                tables.push({
                    name: `Table ${tables.length + 1}`,
                    headers,
                    data
                });
            }
        }
        return tables;
    }

    // Detect JSON array of objects
    isJSONArray(text) {
        const trimmed = text.trimStart();
        if (!trimmed.startsWith('[')) return false;
        try {
            const parsed = JSON.parse(trimmed);
            return Array.isArray(parsed) && parsed.length > 0 &&
                typeof parsed[0] === 'object' && parsed[0] !== null && !Array.isArray(parsed[0]);
        } catch {
            return false;
        }
    }

    // Parse JSON array of objects
    parseJSONArray(text) {
        const arr = JSON.parse(text.trimStart());
        // Collect all keys preserving insertion order
        const keySet = new Set();
        for (const obj of arr) {
            if (typeof obj === 'object' && obj !== null && !Array.isArray(obj)) {
                Object.keys(obj).forEach(k => keySet.add(k));
            }
        }
        const headers = [...keySet];
        const data = arr.filter(obj => typeof obj === 'object' && obj !== null && !Array.isArray(obj))
            .map(obj => headers.map(h => {
                const val = obj[h];
                if (val === undefined || val === null) return '';
                if (typeof val === 'object') return JSON.stringify(val);
                return String(val);
            }));
        return { name: 'JSON Data', headers, data };
    }

    // Detect fixed-width / space-aligned table
    isFixedWidthTable(text) {
        const lines = text.split('\n').filter(l => l.trim());
        if (lines.length < 3) return false;

        // Find column boundaries: positions where 2+ spaces appear across most lines
        const maxLen = Math.max(...lines.map(l => l.length));
        if (maxLen < 5) return false;

        // Skip obvious underline rows for boundary detection
        const contentLines = lines.filter(l => !/^[\s\-=]+$/.test(l));
        if (contentLines.length < 2) return false;

        const spaceCount = new Array(maxLen).fill(0);
        for (const line of contentLines) {
            for (let i = 0; i < maxLen; i++) {
                const ch = i < line.length ? line[i] : ' ';
                if (ch === ' ') spaceCount[i]++;
            }
        }

        // Find positions that are spaces in 80%+ of lines and form 2+ wide gaps
        const threshold = contentLines.length * 0.8;
        let boundaries = 0;
        let inGap = false;
        let gapWidth = 0;
        for (let i = 0; i < maxLen; i++) {
            if (spaceCount[i] >= threshold) {
                gapWidth++;
                if (gapWidth >= 2 && !inGap) {
                    boundaries++;
                    inGap = true;
                }
            } else {
                inGap = false;
                gapWidth = 0;
            }
        }

        // Need at least 1 column boundary to be a fixed-width table
        return boundaries >= 1;
    }

    // Parse fixed-width / space-aligned table
    parseFixedWidthTable(text) {
        const lines = text.split('\n').filter(l => l.trim());
        if (lines.length < 2) return null;

        // Pad all lines to the same length
        const maxLen = Math.max(...lines.map(l => l.length));
        const padded = lines.map(l => l.padEnd(maxLen));

        // Detect underline rows (--- or === rows)
        const isUnderline = (l) => /^[\s\-=]+$/.test(l) && /[-=]{2,}/.test(l);

        // Use only content lines for boundary detection
        const contentLines = padded.filter(l => !isUnderline(l));
        if (contentLines.length < 2) return null;

        // Build space map
        const spaceCount = new Array(maxLen).fill(0);
        for (const line of contentLines) {
            for (let i = 0; i < maxLen; i++) {
                if (line[i] === ' ') spaceCount[i]++;
            }
        }

        // Find column boundary ranges (2+ consecutive spaces in 80%+ of lines)
        const threshold = contentLines.length * 0.8;
        const cuts = []; // split positions
        let inGap = false;
        let gapStart = -1;
        for (let i = 0; i < maxLen; i++) {
            if (spaceCount[i] >= threshold) {
                if (!inGap) { gapStart = i; inGap = true; }
            } else {
                if (inGap) {
                    const gapWidth = i - gapStart;
                    if (gapWidth >= 2) {
                        // Cut at the middle of the gap
                        cuts.push(Math.floor((gapStart + i) / 2));
                    }
                    inGap = false;
                }
            }
        }

        if (cuts.length === 0) return null;

        // Slice each line into cells
        const sliceLine = (line) => {
            const cells = [];
            let prev = 0;
            for (const cut of cuts) {
                cells.push(line.slice(prev, cut).trim());
                prev = cut;
            }
            cells.push(line.slice(prev).trim());
            return cells;
        };

        // First content line is headers, rest are data
        const headers = sliceLine(contentLines[0]);
        const data = [];
        for (let i = 1; i < contentLines.length; i++) {
            const row = sliceLine(contentLines[i]);
            // Skip rows that are all empty
            if (row.some(c => c !== '')) {
                data.push(row);
            }
        }

        if (data.length === 0) return null;
        return { name: 'Table 1', headers, data };
    }

    // Switch to a different table (for multi-table support)
    switchToTable(index) {
        // index -1 means "All Tables"
        if (index === -1) {
            this.currentTableIndex = -1;
            // Combine all tables into one view with table name as first column
            this.headers = ['Table', ...this.getCommonHeaders()];
            this.data = [];
            for (const table of this.tables) {
                for (const row of table.data) {
                    // Pad row to match common headers length
                    const paddedRow = this.padRowToHeaders(row, table.headers);
                    this.data.push([table.name, ...paddedRow]);
                }
            }
            this.filteredData = [...this.data];
            this.sortColumn = -1;
            this.renderTable();
        } else if (index >= 0 && index < this.tables.length) {
            this.currentTableIndex = index;
            const table = this.tables[index];
            this.headers = table.headers;
            this.data = table.data;
            this.filteredData = [...this.data];
            this.sortColumn = -1;
            this.renderTable();
        }
    }

    // Get common headers across all tables (union of all headers)
    getCommonHeaders() {
        if (this.tables.length === 0) return [];

        // Check if all tables have the same headers
        const firstHeaders = this.tables[0].headers;
        const allSame = this.tables.every(t =>
            t.headers.length === firstHeaders.length &&
            t.headers.every((h, i) => h === firstHeaders[i])
        );

        if (allSame) {
            return firstHeaders;
        }

        // If headers differ, collect unique headers in order
        const seen = new Set();
        const allHeaders = [];
        for (const table of this.tables) {
            for (const header of table.headers) {
                if (!seen.has(header)) {
                    seen.add(header);
                    allHeaders.push(header);
                }
            }
        }
        return allHeaders;
    }

    // Pad a row to match the common headers
    padRowToHeaders(row, tableHeaders) {
        const commonHeaders = this.getCommonHeaders();
        const result = new Array(commonHeaders.length).fill('');

        tableHeaders.forEach((header, i) => {
            const commonIndex = commonHeaders.indexOf(header);
            if (commonIndex !== -1 && i < row.length) {
                result[commonIndex] = row[i];
            }
        });

        return result;
    }

    // Update table selector dropdown
    updateTableSelector() {
        if (!this.tableSelector || !this.tableSelectorContainer) return;

        if (this.tables.length > 1) {
            this.tableSelectorContainer.classList.remove('hidden');
            const totalRows = this.tables.reduce((sum, t) => sum + t.data.length, 0);
            const options = [
                `<option value="-1">All Tables (${totalRows} rows)</option>`,
                ...this.tables.map((table, i) =>
                    `<option value="${i}">${table.name} (${table.data.length} rows)</option>`
                )
            ];
            this.tableSelector.innerHTML = options.join('');
        } else {
            this.tableSelectorContainer.classList.add('hidden');
        }
    }

    parseCSV(text, delimiter) {
        const rows = [];
        let currentRow = [];
        let currentCell = '';
        let inQuotes = false;

        for (let i = 0; i < text.length; i++) {
            const char = text[i];
            const nextChar = text[i + 1];

            if (inQuotes) {
                if (char === '"' && nextChar === '"') {
                    currentCell += '"';
                    i++;
                } else if (char === '"') {
                    inQuotes = false;
                } else {
                    currentCell += char;
                }
            } else {
                if (char === '"') {
                    inQuotes = true;
                } else if (char === delimiter) {
                    currentRow.push(currentCell.trim());
                    currentCell = '';
                } else if (char === '\n' || (char === '\r' && nextChar === '\n')) {
                    currentRow.push(currentCell.trim());
                    if (currentRow.some(cell => cell !== '')) {
                        rows.push(currentRow);
                    }
                    currentRow = [];
                    currentCell = '';
                    if (char === '\r') i++;
                } else if (char !== '\r') {
                    currentCell += char;
                }
            }
        }

        // Push last row if exists
        if (currentCell || currentRow.length > 0) {
            currentRow.push(currentCell.trim());
            if (currentRow.some(cell => cell !== '')) {
                rows.push(currentRow);
            }
        }

        return rows;
    }

    formatFromText() {
        const text = this.pasteInput.value.trim();
        if (!text) {
            this.showToast('Please enter some data first');
            return;
        }

        // Check if it's Databricks/SQL ASCII table format
        if (this.isDatabricksFormat(text)) {
            const tables = this.parseDatabricksFormat(text);
            if (tables.length > 0) {
                this.tables = tables;
                this.updateTableSelector();
                // Default to "All Tables" view (-1) when multiple tables
                const defaultView = tables.length > 1 ? -1 : 0;
                this.switchToTable(defaultView);
                if (this.tableSelector) {
                    this.tableSelector.value = defaultView.toString();
                }
                const totalRows = tables.reduce((sum, t) => sum + t.data.length, 0);
                this.showToast(`Loaded ${tables.length} table(s), ${totalRows} total rows`);
                return;
            }
        }

        // Check for Markdown tables
        if (this.isMarkdownTable(text)) {
            const tables = this.parseMarkdownTable(text);
            if (tables.length > 0) {
                this.tables = tables;
                this.updateTableSelector();
                const defaultView = tables.length > 1 ? -1 : 0;
                this.switchToTable(defaultView);
                if (this.tableSelector) this.tableSelector.value = defaultView.toString();
                const totalRows = tables.reduce((sum, t) => sum + t.data.length, 0);
                this.showToast(`Loaded ${tables.length} Markdown table(s), ${totalRows} total rows`);
                return;
            }
        }

        // Check for JSON array
        if (this.isJSONArray(text)) {
            const table = this.parseJSONArray(text);
            if (table) {
                this.tables = [table];
                this.updateTableSelector();
                this.switchToTable(0);
                this.showToast(`Loaded JSON data: ${table.data.length} rows`);
                return;
            }
        }

        // Check for fixed-width tables
        if (this.isFixedWidthTable(text)) {
            const table = this.parseFixedWidthTable(text);
            if (table) {
                this.tables = [table];
                this.updateTableSelector();
                this.switchToTable(0);
                this.showToast(`Loaded fixed-width table: ${table.data.length} rows`);
                return;
            }
        }

        // Fall back to CSV parsing
        let delimiter = this.delimiterSelect.value;
        if (delimiter === 'auto') {
            delimiter = this.detectDelimiter(text);
        } else if (delimiter === '\\t') {
            delimiter = '\t';
        }

        const rows = this.parseCSV(text, delimiter);
        this.tables = []; // Clear multi-table mode
        this.updateTableSelector();
        this.processData(rows);
    }

    handleFileSelect(e) {
        const file = e.target.files[0];
        if (file) this.processFile(file);
    }

    processFile(file) {
        this.fileNameDisplay.textContent = `Selected: ${file.name}`;

        const reader = new FileReader();
        reader.onload = (e) => {
            const text = e.target.result;
            const delimiter = this.detectDelimiter(text);
            const rows = this.parseCSV(text, delimiter);
            this.processData(rows);
        };
        reader.readAsText(file);
    }

    processData(rows) {
        if (rows.length === 0) {
            this.showToast('No data found');
            return;
        }

        // Normalize row lengths
        const maxCols = Math.max(...rows.map(r => r.length));
        rows = rows.map(row => {
            while (row.length < maxCols) row.push('');
            return row;
        });

        if (this.firstRowHeader.checked && rows.length > 0) {
            this.headers = rows[0];
            this.data = rows.slice(1);
        } else {
            this.headers = rows[0].map((_, i) => `Column ${i + 1}`);
            this.data = rows;
        }

        this.filteredData = [...this.data];
        this.sortColumn = -1;
        this.renderTable();
        this.showToast(`Loaded ${this.data.length} rows`);
    }

    renderTable() {
        // Update counts
        this.rowCount.textContent = `${this.filteredData.length} rows`;
        this.colCount.textContent = `${this.headers.length} columns`;

        // Render headers
        this.tableHead.innerHTML = `
            <tr>
                ${this.headers.map((h, i) => `
                    <th onclick="app.sortTable(${i})" class="${this.sortColumn === i ? 'sorted' : ''}">
                        ${this.escapeHtml(h)}
                        <span class="sort-indicator">${this.sortColumn === i ? (this.sortDirection === 'asc' ? '▲' : '▼') : '⇅'}</span>
                    </th>
                `).join('')}
            </tr>
        `;

        // Render body
        const searchTerm = this.searchInput.value.toLowerCase();
        this.tableBody.innerHTML = this.filteredData.map(row => `
            <tr>
                ${row.map(cell => `<td>${this.highlightText(this.escapeHtml(cell), searchTerm)}</td>`).join('')}
            </tr>
        `).join('');

        // Show table, hide empty state
        this.dataTable.classList.add('visible');
        this.emptyState.classList.add('hidden');
    }

    escapeHtml(text) {
        const div = document.createElement('div');
        div.textContent = text;
        return div.innerHTML;
    }

    highlightText(text, searchTerm) {
        if (!searchTerm) return text;
        const regex = new RegExp(`(${searchTerm.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')})`, 'gi');
        return text.replace(regex, '<span class="highlight">$1</span>');
    }

    sortTable(columnIndex) {
        if (this.sortColumn === columnIndex) {
            this.sortDirection = this.sortDirection === 'asc' ? 'desc' : 'asc';
        } else {
            this.sortColumn = columnIndex;
            this.sortDirection = 'asc';
        }

        this.filteredData.sort((a, b) => {
            let valA = a[columnIndex] || '';
            let valB = b[columnIndex] || '';

            // Try numeric sort
            const numA = parseFloat(valA);
            const numB = parseFloat(valB);

            if (!isNaN(numA) && !isNaN(numB)) {
                return this.sortDirection === 'asc' ? numA - numB : numB - numA;
            }

            // String sort
            valA = valA.toLowerCase();
            valB = valB.toLowerCase();

            if (valA < valB) return this.sortDirection === 'asc' ? -1 : 1;
            if (valA > valB) return this.sortDirection === 'asc' ? 1 : -1;
            return 0;
        });

        this.renderTable();
    }

    filterTable() {
        const searchTerm = this.searchInput.value.toLowerCase();

        if (!searchTerm) {
            this.filteredData = [...this.data];
        } else {
            this.filteredData = this.data.filter(row =>
                row.some(cell => cell.toLowerCase().includes(searchTerm))
            );
        }

        this.renderTable();
    }

    copyToClipboard() {
        const text = this.dataToCSV();
        navigator.clipboard.writeText(text).then(() => {
            this.showToast('Copied to clipboard!');
        }).catch(() => {
            this.showToast('Failed to copy');
        });
    }

    copyRichTable() {
        // Generate HTML table that Word/OneNote will recognize
        const html = this.generateRichHTML();
        const plainText = this.dataToCSV();

        // Use ClipboardItem API to copy both HTML and plain text formats
        const htmlBlob = new Blob([html], { type: 'text/html' });
        const textBlob = new Blob([plainText], { type: 'text/plain' });

        navigator.clipboard.write([
            new ClipboardItem({
                'text/html': htmlBlob,
                'text/plain': textBlob
            })
        ]).then(() => {
            this.showToast('Table copied! Paste into Word or OneNote');
        }).catch((err) => {
            // Fallback: try selecting and copying the table element
            console.error('Clipboard API failed:', err);
            this.fallbackRichCopy();
        });
    }

    generateRichHTML() {
        // If we have multiple tables and are in "All Tables" view, generate separate tables with headers
        if (this.tables.length > 1 && this.currentTableIndex === -1) {
            return this.generateMultiTableHTML();
        }

        return this.generateSingleTableHTML(this.headers, this.filteredData);
    }

    generateSingleTableHTML(headers, data, tableName = null) {
        const tableStyle = `
            border-collapse: collapse;
            font-family: Calibri, Arial, sans-serif;
            font-size: 11pt;
            width: 100%;
            margin-bottom: 20px;
        `.replace(/\s+/g, ' ').trim();

        const headerCellStyle = `
            border: 1px solid #5B9BD5;
            background-color: #5B9BD5;
            color: white;
            font-weight: bold;
            padding: 8px 12px;
            text-align: left;
        `.replace(/\s+/g, ' ').trim();

        const cellStyle = `
            border: 1px solid #DDDDDD;
            padding: 8px 12px;
            text-align: left;
        `.replace(/\s+/g, ' ').trim();

        const altRowStyle = `
            border: 1px solid #DDDDDD;
            padding: 8px 12px;
            text-align: left;
            background-color: #F2F2F2;
        `.replace(/\s+/g, ' ').trim();

        const tableTitleStyle = `
            font-family: Calibri, Arial, sans-serif;
            font-size: 14pt;
            font-weight: bold;
            color: #2E75B6;
            padding: 10px 0;
            border-bottom: 3px solid #2E75B6;
            margin-bottom: 10px;
        `.replace(/\s+/g, ' ').trim();

        // Build header row
        const headerRow = `<tr>${headers.map(h =>
            `<th style="${headerCellStyle}">${this.escapeHtml(h)}</th>`
        ).join('')}</tr>`;

        // Build data rows with alternating colors
        const dataRows = data.map((row, index) => {
            const style = index % 2 === 0 ? cellStyle : altRowStyle;
            return `<tr>${row.map(cell =>
                `<td style="${style}">${this.escapeHtml(cell)}</td>`
            ).join('')}</tr>`;
        }).join('');

        const titleHTML = tableName ? `<div style="${tableTitleStyle}">TABLE: ${this.escapeHtml(tableName)}</div>` : '';

        return `
            ${titleHTML}
            <table style="${tableStyle}">
                <thead>${headerRow}</thead>
                <tbody>${dataRows}</tbody>
            </table>
        `.trim();
    }

    generateMultiTableHTML() {
        // Generate HTML with each table clearly separated
        const tablesHTML = this.tables.map(table => {
            return this.generateSingleTableHTML(table.headers, table.data, table.name);
        }).join('<br><br>');

        return `
            <html>
            <body>
            ${tablesHTML}
            </body>
            </html>
        `.trim();
    }

    fallbackRichCopy() {
        // Fallback method using document selection
        const tempDiv = document.createElement('div');
        tempDiv.innerHTML = this.generateRichHTML();
        tempDiv.style.position = 'absolute';
        tempDiv.style.left = '-9999px';
        document.body.appendChild(tempDiv);

        const range = document.createRange();
        range.selectNode(tempDiv);
        const selection = window.getSelection();
        selection.removeAllRanges();
        selection.addRange(range);

        try {
            document.execCommand('copy');
            this.showToast('Table copied! Paste into Word or OneNote');
        } catch (err) {
            this.showToast('Copy failed - try using Ctrl+C');
        }

        selection.removeAllRanges();
        document.body.removeChild(tempDiv);
    }

    dataToCSV() {
        const escape = (cell) => {
            if (cell.includes(',') || cell.includes('"') || cell.includes('\n')) {
                return `"${cell.replace(/"/g, '""')}"`;
            }
            return cell;
        };

        const headerRow = this.headers.map(escape).join(',');
        const dataRows = this.filteredData.map(row => row.map(escape).join(','));

        return [headerRow, ...dataRows].join('\n');
    }

    exportCSV() {
        const csv = this.dataToCSV();
        this.downloadFile(csv, 'data.csv', 'text/csv');
        this.showToast('CSV exported!');
    }

    exportJSON() {
        const json = this.filteredData.map(row => {
            const obj = {};
            this.headers.forEach((header, i) => {
                obj[header] = row[i];
            });
            return obj;
        });

        const jsonStr = JSON.stringify(json, null, 2);
        this.downloadFile(jsonStr, 'data.json', 'application/json');
        this.showToast('JSON exported!');
    }

    downloadFile(content, filename, mimeType) {
        const blob = new Blob([content], { type: mimeType });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = filename;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
    }

    showToast(message) {
        // Remove existing toast
        const existingToast = document.querySelector('.toast');
        if (existingToast) existingToast.remove();

        const toast = document.createElement('div');
        toast.className = 'toast';
        toast.textContent = message;
        document.body.appendChild(toast);

        // Trigger animation
        setTimeout(() => toast.classList.add('show'), 10);

        // Remove after delay
        setTimeout(() => {
            toast.classList.remove('show');
            setTimeout(() => toast.remove(), 300);
        }, 2500);
    }
}

// Initialize app
const app = new FormatKing();
