// Format King - Main Application
class FormatKing {
    constructor() {
        this.data = [];
        this.headers = [];
        this.sortColumn = -1;
        this.sortDirection = 'asc';
        this.filteredData = [];

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

        let delimiter = this.delimiterSelect.value;
        if (delimiter === 'auto') {
            delimiter = this.detectDelimiter(text);
        } else if (delimiter === '\\t') {
            delimiter = '\t';
        }

        const rows = this.parseCSV(text, delimiter);
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
        // Generate a self-contained HTML table with inline styles
        // Word and OneNote recognize this format and preserve the structure
        const tableStyle = `
            border-collapse: collapse;
            font-family: Calibri, Arial, sans-serif;
            font-size: 11pt;
            width: 100%;
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

        // Build header row
        const headerRow = `<tr>${this.headers.map(h =>
            `<th style="${headerCellStyle}">${this.escapeHtml(h)}</th>`
        ).join('')}</tr>`;

        // Build data rows with alternating colors
        const dataRows = this.filteredData.map((row, index) => {
            const style = index % 2 === 0 ? cellStyle : altRowStyle;
            return `<tr>${row.map(cell =>
                `<td style="${style}">${this.escapeHtml(cell)}</td>`
            ).join('')}</tr>`;
        }).join('');

        // Wrap in full HTML document for better compatibility
        return `
            <html>
            <body>
            <table style="${tableStyle}">
                <thead>${headerRow}</thead>
                <tbody>${dataRows}</tbody>
            </table>
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
