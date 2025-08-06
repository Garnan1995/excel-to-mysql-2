// js/dashboard.js

const ExcelDashboard = {
    engine: null,
    currentWorkbook: null,
    currentWorksheet: null,
    selectedCell: null,
    showFormulas: false,
    apiUrl: '../api/excel-api.php',
    
    /**
     * Initialize the dashboard
     */
    init() {
        this.engine = new ExcelEngine();
        this.loadWorkbooks();
        this.setupEventListeners();
        this.setupKeyboardShortcuts();
    },
    
    /**
     * Load workbooks from API
     */
    async loadWorkbooks() {
        try {
            const response = await fetch(`${this.apiUrl}?action=workbooks`);
            const workbooks = await response.json();
            
            const selector = document.getElementById('workbookSelector');
            selector.innerHTML = '<option value="">Select Workbook</option>';
            
            workbooks.forEach(wb => {
                const option = document.createElement('option');
                option.value = wb.id;
                option.textContent = wb.filename;
                selector.appendChild(option);
            });
            
            // Auto-select first workbook if available
            if (workbooks.length > 0) {
                selector.value = workbooks[0].id;
                this.loadWorkbook(workbooks[0].id);
            }
        } catch (error) {
            console.error('Failed to load workbooks:', error);
        }
    },
    
    /**
     * Load a specific workbook
     */
    async loadWorkbook(workbookId) {
        this.showLoading(true);
        this.currentWorkbook = workbookId;
        this.engine.workbookId = workbookId;
        
        try {
            // Load worksheets
            const response = await fetch(`${this.apiUrl}?action=worksheets&workbook_id=${workbookId}`);
            const worksheets = await response.json();
            
            // Create sheet tabs
            this.createSheetTabs(worksheets);
            
            // Load first worksheet
            if (worksheets.length > 0) {
                this.loadWorksheet(worksheets[0].id);
            }
            
            // Load named ranges
            this.loadNamedRanges(workbookId);
            
        } catch (error) {
            console.error('Failed to load workbook:', error);
        } finally {
            this.showLoading(false);
        }
    },
    
    /**
     * Create sheet tabs
     */
    createSheetTabs(worksheets) {
        const container = document.getElementById('sheetTabs');
        container.innerHTML = '';
        
        worksheets.forEach((ws, index) => {
            const tab = document.createElement('div');
            tab.className = 'sheet-tab';
            tab.textContent = ws.sheet_name;
            tab.dataset.worksheetId = ws.id;
            
            if (ws.tab_color) {
                tab.style.borderBottomColor = '#' + ws.tab_color;
            }
            
            tab.onclick = () => this.loadWorksheet(ws.id);
            
            if (index === 0) {
                tab.classList.add('active');
            }
            
            container.appendChild(tab);
        });
    },
    
    /**
     * Load worksheet data
     */
    async loadWorksheet(worksheetId) {
        this.showLoading(true);
        this.currentWorksheet = worksheetId;
        this.engine.worksheetId = worksheetId;
        
        // Update active tab
        document.querySelectorAll('.sheet-tab').forEach(tab => {
            tab.classList.toggle('active', tab.dataset.worksheetId == worksheetId);
        });
        
        try {
            // Load cells for visible range (optimize for performance)
            const response = await fetch(
                `${this.apiUrl}?action=cell-range&worksheet_id=${worksheetId}` +
                `&start_row=1&end_row=100&start_col=1&end_col=50`
            );
            const data = await response.json();
            
            // Load cells into engine
            this.engine.loadCells(data.cells);
            
            // Render spreadsheet
            this.renderSpreadsheet(data);
            
            // Update statistics
            this.updateStatistics(data.cells);
            
        } catch (error) {
            console.error('Failed to load worksheet:', error);
        } finally {
            this.showLoading(false);
        }
    },
    
    /**
     * Render spreadsheet grid
     */
    renderSpreadsheet(data) {
        const header = document.getElementById('spreadsheetHeader');
        const body = document.getElementById('spreadsheetBody');
        
        // Clear existing content
        header.innerHTML = '';
        body.innerHTML = '';
        
        // Create header row
        const headerRow = document.createElement('tr');
        headerRow.innerHTML = '<th class="row-header"></th>';
        
        for (let col = 1; col <= 50; col++) {
            const th = document.createElement('th');
            th.textContent = this.engine.numberToColumn(col);
            headerRow.appendChild(th);
        }
        header.appendChild(headerRow);
        
        // Create data rows
        for (let row = 1; row <= 100; row++) {
            const tr = document.createElement('tr');
            
            // Row header
            const rowHeader = document.createElement('th');
            rowHeader.className = 'row-header';
            rowHeader.textContent = row;
            tr.appendChild(rowHeader);
            
            // Data cells
            for (let col = 1; col <= 50; col++) {
                const td = document.createElement('td');
                const address = this.engine.numberToColumn(col) + row;
                td.dataset.address = address;
                
                const cellData = this.findCellData(data.cells, row, col);
                
                if (cellData) {
                    this.renderCell(td, cellData);
                }
                
                // Add click handler
                td.onclick = (e) => this.selectCell(address, e);
                td.ondblclick = (e) => this.editCell(address, e);
                
                tr.appendChild(td);
            }
            
            body.appendChild(tr);
        }
        
        // Apply merged cells
        this.applyMergedCells(data.merged_cells);
        
        // Apply data validations
        this.applyDataValidations(data.validations);
    },
    
    /**
     * Find cell data by row and column
     */
    findCellData(cells, row, col) {
        return cells.find(c => c.row_num == row && c.col_num == col);
    },
    
    /**
     * Render individual cell
     */
    renderCell(td, cellData) {
        // Apply value
        if (this.showFormulas && cellData.formula) {
            td.textContent = cellData.formula;
            td.classList.add('formula-cell');
        } else if (cellData.formula) {
            // Evaluate formula
            const result = this.engine.evaluateFormula(cellData.formula, cellData.cell_address);
            td.textContent = this.formatValue(result, cellData.number_format);
            td.classList.add('formula-cell');
        } else {
            td.textContent = this.formatValue(cellData.cell_value, cellData.number_format);
        }
        
        // Apply styles
        if (cellData.font_bold) td.classList.add('cell-bold');
        if (cellData.font_italic) td.classList.add('cell-italic');
        if (cellData.font_color) td.style.color = '#' + cellData.font_color;
        if (cellData.fill_color && cellData.fill_color !== 'FFFFFFFF') {
            td.style.backgroundColor = '#' + cellData.fill_color.substring(2);
        }
        
        // Apply alignment
        if (cellData.alignment_horizontal) {
            td.classList.add(`cell-${cellData.alignment_horizontal}`);
        }
        
        // Apply borders
        if (cellData.border_top) td.style.borderTop = `1px solid #000`;
        if (cellData.border_right) td.style.borderRight = `1px solid #000`;
        if (cellData.border_bottom) td.style.borderBottom = `1px solid #000`;
        if (cellData.border_left) td.style.borderLeft = `1px solid #000`;
    },
    
    /**
     * Format value based on number format
     */
    formatValue(value, format) {
        if (value === null || value === undefined || value === '') return '';
        
        // Handle error values
        if (typeof value === 'string' && value.startsWith('#')) {
            return value;
        }
        
        // Handle number formats
        if (format && !isNaN(value)) {
            // Currency format
            if (format.includes('$')) {
                return '$' + parseFloat(value).toFixed(2).replace(/\B(?=(\d{3})+(?!\d))/g, ',');
            }
            // Percentage format
            if (format.includes('%')) {
                return (parseFloat(value) * 100).toFixed(2) + '%';
            }
            // Number with comma separator
            if (format.includes('#,##0')) {
                return parseFloat(value).toFixed(2).replace(/\B(?=(\d{3})+(?!\d))/g, ',');
            }
        }
        
        return value;
    },
    
    /**
     * Select a cell
     */
    selectCell(address, event) {
        // Remove previous selection
        if (this.selectedCell) {
            document.querySelector(`[data-address="${this.selectedCell}"]`)?.classList.remove('selected');
        }
        
        // Add new selection
        this.selectedCell = address;
        const cell = document.querySelector(`[data-address="${address}"]`);
        if (cell) {
            cell.classList.add('selected');
            
            // Update cell reference display
            document.getElementById('cellReference').textContent = address;
            
            // Update formula bar
            const cellData = this.engine.cells[address];
            const formulaInput = document.getElementById('formulaInput');
            if (cellData) {
                formulaInput.value = cellData.formula || cellData.value || '';
            } else {
                formulaInput.value = '';
            }
            
            // Update status bar
            this.updateSelectionStats();
        }
    },
    
    /**
     * Edit a cell
     */
    editCell(address, event) {
        const td = document.querySelector(`[data-address="${address}"]`);
        if (!td) return;
        
        const cellData = this.engine.cells[address] || {};
        
        // Create input element
        const input = document.createElement('input');
        input.type = 'text';
        input.value = cellData.formula || cellData.value || '';
        
        // Replace cell content with input
        td.innerHTML = '';
        td.appendChild(input);
        input.focus();
        input.select();
        
        // Handle input events
        input.onblur = () => this.saveCell(address, input.value);
        input.onkeydown = (e) => {
            if (e.key === 'Enter') {
                this.saveCell(address, input.value);
                this.moveSelection('down');
            } else if (e.key === 'Escape') {
                this.renderCell(td, cellData);
            } else if (e.key === 'Tab') {
                e.preventDefault();
                this.saveCell(address, input.value);
                this.moveSelection(e.shiftKey ? 'left' : 'right');
            }
        };
    },
    
    /**
     * Save cell value
     */
    async saveCell(address, value) {
        const isFormula = value.startsWith('=');
        
        // Update engine
        this.engine.setCellValue(address, isFormula ? null : value, isFormula ? value : null);
        
        // Save to database
        try {
            const response = await fetch(`${this.apiUrl}?action=update-cell`, {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    worksheet_id: this.currentWorksheet,
                    cell_address: address,
                    value: isFormula ? null : value,
                    formula: isFormula ? value : null
                })
            });
            
            if (response.ok) {
                // Re-render the cell
                const td = document.querySelector(`[data-address="${address}"]`);
                if (td) {
                    const cellData = this.engine.cells[address];
                    this.renderCell(td, {
                        cell_address: address,
                        cell_value: cellData.value,
                        formula: cellData.formula
                    });
                }
                
                // Update last saved time
                document.getElementById('lastSaved').textContent = 
                    'Last saved: ' + new Date().toLocaleTimeString();
            }
        } catch (error) {
            console.error('Failed to save cell:', error);
        }
    },
    
    /**
     * Apply merged cells
     */
    applyMergedCells(mergedCells) {
        mergedCells.forEach(merge => {
            const range = merge.cell_range;
            const [start, end] = range.split(':');
            
            const startCell = this.engine.parseCellAddress(start);
            const endCell = this.engine.parseCellAddress(end);
            
            // Get the top-left cell
            const topLeftAddress = this.engine.numberToColumn(startCell.col) + startCell.row;
            const topLeftTd = document.querySelector(`[data-address="${topLeftAddress}"]`);
            
            if (topLeftTd) {
                const rowSpan = endCell.row - startCell.row + 1;
                const colSpan = endCell.col - startCell.col + 1;
                
                topLeftTd.rowSpan = rowSpan;
                topLeftTd.colSpan = colSpan;
                
                // Hide other cells in the merge range
                for (let row = startCell.row; row <= endCell.row; row++) {
                    for (let col = startCell.col; col <= endCell.col; col++) {
                        if (row !== startCell.row || col !== startCell.col) {
                            const address = this.engine.numberToColumn(col) + row;
                            const td = document.querySelector(`[data-address="${address}"]`);
                            if (td) {
                                td.style.display = 'none';
                            }
                        }
                    }
                }
            }
        });
    },
    
    /**
     * Apply data validations
     */
    applyDataValidations(validations) {
        validations.forEach(validation => {
            if (validation.validation_type === 'list' && validation.show_dropdown) {
                const range = validation.cell_range;
                // Parse range and apply dropdown to cells
                // This would add dropdown arrows to cells with list validation
            }
        });
    },
    
    /**
     * Setup event listeners
     */
    setupEventListeners() {
        // Workbook selector
        document.getElementById('workbookSelector').onchange = (e) => {
            if (e.target.value) {
                this.loadWorkbook(e.target.value);
            }
        };
        
        // Formula input
        document.getElementById('formulaInput').onkeydown = (e) => {
            if (e.key === 'Enter') {
                const value = e.target.value;
                if (this.selectedCell) {
                    this.saveCell(this.selectedCell, value);
                }
            }
        };
    },
    
    /**
     * Setup keyboard shortcuts
     */
    setupKeyboardShortcuts() {
        document.addEventListener('keydown', (e) => {
            // Ctrl+S to save
            if (e.ctrlKey && e.key === 's') {
                e.preventDefault();
                this.saveChanges();
            }
            
            // F2 to edit cell
            if (e.key === 'F2' && this.selectedCell) {
                e.preventDefault();
                this.editCell(this.selectedCell);
            }
            
            // Arrow keys for navigation
            if (!e.target.matches('input')) {
                if (e.key === 'ArrowUp') {
                    e.preventDefault();
                    this.moveSelection('up');
                } else if (e.key === 'ArrowDown') {
                    e.preventDefault();
                    this.moveSelection('down');
                } else if (e.key === 'ArrowLeft') {
                    e.preventDefault();
                    this.moveSelection('left');
                } else if (e.key === 'ArrowRight') {
                    e.preventDefault();
                    this.moveSelection('right');
                }
                
                // Delete key to clear cell
                if (e.key === 'Delete' && this.selectedCell) {
                    this.saveCell(this.selectedCell, '');
                }
            }
        });
    },
    
    /**
     * Move selection in a direction
     */
    moveSelection(direction) {
        if (!this.selectedCell) return;
        
        const current = this.engine.parseCellAddress(this.selectedCell);
        let newAddress = null;
        
        switch (direction) {
            case 'up':
                if (current.row > 1) {
                    newAddress = this.engine.numberToColumn(current.col) + (current.row - 1);
                }
                break;
            case 'down':
                if (current.row < 100) {
                    newAddress = this.engine.numberToColumn(current.col) + (current.row + 1);
                }
                break;
            case 'left':
                if (current.col > 1) {
                    newAddress = this.engine.numberToColumn(current.col - 1) + current.row;
                }
                break;
            case 'right':
                if (current.col < 50) {
                    newAddress = this.engine.numberToColumn(current.col + 1) + current.row;
                }
                break;
        }
        
        if (newAddress) {
            this.selectCell(newAddress);
            // Scroll into view if needed
            const cell = document.querySelector(`[data-address="${newAddress}"]`);
            if (cell) {
                cell.scrollIntoViewIfNeeded ? cell.scrollIntoViewIfNeeded() : cell.scrollIntoView();
            }
        }
    },
    
    /**
     * Load named ranges
     */
    async loadNamedRanges(workbookId) {
        try {
            const response = await fetch(`${this.apiUrl}?action=named-ranges&workbook_id=${workbookId}`);
            const ranges = await response.json();
            
            ranges.forEach(range => {
                this.engine.namedRanges[range.name] = range.cell_range;
            });
        } catch (error) {
            console.error('Failed to load named ranges:', error);
        }
    },
    
    /**
     * Toggle formula display
     */
    toggleFormulas() {
        this.showFormulas = !this.showFormulas;
        document.getElementById('formulaToggle').classList.toggle('active', this.showFormulas);
        
        // Re-render all cells
        document.querySelectorAll('#spreadsheetBody td').forEach(td => {
            const address = td.dataset.address;
            if (address && this.engine.cells[address]) {
                const cellData = this.engine.cells[address];
                this.renderCell(td, {
                    cell_address: address,
                    cell_value: cellData.value,
                    formula: cellData.formula,
                    ...cellData.style
                });
            }
        });
    },
    
    /**
     * Recalculate all formulas
     */
    recalculate() {
        document.getElementById('calculationMode').textContent = 'Calculating...';
        
        // Clear cache
        this.engine.cache = {};
        
        // Recalculate all formula cells
        Object.entries(this.engine.cells).forEach(([address, cell]) => {
            if (cell.formula) {
                const result = this.engine.evaluateFormula(cell.formula, address);
                cell.calculatedValue = result;
                
                // Update display
                const td = document.querySelector(`[data-address="${address}"]`);
                if (td) {
                    td.textContent = this.formatValue(result, cell.style?.numberFormat);
                }
            }
        });
        
        document.getElementById('calculationMode').textContent = 'Ready';
        this.updateStatistics();
    },
    
    /**
     * Auto-fit column widths
     */
    autoFitColumns() {
        const table = document.getElementById('spreadsheet');
        const cols = table.querySelectorAll('thead th:not(.row-header)');
        
        cols.forEach((th, colIndex) => {
            let maxWidth = th.offsetWidth;
            
            // Check all cells in this column
            table.querySelectorAll(`tbody td:nth-child(${colIndex + 2})`).forEach(td => {
                const content = td.textContent || '';
                const testDiv = document.createElement('div');
                testDiv.style.position = 'absolute';
                testDiv.style.visibility = 'hidden';
                testDiv.style.whiteSpace = 'nowrap';
                testDiv.style.font = window.getComputedStyle(td).font;
                testDiv.textContent = content;
                document.body.appendChild(testDiv);
                
                maxWidth = Math.max(maxWidth, testDiv.offsetWidth + 16);
                document.body.removeChild(testDiv);
            });
            
            // Set minimum and maximum widths
            maxWidth = Math.max(80, Math.min(maxWidth, 400));
            
            // Apply width to all cells in column
            table.querySelectorAll(`td:nth-child(${colIndex + 2}), th:nth-child(${colIndex + 2})`).forEach(cell => {
                cell.style.minWidth = maxWidth + 'px';
                cell.style.width = maxWidth + 'px';
            });
        });
    },
    
    /**
     * Update statistics
     */
    updateStatistics(cells) {
        const cellCount = Object.keys(this.engine.cells).length;
        const formulaCount = Object.keys(this.engine.formulas).length;
        
        document.getElementById('cellCount').textContent = `${cellCount} cells`;
        document.getElementById('formulaCount').textContent = `${formulaCount} formulas`;
        
        this.updateSelectionStats();
    },
    
    /**
     * Update selection statistics
     */
    updateSelectionStats() {
        // This would calculate SUM, AVG, COUNT for selected range
        // For now, just show current cell value
        if (this.selectedCell && this.engine.cells[this.selectedCell]) {
            const value = this.engine.getCellValue(this.selectedCell);
            if (!isNaN(value)) {
                document.getElementById('sumValue').textContent = value;
                document.getElementById('avgValue').textContent = value;
                document.getElementById('countValue').textContent = '1';
            }
        }
    },
    
    /**
     * Save all changes
     */
    async saveChanges() {
        // This would batch save all modified cells
        console.log('Saving changes...');
        document.getElementById('lastSaved').textContent = 
            'Last saved: ' + new Date().toLocaleTimeString();
    },
    
    /**
     * Export to Excel
     */
    async exportExcel() {
        if (!this.currentWorkbook) return;
        
        window.location.href = `${this.apiUrl}?action=export&workbook_id=${this.currentWorkbook}`;
    },
    
    /**
     * Refresh data from database
     */
    async refreshData() {
        if (this.currentWorksheet) {
            await this.loadWorksheet(this.currentWorksheet);
        }
    },
    
    /**
     * Show/hide loading overlay
     */
    showLoading(show) {
        document.getElementById('loadingOverlay').style.display = show ? 'flex' : 'none';
    }
};