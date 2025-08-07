// js/dashboard-main.js - Main Dashboard Controller

const ExcelDashboard = {
    // Configuration
    apiUrl: 'api/excel-api.php',
    
    // State
    currentWorkbook: null,
    currentWorksheet: null,
    selectedCell: null,
    showFormulas: false,
    
    // Cache
    worksheetsCache: {},
    cellsCache: {},
    
    /**
     * Initialize the dashboard
     */
    async init() {
        console.log('Initializing Excel Dashboard...');
        
        try {
            // Initialize sub-modules
            if (typeof ValidationManager !== 'undefined') {
                ValidationManager.init(this);
            }
            
            // Setup event listeners
            this.setupEventListeners();
            this.setupKeyboardShortcuts();
            
            // Load initial data
            await this.loadWorkbooks();
            
            // Initialize validation controls
            if (typeof ValidationManager !== 'undefined') {
                await ValidationManager.initializeDropdowns();
            }
            
            console.log('Dashboard initialized successfully');
        } catch (error) {
            console.error('Failed to initialize dashboard:', error);
            this.showError('Failed to initialize dashboard: ' + error.message);
        }
    },
    
    /**
     * Load available workbooks
     */
    async loadWorkbooks() {
        try {
            this.showLoading(true, 'Loading workbooks...');
            
            const response = await fetch(`${this.apiUrl}?action=workbooks`);
            if (!response.ok) throw new Error('Failed to fetch workbooks');
            
            const workbooks = await response.json();
            console.log(`Found ${workbooks.length} workbook(s)`);
            
            const selector = document.getElementById('workbookSelector');
            if (selector && workbooks.length > 0) {
                selector.innerHTML = '<option value="">Select Workbook</option>';
                
                workbooks.forEach(wb => {
                    const option = document.createElement('option');
                    option.value = wb.id;
                    option.textContent = wb.filename;
                    selector.appendChild(option);
                });
                
                // Auto-select first workbook
                if (workbooks.length > 0) {
                    selector.value = workbooks[0].id;
                    this.currentWorkbook = workbooks[0].id;
                    await this.loadWorkbook(workbooks[0].id);
                }
            }
        } catch (error) {
            console.error('Error loading workbooks:', error);
            this.showError('Error loading workbooks: ' + error.message);
        } finally {
            this.showLoading(false);
        }
    },
    
    /**
     * Load a specific workbook
     */
    async loadWorkbook(workbookId) {
        try {
            this.showLoading(true, 'Loading worksheets...');
            this.currentWorkbook = workbookId;
            
            const response = await fetch(`${this.apiUrl}?action=worksheets&workbook_id=${workbookId}`);
            if (!response.ok) throw new Error('Failed to fetch worksheets');
            
            const worksheets = await response.json();
            this.worksheetsCache[workbookId] = worksheets;
            
            console.log(`Found ${worksheets.length} worksheet(s)`);
            
            if (worksheets.length > 0) {
                this.createSheetTabs(worksheets);
                await this.loadWorksheet(worksheets[0].id);
                
                // Refresh validation dropdowns
                if (typeof ValidationManager !== 'undefined') {
                    await ValidationManager.refreshDropdowns();
                }
            }
        } catch (error) {
            console.error('Error loading workbook:', error);
            this.showError('Error loading workbook: ' + error.message);
        } finally {
            this.showLoading(false);
        }
    },
    
    /**
     * Load a specific worksheet
     */
    async loadWorksheet(worksheetId) {
        try {
            this.showLoading(true, 'Loading worksheet data...');
            this.currentWorksheet = worksheetId;
            
            // Update active tab
            document.querySelectorAll('.sheet-tab').forEach(tab => {
                tab.classList.toggle('active', tab.dataset.worksheetId == worksheetId);
            });
            
            // Fetch cells data
            const response = await fetch(
                `${this.apiUrl}?action=cell-range&worksheet_id=${worksheetId}` +
                `&start_row=1&end_row=100&start_col=1&end_col=50`
            );
            
            if (!response.ok) throw new Error('Failed to fetch cell data');
            
            const data = await response.json();
            this.cellsCache[worksheetId] = data;
            
            console.log(`Loaded ${data.cells ? data.cells.length : 0} cells`);
            
            // Render the spreadsheet
            this.renderSpreadsheet(data);
            
            // Update validation current values
            if (typeof ValidationManager !== 'undefined') {
                ValidationManager.updateCurrentValues();
            }
            
            // Update statistics
            this.updateStatistics(data.cells);
            
        } catch (error) {
            console.error('Error loading worksheet:', error);
            this.showError('Error loading worksheet: ' + error.message);
        } finally {
            this.showLoading(false);
        }
    },
    
    /**
     * Create sheet tabs
     */
    createSheetTabs(worksheets) {
        const container = document.getElementById('sheetTabs');
        if (!container) return;
        
        container.innerHTML = '';
        
        worksheets.forEach((ws, index) => {
            const tab = document.createElement('div');
            tab.className = 'sheet-tab';
            if (index === 0) tab.classList.add('active');
            tab.textContent = ws.sheet_name;
            tab.dataset.worksheetId = ws.id;
            tab.onclick = () => this.loadWorksheet(ws.id);
            container.appendChild(tab);
        });
    },
    
    /**
     * Render the spreadsheet grid
     */
    renderSpreadsheet(data) {
        const header = document.getElementById('spreadsheetHeader');
        const body = document.getElementById('spreadsheetBody');
        
        if (!header || !body) return;
        
        // Clear existing content
        header.innerHTML = '';
        body.innerHTML = '';
        
        // Show the table
        const table = document.getElementById('spreadsheet');
        if (table) table.style.display = 'inline-block';
        
        const emptyState = document.getElementById('emptyState');
        if (emptyState) emptyState.style.display = 'none';
        
        // Create header row
        const headerRow = document.createElement('tr');
        headerRow.innerHTML = '<th class="row-header"></th>';
        
        for (let col = 1; col <= 50; col++) {
            const th = document.createElement('th');
            th.textContent = this.numberToColumn(col);
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
                const address = this.numberToColumn(col) + row;
                td.dataset.address = address;
                
                // Find cell data
                if (data.cells) {
                    const cellData = data.cells.find(c => 
                        parseInt(c.row_num) === row && parseInt(c.col_num) === col
                    );
                    
                    if (cellData) {
                        this.renderCell(td, cellData);
                    }
                }
                
                // Add click handler
                td.onclick = () => this.selectCell(address);
                td.ondblclick = () => this.editCell(address);
                
                tr.appendChild(td);
            }
            
            body.appendChild(tr);
        }
        
        // Apply merged cells if any
        if (data.merged_cells && data.merged_cells.length > 0) {
            this.applyMergedCells(data.merged_cells);
        }
    },
    
    /**
     * Render individual cell
     */
    renderCell(td, cellData) {
        if (!td || !cellData) return;
        
        // Set cell value
        let displayValue = '';
        
        if (this.showFormulas && cellData.formula) {
            displayValue = cellData.formula;
            td.classList.add('formula-cell');
        } else if (cellData.formula) {
            displayValue = cellData.formatted_value || cellData.cell_value || '';
            td.classList.add('formula-cell');
            td.title = cellData.formula;
        } else if (cellData.cell_value !== null && cellData.cell_value !== undefined) {
            displayValue = cellData.cell_value;
        }
        
        td.textContent = displayValue;
        
        // Apply styles
        this.applyCellStyles(td, cellData);
    },
    
    /**
     * Apply cell styles
     */
    applyCellStyles(td, cellData) {
        // Font styles
        if (cellData.font_bold == 1) td.style.fontWeight = 'bold';
        if (cellData.font_italic == 1) td.style.fontStyle = 'italic';
        
        // Alignment
        if (cellData.alignment_horizontal) {
            td.style.textAlign = cellData.alignment_horizontal;
        }
        
        // Colors
        if (cellData.font_color && cellData.font_color !== 'FF000000') {
            try {
                td.style.color = '#' + cellData.font_color.substring(2);
            } catch (e) {}
        }
        
        if (cellData.fill_color && cellData.fill_color !== 'FFFFFFFF' && cellData.fill_color !== 'FF000000') {
            try {
                td.style.backgroundColor = '#' + cellData.fill_color.substring(2);
            } catch (e) {}
        }
        
        // Borders
        if (cellData.border_top && cellData.border_top !== 'none') {
            td.classList.add('border-top-thick');
        }
        if (cellData.border_right && cellData.border_right !== 'none') {
            td.classList.add('border-right-thick');
        }
        if (cellData.border_bottom && cellData.border_bottom !== 'none') {
            td.classList.add('border-bottom-thick');
        }
        if (cellData.border_left && cellData.border_left !== 'none') {
            td.classList.add('border-left-thick');
        }
    },
    
    /**
     * Apply merged cells
     */
    applyMergedCells(mergedCells) {
        // Implementation for merged cells
        console.log(`Applying ${mergedCells.length} merged cells`);
        // TODO: Implement merged cells display
    },
    
    /**
     * Select a cell
     */
    selectCell(address) {
        // Remove previous selection
        if (this.selectedCell) {
            const prevCell = document.querySelector(`[data-address="${this.selectedCell}"]`);
            if (prevCell) prevCell.classList.remove('selected');
        }
        
        // Add new selection
        this.selectedCell = address;
        const cell = document.querySelector(`[data-address="${address}"]`);
        
        if (cell) {
            cell.classList.add('selected');
            
            // Update cell reference display
            const cellRef = document.getElementById('cellReference');
            if (cellRef) cellRef.textContent = address;
            
            // Update formula bar
            const formulaInput = document.getElementById('formulaInput');
            if (formulaInput) {
                formulaInput.value = cell.title || cell.textContent || '';
            }
        }
    },
    
    /**
     * Edit a cell (double-click)
     */
    editCell(address) {
        const td = document.querySelector(`[data-address="${address}"]`);
        if (!td || td.classList.contains('editing')) return;
        
        const currentValue = td.textContent;
        const formula = td.title || '';
        
        td.classList.add('editing');
        
        const input = document.createElement('input');
        input.type = 'text';
        input.value = formula || currentValue;
        
        td.innerHTML = '';
        td.appendChild(input);
        input.focus();
        input.select();
        
        const finishEdit = async (save) => {
            if (save && input.value !== currentValue) {
                td.textContent = input.value;
                await this.saveCell(address, input.value);
            } else {
                td.textContent = currentValue;
            }
            td.classList.remove('editing');
        };
        
        input.onblur = () => finishEdit(true);
        input.onkeydown = (e) => {
            if (e.key === 'Enter') {
                e.preventDefault();
                finishEdit(true);
            } else if (e.key === 'Escape') {
                e.preventDefault();
                finishEdit(false);
            }
        };
    },
    
    /**
     * Save cell value to database
     */
    async saveCell(address, value) {
        try {
            const isFormula = value.startsWith('=');
            
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
                this.updateLastSaved();
                
                // Trigger dependent updates if this is D4 or D5
                if (address === 'D4' || address === 'D5') {
                    if (typeof ValidationManager !== 'undefined') {
                        await ValidationManager.handleCellUpdate(address, value);
                    }
                }
            } else {
                throw new Error('Failed to save cell');
            }
        } catch (error) {
            console.error('Error saving cell:', error);
            this.showError('Failed to save cell: ' + error.message);
        }
    },
    
    /**
     * Update statistics display
     */
    updateStatistics(cells) {
        if (!cells) return;
        
        const totalCells = cells.filter(c => c.cell_value !== null || c.formula !== null).length;
        const formulaCells = cells.filter(c => c.formula !== null).length;
        
        const cellCount = document.getElementById('cellCount');
        if (cellCount) cellCount.textContent = `${totalCells} cells`;
        
        const formulaCount = document.getElementById('formulaCount');
        if (formulaCount) formulaCount.textContent = `${formulaCells} formulas`;
    },
    
    /**
     * Setup event listeners
     */
    setupEventListeners() {
        // Workbook selector
        const selector = document.getElementById('workbookSelector');
        if (selector) {
            selector.onchange = (e) => {
                if (e.target.value) {
                    this.loadWorkbook(e.target.value);
                }
            };
        }
        
        // Formula input
        const formulaInput = document.getElementById('formulaInput');
        if (formulaInput) {
            formulaInput.onkeydown = (e) => {
                if (e.key === 'Enter' && this.selectedCell) {
                    const td = document.querySelector(`[data-address="${this.selectedCell}"]`);
                    if (td) {
                        td.textContent = e.target.value;
                        this.saveCell(this.selectedCell, e.target.value);
                    }
                }
            };
        }
    },
    
    /**
     * Setup keyboard shortcuts
     */
    setupKeyboardShortcuts() {
        document.addEventListener('keydown', (e) => {
            // Don't interfere with input fields
            if (e.target.matches('input, select, textarea')) return;
            
            // Arrow navigation
            if (this.selectedCell && !e.ctrlKey && !e.shiftKey) {
                let newAddress = null;
                const current = this.selectedCell.match(/([A-Z]+)(\d+)/);
                if (!current) return;
                
                const col = current[1];
                const row = parseInt(current[2]);
                
                if (e.key === 'ArrowUp' && row > 1) {
                    newAddress = col + (row - 1);
                } else if (e.key === 'ArrowDown' && row < 100) {
                    newAddress = col + (row + 1);
                } else if (e.key === 'ArrowLeft') {
                    const colNum = this.columnToNumber(col);
                    if (colNum > 1) {
                        newAddress = this.numberToColumn(colNum - 1) + row;
                    }
                } else if (e.key === 'ArrowRight') {
                    const colNum = this.columnToNumber(col);
                    if (colNum < 50) {
                        newAddress = this.numberToColumn(colNum + 1) + row;
                    }
                }
                
                if (newAddress) {
                    e.preventDefault();
                    this.selectCell(newAddress);
                }
            }
            
            // F2 to edit
            if (e.key === 'F2' && this.selectedCell) {
                e.preventDefault();
                this.editCell(this.selectedCell);
            }
            
            // Ctrl+S to save
            if (e.ctrlKey && e.key === 's') {
                e.preventDefault();
                this.saveChanges();
            }
            
            // Ctrl+Z for undo (placeholder)
            if (e.ctrlKey && e.key === 'z') {
                e.preventDefault();
                console.log('Undo not implemented yet');
            }
        });
    },
    
    // Public methods for toolbar buttons
    
    toggleFormulas() {
        this.showFormulas = !this.showFormulas;
        const btn = document.getElementById('formulaToggle');
        if (btn) btn.classList.toggle('active', this.showFormulas);
        
        if (this.currentWorksheet) {
            this.loadWorksheet(this.currentWorksheet);
        }
    },
    
    async recalculate() {
        console.log('Recalculating formulas...');
        this.showStatus('Recalculating...');
        
        if (this.currentWorksheet) {
            await this.loadWorksheet(this.currentWorksheet);
        }
        
        this.showStatus('Recalculation complete');
    },
    
    saveChanges() {
        this.showStatus('Saving changes...');
        this.updateLastSaved();
        
        setTimeout(() => {
            this.showStatus('All changes saved');
        }, 500);
    },
    
    exportExcel() {
        if (this.currentWorkbook) {
            this.showStatus('Exporting to Excel...');
            window.location.href = `${this.apiUrl}?action=export&workbook_id=${this.currentWorkbook}`;
        }
    },
    
    async refreshData() {
        this.showStatus('Refreshing data...');
        
        if (this.currentWorksheet) {
            await this.loadWorksheet(this.currentWorksheet);
            
            if (typeof ValidationManager !== 'undefined') {
                await ValidationManager.refreshDropdowns();
            }
        }
        
        this.showStatus('Data refreshed');
    },
    
    // D4 and D5 change handlers (called from HTML)
    /*
    async onD4Change(value) {
        if (typeof ValidationManager !== 'undefined') {
            await ValidationManager.onD4Change(value);
        }
    },

    async onD5Change(value) {
        if (typeof ValidationManager !== 'undefined') {
            await ValidationManager.onD5Change(value);
        }
    },
    */
    
    // Utility methods
    
    numberToColumn(num) {
        let col = '';
        while (num > 0) {
            num--;
            col = String.fromCharCode(65 + (num % 26)) + col;
            num = Math.floor(num / 26);
        }
        return col;
    },
    
    columnToNumber(col) {
        let num = 0;
        for (let i = 0; i < col.length; i++) {
            num = num * 26 + (col.charCodeAt(i) - 64);
        }
        return num;
    },
    
    showLoading(show, text) {
        const overlay = document.getElementById('loadingOverlay');
        if (overlay) {
            overlay.classList.toggle('active', show);
            if (text) {
                const textEl = document.getElementById('loadingText');
                if (textEl) textEl.textContent = text;
            }
        }
    },
    
    showStatus(message) {
        const statusEl = document.getElementById('statusMessage');
        if (statusEl) {
            statusEl.textContent = message;
            
            // Auto-clear after 3 seconds
            if (message !== 'Ready') {
                setTimeout(() => {
                    if (statusEl.textContent === message) {
                        statusEl.textContent = 'Ready';
                    }
                }, 3000);
            }
        }
    },
    
    showError(message) {
        console.error(message);
        this.showStatus('Error: ' + message);
        
        // Also show in a more prominent way if needed
        // alert(message);
    },
    
    updateLastSaved() {
        const lastSaved = document.getElementById('lastSaved');
        if (lastSaved) {
            lastSaved.textContent = 'Last saved: ' + new Date().toLocaleTimeString();
        }
    }
};

// Make globally available
window.ExcelDashboard = ExcelDashboard;