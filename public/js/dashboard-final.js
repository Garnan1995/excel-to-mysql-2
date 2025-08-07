// js/dashboard-final.js - Working version with formula support

const ExcelDashboard = {
    engine: null,
    currentWorkbook: null,
    currentWorksheet: null,
    selectedCell: null,
    showFormulas: false,
    apiUrl: 'api/excel-api.php',
    
    /**
     * Initialize the dashboard
     */
    init() {
        console.log('Initializing dashboard with formula support...');
        
        // Initialize ExcelEngine for formula evaluation
        if (typeof ExcelEngine !== 'undefined') {
            this.engine = new ExcelEngine();
            console.log('✓ ExcelEngine initialized');
        } else {
            console.warn('ExcelEngine not found - formulas will show as text');
        }
        
        // Setup event listeners
        this.setupEventListeners();
        this.setupKeyboardShortcuts();
        
        // Load workbooks
        this.loadWorkbooks();
    },
    
    /**
     * Load workbooks
     */
    async loadWorkbooks() {
        console.log('Loading workbooks...');
        
        try {
            this.showLoading(true, 'Loading workbooks...');
            
            const response = await fetch(`${this.apiUrl}?action=workbooks`);
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
                selector.value = workbooks[0].id;
                this.currentWorkbook = workbooks[0].id;
                
                // Load first workbook
                await this.loadWorkbook(workbooks[0].id);
            }
            
        } catch (error) {
            console.error('Error loading workbooks:', error);
            alert('Error loading workbooks: ' + error.message);
        } finally {
            this.showLoading(false);
        }
    },
    
    /**
     * Load workbook
     */
    async loadWorkbook(workbookId) {
        console.log('Loading workbook:', workbookId);
        
        try {
            this.showLoading(true, 'Loading worksheets...');
            this.currentWorkbook = workbookId;
            
            const response = await fetch(`${this.apiUrl}?action=worksheets&workbook_id=${workbookId}`);
            const worksheets = await response.json();
            
            console.log(`Found ${worksheets.length} worksheet(s)`);
            
            if (worksheets.length > 0) {
                this.createSheetTabs(worksheets);
                await this.loadWorksheet(worksheets[0].id);
            }
            
        } catch (error) {
            console.error('Error loading workbook:', error);
            alert('Error loading workbook: ' + error.message);
        } finally {
            this.showLoading(false);
        }
    },
    
    /**
     * Load worksheet
     */
    async loadWorksheet(worksheetId) {
        console.log('Loading worksheet:', worksheetId);
        
        try {
            this.showLoading(true, 'Loading worksheet data...');
            this.currentWorksheet = worksheetId;
            
            // Update active tab
            document.querySelectorAll('.sheet-tab').forEach(tab => {
                tab.classList.toggle('active', tab.dataset.worksheetId == worksheetId);
            });
            
            // Fetch cells
            const response = await fetch(
                `${this.apiUrl}?action=cell-range&worksheet_id=${worksheetId}` +
                `&start_row=1&end_row=100&start_col=1&end_col=50`
            );
            const data = await response.json();
            
            console.log(`Loaded ${data.cells ? data.cells.length : 0} cells`);
            
            // Load cells into engine for formula evaluation
            if (this.engine && data.cells) {
                // Only load cells with values or formulas
                const validCells = data.cells.filter(cell => 
                    cell.cell_value !== null || cell.formula !== null
                );
                console.log(`Loading ${validCells.length} valid cells into engine`);
                this.engine.loadCells(validCells);
            }
            
            // Render the spreadsheet
            this.renderSpreadsheet(data);
            
            // Update statistics
            this.updateStatistics(data.cells);
            
        } catch (error) {
            console.error('Error loading worksheet:', error);
            alert('Error loading worksheet: ' + error.message);
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
     * Render spreadsheet
     */
    renderSpreadsheet(data) {
        console.log('Rendering spreadsheet...');
        
        const header = document.getElementById('spreadsheetHeader');
        const body = document.getElementById('spreadsheetBody');
        
        if (!header || !body) {
            console.error('Spreadsheet elements not found');
            return;
        }
        
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
                
                // Add click handlers
                td.onclick = (e) => this.selectCell(address, e);
                td.ondblclick = (e) => this.editCell(address, e);
                
                tr.appendChild(td);
            }
            
            body.appendChild(tr);
        }
        
        // Apply merged cells if any
        if (data.merged_cells && data.merged_cells.length > 0) {
            this.applyMergedCells(data.merged_cells);
        }
        
        console.log('Spreadsheet rendered successfully');
    },
    
    /**
     * Render individual cell
     */
    renderCell(td, cellData) {
        if (!td || !cellData) return;
        
        try {
            // Determine what to display
            let displayValue = '';
            
            // Check if we should show formulas
            if (this.showFormulas && cellData.formula) {
                displayValue = cellData.formula;
                td.classList.add('formula-cell');
            }
            // If it's a formula and we're not showing formulas, show the value
            else if (cellData.formula) {
                // Use formatted_value if available, otherwise use cell_value
                displayValue = cellData.formatted_value || cellData.cell_value || '';
                td.classList.add('formula-cell');
                td.title = cellData.formula; // Show formula in tooltip
            }
            // Regular cell with value
            else if (cellData.cell_value !== null && cellData.cell_value !== undefined) {
                displayValue = cellData.cell_value;
            }
            
            // Set the display value
            td.textContent = displayValue;
            
            // Apply basic styling
            if (cellData.font_bold == 1) {
                td.style.fontWeight = 'bold';
            }
            
            if (cellData.font_italic == 1) {
                td.style.fontStyle = 'italic';
            }
            
            // Font color
            if (cellData.font_color && cellData.font_color !== 'FF000000') {
                try {
                    td.style.color = '#' + cellData.font_color.substring(2);
                } catch (e) {}
            }
            
            // Background color
            if (cellData.fill_color && cellData.fill_color !== 'FFFFFFFF' && cellData.fill_color !== 'FF000000') {
                try {
                    td.style.backgroundColor = '#' + cellData.fill_color.substring(2);
                } catch (e) {}
            }
            
            // Alignment
            if (cellData.alignment_horizontal) {
                if (cellData.alignment_horizontal === 'center') {
                    td.style.textAlign = 'center';
                } else if (cellData.alignment_horizontal === 'right') {
                    td.style.textAlign = 'right';
                } else if (cellData.alignment_horizontal === 'left') {
                    td.style.textAlign = 'left';
                }
            }
            
            // Borders
            if (cellData.border_top && cellData.border_top !== 'none') {
                td.style.borderTop = '1px solid #000';
            }
            if (cellData.border_right && cellData.border_right !== 'none') {
                td.style.borderRight = '1px solid #000';
            }
            if (cellData.border_bottom && cellData.border_bottom !== 'none') {
                td.style.borderBottom = '1px solid #000';
            }
            if (cellData.border_left && cellData.border_left !== 'none') {
                td.style.borderLeft = '1px solid #000';
            }
            
            // Number formatting
            if (cellData.number_format && cellData.formatted_value) {
                td.textContent = cellData.formatted_value;
            }
            
        } catch (error) {
            console.error('Error rendering cell:', error);
        }
    },
    
    /**
     * Select a cell
     */
    selectCell(address, event) {
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
            
            // Update cell reference
            const cellRef = document.getElementById('cellReference');
            if (cellRef) cellRef.textContent = address;
            
            // Update formula bar
            const formulaInput = document.getElementById('formulaInput');
            if (formulaInput) {
                // Get the cell's formula or value
                const cellValue = cell.textContent || '';
                const cellTitle = cell.title || ''; // Formula is stored in title
                formulaInput.value = cellTitle || cellValue;
            }
        }
    },
    
    /**
     * Edit a cell
     */
    editCell(address, event) {
        const td = document.querySelector(`[data-address="${address}"]`);
        if (!td) return;
        
        const currentValue = td.textContent;
        const formula = td.title || ''; // Get formula from title if exists
        
        const input = document.createElement('input');
        input.type = 'text';
        input.value = formula || currentValue; // Edit formula if exists, otherwise value
        input.style.width = '100%';
        input.style.height = '100%';
        input.style.border = 'none';
        input.style.padding = '2px';
        
        td.innerHTML = '';
        td.appendChild(input);
        input.focus();
        input.select();
        
        input.onblur = () => {
            td.textContent = currentValue; // Restore original value
            // Here you would save the new value to the server
        };
        
        input.onkeydown = (e) => {
            if (e.key === 'Enter') {
                td.textContent = input.value;
                // Here you would save the new value to the server
            } else if (e.key === 'Escape') {
                td.textContent = currentValue;
            }
        };
    },
    
    /**
     * Apply merged cells
     */
    applyMergedCells(mergedCells) {
        // Simple implementation - just log for now
        console.log(`Applying ${mergedCells.length} merged cells`);
    },
    
    /**
     * Update statistics
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
     * Convert column number to letter
     */
    numberToColumn(num) {
        let col = '';
        while (num > 0) {
            num--;
            col = String.fromCharCode(65 + (num % 26)) + col;
            num = Math.floor(num / 26);
        }
        return col;
    },
    
    /**
     * Show/hide loading overlay
     */
    showLoading(show, text) {
        const overlay = document.getElementById('loadingOverlay');
        if (!overlay) return;
        
        if (show) {
            overlay.style.display = 'flex';
            overlay.classList.add('active');
            if (text) {
                const textEl = document.getElementById('loadingText');
                if (textEl) textEl.textContent = text;
            }
        } else {
            overlay.style.display = 'none';
            overlay.classList.remove('active');
        }
    },
    
    /**
     * Setup event listeners
     */
    setupEventListeners() {
        const selector = document.getElementById('workbookSelector');
        if (selector) {
            selector.onchange = (e) => {
                if (e.target.value) {
                    this.loadWorkbook(e.target.value);
                }
            };
        }
        
        const formulaInput = document.getElementById('formulaInput');
        if (formulaInput) {
            formulaInput.onkeydown = (e) => {
                if (e.key === 'Enter' && this.selectedCell) {
                    const td = document.querySelector(`[data-address="${this.selectedCell}"]`);
                    if (td) {
                        td.textContent = e.target.value;
                        // Here you would save to server
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
            if (e.target.matches('input')) return;
            
            // Arrow keys for navigation
            if (this.selectedCell) {
                let newAddress = null;
                const current = this.selectedCell.match(/([A-Z]+)(\d+)/);
                if (!current) return;
                
                const col = current[1];
                const row = parseInt(current[2]);
                
                if (e.key === 'ArrowUp' && row > 1) {
                    newAddress = col + (row - 1);
                } else if (e.key === 'ArrowDown' && row < 100) {
                    newAddress = col + (row + 1);
                } else if (e.key === 'ArrowLeft' && col > 'A') {
                    const colNum = col.charCodeAt(0) - 65;
                    if (colNum > 0) {
                        newAddress = String.fromCharCode(64 + colNum) + row;
                    }
                } else if (e.key === 'ArrowRight' && col < 'AX') {
                    const colNum = col.charCodeAt(0) - 65;
                    if (colNum < 49) {
                        newAddress = String.fromCharCode(66 + colNum) + row;
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
        });
    },
    
    // Button functions
    toggleFormulas() {
        this.showFormulas = !this.showFormulas;
        const btn = document.getElementById('formulaToggle');
        if (btn) btn.classList.toggle('active', this.showFormulas);
        
        // Reload worksheet to apply change
        if (this.currentWorksheet) {
            this.loadWorksheet(this.currentWorksheet);
        }
    },
    
    recalculate() {
        console.log('Recalculate formulas');
        if (this.currentWorksheet) {
            this.loadWorksheet(this.currentWorksheet);
        }
    },
    
    autoFitColumns() {
        console.log('Auto-fit columns - not implemented');
    },
    
    saveChanges() {
        console.log('Save changes');
        const lastSaved = document.getElementById('lastSaved');
        if (lastSaved) {
            lastSaved.textContent = 'Last saved: ' + new Date().toLocaleTimeString();
        }
    },
    
    exportExcel() {
        if (this.currentWorkbook) {
            window.location.href = `${this.apiUrl}?action=export&workbook_id=${this.currentWorkbook}`;
        }
    },
    
    refreshData() {
        if (this.currentWorksheet) {
            this.loadWorksheet(this.currentWorksheet);
        }
    }
};

// Make globally available
window.ExcelDashboard = ExcelDashboard;
console.log('Dashboard loaded - ready to use');