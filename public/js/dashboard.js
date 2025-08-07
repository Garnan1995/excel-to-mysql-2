// js/dashboard.js - Complete final version with all fixes and data validation support

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
        console.log('Initializing dashboard with formula and validation support...');
        
        // Initialize ExcelEngine for formula evaluation
        if (typeof ExcelEngine !== 'undefined') {
            this.engine = new ExcelEngine();
            console.log('✓ ExcelEngine initialized');
        } else {
            console.warn('ExcelEngine not found - formulas will show as text');
        }
        
        // Add validation styles
        this.addValidationStyles();
        
        // Setup event listeners
        this.setupEventListeners();
        this.setupKeyboardShortcuts();
        
        // Load workbooks
        this.loadWorkbooks();
    },
    
    /**
     * Load workbooks from API
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
            console.log(`Loaded ${data.validations ? data.validations.length : 0} validations`);
            
            // Load cells into engine for formula evaluation
            if (this.engine && data.cells) {
                const validCells = data.cells.filter(cell => 
                    cell.cell_value !== null || cell.formula !== null
                );
                console.log(`Loading ${validCells.length} valid cells into engine`);
                this.engine.loadCells(validCells);
            }
            
            // Render the spreadsheet
            this.renderSpreadsheet(data);
            
            // Apply data validations AFTER rendering
            if (data.validations && data.validations.length > 0) {
                this.applyDataValidations(data.validations);
            }
            
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
     * Apply data validations to cells - FIXED VERSION
     */
    applyDataValidations(validations) {
        console.log('Applying', validations.length, 'data validations');
        
        if (!validations || validations.length === 0) return;
        
        validations.forEach(validation => {
            // Fix: Check show_dropdown properly (could be string "1" or number 1)
            const showDropdown = validation.show_dropdown == 1 || validation.show_dropdown == "1" || validation.show_dropdown === true;
            
            if (validation.validation_type === 'list' && showDropdown) {
                console.log('Processing validation for range:', validation.cell_range);
                
                // Parse the cell range to apply validation
                const cellRange = validation.cell_range;
                const cells = this.getCellsFromRange(cellRange);
                
                console.log('Applying to cells:', cells);
                
                cells.forEach(cellAddress => {
                    const td = document.querySelector(`[data-address="${cellAddress}"]`);
                    if (td) {
                        // Add validation data to the cell
                        td.dataset.validation = JSON.stringify(validation);
                        td.dataset.validationType = 'list';
                        td.classList.add('has-validation');
                        
                        // Ensure cell is positioned relative for absolute arrow
                        td.style.position = 'relative';
                        
                        // Remove any existing arrow first
                        const existingArrow = td.querySelector('.dropdown-arrow');
                        if (existingArrow) {
                            existingArrow.remove();
                        }
                        
                        // Add visual indicator (small arrow)
                        const arrow = document.createElement('span');
                        arrow.className = 'dropdown-arrow';
                        arrow.innerHTML = '▼';
                        arrow.style.cssText = `
                            position: absolute;
                            right: 2px;
                            top: 50%;
                            transform: translateY(-50%);
                            font-size: 10px;
                            color: #666;
                            pointer-events: none;
                            z-index: 2;
                        `;
                        td.appendChild(arrow);
                        
                        // Add subtle background to indicate validation
                        if (!td.style.backgroundColor || td.style.backgroundColor === '' || td.style.backgroundColor === 'white' || td.style.backgroundColor === 'rgb(255, 255, 255)') {
                            td.style.backgroundColor = '#fafafa';
                        }
                        
                        console.log('Added dropdown to cell:', cellAddress);
                    }
                });
            }
        });
    },
    
    /**
     * Get cells from a range string (e.g., "D4" or "D4:F5")
     */
    getCellsFromRange(range) {
        const cells = [];
        
        // Remove $ signs for easier parsing
        range = range.replace(/\$/g, '');
        
        // Handle single cell
        if (!range.includes(':')) {
            cells.push(range);
            return cells;
        }
        
        // Handle range (e.g., D4:F4)
        const [start, end] = range.split(':');
        const startMatch = start.match(/([A-Z]+)(\d+)/i);
        const endMatch = end.match(/([A-Z]+)(\d+)/i);
        
        if (startMatch && endMatch) {
            const startCol = this.columnToNumber(startMatch[1]);
            const endCol = this.columnToNumber(endMatch[1]);
            const startRow = parseInt(startMatch[2]);
            const endRow = parseInt(endMatch[2]);
            
            for (let row = startRow; row <= endRow; row++) {
                for (let col = startCol; col <= endCol; col++) {
                    cells.push(this.numberToColumn(col) + row);
                }
            }
        }
        
        return cells;
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
                const cellValue = cell.textContent || '';
                const cellTitle = cell.title || ''; // Formula is stored in title
                formulaInput.value = cellTitle || cellValue;
            }
        }
    },
    
    /**
     * Edit a cell with validation support
     */
    editCell(address, event) {
        const td = document.querySelector(`[data-address="${address}"]`);
        if (!td) return;
        
        const currentValue = td.textContent;
        const formula = td.title || '';
        const validation = td.dataset.validation ? JSON.parse(td.dataset.validation) : null;
        
        // Check if this cell has list validation
        if (validation && validation.validation_type === 'list') {
            this.showDropdownList(td, address, validation, currentValue);
        } else {
            // Regular edit mode
            const input = document.createElement('input');
            input.type = 'text';
            input.value = formula || currentValue;
            input.style.width = '100%';
            input.style.height = '100%';
            input.style.border = 'none';
            input.style.padding = '2px';
            
            // Clear cell content except dropdown arrow
            const arrow = td.querySelector('.dropdown-arrow');
            td.innerHTML = '';
            if (arrow) td.appendChild(arrow);
            td.appendChild(input);
            
            input.focus();
            input.select();
            
            input.onblur = () => {
                td.textContent = currentValue;
                if (arrow) td.appendChild(arrow);
            };
            
            input.onkeydown = (e) => {
                if (e.key === 'Enter') {
                    td.textContent = input.value;
                    if (arrow) td.appendChild(arrow);
                    this.saveCell(address, input.value, td);
                } else if (e.key === 'Escape') {
                    td.textContent = currentValue;
                    if (arrow) td.appendChild(arrow);
                }
            };
        }
    },
    
    /**
     * Show dropdown list for data validation
     */
    showDropdownList(td, address, validation, currentValue) {
        console.log('Showing dropdown for cell:', address);
        
        // Remove any existing dropdown
        const existingDropdown = document.getElementById('validationDropdown');
        if (existingDropdown) {
            existingDropdown.remove();
        }
        
        // Create dropdown container
        const dropdown = document.createElement('div');
        dropdown.id = 'validationDropdown';
        dropdown.style.cssText = `
            position: fixed;
            background: white;
            border: 1px solid #ccc;
            box-shadow: 0 2px 8px rgba(0,0,0,0.15);
            max-height: 250px;
            overflow-y: auto;
            z-index: 10000;
            min-width: 200px;
            border-radius: 4px;
        `;
        
        // Add loading message
        dropdown.innerHTML = '<div style="padding: 10px; color: #666;">Loading options...</div>';
        
        // Position the dropdown
        const rect = td.getBoundingClientRect();
        dropdown.style.left = rect.left + 'px';
        dropdown.style.top = (rect.bottom + 2) + 'px';
        
        // Adjust if dropdown goes off screen
        document.body.appendChild(dropdown);
        
        const dropdownRect = dropdown.getBoundingClientRect();
        if (dropdownRect.bottom > window.innerHeight) {
            dropdown.style.top = (rect.top - dropdownRect.height - 2) + 'px';
        }
        if (dropdownRect.right > window.innerWidth) {
            dropdown.style.left = (window.innerWidth - dropdownRect.width - 10) + 'px';
        }
        
        // Parse formula1 to get the list values
        const listValues = this.getValidationListValues(validation.formula1);
        
        if (listValues.length === 0) {
            // Fetch values from the referenced range
            this.fetchValidationValues(validation.formula1, (values) => {
                dropdown.innerHTML = ''; // Clear loading message
                this.populateDropdown(dropdown, values, td, address, currentValue);
            });
        } else {
            dropdown.innerHTML = ''; // Clear loading message
            this.populateDropdown(dropdown, listValues, td, address, currentValue);
        }
        
        // Close dropdown when clicking outside
        const closeDropdown = (e) => {
            if (!dropdown.contains(e.target) && e.target !== td) {
                dropdown.remove();
                document.removeEventListener('click', closeDropdown);
            }
        };
        
        // Use setTimeout to avoid immediate trigger
        setTimeout(() => {
            document.addEventListener('click', closeDropdown);
        }, 100);
    },
    
    /**
     * Populate dropdown with values
     */
    populateDropdown(dropdown, values, td, address, currentValue) {
        values.forEach(value => {
            const option = document.createElement('div');
            option.textContent = value;
            option.style.cssText = `
                padding: 5px 10px;
                cursor: pointer;
            `;
            
            if (value === currentValue) {
                option.style.background = '#e3f2fd';
            }
            
            option.onmouseover = () => {
                option.style.background = '#f0f0f0';
            };
            
            option.onmouseout = () => {
                option.style.background = value === currentValue ? '#e3f2fd' : 'white';
            };
            
            option.onclick = () => {
                td.textContent = value;
                
                // Re-add dropdown arrow
                const arrow = td.querySelector('.dropdown-arrow');
                if (!arrow) {
                    const newArrow = document.createElement('span');
                    newArrow.className = 'dropdown-arrow';
                    newArrow.innerHTML = '▼';
                    newArrow.style.cssText = `
                        position: absolute;
                        right: 2px;
                        top: 50%;
                        transform: translateY(-50%);
                        font-size: 10px;
                        color: #666;
                        pointer-events: none;
                    `;
                    td.appendChild(newArrow);
                }
                
                this.saveCell(address, value, td);
                dropdown.remove();
            };
            
            dropdown.appendChild(option);
        });
        
        if (values.length === 0) {
            const emptyMsg = document.createElement('div');
            emptyMsg.textContent = 'No options available';
            emptyMsg.style.cssText = 'padding: 5px 10px; color: #999; font-style: italic;';
            dropdown.appendChild(emptyMsg);
        }
    },
    
    /**
     * Get validation list values from formula
     */
    getValidationListValues(formula) {
        // Handle comma-separated list (e.g., "Yes,No,Maybe")
        if (formula && !formula.includes('!') && formula.includes(',')) {
            return formula.split(',').map(v => v.trim());
        }
        
        // For range references, we'll need to fetch from the server
        return [];
    },
    
    /**
     * Fetch validation values from a range reference - ENHANCED VERSION
     */
    async fetchValidationValues(formula, callback) {
        console.log('Fetching validation values for:', formula);
        
        // Parse the formula to get sheet and range
        // Handle both 'Sheet Name'!Range and Sheet!Range formats
        let sheetName, range;
        
        // Check if sheet name is quoted
        if (formula.startsWith("'")) {
            const match = formula.match(/^'([^']+)'!(.+)$/);
            if (match) {
                sheetName = match[1];
                range = match[2];
            }
        } else {
            const match = formula.match(/^([^!]+)!(.+)$/);
            if (match) {
                sheetName = match[1];
                range = match[2];
            }
        }
        
        if (!sheetName || !range) {
            console.error('Could not parse validation formula:', formula);
            callback([]);
            return;
        }
        
        console.log('Sheet:', sheetName, 'Range:', range);
        range = range.replace(/\$/g, ''); // Remove $ signs
        
        try {
            // Fetch the worksheet ID for the sheet name
            const worksheetsResponse = await fetch(
                `${this.apiUrl}?action=worksheets&workbook_id=${this.currentWorkbook}`
            );
            const worksheets = await worksheetsResponse.json();
            
            // Try to find the sheet (handle trailing spaces)
            const targetSheet = worksheets.find(ws => 
                ws.sheet_name === sheetName || 
                ws.sheet_name.trim() === sheetName.trim()
            );
            
            if (!targetSheet) {
                console.error('Sheet not found:', sheetName);
                console.log('Available sheets:', worksheets.map(ws => ws.sheet_name));
                callback([]);
                return;
            }
            
            console.log('Found sheet:', targetSheet.sheet_name, 'ID:', targetSheet.id);
            
            // Parse the range
            const [startCell, endCell] = range.split(':');
            const startMatch = startCell.match(/([A-Z]+)(\d+)/i);
            const endMatch = endCell.match(/([A-Z]+)(\d+)/i);
            
            if (!startMatch || !endMatch) {
                console.error('Invalid range:', range);
                callback([]);
                return;
            }
            
            const startCol = this.columnToNumber(startMatch[1]);
            const endCol = this.columnToNumber(endMatch[1]);
            const startRow = parseInt(startMatch[2]);
            const endRow = parseInt(endMatch[2]);
            
            console.log(`Fetching cells: rows ${startRow}-${endRow}, cols ${startCol}-${endCol}`);
            
            // Fetch cells from the range
            const cellsResponse = await fetch(
                `${this.apiUrl}?action=cell-range&worksheet_id=${targetSheet.id}` +
                `&start_row=${startRow}&end_row=${endRow}` +
                `&start_col=${startCol}&end_col=${endCol}`
            );
            const data = await cellsResponse.json();
            
            // Extract unique values from the cells
            const values = [];
            if (data.cells) {
                data.cells.forEach(cell => {
                    if (cell.cell_value !== null && cell.cell_value !== '' && cell.cell_value !== undefined) {
                        const value = String(cell.cell_value).trim();
                        if (value && !values.includes(value)) {
                            values.push(value);
                        }
                    }
                });
            }
            
            console.log(`Found ${values.length} unique values:`, values.slice(0, 5), '...');
            callback(values);
            
        } catch (error) {
            console.error('Error fetching validation values:', error);
            callback([]);
        }
    },
    
    /**
     * Save cell value
     */
    async saveCell(address, value, td) {
        const isFormula = value.startsWith('=');
        
        // Update display immediately if td provided
        if (td) {
            td.textContent = value;
        }
        
        // Update engine if available
        if (this.engine) {
            this.engine.setCellValue(address, isFormula ? null : value, isFormula ? value : null);
        }
        
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
                // Update last saved time
                const lastSaved = document.getElementById('lastSaved');
                if (lastSaved) {
                    lastSaved.textContent = 'Last saved: ' + new Date().toLocaleTimeString();
                }
            }
        } catch (error) {
            console.error('Failed to save cell:', error);
        }
    },
    
    /**
     * Apply merged cells
     */
    applyMergedCells(mergedCells) {
        console.log(`Applying ${mergedCells.length} merged cells`);
        // Implementation for merged cells would go here
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
     * Convert column letter to number
     */
    columnToNumber(col) {
        let num = 0;
        for (let i = 0; i < col.length; i++) {
            num = num * 26 + (col.charCodeAt(i) - 64);
        }
        return num;
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
                        this.saveCell(this.selectedCell, e.target.value, td);
                        td.textContent = e.target.value;
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
            
            // Ctrl+S to save
            if (e.ctrlKey && e.key === 's') {
                e.preventDefault();
                this.saveChanges();
            }
        });
    },
    
    /**
     * Add validation styles
     */
    addValidationStyles() {
        if (!document.getElementById('validationStyles')) {
            const style = document.createElement('style');
            style.id = 'validationStyles';
            style.textContent = `
                td.has-validation {
                    background-color: #fafafa !important;
                }
                
                td.has-validation:hover {
                    background-color: #f0f0f0 !important;
                }
                
                .dropdown-arrow {
                    display: inline-block;
                }
                
                td.has-validation.selected .dropdown-arrow {
                    color: #0066cc;
                }
                
                #validationDropdown {
                    border-radius: 4px;
                }
                
                #validationDropdown > div:first-child {
                    border-top-left-radius: 4px;
                    border-top-right-radius: 4px;
                }
                
                #validationDropdown > div:last-child {
                    border-bottom-left-radius: 4px;
                    border-bottom-right-radius: 4px;
                }
                
                #validationDropdown > div:hover {
                    background-color: #f0f0f0;
                }
            `;
            document.head.appendChild(style);
        }
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


// Add this fix to your dashboard.js to properly handle validation data storage

// Fixed applyDataValidations that stores validation data properly
ExcelDashboard.applyDataValidations = function(validations) {
    console.log('Applying', validations.length, 'data validations');
    
    if (!validations || validations.length === 0) return;
    
    validations.forEach(validation => {
        // Fix: Check show_dropdown properly
        const showDropdown = validation.show_dropdown == 1 || validation.show_dropdown == "1" || validation.show_dropdown === true;
        
        if (validation.validation_type === 'list' && showDropdown) {
            console.log('Processing validation for range:', validation.cell_range);
            
            const cells = this.getCellsFromRange(validation.cell_range);
            console.log('Applying to cells:', cells);
            
            cells.forEach(cellAddress => {
                const td = document.querySelector(`[data-address="${cellAddress}"]`);
                if (td) {
                    // Store validation data in a custom property instead of dataset
                    // This avoids issues with JSON parsing
                    td._validationData = validation;
                    td.classList.add('has-validation');
                    
                    // Also set a simple flag
                    td.setAttribute('data-has-validation', 'true');
                    td.setAttribute('data-validation-type', 'list');
                    
                    // Ensure cell is positioned relative
                    td.style.position = 'relative';
                    
                    // Remove any existing arrow
                    const existingArrow = td.querySelector('.dropdown-arrow');
                    if (existingArrow) {
                        existingArrow.remove();
                    }
                    
                    // Add dropdown arrow
                    const arrow = document.createElement('span');
                    arrow.className = 'dropdown-arrow';
                    arrow.innerHTML = '▼';
                    arrow.style.cssText = `
                        position: absolute;
                        right: 2px;
                        top: 50%;
                        transform: translateY(-50%);
                        font-size: 10px;
                        color: #666;
                        pointer-events: none;
                        z-index: 2;
                    `;
                    td.appendChild(arrow);
                    
                    // Add subtle background
                    if (!td.style.backgroundColor || td.style.backgroundColor === '' || 
                        td.style.backgroundColor === 'white' || td.style.backgroundColor === 'rgb(255, 255, 255)') {
                        td.style.backgroundColor = '#fafafa';
                    }
                    
                    console.log('✓ Added validation to cell:', cellAddress);
                }
            });
        }
    });
    
    // Verify D4 and D5 specifically
    setTimeout(() => {
        ['D4', 'D5'].forEach(addr => {
            const cell = document.querySelector(`[data-address="${addr}"]`);
            if (cell && cell._validationData) {
                console.log(`${addr} validation confirmed:`, cell._validationData);
            }
        });
    }, 100);
};

// Fixed editCell that uses the stored validation data
ExcelDashboard.editCell = function(address, event) {
    const td = document.querySelector(`[data-address="${address}"]`);
    if (!td) return;
    
    const currentValue = td.textContent;
    const formula = td.title || '';
    
    // Check for validation data stored as a property
    const validation = td._validationData;
    
    // Check if this cell has list validation
    if (validation && validation.validation_type === 'list') {
        console.log('Cell has validation, showing dropdown');
        this.showDropdownList(td, address, validation, currentValue);
    } else {
        // Regular edit mode
        const input = document.createElement('input');
        input.type = 'text';
        input.value = formula || currentValue;
        input.style.width = '100%';
        input.style.height = '100%';
        input.style.border = 'none';
        input.style.padding = '2px';
        
        // Clear cell content except dropdown arrow
        const arrow = td.querySelector('.dropdown-arrow');
        td.innerHTML = '';
        if (arrow) td.appendChild(arrow);
        td.appendChild(input);
        
        input.focus();
        input.select();
        
        input.onblur = () => {
            td.textContent = currentValue;
            if (arrow) td.appendChild(arrow);
        };
        
        input.onkeydown = (e) => {
            if (e.key === 'Enter') {
                td.textContent = input.value;
                if (arrow) td.appendChild(arrow);
                this.saveCell(address, input.value, td);
            } else if (e.key === 'Escape') {
                td.textContent = currentValue;
                if (arrow) td.appendChild(arrow);
            }
        };
    }
};

// Add a manual trigger function for testing
ExcelDashboard.testDropdown = function(cellAddress) {
    const td = document.querySelector(`[data-address="${cellAddress}"]`);
    if (!td) {
        console.error('Cell not found:', cellAddress);
        return;
    }
    
    const validation = td._validationData;
    if (!validation) {
        console.error('No validation data for cell:', cellAddress);
        return;
    }
    
    console.log('Triggering dropdown for:', cellAddress);
    this.showDropdownList(td, cellAddress, validation, td.textContent);
};

console.log('Validation fix applied - try double-clicking D4 or D5');