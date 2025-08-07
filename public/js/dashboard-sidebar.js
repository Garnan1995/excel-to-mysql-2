// js/dashboard-sidebar.js - Enhanced dashboard with validation sidebar and dependency tracking

const ExcelDashboard = {
    engine: null,
    currentWorkbook: null,
    currentWorksheet: null,
    selectedCell: null,
    showFormulas: false,
    apiUrl: 'api/excel-api.php',
    validationMap: {}, // Maps cell addresses to validation data
    validationGroups: {}, // Groups validations by type/category
    dependencyMap: {}, // Maps cells to their dependencies
    reverseDependencyMap: {}, // Maps cells to cells that depend on them
    formulaCells: [], // Array of cells containing formulas
    
    /**
     * Initialize the dashboard
     */
    init() {
        console.log('Initializing dashboard with validation sidebar...');
        
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
     * Initialize validation sidebar
     */
    initializeValidationSidebar() {
        console.log('Initializing validation sidebar...');
        
        try {
            // Group validations by category
            this.groupValidations();
            
            // Populate sidebar
            this.populateValidationSidebar();
            
            // Setup sidebar event handlers
            this.setupSidebarEventHandlers();
        } catch (error) {
            console.error('Error initializing sidebar:', error);
            // Continue anyway - sidebar is not critical
        }
    },
    
    /**
     * Group validations by category/type
     */
    groupValidations() {
        // Only show D4 and D5 validation cells
        this.validationGroups = {
            'Primary Controls': {
                cells: ['D4', 'D5'],
                description: 'Main selection fields',
                icon: '📋'
            }
        };
    },
    
    /**
     * Populate validation sidebar with controls
     */
    populateValidationSidebar() {
        const sidebarContent = document.getElementById('sidebarContent');
        if (!sidebarContent) return;
        
        sidebarContent.innerHTML = '';
        
        // Only create controls for D4 and D5
        const allowedCells = ['D4', 'D5'];
        
        // Create group for the allowed validation cells
        Object.entries(this.validationGroups).forEach(([groupName, groupData]) => {
            const groupDiv = document.createElement('div');
            groupDiv.className = 'validation-group';
            groupDiv.innerHTML = `
                <div class="validation-group-header">
                    ${groupData.icon || ''} ${groupName}
                </div>
            `;
            
            // Add validation controls only for D4 and D5
            const validationItems = document.createElement('div');
            
            groupData.cells.forEach(cellAddress => {
                if (allowedCells.includes(cellAddress)) {
                    const validationData = this.validationMap[cellAddress] || this.getKnownValidationData(cellAddress);
                    if (validationData || this.isKnownValidationCell(cellAddress)) {
                        const itemDiv = this.createValidationControl(cellAddress, validationData);
                        validationItems.appendChild(itemDiv);
                    }
                }
            });
            
            if (validationItems.children.length > 0) {
                groupDiv.appendChild(validationItems);
                sidebarContent.appendChild(groupDiv);
            }
        });
        
        // If no validations found, show a message
        if (sidebarContent.children.length === 0) {
            sidebarContent.innerHTML = `
                <div class="validation-group">
                    <div class="validation-group-header">Loading Validations</div>
                    <p style="color: #666; font-size: 13px; padding: 10px;">
                        Setting up validation controls for cells D4 and D5...
                    </p>
                </div>
            `;
            
            // Try to set up the known validation cells
            setTimeout(() => {
                this.setupKnownValidations();
            }, 500);
        }
    },
    
    /**
     * Setup known validations for D4 and D5
     */
    setupKnownValidations() {
        const sidebarContent = document.getElementById('sidebarContent');
        if (!sidebarContent) return;
        
        sidebarContent.innerHTML = '';
        
        const groupDiv = document.createElement('div');
        groupDiv.className = 'validation-group';
        groupDiv.innerHTML = `
            <div class="validation-group-header">
                📋 Primary Controls
            </div>
        `;
        
        const validationItems = document.createElement('div');
        
        // Create controls for D4 and D5
        ['D4', 'D5'].forEach(cellAddress => {
            const validationData = this.getKnownValidationData(cellAddress);
            const itemDiv = this.createValidationControl(cellAddress, validationData);
            validationItems.appendChild(itemDiv);
        });
        
        groupDiv.appendChild(validationItems);
        sidebarContent.appendChild(groupDiv);
    },
    
    /**
     * Create validation control for a cell
     */
    createValidationControl(cellAddress, validationData) {
        const itemDiv = document.createElement('div');
        itemDiv.className = 'validation-item';
        itemDiv.dataset.cellAddress = cellAddress;
        
        // Get current cell value
        const cell = document.querySelector(`[data-address="${cellAddress}"]`);
        const currentValue = cell ? cell.textContent : '';
        
        // Create label with cell reference
        const labelDiv = document.createElement('div');
        labelDiv.className = 'validation-label';
        labelDiv.innerHTML = `
            <span class="validation-label-text">${this.getValidationLabel(cellAddress)}</span>
            <span class="validation-cell-ref" onclick="ExcelDashboard.focusCell('${cellAddress}')">${cellAddress}</span>
        `;
        
        // Check if this should be a searchable dropdown
        const isLargeList = validationData && validationData.formula1 && 
                           this.isLargeValidationList(validationData.formula1);
        
        if (isLargeList) {
            // Create searchable dropdown
            const searchInput = document.createElement('input');
            searchInput.type = 'text';
            searchInput.className = 'validation-search';
            searchInput.placeholder = 'Type to search...';
            searchInput.value = currentValue;
            searchInput.dataset.cellAddress = cellAddress;
            
            const selectDiv = document.createElement('div');
            selectDiv.className = 'validation-options-list';
            selectDiv.style.display = 'none';
            
            itemDiv.appendChild(labelDiv);
            itemDiv.appendChild(searchInput);
            itemDiv.appendChild(selectDiv);
            
            // Setup search functionality
            this.setupSearchableDropdown(searchInput, selectDiv, cellAddress, validationData);
        } else {
            // Create standard dropdown
            const select = document.createElement('select');
            select.className = 'validation-select';
            select.dataset.cellAddress = cellAddress;
            
            // Add empty option
            select.innerHTML = '<option value="">-- Select --</option>';
            
            // Set current value if exists
            if (currentValue) {
                const option = document.createElement('option');
                option.value = currentValue;
                option.textContent = currentValue;
                option.selected = true;
                select.appendChild(option);
            }
            
            itemDiv.appendChild(labelDiv);
            itemDiv.appendChild(select);
            
            // Load options asynchronously
            this.loadValidationOptions(select, cellAddress, validationData);
        }
        
        return itemDiv;
    },
    
    /**
     * Get user-friendly label for validation cell
     */
    getValidationLabel(cellAddress) {
        // Map cell addresses to user-friendly labels - Only D4 and D5
        const labelMap = {
            'D4': 'Aircraft Type',
            'D5': 'Aircraft Registration'
        };
        
        return labelMap[cellAddress] || `Cell ${cellAddress}`;
    },
    
    /**
     * Check if validation list is large (needs search)
     */
    isLargeValidationList(formula) {
        // Check if the range is large (e.g., more than 20 items)
        if (formula && formula.includes(':')) {
            const match = formula.match(/\$?([A-Z]+)\$?(\d+):\$?([A-Z]+)\$?(\d+)/);
            if (match) {
                const startRow = parseInt(match[2]);
                const endRow = parseInt(match[4]);
                return (endRow - startRow) > 20;
            }
        }
        return false;
    },
    
    /**
     * Setup searchable dropdown
     */
    setupSearchableDropdown(searchInput, selectDiv, cellAddress, validationData) {
        let allOptions = [];
        let filteredOptions = [];
        
        // Load options
        this.fetchValidationValues(validationData.formula1, (options) => {
            allOptions = options;
            filteredOptions = options;
            
            searchInput.addEventListener('focus', () => {
                this.showSearchResults(selectDiv, filteredOptions, searchInput, cellAddress);
            });
            
            searchInput.addEventListener('input', (e) => {
                const searchTerm = e.target.value.toLowerCase();
                filteredOptions = allOptions.filter(opt => 
                    opt.toLowerCase().includes(searchTerm)
                );
                this.showSearchResults(selectDiv, filteredOptions, searchInput, cellAddress);
            });
            
            // Hide on click outside
            document.addEventListener('click', (e) => {
                if (!searchInput.contains(e.target) && !selectDiv.contains(e.target)) {
                    selectDiv.style.display = 'none';
                }
            });
        });
    },
    
    /**
     * Show search results dropdown
     */
    showSearchResults(selectDiv, options, searchInput, cellAddress) {
        selectDiv.innerHTML = '';
        selectDiv.style.cssText = `
            display: block;
            position: absolute;
            background: white;
            border: 1px solid #ddd;
            border-radius: 5px;
            max-height: 200px;
            overflow-y: auto;
            width: ${searchInput.offsetWidth}px;
            margin-top: 2px;
            box-shadow: 0 2px 8px rgba(0,0,0,0.1);
            z-index: 1000;
        `;
        
        if (options.length === 0) {
            selectDiv.innerHTML = '<div style="padding: 8px; color: #999;">No matches found</div>';
            return;
        }
        
        options.slice(0, 50).forEach(option => {
            const optionDiv = document.createElement('div');
            optionDiv.style.cssText = 'padding: 8px; cursor: pointer;';
            optionDiv.textContent = option;
            optionDiv.onmouseover = () => {
                optionDiv.style.background = '#f0f0f0';
            };
            optionDiv.onmouseout = () => {
                optionDiv.style.background = 'white';
            };
            optionDiv.onclick = () => {
                searchInput.value = option;
                selectDiv.style.display = 'none';
                this.updateCellValue(cellAddress, option);
            };
            selectDiv.appendChild(optionDiv);
        });
    },
    
    /**
     * Load validation options for a dropdown
     */
    async loadValidationOptions(select, cellAddress, validationData) {
        // Set loading state
        select.innerHTML = '<option value="">Loading...</option>';
        
        if (!validationData || !validationData.formula1) {
            // Try to get validation from known cells
            validationData = this.getKnownValidationData(cellAddress);
        }
        
        if (validationData && validationData.formula1) {
            this.fetchValidationValues(validationData.formula1, (options) => {
                select.innerHTML = '<option value="">-- Select --</option>';
                
                const currentValue = select.dataset.currentValue || '';
                
                options.forEach(option => {
                    const optionEl = document.createElement('option');
                    optionEl.value = option;
                    optionEl.textContent = option;
                    if (option === currentValue) {
                        optionEl.selected = true;
                    }
                    select.appendChild(optionEl);
                });
                
                // Setup change handler
                select.onchange = () => {
                    this.updateCellValue(cellAddress, select.value);
                };
            });
        }
    },
    
    /**
     * Get known validation data for specific cells
     */
    getKnownValidationData(cellAddress) {
        // Only return validation data for D4 and D5
        const knownValidations = {
            'D4': { formula1: "'Standar Jumlah'!$B$5:$B$83", validation_type: 'list' },
            'D5': { formula1: "'List Aircraft'!$R$3:$R$42", validation_type: 'list' }
        };
        
        return knownValidations[cellAddress];
    },
    
    /**
     * Check if cell is known to have validation
     */
    isKnownValidationCell(cellAddress) {
        // Only D4 and D5 are allowed
        const knownCells = ['D4', 'D5'];
        return knownCells.includes(cellAddress);
    },
    
    /**
     * Check if cell is in any group
     */
    isCellInAnyGroup(cellAddress) {
        for (const groupData of Object.values(this.validationGroups)) {
            if (groupData.cells.includes(cellAddress)) {
                return true;
            }
        }
        return false;
    },
    
    /**
     * Update cell value from sidebar with cascade updates
     */
    updateCellValue(cellAddress, value) {
        const cell = document.querySelector(`[data-address="${cellAddress}"]`);
        if (cell) {
            // Update visual
            cell.textContent = value;
            cell.classList.add('validation-linked');
            
            // Flash to indicate update
            cell.style.transition = 'background-color 0.3s';
            const originalBg = cell.style.backgroundColor;
            cell.style.backgroundColor = '#90EE90';
            setTimeout(() => {
                cell.style.backgroundColor = originalBg || '';
                cell.classList.add('validation-linked');
            }, 300);
            
            // Save to database
            this.saveCell(cellAddress, value, cell);
            
            // Update engine with new value
            if (this.engine) {
                this.engine.setCellValue(cellAddress, value, null);
            }
            
            // Trigger cascade updates for dependent cells
            this.processDependentCells(cellAddress, value);
            
            // Update sync indicator
            this.showSyncStatus('syncing');
            setTimeout(() => this.showSyncStatus('synced'), 1000);
        }
    },
    
    /**
     * Process dependent cells that rely on the changed cell
     */
    processDependentCells(changedCell, newValue) {
        console.log(`Processing dependencies for ${changedCell} with value: ${newValue}`);
        
        // Track cells that need updating
        const cellsToUpdate = [];
        
        // Check all cells for formulas that reference the changed cell
        document.querySelectorAll('td[data-address]').forEach(td => {
            const address = td.dataset.address;
            const formula = td.title; // Formulas are stored in title attribute
            
            if (formula && formula.includes(changedCell)) {
                cellsToUpdate.push({
                    address: address,
                    element: td,
                    formula: formula
                });
            }
        });
        
        // Special handling for D4 and D5 selections
        if (changedCell === 'D4') {
            this.processD4Dependencies(newValue, cellsToUpdate);
        } else if (changedCell === 'D5') {
            this.processD5Dependencies(newValue, cellsToUpdate);
        }
        
        // Process VLOOKUP and other formula dependencies
        this.processFormulaDependencies(changedCell, newValue, cellsToUpdate);
        
        // Update all dependent cells
        cellsToUpdate.forEach(cellInfo => {
            this.evaluateAndUpdateCell(cellInfo);
        });
    },
    
    /**
     * Process dependencies specific to D4 (Aircraft Type)
     */
    processD4Dependencies(aircraftType, cellsToUpdate) {
        console.log('Processing D4 (Aircraft Type) dependencies:', aircraftType);
        
        // Common patterns for aircraft type dependencies
        const dependentRanges = [
            { range: 'E4:F4', description: 'Related aircraft fields' },
            { range: 'D6:F10', description: 'Aircraft specifications' },
            { range: 'D11:F20', description: 'Configuration details' }
        ];
        
        dependentRanges.forEach(({ range }) => {
            const cells = this.getCellsFromRange(range);
            cells.forEach(cellAddress => {
                const td = document.querySelector(`[data-address="${cellAddress}"]`);
                if (td && td.title) { // Has formula
                    // Check for VLOOKUP formulas that might use D4
                    if (td.title.includes('VLOOKUP') || td.title.includes('MATCH') || td.title.includes('INDEX')) {
                        cellsToUpdate.push({
                            address: cellAddress,
                            element: td,
                            formula: td.title,
                            dependsOn: 'D4'
                        });
                    }
                }
            });
        });
    },
    
    /**
     * Process dependencies specific to D5 (Registration)
     */
    processD5Dependencies(registration, cellsToUpdate) {
        console.log('Processing D5 (Registration) dependencies:', registration);
        
        // Common patterns for registration dependencies
        const dependentRanges = [
            { range: 'E5:F5', description: 'Related registration fields' },
            { range: 'D21:F30', description: 'Registration details' }
        ];
        
        dependentRanges.forEach(({ range }) => {
            const cells = this.getCellsFromRange(range);
            cells.forEach(cellAddress => {
                const td = document.querySelector(`[data-address="${cellAddress}"]`);
                if (td && td.title) { // Has formula
                    if (td.title.includes('VLOOKUP') || td.title.includes('MATCH') || td.title.includes('INDEX')) {
                        cellsToUpdate.push({
                            address: cellAddress,
                            element: td,
                            formula: td.title,
                            dependsOn: 'D5'
                        });
                    }
                }
            });
        });
    },
    
    /**
     * Process formula dependencies using the Excel engine
     */
    async processFormulaDependencies(changedCell, newValue, cellsToUpdate) {
        if (!this.engine) return;
        
        // Get all formulas from the current worksheet
        const formulas = await this.fetchWorksheetFormulas();
        
        formulas.forEach(formulaInfo => {
            // Check if formula references the changed cell
            if (this.formulaReferencesCell(formulaInfo.formula, changedCell)) {
                const td = document.querySelector(`[data-address="${formulaInfo.address}"]`);
                if (td && !cellsToUpdate.find(c => c.address === formulaInfo.address)) {
                    cellsToUpdate.push({
                        address: formulaInfo.address,
                        element: td,
                        formula: formulaInfo.formula,
                        originalValue: formulaInfo.value
                    });
                }
            }
        });
    },
    
    /**
     * Check if a formula references a specific cell
     */
    formulaReferencesCell(formula, cellAddress) {
        if (!formula) return false;
        
        // Check for direct cell references
        const cellPattern = new RegExp(`\\b${cellAddress}\\b`, 'i');
        if (cellPattern.test(formula)) return true;
        
        // Check for range references that include the cell
        const rangePattern = /([A-Z]+\d+):([A-Z]+\d+)/gi;
        let match;
        while ((match = rangePattern.exec(formula)) !== null) {
            const range = match[0];
            const cells = this.getCellsFromRange(range);
            if (cells.includes(cellAddress)) return true;
        }
        
        return false;
    },
    
    /**
     * Evaluate and update a cell with formula
     */
    async evaluateAndUpdateCell(cellInfo) {
        console.log(`Evaluating cell ${cellInfo.address} with formula: ${cellInfo.formula}`);
        
        try {
            let newValue = '';
            
            // Use the Excel engine to evaluate the formula
            if (this.engine && cellInfo.formula) {
                newValue = await this.evaluateFormula(cellInfo.formula, cellInfo.address);
            }
            
            // Update the cell display
            if (cellInfo.element) {
                const oldValue = cellInfo.element.textContent;
                if (oldValue !== newValue) {
                    cellInfo.element.textContent = newValue;
                    
                    // Highlight the updated cell
                    cellInfo.element.style.transition = 'background-color 0.5s';
                    const originalBg = cellInfo.element.style.backgroundColor;
                    cellInfo.element.style.backgroundColor = '#FFE4B5'; // Light orange for dependent updates
                    setTimeout(() => {
                        cellInfo.element.style.backgroundColor = originalBg || '';
                    }, 1000);
                    
                    // Save the updated value to database
                    this.saveDependentCell(cellInfo.address, newValue, cellInfo.formula);
                }
            }
        } catch (error) {
            console.error(`Error evaluating cell ${cellInfo.address}:`, error);
        }
    },
    
    /**
     * Evaluate formula using Excel engine or API
     */
    async evaluateFormula(formula, cellAddress) {
        // First try using the local engine
        if (this.engine && this.engine.evaluateFormula) {
            try {
                return this.engine.evaluateFormula(formula, cellAddress);
            } catch (error) {
                console.debug('Local evaluation failed, trying server:', error.message);
            }
        }
        
        // Fallback to server-side evaluation (only if endpoint exists)
        try {
            const response = await fetch(`${this.apiUrl}?action=evaluate-formula`, {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    formula: formula,
                    worksheet_id: this.currentWorksheet,
                    cell_address: cellAddress
                })
            });
            
            if (response.ok) {
                const data = await response.json();
                return data.result || '';
            }
        } catch (error) {
            console.debug('Server evaluation not available');
        }
        
        // Return the original formula if evaluation fails
        return formula;
    },
    
    /**
     * Save dependent cell update to database
     */
    async saveDependentCell(address, value, formula) {
        try {
            const response = await fetch(`${this.apiUrl}?action=update-cell`, {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    worksheet_id: this.currentWorksheet,
                    cell_address: address,
                    value: value,
                    formula: formula,
                    formatted_value: value,
                    is_dependent_update: true
                })
            });
            
            if (response.ok) {
                console.log(`Saved dependent cell ${address} with value: ${value}`);
            }
        } catch (error) {
            console.error(`Failed to save dependent cell ${address}:`, error);
        }
    },
    
    /**
     * Fetch all formulas in the worksheet
     */
    async fetchWorksheetFormulas() {
        try {
            // Check if endpoint exists first
            const response = await fetch(
                `${this.apiUrl}?action=get-formulas&worksheet_id=${this.currentWorksheet}`,
                { method: 'HEAD' }
            ).catch(() => null);
            
            if (!response || !response.ok) {
                // Endpoint doesn't exist, return formulas from DOM
                const formulas = [];
                document.querySelectorAll('td[title]').forEach(td => {
                    if (td.title && td.title.startsWith('=')) {
                        formulas.push({
                            address: td.dataset.address,
                            formula: td.title,
                            value: td.textContent
                        });
                    }
                });
                return formulas;
            }
            
            // Endpoint exists, fetch data
            const dataResponse = await fetch(
                `${this.apiUrl}?action=get-formulas&worksheet_id=${this.currentWorksheet}`
            );
            
            if (dataResponse.ok) {
                const data = await dataResponse.json();
                return data.formulas || [];
            }
        } catch (error) {
            console.debug('Formula fetch not available, using DOM data');
        }
        
        // Fallback to DOM data
        const formulas = [];
        document.querySelectorAll('td[title]').forEach(td => {
            if (td.title && td.title.startsWith('=')) {
                formulas.push({
                    address: td.dataset.address,
                    formula: td.title,
                    value: td.textContent
                });
            }
        });
        return formulas;
    },
    
    /**
     * Recalculate all formulas in the worksheet
     */
    async recalculateAll() {
        console.log('Recalculating all formulas...');
        this.showSyncStatus('syncing');
        
        // Get all cells with formulas
        const formulaCells = [];
        document.querySelectorAll('td.formula-cell').forEach(td => {
            if (td.title) { // Has formula
                formulaCells.push({
                    address: td.dataset.address,
                    element: td,
                    formula: td.title
                });
            }
        });
        
        // Sort by dependencies (simple cells first, complex formulas last)
        formulaCells.sort((a, b) => {
            const aComplexity = (a.formula.match(/VLOOKUP|INDEX|MATCH|SUMIF/gi) || []).length;
            const bComplexity = (b.formula.match(/VLOOKUP|INDEX|MATCH|SUMIF/gi) || []).length;
            return aComplexity - bComplexity;
        });
        
        // Recalculate each formula
        for (const cellInfo of formulaCells) {
            await this.evaluateAndUpdateCell(cellInfo);
        }
        
        this.showSyncStatus('synced');
        console.log('Recalculation complete');
    },
    
    /**
     * Focus on a specific cell
     */
    focusCell(cellAddress) {
        const cell = document.querySelector(`[data-address="${cellAddress}"]`);
        if (cell) {
            // Scroll to cell
            cell.scrollIntoView({ behavior: 'smooth', block: 'center', inline: 'center' });
            
            // Select the cell
            this.selectCell(cellAddress);
            
            // Highlight briefly
            cell.style.transition = 'box-shadow 0.3s';
            cell.style.boxShadow = '0 0 10px rgba(102, 126, 234, 0.5)';
            setTimeout(() => {
                cell.style.boxShadow = '';
            }, 1000);
        }
    },
    
    /**
     * Show sync status
     */
    showSyncStatus(status) {
        const indicator = document.getElementById('syncIndicator');
        if (indicator) {
            if (status === 'syncing') {
                indicator.classList.add('syncing');
                indicator.querySelector('span:last-child').textContent = 'Syncing...';
            } else {
                indicator.classList.remove('syncing');
                indicator.querySelector('span:last-child').textContent = 'Synced';
            }
        }
    },
    
    /**
     * Setup sidebar event handlers
     */
    setupSidebarEventHandlers() {
        console.log('Setting up sidebar event handlers');
        
        // Remove any existing handlers first
        const oldHandlers = document.querySelectorAll('.validation-select');
        oldHandlers.forEach(select => {
            const newSelect = select.cloneNode(true);
            select.parentNode.replaceChild(newSelect, select);
        });
        
        // Add event delegation for validation selects
        document.addEventListener('change', (e) => {
            if (e.target && e.target.classList.contains('validation-select')) {
                const cellAddress = e.target.dataset.cellAddress;
                const value = e.target.value;
                console.log(`Validation change detected - Cell: ${cellAddress}, Value: ${value}`);
                if (cellAddress && value) {
                    this.updateCellValue(cellAddress, value);
                }
            }
        });
        
        // Add event delegation for cell references
        document.addEventListener('click', (e) => {
            if (e.target && e.target.classList.contains('validation-cell-ref')) {
                const cellAddress = e.target.textContent;
                console.log(`Cell reference clicked: ${cellAddress}`);
                this.focusCell(cellAddress);
            }
        });
        
        console.log('Sidebar event handlers setup complete');
    },
    
    /**
     * Apply bulk validations
     */
    applyBulkValidations(validations) {
        console.log('Applying bulk validations:', validations);
        
        this.showSyncStatus('syncing');
        
        Object.entries(validations).forEach(([cellAddress, value]) => {
            this.updateCellValue(cellAddress, value);
        });
        
        setTimeout(() => {
            this.showSyncStatus('synced');
            alert('All validations applied successfully!');
        }, 500);
    },
    
    /**
     * Fetch validation values from a range reference
     */
    async fetchValidationValues(formula, callback) {
        console.log('Fetching validation values for:', formula);
        
        // Parse the formula to get sheet and range
        let sheetName, range;
        
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
        range = range.replace(/\$/g, '');
        
        try {
            // Fetch the worksheet ID for the sheet name
            const worksheetsResponse = await fetch(
                `${this.apiUrl}?action=worksheets&workbook_id=${this.currentWorkbook}`
            );
            const worksheets = await worksheetsResponse.json();
            
            const targetSheet = worksheets.find(ws => 
                ws.sheet_name === sheetName || 
                ws.sheet_name.trim() === sheetName.trim()
            );
            
            if (!targetSheet) {
                console.error('Sheet not found:', sheetName);
                callback([]);
                return;
            }
            
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
            
            // Fetch cells from the range
            const cellsResponse = await fetch(
                `${this.apiUrl}?action=cell-range&worksheet_id=${targetSheet.id}` +
                `&start_row=${startRow}&end_row=${endRow}` +
                `&start_col=${startCol}&end_col=${endCol}`
            );
            const data = await cellsResponse.json();
            
            // Extract unique values
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
            
            console.log(`Found ${values.length} unique values`);
            callback(values.sort());
            
        } catch (error) {
            console.error('Error fetching validation values:', error);
            callback([]);
        }
    },
    
    // ... (Include all other existing methods from the original dashboard.js)
    // The rest of the methods remain the same as in the original implementation
    // Including: loadWorkbooks, loadWorkbook, loadWorksheet, renderSpreadsheet, etc.
    
    /**
     * Apply data validations - Modified to only handle D4 and D5
     */
    applyDataValidations(validations) {
        console.log('Applying data validations (filtering for D4 and D5 only)');
        
        if (!validations || validations.length === 0) {
            // Even if no validations from database, set up D4 and D5
            this.setupKnownValidations();
            return;
        }
        
        // Clear existing validation map
        this.validationMap = {};
        
        // Only process D4 and D5
        const allowedCells = ['D4', 'D5'];
        
        validations.forEach(validation => {
            const showDropdown = validation.show_dropdown == 1 || validation.show_dropdown == "1" || validation.show_dropdown === true;
            
            if (validation.validation_type === 'list' && showDropdown) {
                const cells = this.getCellsFromRange(validation.cell_range);
                
                cells.forEach(cellAddress => {
                    // Only store validation for D4 and D5
                    if (allowedCells.includes(cellAddress)) {
                        // Store validation data
                        this.validationMap[cellAddress] = validation;
                        
                        // Mark cells in spreadsheet
                        const td = document.querySelector(`[data-address="${cellAddress}"]`);
                        if (td) {
                            td.classList.add('validation-linked');
                            td.title = 'This cell has data validation - see sidebar';
                        }
                    }
                });
            }
        });
        
        // Always initialize sidebar with D4 and D5
        this.initializeValidationSidebar();
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
     * Load worksheet with enhanced formula processing
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
            
            if (!response.ok) {
                throw new Error(`HTTP error! status: ${response.status}`);
            }
            
            const data = await response.json();
            
            console.log(`Loaded ${data.cells ? data.cells.length : 0} cells`);
            console.log(`Loaded ${data.validations ? data.validations.length : 0} validations`);
            
            // Build dependency map for formulas (safe)
            if (data.cells) {
                this.buildDependencyMap(data.cells);
            }
            
            // Load cells into engine for formula evaluation
            if (this.engine && data.cells) {
                const validCells = data.cells.filter(cell => 
                    cell.cell_value !== null || cell.formula !== null
                );
                console.log(`Loading ${validCells.length} valid cells into engine`);
                this.engine.loadCells(validCells);
                
                // Track formula cells for dependency resolution
                this.formulaCells = validCells.filter(cell => cell.formula !== null);
            }
            
            // Render the spreadsheet
            this.renderSpreadsheet(data);
            
            // Apply data validations AFTER rendering
            if (data.validations && data.validations.length > 0) {
                this.applyDataValidations(data.validations);
            } else {
                // Set up known validations for D4 and D5
                setTimeout(() => {
                    this.setupKnownValidations();
                }, 100);
            }
            
            // Process formulas after a short delay to ensure DOM is ready
            setTimeout(() => {
                this.processAllFormulas().catch(err => {
                    console.warn('Formula processing skipped:', err);
                });
            }, 500);
            
            // Update statistics
            this.updateStatistics(data.cells);
            
        } catch (error) {
            console.error('Error loading worksheet:', error);
            // Show the spreadsheet even if there's an error
            const table = document.getElementById('spreadsheet');
            if (table) table.style.display = 'inline-block';
            
            const emptyState = document.getElementById('emptyState');
            if (emptyState) emptyState.style.display = 'none';
        } finally {
            this.showLoading(false);
        }
    },
    
    /**
     * Build dependency map for formula cells
     */
    buildDependencyMap(cells) {
        this.dependencyMap = {};
        this.reverseDependencyMap = {};
        
        if (!cells) return;
        
        cells.forEach(cell => {
            if (cell.formula) {
                const dependencies = this.extractDependencies(cell.formula);
                this.dependencyMap[cell.cell_address] = dependencies;
                
                // Build reverse map (which cells depend on this cell)
                dependencies.forEach(dep => {
                    if (!this.reverseDependencyMap[dep]) {
                        this.reverseDependencyMap[dep] = [];
                    }
                    this.reverseDependencyMap[dep].push(cell.cell_address);
                });
            }
        });
        
        console.log('Dependency map built:', {
            dependencies: Object.keys(this.dependencyMap).length,
            reverseDependencies: Object.keys(this.reverseDependencyMap).length
        });
    },
    
    /**
     * Extract cell dependencies from a formula
     */
    extractDependencies(formula) {
        const dependencies = [];
        if (!formula) return dependencies;
        
        // Match individual cell references (e.g., D4, $D$4)
        const cellPattern = /\$?([A-Z]+)\$?(\d+)/gi;
        let match;
        while ((match = cellPattern.exec(formula)) !== null) {
            const cellAddr = match[1] + match[2];
            if (!dependencies.includes(cellAddr)) {
                dependencies.push(cellAddr);
            }
        }
        
        // Match range references and expand them
        const rangePattern = /\$?([A-Z]+)\$?(\d+):\$?([A-Z]+)\$?(\d+)/gi;
        while ((match = rangePattern.exec(formula)) !== null) {
            const startCol = this.columnToNumber(match[1]);
            const startRow = parseInt(match[2]);
            const endCol = this.columnToNumber(match[3]);
            const endRow = parseInt(match[4]);
            
            for (let row = startRow; row <= endRow; row++) {
                for (let col = startCol; col <= endCol; col++) {
                    const addr = this.numberToColumn(col) + row;
                    if (!dependencies.includes(addr)) {
                        dependencies.push(addr);
                    }
                }
            }
        }
        
        return dependencies;
    },
    
    /**
     * Process all formulas in dependency order
     */
    async processAllFormulas() {
        if (!this.formulaCells || this.formulaCells.length === 0) {
            console.log('No formulas to process');
            return;
        }
        
        console.log(`Processing ${this.formulaCells.length} formulas...`);
        
        try {
            // Sort formulas by dependency depth
            const sortedFormulas = this.topologicalSort(this.formulaCells);
            
            // Process only the first few formulas to avoid blocking
            const maxFormulas = Math.min(10, sortedFormulas.length);
            
            for (let i = 0; i < maxFormulas; i++) {
                const cellData = sortedFormulas[i];
                const td = document.querySelector(`[data-address="${cellData.cell_address}"]`);
                if (td && cellData.formula) {
                    try {
                        // Skip complex formulas on initial load
                        if (!cellData.formula.includes('VLOOKUP') && !cellData.formula.includes('INDEX')) {
                            const result = this.evaluateSimpleFormula(cellData.formula, cellData.cell_address);
                            if (result && result !== '#ERROR') {
                                td.textContent = result;
                            }
                        }
                    } catch (error) {
                        console.debug(`Skipping formula for ${cellData.cell_address}`);
                    }
                }
            }
        } catch (error) {
            console.warn('Formula processing error (non-critical):', error);
        }
    },
    
    /**
     * Evaluate simple formulas locally
     */
    evaluateSimpleFormula(formula, cellAddress) {
        // Only handle very simple formulas locally
        if (formula.startsWith('=') && /^=[A-Z]\d+$/.test(formula)) {
            // Simple cell reference
            const refCell = formula.substring(1);
            const refTd = document.querySelector(`[data-address="${refCell}"]`);
            return refTd ? refTd.textContent : '';
        }
        return null;
    },
    
    /**
     * Topological sort for formula dependencies
     */
    topologicalSort(formulaCells) {
        const sorted = [];
        const visited = new Set();
        const visiting = new Set();
        
        const visit = (cellData) => {
            if (visited.has(cellData.cell_address)) return;
            if (visiting.has(cellData.cell_address)) {
                console.warn('Circular dependency detected at:', cellData.cell_address);
                return;
            }
            
            visiting.add(cellData.cell_address);
            
            // Visit dependencies first
            const deps = this.dependencyMap[cellData.cell_address] || [];
            deps.forEach(dep => {
                const depCell = formulaCells.find(c => c.cell_address === dep);
                if (depCell) {
                    visit(depCell);
                }
            });
            
            visiting.delete(cellData.cell_address);
            visited.add(cellData.cell_address);
            sorted.push(cellData);
        };
        
        formulaCells.forEach(cellData => visit(cellData));
        
        return sorted;
    },
    
    /**
     * Process dependent cells with cascade updates
     */
    processDependentCells(changedCell, newValue) {
        console.log(`Processing dependencies for ${changedCell} with value: ${newValue}`);
        
        // Get cells that depend on the changed cell
        const dependentCells = this.reverseDependencyMap[changedCell] || [];
        console.log(`Found ${dependentCells.length} dependent cells:`, dependentCells);
        
        // Also check for special Excel patterns
        const additionalDependents = this.findAdditionalDependents(changedCell, newValue);
        
        // Combine all dependents
        const allDependents = [...new Set([...dependentCells, ...additionalDependents])];
        
        // Process each dependent cell
        allDependents.forEach(async cellAddress => {
            const td = document.querySelector(`[data-address="${cellAddress}"]`);
            if (td) {
                const formula = td.title || '';
                if (formula) {
                    await this.evaluateAndUpdateCell({
                        address: cellAddress,
                        element: td,
                        formula: formula,
                        triggeredBy: changedCell
                    });
                }
            }
        });
        
        // Handle cascade updates (dependents of dependents)
        this.processCascadeUpdates(allDependents, changedCell);
    },
    
    /**
     * Find additional dependent cells using pattern matching
     */
    findAdditionalDependents(changedCell, newValue) {
        const dependents = [];
        
        // Special handling for D4 and D5
        if (changedCell === 'D4' || changedCell === 'D5') {
            // Check all VLOOKUP formulas that might use these as lookup values
            document.querySelectorAll('td[title*="VLOOKUP"]').forEach(td => {
                const formula = td.title;
                if (formula.includes(changedCell)) {
                    dependents.push(td.dataset.address);
                }
            });
            
            // Check INDEX/MATCH formulas
            document.querySelectorAll('td[title*="INDEX"], td[title*="MATCH"]').forEach(td => {
                const formula = td.title;
                if (formula.includes(changedCell)) {
                    dependents.push(td.dataset.address);
                }
            });
        }
        
        return dependents;
    },
    
    /**
     * Process cascade updates for multiple levels of dependencies
     */
    async processCascadeUpdates(updatedCells, originalTrigger, depth = 0) {
        if (depth > 10) {
            console.warn('Maximum cascade depth reached, stopping to prevent infinite loop');
            return;
        }
        
        const nextLevelDependents = [];
        
        updatedCells.forEach(cellAddress => {
            const deps = this.reverseDependencyMap[cellAddress] || [];
            deps.forEach(dep => {
                if (!updatedCells.includes(dep) && dep !== originalTrigger) {
                    nextLevelDependents.push(dep);
                }
            });
        });
        
        if (nextLevelDependents.length > 0) {
            console.log(`Processing cascade level ${depth + 1} with ${nextLevelDependents.length} cells`);
            
            for (const cellAddress of nextLevelDependents) {
                const td = document.querySelector(`[data-address="${cellAddress}"]`);
                if (td && td.title) {
                    await this.evaluateAndUpdateCell({
                        address: cellAddress,
                        element: td,
                        formula: td.title,
                        cascadeLevel: depth + 1
                    });
                }
            }
            
            // Recursively process next level
            await this.processCascadeUpdates(nextLevelDependents, originalTrigger, depth + 1);
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
        const formula = td.title || '';
        
        const input = document.createElement('input');
        input.type = 'text';
        input.value = formula || currentValue;
        input.style.width = '100%';
        input.style.height = '100%';
        input.style.border = 'none';
        input.style.padding = '2px';
        
        td.innerHTML = '';
        td.appendChild(input);
        input.focus();
        input.select();
        
        input.onblur = () => {
            td.textContent = currentValue;
        };
        
        input.onkeydown = (e) => {
            if (e.key === 'Enter') {
                td.textContent = input.value;
                this.saveCell(address, input.value, td);
            } else if (e.key === 'Escape') {
                td.textContent = currentValue;
            }
        };
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
        console.log('Recalculate formulas with dependency resolution');
        if (this.currentWorksheet) {
            this.recalculateAll();
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
    },
    
    /**
     * Column number to letter conversion
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
     * Column letter to number conversion
     */
    columnToNumber(col) {
        let num = 0;
        for (let i = 0; i < col.length; i++) {
            num = num * 26 + (col.charCodeAt(i) - 64);
        }
        return num;
    },
    
    /**
     * Get cells from a range string
     */
    getCellsFromRange(range) {
        const cells = [];
        range = range.replace(/\$/g, '');
        
        if (!range.includes(':')) {
            cells.push(range);
            return cells;
        }
        
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
    }
};

// Make globally available
window.ExcelDashboard = ExcelDashboard;
console.log('Dashboard with validation sidebar loaded - ready to use');

// Debug helper function
window.testValidation = function(cellAddress) {
    console.log('Testing validation for', cellAddress);
    const select = document.querySelector(`#validation-select-${cellAddress}`);
    if (select) {
        console.log('Select found:', select);
        console.log('Options:', Array.from(select.options).map(o => o.value));
        console.log('Current value:', select.value);
        console.log('Event handlers:', select.onchange ? 'Present' : 'Missing');
    } else {
        console.log('Select not found for', cellAddress);
    }
    
    // Test fetching values directly
    const validationData = ExcelDashboard.getKnownValidationData(cellAddress);
    if (validationData) {
        console.log('Validation data:', validationData);
        ExcelDashboard.fetchValidationValues(validationData.formula1, (values) => {
            console.log('Fetched values:', values.length, 'items');
            if (values.length > 0) {
                console.log('First 5 values:', values.slice(0, 5));
            }
        });
    }
};

// Auto-test after load
window.addEventListener('load', function() {
    setTimeout(() => {
        console.log('Running validation diagnostics...');
        console.log('Current workbook:', ExcelDashboard.currentWorkbook);
        console.log('Current worksheet:', ExcelDashboard.currentWorksheet);
        
        // Check if dropdowns exist
        const d4Select = document.querySelector('#validation-select-D4');
        const d5Select = document.querySelector('#validation-select-D5');
        
        if (d4Select) {
            console.log('D4 dropdown found, options:', d4Select.options.length);
        } else {
            console.log('D4 dropdown not found');
        }
        
        if (d5Select) {
            console.log('D5 dropdown found, options:', d5Select.options.length);
        } else {
            console.log('D5 dropdown not found');
        }
    }, 3000);
});