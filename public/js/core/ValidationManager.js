// js/validation-manager.js - Handles D4 and D5 validation controls (WITHOUT cell updates)

const ValidationManager = {
    // Reference to main dashboard
    dashboard: null,
    
    // Validation configurations
    validationData: {
        D4: {
            formula: "'Standar Jumlah'!$B$5:$B$83",
            sheetName: "Standar Jumlah",
            range: "B5:B83",
            values: [],
            currentValue: '',
            label: "Aircraft Type"
        },
        D5: {
            formula: "'List Aircraft'!$R$3:$R$42",
            sheetName: "List Aircraft",
            range: "R3:R42",
            values: [],
            currentValue: '',
            label: "Registration"
        }
    },
    
    // Cache for validation values
    cache: {},
    
    /**
     * Initialize the validation manager
     */
    init(dashboardInstance) {
        this.dashboard = dashboardInstance;
        console.log('ValidationManager initialized');
    },
    
    /**
     * Initialize validation dropdowns
     */
    async initializeDropdowns() {
        console.log('Initializing validation dropdowns for D4 and D5...');
        
        if (!this.dashboard.currentWorkbook) {
            console.warn('No workbook loaded, skipping validation initialization');
            return;
        }
        
        try {
            // Load values for both D4 and D5
            await Promise.all([
                this.loadValidationValues('D4'),
                this.loadValidationValues('D5')
            ]);
            
            // Populate the dropdowns
            this.populateDropdown('d4Select', 'D4');
            this.populateDropdown('d5Select', 'D5');
            
            console.log('Validation dropdowns initialized successfully');
        } catch (error) {
            console.error('Error initializing validation dropdowns:', error);
        }
    },
    
    /**
     * Load validation values for a specific cell
     */
    async loadValidationValues(cellAddress) {
        const config = this.validationData[cellAddress];
        if (!config) {
            console.error(`No validation config for ${cellAddress}`);
            return;
        }
        
        console.log(`Loading validation values for ${cellAddress} from ${config.sheetName}`);
        
        try {
            // Check cache first
            const cacheKey = `${this.dashboard.currentWorkbook}_${config.formula}`;
            if (this.cache[cacheKey]) {
                config.values = this.cache[cacheKey];
                console.log(`Using cached values for ${cellAddress}`);
                return;
            }
            
            // Fetch fresh values
            const values = await this.fetchValidationRange(config.sheetName, config.range);
            config.values = values;
            
            // Cache the results
            this.cache[cacheKey] = values;
            
            console.log(`Loaded ${values.length} values for ${cellAddress}`);
        } catch (error) {
            console.error(`Error loading validation for ${cellAddress}:`, error);
            config.values = [];
        }
    },
    
    /**
     * Fetch validation values from a specific sheet and range
     */
    async fetchValidationRange(sheetName, range) {
        if (!this.dashboard.currentWorkbook) {
            throw new Error('No workbook loaded');
        }
        
        try {
            // Get all worksheets
            const worksheetsResponse = await fetch(
                `${this.dashboard.apiUrl}?action=worksheets&workbook_id=${this.dashboard.currentWorkbook}`
            );
            
            if (!worksheetsResponse.ok) {
                throw new Error('Failed to fetch worksheets');
            }
            
            const worksheets = await worksheetsResponse.json();
            
            // Find the target sheet
            const targetSheet = worksheets.find(ws => 
                ws.sheet_name === sheetName || 
                ws.sheet_name.trim() === sheetName.trim()
            );
            
            if (!targetSheet) {
                console.error(`Sheet "${sheetName}" not found. Available sheets:`, 
                    worksheets.map(ws => ws.sheet_name));
                return [];
            }
            
            // Parse the range
            const [startCell, endCell] = range.split(':');
            const startMatch = startCell.match(/([A-Z]+)(\d+)/i);
            const endMatch = endCell.match(/([A-Z]+)(\d+)/i);
            
            if (!startMatch || !endMatch) {
                throw new Error(`Invalid range format: ${range}`);
            }
            
            const startCol = this.columnToNumber(startMatch[1]);
            const endCol = this.columnToNumber(endMatch[1]);
            const startRow = parseInt(startMatch[2]);
            const endRow = parseInt(endMatch[2]);
            
            // Fetch cells from the range
            const cellsResponse = await fetch(
                `${this.dashboard.apiUrl}?action=cell-range&worksheet_id=${targetSheet.id}` +
                `&start_row=${startRow}&end_row=${endRow}` +
                `&start_col=${startCol}&end_col=${endCol}`
            );
            
            if (!cellsResponse.ok) {
                throw new Error('Failed to fetch cell range');
            }
            
            const data = await cellsResponse.json();
            
            // Extract unique, non-empty values
            const values = [];
            const seen = new Set();
            
            if (data.cells) {
                data.cells.forEach(cell => {
                    if (cell.cell_value !== null && cell.cell_value !== '') {
                        const value = String(cell.cell_value).trim();
                        if (value && !seen.has(value)) {
                            seen.add(value);
                            values.push(value);
                        }
                    }
                });
            }
            
            return values.sort();
            
        } catch (error) {
            console.error(`Error fetching range ${sheetName}!${range}:`, error);
            return [];
        }
    },
    
    /**
     * Populate a dropdown with values
     */
    populateDropdown(selectId, cellAddress) {
        const select = document.getElementById(selectId);
        const config = this.validationData[cellAddress];
        
        if (!select || !config) return;
        
        // Clear existing options
        select.innerHTML = '<option value="">-- Select --</option>';
        
        // Add values as options
        config.values.forEach(value => {
            const option = document.createElement('option');
            option.value = value;
            option.textContent = value;
            
            if (value === config.currentValue) {
                option.selected = true;
            }
            
            select.appendChild(option);
        });
        
        console.log(`Populated ${selectId} with ${config.values.length} options`);
    },
    
    /**
     * Update current values from the spreadsheet
     */
    updateCurrentValues() {
        // Check D4 cell
        const d4Cell = document.querySelector('[data-address="D4"]');
        if (d4Cell) {
            const value = d4Cell.textContent;
            this.validationData.D4.currentValue = value;
            
            const d4Select = document.getElementById('d4Select');
            if (d4Select) {
                d4Select.value = value;
            }
        }
        
        // Check D5 cell
        const d5Cell = document.querySelector('[data-address="D5"]');
        if (d5Cell) {
            const value = d5Cell.textContent;
            this.validationData.D5.currentValue = value;
            
            const d5Select = document.getElementById('d5Select');
            if (d5Select) {
                d5Select.value = value;
            }
        }
    },
    
    /**
     * Handle D4 (Aircraft Type) change - MODIFIED: No cell update
     */
    async onD4Change(value) {
        console.log('D4 (Aircraft Type) dropdown changed to:', value);
        
        // Update sync status
        this.setSyncStatus('syncing');
        
        // Store the new value internally
        this.validationData.D4.currentValue = value;
        
        // DO NOT UPDATE THE CELL - REMOVED THIS SECTION
        // const d4Cell = document.querySelector('[data-address="D4"]');
        // if (d4Cell) {
        //     d4Cell.textContent = value;
        // }
        
        // DO NOT SAVE TO DATABASE - REMOVED THIS SECTION
        // await this.dashboard.saveCell('D4', value);
        
        // You can still process dependent cells if needed (optional)
        // await this.updateDependentCells('D4', value);
        
        // Update status
        this.setSyncStatus('synced');
        this.dashboard.showStatus(`Aircraft Type selected: ${value} (cell not updated)`);
    },
    
    /**
     * Handle D5 (Registration) change - MODIFIED: No cell update
     */
    async onD5Change(value) {
        console.log('D5 (Registration) dropdown changed to:', value);
        
        // Update sync status
        this.setSyncStatus('syncing');
        
        // Store the new value internally
        this.validationData.D5.currentValue = value;
        
        // DO NOT UPDATE THE CELL - REMOVED THIS SECTION
        // const d5Cell = document.querySelector('[data-address="D5"]');
        // if (d5Cell) {
        //     d5Cell.textContent = value;
        // }
        
        // DO NOT SAVE TO DATABASE - REMOVED THIS SECTION
        // await this.dashboard.saveCell('D5', value);
        
        // You can still process dependent cells if needed (optional)
        // await this.updateDependentCells('D5', value);
        
        // Update status
        this.setSyncStatus('synced');
        this.dashboard.showStatus(`Registration selected: ${value} (cell not updated)`);
    },
    
    /**
     * Update cells that depend on the changed cell - OPTIONAL: You can remove this entire function if not needed
     */
    async updateDependentCells(changedCell, newValue) {
        console.log(`Checking for cells dependent on ${changedCell}...`);
        
        // Since we're not updating the main cell, you might not want to update dependents either
        // Comment out or remove this function's content if you don't want any cascade updates
        
        /*
        // Find all cells with formulas that reference the changed cell
        const allCells = document.querySelectorAll('td[data-address]');
        const dependentCells = [];
        
        allCells.forEach(td => {
            const formula = td.title; // Formulas are stored in title attribute
            if (formula && formula.includes(changedCell)) {
                dependentCells.push({
                    address: td.dataset.address,
                    element: td,
                    formula: formula
                });
            }
        });
        
        console.log(`Found ${dependentCells.length} dependent cells`);
        */
        
        console.log('Dependent cell updates disabled');
    },
    
    /**
     * Recalculate formulas for dependent cells - OPTIONAL: Can be removed if not needed
     */
    async recalculateFormulas(dependentCells) {
        console.log('Formula recalculation disabled');
        // Removed all recalculation logic since we're not updating cells
    },
    
    /**
     * Handle cell update from spreadsheet - MODIFIED: Don't update dropdowns from cell changes
     */
    async handleCellUpdate(address, value) {
        // Optionally update the dropdown to reflect cell changes
        // Or comment this out if you don't want dropdowns to follow cell changes
        
        if (address === 'D4') {
            this.validationData.D4.currentValue = value;
            const select = document.getElementById('d4Select');
            if (select) select.value = value;
        } else if (address === 'D5') {
            this.validationData.D5.currentValue = value;
            const select = document.getElementById('d5Select');
            if (select) select.value = value;
        }
    },
    
    /**
     * Refresh dropdowns after workbook change
     */
    async refreshDropdowns() {
        console.log('Refreshing validation dropdowns...');
        
        // Clear cache for new workbook
        this.cache = {};
        
        // Reload validation values
        await this.initializeDropdowns();
        
        // Update current values from spreadsheet
        this.updateCurrentValues();
    },
    
    /**
     * Set sync status indicator
     */
    setSyncStatus(status) {
        const syncEl = document.getElementById('syncStatus');
        if (syncEl) {
            if (status === 'syncing') {
                syncEl.classList.add('syncing');
                syncEl.innerHTML = '<span>●</span><span>Syncing...</span>';
            } else {
                syncEl.classList.remove('syncing');
                syncEl.innerHTML = '<span>●</span><span>Synced</span>';
            }
        }
    },
    
    /**
     * Utility: Convert column letter to number
     */
    columnToNumber(col) {
        let num = 0;
        for (let i = 0; i < col.length; i++) {
            num = num * 26 + (col.charCodeAt(i) - 64);
        }
        return num;
    },
    
    /**
     * Clear validation cache
     */
    clearCache() {
        this.cache = {};
        console.log('Validation cache cleared');
    },
    
    /**
     * Get validation info for a cell
     */
    getValidationInfo(cellAddress) {
        return this.validationData[cellAddress] || null;
    },
    
    /**
     * Check if a cell has validation
     */
    hasValidation(cellAddress) {
        return cellAddress in this.validationData;
    },
    
    /**
     * Export current validation values
     */
    exportValidationData() {
        return {
            D4: {
                label: this.validationData.D4.label,
                currentValue: this.validationData.D4.currentValue,
                possibleValues: this.validationData.D4.values
            },
            D5: {
                label: this.validationData.D5.label,
                currentValue: this.validationData.D5.currentValue,
                possibleValues: this.validationData.D5.values
            }
        };
    }
};

// Make globally available
window.ValidationManager = ValidationManager;