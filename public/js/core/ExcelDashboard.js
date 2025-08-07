// ===================================================================
// js/core/ExcelDashboard.js - Main controller (simplified)
// ===================================================================
class ExcelDashboard {
    constructor(config) {
        this.config = config;
        this.currentWorkbook = null;
        this.currentWorksheet = null;
        
        // Initialize services
        this.dataService = new DataService(config.apiUrl);
        this.cellRenderer = new CellRenderer();
        this.gridManager = new GridManager(
            document.getElementById('spreadsheetContainer'),
            this.cellRenderer
        );
        this.validationManager = new ValidationManager(this.dataService);
        
        // Setup callbacks
        this.gridManager.onCellSelect = this.handleCellSelect.bind(this);
        this.gridManager.onCellEdit = this.handleCellEdit.bind(this);
    }

    async init() {
        console.log('Initializing Excel Dashboard...');
        
        try {
            this.setupEventListeners();
            await this.loadWorkbooks();
            console.log('Dashboard initialized successfully');
        } catch (error) {
            console.error('Failed to initialize dashboard:', error);
            this.showError('Failed to initialize: ' + error.message);
        }
    }

    async loadWorkbooks() {
        this.showLoading(true, 'Loading workbooks...');
        
        try {
            const workbooks = await this.dataService.getWorkbooks();
            this.populateWorkbookSelector(workbooks);
            
            if (workbooks.length > 0) {
                await this.loadWorkbook(workbooks[0].id);
            }
        } finally {
            this.showLoading(false);
        }
    }

    async loadWorkbook(workbookId) {
        this.showLoading(true, 'Loading worksheets...');
        
        try {
            this.currentWorkbook = workbookId;
            const worksheets = await this.dataService.getWorksheets(workbookId);
            
            this.createSheetTabs(worksheets);
            if (worksheets.length > 0) {
                await this.loadWorksheet(worksheets[0].id);
            }
        } finally {
            this.showLoading(false);
        }
    }

    async loadWorksheet(worksheetId) {
        this.showLoading(true, 'Loading worksheet data...');
        
        try {
            this.currentWorksheet = worksheetId;
            this.updateActiveTab(worksheetId);
            
            const data = await this.dataService.getCellRange(worksheetId);
            
            // Load validations
            await this.validationManager.loadValidations(worksheetId);
            
            // Render grid
            this.gridManager.render(data);
            
            // Apply validations to cells
            this.validationManager.validations.forEach((validation, cellAddress) => {
                const td = document.querySelector(`[data-address="${cellAddress}"]`);
                if (td) this.validationManager.applyToCell(td, cellAddress);
            });
            
            this.updateStatistics(data.cells);
        } finally {
            this.showLoading(false);
        }
    }

    handleCellSelect(address, cell) {
        const cellRef = document.getElementById('cellReference');
        if (cellRef) cellRef.textContent = address;
        
        const formulaInput = document.getElementById('formulaInput');
        if (formulaInput) {
            formulaInput.value = cell.title || cell.textContent || '';
        }
    }

    async handleCellEdit(address, td) {
        const validation = this.validationManager.getValidation(address);
        
        if (validation) {
            await this.showValidationDropdown(td, address, validation);
        } else {
            this.showCellEditor(td, address);
        }
    }

    showCellEditor(td, address) {
        const currentValue = td.textContent;
        const formula = td.title || '';
        
        const input = document.createElement('input');
        input.type = 'text';
        input.value = formula || currentValue;
        input.style.cssText = 'width:100%;height:100%;border:none;padding:2px';
        
        td.innerHTML = '';
        td.appendChild(input);
        input.focus();
        input.select();
        
        const save = async () => {
            const newValue = input.value;
            td.textContent = newValue;
            
            await this.dataService.updateCell(
                this.currentWorksheet,
                address,
                newValue,
                newValue.startsWith('=') ? newValue : null
            );
            
            this.updateLastSaved();
        };
        
        input.addEventListener('blur', save);
        input.addEventListener('keydown', (e) => {
            if (e.key === 'Enter') {
                e.preventDefault();
                save();
            } else if (e.key === 'Escape') {
                td.textContent = currentValue;
            }
        });
    }

    async showValidationDropdown(td, address, validation) {
        // Implementation for validation dropdown
        console.log('Show validation dropdown for', address);
    }

    populateWorkbookSelector(workbooks) {
        const selector = document.getElementById('workbookSelector');
        if (!selector) return;
        
        selector.innerHTML = '<option value="">Select Workbook</option>';
        workbooks.forEach(wb => {
            const option = document.createElement('option');
            option.value = wb.id;
            option.textContent = wb.filename;
            selector.appendChild(option);
        });
        
        if (workbooks.length > 0) {
            selector.value = workbooks[0].id;
        }
    }

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
    }

    updateActiveTab(worksheetId) {
        document.querySelectorAll('.sheet-tab').forEach(tab => {
            tab.classList.toggle('active', tab.dataset.worksheetId == worksheetId);
        });
    }

    updateStatistics(cells) {
        if (!cells) return;
        
        const totalCells = cells.filter(c => c.cell_value !== null || c.formula !== null).length;
        const formulaCells = cells.filter(c => c.formula !== null).length;
        
        const cellCount = document.getElementById('cellCount');
        if (cellCount) cellCount.textContent = `${totalCells} cells`;
        
        const formulaCount = document.getElementById('formulaCount');
        if (formulaCount) formulaCount.textContent = `${formulaCells} formulas`;
    }

    updateLastSaved() {
        const lastSaved = document.getElementById('lastSaved');
        if (lastSaved) {
            lastSaved.textContent = 'Last saved: ' + new Date().toLocaleTimeString();
        }
    }

    showLoading(show, text) {
        const overlay = document.getElementById('loadingOverlay');
        if (!overlay) return;
        
        overlay.classList.toggle('active', show);
        if (text) {
            const textEl = document.getElementById('loadingText');
            if (textEl) textEl.textContent = text;
        }
    }

    showError(message) {
        console.error(message);
        const statusEl = document.getElementById('statusMessage');
        if (statusEl) statusEl.textContent = 'Error: ' + message;
    }

    setupEventListeners() {
        const selector = document.getElementById('workbookSelector');
        if (selector) {
            selector.addEventListener('change', (e) => {
                if (e.target.value) this.loadWorkbook(e.target.value);
            });
        }

        // Keyboard shortcuts
        document.addEventListener('keydown', (e) => {
            if (e.ctrlKey && e.key === 's') {
                e.preventDefault();
                this.saveChanges();
            }
        });
    }

    // Public methods for toolbar
    toggleFormulas() {
        const showFormulas = this.cellRenderer.toggleFormulas();
        const btn = document.getElementById('formulaToggle');
        if (btn) btn.classList.toggle('active', showFormulas);
        
        if (this.currentWorksheet) {
            this.loadWorksheet(this.currentWorksheet);
        }
    }

    async refreshData() {
        if (this.currentWorksheet) {
            await this.loadWorksheet(this.currentWorksheet);
        }
    }

    saveChanges() {
        this.updateLastSaved();
        console.log('Changes saved');
    }

    exportExcel() {
        if (this.currentWorkbook) {
            window.location.href = `${this.config.apiUrl}?action=export&workbook_id=${this.currentWorkbook}`;
        }
    }
}
