// ===================================================================
// js/core/ValidationManager.js - Simplified validation handling
// ===================================================================
class ValidationManager {
    constructor(dataService) {
        this.dataService = dataService;
        this.validations = new Map();
    }

    async loadValidations(worksheetId) {
        const data = await this.dataService.getCellRange(worksheetId);
        if (data.validations) {
            this.processValidations(data.validations);
        }
    }

    processValidations(validations) {
        validations.forEach(validation => {
            if (validation.validation_type === 'list' && validation.show_dropdown) {
                const cells = this.getCellsFromRange(validation.cell_range);
                cells.forEach(cell => {
                    this.validations.set(cell, validation);
                });
            }
        });
    }

    getValidation(cellAddress) {
        return this.validations.get(cellAddress);
    }

    applyToCell(td, cellAddress) {
        const validation = this.validations.get(cellAddress);
        if (!validation) return;

        td.classList.add('has-validation');
        td.dataset.hasValidation = 'true';
        
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
        `;
        td.style.position = 'relative';
        td.appendChild(arrow);
    }

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

    columnToNumber(col) {
        let num = 0;
        for (let i = 0; i < col.length; i++) {
            num = num * 26 + (col.charCodeAt(i) - 64);
        }
        return num;
    }

    numberToColumn(num) {
        let col = '';
        while (num > 0) {
            num--;
            col = String.fromCharCode(65 + (num % 26)) + col;
            num = Math.floor(num / 26);
        }
        return col;
    }
}
