// js/excel-engine.js

class ExcelEngine {
    constructor() {
        this.workbookId = null;
        this.worksheetId = null;
        this.cells = {};
        this.formulas = {};
        this.namedRanges = {};
        this.validations = {};
        this.mergedCells = [];
        this.dependencies = {};
        this.cache = {};
    }

    /**
     * Parse Excel formula and evaluate it
     */
    evaluateFormula(formula, cellAddress) {
        // Remove leading '='
        formula = formula.substring(1).trim();
        
        // Handle IFERROR
        if (formula.startsWith('IFERROR')) {
            return this.evaluateIFERROR(formula, cellAddress);
        }
        
        // Handle VLOOKUP
        if (formula.startsWith('VLOOKUP')) {
            return this.evaluateVLOOKUP(formula, cellAddress);
        }
        
        // Handle MATCH
        if (formula.startsWith('MATCH')) {
            return this.evaluateMATCH(formula, cellAddress);
        }
        
        // Handle direct cell references
        if (/^[A-Z]+\d+$/.test(formula)) {
            return this.getCellValue(formula);
        }
        
        // Handle ranges
        if (/^[A-Z]+\d+:[A-Z]+\d+$/.test(formula)) {
            return this.getRange(formula);
        }
        
        // Handle sheet references
        if (formula.includes('!')) {
            return this.evaluateSheetReference(formula);
        }
        
        // Handle basic arithmetic
        try {
            return this.evaluateExpression(formula);
        } catch (e) {
            return '#ERROR';
        }
    }

    /**
     * VLOOKUP implementation
     */
    evaluateVLOOKUP(formula, cellAddress) {
        // Parse VLOOKUP(lookup_value, table_array, col_index, [range_lookup])
        const params = this.parseFormulaParams(formula);
        if (params.length < 3) return '#VALUE!';
        
        const lookupValue = this.evaluateParameter(params[0]);
        const tableRange = params[1];
        const colIndex = parseInt(this.evaluateParameter(params[2]));
        const exactMatch = params[3] ? params[3].toLowerCase() === 'false' || params[3] === '0' : false;
        
        // Parse the table range
        const rangeData = this.parseRangeReference(tableRange);
        if (!rangeData) return '#REF!';
        
        // Get data from the range
        const tableData = this.getRangeData(rangeData);
        
        // Perform lookup
        for (let row of tableData) {
            if (exactMatch) {
                if (row[0] == lookupValue) {
                    return row[colIndex - 1] || '#REF!';
                }
            } else {
                // Approximate match (assuming sorted data)
                if (row[0] <= lookupValue) {
                    return row[colIndex - 1] || '#REF!';
                }
            }
        }
        
        return '#N/A';
    }

    /**
     * MATCH implementation
     */
    evaluateMATCH(formula, cellAddress) {
        const params = this.parseFormulaParams(formula);
        if (params.length < 2) return '#VALUE!';
        
        const lookupValue = this.evaluateParameter(params[0]);
        const lookupRange = params[1];
        const matchType = params[2] ? parseInt(this.evaluateParameter(params[2])) : 1;
        
        const rangeData = this.parseRangeReference(lookupRange);
        if (!rangeData) return '#REF!';
        
        const values = this.getRangeData(rangeData).flat();
        
        for (let i = 0; i < values.length; i++) {
            if (matchType === 0) {
                // Exact match
                if (values[i] == lookupValue) {
                    return i + 1;
                }
            } else if (matchType === 1) {
                // Less than or equal (assuming sorted ascending)
                if (values[i] <= lookupValue) {
                    if (i === values.length - 1 || values[i + 1] > lookupValue) {
                        return i + 1;
                    }
                }
            } else if (matchType === -1) {
                // Greater than or equal (assuming sorted descending)
                if (values[i] >= lookupValue) {
                    return i + 1;
                }
            }
        }
        
        return '#N/A';
    }

    /**
     * IFERROR implementation
     */
    evaluateIFERROR(formula, cellAddress) {
        const params = this.parseFormulaParams(formula);
        if (params.length < 2) return '#VALUE!';
        
        try {
            const result = this.evaluateParameter(params[0]);
            if (result && result.toString().startsWith('#')) {
                return this.evaluateParameter(params[1]);
            }
            return result;
        } catch (e) {
            return this.evaluateParameter(params[1]);
        }
    }

    /**
     * Parse formula parameters
     */
    parseFormulaParams(formula) {
        // Remove function name and parentheses
        const start = formula.indexOf('(');
        const end = formula.lastIndexOf(')');
        if (start === -1 || end === -1) return [];
        
        const paramsStr = formula.substring(start + 1, end);
        const params = [];
        let current = '';
        let depth = 0;
        let inQuotes = false;
        
        for (let i = 0; i < paramsStr.length; i++) {
            const char = paramsStr[i];
            
            if (char === '"') {
                inQuotes = !inQuotes;
            }
            
            if (!inQuotes) {
                if (char === '(') depth++;
                if (char === ')') depth--;
                if (char === ',' && depth === 0) {
                    params.push(current.trim());
                    current = '';
                    continue;
                }
            }
            
            current += char;
        }
        
        if (current.trim()) {
            params.push(current.trim());
        }
        
        return params;
    }

    /**
     * Evaluate a parameter (could be a value, cell reference, or nested formula)
     */
    evaluateParameter(param) {
        param = param.trim();
        
        // Remove leading/trailing semicolons from separator
        param = param.replace(/^;|;$/g, '');
        
        // Direct number
        if (!isNaN(param)) {
            return parseFloat(param);
        }
        
        // String literal
        if (param.startsWith('"') && param.endsWith('"')) {
            return param.slice(1, -1);
        }
        
        // Cell reference
        if (/^\$?[A-Z]+\$?\d+$/.test(param)) {
            return this.getCellValue(param.replace(/\$/g, ''));
        }
        
        // Formula
        if (param.includes('(')) {
            return this.evaluateFormula('=' + param, null);
        }
        
        return param;
    }

    /**
     * Parse range reference (e.g., 'Sheet1'!A1:C10 or A1:C10)
     */
    parseRangeReference(range) {
        let sheetName = null;
        let rangeStr = range;
        
        // Handle sheet reference
        if (range.includes('!')) {
            const parts = range.split('!');
            sheetName = parts[0].replace(/'/g, '');
            rangeStr = parts[1];
        }
        
        // Remove absolute references
        rangeStr = rangeStr.replace(/\$/g, '');
        
        // Parse range
        const rangeParts = rangeStr.split(':');
        if (rangeParts.length !== 2) return null;
        
        const start = this.parseCellAddress(rangeParts[0]);
        const end = this.parseCellAddress(rangeParts[1]);
        
        return {
            sheet: sheetName,
            startRow: start.row,
            startCol: start.col,
            endRow: end.row,
            endCol: end.col
        };
    }

    /**
     * Parse cell address to row and column
     */
    parseCellAddress(address) {
        const match = address.match(/([A-Z]+)(\d+)/i);
        if (!match) return { row: 0, col: 0 };
        
        const col = this.columnToNumber(match[1]);
        const row = parseInt(match[2]);
        
        return { row, col };
    }

    /**
     * Convert column letter to number (A=1, B=2, etc.)
     */
    columnToNumber(col) {
        let num = 0;
        for (let i = 0; i < col.length; i++) {
            num = num * 26 + (col.charCodeAt(i) - 64);
        }
        return num;
    }

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
    }

    /**
     * Get cell value
     */
    getCellValue(address) {
        const cell = this.cells[address];
        if (!cell) return '';
        
        if (cell.formula) {
            // Check cache
            if (this.cache[address]) {
                return this.cache[address];
            }
            
            // Evaluate formula
            const result = this.evaluateFormula(cell.formula, address);
            this.cache[address] = result;
            return result;
        }
        
        return cell.value || '';
    }

    /**
     * Get range data
     */
    getRangeData(rangeInfo) {
        const data = [];
        
        for (let row = rangeInfo.startRow; row <= rangeInfo.endRow; row++) {
            const rowData = [];
            for (let col = rangeInfo.startCol; col <= rangeInfo.endCol; col++) {
                const address = this.numberToColumn(col) + row;
                rowData.push(this.getCellValue(address));
            }
            data.push(rowData);
        }
        
        return data;
    }

    /**
     * Set cell value
     */
    setCellValue(address, value, formula = null) {
        if (!this.cells[address]) {
            this.cells[address] = {};
        }
        
        this.cells[address].value = value;
        this.cells[address].formula = formula;
        
        // Clear cache for this cell and dependents
        this.clearCache(address);
        
        // Recalculate dependents
        this.recalculateDependents(address);
    }

    /**
     * Clear cache for cell and its dependents
     */
    clearCache(address) {
        delete this.cache[address];
        
        if (this.dependencies[address]) {
            for (let dependent of this.dependencies[address]) {
                this.clearCache(dependent);
            }
        }
    }

    /**
     * Track formula dependencies
     */
    trackDependency(formula, cellAddress) {
        // Extract all cell references from formula
        const references = formula.match(/\$?[A-Z]+\$?\d+/gi) || [];
        
        for (let ref of references) {
            ref = ref.replace(/\$/g, '');
            if (!this.dependencies[ref]) {
                this.dependencies[ref] = [];
            }
            if (!this.dependencies[ref].includes(cellAddress)) {
                this.dependencies[ref].push(cellAddress);
            }
        }
    }

    /**
     * Recalculate dependent cells
     */
    recalculateDependents(address) {
        if (this.dependencies[address]) {
            for (let dependent of this.dependencies[address]) {
                const cell = this.cells[dependent];
                if (cell && cell.formula) {
                    const result = this.evaluateFormula(cell.formula, dependent);
                    this.cells[dependent].calculatedValue = result;
                }
            }
        }
    }

    /**
     * Load cells from API response
     */
    loadCells(cellsData) {
        this.cells = {};
        this.formulas = {};
        
        for (let cellData of cellsData) {
            const address = cellData.cell_address;
            
            this.cells[address] = {
                value: cellData.cell_value,
                formula: cellData.formula,
                type: cellData.cell_type,
                style: {
                    fontBold: cellData.font_bold,
                    fontItalic: cellData.font_italic,
                    fontColor: cellData.font_color,
                    fillColor: cellData.fill_color,
                    alignment: cellData.alignment_horizontal,
                    numberFormat: cellData.number_format
                }
            };
            
            if (cellData.formula) {
                this.formulas[address] = cellData.formula;
                this.trackDependency(cellData.formula, address);
            }
        }
    }

    /**
     * Evaluate basic expression (numbers and operators)
     */
    evaluateExpression(expr) {
        // Replace cell references with values
        expr = expr.replace(/\$?[A-Z]+\$?\d+/gi, (match) => {
            const value = this.getCellValue(match.replace(/\$/g, ''));
            return isNaN(value) ? 0 : value;
        });
        
        // Safely evaluate the expression
        try {
            // Only allow numbers, operators, and parentheses
            if (!/^[\d\s+\-*/().,]+$/.test(expr)) {
                return '#ERROR';
            }
            return Function('"use strict"; return (' + expr + ')')();
        } catch (e) {
            return '#ERROR';
        }
    }
}

// Export for use in other scripts
window.ExcelEngine = ExcelEngine;