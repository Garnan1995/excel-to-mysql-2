// Additional formula functions for Excel compatibility

// Extend the ExcelEngine with more Excel functions
ExcelEngine.prototype.evaluateSUM = function(range) {
    const data = this.getRange(range);
    return data.flat().reduce((sum, val) => sum + (parseFloat(val) || 0), 0);
};

ExcelEngine.prototype.evaluateAVERAGE = function(range) {
    const data = this.getRange(range);
    const values = data.flat().filter(val => !isNaN(val));
    return values.reduce((sum, val) => sum + parseFloat(val), 0) / values.length;
};

ExcelEngine.prototype.evaluateCOUNT = function(range) {
    const data = this.getRange(range);
    return data.flat().filter(val => !isNaN(val) && val !== '').length;
};

ExcelEngine.prototype.evaluateCOUNTA = function(range) {
    const data = this.getRange(range);
    return data.flat().filter(val => val !== '').length;
};

ExcelEngine.prototype.evaluateMAX = function(range) {
    const data = this.getRange(range);
    const values = data.flat().filter(val => !isNaN(val)).map(val => parseFloat(val));
    return Math.max(...values);
};

ExcelEngine.prototype.evaluateMIN = function(range) {
    const data = this.getRange(range);
    const values = data.flat().filter(val => !isNaN(val)).map(val => parseFloat(val));
    return Math.min(...values);
};

// Handle sheet references like 'Standar Jumlah'!$B$5:$B$83
ExcelEngine.prototype.evaluateSheetReference = function(formula) {
    const match = formula.match(/^'([^']+)'!(.+)$/);
    if (!match) return '#REF!';
    
    const sheetName = match[1];
    const range = match[2];
    
    // This would fetch data from the specified sheet
    // For now, return a placeholder
    return this.fetchSheetData(sheetName, range);
};

// Fetch data from another sheet
ExcelEngine.prototype.fetchSheetData = async function(sheetName, range) {
    // Make API call to get data from another sheet
    try {
        const response = await fetch(
            `/api/excel-api.php?action=sheet-range&sheet_name=${encodeURIComponent(sheetName)}&range=${encodeURIComponent(range)}`
        );
        const data = await response.json();
        return data.values || [];
    } catch (error) {
        console.error('Failed to fetch sheet data:', error);
        return '#REF!';
    }
};