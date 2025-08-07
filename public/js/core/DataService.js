// ===================================================================
// js/core/DataService.js - Handles all API communication
// ===================================================================
class DataService {
    constructor(apiUrl) {
        this.apiUrl = apiUrl;
        this.cache = new Map();
    }

    async request(action, params = {}, method = 'GET', body = null) {
        const url = new URL(this.apiUrl);
        url.searchParams.append('action', action);
        
        if (method === 'GET') {
            Object.entries(params).forEach(([key, value]) => {
                url.searchParams.append(key, value);
            });
        }

        const options = {
            method,
            headers: method === 'POST' ? { 'Content-Type': 'application/json' } : {}
        };

        if (body) options.body = JSON.stringify(body);

        const response = await fetch(url, options);
        if (!response.ok) throw new Error(`API Error: ${response.status}`);
        return response.json();
    }

    async getWorkbooks() {
        return this.request('workbooks');
    }

    async getWorksheets(workbookId) {
        const cacheKey = `worksheets_${workbookId}`;
        if (this.cache.has(cacheKey)) return this.cache.get(cacheKey);
        
        const data = await this.request('worksheets', { workbook_id: workbookId });
        this.cache.set(cacheKey, data);
        return data;
    }

    async getCellRange(worksheetId, startRow = 1, endRow = 100, startCol = 1, endCol = 50) {
        return this.request('cell-range', {
            worksheet_id: worksheetId,
            start_row: startRow,
            end_row: endRow,
            start_col: startCol,
            end_col: endCol
        });
    }

    async updateCell(worksheetId, cellAddress, value, formula = null) {
        return this.request('update-cell', {}, 'POST', {
            worksheet_id: worksheetId,
            cell_address: cellAddress,
            value: formula ? null : value,
            formula: formula
        });
    }

    clearCache() {
        this.cache.clear();
    }
}
