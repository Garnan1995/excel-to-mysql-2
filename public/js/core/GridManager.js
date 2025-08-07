// ===================================================================
// js/core/GridManager.js - Manages the spreadsheet grid
// ===================================================================
class GridManager {
    constructor(container, cellRenderer) {
        this.container = container;
        this.cellRenderer = cellRenderer;
        this.selectedCell = null;
        this.data = null;
    }

    render(data) {
        this.data = data;
        this.clear();
        
        const table = this.createTable();
        const header = this.createHeader();
        const body = this.createBody(data);
        
        table.appendChild(header);
        table.appendChild(body);
        this.container.appendChild(table);

        if (data.merged_cells?.length) {
            this.applyMergedCells(data.merged_cells);
        }
    }

    createTable() {
        const table = document.createElement('table');
        table.className = 'spreadsheet';
        return table;
    }

    createHeader() {
        const thead = document.createElement('thead');
        const row = document.createElement('tr');
        row.innerHTML = '<th class="row-header"></th>';
        
        for (let col = 1; col <= 50; col++) {
            const th = document.createElement('th');
            th.textContent = this.numberToColumn(col);
            row.appendChild(th);
        }
        
        thead.appendChild(row);
        return thead;
    }

    createBody(data) {
        const tbody = document.createElement('tbody');
        
        for (let row = 1; row <= 100; row++) {
            const tr = this.createRow(row, data);
            tbody.appendChild(tr);
        }
        
        return tbody;
    }

    createRow(rowNum, data) {
        const tr = document.createElement('tr');
        
        // Row header
        const th = document.createElement('th');
        th.className = 'row-header';
        th.textContent = rowNum;
        tr.appendChild(th);
        
        // Data cells
        for (let col = 1; col <= 50; col++) {
            const td = this.createCell(rowNum, col, data);
            tr.appendChild(td);
        }
        
        return tr;
    }

    createCell(row, col, data) {
        const td = document.createElement('td');
        const address = this.numberToColumn(col) + row;
        td.dataset.address = address;
        
        if (data.cells) {
            const cellData = data.cells.find(c => 
                parseInt(c.row_num) === row && parseInt(c.col_num) === col
            );
            
            if (cellData) {
                this.cellRenderer.renderCell(td, cellData);
            }
        }
        
        td.addEventListener('click', () => this.selectCell(address));
        td.addEventListener('dblclick', () => this.editCell(address));
        
        return td;
    }

    selectCell(address) {
        // Remove previous selection
        if (this.selectedCell) {
            const prev = this.container.querySelector(`[data-address="${this.selectedCell}"]`);
            if (prev) prev.classList.remove('selected');
        }
        
        // Add new selection
        this.selectedCell = address;
        const cell = this.container.querySelector(`[data-address="${address}"]`);
        if (cell) {
            cell.classList.add('selected');
            this.onCellSelect?.(address, cell);
        }
    }

    editCell(address) {
        const td = this.container.querySelector(`[data-address="${address}"]`);
        if (!td) return;
        
        this.onCellEdit?.(address, td);
    }

    clear() {
        this.container.innerHTML = '';
    }

    applyMergedCells(mergedCells) {
        // Implementation for merged cells
        console.log(`Applying ${mergedCells.length} merged cells`);
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

    columnToNumber(col) {
        let num = 0;
        for (let i = 0; i < col.length; i++) {
            num = num * 26 + (col.charCodeAt(i) - 64);
        }
        return num;
    }
}
