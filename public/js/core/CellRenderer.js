// ===================================================================
// js/core/CellRenderer.js - Handles cell rendering and styling
// ===================================================================
class CellRenderer {
    constructor() {
        this.showFormulas = false;
    }

    renderCell(td, cellData) {
        if (!td || !cellData) return;

        // Content
        const content = this.getCellContent(cellData);
        td.textContent = content.display;
        if (content.title) td.title = content.title;
        if (content.className) td.className = content.className;

        // Styling
        this.applyCellStyles(td, cellData);
    }

    getCellContent(cellData) {
        const result = { display: '', title: '', className: '' };

        if (this.showFormulas && cellData.formula) {
            result.display = cellData.formula;
            result.className = 'formula-cell';
        } else if (cellData.formula) {
            result.display = cellData.formatted_value || cellData.cell_value || '';
            result.className = 'formula-cell';
            result.title = cellData.formula;
        } else if (cellData.cell_value !== null && cellData.cell_value !== undefined) {
            result.display = cellData.cell_value;
        }

        return result;
    }

    applyCellStyles(td, cellData) {
        const styles = {};
        const classes = [];

        // Font styles
        if (cellData.font_bold == 1) styles.fontWeight = 'bold';
        if (cellData.font_italic == 1) styles.fontStyle = 'italic';

        // Colors
        if (cellData.font_color && cellData.font_color !== 'FF000000') {
            styles.color = '#' + cellData.font_color.substring(2);
        }
        if (cellData.fill_color && !['FFFFFFFF', 'FF000000'].includes(cellData.fill_color)) {
            styles.backgroundColor = '#' + cellData.fill_color.substring(2);
        }

        // Alignment
        if (cellData.alignment_horizontal) {
            styles.textAlign = cellData.alignment_horizontal;
        }

        // Borders
        ['top', 'right', 'bottom', 'left'].forEach(side => {
            if (cellData[`border_${side}`] && cellData[`border_${side}`] !== 'none') {
                classes.push(`border-${side}-thick`);
            }
        });

        // Apply styles
        Object.assign(td.style, styles);
        classes.forEach(cls => td.classList.add(cls));
    }

    toggleFormulas() {
        this.showFormulas = !this.showFormulas;
        return this.showFormulas;
    }
}
