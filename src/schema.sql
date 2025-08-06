-- Main workbook table
CREATE TABLE workbooks (
    id INT AUTO_INCREMENT PRIMARY KEY,
    filename VARCHAR(255),
    file_hash VARCHAR(64),
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
    metadata JSON -- Store workbook-level properties
);

-- Worksheets table
CREATE TABLE worksheets (
    id INT AUTO_INCREMENT PRIMARY KEY,
    workbook_id INT,
    sheet_name VARCHAR(255),
    sheet_index INT,
    sheet_state VARCHAR(50), -- visible, hidden, veryHidden
    tab_color VARCHAR(8), -- Hex color
    zoom_scale INT,
    properties JSON, -- Additional sheet properties
    FOREIGN KEY (workbook_id) REFERENCES workbooks(id) ON DELETE CASCADE,
    INDEX idx_workbook_sheet (workbook_id, sheet_index)
);

-- Cells table - stores all cell data
CREATE TABLE cells (
    id BIGINT AUTO_INCREMENT PRIMARY KEY,
    worksheet_id INT,
    cell_address VARCHAR(20), -- e.g., 'A1', 'B2'
    row_num INT,
    col_num INT,
    cell_value TEXT, -- Raw value
    cell_type VARCHAR(50), -- string, numeric, boolean, formula, etc.
    formatted_value TEXT, -- Display value
    formula TEXT, -- Store formula if exists
    formula_attributes JSON, -- Formula metadata
    hyperlink VARCHAR(500),
    comment TEXT,
    FOREIGN KEY (worksheet_id) REFERENCES worksheets(id) ON DELETE CASCADE,
    INDEX idx_worksheet_cell (worksheet_id, row_num, col_num),
    INDEX idx_cell_address (worksheet_id, cell_address)
);

-- Cell styles table
CREATE TABLE cell_styles (
    id INT AUTO_INCREMENT PRIMARY KEY,
    cell_id BIGINT,
    font_name VARCHAR(100),
    font_size DECIMAL(5,2),
    font_bold BOOLEAN DEFAULT FALSE,
    font_italic BOOLEAN DEFAULT FALSE,
    font_underline VARCHAR(20),
    font_strike BOOLEAN DEFAULT FALSE,
    font_color VARCHAR(8), -- Hex color
    fill_type VARCHAR(50),
    fill_color VARCHAR(8), -- Background color
    fill_pattern VARCHAR(50),
    border_top VARCHAR(50),
    border_right VARCHAR(50),
    border_bottom VARCHAR(50),
    border_left VARCHAR(50),
    border_diagonal VARCHAR(50),
    border_color_top VARCHAR(8),
    border_color_right VARCHAR(8),
    border_color_bottom VARCHAR(8),
    border_color_left VARCHAR(8),
    alignment_horizontal VARCHAR(20),
    alignment_vertical VARCHAR(20),
    alignment_wrap_text BOOLEAN DEFAULT FALSE,
    alignment_text_rotation INT,
    alignment_indent INT,
    number_format VARCHAR(255),
    protection_locked BOOLEAN,
    protection_hidden BOOLEAN,
    FOREIGN KEY (cell_id) REFERENCES cells(id) ON DELETE CASCADE,
    INDEX idx_cell_style (cell_id)
);

-- Data validation rules
CREATE TABLE data_validations (
    id INT AUTO_INCREMENT PRIMARY KEY,
    worksheet_id INT,
    cell_range VARCHAR(50), -- e.g., 'A1:A10'
    validation_type VARCHAR(50), -- list, whole, decimal, date, time, textLength, custom
    operator VARCHAR(20), -- between, notBetween, equal, notEqual, etc.
    formula1 TEXT,
    formula2 TEXT,
    allow_blank BOOLEAN DEFAULT TRUE,
    show_dropdown BOOLEAN DEFAULT TRUE,
    show_input_message BOOLEAN DEFAULT FALSE,
    show_error_message BOOLEAN DEFAULT FALSE,
    error_title VARCHAR(255),
    error_message TEXT,
    input_title VARCHAR(255),
    input_message TEXT,
    error_style VARCHAR(20), -- stop, warning, information
    FOREIGN KEY (worksheet_id) REFERENCES worksheets(id) ON DELETE CASCADE,
    INDEX idx_worksheet_validation (worksheet_id)
);

-- Conditional formatting rules
CREATE TABLE conditional_formats (
    id INT AUTO_INCREMENT PRIMARY KEY,
    worksheet_id INT,
    cell_range VARCHAR(50),
    rule_type VARCHAR(50), -- cellIs, expression, containsText, colorScale, dataBar, iconSet
    priority INT,
    stop_if_true BOOLEAN DEFAULT FALSE,
    formula TEXT,
    operator VARCHAR(20),
    text VARCHAR(255),
    time_period VARCHAR(50),
    style_attributes JSON, -- Store formatting to apply
    FOREIGN KEY (worksheet_id) REFERENCES worksheets(id) ON DELETE CASCADE,
    INDEX idx_worksheet_conditional (worksheet_id)
);

-- Merged cells
CREATE TABLE merged_cells (
    id INT AUTO_INCREMENT PRIMARY KEY,
    worksheet_id INT,
    cell_range VARCHAR(50), -- e.g., 'A1:C3'
    top_left_cell VARCHAR(20),
    bottom_right_cell VARCHAR(20),
    FOREIGN KEY (worksheet_id) REFERENCES worksheets(id) ON DELETE CASCADE,
    INDEX idx_worksheet_merged (worksheet_id)
);

-- Named ranges
CREATE TABLE named_ranges (
    id INT AUTO_INCREMENT PRIMARY KEY,
    workbook_id INT,
    name VARCHAR(255),
    cell_range VARCHAR(255), -- Can include sheet reference
    scope VARCHAR(255), -- Workbook or specific worksheet
    comment TEXT,
    FOREIGN KEY (workbook_id) REFERENCES workbooks(id) ON DELETE CASCADE,
    INDEX idx_workbook_named (workbook_id)
);

-- Charts (store as JSON for complex structure)
CREATE TABLE charts (
    id INT AUTO_INCREMENT PRIMARY KEY,
    worksheet_id INT,
    chart_name VARCHAR(255),
    chart_type VARCHAR(50),
    position_from VARCHAR(20), -- Top-left cell
    position_to VARCHAR(20), -- Bottom-right cell
    chart_data JSON, -- Complete chart configuration
    FOREIGN KEY (worksheet_id) REFERENCES worksheets(id) ON DELETE CASCADE,
    INDEX idx_worksheet_chart (worksheet_id)
);

-- Pivot tables
CREATE TABLE pivot_tables (
    id INT AUTO_INCREMENT PRIMARY KEY,
    worksheet_id INT,
    pivot_name VARCHAR(255),
    source_range VARCHAR(255),
    target_cell VARCHAR(20),
    configuration JSON, -- Store complete pivot configuration
    FOREIGN KEY (worksheet_id) REFERENCES worksheets(id) ON DELETE CASCADE,
    INDEX idx_worksheet_pivot (worksheet_id)
);

-- Row and column dimensions
CREATE TABLE dimensions (
    id INT AUTO_INCREMENT PRIMARY KEY,
    worksheet_id INT,
    dimension_type ENUM('row', 'column'),
    dimension_index INT, -- Row number or column number
    height_or_width DECIMAL(10,2), -- Height for rows, width for columns
    hidden BOOLEAN DEFAULT FALSE,
    outline_level INT,
    collapsed BOOLEAN DEFAULT FALSE,
    FOREIGN KEY (worksheet_id) REFERENCES worksheets(id) ON DELETE CASCADE,
    INDEX idx_worksheet_dimension (worksheet_id, dimension_type, dimension_index)
);

-- Images and objects
CREATE TABLE embedded_objects (
    id INT AUTO_INCREMENT PRIMARY KEY,
    worksheet_id INT,
    object_type VARCHAR(50), -- image, shape, textbox, etc.
    object_name VARCHAR(255),
    position_from VARCHAR(20),
    position_to VARCHAR(20),
    width INT,
    height INT,
    object_data LONGBLOB, -- Store actual image/object data
    properties JSON, -- Additional properties
    FOREIGN KEY (worksheet_id) REFERENCES worksheets(id) ON DELETE CASCADE,
    INDEX idx_worksheet_object (worksheet_id)
);

-- Macros and VBA (if needed)
CREATE TABLE macros (
    id INT AUTO_INCREMENT PRIMARY KEY,
    workbook_id INT,
    macro_name VARCHAR(255),
    macro_type VARCHAR(50),
    macro_code LONGTEXT,
    FOREIGN KEY (workbook_id) REFERENCES workbooks(id) ON DELETE CASCADE,
    INDEX idx_workbook_macro (workbook_id)
);