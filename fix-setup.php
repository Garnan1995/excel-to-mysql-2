<?php
// fix-setup.php - Run this to fix your setup
echo "Excel to MySQL Setup Fixer\n";
echo "==========================\n\n";

// 1. Check and rename .env.txt to .env if needed
$projectDir = __DIR__;
if (file_exists($projectDir . '/.env.txt') && !file_exists($projectDir . '/.env')) {
    echo "Renaming .env.txt to .env... ";
    if (rename($projectDir . '/.env.txt', $projectDir . '/.env')) {
        echo "OK\n";
    } else {
        echo "FAILED (please rename manually)\n";
    }
} elseif (file_exists($projectDir . '/.env')) {
    echo ".env file already exists: OK\n";
} else {
    echo "Creating .env file... ";
    $envContent = "# Database Configuration
DB_HOST=localhost
DB_PORT=3306
DB_DATABASE=excel_storage
DB_USERNAME=root
DB_PASSWORD=

# Application Settings
APP_ENV=development
APP_DEBUG=true
DEFAULT_TIMEZONE=America/New_York";
    
    file_put_contents($projectDir . '/.env', $envContent);
    echo "OK\n";
}

// 2. Check if vendor directory exists
if (!file_exists($projectDir . '/vendor/autoload.php')) {
    echo "\nERROR: Vendor directory not found!\n";
    echo "Please run: composer install\n";
    exit(1);
}

// 3. Create schema.sql if it doesn't exist
if (!file_exists($projectDir . '/src/schema.sql')) {
    echo "\nCreating schema.sql file... ";
    $schema = "-- Main workbook table
CREATE TABLE IF NOT EXISTS workbooks (
    id INT AUTO_INCREMENT PRIMARY KEY,
    filename VARCHAR(255),
    file_hash VARCHAR(64),
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
    metadata JSON
);

-- Worksheets table
CREATE TABLE IF NOT EXISTS worksheets (
    id INT AUTO_INCREMENT PRIMARY KEY,
    workbook_id INT,
    sheet_name VARCHAR(255),
    sheet_index INT,
    sheet_state VARCHAR(50),
    tab_color VARCHAR(8),
    zoom_scale INT,
    properties JSON,
    FOREIGN KEY (workbook_id) REFERENCES workbooks(id) ON DELETE CASCADE,
    INDEX idx_workbook_sheet (workbook_id, sheet_index)
);

-- Cells table
CREATE TABLE IF NOT EXISTS cells (
    id BIGINT AUTO_INCREMENT PRIMARY KEY,
    worksheet_id INT,
    cell_address VARCHAR(20),
    row_num INT,
    col_num INT,
    cell_value TEXT,
    cell_type VARCHAR(50),
    formatted_value TEXT,
    formula TEXT,
    formula_attributes JSON,
    hyperlink VARCHAR(500),
    comment TEXT,
    FOREIGN KEY (worksheet_id) REFERENCES worksheets(id) ON DELETE CASCADE,
    INDEX idx_worksheet_cell (worksheet_id, row_num, col_num),
    INDEX idx_cell_address (worksheet_id, cell_address)
);

-- Cell styles table
CREATE TABLE IF NOT EXISTS cell_styles (
    id INT AUTO_INCREMENT PRIMARY KEY,
    cell_id BIGINT,
    font_name VARCHAR(100),
    font_size DECIMAL(5,2),
    font_bold BOOLEAN DEFAULT FALSE,
    font_italic BOOLEAN DEFAULT FALSE,
    font_underline VARCHAR(20),
    font_strike BOOLEAN DEFAULT FALSE,
    font_color VARCHAR(8),
    fill_type VARCHAR(50),
    fill_color VARCHAR(8),
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
CREATE TABLE IF NOT EXISTS data_validations (
    id INT AUTO_INCREMENT PRIMARY KEY,
    worksheet_id INT,
    cell_range VARCHAR(50),
    validation_type VARCHAR(50),
    operator VARCHAR(20),
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
    error_style VARCHAR(20),
    FOREIGN KEY (worksheet_id) REFERENCES worksheets(id) ON DELETE CASCADE,
    INDEX idx_worksheet_validation (worksheet_id)
);

-- Conditional formatting rules
CREATE TABLE IF NOT EXISTS conditional_formats (
    id INT AUTO_INCREMENT PRIMARY KEY,
    worksheet_id INT,
    cell_range VARCHAR(50),
    rule_type VARCHAR(50),
    priority INT,
    stop_if_true BOOLEAN DEFAULT FALSE,
    formula TEXT,
    operator VARCHAR(20),
    text VARCHAR(255),
    time_period VARCHAR(50),
    style_attributes JSON,
    FOREIGN KEY (worksheet_id) REFERENCES worksheets(id) ON DELETE CASCADE,
    INDEX idx_worksheet_conditional (worksheet_id)
);

-- Merged cells
CREATE TABLE IF NOT EXISTS merged_cells (
    id INT AUTO_INCREMENT PRIMARY KEY,
    worksheet_id INT,
    cell_range VARCHAR(50),
    top_left_cell VARCHAR(20),
    bottom_right_cell VARCHAR(20),
    FOREIGN KEY (worksheet_id) REFERENCES worksheets(id) ON DELETE CASCADE,
    INDEX idx_worksheet_merged (worksheet_id)
);

-- Named ranges
CREATE TABLE IF NOT EXISTS named_ranges (
    id INT AUTO_INCREMENT PRIMARY KEY,
    workbook_id INT,
    name VARCHAR(255),
    cell_range VARCHAR(255),
    scope VARCHAR(255),
    comment TEXT,
    FOREIGN KEY (workbook_id) REFERENCES workbooks(id) ON DELETE CASCADE,
    INDEX idx_workbook_named (workbook_id)
);

-- Charts
CREATE TABLE IF NOT EXISTS charts (
    id INT AUTO_INCREMENT PRIMARY KEY,
    worksheet_id INT,
    chart_name VARCHAR(255),
    chart_type VARCHAR(50),
    position_from VARCHAR(20),
    position_to VARCHAR(20),
    chart_data JSON,
    FOREIGN KEY (worksheet_id) REFERENCES worksheets(id) ON DELETE CASCADE,
    INDEX idx_worksheet_chart (worksheet_id)
);

-- Pivot tables
CREATE TABLE IF NOT EXISTS pivot_tables (
    id INT AUTO_INCREMENT PRIMARY KEY,
    worksheet_id INT,
    pivot_name VARCHAR(255),
    source_range VARCHAR(255),
    target_cell VARCHAR(20),
    configuration JSON,
    FOREIGN KEY (worksheet_id) REFERENCES worksheets(id) ON DELETE CASCADE,
    INDEX idx_worksheet_pivot (worksheet_id)
);

-- Row and column dimensions
CREATE TABLE IF NOT EXISTS dimensions (
    id INT AUTO_INCREMENT PRIMARY KEY,
    worksheet_id INT,
    dimension_type ENUM('row', 'column'),
    dimension_index INT,
    height_or_width DECIMAL(10,2),
    hidden BOOLEAN DEFAULT FALSE,
    outline_level INT,
    collapsed BOOLEAN DEFAULT FALSE,
    FOREIGN KEY (worksheet_id) REFERENCES worksheets(id) ON DELETE CASCADE,
    INDEX idx_worksheet_dimension (worksheet_id, dimension_type, dimension_index)
);

-- Images and objects
CREATE TABLE IF NOT EXISTS embedded_objects (
    id INT AUTO_INCREMENT PRIMARY KEY,
    worksheet_id INT,
    object_type VARCHAR(50),
    object_name VARCHAR(255),
    position_from VARCHAR(20),
    position_to VARCHAR(20),
    width INT,
    height INT,
    object_data LONGBLOB,
    properties JSON,
    FOREIGN KEY (worksheet_id) REFERENCES worksheets(id) ON DELETE CASCADE,
    INDEX idx_worksheet_object (worksheet_id)
);

-- Macros and VBA
CREATE TABLE IF NOT EXISTS macros (
    id INT AUTO_INCREMENT PRIMARY KEY,
    workbook_id INT,
    macro_name VARCHAR(255),
    macro_type VARCHAR(50),
    macro_code LONGTEXT,
    FOREIGN KEY (workbook_id) REFERENCES workbooks(id) ON DELETE CASCADE,
    INDEX idx_workbook_macro (workbook_id)
);";
    
    file_put_contents($projectDir . '/src/schema.sql', $schema);
    echo "OK\n";
} else {
    echo "\nschema.sql already exists: OK\n";
}

// 4. Test database connection
echo "\n----------------------------------------\n";
echo "Please enter your MySQL password (press Enter for no password): ";
$password = trim(fgets(STDIN));

// Update .env with the password
$envFile = $projectDir . '/.env';
if (file_exists($envFile)) {
    $envContent = file_get_contents($envFile);
    $envContent = preg_replace('/DB_PASSWORD=.*/', 'DB_PASSWORD=' . $password, $envContent);
    file_put_contents($envFile, $envContent);
    echo "Updated .env file with password\n";
}

echo "\n✅ Setup fixed! Now run:\n";
echo "   php scripts/setup-database.php\n";
echo "   php scripts/extract.php input/yourfile.xlsx\n";