<?php
// api/excel-api.php
header('Content-Type: application/json');
header('Access-Control-Allow-Origin: *');
header('Access-Control-Allow-Methods: GET, POST, PUT, DELETE');
header('Access-Control-Allow-Headers: Content-Type');

require_once __DIR__ . '/../config/database.php';  // <-- Fixed path

class ExcelAPI {
    private $db;
    
    public function __construct() {
        $config = require __DIR__ . '/../config/database.php';  // <-- Also fix here
        try {
            $this->db = new PDO(
                "mysql:host={$config['host']};dbname={$config['database']};charset=utf8mb4",
                $config['username'],
                $config['password']
            );
            $this->db->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION);
        } catch (PDOException $e) {
            $this->sendError('Database connection failed');
        }
    }
    
    /**
     * Handle API requests
     */
    public function handleRequest() {
        $method = $_SERVER['REQUEST_METHOD'];
        $action = $_GET['action'] ?? '';
        
        switch ($action) {
            case 'workbooks':
                $this->getWorkbooks();
                break;
            case 'worksheets':
                $this->getWorksheets($_GET['workbook_id'] ?? 0);
                break;
            case 'cells':
                $this->getCells($_GET['worksheet_id'] ?? 0);
                break;
            case 'cell-range':
                $this->getCellRange(
                    $_GET['worksheet_id'] ?? 0,
                    $_GET['start_row'] ?? 1,
                    $_GET['end_row'] ?? 100,
                    $_GET['start_col'] ?? 1,
                    $_GET['end_col'] ?? 26
                );
                break;
            case 'update-cell':
                $this->updateCell();
                break;
            case 'evaluate-formula':
                $this->evaluateFormula();
                break;
            case 'vlookup':
                $this->performVLookup();
                break;
            case 'data-validation':
                $this->getDataValidation($_GET['worksheet_id'] ?? 0);
                break;
            case 'named-ranges':
                $this->getNamedRanges($_GET['workbook_id'] ?? 0);
                break;
            case 'export':
                $this->exportToExcel($_GET['workbook_id'] ?? 0);
                break;
            default:
                $this->sendError('Invalid action');
        }
    }
    
    /**
     * Get all workbooks
     */
    private function getWorkbooks() {
        $stmt = $this->db->query("
            SELECT id, filename, created_at, updated_at 
            FROM workbooks 
            ORDER BY updated_at DESC
        ");
        $workbooks = $stmt->fetchAll(PDO::FETCH_ASSOC);
        $this->sendResponse($workbooks);
    }
    
    /**
     * Get worksheets for a workbook
     */
    private function getWorksheets($workbookId) {
        $stmt = $this->db->prepare("
            SELECT id, sheet_name, sheet_index, tab_color, properties
            FROM worksheets 
            WHERE workbook_id = :workbook_id
            ORDER BY sheet_index
        ");
        $stmt->execute([':workbook_id' => $workbookId]);
        $worksheets = $stmt->fetchAll(PDO::FETCH_ASSOC);
        
        foreach ($worksheets as &$ws) {
            $ws['properties'] = json_decode($ws['properties'], true);
        }
        
        $this->sendResponse($worksheets);
    }
    
    /**
     * Get cells for a worksheet with pagination
     */
    private function getCellRange($worksheetId, $startRow, $endRow, $startCol, $endCol) {
        $stmt = $this->db->prepare("
            SELECT 
                c.id,
                c.cell_address,
                c.row_num,
                c.col_num,
                c.cell_value,
                c.cell_type,
                c.formatted_value,
                c.formula,
                c.formula_attributes,
                c.hyperlink,
                c.comment,
                cs.font_name,
                cs.font_size,
                cs.font_bold,
                cs.font_italic,
                cs.font_color,
                cs.fill_color,
                cs.border_top,
                cs.border_right,
                cs.border_bottom,
                cs.border_left,
                cs.alignment_horizontal,
                cs.alignment_vertical,
                cs.number_format
            FROM cells c
            LEFT JOIN cell_styles cs ON c.id = cs.cell_id
            WHERE c.worksheet_id = :worksheet_id
            AND c.row_num >= :start_row AND c.row_num <= :end_row
            AND c.col_num >= :start_col AND c.col_num <= :end_col
            ORDER BY c.row_num, c.col_num
        ");
        
        $stmt->execute([
            ':worksheet_id' => $worksheetId,
            ':start_row' => $startRow,
            ':end_row' => $endRow,
            ':start_col' => $startCol,
            ':end_col' => $endCol
        ]);
        
        $cells = $stmt->fetchAll(PDO::FETCH_ASSOC);
        
        // Parse JSON fields
        foreach ($cells as &$cell) {
            if ($cell['formula_attributes']) {
                $cell['formula_attributes'] = json_decode($cell['formula_attributes'], true);
            }
        }
        
        // Get merged cells
        $mergedStmt = $this->db->prepare("
            SELECT cell_range, top_left_cell, bottom_right_cell
            FROM merged_cells
            WHERE worksheet_id = :worksheet_id
        ");
        $mergedStmt->execute([':worksheet_id' => $worksheetId]);
        $mergedCells = $mergedStmt->fetchAll(PDO::FETCH_ASSOC);
        
        // Get data validations
        $validationStmt = $this->db->prepare("
            SELECT * FROM data_validations
            WHERE worksheet_id = :worksheet_id
        ");
        $validationStmt->execute([':worksheet_id' => $worksheetId]);
        $validations = $validationStmt->fetchAll(PDO::FETCH_ASSOC);
        
        $this->sendResponse([
            'cells' => $cells,
            'merged_cells' => $mergedCells,
            'validations' => $validations
        ]);
    }
    
    /**
     * Update a cell value
     */
    private function updateCell() {
        $data = json_decode(file_get_contents('php://input'), true);
        
        $worksheetId = $data['worksheet_id'] ?? 0;
        $cellAddress = $data['cell_address'] ?? '';
        $value = $data['value'] ?? '';
        $formula = $data['formula'] ?? null;
        
        // Check if cell exists
        $stmt = $this->db->prepare("
            SELECT id FROM cells 
            WHERE worksheet_id = :worksheet_id AND cell_address = :cell_address
        ");
        $stmt->execute([
            ':worksheet_id' => $worksheetId,
            ':cell_address' => $cellAddress
        ]);
        $cell = $stmt->fetch(PDO::FETCH_ASSOC);
        
        if ($cell) {
            // Update existing cell
            $updateStmt = $this->db->prepare("
                UPDATE cells 
                SET cell_value = :value,
                    formula = :formula,
                    formatted_value = :formatted_value,
                    cell_type = :cell_type
                WHERE id = :id
            ");
            
            $cellType = $formula ? 'f' : (is_numeric($value) ? 'n' : 's');
            
            $updateStmt->execute([
                ':value' => $value,
                ':formula' => $formula,
                ':formatted_value' => $value, // Will be calculated client-side
                ':cell_type' => $cellType,
                ':id' => $cell['id']
            ]);
        } else {
            // Insert new cell
            list($col, $row) = $this->parseAddress($cellAddress);
            
            $insertStmt = $this->db->prepare("
                INSERT INTO cells (worksheet_id, cell_address, row_num, col_num, 
                                 cell_value, cell_type, formula, formatted_value)
                VALUES (:worksheet_id, :cell_address, :row_num, :col_num,
                        :value, :cell_type, :formula, :formatted_value)
            ");
            
            $cellType = $formula ? 'f' : (is_numeric($value) ? 'n' : 's');
            
            $insertStmt->execute([
                ':worksheet_id' => $worksheetId,
                ':cell_address' => $cellAddress,
                ':row_num' => $row,
                ':col_num' => $col,
                ':value' => $value,
                ':cell_type' => $cellType,
                ':formula' => $formula,
                ':formatted_value' => $value
            ]);
        }
        
        $this->sendResponse(['success' => true]);
    }
    
    /**
     * Evaluate Excel formulas (simplified version)
     */
    private function evaluateFormula() {
        $data = json_decode(file_get_contents('php://input'), true);
        $formula = $data['formula'] ?? '';
        $worksheetId = $data['worksheet_id'] ?? 0;
        
        // Handle VLOOKUP
        if (strpos($formula, 'VLOOKUP') !== false) {
            $result = $this->evaluateVLookup($formula, $worksheetId);
        }
        // Handle IFERROR
        elseif (strpos($formula, 'IFERROR') !== false) {
            $result = $this->evaluateIferror($formula, $worksheetId);
        }
        // Handle MATCH
        elseif (strpos($formula, 'MATCH') !== false) {
            $result = $this->evaluateMatch($formula, $worksheetId);
        }
        // Handle basic cell references
        else {
            $result = $this->evaluateCellReference($formula, $worksheetId);
        }
        
        $this->sendResponse(['result' => $result]);
    }
    
    /**
     * Evaluate VLOOKUP function
     */
    private function evaluateVLookup($formula, $worksheetId) {
        // Parse VLOOKUP(lookup_value, table_array, col_index, [range_lookup])
        preg_match('/VLOOKUP\s*\(\s*([^,]+),\s*([^,]+),\s*([^,]+)(?:,\s*([^)]+))?\s*\)/i', $formula, $matches);
        
        if (!$matches) {
            return '#ERROR';
        }
        
        $lookupValue = $this->getCellValue($matches[1], $worksheetId);
        $tableRange = $matches[2];
        $colIndex = intval($matches[3]);
        $exactMatch = isset($matches[4]) && (strtoupper(trim($matches[4])) === 'FALSE' || trim($matches[4]) === '0');
        
        // Parse table range (e.g., 'Sheet1'!$A$1:$C$10)
        $rangeData = $this->parseRange($tableRange, $worksheetId);
        
        if (!$rangeData) {
            return '#REF!';
        }
        
        // Perform VLOOKUP
        $stmt = $this->db->prepare("
            SELECT c2.cell_value
            FROM cells c1
            JOIN cells c2 ON c1.worksheet_id = c2.worksheet_id 
                AND c1.row_num = c2.row_num
                AND c2.col_num = c1.col_num + :col_offset
            WHERE c1.worksheet_id = :worksheet_id
                AND c1.col_num = :first_col
                AND c1.row_num >= :start_row
                AND c1.row_num <= :end_row
                AND c1.cell_value = :lookup_value
            LIMIT 1
        ");
        
        $stmt->execute([
            ':worksheet_id' => $rangeData['worksheet_id'],
            ':first_col' => $rangeData['start_col'],
            ':start_row' => $rangeData['start_row'],
            ':end_row' => $rangeData['end_row'],
            ':col_offset' => $colIndex - 1,
            ':lookup_value' => $lookupValue
        ]);
        
        $result = $stmt->fetch(PDO::FETCH_ASSOC);
        
        return $result ? $result['cell_value'] : '#N/A';
    }
    
    /**
     * Parse cell address to row and column numbers
     */
    private function parseAddress($address) {
        preg_match('/([A-Z]+)(\d+)/i', $address, $matches);
        if (!$matches) return [0, 0];
        
        $col = 0;
        $colStr = strtoupper($matches[1]);
        for ($i = 0; $i < strlen($colStr); $i++) {
            $col = $col * 26 + (ord($colStr[$i]) - ord('A') + 1);
        }
        
        return [$col, intval($matches[2])];
    }
    
    /**
     * Parse range reference
     */
    private function parseRange($range, $defaultWorksheetId) {
        // Handle sheet references like 'Sheet1'!A1:C10
        if (strpos($range, '!') !== false) {
            list($sheetName, $range) = explode('!', $range);
            $sheetName = trim($sheetName, "'\"");
            
            // Get worksheet ID
            $stmt = $this->db->prepare("
                SELECT id FROM worksheets 
                WHERE sheet_name = :sheet_name
                LIMIT 1
            ");
            $stmt->execute([':sheet_name' => $sheetName]);
            $ws = $stmt->fetch(PDO::FETCH_ASSOC);
            $worksheetId = $ws ? $ws['id'] : $defaultWorksheetId;
        } else {
            $worksheetId = $defaultWorksheetId;
        }
        
        // Parse range A1:C10
        $range = str_replace('$', '', $range);
        list($start, $end) = explode(':', $range);
        list($startCol, $startRow) = $this->parseAddress($start);
        list($endCol, $endRow) = $this->parseAddress($end);
        
        return [
            'worksheet_id' => $worksheetId,
            'start_row' => $startRow,
            'end_row' => $endRow,
            'start_col' => $startCol,
            'end_col' => $endCol
        ];
    }
    
    /**
     * Get cell value
     */
    private function getCellValue($reference, $worksheetId) {
        $reference = trim($reference);
        
        // If it's a direct value (number or string in quotes)
        if (is_numeric($reference)) {
            return $reference;
        }
        if (preg_match('/^"([^"]*)"$/', $reference, $matches)) {
            return $matches[1];
        }
        
        // If it's a cell reference
        list($col, $row) = $this->parseAddress($reference);
        
        $stmt = $this->db->prepare("
            SELECT cell_value FROM cells
            WHERE worksheet_id = :worksheet_id
            AND row_num = :row AND col_num = :col
        ");
        $stmt->execute([
            ':worksheet_id' => $worksheetId,
            ':row' => $row,
            ':col' => $col
        ]);
        
        $result = $stmt->fetch(PDO::FETCH_ASSOC);
        return $result ? $result['cell_value'] : '';
    }
    
    /**
     * Get named ranges
     */
    private function getNamedRanges($workbookId) {
        $stmt = $this->db->prepare("
            SELECT name, cell_range, scope
            FROM named_ranges
            WHERE workbook_id = :workbook_id
        ");
        $stmt->execute([':workbook_id' => $workbookId]);
        $ranges = $stmt->fetchAll(PDO::FETCH_ASSOC);
        $this->sendResponse($ranges);
    }
    
    /**
     * Get data validation rules
     */
    private function getDataValidation($worksheetId) {
        $stmt = $this->db->prepare("
            SELECT * FROM data_validations
            WHERE worksheet_id = :worksheet_id
        ");
        $stmt->execute([':worksheet_id' => $worksheetId]);
        $validations = $stmt->fetchAll(PDO::FETCH_ASSOC);
        $this->sendResponse($validations);
    }
    
    /**
     * Send JSON response
     */
    private function sendResponse($data) {
        echo json_encode($data);
        exit;
    }
    
    /**
     * Send error response
     */
    private function sendError($message, $code = 400) {
        http_response_code($code);
        echo json_encode(['error' => $message]);
        exit;
    }
}

// Initialize and handle request
$api = new ExcelAPI();
$api->handleRequest();