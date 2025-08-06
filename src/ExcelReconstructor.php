<?php
require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use PhpOffice\PhpSpreadsheet\Style\Color;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Protection;
use PhpOffice\PhpSpreadsheet\Cell\DataValidation;

class ExcelReconstructor {
    private $db;
    private $spreadsheet;
    private $workbookId;
    
    public function __construct($host, $dbname, $username, $password) {
        try {
            $this->db = new PDO("mysql:host=$host;dbname=$dbname;charset=utf8mb4", 
                               $username, $password);
            $this->db->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION);
        } catch (PDOException $e) {
            throw new Exception("Connection failed: " . $e->getMessage());
        }
    }
    
    /**
     * Reconstruct Excel file from database
     */
    public function reconstructExcel($workbookId, $outputPath = null) {
        $this->workbookId = $workbookId;
        $this->spreadsheet = new Spreadsheet();
        
        try {
            // Load workbook metadata
            $this->loadWorkbookMetadata();
            
            // Remove default worksheet
            $this->spreadsheet->removeSheetByIndex(0);
            
            // Load all worksheets
            $this->loadWorksheets();
            
            // Load named ranges
            $this->loadNamedRanges();
            
            // Set active sheet to first one
            $this->spreadsheet->setActiveSheetIndex(0);
            
            // Save the file
            if ($outputPath) {
                $writer = new Xlsx($this->spreadsheet);
                $writer->save($outputPath);
                echo "Excel file reconstructed successfully at: $outputPath\n";
            }
            
            return $this->spreadsheet;
            
        } catch (Exception $e) {
            throw new Exception("Reconstruction failed: " . $e->getMessage());
        }
    }
    
    /**
     * Load workbook metadata
     */
    private function loadWorkbookMetadata() {
        $stmt = $this->db->prepare("SELECT * FROM workbooks WHERE id = :id");
        $stmt->execute([':id' => $this->workbookId]);
        $workbook = $stmt->fetch(PDO::FETCH_ASSOC);
        
        if (!$workbook) {
            throw new Exception("Workbook not found");
        }
        
        $metadata = json_decode($workbook['metadata'], true);
        $properties = $this->spreadsheet->getProperties();
        
        if ($metadata) {
            $properties->setCreator($metadata['creator'] ?? '');
            $properties->setLastModifiedBy($metadata['last_modified_by'] ?? '');
            $properties->setTitle($metadata['title'] ?? '');
            $properties->setSubject($metadata['subject'] ?? '');
            $properties->setDescription($metadata['description'] ?? '');
            $properties->setKeywords($metadata['keywords'] ?? '');
            $properties->setCategory($metadata['category'] ?? '');
            $properties->setCompany($metadata['company'] ?? '');
        }
    }
    
    /**
     * Load all worksheets
     */
    private function loadWorksheets() {
        $stmt = $this->db->prepare("
            SELECT * FROM worksheets 
            WHERE workbook_id = :workbook_id 
            ORDER BY sheet_index
        ");
        $stmt->execute([':workbook_id' => $this->workbookId]);
        $worksheets = $stmt->fetchAll(PDO::FETCH_ASSOC);
        
        foreach ($worksheets as $wsData) {
            $this->createWorksheet($wsData);
        }
    }
    
    /**
     * Create individual worksheet
     */
    private function createWorksheet($wsData) {
        $worksheet = new \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet($this->spreadsheet, $wsData['sheet_name']);
        $this->spreadsheet->addSheet($worksheet);
        
        // Set worksheet properties
        if ($wsData['sheet_state']) {
            $worksheet->setSheetState($wsData['sheet_state']);
        }
        
        if ($wsData['tab_color']) {
            $worksheet->getTabColor()->setARGB($wsData['tab_color']);
        }
        
        if ($wsData['zoom_scale']) {
            $worksheet->getSheetView()->setZoomScale($wsData['zoom_scale']);
        }
        
        $properties = json_decode($wsData['properties'], true);
        if ($properties) {
            $worksheet->setRightToLeft($properties['right_to_left'] ?? false);
            $worksheet->setShowGridlines($properties['show_gridlines'] ?? true);
            $worksheet->setShowRowColHeaders($properties['show_row_col_headers'] ?? true);
        }
        
        // Load cells
        $this->loadCells($worksheet, $wsData['id']);
        
        // Load merged cells
        $this->loadMergedCells($worksheet, $wsData['id']);
        
        // Load data validations
        $this->loadDataValidations($worksheet, $wsData['id']);
        
        // Load conditional formatting
        $this->loadConditionalFormatting($worksheet, $wsData['id']);
        
        // Load dimensions
        $this->loadDimensions($worksheet, $wsData['id']);
        
        // Load charts
        $this->loadCharts($worksheet, $wsData['id']);
        
        // Load images
        $this->loadImages($worksheet, $wsData['id']);
    }
    
    /**
     * Load cells for worksheet
     */
    private function loadCells($worksheet, $worksheetId) {
        $stmt = $this->db->prepare("
            SELECT c.*, cs.*
            FROM cells c
            LEFT JOIN cell_styles cs ON c.id = cs.cell_id
            WHERE c.worksheet_id = :worksheet_id
            ORDER BY c.row_num, c.col_num
        ");
        $stmt->execute([':worksheet_id' => $worksheetId]);
        
        while ($cellData = $stmt->fetch(PDO::FETCH_ASSOC)) {
            $cell = $worksheet->getCell($cellData['cell_address']);
            
            // Set cell value
            if ($cellData['formula']) {
                $cell->setValue($cellData['formula']);
            } else {
                switch ($cellData['cell_type']) {
                    case 'n': // Numeric
                        $cell->setValueExplicit($cellData['cell_value'], \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_NUMERIC);
                        break;
                    case 'b': // Boolean
                        $cell->setValueExplicit($cellData['cell_value'], \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_BOOL);
                        break;
                    case 's': // String
                    default:
                        $cell->setValueExplicit($cellData['cell_value'], \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_STRING);
                        break;
                }
            }
            
            // Set hyperlink
            if ($cellData['hyperlink']) {
                $cell->getHyperlink()->setUrl($cellData['hyperlink']);
            }
            
            // Set comment
            if ($cellData['comment']) {
                $worksheet->getComment($cellData['cell_address'])->getText()->createText($cellData['comment']);
            }
            
            // Apply cell style
            $this->applyCellStyle($cell, $cellData);
        }
    }
    
    /**
     * Apply cell styling
     */
    private function applyCellStyle($cell, $styleData) {
        $style = $cell->getStyle();
        
        // Font styling
        if ($styleData['font_name']) {
            $font = $style->getFont();
            $font->setName($styleData['font_name']);
            $font->setSize($styleData['font_size']);
            $font->setBold($styleData['font_bold']);
            $font->setItalic($styleData['font_italic']);
            $font->setUnderline($styleData['font_underline'] ?: false);
            $font->setStrikethrough($styleData['font_strike']);
            
            if ($styleData['font_color']) {
                $font->getColor()->setARGB($styleData['font_color']);
            }
        }
        
        // Fill styling
        if ($styleData['fill_color'] && $styleData['fill_color'] !== 'FF000000') {
            $fill = $style->getFill();
            $fill->setFillType($styleData['fill_type'] ?: Fill::FILL_SOLID);
            $fill->getStartColor()->setARGB($styleData['fill_color']);
        }
        
        // Border styling
        $borders = $style->getBorders();
        
        if ($styleData['border_top']) {
            $borders->getTop()->setBorderStyle($styleData['border_top']);
            if ($styleData['border_color_top']) {
                $borders->getTop()->getColor()->setARGB($styleData['border_color_top']);
            }
        }
        
        if ($styleData['border_right']) {
            $borders->getRight()->setBorderStyle($styleData['border_right']);
            if ($styleData['border_color_right']) {
                $borders->getRight()->getColor()->setARGB($styleData['border_color_right']);
            }
        }
        
        if ($styleData['border_bottom']) {
            $borders->getBottom()->setBorderStyle($styleData['border_bottom']);
            if ($styleData['border_color_bottom']) {
                $borders->getBottom()->getColor()->setARGB($styleData['border_color_bottom']);
            }
        }
        
        if ($styleData['border_left']) {
            $borders->getLeft()->setBorderStyle($styleData['border_left']);
            if ($styleData['border_color_left']) {
                $borders->getLeft()->getColor()->setARGB($styleData['border_color_left']);
            }
        }
        
        // Alignment
        $alignment = $style->getAlignment();
        if ($styleData['alignment_horizontal']) {
            $alignment->setHorizontal($styleData['alignment_horizontal']);
        }
        if ($styleData['alignment_vertical']) {
            $alignment->setVertical($styleData['alignment_vertical']);
        }
        $alignment->setWrapText($styleData['alignment_wrap_text']);
        if ($styleData['alignment_text_rotation']) {
            $alignment->setTextRotation($styleData['alignment_text_rotation']);
        }
        if ($styleData['alignment_indent']) {
            $alignment->setIndent($styleData['alignment_indent']);
        }
        
        // Number format
        if ($styleData['number_format']) {
            $style->getNumberFormat()->setFormatCode($styleData['number_format']);
        }
        
        // Protection
        $protection = $style->getProtection();
        $protection->setLocked($styleData['protection_locked']);
        $protection->setHidden($styleData['protection_hidden']);
    }
    
    /**
     * Load merged cells
     */
    private function loadMergedCells($worksheet, $worksheetId) {
        $stmt = $this->db->prepare("
            SELECT * FROM merged_cells 
            WHERE worksheet_id = :worksheet_id
        ");
        $stmt->execute([':worksheet_id' => $worksheetId]);
        
        while ($merge = $stmt->fetch(PDO::FETCH_ASSOC)) {
            $worksheet->mergeCells($merge['cell_range']);
        }
    }
    
    /**
     * Load data validations
     */
    private function loadDataValidations($worksheet, $worksheetId) {
        $stmt = $this->db->prepare("
            SELECT * FROM data_validations 
            WHERE worksheet_id = :worksheet_id
        ");
        $stmt->execute([':worksheet_id' => $worksheetId]);
        
        while ($valData = $stmt->fetch(PDO::FETCH_ASSOC)) {
            $validation = $worksheet->getCell($valData['cell_range'])->getDataValidation();
            
            $validation->setType($valData['validation_type']);
            $validation->setOperator($valData['operator']);
            $validation->setFormula1($valData['formula1']);
            $validation->setFormula2($valData['formula2']);
            $validation->setAllowBlank($valData['allow_blank']);
            $validation->setShowDropDown($valData['show_dropdown']);
            $validation->setShowInputMessage($valData['show_input_message']);
            $validation->setShowErrorMessage($valData['show_error_message']);
            $validation->setErrorTitle($valData['error_title']);
            $validation->setError($valData['error_message']);
            $validation->setPromptTitle($valData['input_title']);
            $validation->setPrompt($valData['input_message']);
            $validation->setErrorStyle($valData['error_style']);
        }
    }
    
    /**
     * Load conditional formatting
     */
    private function loadConditionalFormatting($worksheet, $worksheetId) {
        $stmt = $this->db->prepare("
            SELECT * FROM conditional_formats 
            WHERE worksheet_id = :worksheet_id
            ORDER BY priority
        ");
        $stmt->execute([':worksheet_id' => $worksheetId]);
        
        while ($cfData = $stmt->fetch(PDO::FETCH_ASSOC)) {
            $conditional = new \PhpOffice\PhpSpreadsheet\Style\Conditional();
            
            $conditional->setConditionType($cfData['rule_type']);
            $conditional->setOperatorType($cfData['operator']);
            
            if ($cfData['formula']) {
                $conditional->addCondition($cfData['formula']);
            }
            
            if ($cfData['text']) {
                $conditional->setText($cfData['text']);
            }
            
            $conditional->setStopIfTrue($cfData['stop_if_true']);
            
            // Apply style attributes
            if ($cfData['style_attributes']) {
                $styleAttrs = json_decode($cfData['style_attributes'], true);
                $conditionalStyle = $conditional->getStyle();
                
                if (isset($styleAttrs['font'])) {
                    $conditionalStyle->getFont()->setBold($styleAttrs['font']['bold'] ?? false);
                    $conditionalStyle->getFont()->setItalic($styleAttrs['font']['italic'] ?? false);
                    if (isset($styleAttrs['font']['color'])) {
                        $conditionalStyle->getFont()->getColor()->setARGB($styleAttrs['font']['color']);
                    }
                }
                
                if (isset($styleAttrs['fill'])) {
                    $conditionalStyle->getFill()->setFillType($styleAttrs['fill']['type'] ?? Fill::FILL_SOLID);
                    if (isset($styleAttrs['fill']['color'])) {
                        $conditionalStyle->getFill()->getStartColor()->setARGB($styleAttrs['fill']['color']);
                    }
                }
            }
            
            $existingConditionals = $worksheet->getConditionalStyles($cfData['cell_range']);
            $existingConditionals[] = $conditional;
            $worksheet->setConditionalStyles($cfData['cell_range'], $existingConditionals);
        }
    }
    
    /**
     * Load dimensions
     */
    private function loadDimensions($worksheet, $worksheetId) {
        $stmt = $this->db->prepare("
            SELECT * FROM dimensions 
            WHERE worksheet_id = :worksheet_id
        ");
        $stmt->execute([':worksheet_id' => $worksheetId]);
        
        while ($dim = $stmt->fetch(PDO::FETCH_ASSOC)) {
            if ($dim['dimension_type'] === 'row') {
                $rowDimension = $worksheet->getRowDimension($dim['dimension_index']);
                $rowDimension->setRowHeight($dim['height_or_width']);
                $rowDimension->setVisible(!$dim['hidden']);
                $rowDimension->setOutlineLevel($dim['outline_level']);
                $rowDimension->setCollapsed($dim['collapsed']);
            } else {
                $colLetter = Coordinate::stringFromColumnIndex($dim['dimension_index']);
                $colDimension = $worksheet->getColumnDimension($colLetter);
                $colDimension->setWidth($dim['height_or_width']);
                $colDimension->setVisible(!$dim['hidden']);
                $colDimension->setOutlineLevel($dim['outline_level']);
                $colDimension->setCollapsed($dim['collapsed']);
            }
        }
    }
    
    /**
     * Load charts
     */
    private function loadCharts($worksheet, $worksheetId) {
        // Chart reconstruction is complex and depends on specific chart types
        // This is a placeholder for chart loading logic
        $stmt = $this->db->prepare("
            SELECT * FROM charts 
            WHERE worksheet_id = :worksheet_id
        ");
        $stmt->execute([':worksheet_id' => $worksheetId]);
        
        // Chart reconstruction would go here
        // PhpSpreadsheet has limited chart support compared to Excel
    }
    
    /**
     * Load images
     */
    private function loadImages($worksheet, $worksheetId) {
        $stmt = $this->db->prepare("
            SELECT * FROM embedded_objects 
            WHERE worksheet_id = :worksheet_id AND object_type = 'image'
        ");
        $stmt->execute([':worksheet_id' => $worksheetId]);
        
        while ($img = $stmt->fetch(PDO::FETCH_ASSOC)) {
            if ($img['object_data']) {
                // Create temporary file for image
                $tempFile = tempnam(sys_get_temp_dir(), 'img');
                file_put_contents($tempFile, $img['object_data']);
                
                $drawing = new \PhpOffice\PhpSpreadsheet\Worksheet\Drawing();
                $drawing->setName($img['object_name']);
                $drawing->setPath($tempFile);
                $drawing->setCoordinates($img['position_from']);
                
                if ($img['width']) {
                    $drawing->setWidth($img['width']);
                }
                if ($img['height']) {
                    $drawing->setHeight($img['height']);
                }
                
                $properties = json_decode($img['properties'], true);
                if ($properties) {
                    $drawing->setRotation($properties['rotation'] ?? 0);
                    $drawing->setResizeProportional($properties['resizeProportional'] ?? true);
                }
                
                $drawing->setWorksheet($worksheet);
            }
        }
    }
    
    /**
     * Load named ranges
     */
    private function loadNamedRanges() {
        $stmt = $this->db->prepare("
            SELECT * FROM named_ranges 
            WHERE workbook_id = :workbook_id
        ");
        $stmt->execute([':workbook_id' => $this->workbookId]);
        
        while ($nr = $stmt->fetch(PDO::FETCH_ASSOC)) {
            $namedRange = new \PhpOffice\PhpSpreadsheet\NamedRange(
                $nr['name'],
                $this->spreadsheet->getSheetByName($nr['scope']) ?: $this->spreadsheet->getActiveSheet(),
                $nr['cell_range']
            );
            $this->spreadsheet->addNamedRange($namedRange);
        }
    }
    
    /**
     * Get cell data as JSON for API/Dashboard use
     */
    public function getCellDataAsJson($worksheetId, $cellRange = null) {
        $query = "
            SELECT c.*, cs.*
            FROM cells c
            LEFT JOIN cell_styles cs ON c.id = cs.cell_id
            WHERE c.worksheet_id = :worksheet_id
        ";
        
        $params = [':worksheet_id' => $worksheetId];
        
        if ($cellRange) {
            // Parse cell range and add to query
            // Implementation depends on your needs
        }
        
        $stmt = $this->db->prepare($query);
        $stmt->execute($params);
        
        $data = [];
        while ($row = $stmt->fetch(PDO::FETCH_ASSOC)) {
            $data[$row['cell_address']] = [
                'value' => $row['cell_value'],
                'formula' => $row['formula'],
                'formatted' => $row['formatted_value'],
                'style' => [
                    'font' => [
                        'name' => $row['font_name'],
                        'size' => $row['font_size'],
                        'bold' => $row['font_bold'],
                        'color' => $row['font_color']
                    ],
                    'fill' => [
                        'color' => $row['fill_color']
                    ],
                    'alignment' => [
                        'horizontal' => $row['alignment_horizontal'],
                        'vertical' => $row['alignment_vertical']
                    ]
                ]
            ];
        }
        
        return json_encode($data);
    }
}

// Usage example
try {
    $reconstructor = new ExcelReconstructor('localhost', 'your_database', 'username', 'password');
    
    // Reconstruct the Excel file
    $reconstructor->reconstructExcel(1, 'output/reconstructed.xlsx');
    
    // Or get data for web dashboard
    $jsonData = $reconstructor->getCellDataAsJson(1);
    
} catch (Exception $e) {
    echo "Error: " . $e->getMessage() . "\n";
}