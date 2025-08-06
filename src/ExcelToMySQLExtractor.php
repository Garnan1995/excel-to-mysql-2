<?php
require __DIR__ . '/../vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use PhpOffice\PhpSpreadsheet\Style\NumberFormat;
use PhpOffice\PhpSpreadsheet\RichText\RichText;

class ExcelToMySQLExtractor {
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
     * Main extraction method
     */
    public function extractExcelFile($filePath) {
        try {
            // Load the spreadsheet
            $this->spreadsheet = IOFactory::load($filePath);
            
            // Start transaction for data integrity
            $this->db->beginTransaction();
            
            // Save workbook information
            $this->saveWorkbook($filePath);
            
            // Process each worksheet
            foreach ($this->spreadsheet->getWorksheetIterator() as $worksheet) {
                $this->processWorksheet($worksheet);
            }
            
            // Save named ranges
            $this->saveNamedRanges();
            
            // Commit transaction
            $this->db->commit();
            
            echo "Excel file successfully extracted to MySQL!\n";
            echo "Workbook ID: " . $this->workbookId . "\n";
            
        } catch (Exception $e) {
            $this->db->rollBack();
            throw new Exception("Extraction failed: " . $e->getMessage());
        }
    }
    
    /**
     * Save workbook metadata
     */
    private function saveWorkbook($filePath) {
        $properties = $this->spreadsheet->getProperties();
        
        // Handle dates properly - they might be DateTime objects or integers
        $createdDate = $properties->getCreated();
        $modifiedDate = $properties->getModified();
        
        $metadata = [
            'creator' => $properties->getCreator(),
            'last_modified_by' => $properties->getLastModifiedBy(),
            'title' => $properties->getTitle(),
            'subject' => $properties->getSubject(),
            'description' => $properties->getDescription(),
            'keywords' => $properties->getKeywords(),
            'category' => $properties->getCategory(),
            'company' => $properties->getCompany(),
            'created' => null,
            'modified' => null
        ];
        
        // Handle created date
        if ($createdDate instanceof \DateTime) {
            $metadata['created'] = $createdDate->format('Y-m-d H:i:s');
        } elseif (is_numeric($createdDate) && $createdDate > 0) {
            $metadata['created'] = date('Y-m-d H:i:s', $createdDate);
        }
        
        // Handle modified date
        if ($modifiedDate instanceof \DateTime) {
            $metadata['modified'] = $modifiedDate->format('Y-m-d H:i:s');
        } elseif (is_numeric($modifiedDate) && $modifiedDate > 0) {
            $metadata['modified'] = date('Y-m-d H:i:s', $modifiedDate);
        }
        
        $stmt = $this->db->prepare("
            INSERT INTO workbooks (filename, file_hash, metadata) 
            VALUES (:filename, :file_hash, :metadata)
        ");
        
        $stmt->execute([
            ':filename' => basename($filePath),
            ':file_hash' => md5_file($filePath),
            ':metadata' => json_encode($metadata)
        ]);
        
        $this->workbookId = $this->db->lastInsertId();
    }
    
    /**
     * Process individual worksheet
     */
    private function processWorksheet(Worksheet $worksheet) {
        echo "Processing worksheet: " . $worksheet->getTitle() . "\n";
        
        // Save worksheet information
        $worksheetId = $this->saveWorksheetInfo($worksheet);
        
        // Process cells
        $this->processCells($worksheet, $worksheetId);
        
        // Save data validations
        $this->saveDataValidations($worksheet, $worksheetId);
        
        // Save conditional formatting
        $this->saveConditionalFormatting($worksheet, $worksheetId);
        
        // Save merged cells
        $this->saveMergedCells($worksheet, $worksheetId);
        
        // Save row and column dimensions
        $this->saveDimensions($worksheet, $worksheetId);
        
        // Save charts
        $this->saveCharts($worksheet, $worksheetId);
        
        // Save images
        $this->saveImages($worksheet, $worksheetId);
    }
    
    /**
     * Save worksheet metadata
     */
    private function saveWorksheetInfo($worksheet) {
        $sheetState = $worksheet->getSheetState();
        $tabColor = null;
        
        if ($worksheet->getTabColor()) {
            $tabColor = $worksheet->getTabColor()->getARGB();
        }
        
        $properties = [
            'code_name' => $worksheet->getCodeName(),
            'right_to_left' => $worksheet->getRightToLeft(),
            'show_gridlines' => $worksheet->getShowGridlines(),
            'show_row_col_headers' => $worksheet->getShowRowColHeaders(),
            'show_summary_below' => $worksheet->getShowSummaryBelow(),
            'show_summary_right' => $worksheet->getShowSummaryRight()
        ];
        
        $stmt = $this->db->prepare("
            INSERT INTO worksheets (workbook_id, sheet_name, sheet_index, sheet_state, 
                                  tab_color, zoom_scale, properties)
            VALUES (:workbook_id, :sheet_name, :sheet_index, :sheet_state, 
                    :tab_color, :zoom_scale, :properties)
        ");
        
        $stmt->execute([
            ':workbook_id' => $this->workbookId,
            ':sheet_name' => $worksheet->getTitle(),
            ':sheet_index' => $this->spreadsheet->getIndex($worksheet),
            ':sheet_state' => $sheetState,
            ':tab_color' => $tabColor,
            ':zoom_scale' => $worksheet->getSheetView()->getZoomScale(),
            ':properties' => json_encode($properties)
        ]);
        
        return $this->db->lastInsertId();
    }
    
    /**
     * Process and save all cells
     */
    private function processCells($worksheet, $worksheetId) {
        $highestRow = $worksheet->getHighestRow();
        $highestColumn = $worksheet->getHighestColumn();
        $highestColumnIndex = Coordinate::columnIndexFromString($highestColumn);
        
        echo "  Processing cells (rows: $highestRow, columns: $highestColumnIndex)...\n";
        
        // Prepare batch insert for better performance
        $cellStmt = $this->db->prepare("
            INSERT INTO cells (worksheet_id, cell_address, row_num, col_num, 
                             cell_value, cell_type, formatted_value, formula, 
                             formula_attributes, hyperlink, comment)
            VALUES (:worksheet_id, :cell_address, :row_num, :col_num, 
                    :cell_value, :cell_type, :formatted_value, :formula, 
                    :formula_attributes, :hyperlink, :comment)
        ");
        
        $cellCount = 0;
        for ($row = 1; $row <= $highestRow; $row++) {
            for ($col = 1; $col <= $highestColumnIndex; $col++) {
                $cellAddress = Coordinate::stringFromColumnIndex($col) . $row;
                $cell = $worksheet->getCell($cellAddress);
                
                // Skip completely empty cells (no value and no style)
                if ($cell->getValue() === null && 
                    $cell->getStyle()->getFont()->getName() === 'Calibri' &&
                    $cell->getStyle()->getFont()->getSize() === 11) {
                    continue;
                }
                
                // Extract cell data
                $cellData = $this->extractCellData($cell, $worksheet);
                
                // Save cell
                $cellStmt->execute([
                    ':worksheet_id' => $worksheetId,
                    ':cell_address' => $cellAddress,
                    ':row_num' => $row,
                    ':col_num' => $col,
                    ':cell_value' => $cellData['value'],
                    ':cell_type' => $cellData['type'],
                    ':formatted_value' => $cellData['formatted_value'],
                    ':formula' => $cellData['formula'],
                    ':formula_attributes' => $cellData['formula_attributes'] ? 
                                           json_encode($cellData['formula_attributes']) : null,
                    ':hyperlink' => $cellData['hyperlink'],
                    ':comment' => $cellData['comment']
                ]);
                
                $cellId = $this->db->lastInsertId();
                
                // Save cell style
                $this->saveCellStyle($cell, $cellId);
                $cellCount++;
            }
            
            // Show progress every 100 rows
            if ($row % 100 == 0) {
                echo "    Processed $row rows...\n";
            }
        }
        
        echo "  Total cells saved: $cellCount\n";
    }
    
    /**
     * Extract comprehensive cell data
     */
    private function extractCellData($cell, $worksheet) {
        $data = [
            'value' => null,
            'type' => null,
            'formatted_value' => null,
            'formula' => null,
            'formula_attributes' => null,
            'hyperlink' => null,
            'comment' => null
        ];
        
        // Get cell value
        $value = $cell->getValue();
        if ($value instanceof RichText) {
            $data['value'] = $value->getPlainText();
            $data['type'] = 'richtext';
        } else {
            $data['value'] = $value;
            $data['type'] = $cell->getDataType();
        }
        
        // Get formatted value
        try {
            $data['formatted_value'] = $cell->getFormattedValue();
        } catch (Exception $e) {
            $data['formatted_value'] = $data['value'];
        }
        
        // Check for hyperlink
        if ($cell->hasHyperlink()) {
            $data['hyperlink'] = $cell->getHyperlink()->getUrl();
        }
        
        // Check for formula
        if ($cell->isFormula()) {
            $data['formula'] = $cell->getValue();
            try {
                $data['formula_attributes'] = [
                    'calculated_value' => $cell->getCalculatedValue(),
                    'formula_attributes' => $cell->getFormulaAttributes()
                ];
            } catch (Exception $e) {
                $data['formula_attributes'] = ['error' => $e->getMessage()];
            }
        }
        
        // Check for comment
        if ($worksheet->getComment($cell->getCoordinate())) {
            $comment = $worksheet->getComment($cell->getCoordinate());
            $data['comment'] = $comment->getText()->getPlainText();
        }
        
        return $data;
    }
    
    /**
     * Save cell styling information
     */
    private function saveCellStyle($cell, $cellId) {
        $style = $cell->getStyle();
        
        // Font information
        $font = $style->getFont();
        $fontColor = $font->getColor()->getARGB();
        
        // Fill information
        $fill = $style->getFill();
        $fillColor = $fill->getStartColor()->getARGB();
        
        // Border information
        $borders = $style->getBorders();
        
        // Alignment information
        $alignment = $style->getAlignment();
        
        // Number format
        $numberFormat = $style->getNumberFormat()->getFormatCode();
        
        // Protection
        $protection = $style->getProtection();
        
        $stmt = $this->db->prepare("
            INSERT INTO cell_styles (
                cell_id, font_name, font_size, font_bold, font_italic, 
                font_underline, font_strike, font_color,
                fill_type, fill_color, fill_pattern,
                border_top, border_right, border_bottom, border_left,
                border_color_top, border_color_right, border_color_bottom, border_color_left,
                alignment_horizontal, alignment_vertical, alignment_wrap_text, 
                alignment_text_rotation, alignment_indent,
                number_format, protection_locked, protection_hidden
            ) VALUES (
                :cell_id, :font_name, :font_size, :font_bold, :font_italic,
                :font_underline, :font_strike, :font_color,
                :fill_type, :fill_color, :fill_pattern,
                :border_top, :border_right, :border_bottom, :border_left,
                :border_color_top, :border_color_right, :border_color_bottom, :border_color_left,
                :alignment_horizontal, :alignment_vertical, :alignment_wrap_text,
                :alignment_text_rotation, :alignment_indent,
                :number_format, :protection_locked, :protection_hidden
            )
        ");
        
        $stmt->execute([
            ':cell_id' => $cellId,
            ':font_name' => $font->getName(),
            ':font_size' => $font->getSize(),
            ':font_bold' => $font->getBold(),
            ':font_italic' => $font->getItalic(),
            ':font_underline' => $font->getUnderline(),
            ':font_strike' => $font->getStrikethrough(),
            ':font_color' => $fontColor,
            ':fill_type' => $fill->getFillType(),
            ':fill_color' => $fillColor,
            ':fill_pattern' => null,
            ':border_top' => $borders->getTop()->getBorderStyle(),
            ':border_right' => $borders->getRight()->getBorderStyle(),
            ':border_bottom' => $borders->getBottom()->getBorderStyle(),
            ':border_left' => $borders->getLeft()->getBorderStyle(),
            ':border_color_top' => $borders->getTop()->getColor()->getARGB(),
            ':border_color_right' => $borders->getRight()->getColor()->getARGB(),
            ':border_color_bottom' => $borders->getBottom()->getColor()->getARGB(),
            ':border_color_left' => $borders->getLeft()->getColor()->getARGB(),
            ':alignment_horizontal' => $alignment->getHorizontal(),
            ':alignment_vertical' => $alignment->getVertical(),
            ':alignment_wrap_text' => $alignment->getWrapText(),
            ':alignment_text_rotation' => $alignment->getTextRotation(),
            ':alignment_indent' => $alignment->getIndent(),
            ':number_format' => $numberFormat,
            ':protection_locked' => $protection->getLocked(),
            ':protection_hidden' => $protection->getHidden()
        ]);
    }
    
    /**
     * Save data validation rules
     */
    private function saveDataValidations($worksheet, $worksheetId) {
        $validations = $worksheet->getDataValidationCollection();
        
        foreach ($validations as $address => $validation) {
            $stmt = $this->db->prepare("
                INSERT INTO data_validations (
                    worksheet_id, cell_range, validation_type, operator,
                    formula1, formula2, allow_blank, show_dropdown,
                    show_input_message, show_error_message,
                    error_title, error_message, input_title, input_message,
                    error_style
                ) VALUES (
                    :worksheet_id, :cell_range, :validation_type, :operator,
                    :formula1, :formula2, :allow_blank, :show_dropdown,
                    :show_input_message, :show_error_message,
                    :error_title, :error_message, :input_title, :input_message,
                    :error_style
                )
            ");
            
            $stmt->execute([
                ':worksheet_id' => $worksheetId,
                ':cell_range' => $address,
                ':validation_type' => $validation->getType(),
                ':operator' => $validation->getOperator(),
                ':formula1' => $validation->getFormula1(),
                ':formula2' => $validation->getFormula2(),
                ':allow_blank' => $validation->getAllowBlank(),
                ':show_dropdown' => $validation->getShowDropDown(),
                ':show_input_message' => $validation->getShowInputMessage(),
                ':show_error_message' => $validation->getShowErrorMessage(),
                ':error_title' => $validation->getErrorTitle(),
                ':error_message' => $validation->getError(),
                ':input_title' => $validation->getPromptTitle(),
                ':input_message' => $validation->getPrompt(),
                ':error_style' => $validation->getErrorStyle()
            ]);
        }
    }
    
    /**
     * Save conditional formatting rules
     */
    private function saveConditionalFormatting($worksheet, $worksheetId) {
        $conditionals = $worksheet->getConditionalStylesCollection();
        
        foreach ($conditionals as $coordinate => $conditionalStyles) {
            foreach ($conditionalStyles as $conditional) {
                $styleAttributes = [
                    'font' => [
                        'bold' => $conditional->getStyle()->getFont()->getBold(),
                        'italic' => $conditional->getStyle()->getFont()->getItalic(),
                        'color' => $conditional->getStyle()->getFont()->getColor()->getARGB()
                    ],
                    'fill' => [
                        'type' => $conditional->getStyle()->getFill()->getFillType(),
                        'color' => $conditional->getStyle()->getFill()->getStartColor()->getARGB()
                    ]
                ];
                
                $stmt = $this->db->prepare("
                    INSERT INTO conditional_formats (
                        worksheet_id, cell_range, rule_type, priority,
                        stop_if_true, formula, operator, text, style_attributes
                    ) VALUES (
                        :worksheet_id, :cell_range, :rule_type, :priority,
                        :stop_if_true, :formula, :operator, :text, :style_attributes
                    )
                ");
                
                $conditions = $conditional->getConditions();
                $formula = count($conditions) > 0 ? $conditions[0] : null;
                
                $stmt->execute([
                    ':worksheet_id' => $worksheetId,
                    ':cell_range' => $coordinate,
                    ':rule_type' => $conditional->getConditionType(),
                    ':priority' => $conditional->getPriority(),
                    ':stop_if_true' => $conditional->getStopIfTrue(),
                    ':formula' => $formula,
                    ':operator' => $conditional->getOperatorType(),
                    ':text' => $conditional->getText(),
                    ':style_attributes' => json_encode($styleAttributes)
                ]);
            }
        }
    }
    
    /**
     * Save merged cells
     */
    private function saveMergedCells($worksheet, $worksheetId) {
        $mergedCells = $worksheet->getMergeCells();
        
        foreach ($mergedCells as $mergeRange) {
            $range = Coordinate::splitRange($mergeRange);
            $topLeft = $range[0][0];
            $bottomRight = $range[0][1];
            
            $stmt = $this->db->prepare("
                INSERT INTO merged_cells (worksheet_id, cell_range, top_left_cell, bottom_right_cell)
                VALUES (:worksheet_id, :cell_range, :top_left_cell, :bottom_right_cell)
            ");
            
            $stmt->execute([
                ':worksheet_id' => $worksheetId,
                ':cell_range' => $mergeRange,
                ':top_left_cell' => $topLeft,
                ':bottom_right_cell' => $bottomRight
            ]);
        }
    }
    
    /**
     * Save row and column dimensions
     */
    private function saveDimensions($worksheet, $worksheetId) {
        $stmt = $this->db->prepare("
            INSERT INTO dimensions (worksheet_id, dimension_type, dimension_index, 
                                  height_or_width, hidden, outline_level, collapsed)
            VALUES (:worksheet_id, :dimension_type, :dimension_index, 
                    :height_or_width, :hidden, :outline_level, :collapsed)
        ");
        
        // Save row dimensions
        foreach ($worksheet->getRowDimensions() as $rowIndex => $dimension) {
            $stmt->execute([
                ':worksheet_id' => $worksheetId,
                ':dimension_type' => 'row',
                ':dimension_index' => $rowIndex,
                ':height_or_width' => $dimension->getRowHeight(),
                ':hidden' => $dimension->getVisible() ? 0 : 1,
                ':outline_level' => $dimension->getOutlineLevel(),
                ':collapsed' => $dimension->getCollapsed()
            ]);
        }
        
        // Save column dimensions
        foreach ($worksheet->getColumnDimensions() as $columnIndex => $dimension) {
            $stmt->execute([
                ':worksheet_id' => $worksheetId,
                ':dimension_type' => 'column',
                ':dimension_index' => Coordinate::columnIndexFromString($columnIndex),
                ':height_or_width' => $dimension->getWidth(),
                ':hidden' => $dimension->getVisible() ? 0 : 1,
                ':outline_level' => $dimension->getOutlineLevel(),
                ':collapsed' => $dimension->getCollapsed()
            ]);
        }
    }
    
    /**
     * Save charts
     */
    private function saveCharts($worksheet, $worksheetId) {
        $chartCollection = $worksheet->getChartCollection();
        
        foreach ($chartCollection as $chart) {
            $chartData = [
                'title' => $chart->getTitle() ? $chart->getTitle()->getCaption() : null,
                'plot_area' => [],
                'legend' => [
                    'position' => $chart->getLegend() ? $chart->getLegend()->getPosition() : null,
                    'overlay' => $chart->getLegend() ? $chart->getLegend()->getOverlay() : null
                ],
                'plot_series' => []
            ];
            
            $stmt = $this->db->prepare("
                INSERT INTO charts (worksheet_id, chart_name, chart_type, 
                                  position_from, position_to, chart_data)
                VALUES (:worksheet_id, :chart_name, :chart_type, 
                        :position_from, :position_to, :chart_data)
            ");
            
            $topLeft = $chart->getTopLeftCell();
            $bottomRight = $chart->getBottomRightCell();
            
            $stmt->execute([
                ':worksheet_id' => $worksheetId,
                ':chart_name' => $chart->getName(),
                ':chart_type' => 'chart',
                ':position_from' => $topLeft,
                ':position_to' => $bottomRight,
                ':chart_data' => json_encode($chartData)
            ]);
        }
    }
    
    /**
     * Save images and drawings
     */
    private function saveImages($worksheet, $worksheetId) {
        $drawingCollection = $worksheet->getDrawingCollection();
        
        foreach ($drawingCollection as $drawing) {
            if ($drawing instanceof \PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing) {
                ob_start();
                switch ($drawing->getMimeType()) {
                    case \PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing::MIMETYPE_PNG:
                        imagepng($drawing->getImageResource());
                        break;
                    case \PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing::MIMETYPE_GIF:
                        imagegif($drawing->getImageResource());
                        break;
                    case \PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing::MIMETYPE_JPEG:
                        imagejpeg($drawing->getImageResource());
                        break;
                }
                $imageContents = ob_get_contents();
                ob_end_clean();
            } elseif ($drawing instanceof \PhpOffice\PhpSpreadsheet\Worksheet\Drawing) {
                $imageContents = file_get_contents($drawing->getPath());
            } else {
                continue;
            }
            
            $properties = [
                'name' => $drawing->getName(),
                'description' => $drawing->getDescription(),
                'resizeProportional' => $drawing->getResizeProportional(),
                'rotation' => $drawing->getRotation(),
                'shadow' => $drawing->getShadow() ? [
                    'visible' => $drawing->getShadow()->getVisible(),
                    'direction' => $drawing->getShadow()->getDirection(),
                    'distance' => $drawing->getShadow()->getDistance()
                ] : null
            ];
            
            $stmt = $this->db->prepare("
                INSERT INTO embedded_objects (worksheet_id, object_type, object_name,
                                             position_from, width, height, 
                                             object_data, properties)
                VALUES (:worksheet_id, :object_type, :object_name,
                        :position_from, :width, :height, 
                        :object_data, :properties)
            ");
            
            $stmt->execute([
                ':worksheet_id' => $worksheetId,
                ':object_type' => 'image',
                ':object_name' => $drawing->getName(),
                ':position_from' => $drawing->getCoordinates(),
                ':width' => $drawing->getWidth(),
                ':height' => $drawing->getHeight(),
                ':object_data' => $imageContents,
                ':properties' => json_encode($properties)
            ]);
        }
    }
    
    /**
     * Save named ranges
     */
    private function saveNamedRanges() {
        $namedRanges = $this->spreadsheet->getNamedRanges();
        
        foreach ($namedRanges as $namedRange) {
            $stmt = $this->db->prepare("
                INSERT INTO named_ranges (workbook_id, name, cell_range, scope, comment)
                VALUES (:workbook_id, :name, :cell_range, :scope, :comment)
            ");
            
            $stmt->execute([
                ':workbook_id' => $this->workbookId,
                ':name' => $namedRange->getName(),
                ':cell_range' => $namedRange->getRange(),
                ':scope' => $namedRange->getScope() ? $namedRange->getScope()->getTitle() : 'Workbook',
                ':comment' => $namedRange->getComment()
            ]);
        }
    }
}