<?php
require_once __DIR__ . '/../vendor/autoload.php';
require_once __DIR__ . '/../src/ExcelReconstructor.php';

$config = require __DIR__ . '/../config/database.php';

// Check command line arguments
if ($argc < 2) {
    echo "Usage: php reconstruct.php <workbook-id> [output-filename]\n";
    echo "Example: php reconstruct.php 1 reconstructed.xlsx\n";
    exit(1);
}

$workbookId = $argv[1];
$outputFilename = $argv[2] ?? 'reconstructed_' . date('Y-m-d_H-i-s') . '.xlsx';
$outputPath = __DIR__ . '/../output/' . $outputFilename;

try {
    echo "Starting reconstruction of workbook ID: $workbookId\n";
    echo "----------------------------------------\n";
    
    // Create reconstructor instance
    $reconstructor = new ExcelReconstructor(
        $config['host'],
        $config['database'],
        $config['username'],
        $config['password']
    );
    
    // Reconstruct the Excel file
    $startTime = microtime(true);
    $reconstructor->reconstructExcel($workbookId, $outputPath);
    $endTime = microtime(true);
    
    $executionTime = round($endTime - $startTime, 2);
    echo "----------------------------------------\n";
    echo "Reconstruction completed in {$executionTime} seconds\n";
    echo "Output file: $outputPath\n";
    
} catch (Exception $e) {
    echo "Error: " . $e->getMessage() . "\n";
    exit(1);
}