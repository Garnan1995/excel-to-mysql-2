<?php
require_once __DIR__ . '/../vendor/autoload.php';
require_once __DIR__ . '/../src/ExcelToMySQLExtractor.php';

$config = require __DIR__ . '/../config/database.php';

// Check command line arguments
if ($argc < 2) {
    echo "Usage: php extract.php <excel-file-path>\n";
    echo "Example: php extract.php ../input/data.xlsx\n";
    exit(1);
}

$filePath = $argv[1];

// Check if file exists
if (!file_exists($filePath)) {
    // Try checking in input directory
    $altPath = __DIR__ . '/../input/' . basename($filePath);
    if (file_exists($altPath)) {
        $filePath = $altPath;
    } else {
        echo "Error: File not found: $filePath\n";
        exit(1);
    }
}

try {
    echo "Starting extraction of: " . basename($filePath) . "\n";
    echo "----------------------------------------\n";
    
    // Create extractor instance
    $extractor = new ExcelToMySQLExtractor(
        $config['host'],
        $config['database'],
        $config['username'],
        $config['password']
    );
    
    // Extract the Excel file
    $startTime = microtime(true);
    $extractor->extractExcelFile($filePath);
    $endTime = microtime(true);
    
    $executionTime = round($endTime - $startTime, 2);
    echo "----------------------------------------\n";
    echo "Extraction completed in {$executionTime} seconds\n";
    
} catch (Exception $e) {
    echo "Error: " . $e->getMessage() . "\n";
    exit(1);
}