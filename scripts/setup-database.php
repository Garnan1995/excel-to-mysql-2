<?php
require_once __DIR__ . '/../vendor/autoload.php';

$config = require __DIR__ . '/../config/database.php';

try {
    // Connect to MySQL without selecting a database
    $dsn = "mysql:host={$config['host']};charset={$config['charset']}";
    $pdo = new PDO($dsn, $config['username'], $config['password'], $config['options']);
    
    // Create database if it doesn't exist
    $dbName = $config['database'];
    $pdo->exec("CREATE DATABASE IF NOT EXISTS `$dbName` 
                CHARACTER SET utf8mb4 
                COLLATE utf8mb4_unicode_ci");
    
    echo "Database '$dbName' created successfully (or already exists).\n";
    
    // Connect to the specific database
    $pdo->exec("USE `$dbName`");
    
    // Read and execute the SQL schema file
    $sqlFile = __DIR__ . '/../src/schema.sql';
    
    // Check if schema file exists
    if (!file_exists($sqlFile)) {
        echo "Creating schema.sql file...\n";
        // Copy the SQL schema from our artifacts into a file
        $schemaSQL = file_get_contents(__DIR__ . '/../src/schema.sql');
        file_put_contents($sqlFile, $schemaSQL);
    }
    
    // Execute the schema
    $sql = file_get_contents($sqlFile);
    
    // Split by semicolon to execute multiple statements
    $statements = array_filter(array_map('trim', explode(';', $sql)));
    
    foreach ($statements as $statement) {
        if (!empty($statement)) {
            $pdo->exec($statement);
        }
    }
    
    echo "Database tables created successfully!\n";
    echo "Setup complete. You can now run the extraction script.\n";
    
} catch (PDOException $e) {
    die("Database setup failed: " . $e->getMessage() . "\n");
}