<?php
// router.php

$uri = parse_url($_SERVER['REQUEST_URI'], PHP_URL_PATH);

// Handle API requests
if (strpos($uri, '/api/') === 0) {
    $apiFile = __DIR__ . $uri;
    if (file_exists($apiFile)) {
        require $apiFile;
        exit;
    }
    http_response_code(404);
    echo "API endpoint not found";
    exit;
}

// Handle static files (JS, CSS, images)
$staticFile = __DIR__ . '/public' . $uri;
if (file_exists($staticFile) && !is_dir($staticFile)) {
    // Set correct content type
    $ext = pathinfo($staticFile, PATHINFO_EXTENSION);
    switch($ext) {
        case 'js':
            header('Content-Type: application/javascript');
            break;
        case 'css':
            header('Content-Type: text/css');
            break;
        case 'html':
            header('Content-Type: text/html');
            break;
    }
    readfile($staticFile);
    exit;
}

// Default to index.html for root or any other request
header('Content-Type: text/html');
readfile(__DIR__ . '/public/index.html');