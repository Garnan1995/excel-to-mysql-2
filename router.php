<?php
// router.php

// Handle API requests
if (preg_match('/^\/api\/(.*)/', $_SERVER["REQUEST_URI"], $matches)) {
    $apiFile = __DIR__ . '/api/' . $matches[1];
    if (file_exists($apiFile)) {
        require $apiFile;
        return true;
    }
}

// Handle static files from public directory
$publicPath = __DIR__ . '/public' . $_SERVER["REQUEST_URI"];
if (file_exists($publicPath) && !is_dir($publicPath)) {
    return false; // Let PHP built-in server handle it
}

// Default to index.html
require __DIR__ . '/public/index.html';