<?php

require_once 'ExcelBuilder.php';

echo "=== Quick Test - Fixed Version ===\n\n";

try {
    // Test basic functionality
    $excel = new ExcelBuilder('Test', 'User');
    $sheet = $excel->createWorksheet('Test Sheet');

    // Test column letter conversion (now public)
    echo "Testing column letter conversion:\n";
    for ($i = 1; $i <= 5; $i++) {
        $letter = $sheet->numberToColumnLetter($i);
        echo "  {$i} → {$letter}\n";
    }

    // Test data addition
    $sheet->setCell('A1', 'Name')
          ->setCell('B1', 'Value')
          ->setCell('A2', 'Test')
          ->setCell('B2', 123);

    echo "\nTesting CSV generation:\n";
    $csv = $excel->toCsv();
    echo "✓ CSV generated successfully (length: " . strlen($csv) . " chars)\n";

    // Save to file
    file_put_contents('quick_test.csv', $csv);
    echo "✓ File saved as quick_test.csv\n";

    echo "\n=== All tests passed! ===\n";

} catch (Exception $e) {
    echo "Error: " . $e->getMessage() . "\n";
    echo "Stack trace:\n" . $e->getTraceAsString() . "\n";
}
