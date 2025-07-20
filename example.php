<?php

require_once 'ExcelBuilder.php';

/**
 * Example usage of the Excel Builder
 * This demonstrates all the key features for the interview
 */

echo "=== PHP Excel Builder Interview Example ===\n\n";

try {
    // 1. Create a new Excel Builder instance
    $excel = new ExcelBuilder('Sales Report 2024', 'Interview Candidate');

    // 2. Create worksheets
    $salesSheet = $excel->createWorksheet('Sales Data');
    $summarySheet = $excel->createWorksheet('Summary');

    // 3. Add data to Sales Data worksheet
    echo "Adding data to Sales Data worksheet...\n";

    // Individual cell setting
    $salesSheet->setCell('A1', 'Product')
             ->setCell('B1', 'Quantity')
             ->setCell('C1', 'Price')
             ->setCell('D1', 'Total');

    // Set cells with formatting
    $salesSheet->setCellWithFormat('A2', 'Laptop', 'general')
             ->setCellWithFormat('B2', 10, 'number')
             ->setCellWithFormat('C2', 999.99, 'currency')
             ->setCellWithFormat('D2', 9999.90, 'currency');

    $salesSheet->setCellWithFormat('A3', 'Mouse', 'general')
             ->setCellWithFormat('B3', 50, 'number')
             ->setCellWithFormat('C3', 25.50, 'currency')
             ->setCellWithFormat('D3', 1275.00, 'currency');

    $salesSheet->setCellWithFormat('A4', 'Keyboard', 'general')
             ->setCellWithFormat('B4', 30, 'number')
             ->setCellWithFormat('C4', 75.00, 'currency')
             ->setCellWithFormat('D4', 2250.00, 'currency');

    // 4. Use recursive function to import array data
    echo "Importing array data using recursion...\n";

    $arrayData = [
        ['Monitor', 15, 299.99, 4499.85],
        ['Headphones', 25, 89.99, 2249.75],
        ['Webcam', 20, 149.99, 2999.80]
    ];

    // This will use the recursive setCells method
    $row = 5;
    foreach ($arrayData as $rowData) {
        $col = 1;
        foreach ($rowData as $cellValue) {
            $coordinate = chr(64 + $col) . $row; // A5, B5, C5, etc.
            if ($col == 3 || $col == 4) { // Price and Total columns
                $salesSheet->setCellWithFormat($coordinate, $cellValue, 'currency');
            } else {
                $salesSheet->setCell($coordinate, $cellValue);
            }
            $col++;
        }
        $row++;
    }

    // 5. Create summary using recursive search
    echo "Creating summary using recursive functions...\n";

    $summarySheet->setCell('A1', 'Summary Report')
               ->setCell('A3', 'Total Products:')
               ->setCell('B3', 7)
               ->setCell('A4', 'Generated:')
               ->setCell('B4', date('Y-m-d H:i:s'));

    // 6. Demonstrate recursive search functionality
    echo "Searching for values using recursion...\n";

    $searchResults = $excel->searchValue('Laptop');
    if (!empty($searchResults)) {
        echo "Found 'Laptop' in cells: " . json_encode($searchResults) . "\n";
    }

    // 7. Clone worksheet using recursion
    echo "Cloning worksheet using recursion...\n";
    $excel->cloneWorksheet('Sales Data', 'Backup Sales Data');

    // 8. Get statistics
    $stats = $excel->getStatistics();
    echo "Workbook Statistics:\n";
    echo "- Total Worksheets: " . $stats['worksheets'] . "\n";
    echo "- Total Cells: " . $stats['total_cells'] . "\n";

    foreach ($stats['worksheet_details'] as $name => $details) {
        echo "- {$name}: {$details['cells']} cells, " .
             "{$details['dimensions']['rows']} rows x {$details['dimensions']['cols']} cols\n";
    }

    // 9. Save to file
    echo "\nSaving Excel file...\n";
    if ($excel->save('sales_report.xml')) {
        echo "✓ Excel file saved as 'sales_report.xml'\n";
    } else {
        echo "✗ Failed to save Excel file\n";
    }

    // 10. Export to CSV
    echo "\nExporting to CSV...\n";
    $csvContent = $excel->toCsv();
    if (file_put_contents('sales_report.csv', $csvContent)) {
        echo "✓ CSV file saved as 'sales_report.csv'\n";
    } else {
        echo "✗ Failed to save CSV file\n";
    }

    // 11. Demonstrate cell formatting
    echo "\nDemonstrating cell formatting:\n";
    $testSheet = $excel->createWorksheet('Format Test');

    $testSheet->setCellWithFormat('A1', 1234.56, 'currency')
             ->setCellWithFormat('A2', 0.75, 'percentage')
             ->setCellWithFormat('A3', time(), 'date')
             ->setCellWithFormat('A4', 9876.54, 'number');

    echo "Currency: " . $testSheet->getCellObject('A1')->getFormattedValue() . "\n";
    echo "Percentage: " . $testSheet->getCellObject('A2')->getFormattedValue() . "\n";
    echo "Date: " . $testSheet->getCellObject('A3')->getFormattedValue() . "\n";
    echo "Number: " . $testSheet->getCellObject('A4')->getFormattedValue() . "\n";

    // 12. Show recursive column letter conversion
    echo "\nTesting recursive column letter conversion:\n";
    $worksheet = new Worksheet('Test');
    for ($i = 1; $i <= 30; $i++) {
        $letter = $worksheet->numberToColumnLetter($i);
        echo "Column {$i} = {$letter} ";
        if ($i % 10 == 0) echo "\n";
    }
    echo "\n";

    echo "\n=== Example completed successfully! ===\n";

} catch (Exception $e) {
    echo "Error: " . $e->getMessage() . "\n";
    echo "Stack trace:\n" . $e->getTraceAsString() . "\n";
}
