<?php

require_once 'ExcelBuilder.php';

/**
 * Interview Demo - PHP Excel Builder
 * Focuses on the core requirements: OOP, Interfaces, and Recursion
 */

echo "=== PHP Excel Builder Interview Demo ===\n\n";

// 1. DEMONSTRATE OBJECT-ORIENTED PROGRAMMING
echo "1. Object-Oriented Programming Demo:\n";
$excel = new ExcelBuilder('Interview Demo', 'Candidate');
echo "✓ Created ExcelBuilder instance with constructor parameters\n";

// 2. DEMONSTRATE INTERFACE IMPLEMENTATION
echo "\n2. Interface Implementation Demo:\n";
$sheet = $excel->createWorksheet('Demo Sheet');
echo "✓ ExcelBuilder implements ExcelBuilderInterface\n";
echo "✓ Worksheet implements WorksheetInterface\n";
echo "✓ Method chaining with fluent interface\n";

// 3. DEMONSTRATE RECURSIVE FUNCTIONS
echo "\n3. Recursive Functions Demo:\n";

// Recursive column letter conversion
echo "Column Letter Conversion (Recursive):\n";
for ($i = 1; $i <= 30; $i++) {
    $letter = $sheet->numberToColumnLetter($i);
    echo "  {$i} → {$letter}";
    if ($i % 10 == 0) echo "\n";
    if ($i % 5 == 0 && $i % 10 != 0) echo " | ";
}
echo "\n";

// 4. DEMONSTRATE PRACTICAL EXCEL FUNCTIONALITY
echo "\n4. Excel-like Functionality:\n";

// Add headers
$sheet->setCell('A1', 'Product')
      ->setCell('B1', 'Quantity')
      ->setCell('C1', 'Unit Price')
      ->setCell('D1', 'Total');

// Add data with different formats
$products = [
    ['Laptop', 5, 999.99],
    ['Mouse', 20, 25.50],
    ['Keyboard', 15, 75.00]
];

$row = 2;
foreach ($products as $product) {
    $sheet->setCell("A{$row}", $product[0])
          ->setCellWithFormat("B{$row}", $product[1], 'number')
          ->setCellWithFormat("C{$row}", $product[2], 'currency')
          ->setCellWithFormat("D{$row}", $product[1] * $product[2], 'currency');
    $row++;
}

echo "✓ Added data with different cell formats\n";

// 5. DEMONSTRATE RECURSIVE SEARCH
echo "\n5. Recursive Search Demo:\n";
$found = $sheet->findCellsByValue('Laptop');
echo "✓ Found 'Laptop' in cells: " . implode(', ', $found) . "\n";

// 6. SAVE AS EXCEL-COMPATIBLE FORMAT
echo "\n6. File Generation:\n";

// Save as CSV (Excel-compatible)
$csvContent = $excel->toCsv();
file_put_contents('demo_output.csv', $csvContent);
echo "✓ Generated Excel-compatible CSV file\n";

// 7. INTERVIEW TALKING POINTS
echo "\n=== INTERVIEW TALKING POINTS ===\n";

echo "\n🎯 Object-Oriented Programming:\n";
echo "- Encapsulation: Private properties with public methods\n";
echo "- Inheritance: Interface implementations\n";
echo "- Polymorphism: Different cell formats, method overloading\n";
echo "- Abstraction: Clean interface contracts\n";

echo "\n🔄 Recursive Functions:\n";
echo "- Column conversion: Base case (≤26), recursive case (>26)\n";
echo "- Data processing: Handles nested arrays recursively\n";
echo "- Search operations: Traverses complex structures\n";

echo "\n📊 Excel Format Reality:\n";
echo "- Real .xlsx files require complex ZIP/XML structure\n";
echo "- Native PHP solution: CSV format (Excel-compatible)\n";
echo "- Production solution: Use PhpSpreadsheet library\n";

echo "\n✅ Core Requirements Met:\n";
echo "- ✓ Object-oriented programming with classes\n";
echo "- ✓ Interface creation and implementation\n";
echo "- ✓ Recursive functions with proper base cases\n";
echo "- ✓ Native PHP (no extensions)\n";
echo "- ✓ Excel-like functionality\n";

// 8. DEMONSTRATE ERROR HANDLING
echo "\n8. Error Handling Demo:\n";
try {
    $excel->createWorksheet('Demo Sheet'); // Duplicate name
} catch (InvalidArgumentException $e) {
    echo "✓ Proper exception handling: " . $e->getMessage() . "\n";
}

echo "\n=== Demo Complete! ===\n";
echo "Files generated: demo_output.csv\n";
echo "Ready for interview questions about OOP, interfaces, and recursion!\n";
