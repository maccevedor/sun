<?php

require_once 'ExcelBuilder.php';

/**
 * Test suite for the Excel Builder
 * This validates all functionality for the interview
 */

class ExcelBuilderTest
{
    private int $testsPassed = 0;
    private int $testsTotal = 0;
    private array $errors = [];

    public function runAllTests(): void
    {
        echo "=== PHP Excel Builder Test Suite ===\n\n";

        $this->testInterfaceImplementation();
        $this->testCellFunctionality();
        $this->testWorksheetFunctionality();
        $this->testRecursiveFunctions();
        $this->testExcelBuilderCore();
        $this->testFileOperations();
        $this->testErrorHandling();

        $this->printResults();
    }

    private function assert(bool $condition, string $message): void
    {
        $this->testsTotal++;
        if ($condition) {
            $this->testsPassed++;
            echo "✓ {$message}\n";
        } else {
            $this->errors[] = $message;
            echo "✗ {$message}\n";
        }
    }

    private function testInterfaceImplementation(): void
    {
        echo "Testing Interface Implementation...\n";

        $excel = new ExcelBuilder();
        $this->assert($excel instanceof ExcelBuilderInterface,
                     "ExcelBuilder implements ExcelBuilderInterface");

        $worksheet = $excel->createWorksheet('Test');
        $this->assert($worksheet instanceof WorksheetInterface,
                     "Worksheet implements WorksheetInterface");

        $cell = new Cell('A1', 'test');
        $this->assert($cell instanceof CellInterface,
                     "Cell implements CellInterface");

        echo "\n";
    }

    private function testCellFunctionality(): void
    {
        echo "Testing Cell Functionality...\n";

        $cell = new Cell('A1', 'Hello World');
        $this->assert($cell->getValue() === 'Hello World',
                     "Cell stores and retrieves value correctly");

        $cell->setValue(123.45);
        $this->assert($cell->getValue() === 123.45,
                     "Cell value can be updated");

        $cell->setFormat('currency');
        $this->assert($cell->getFormat() === 'currency',
                     "Cell format can be set");

        $this->assert($cell->getFormattedValue() === '$123.45',
                     "Currency formatting works correctly");

        $cell->setFormat('percentage')->setValue(0.75);
        $this->assert($cell->getFormattedValue() === '75.00%',
                     "Percentage formatting works correctly");

        $this->assert($cell->getCoordinate() === 'A1',
                     "Cell coordinate is correct");

        echo "\n";
    }

    private function testWorksheetFunctionality(): void
    {
        echo "Testing Worksheet Functionality...\n";

        $worksheet = new Worksheet('TestSheet');
        $this->assert($worksheet->getName() === 'TestSheet',
                     "Worksheet name is set correctly");

        $worksheet->setCell('A1', 'Test Value');
        $this->assert($worksheet->getCell('A1') === 'Test Value',
                     "Cell value can be set and retrieved");

        $worksheet->setCellWithFormat('B1', 100, 'currency');
        $cellObj = $worksheet->getCellObject('B1');
        $this->assert($cellObj->getFormattedValue() === '$100.00',
                     "Cell with format works correctly");

        // Test array data setting
        $data = [
            ['Name', 'Age', 'Salary'],
            ['John', 30, 50000],
            ['Jane', 25, 45000]
        ];
        $worksheet->setCells($data);
        $this->assert($worksheet->getCell('A1') === 'Name',
                     "Array data import works - header");
        $this->assert($worksheet->getCell('C2') === 50000,
                     "Array data import works - numeric value");

        // Test search functionality
        $found = $worksheet->findCellsByValue('John');
        $this->assert(!empty($found),
                     "Cell search functionality works");

        echo "\n";
    }

    private function testRecursiveFunctions(): void
    {
        echo "Testing Recursive Functions...\n";

        $worksheet = new Worksheet('RecursiveTest');

        // Test recursive column letter conversion
        $this->assert($worksheet->numberToColumnLetter(1) === 'A',
                     "Column 1 converts to A");
        $this->assert($worksheet->numberToColumnLetter(26) === 'Z',
                     "Column 26 converts to Z");
        $this->assert($worksheet->numberToColumnLetter(27) === 'AA',
                     "Column 27 converts to AA (recursive case)");

        // Test nested array processing
        $nestedData = [
            ['Product', 'Q1', 'Q2', 'Q3', 'Q4'],
            ['Laptops', [100, 120, 110, 130]],
            ['Phones', [200, 180, 220, 250]]
        ];

        $worksheet->setCells($nestedData);
        $this->assert($worksheet->getCell('A1') === 'Product',
                     "Nested array processing - header");

        // Test recursive search across complex data
        $worksheet->setCell('Z99', 'DeepValue');
        $found = $worksheet->findCellsByValue('DeepValue');
        $this->assert(in_array('Z99', $found),
                     "Recursive search finds deep values");

        echo "\n";
    }

    private function testExcelBuilderCore(): void
    {
        echo "Testing ExcelBuilder Core Functionality...\n";

        $excel = new ExcelBuilder('Test Workbook', 'Test Author');

        // Test worksheet creation
        $sheet1 = $excel->createWorksheet('Sheet1');
        $this->assert($sheet1 instanceof WorksheetInterface,
                     "Worksheet creation returns correct interface");

        $sheet2 = $excel->createWorksheet('Sheet2');
        $worksheets = $excel->getWorksheets();
        $this->assert(count($worksheets) === 2,
                     "Multiple worksheets can be created");

        // Test worksheet retrieval
        $retrieved = $excel->getWorksheet('Sheet1');
        $this->assert($retrieved === $sheet1,
                     "Worksheet can be retrieved by name");

        $this->assert($excel->getWorksheet('NonExistent') === null,
                     "Non-existent worksheet returns null");

        // Test data import
        $testData = [
            ['Item', 'Price'],
            ['Apple', 1.50],
            ['Banana', 0.75]
        ];
        $excel->importFromArray('ImportTest', $testData);
        $importSheet = $excel->getWorksheet('ImportTest');
        $this->assert($importSheet->getCell('A1') === 'Item',
                     "Array import creates worksheet correctly");

        // Test search across workbooks
        $sheet1->setCell('A1', 'SearchMe');
        $sheet2->setCell('B2', 'SearchMe');
        $results = $excel->searchValue('SearchMe');
        $this->assert(count($results) === 2,
                     "Search across multiple worksheets works");

        // Test cloning
        $sheet1->setCell('C3', 'CloneTest');
        $cloned = $excel->cloneWorksheet('Sheet1', 'ClonedSheet');
        $this->assert($cloned === true,
                     "Worksheet cloning succeeds");

        $clonedSheet = $excel->getWorksheet('ClonedSheet');
        $this->assert($clonedSheet->getCell('C3') === 'CloneTest',
                     "Cloned worksheet contains original data");

        // Test statistics
        $stats = $excel->getStatistics();
        $this->assert($stats['worksheets'] >= 4,
                     "Statistics show correct worksheet count");
        $this->assert($stats['total_cells'] > 0,
                     "Statistics show cell count");

        echo "\n";
    }

    private function testFileOperations(): void
    {
        echo "Testing File Operations...\n";

        $excel = new ExcelBuilder('File Test', 'Tester');
        $sheet = $excel->createWorksheet('FileTest');
        $sheet->setCell('A1', 'File Test Data')
              ->setCellWithFormat('B1', 123.45, 'currency');

        // Test XML generation
        $saved = $excel->save('test_output.xml');
        $this->assert($saved === true,
                     "Excel file can be saved");

        if (file_exists('test_output.xml')) {
            $content = file_get_contents('test_output.xml');
            $this->assert(strpos($content, '<workbook>') !== false,
                         "Saved file contains XML structure");
            $this->assert(strpos($content, 'File Test Data') !== false,
                         "Saved file contains cell data");
            unlink('test_output.xml'); // Clean up
        }

        // Test CSV generation
        $csv = $excel->toCsv();
        $this->assert(!empty($csv),
                     "CSV content is generated");
        $this->assert(strpos($csv, 'FileTest') !== false,
                     "CSV contains worksheet name");

        echo "\n";
    }

    private function testErrorHandling(): void
    {
        echo "Testing Error Handling...\n";

        $excel = new ExcelBuilder();
        $excel->createWorksheet('Test');

        try {
            $excel->createWorksheet('Test'); // Duplicate name
            $this->assert(false, "Duplicate worksheet should throw exception");
        } catch (InvalidArgumentException $e) {
            $this->assert(true, "Duplicate worksheet throws correct exception");
        }

        // Test cloning non-existent worksheet
        $result = $excel->cloneWorksheet('NonExistent', 'Clone');
        $this->assert($result === false,
                     "Cloning non-existent worksheet returns false");

        echo "\n";
    }

    private function printResults(): void
    {
        echo "=== Test Results ===\n";
        echo "Tests Passed: {$this->testsPassed}/{$this->testsTotal}\n";
        echo "Success Rate: " . round(($this->testsPassed / $this->testsTotal) * 100, 2) . "%\n";

        if (!empty($this->errors)) {
            echo "\nFailed Tests:\n";
            foreach ($this->errors as $error) {
                echo "- {$error}\n";
            }
        } else {
            echo "\n🎉 All tests passed!\n";
        }

        echo "\n=== Interview Readiness Assessment ===\n";
        if ($this->testsPassed === $this->testsTotal) {
            echo "✅ EXCELLENT: Code demonstrates strong OOP, interface usage, and recursion\n";
            echo "✅ All core functionality working correctly\n";
            echo "✅ Error handling implemented\n";
            echo "✅ Ready for interview!\n";
        } else {
            $percentage = ($this->testsPassed / $this->testsTotal) * 100;
            if ($percentage >= 80) {
                echo "✅ GOOD: Most functionality working, minor issues to address\n";
            } elseif ($percentage >= 60) {
                echo "⚠️  NEEDS WORK: Core functionality present but several issues\n";
            } else {
                echo "❌ MAJOR ISSUES: Significant problems need to be resolved\n";
            }
        }
    }
}

// Run the tests
$tester = new ExcelBuilderTest();
$tester->runAllTests();
