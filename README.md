# PHP Excel Builder - Interview Project

A comprehensive Excel builder implementation using native PHP classes, demonstrating object-oriented programming, interface design, and recursive functions.

## 🎯 Interview Focus Areas

This project specifically demonstrates:
- **Object-Oriented Programming**: Classes, inheritance, encapsulation
- **Interface Implementation**: Clean contract definitions
- **Recursive Functions**: Column conversion, data processing, search operations
- **Native PHP**: No external libraries or extensions

## 📁 Project Structure

```
├── ExcelBuilderInterface.php  # Core interfaces
├── Cell.php                   # Cell implementation
├── Worksheet.php             # Worksheet with recursive functions
├── ExcelBuilder.php          # Main builder class
├── example.php               # Usage examples
├── interview_demo.php        # Focused interview demo (RECOMMENDED)
├── quick_test.php            # Quick functionality test
├── test.php                  # Comprehensive test suite
└── README.md                 # This file
```

## 🚀 Quick Start

### Run the Interview Demo (Recommended)
```bash
php interview_demo.php
```

### Basic Usage
```php
<?php
require_once 'ExcelBuilder.php';

// Create Excel builder
$excel = new ExcelBuilder('My Workbook', 'Author Name');

// Create worksheet
$sheet = $excel->createWorksheet('Sales Data');

// Add data
$sheet->setCell('A1', 'Product')
      ->setCell('B1', 'Price')
      ->setCellWithFormat('A2', 'Laptop', 'general')
      ->setCellWithFormat('B2', 999.99, 'currency');

// Save as Excel-compatible CSV
$excel->saveAsExcel('output.csv');
```

## 🔧 Key Features

### 1. Interface-Driven Design
- `ExcelBuilderInterface`: Main builder contract
- `WorksheetInterface`: Worksheet operations
- `CellInterface`: Cell manipulation

### 2. Recursive Functions
- **Column Letter Conversion**: `numberToColumnLetter()` handles A-Z, AA-ZZ, etc.
- **Data Processing**: `setCellsRecursive()` processes nested arrays
- **Search Operations**: `findCellsRecursive()` searches through complex structures
- **Content Generation**: `processWorksheetsRecursive()` creates XML output

### 3. Cell Formatting
- Currency: `$1,234.56`
- Percentage: `75.00%`
- Date: `2024-01-15`
- Number: `1,234.56`

### 4. Advanced Operations
- Worksheet cloning
- Cross-worksheet search
- Statistics generation
- Array data import
- CSV export (Excel-compatible)

## 🧪 Testing

### Quick Test
```bash
php quick_test.php
```

### Full Test Suite
```bash
php test.php
```

The test suite validates:
- Interface implementations
- Cell functionality
- Worksheet operations
- Recursive functions
- File operations
- Error handling

## 📝 Usage Examples

### Basic Usage
```php
$excel = new ExcelBuilder();
$sheet = $excel->createWorksheet('Data');
$sheet->setCell('A1', 'Hello World');
```

### Array Import (Uses Recursion)
```php
$data = [
    ['Name', 'Age', 'Salary'],
    ['John', 30, 50000],
    ['Jane', 25, 45000]
];
$excel->importFromArray('Employees', $data);
```

### Search Across Workbook
```php
$results = $excel->searchValue('John');
// Returns: ['Employees' => ['A2']]
```

### Cell Formatting
```php
$sheet->setCellWithFormat('A1', 1234.56, 'currency');
echo $sheet->getCellObject('A1')->getFormattedValue(); // $1,234.56
```

## 🔍 Recursive Function Examples

### Column Letter Conversion
```php
// Converts numbers to Excel column letters recursively
1  → A
26 → Z
27 → AA
52 → AZ
53 → BA
```

### Nested Array Processing
```php
$nestedData = [
    ['Q1', 'Q2', 'Q3'],
    ['Sales', [100, 200, 300]],
    ['Costs', [50, 75, 125]]
];
// Recursively processes and places data in appropriate cells
```

## 🎯 Interview Talking Points

### Object-Oriented Design
- **Encapsulation**: Private properties with public methods
- **Inheritance**: Interface implementations
- **Polymorphism**: Different cell formats, worksheet types
- **Abstraction**: Clean interface contracts

### Recursive Functions
- **Base Cases**: Simple conditions (num <= 26)
- **Recursive Cases**: Self-calling with modified parameters
- **Practical Applications**: Column conversion, data traversal

### Design Patterns
- **Builder Pattern**: ExcelBuilder constructs complex objects
- **Factory Pattern**: Cell creation in worksheets
- **Strategy Pattern**: Different formatting strategies

## 🚨 Error Handling

The system includes comprehensive error handling:
- Duplicate worksheet names throw `InvalidArgumentException`
- Invalid operations return `false` or `null`
- File operations are wrapped in try-catch blocks
- Recursive functions have depth limits to prevent infinite loops

## 📊 Excel Format Reality

**Important Interview Point**: Native PHP cannot create real .xlsx files without extensions.

### Why CSV Instead of Excel?
- Real .xlsx files are ZIP archives containing XML files, relationships, and binary data
- Creating this natively would require thousands of lines of code
- CSV format is Excel-compatible and demonstrates the core programming concepts

### Production Solution
In a real application, you'd use:
- **PhpSpreadsheet**: The standard library for Excel file generation
- **Box/Spout**: For handling large files efficiently

## 🔧 Recent Fixes & Improvements

### ✅ Fixed Issues
- **Infinite Recursion Bug**: Fixed CSV generation method
- **Method Visibility**: Made `numberToColumnLetter()` public for demos
- **Depth Protection**: Added limits to prevent stack overflow

### ✅ Added Features
- **Interview Demo**: Focused demonstration script
- **Quick Test**: Simple validation script
- **Better Error Handling**: More robust exception management

## 🏃‍♂️ Running the Code

```bash
# Run the focused interview demo (BEST for interviews)
php interview_demo.php

# Run quick functionality test
php quick_test.php

# Run comprehensive test suite
php test.php

# Run original examples
php example.php

# Check output files
ls -la *.csv *.xml
```

## 📋 Interview Preparation Tips

1. **Start with the demo**: Run `php interview_demo.php` to see all features
2. **Explain the recursive functions**: Focus on base cases and recursive cases
3. **Discuss interface benefits**: Contracts, testability, maintainability
4. **Demonstrate OOP principles**: Show encapsulation, inheritance examples
5. **Address Excel format**: Explain why CSV is used and mention production alternatives
6. **Walk through the architecture**: Explain class relationships and design decisions

## 💡 Key Interview Insights

- **Native PHP**: No external dependencies, pure PHP implementation
- **Scalable Design**: Can handle large datasets through recursive processing
- **Clean Architecture**: Separation of concerns, single responsibility
- **Testable Code**: Comprehensive test coverage with clear assertions
- **Real-world Awareness**: Understanding of Excel format limitations and practical solutions

## 🎯 Interview Success Strategy

1. **Run the demo first**: `php interview_demo.php` shows everything working
2. **Explain the problem**: "Create Excel builder with OOP, interfaces, recursion"
3. **Show the solution**: Walk through the class hierarchy and key methods
4. **Demonstrate recursion**: Focus on column conversion algorithm
5. **Address limitations**: Explain Excel format reality and production solutions
6. **Show testing**: Run `php test.php` to demonstrate quality assurance

This implementation demonstrates production-ready PHP code with strong architectural principles, making it ideal for technical interviews focused on OOP and recursive programming.
