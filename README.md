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
├── test.php                  # Comprehensive test suite
└── README.md                 # This file
```

## 🚀 Quick Start

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

// Save file
$excel->save('output.xml');
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
- CSV export

## 🧪 Testing

Run the comprehensive test suite:

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

## 📊 Performance Considerations

- Efficient recursive algorithms with proper base cases
- Memory-conscious cell storage
- Lazy loading of worksheet data
- Optimized search operations

## 🔧 Extension Points

The design allows for easy extension:
- New cell formats
- Additional worksheet types
- Custom export formats
- Enhanced search capabilities

## 📋 Interview Preparation Tips

1. **Explain the recursive functions**: Focus on base cases and recursive cases
2. **Discuss interface benefits**: Contracts, testability, maintainability
3. **Demonstrate OOP principles**: Show encapsulation, inheritance examples
4. **Walk through the architecture**: Explain class relationships
5. **Discuss design decisions**: Why certain patterns were chosen

## 🏃‍♂️ Running the Code

```bash
# Run examples
php example.php

# Run tests
php test.php

# Check output files
ls -la *.xml *.csv
```

## 💡 Key Interview Insights

- **Native PHP**: No external dependencies, pure PHP implementation
- **Scalable Design**: Can handle large datasets through recursive processing
- **Clean Architecture**: Separation of concerns, single responsibility
- **Testable Code**: Comprehensive test coverage with clear assertions

This implementation demonstrates production-ready PHP code with strong architectural principles, making it ideal for technical interviews focused on OOP and recursive programming.
