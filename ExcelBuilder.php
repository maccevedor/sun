<?php

require_once 'ExcelBuilderInterface.php';
require_once 'Worksheet.php';

/**
 * Main ExcelBuilder class for creating Excel-like files using native PHP
 */
class ExcelBuilder implements ExcelBuilderInterface
{
    private array $worksheets = [];
    private string $title;
    private string $author;

    public function __construct(string $title = 'Workbook', string $author = 'PHP Excel Builder')
    {
        $this->title = $title;
        $this->author = $author;
    }

    /**
     * Create a new worksheet
     */
    public function createWorksheet(string $name): WorksheetInterface
    {
        if (isset($this->worksheets[$name])) {
            throw new InvalidArgumentException("Worksheet '{$name}' already exists");
        }

        $worksheet = new Worksheet($name);
        $this->worksheets[$name] = $worksheet;
        return $worksheet;
    }

    /**
     * Get worksheet by name
     */
    public function getWorksheet(string $name): ?WorksheetInterface
    {
        return $this->worksheets[$name] ?? null;
    }

    /**
     * Get all worksheets
     */
    public function getWorksheets(): array
    {
        return $this->worksheets;
    }

    /**
     * Save the Excel file (as CSV for simplicity, but structured like Excel)
     */
    public function save(string $filename): bool
    {
        try {
            $content = $this->generateExcelContent();
            return file_put_contents($filename, $content) !== false;
        } catch (Exception $e) {
            return false;
        }
    }

    /**
     * Convert to CSV format
     */
    public function toCsv(): string
    {
        return $this->generateCsvContent();
    }

    /**
     * Generate Excel-like content using recursion
     */
    private function generateExcelContent(): string
    {
        $content = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n";
        $content .= "<workbook>\n";
        $content .= "  <properties>\n";
        $content .= "    <title>{$this->title}</title>\n";
        $content .= "    <author>{$this->author}</author>\n";
        $content .= "    <created>" . date('Y-m-d H:i:s') . "</created>\n";
        $content .= "  </properties>\n";
        $content .= "  <worksheets>\n";

        // Recursively process all worksheets
        $content .= $this->processWorksheetsRecursive($this->worksheets);

        $content .= "  </worksheets>\n";
        $content .= "</workbook>";

        return $content;
    }

    /**
     * Recursively process worksheets
     */
    private function processWorksheetsRecursive(array $worksheets, int $depth = 0): string
    {
        $content = '';
        $indent = str_repeat('  ', $depth + 2);

        foreach ($worksheets as $name => $worksheet) {
            $content .= "{$indent}<worksheet name=\"{$name}\">\n";

            if ($worksheet instanceof Worksheet) {
                $content .= $this->processCellsRecursive($worksheet->getCells(), $depth + 1);
            }

            $content .= "{$indent}</worksheet>\n";
        }

        return $content;
    }

    /**
     * Recursively process cells
     */
    private function processCellsRecursive(array $cells, int $depth = 0): string
    {
        $content = '';
        $indent = str_repeat('  ', $depth + 2);

        foreach ($cells as $coordinate => $cell) {
            if ($cell instanceof Cell) {
                $value = htmlspecialchars($cell->getFormattedValue());
                $format = $cell->getFormat();
                $content .= "{$indent}<cell coordinate=\"{$coordinate}\" format=\"{$format}\">{$value}</cell>\n";
            }
        }

        return $content;
    }

    /**
     * Generate CSV content
     */
    private function generateCsvContent(): string
    {
        $csvContent = '';

        foreach ($this->worksheets as $name => $worksheet) {
            $csvContent .= "=== Worksheet: {$name} ===\n";

            if ($worksheet instanceof Worksheet) {
                $array = $worksheet->toArray();
                $csvContent .= $this->arrayToCsvRecursive($array);
            }

            $csvContent .= "\n";
        }

        return $csvContent;
    }

    /**
     * Convert array to CSV using recursion
     */
    private function arrayToCsvRecursive(array $data, int $depth = 0): string
    {
        $csv = '';

        foreach ($data as $row) {
            if (is_array($row)) {
                $csv .= $this->arrayToCsvRecursive([$row], $depth + 1);
            } else {
                $csv .= implode(',', array_map(function($cell) {
                    return '"' . str_replace('"', '""', (string)$cell) . '"';
                }, is_array($row) ? $row : [$row])) . "\n";
            }
        }

        return $csv;
    }

    /**
     * Import data from array using recursion
     */
    public function importFromArray(string $worksheetName, array $data): self
    {
        $worksheet = $this->createWorksheet($worksheetName);
        $this->importArrayRecursive($worksheet, $data, 1, 1);
        return $this;
    }

    /**
     * Recursively import array data
     */
    private function importArrayRecursive(Worksheet $worksheet, array $data, int $startRow, int $startCol): void
    {
        $row = $startRow;

        foreach ($data as $rowData) {
            if (is_array($rowData)) {
                $col = $startCol;
                foreach ($rowData as $cellValue) {
                    if (is_array($cellValue)) {
                        // Handle nested arrays recursively
                        $this->importArrayRecursive($worksheet, [$cellValue], $row, $col);
                    } else {
                        $coordinate = $this->numberToColumnLetter($col) . $row;
                        $worksheet->setCell($coordinate, $cellValue);
                        $col++;
                    }
                }
            } else {
                $coordinate = $this->numberToColumnLetter($startCol) . $row;
                $worksheet->setCell($coordinate, $rowData);
            }
            $row++;
        }
    }

    /**
     * Convert column number to letter (helper method)
     */
    private function numberToColumnLetter(int $num): string
    {
        if ($num <= 0) {
            return '';
        }

        if ($num <= 26) {
            return chr(64 + $num);
        }

        return $this->numberToColumnLetter(intval(($num - 1) / 26)) .
               chr(65 + (($num - 1) % 26));
    }

    /**
     * Search for values across all worksheets using recursion
     */
    public function searchValue($searchValue): array
    {
        return $this->searchValueRecursive($this->worksheets, $searchValue);
    }

    /**
     * Recursively search for values
     */
    private function searchValueRecursive(array $worksheets, $searchValue, array $results = []): array
    {
        foreach ($worksheets as $name => $worksheet) {
            if ($worksheet instanceof Worksheet) {
                $found = $worksheet->findCellsByValue($searchValue);
                if (!empty($found)) {
                    $results[$name] = $found;
                }
            }
        }

        return $results;
    }

    /**
     * Get workbook statistics
     */
    public function getStatistics(): array
    {
        $stats = [
            'worksheets' => count($this->worksheets),
            'total_cells' => 0,
            'worksheet_details' => []
        ];

        foreach ($this->worksheets as $name => $worksheet) {
            if ($worksheet instanceof Worksheet) {
                $cellCount = count($worksheet->getCells());
                $dimensions = $worksheet->getDimensions();

                $stats['total_cells'] += $cellCount;
                $stats['worksheet_details'][$name] = [
                    'cells' => $cellCount,
                    'dimensions' => $dimensions
                ];
            }
        }

        return $stats;
    }

    /**
     * Clone worksheet using recursion
     */
    public function cloneWorksheet(string $sourceName, string $targetName): bool
    {
        $sourceWorksheet = $this->getWorksheet($sourceName);
        if (!$sourceWorksheet) {
            return false;
        }

        $targetWorksheet = $this->createWorksheet($targetName);
        $this->cloneWorksheetRecursive($sourceWorksheet, $targetWorksheet);

        return true;
    }

    /**
     * Recursively clone worksheet data
     */
    private function cloneWorksheetRecursive(Worksheet $source, Worksheet $target): void
    {
        foreach ($source->getCells() as $coordinate => $cell) {
            if ($cell instanceof Cell) {
                $target->setCellWithFormat(
                    $coordinate,
                    $cell->getValue(),
                    $cell->getFormat()
                );
            }
        }
    }
}
