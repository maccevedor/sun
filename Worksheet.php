<?php

require_once 'ExcelBuilderInterface.php';
require_once 'Cell.php';

/**
 * Worksheet class representing an Excel worksheet
 */
class Worksheet implements WorksheetInterface
{
    private string $name;
    private array $cells = [];

    public function __construct(string $name)
    {
        $this->name = $name;
    }

    /**
     * Set cell value
     */
    public function setCell(string $coordinate, $value): self
    {
        $this->cells[$coordinate] = new Cell($coordinate, $value);
        return $this;
    }

    /**
     * Get cell value
     */
    public function getCell(string $coordinate)
    {
        return isset($this->cells[$coordinate]) ? $this->cells[$coordinate]->getValue() : null;
    }

    /**
     * Set multiple cells from array using recursion
     */
    public function setCells(array $data): self
    {
        $this->setCellsRecursive($data, 1, 1);
        return $this;
    }

    /**
     * Recursive function to set cells from nested arrays
     */
    private function setCellsRecursive(array $data, int $row, int $col): void
    {
        foreach ($data as $key => $value) {
            if (is_array($value)) {
                // If value is array, recursively process it
                $this->setCellsRecursive($value, $row, $col);
                $row++;
            } else {
                // Convert column number to letter (A, B, C, etc.)
                $coordinate = $this->numberToColumnLetter($col) . $row;
                $this->setCell($coordinate, $value);
                $col++;
            }
        }
    }

    /**
     * Convert column number to Excel column letter using recursion
     */
    public function numberToColumnLetter(int $num): string
    {
        if ($num <= 0) {
            return '';
        }

        if ($num <= 26) {
            return chr(64 + $num); // A=65, so 64+1=65
        }

        // Recursive case for columns beyond Z
        return $this->numberToColumnLetter(intval(($num - 1) / 26)) .
               chr(65 + (($num - 1) % 26));
    }

    /**
     * Get worksheet name
     */
    public function getName(): string
    {
        return $this->name;
    }

    /**
     * Get all cells
     */
    public function getCells(): array
    {
        return $this->cells;
    }

    /**
     * Set cell with format
     */
    public function setCellWithFormat(string $coordinate, $value, string $format): self
    {
        $cell = new Cell($coordinate, $value, $format);
        $this->cells[$coordinate] = $cell;
        return $this;
    }

    /**
     * Get cell object
     */
    public function getCellObject(string $coordinate): ?Cell
    {
        return $this->cells[$coordinate] ?? null;
    }

    /**
     * Find cells by value using recursion
     */
    public function findCellsByValue($searchValue): array
    {
        return $this->findCellsRecursive($this->cells, $searchValue);
    }

    /**
     * Recursive function to find cells by value
     */
    private function findCellsRecursive(array $cells, $searchValue, array $found = []): array
    {
        foreach ($cells as $coordinate => $cell) {
            if ($cell instanceof Cell) {
                if ($cell->getValue() === $searchValue) {
                    $found[] = $coordinate;
                }
            } elseif (is_array($cell)) {
                // If somehow we have nested arrays, handle recursively
                $found = array_merge($found, $this->findCellsRecursive($cell, $searchValue, $found));
            }
        }
        return $found;
    }

    /**
     * Get worksheet dimensions
     */
    public function getDimensions(): array
    {
        if (empty($this->cells)) {
            return ['rows' => 0, 'cols' => 0];
        }

        $maxRow = 0;
        $maxCol = 0;

        foreach (array_keys($this->cells) as $coordinate) {
            $parsed = $this->parseCoordinate($coordinate);
            $maxRow = max($maxRow, $parsed['row']);
            $maxCol = max($maxCol, $parsed['col']);
        }

        return ['rows' => $maxRow, 'cols' => $maxCol];
    }

    /**
     * Parse coordinate (e.g., "A1" -> ['row' => 1, 'col' => 1])
     */
    private function parseCoordinate(string $coordinate): array
    {
        preg_match('/([A-Z]+)(\d+)/', $coordinate, $matches);
        return [
            'row' => (int)$matches[2],
            'col' => $this->columnLetterToNumber($matches[1])
        ];
    }

    /**
     * Convert column letter to number
     */
    private function columnLetterToNumber(string $letters): int
    {
        $number = 0;
        $length = strlen($letters);

        for ($i = 0; $i < $length; $i++) {
            $number = $number * 26 + (ord($letters[$i]) - ord('A') + 1);
        }

        return $number;
    }

    /**
     * Convert worksheet to array representation
     */
    public function toArray(): array
    {
        $dimensions = $this->getDimensions();
        $result = [];

        for ($row = 1; $row <= $dimensions['rows']; $row++) {
            for ($col = 1; $col <= $dimensions['cols']; $col++) {
                $coordinate = $this->numberToColumnLetter($col) . $row;
                $result[$row - 1][$col - 1] = $this->getCell($coordinate);
            }
        }

        return $result;
    }
}
