<?php

/**
 * Interface for Excel Builder functionality
 */
interface ExcelBuilderInterface
{
    /**
     * Create a new worksheet
     */
    public function createWorksheet(string $name): WorksheetInterface;

    /**
     * Get worksheet by name
     */
    public function getWorksheet(string $name): ?WorksheetInterface;

    /**
     * Get all worksheets
     */
    public function getWorksheets(): array;

    /**
     * Save the Excel file
     */
    public function save(string $filename): bool;

    /**
     * Convert to CSV format
     */
    public function toCsv(): string;
}

/**
 * Interface for Worksheet functionality
 */
interface WorksheetInterface
{
    /**
     * Set cell value
     */
    public function setCell(string $coordinate, $value): self;

    /**
     * Get cell value
     */
    public function getCell(string $coordinate);

    /**
     * Set multiple cells from array
     */
    public function setCells(array $data): self;

    /**
     * Get worksheet name
     */
    public function getName(): string;

    /**
     * Get all cells
     */
    public function getCells(): array;
}

/**
 * Interface for Cell functionality
 */
interface CellInterface
{
    /**
     * Set cell value
     */
    public function setValue($value): self;

    /**
     * Get cell value
     */
    public function getValue();

    /**
     * Set cell format
     */
    public function setFormat(string $format): self;

    /**
     * Get cell format
     */
    public function getFormat(): string;

    /**
     * Get cell coordinate
     */
    public function getCoordinate(): string;
}
