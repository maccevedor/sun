<?php

require_once 'ExcelBuilderInterface.php';

/**
 * Cell class representing a single Excel cell
 */
class Cell implements CellInterface
{
    private $value;
    private string $format;
    private string $coordinate;

    public function __construct(string $coordinate, $value = null, string $format = 'general')
    {
        $this->coordinate = $coordinate;
        $this->value = $value;
        $this->format = $format;
    }

    /**
     * Set cell value
     */
    public function setValue($value): self
    {
        $this->value = $value;
        return $this;
    }

    /**
     * Get cell value
     */
    public function getValue()
    {
        return $this->value;
    }

    /**
     * Set cell format
     */
    public function setFormat(string $format): self
    {
        $this->format = $format;
        return $this;
    }

    /**
     * Get cell format
     */
    public function getFormat(): string
    {
        return $this->format;
    }

    /**
     * Get cell coordinate
     */
    public function getCoordinate(): string
    {
        return $this->coordinate;
    }

    /**
     * Format value based on cell format
     */
    public function getFormattedValue(): string
    {
        if ($this->value === null) {
            return '';
        }

        switch ($this->format) {
            case 'currency':
                return '$' . number_format((float)$this->value, 2);
            case 'percentage':
                return number_format((float)$this->value * 100, 2) . '%';
            case 'date':
                if (is_numeric($this->value)) {
                    return date('Y-m-d', (int)$this->value);
                }
                return (string)$this->value;
            case 'number':
                return number_format((float)$this->value, 2);
            default:
                return (string)$this->value;
        }
    }

    /**
     * Convert cell to string representation
     */
    public function __toString(): string
    {
        return $this->getFormattedValue();
    }
}
