<?php

namespace mradang\LaravelExcel\Services;

use Closure;

class ExcelService
{
    public static function read(
        string $pathname,
        int $fieldRow,
        array $fields,
        int $firstDataRow,
        Closure $callback
    ): void {
        Spout::read($pathname, $fieldRow, $fields, $firstDataRow, $callback);
    }

    public static function write(array $fields, $values = null, ?Closure $rowCallback = null): string
    {
        return Spout::write($fields, $values, $rowCallback);
    }

    public static function makeUsePhpSpreadsheet(
        string $title,
        array $fields,
        array $values,
        int $freezeColumnIndex = 0
    ): string {
        return PhpSpreadsheet::make($title, $fields, $values, $freezeColumnIndex);
    }

    public static function writeUsePhpSpreadsheet(
        string $title,
        array $fields,
        array $numericColumns,
        $values,
        ?Closure $rowCallback = null,
        int $freezeColumnIndex = 0
    ): string {
        return PhpSpreadsheet::write(
            $title,
            $fields,
            $numericColumns,
            $values,
            $rowCallback,
            $freezeColumnIndex
        );
    }

    public static function getHighestRow(string $inputFileName, int $sheetIndex = 0): int
    {
        return PhpSpreadsheet::getHighestRow($inputFileName, $sheetIndex);
    }
}
