<?php

namespace mradang\LaravelExcel\Services;

use Closure;
use Illuminate\Database\Eloquent\Builder;
use Illuminate\Support\Arr;
use Illuminate\Support\Collection;
use Illuminate\Support\Str;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Cell\DataType;
use PhpOffice\PhpSpreadsheet\IOFactory;

class PhpSpreadsheet
{
    const DEFAULT_ROW_HEIGHT = 14.4;
    const SINGLE_CHARACTER_WIDTH = 1.1;
    const MAX_COLUMN_WIDTH = 100;

    // 生成excel临时文件
    // $fields 字段名数组
    // $values 字段值数组
    // 返回全路径文件名
    public static function make(string $title, array $fields, array $values, int $freezeColumnIndex = 0): string
    {
        ini_set('memory_limit', '512M');
        $fields = array_values($fields);

        $spreadsheet = new Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();
        $w = []; // 每列最大宽度值

        // 设置标题
        if ($title) {
            $sheet->mergeCellsByColumnAndRow(1, 1, count($fields), 1);
            $sheet->setCellValue('A1', $title);
            $sheet->getStyle('A1')->getAlignment()->setWrapText(true);
            $sheet->getStyle('A1')->getAlignment()->setVertical('center');
        }

        // 设置行号常量
        $title_row = $title ? 1 : 0;
        $field_row = $title_row + 1;
        $data_row = $field_row + 1;

        // 写入字段名
        for ($i = 0; $i < count($fields); $i++) {
            $sheet->setCellValueExplicitByColumnAndRow($i + 1, $field_row, $fields[$i], DataType::TYPE_STRING);
            $width = mb_strwidth($fields[$i], 'UTF-8');
            $w[$i] = max($width, 4 * 2); // 最小宽度 4 个汉字
        }

        // 冻结表头
        if ($freezeColumnIndex) {
            $sheet->freezePaneByColumnAndRow($freezeColumnIndex, $data_row);
        }

        // 写入字段值
        for ($i = 0; $i < count($values); $i++) {
            for ($j = 0; $j < count($values[$i]); $j++) {
                $sheet->setCellValueExplicitByColumnAndRow(
                    $j + 1,
                    $i + $data_row,
                    array_values($values[$i])[$j],
                    DataType::TYPE_STRING
                );
                // $width = mb_strlen($values[$i][$j], 'GB2312');
                $width = mb_strwidth(array_values($values[$i])[$j], 'UTF-8');
                $w[$j] = max($w[$j], $width);
            }
        }

        // 调整列宽和标题行高
        for ($i = 0; $i < count($fields); $i++) {
            $sheet->getColumnDimensionByColumn($i + 1)->setWidth(min(
                $w[$i] * self::SINGLE_CHARACTER_WIDTH,
                self::MAX_COLUMN_WIDTH
            ));
        }
        if (strpos($title, "\n")) {
            self::autofitTitleHeight($sheet, $w);
        }

        // 写excel文件
        $writer = new Xlsx($spreadsheet);
        $writer->setPreCalculateFormulas(false);
        $pathname = storage_path('app/' . md5(Str::random(40)) . '.xlsx');
        $writer->save($pathname);

        return $pathname;
    }

    /**
     * 生成 excel 文件
     *
     * @param string $title 标题
     * @param array $fields 字段数组 [字段名 => 字段中文名]
     * @param array $numericColumns 数字字段列数组 [字段名...]
     * @param array|Collection $values 字段值
     * @param integer $freezeColumnIndex 冻结字段列（0不冻结）
     * @param Closure $rowCallback 行处理程序，需返回数组 [字段名 => 字段值]
     * @return string 返回完整文件名
     */
    public static function write(
        string $title,
        array $fields,
        array $numericColumns,
        $values,
        Closure $rowCallback = null,
        int $freezeColumnIndex = 0
    ): string {
        ini_set('memory_limit', '512M');
        $spreadsheet = new Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();
        $w = []; // 每列最大宽度值

        // 设置标题
        if ($title) {
            $sheet->mergeCellsByColumnAndRow(1, 1, count($fields), 1);
            $sheet->setCellValue('A1', $title);
            $sheet->getStyle('A1')->getAlignment()->setWrapText(true);
            $sheet->getStyle('A1')->getAlignment()->setVertical('center');
        }

        // 设置行号常量
        $title_row = $title ? 1 : 0;
        $field_row = $title_row + 1;
        $data_row = $field_row + 1;

        // 写入字段名
        $i = 0;
        foreach ($fields as $field) {
            $sheet->setCellValueExplicitByColumnAndRow($i + 1, $field_row, $field, DataType::TYPE_STRING);
            $width = mb_strwidth($field, 'UTF-8');
            $w[$i] = max($width, 4 * 2); // 最小宽度 4 个汉字
            $i++;
        }

        // 冻结表头
        if ($freezeColumnIndex) {
            $sheet->freezePaneByColumnAndRow($freezeColumnIndex, $data_row);
        }

        // 添加一行的方法
        $addRow = function ($sheet, $row, $fields, $rowIndex) use ($data_row, $numericColumns, &$w) {
            $keys = array_keys($fields);
            foreach ($keys as $index => $key) {
                $value = Arr::get($row, $key) ?? Arr::get($row, $index);

                $sheet->setCellValueExplicitByColumnAndRow(
                    $index + 1,
                    $rowIndex + $data_row,
                    $value,
                    in_array($key, $numericColumns) ? DataType::TYPE_NUMERIC : DataType::TYPE_STRING
                );
                $width = mb_strwidth($value, 'UTF-8');
                $w[$index] = max($w[$index], $width);
            }
        };

        // 写入字段值
        $i = 0;
        if (is_array($values) || $values instanceof Collection) {
            foreach ($values as $row) {
                $_row = $rowCallback ? $rowCallback($i, $row) : $row;
                if ($_row) {
                    $addRow($sheet, $_row, $fields, $i);
                    $i++;
                }
            }
        } else if ($values instanceof Builder) {
            $values->chunk(100, function ($rows) use (&$i, $rowCallback, $sheet, $fields, $addRow) {
                foreach ($rows as $row) {
                    $_row = $rowCallback ? $rowCallback($i, $row) : $row;
                    if ($_row) {
                        $addRow($sheet, $_row, $fields, $i);
                        $i++;
                    }
                }
            });
        }

        // 调整列宽和标题行高
        for ($i = 0; $i < count($fields); $i++) {
            $sheet->getColumnDimensionByColumn($i + 1)->setWidth(min(
                $w[$i] * self::SINGLE_CHARACTER_WIDTH,
                self::MAX_COLUMN_WIDTH
            ));
        }
        if ($title && strpos($title, "\n")) {
            self::autofitTitleHeight($sheet, $w);
        }

        // 写excel文件
        $writer = new Xlsx($spreadsheet);
        $writer->setPreCalculateFormulas(false);
        $pathname = storage_path('app/' . md5(Str::random(40)) . '.xlsx');
        $writer->save($pathname);

        return $pathname;
    }

    private static function autofitTitleHeight($sheet, $w)
    {
        $cell = $sheet->getCell('A1');
        $cellWidth = array_sum($w);

        $lines = explode("\n", $cell->getValue());
        $cellLines = 0;
        foreach ($lines as $line) {
            $cellLines += ceil(mb_strwidth($line, 'UTF-8') / $cellWidth);
        }

        $rowDimension = $sheet->getRowDimension(1);
        $rowHeight = $rowDimension->getRowHeight();
        if ($rowHeight === -1) {
            $rowDimension->setRowHeight(self::DEFAULT_ROW_HEIGHT);
            $rowHeight = $rowDimension->getRowHeight();
        }

        $rowDimension->setRowHeight($rowHeight * ($cellLines + 1));
    }

    public static function getHighestRow(string $inputFileName, int $sheetIndex = 0): int
    {
        ini_set('memory_limit', '512M');

        $inputFileType = IOFactory::identify($inputFileName);
        $reader = IOFactory::createReader($inputFileType);

        $worksheetData = $reader->listWorksheetInfo($inputFileName);

        $totalRows = 0;
        foreach ($worksheetData as $index => $worksheet) {
            if ($index === $sheetIndex) {
                $totalRows = $worksheet['totalRows'];
                break;
            }
        }

        return $totalRows;
    }
}
