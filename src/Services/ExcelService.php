<?php

namespace mradang\LaravelExcel\Services;

use Closure;
use Illuminate\Database\Eloquent\Builder;
use Illuminate\Support\Arr;
use Illuminate\Support\Collection;
use Illuminate\Support\Str;
use PhpOffice\PhpSpreadsheet\Cell\DataType;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\RichText\RichText;
use PhpOffice\PhpSpreadsheet\Shared\Date;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

class ExcelService
{
    const DEFAULT_ROW_HEIGHT = 14.4;

    const SINGLE_CHARACTER_WIDTH = 1.1;

    const MAX_COLUMN_WIDTH = 100;

    /**
     * 读取 excel 文件
     *
     * @param  string  $pathname  文件绝对路径
     * @param  int  $fieldRow  字段名行号
     * @param  array  $fields  字段数组 [字段名 => 字段中文名]
     * @param  int  $firstDataRow  首行数据行号
     * @param  Closure  $callback  行处理程序
     */
    public static function read(
        string $pathname,
        int $fieldRow,
        array $fields,
        int $firstDataRow,
        Closure $callback
    ): void {
        ini_set('memory_limit', '512M');

        $inputFileType = IOFactory::identify($pathname);
        $reader = IOFactory::createReader($inputFileType);

        $spreadsheet = $reader->load($pathname);
        $spreadsheet->setActiveSheetIndex(0);

        $columns = null;

        foreach ($spreadsheet->getActiveSheet()->getRowIterator() as $index => $row) {
            $cellIterator = $row->getCellIterator();
            $cellIterator->setIterateOnlyExistingCells(false);

            $cells = [];
            foreach ($cellIterator as $cell) {
                $value = $cell->getValue();
                if (Date::isDateTime($cell)) {
                    $cells[] = Date::excelToDateTimeObject($value);
                } elseif ($value instanceof RichText) {
                    $cells[] = $value->getPlainText();
                } else {
                    $cells[] = $value;
                }
            }

            if ($index === $fieldRow) {
                $columns = self::handleFieldRow($fields, $cells);
            } elseif ($index >= $firstDataRow && $columns) {
                $callback($index - $firstDataRow, self::handleDataRow($columns, $cells));
            }
        }
    }

    private static function handleFieldRow(array $fields, array $cells)
    {
        $columns = [];
        foreach ($fields as $key => $value) {
            $columns[] = [
                'field' => $key,
                'column' => array_search($value, $cells),
            ];
        }

        return $columns;
    }

    private static function handleDataRow(array $columns, array $cells)
    {
        $ret = [];
        foreach ($columns as $col) {
            if ($col['column'] !== false) {
                $ret[$col['field']] = Arr::get($cells, $col['column']) ?? '';
            } else {
                $ret[$col['field']] = '';
            }
        }

        return $ret;
    }

    /**
     * 生成 excel 文件
     *
     * @param  string  $title  标题（空字符串时第一行开始字段名）
     * @param  array  $fields  字段数组 [字段名 => 字段中文名]
     * @param  array  $numericColumns  数字字段列数组 [字段名...]
     * @param  array|Collection  $values  字段值
     * @param  int  $freezeColumnIndex  冻结字段列（0不冻结）
     * @param  Closure  $rowCallback  行处理程序，需返回数组 [字段名 => 字段值]
     * @return string 返回完整文件名
     */
    public static function write(
        string $title,
        array $fields,
        array $numericColumns,
        $values,
        ?Closure $rowCallback = null,
        int $freezeColumnIndex = 0,
    ): string {
        ini_set('memory_limit', '512M');

        $spreadsheet = new Spreadsheet;
        $sheet = $spreadsheet->getActiveSheet();
        $w = []; // 每列最大宽度值

        // 设置标题
        if ($title) {
            $sheet->mergeCells([1, 1, count($fields), 1]);
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
            $sheet->setCellValueExplicit([$i + 1, $field_row], $field, DataType::TYPE_STRING);
            $width = mb_strwidth($field, 'UTF-8');
            $w[$i] = max($width, 4 * 2); // 最小宽度 4 个汉字
            $i++;
        }

        // 冻结表头
        if ($freezeColumnIndex) {
            $sheet->freezePane([$freezeColumnIndex, $data_row]);
        }

        // 添加一行的方法
        $addRow = function ($row, $rowIndex) use ($data_row, $numericColumns, $sheet, $fields, &$w) {
            $keys = array_keys($fields);
            foreach ($keys as $index => $key) {
                $value = Arr::get($row, $key) ?? Arr::get($row, $index);

                $sheet->setCellValueExplicit(
                    [$index + 1, $rowIndex + $data_row],
                    $value,
                    in_array($key, $numericColumns) && is_numeric($value) ? DataType::TYPE_NUMERIC : DataType::TYPE_STRING
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
                    $addRow($_row, $i);
                    $i++;
                }
            }
        } elseif ($values instanceof Builder) {
            $values->chunk(100, function ($rows) use (&$i, $rowCallback, $addRow) {
                foreach ($rows as $row) {
                    $_row = $rowCallback ? $rowCallback($i, $row) : $row;
                    if ($_row) {
                        $addRow($_row, $i);
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
        $pathname = storage_path('app/'.md5(Str::random(40)).'.xlsx');
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

        return Arr::get($worksheetData, $sheetIndex.'.totalRows') ?: 0;
    }
}
