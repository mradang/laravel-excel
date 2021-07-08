<?php

namespace mradang\LaravelExcel\Services;

use Box\Spout\Reader\Common\Creator\ReaderEntityFactory;
use Box\Spout\Writer\Common\Creator\WriterEntityFactory;
use Closure;
use Illuminate\Database\Eloquent\Builder;
use Illuminate\Support\Arr;
use Illuminate\Support\Collection;

class ExcelService
{
    public static function read(
        string $pathname,
        int $fieldRow,
        array $fields,
        int $firstDataRow,
        Closure $callback
    ) {
        ini_set('memory_limit', '512M');

        $reader = ReaderEntityFactory::createReaderFromFile($pathname);
        $reader->open($pathname);

        $columns = null;

        foreach ($reader->getSheetIterator() as $sheet) {
            if ($sheet->getIndex() === 0) {
                foreach ($sheet->getRowIterator() as $index => $row) {
                    if ($index === $fieldRow) {
                        $cells = array_map(function ($item) {
                            return $item->getValue();
                        }, $row->getCells());
                        $diff = array_diff($fields, $cells);
                        throw_if(count($diff) > 0, 'RuntimeException', sprintf(
                            '数据文件(%s)缺少必要列：%s',
                            $pathname,
                            implode(',', $diff),
                        ));
                        $columns = self::handleFieldRow($fields, $row->getCells());
                    } elseif ($index >= $firstDataRow && $columns) {
                        $callback($index - $firstDataRow, self::handleDataRow($columns, $row->getCells()));
                    }
                }
                break;
            }
        }

        $reader->close();
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
                $ret[$col['field']] = optional(Arr::get($cells, $col['column']))->getValue() ?? '';
            } else {
                $ret[$col['field']] = '';
            }
        }
        return $ret;
    }

    public static function write(array $fields, $values = null, Closure $rowCallback = null)
    {
        ini_set('memory_limit', '512M');

        $pathname = tempnam(sys_get_temp_dir(), '') . '.xlsx';
        $writer = WriterEntityFactory::createXLSXWriter();
        $writer->openToFile($pathname);

        $writer->addRow(WriterEntityFactory::createRowFromArray($fields));

        $addRow = function($writer, $row, $fields) {
            $_row = [];
            $keys = array_keys($fields);
            foreach ($keys as $index => $key) {
                $_row[$key] = Arr::get($row, $key) ?? Arr::get($row, $index);
            }
            $writer->addRow(WriterEntityFactory::createRowFromArray($_row));
        };

        if (is_array($values) || $values instanceof Collection) {
            foreach ($values as $index => $value) {
                $row = $rowCallback ? $rowCallback($index, $value) : $value;
                $addRow($writer, $row, $fields);
            }
        } else if ($values instanceof Builder) {
            $index = 0;
            $values->chunk(100, function ($rows) use (&$index, $rowCallback, $writer, $addRow, $fields) {
                foreach ($rows as $value) {
                    $row = $rowCallback ? $rowCallback($index, $value) : $value;
                    $addRow($writer, $row, $fields);
                    $index++;
                }
            });
        }

        $writer->close();
        return $pathname;
    }
}
