<?php

namespace mradang\LaravelExcel\Services;

use Box\Spout\Reader\Common\Creator\ReaderEntityFactory;
use Box\Spout\Writer\Common\Creator\WriterEntityFactory;
use Closure;
use Illuminate\Database\Eloquent\Builder;
use Illuminate\Support\Arr;

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
            $ret[$col['field']] = optional(Arr::get($cells, $col['column']))->getValue() ?? '';
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

        if (is_array($values)) {
            foreach ($values as $index => $value) {
                $row = $rowCallback($index, $value);
                $writer->addRow(WriterEntityFactory::createRowFromArray($row));
            }
        } else if ($values instanceof Builder) {
            $index = 0;
            $values->chunk(100, function ($rows) use (&$index, $rowCallback, $writer) {
                foreach ($rows as $value) {
                    $row = $rowCallback($index, $value);
                    $writer->addRow(WriterEntityFactory::createRowFromArray($row));
                    $index++;
                }
            });
        }

        $writer->close();
        return $pathname;
    }
}
