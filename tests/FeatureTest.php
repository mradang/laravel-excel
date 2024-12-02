<?php

namespace mradang\LaravelExcel\Test;

use Carbon\Carbon;
use mradang\LaravelExcel\Services\ExcelService;

class FeatureTest extends TestCase
{
    public function setUp(): void
    {
        parent::setUp();
    }

    public $fields = [
        'id' => '编号',
        'name' => '姓名',
        'age' => '年龄',
        'date1' => '日期1',
        'date2' => '日期2',
    ];

    public function testRead()
    {
        $pathname = __DIR__ . DIRECTORY_SEPARATOR . 'data.xlsx';
        $data = [];

        ExcelService::read(
            $pathname,
            1,
            $this->fields,
            2,
            function (int $index, array $cells) use (&$data) {
                if ($index === 0) {
                    $this->assertEquals($this->formatCells($cells), [
                        'id' => 1,
                        'name' => '张三',
                        'age' => 21,
                        'date1' => '2024-04-01',
                        'date2' => '2024-01-01',
                    ]);
                }
                $data[] = $cells;
            },
        );

        $this->assertEquals(count($data), 3);

        $this->assertEquals($this->formatCells($data[1]), [
            'id' => 2,
            'name' => '李四',
            'age' => 22,
            'date1' => '2024-05-01',
            'date2' => '2024-02-01',
        ]);

        $this->assertEquals($this->formatCells($data[2]), [
            'id' => 3,
            'name' => '王五',
            'age' => 23,
            'date1' => '2024-06-01',
            'date2' => '2024-03-01',
        ]);

        $this->assertEquals(ExcelService::getHighestRow($pathname), 4);
    }

    public function testWriteUseArray()
    {
        $values = [
            ['id' => 4, 'name' => '赵六', 'age' => 24, 'date1' => '2024-07-01', 'date2' => '2024-04-01'],
            ['id' => 5, 'name' => '孙七', 'age' => 25, 'date1' => '2024-08-01', 'date2' => '2024-05-01'],
            ['id' => 6, 'name' => '周八', 'age' => 26, 'date1' => '2024-09-01', 'date2' => '2024-06-01'],
        ];

        $pathname = ExcelService::write(
            '测试标题',
            $this->fields,
            ['id', 'age'],
            $values,
            function (int $index, array $row) {
                return array_values($row);
            },
            2,
        );

        $this->assertFileExists($pathname);
        $data = $this->readTestExcelData($pathname, 2);
        $this->assertEquals(count($data), 3);
        $this->assertEquals($data[2], $values[2]);
        @unlink($pathname);

        // 不使用回调函数处理行数据
        $pathname = ExcelService::write(
            '',
            $this->fields,
            [],
            array_map(function ($row) {
                return array_values($row);
            }, $values)
        );

        $this->assertFileExists($pathname);
        $data = $this->readTestExcelData($pathname, 1);
        $this->assertEquals(count($data), 3);
        $this->assertEquals($data[1], $values[1]);
        @unlink($pathname);
    }

    public function testWriteUseEloquentBuilder()
    {
        User::create(['name' => '张三', 'age' => 33, 'date1' => '2024-07-01', 'date2' => '2024-04-01']);
        User::create(['name' => '李四', 'age' => 34, 'date1' => '2024-08-01', 'date2' => '2024-05-01']);

        $pathname = ExcelService::write(
            '',
            $this->fields,
            ['id', 'age'],
            User::query(),
            function (int $index, $row) {
                return [$row->id, $row->name, $row->age, $row->date1, $row->date2];
            },
        );

        $this->assertFileExists($pathname);
        $data = $this->readTestExcelData($pathname, 1);
        $this->assertEquals(count($data), 2);
        $this->assertEquals($data[1], [
            'id' => 2,
            'name' => '李四',
            'age' => 34,
            'date1' => '2024-08-01',
            'date2' => '2024-05-01',
        ]);
        @unlink($pathname);
    }

    private function readTestExcelData(string $pathname, int $fieldRow)
    {
        $data = [];
        ExcelService::read(
            $pathname,
            $fieldRow,
            $this->fields,
            $fieldRow + 1,
            function (int $index, array $cells) use (&$data) {
                $data[] = $this->formatCells($cells);
            },
        );

        return $data;
    }

    private function formatCells(array $cells): array
    {
        foreach ($cells as $key => $cell) {
            if (is_a($cell, 'DateTime')) {
                $cells[$key] = Carbon::parse($cell)->format('Y-m-d');
            }
        }
        return $cells;
    }
}
