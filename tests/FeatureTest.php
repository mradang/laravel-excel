<?php

namespace mradang\LaravelExcel\Test;

use mradang\LaravelExcel\Services\ExcelService;

class FeatureTest extends TestCase
{
    public function setUp(): void
    {
        parent::setUp();
    }

    public function testRead()
    {
        $pathname = __DIR__ . DIRECTORY_SEPARATOR . 'data.xlsx';
        $data = [];

        ExcelService::read(
            $pathname,
            1,
            [
                'id' => '编号',
                'name' => '姓名',
                'age' => '年龄',
            ],
            2,
            function (int $index, array $cells) use (&$data) {
                if ($index === 0) {
                    $this->assertEquals($cells, [
                        'id' => 1,
                        'name' => '张三',
                        'age' => 21,
                    ]);
                }
                $data[] = $cells;
            },
        );

        $this->assertEquals(count($data), 3);

        $this->assertEquals($data[1], [
            'id' => 2,
            'name' => '李四',
            'age' => 22,
        ]);

        $this->assertEquals($data[2], [
            'id' => 3,
            'name' => '王五',
            'age' => 23,
        ]);
    }

    public function testWriteUseArray()
    {
        $fields = [
            'id' => '编号',
            'name' => '姓名',
            'age' => '年龄',
        ];

        $values = [
            ['id' => 4, 'name' => '赵六', 'age' => 24],
            ['id' => 5, 'name' => '孙七', 'age' => 25],
            ['id' => 6, 'name' => '周八', 'age' => 26],
        ];

        $pathname = ExcelService::write(
            '测试标题',
            $fields,
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
            $fields,
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
        User::create(['name' => '张三', 'age' => 33]);
        User::create(['name' => '李四', 'age' => 34]);

        $fields = [
            'id' => '编号',
            'name' => '姓名',
            'age' => '年龄',
        ];

        $pathname = ExcelService::write(
            '',
            $fields,
            ['id', 'age'],
            User::query(),
            function (int $index, $row) {
                return [$row->id, $row->name, $row->age];
            },
        );

        $this->assertFileExists($pathname);
        $data = $this->readTestExcelData($pathname, 1);
        $this->assertEquals(count($data), 2);
        $this->assertEquals($data[1], [
            'id' => 2,
            'name' => '李四',
            'age' => 34,
        ]);
        @unlink($pathname);
    }

    private function readTestExcelData(string $pathname, int $fieldRow)
    {
        $data = [];
        ExcelService::read(
            $pathname,
            $fieldRow,
            [
                'id' => '编号',
                'name' => '姓名',
                'age' => '年龄',
            ],
            $fieldRow + 1,
            function (int $index, array $cells) use (&$data) {
                $data[] = $cells;
            },
        );

        return $data;
    }
}
