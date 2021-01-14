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

        $this->assertEquals($data[0], [
            'id' => 1,
            'name' => '张三',
            'age' => 21,
        ]);

        $this->assertEquals($data[2], [
            'id' => 3,
            'name' => '王五',
            'age' => 23,
        ]);
    }

    public function testWriteUseArray()
    {
        $pathname = ExcelService::write(
            [
                'id' => '编号',
                'name' => '姓名',
                'age' => '年龄',
            ],
            [
                ['id' => 4, 'name' => '赵六', 'age' => 24],
                ['id' => 5, 'name' => '孙七', 'age' => 25],
                ['id' => 6, 'name' => '周八', 'age' => 26],
            ],
            function (int $index, array $row) {
                return array_values($row);
            },
        );

        $this->assertFileExists($pathname);

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
                $data[] = $cells;
            },
        );

        $this->assertEquals($data[2], [
            'id' => 6,
            'name' => '周八',
            'age' => 26,
        ]);

        @unlink($pathname);
    }

    public function testWriteUseEloquentBuilder()
    {
        User::create(['name' => '张三']);
        User::create(['name' => '李四']);

        $pathname = ExcelService::write(
            [
                'id' => '编号',
                'name' => '姓名',
            ],
            User::query(),
            function (int $index, $row) {
                return [$row->id, $row->name];
            },
        );

        $this->assertFileExists($pathname);

        $data = [];
        ExcelService::read(
            $pathname,
            1,
            [
                'id' => '编号',
                'name' => '姓名',
            ],
            2,
            function (int $index, array $cells) use (&$data) {
                $data[] = $cells;
            },
        );

        $this->assertEquals($data, [
            ['id' => 1, 'name' => '张三'],
            ['id' => 2, 'name' => '李四'],
        ]);

        @unlink($pathname);
    }
}
