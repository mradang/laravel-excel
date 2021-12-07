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

        $this->assertEquals(4, ExcelService::getHighestRow($pathname));
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
            $fields,
            $values,
            function (int $index, array $row) {
                return array_values($row);
            },
        );

        $this->assertFileExists($pathname);
        $this->assertEquals(4, ExcelService::getHighestRow($pathname));

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

        // 不试用回调函数处理行数据
        $pathname = ExcelService::write(
            $fields,
            array_map(function($row) {
                return array_values($row);
            }, $values)
        );

        $this->assertFileExists($pathname);
        $this->assertEquals(4, ExcelService::getHighestRow($pathname));

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

        $this->assertEquals($data[1], [
            'id' => 5,
            'name' => '孙七',
            'age' => 25,
        ]);

        @unlink($pathname);

        // 用 PhpSpreadsheet 生成
        $pathname = ExcelService::makeUsePhpSpreadsheet('测试标题', $fields, $values, 1);
        $this->assertFileExists($pathname);
        $this->assertEquals(5, ExcelService::getHighestRow($pathname));
        @unlink($pathname);
    }

    public function testWriteUseEloquentBuilder()
    {
        User::create(['name' => '张三']);
        User::create(['name' => '李四']);

        $fields = [
            'id' => '编号',
            'name' => '姓名',
        ];

        $pathname = ExcelService::write(
            $fields,
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
            $fields,
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

        // 用 PhpSpreadsheet 生成
        $pathname = ExcelService::writeUsePhpSpreadsheet(
            '测试标题',
            $fields,
            ['id'],
            User::query(),
            function (int $index, $row) {
                return [$row->id, $row->name];
            },
        );
        $this->assertFileExists($pathname);
        $this->assertEquals(4, ExcelService::getHighestRow($pathname));
        @unlink($pathname);
    }
}
