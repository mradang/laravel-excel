# laravel-excel

仅支持 XLSX 文件

## 安装

```shell
composer require mradang/laravel-options -vvv
```

## 读取

|编号|姓名|年龄|
|---|---|---|
|1|张三|21|
|2|李四|22|
|3|王五|23|

```php
ExcelService::read(
    $pathname, // excel 文件绝对路径
    1, // 字段名行号
    [
        // 字段定义
        'id' => '编号',
        'name' => '姓名',
        'age' => '年龄',
    ],
    2, // 数据开始的行号
    function (array $cells) {
        // 行数据处理函数，每行调用一次
        /**
         * $cells = [
         *   'id' => 1,
         *   'name' => '张三',
         *   'age' => 21,
         * ];
         */
    },
);
```

## 写入

$values 参数支持 Illuminate\Database\Eloquent\Builder 和 array

当使用 Builder 类型时，会调用实例的 chunk 方法，分块读取数据

```php
$pathname = ExcelService::write(
    [
        // 字段定义
        'id' => '编号',
        'name' => '姓名',
        'age' => '年龄',
    ],
    [
        // 数据
        ['id' => 4, 'name' => '赵六', 'age' => 24],
        ['id' => 5, 'name' => '孙七', 'age' => 25],
        ['id' => 6, 'name' => '周八', 'age' => 26],
    ],
    function (int $index, $row) {
        // 行数据处理函数，每行调用一次，返回字段值数组
        return array_values($row);
    },
);

```