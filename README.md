# [PHPExcel for Yii2](http://phpexcel.codeplex.com/)
========
基于 PHPExcel 写的 Yii2 扩展，用于将数据表导出为Excel表格，或将Excel表格的内存导入到数据库。

安装:
------------
使用 [composer](http://getcomposer.org/download/) 下载:
```
composer require --prefer-dist moxuandi/yii2-phpexcel:"*"
composer require --prefer-dist moxuandi/yii2-phpexcel:"dev-master"
```


导出(export):
-----
参数:
```
// 公用参数:
$mode;  string  // 模式, 导出(export) or 导入(import)
$format;  string  // Excel 的版本, 值有'Excel2007', 'Excel5', 'Excel2003XML', 'OOCalc', 'SYLK', 'Gnumeric', 'HTML', 'CSV'

// 导出(export)的参数:
$models;  // 数据提供者, eg: Post::find()->all()
$columns = [];  array  // 从模型中获取属性, 未设置则获取该模型的所有属性
$headers = [];  array  // 设置第一行的标题栏, 未设置则获取该模型的属性标签
$setFirstTitle;  boolean  // 是否在第一行设置标题行
$asAttachment;  boolean  // 是否下载导出结果, 为 true 时则仅下载, 为 false 时仅保存结果到服务器,
$fileName;  // 导出的文件名
$savePath;  string  // 保存到服务器的路径, 仅 asAttachment=false 时生效
$isMultipleSheet;  boolean  // 是否同时导出多个表, 导出多个表时必须为 true
$formatter;
$properties = [];

// 导入(import)的参数:
$importFile;  string|array  // 导入的文件, 可以是单文件也可以是多文件的数组
$setIndexSheetByName;  boolean  // 如果Excel文件中有多个表, 是否以表名(eg:sheet1,sheet2)作为键名, 为 false 时使用数字(eg:0,1,2)
$setFirstRecordAsKeys;  boolean  // 将Excel文件中的第一行记录设置为每行数组的键, 为 false 时使用Excel的字母列(eg:A,B,C)
$getOnlyRecordByIndex = [];  //
$leaveRecordByIndex = [];  //
$getOnlySheet;  string  // 当Excel文件中有多个表时, 指定仅获取某个表(eg:sheet1),
```

用法示例:
```
// 导出单个表, 并下载导出的Excel文件
Phpexcel::widget([
    'mode' => 'export',  // 必须
    'models' => Upload::find()->all(),  // 必须
    'asAttachment' => true,  // 默认值, 可忽略
]);

// 导出单个表, 并将文件保存到服务器, 返回导出后的Excel文件路径
$url = Phpexcel::widget([
    'mode' => 'export',  // 必须
    'models' => Upload::find()->all(),  // 必须
    'asAttachment' => false,  // 为 false 时保存到服务器
    'fileName' => time() . '.xls',  // 默认为:'excel.xls'
    'savePath' => 'uploads/excel', // 默认为:'uploads/excel/'
]);
// return: $url = 'uploads/excel1500597563.xls';

// 导出单个表中指定的列
Phpexcel::widget([
    'mode' => 'export',  // 必须
    'models' => Upload::find()->all(),  // 必须
    'columns' => ['id', 'real_name', 'file_name', 'file_size'],
    // 'headers'数组中的键名必须是'columns'数组的值, 否则无效
    'headers' => ['id'=>'ID', 'real_name'=>'源文件名', 'file_name'=>'新文件路径', 'file_size'=>'大小(B)'],
]);

// 导出多个表, 一个Excel文件多个表
Phpexcel::widget([
    'mode' => 'export',  // 必须
    'isMultipleSheet' => true,  // 导出多个表时, 必须为 true
    'models' => [
        'sheet1' => Upload::find()->all(),
        'sheet2' => Article::find()->all(),
        'sheet3' => Effect::find()->all(),
    ],
    //指定导出的列
    'columns' => [
        'sheet1' => ['id', 'real_name', 'file_name', 'file_size'],
        'sheet2' => ['id', 'title', 'sort'],
        'sheet3' => ['id', 'title', 'summary', 'method', 'demo_url'],
    ],
    'headers' => [
        'sheet1' => ['id'=>'ID', 'real_name'=>'源文件名', 'file_name'=>'新文件路径', 'file_size'=>'大小(B)'],
        'sheet2' => ['id'=>'ID', 'title'=>'文章标题', 'sort'=>'排序值'],
        'sheet3' => ['id'=>'ID', 'title'=>'插件标题', 'summary'=>'插件介绍', 'demo_url'=>'演示地址'],
    ],
]);

// 【失败】更强的导出功能
Phpexcel::widget([
    'mode' => 'export',  // 必须
    'models' => Upload::find()->all(),  // 必须
    'columns' => [
        'id',
        'real_name',
        'file_name',
        [
            'attribute' => 'file_size',
            'header' => '文件大小',
            'format' => 'text',
            'value' => function($model){
                return Helper::byteFormat($model->file_size);
            }
        ],
        'created_at:datetime',
        [
            'attribute' => 'updated_at',
            'format' => 'date'
        ]
    ],
    'headers' => ['id'=>'ID', 'real_name'=>'源文件名', 'file_name'=>'新文件路径', 'file_size'=>'大小(B)'],
]);
```

导入(import):
-----
```
// 导入一个Excel文件(默认值)
$data = Phpexcel::widget([
    'mode' => 'import',  // 必须
    'importFile' => 'uploads/excel/excel.xls',  // 必须, 要导入的Excel文件
    //'setIndexSheetByName' => false,  // 默认为 false
    //'setFirstRecordAsKeys' => true,  // 默认为 true
    //'getOnlyRecordByIndex' => [],  // 默认为空
    //'leaveRecordByIndex' => [],  // 默认为空
    //'getOnlySheet' => '',  // 默认为空
]);

// 导入一个多表Excel文件, 以表名作为索引
$data = Phpexcel::widget([
    'mode' => 'import',  // 必须
    'importFile' => 'uploads/excel/excel.xls',  // 必须, 要导入的Excel文件
    'setIndexSheetByName' => true,  // 默认为 false
]);

// 导入一个多表Excel文件中指定的一个表
$data = Phpexcel::widget([
    'mode' => 'import',  // 必须
    'importFile' => 'uploads/excel/excel.xls',  // 必须, 要导入的Excel文件
    'getOnlySheet' => 'sheet1',  // 默认为空
]);
```
