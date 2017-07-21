<?php
namespace moxuandi\phpexcel;

use Yii;
use yii\base\InvalidConfigException;
use yii\base\InvalidParamException;
use yii\base\Widget;
use yii\helpers\ArrayHelper;

class Phpexcel extends Widget
{
    public $mode;  // 模式, 导出(export) or 导入(import)
    public $format;  // Excel 的版本, 值有'Excel2007', 'Excel5', 'Excel2003XML', 'OOCalc', 'SYLK', 'Gnumeric', 'HTML', 'CSV'

    // 导出(export)的参数:
    public $models;  // 数据提供者, eg: Post::find()->all()
    public $columns = [];  // 从模型中获取属性, 未设置则获取该模型的所有属性
    public $headers = [];  // 设置第一行的标题栏, 未设置则获取该模型的属性标签
    public $setFirstTitle;  // 是否在第一行设置标题行
    public $asAttachment;  // 是否下载导出结果, 为 true 时则仅下载, 为 false 时仅保存结果到服务器,
    public $fileName;  // 导出的文件名
    public $savePath;  // 保存到服务器的路径, 仅 asAttachment=false 时生效
    public $isMultipleSheet;  // 是否同时导出多个表, 导出多个表时必须为 true
    public $formatter;
    public $properties = [];

    // 导入(import)的参数:
    public $importFile;  // 导入的文件, 可以是单文件也可以是多文件的数组
    public $setIndexSheetByName;  // 如果Excel文件中有多个表, 是否以表名(eg:sheet1,sheet2)作为键名, 为 false 时使用数字(eg:0,1,2)
    public $setFirstRecordAsKeys;  // 将Excel文件中的第一行记录设置为每行数组的键, 为 false 时使用Excel的字母列(eg:A,B,C)
    public $getOnlyRecordByIndex = [];  //
    public $leaveRecordByIndex = [];  //
    public $getOnlySheet;  // 当Excel文件中有多个表时, 指定仅获取某个表(eg:sheet1),

    public function __construct(array $config = [])
    {
        $config = ArrayHelper::merge([
            'mode' => 'export',
            'format' => 'Excel2007',
            'models' => '',
            //'columns' => [],
            //'headers' => [],
            'setFirstTitle' => true,
            'asAttachment' => true,
            'fileName' => 'excel.xls',
            'savePath' => 'uploads/excel/',
            'isMultipleSheet' => false,
            //'formatter' => '',  // 不能设置, 否则无法格式化导出内容
            //'properties' => [],

            'importFile' => '',
            'setIndexSheetByName' => false,
            'setFirstRecordAsKeys' => true,
            //'getOnlyRecordByIndex' => [],
            //'leaveRecordByIndex' => [],
            //'getOnlySheet' => '',
        ], $config);
        parent::__construct($config);
    }

    /**
     * 解决 mode='import' 时无法返回数组的错误
     */
    public static function widget($config=[])
    {
        if($config['mode'] == 'import'){
            $config['class'] = get_called_class();
            $widget = yii::createObject($config);
            return $widget->run();
        }
        return parent::widget($config);
    }

    public function run()
    {
        if($this->mode === 'export'){
            return self::Export();
        }elseif($this->mode === 'import'){
            return self::Import();
        }else{
            return '访问错误！';
        }
    }

    /**
     * 导出操作
     * @return bool|string
     * @throws InvalidConfigException
     */
    public function Export()
    {
        $sheet = new \PHPExcel();
        if(!isset($this->models)){
            throw new InvalidConfigException('Config models must be set.');
        }
        if(isset($this->properties)){
            self::properties($sheet, $this->properties);
        }
        if($this->isMultipleSheet){
            $index = 0;
            $worksheet = [];
            foreach($this->models as $title=>$model){
                $sheet->createSheet($index);
                $sheet->getSheet($index)->setTitle($title);
                $worksheet[$index] = $sheet->getSheet($index);
                $columns = isset($this->columns[$title]) ? $this->columns[$title] : [];
                $headers = isset($this->headers[$title]) ? $this->headers[$title] : [];
                self::executeColumns($worksheet[$index], $model, self::populateColumns($columns), $headers);
                $index++;
            }
        }else{
            $worksheet = $sheet->getActiveSheet();
            self::executeColumns($worksheet, $this->models, isset($this->columns) ? self::populateColumns($this->columns) : [], isset($this->headers) ? $this->headers : []);
        }
        $objectwriter = \PHPExcel_IOFactory::createWriter($sheet, $this->format);
        if($this->asAttachment){
            header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
            header('Content-Disposition: attachment;filename="' . $this->fileName .'"');
            header('Cache-Control: max-age=0');
            $objectwriter->save('php://output');
            exit();
        }elseif(self::createDir($this->savePath)){
            $path = $this->savePath . $this->fileName;
            $objectwriter->save($path);
            return $path;
        }else{
            return false;
        }
    }

    /**
     * 导入操作, 仅返回处理后的数组, 不执行导入到数据库操作
     * @return array
     */
    public function Import()
    {
        if(is_array($this->importFile)){
            $datas = [];
            foreach($this->importFile as $key=>$file){
                $datas[$key] = self::readFile($file);
            }
            return $datas;
        }else{
            return self::readFile($this->importFile);
        }
    }

    /**
     * 导出操作: 从模型中获取数据
     * @param null $activeSheet
     * @param $models
     * @param array $columns
     * @param array $headers
     */
    private function executeColumns(&$activeSheet = null, $models, $columns=[], $headers=[])
    {
        if($activeSheet === null){
            $activeSheet = $this->activeSheet;  // 不理解
        }
        $hasHeader = false;
        $row = 1;
        $char = 26;
        foreach($models as $model){
            if(empty($columns)){
                $columns = $model->attributes();
            }
            if($this->setFirstTitle && !$hasHeader){
                $isPlus = false;
                $colPlus = 0;
                $colNum = 1;
                foreach($columns as $key=>$column){
                    $col = '';
                    if($colNum > $char){
                        $colPlus += 1;
                        $colNum = 1;
                        $isPlus = true;
                    }
                    if($isPlus){
                        $col .= chr(64 + $colPlus);
                    }
                    $col .= chr(64 + $colNum);
                    $header = '';
                    if(is_array($column)){
                        if(isset($column['header'])){
                            $header = $column['header'];
                        }elseif(isset($column['attribute'])){
                            if(isset($headers[$column['attribute']])){
                                $header = $headers[$column['attribute']];
                            }else{
                                $header = $model->getAttributeLabel($column['attribute']);
                            }
                        }
                    }else{
                        $header = $model->getAttributeLabel($column);
                    }
                    $activeSheet->setCellValue($col . $row, $header);
                    $colNum++;
                }
                $hasHeader = true;
                $row++;
            }else{
                $isPlus = false;
                $colPlus = 0;
                $colNum = 1;
                foreach($columns as $key=>$column){
                    $col = '';
                    if($colNum > $char){
                        $colPlus += 1;
                        $colNum = 1;
                        $isPlus = true;
                    }
                    if($isPlus){
                        $col .= chr(64 + $colPlus);
                    }
                    $col .= chr(64 + $colNum);
                    $header = '';
                    if(is_array($column)){
                        $column_value = self::executeGetColumnData($model, $column);
                    }else{
                        $column_value = self::executeGetColumnData($model, ['attribute'=>$column]);
                    }
                    $activeSheet->setCellValue($col . $row, $column_value);
                    $colNum++;
                }
                $row++;
            }
        }
    }

    /**
     * 导出操作: 获取每一列的值
     * @param $model
     * @param array $params
     * @return mixed|null|string
     */
    private function executeGetColumnData($model, $params=[])
    {
        $value = null;
        if(isset($params['value']) && $params['value'] !== null){
            if(is_string($params['value'])){
                $value = ArrayHelper::getValue($model, $params['value']);
            }else{
                $value = call_user_func($params['value'], $model, $this);
            }
        }elseif(isset($params['attribute']) && $params['attribute'] !== null){
            $value = ArrayHelper::getValue($model, $params['attribute']);
        }
        if(isset($params['format']) && $params['format'] !== null){
            $value = self::formatter()->format($value, $params['format']);
        }
        return $value;
    }

    /**
     * 导出操作: Populating columns for checking the column is string or array. if is string this will be checking have a formatter or header.
     * @param array $columns
     * @return array
     */
    private function populateColumns($columns=[])
    {
        $_columns = [];
        foreach($columns as $key=>$value){
            if(is_string($value)){
                $value_log = explode(':', $value);
                $_columns[$key] = ['attribute'=>$value_log[0]];
                if(isset($value_log[1]) && $value_log[1] !== null){
                    $_columns[$key]['format'] = $value_log[1];
                }
                if(isset($value_log[2]) && $value_log[2] !== null){
                    $_columns[$key]['header'] = $value_log[2];
                }
            }else{
                if(!isset($value['attribute']) && !isset($value['value'])){
                    throw new InvalidParamException('Attribute or Value must be defined.');
                }else{
                    $_columns[$key] = $value;
                }
            }

        }
        return $_columns;
    }

    /**
     * 导出操作: Setting properties for excel file
     * @param $objectExcel
     * @param array $properties
     */
    private function properties(&$objectExcel, $properties=[])
    {
        foreach($properties as $key=>$value){
            $keyname = 'set' . ucfirst($key);
            $objectExcel->getProperties()->{$keyname}($value);
        }
    }

    /**
     * 导入操作: 读取 Excel 文件的内容, 返回处理后的数组
     * @param string|array $fileName
     * @return array
     */
    private function readFile($fileName)
    {
        if(!isset($this->format) || $this->format == ''){
            $this->format = \PHPExcel_IOFactory::identify($fileName);
        }
        $objectreader = \PHPExcel_IOFactory::createReader($this->format);
        $objectPhpExcel = $objectreader->load($fileName);

        $sheetCount = $objectPhpExcel->getSheetCount();
        $sheetDatas = [];

        if($sheetCount > 1){
            foreach($objectPhpExcel->getSheetNames() as $sheetIndex=>$sheetName){
                if(isset($this->getOnlySheet) && $this->getOnlySheet != null){
                    if(!$objectPhpExcel->getSheetByName($this->getOnlySheet)){
                        return $sheetDatas;
                    }
                    $objectPhpExcel->setActiveSheetIndexByName($this->getOnlySheet);
                    $indexed = $this->getOnlySheet;
                    $sheetDatas[$indexed] = $objectPhpExcel->getActiveSheet()->toArray(null, true, true, true);
                    if($this->setFirstRecordAsKeys){
                        $sheetDatas[$indexed] = self::executeArrayLabel($sheetDatas[$indexed]);
                    }
                    if(!empty($this->getOnlyRecordByIndex)){
                        $sheetDatas[$indexed] = self::executeGetOnlyRecords($sheetDatas[$indexed], $this->getOnlyRecordByIndex);
                    }
                    if(!empty($this->leaveRecordByIndex)){
                        $sheetDatas[$indexed] = self::executeLeaveRecords($sheetDatas[$indexed], $this->leaveRecordByIndex);
                    }
                    return $sheetDatas[$indexed];
                }else{
                    $objectPhpExcel->setActiveSheetIndexByName($sheetName);
                    $indexed = $this->setIndexSheetByName==true ? $sheetName : $sheetIndex;
                    $sheetDatas[$indexed] = $objectPhpExcel->getActiveSheet()->toArray(null, true, true, true);
                    if($this->setFirstRecordAsKeys){
                        $sheetDatas[$indexed] = self::executeArrayLabel($sheetDatas[$indexed]);
                    }
                    if(!empty($this->getOnlyRecordByIndex) && isset($this->getOnlyRecordByIndex[$indexed]) && is_array($this->getOnlyRecordByIndex[$indexed])){
                        $sheetDatas = self::executeGetOnlyRecords($sheetDatas, $this->getOnlyRecordByIndex[$indexed]);
                    }
                    if(!empty($this->leaveRecordByIndex) && isset($this->leaveRecordByIndex[$indexed]) && is_array($this->leaveRecordByIndex[$indexed])){
                        $sheetDatas[$indexed] = self::executeLeaveRecords($sheetDatas[$indexed], $this->leaveRecordByIndex[$indexed]);
                    }
                }
            }
        }else{
            // 以数组的形式返回 excel 表格中的数据, eg: [1=>['A'=>'a', 'B'=>'b', 'C'=>'c'], 2=>['A'=>'aa', 'B'=>'bb', 'C'=>'cc'], 3=>['A'=>'aaa', 'B'=>'bbb', 'C'=>'ccc']];
            $sheetDatas = $objectPhpExcel->getActiveSheet()->toArray(null, true, true, true);
            if($this->setFirstRecordAsKeys){
                $sheetDatas = self::executeArrayLabel($sheetDatas);
            }
            if(!empty($this->getOnlyRecordByIndex)){
                $sheetDatas = self::executeGetOnlyRecords($sheetDatas, $this->getOnlyRecordByIndex);
            }
            if(!empty($this->leaveRecordByIndex)){
                $sheetDatas = self::executeLeaveRecords($sheetDatas, $this->leaveRecordByIndex);
            }
        }
        return $sheetDatas;
    }

    /**
     * 导入操作: Setting label or keys on every record if setFirstRecordAsKeys is true.
     * @param array $sheetData
     * @return array
     */
    private function executeArrayLabel($sheetData)
    {
        $keys = ArrayHelper::remove($sheetData, '1');  // 从数组移除键=1并返回该键的值
        $newData = [];
        foreach($sheetData as $v){
            $newData[] = array_combine($keys, $v);  //合并两个数组来创建一个新数组, $keys为键名, $v为键值
        }
        return $newData;
    }

    /**
     * 导入操作: Read record with same index number.
     * @param array $sheetData
     * @param array $index
     * @return array
     */
    private function executeGetOnlyRecords($sheetData=[], $index=[])
    {
        foreach($sheetData as $key=>$data){
            if(!in_array($key, $index)){
                unset($sheetData[$key]);
            }
        }
        return $sheetData;
    }

    /**
     * 导入操作: Leave record with same index number.
     * @param array $sheetData
     * @param array $index
     * @return array
     */
    private function executeLeaveRecords($sheetData=[], $index=[])
    {
        foreach($sheetData as $key=>$data){
            if(in_array($key, $index)){
                unset($sheetData[$key]);
            }
        }
        return $sheetData;
    }

    /**
     * Formatter for i18n.
     * @return \yii\i18n\Formatter
     */
    private function formatter()
    {
        if(!isset($this->formatter)){
            $this->formatter = Yii::$app->getFormatter();
        }
        return $this->formatter;
    }

    /**
     * 创建文件夹
     * @param string $dirname
     * @return bool
     */
    private function createDir($dirname)
    {
        if(!file_exists($dirname) && !mkdir($dirname, 0777, true)){
            return false;
        }elseif(!is_writable($dirname)){
            return false;
        }
        return true;
    }
}
