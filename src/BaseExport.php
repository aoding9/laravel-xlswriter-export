<?php
/**
 * @Desc 导出基类
 * @User yangyang
 * @Date 2022/8/24 10:56
 */

namespace Aoding9\Laravel\Xlswriter\Export;

use Exception;

// use Illuminate\Database\Eloquent\Collection;
use Illuminate\Database\Eloquent\Builder;
use Illuminate\Support\Collection;
use Illuminate\Database\Eloquent\Model;
use Vtiful\Kernel\Excel;
use Vtiful\Kernel\Format;

abstract class BaseExport {
    public $header = [];
    public $fileName = '文件名';
    public $tableTitle = '表名';
    /**
     * @var Collection
     */
    public $data;
    
    public function getTmpDir(): string {
        $tmp = ini_get('upload_tmp_dir');
        
        if ($tmp !== false && file_exists($tmp)) {
            return realpath($tmp);
        }
        
        return realpath(sys_get_temp_dir());
    }
    
    /**
     * @Desc 导出文件的保存路径
     * @return string
     * @Date 2023/6/21 21:34
     */
    public function getStoreFilePath() {
        return $this->getTmpDir() . '/';
    }
    
    public function setFilename($filename) {
        $this->fileName = $filename . Date('YmdHis') . '.xlsx';
        return $this;
    }
    
    public function getFilename() {
        return $this->fileName;
    }
    
    public function getHeader() {
        return $this->header;
    }
    
    public function getTableTitle() {
        return $this->tableTitle;
    }
    
    public function getData() {
        return $this->data;
    }
    
    public function getChunkData() {
        return $this->chunkData;
    }
    
    public $index;
    public $dataSourceType;
    
    /**
     * @Desc 初始化数据源，判断数据源的类型
     * @param array|Collection|Builder $dataSource
     * @return $this
     * @Date 2023/6/21 22:02
     */
    public function initDataSource($dataSource) {
        if ($dataSource instanceof Builder) {
            $this->dataSourceType = 'query';
            $this->setQuery($dataSource);
            $dataSource = [];
        } else if (is_array($dataSource) || $dataSource instanceof Collection) {
            $this->dataSourceType = 'collection';
        }
        $this->setData($dataSource);
        $this->index = 1;
        return $this;
    }
    
    public function setData($data) {
        if (!$data instanceof Collection) {
            $data = collect($data);
        }
        $this->data = $data;
        
        return $this;
    }
    
    abstract public function eachRow($row);
    
    public $fontFamily = '微软雅黑';
    public $rowHeight = 40;
    public $headerRowHeight = 40;
    public $titleRowHeight = 50;
    public $filePath;
    /**
     * @var Excel $excel
     */
    public $excel;
    public $headerLen;
    public $end;
    /**
     * @var Collection
     */
    public $headerData;
    
    public function setHeaderData() {
        $this->headerData = collect([]);
        if ($this->useTitle) {
            $this->headerData->push([$this->getTableTitle()]);
        }
        $this->headerData->push(array_column($this->getHeader(), 'name'));
        return $this;
    }
    
    public $query;
    
    /**
     * BaseExport constructor.
     * @param Builder|array|Collection|null $dataSource
     */
    public function __construct($dataSource, $time = null) {
        if ($this->debug) {
            $this->time = $time ?? microtime(true);
            dump('开始内存占用：' . memory_get_peak_usage() / 1024000);
        }
        $this->init($dataSource);
    }
    
    public $config;
    
    public function setConfig($config = null) {
        $this->config = ['path' => $this->getStoreFilePath()];
        return $this;
    }
    
    public function setQuery($query) {
        $this->query = $query;
        return $this;
    }
    
    public function getFinalFileName() {
        return $this->setFilename($this->fileName)->getFilename();
    }
    
    public $sheetName = 'Sheet1';
    
    public function setSheet($name) {
        $this->sheetName = $name;
        return $this;
    }
    
    public function newExcel($config) {
        $this->setExcel(new Excel($config));
        return $this;
    }
    
    public function init($dataSource = null) {
        $this->setConfig()
             ->initDataSource($dataSource)
             ->newExcel($this->config)
            ->excel
            ->fileName($this->getFinalFileName(), $this->sheetName);
        
        return $this;
    }
    
    public function setExcel(Excel $excel) {
        $this->excel = $excel;
        return $this;
    }
    
    public function getExcel() {
        return $this->excel;
    }
    
    /**
     * @Desc 设置表格冻结
     * @param int $row
     * @param int $column
     * @return $this
     * @Date 2023/6/25 17:59
     */
    public function freezePanes(int $row = 2, int $column = 0) {
        if ($this->useFreezePanes) {
            $this->excel->freezePanes($row, $column);        // 冻结前两行，列不冻结
        }
        return $this;
    }
    
    public $useFreezePanes = false;
    
    public function useFreezePanes(){
        $this->useFreezePanes=true;
        return $this;
    }
    
    public function beforeInsertData() {
        return $this;
    }
    
    // 是否使用首行标题
    public $useTitle = true;
    public $titleStyle;
    
    public function setTitleStyle() {
        // title style
        $this->titleStyle =
            (new Format($this->fileHandle))
                ->fontSize(16)
                ->bold()
                ->font($this->fontFamily)
                ->align(Format::FORMAT_ALIGN_CENTER, Format::FORMAT_ALIGN_VERTICAL_CENTER)
                ->wrap()
                ->toResource();
    }
    
    public $headerStyle;
    
    public function setHeaderStyle() {
        // title style
        $this->headerStyle =
            (new Format($this->fileHandle))
                ->fontSize(10)
                ->font($this->fontFamily)
                ->bold()
                ->align(Format::FORMAT_ALIGN_CENTER, Format::FORMAT_ALIGN_VERTICAL_CENTER)
                ->border(Format::BORDER_THIN)
                ->wrap()
                ->toResource();
    }
    
    public $useGlobalStyle = false;
    public $globalStyle;
    
    public function setGlobalStyle() {
        // global style
        $this->globalStyle = (new Format($this->fileHandle))
            ->fontSize(10)
            ->font($this->fontFamily)
            ->align(Format::FORMAT_ALIGN_CENTER, Format::FORMAT_ALIGN_VERTICAL_CENTER)
            ->border(Format::PATTERN_NONE)
            ->wrap()
            ->toResource();
        $this->excel->defaultFormat($this->globalStyle); // 默认样式
        return $this;
    }
    
    public $normalStyle;
    
    public function getNormalStyle() {
        return $this->normalStyle ?: $this->normalStyle = (new Format($this->fileHandle))
            ->fontSize(10)
            ->font($this->fontFamily)
            ->align(Format::FORMAT_ALIGN_CENTER, Format::FORMAT_ALIGN_VERTICAL_CENTER)
            ->border(Format::BORDER_THIN)
            ->wrap()
            ->toResource();
    }
    
    public function setColumnStyle() {
        $this->columnWidths = array_column($this->getHeader(), 'width');
        
        // 设置列宽 以及默认样式
        foreach ($this->columnWidths as $k => $columnWidth) {
            $column = $this->getColumn($k);
            if ($this->useGlobalStyle) {
                $this->excel->setColumn($column . ':' . $column, $columnWidth);
            } else {
                $this->excel->setColumn($column . ':' . $column, $columnWidth, $this->getNormalStyle());
            }
    
        }
    }
    
    // 开始插入数据
    public function startInsertData() {
        if ($this->useGlobalStyle) {
            $this->setGlobalStyle();
        }
        $this->setTitleStyle();
        $this->setHeaderStyle();
        $this->setColumnStyle();
        
        // 全部导出时，分块插入数据
        $this->filePath = $this->insertHeaderData()
                               ->chunk(function(int $times, $perPage) {
                                   return $this->buildData($times, $perPage);
                               });
        unset($this->data);
        
        return $this;
    }
    
    public function afterStore() {
    }
    
    public $fileHandle;
    public $columnWidths;
    
    public function setEnd($end = null) {
        $this->end = $end ?? $this->getColumn($this->headerLen - 1);
        return $this;
    }
    
    public function setHeaderLen($headerLen = null) {
        $this->headerLen = $headerLen ?? count($this->getHeader());
        return $this;
    }
    
    public function setFileHandle($fileHandle = null) {
        $this->fileHandle = $fileHandle ?? $this->excel->getHandle();
        return $this;
    }
    
    public function setFilePath($filePath = null) {
        $this->filePath = $filePath ?? $this->output();
        return $this;
    }
    
    public function store() {
        return $this->setFileHandle()  // 设置文件处理对象
                    ->freezePanes()       // 冻结前两行，列不冻结
                    ->setHeaderLen() // 设置最大列数
                    ->setEnd() // 设置末尾的列名
                    ->setHeaderData() // 设置表头数据
                    ->beforeInsertData() // 插入正式数据前回调
                    ->startInsertData() // 开始插入数据
                    ->afterInsertData() // 插入数据完成回调
                    ->setFilePath(); // 输出文件到临时目录，并设置文件地址
    }
    
    /**
     * @Desc 设置普通行的样式
     * @return Excel
     * @Date 2023/6/14 18:12
     */
    public function setRowHeight() {
        return $this->excel->setRow($this->currentLine + 1, $this->rowHeight);
    }
    
    public function setTitleHeight() {
        return $this->excel->setRow("A{$this->getCurrentLine()}", $this->titleRowHeight);                                  // title样式
    }
    
    public function setHeaderHeight() {
        return $this->excel->setRow("A{$this->getCurrentLine()}", $this->headerRowHeight);
    }
    
    public $shouldDelete = false;
    public $startDataRow;     // 第三行开始数据行(0是第一行）
    public $currentLine = 0;  // 当前数据插入行
    
    public function getCurrentLine() {
        return $this->currentLine + 1;
    }
    
    public function insertHeaderData() {
        $this->startDataRow = count($this->headerData);
        foreach ($this->headerData as $row => $rowData) {
            $isHeader = true;
            if ($this->currentLine === 0 && $this->useTitle) {
                $isHeader = false;
                $this->setTitleHeight();
            } else {
                $this->setHeaderHeight();
            }
            
            foreach ($rowData as $column => $columnData) {
                $this->insertCell(
                    $this->currentLine
                    , $column
                    , $columnData
                    , null
                    , $isHeader ? $this->headerStyle : null
                );
            }
            
            $this->currentLine++;
        }
        return $this;
    }
    
    public function getIndex() {
        return $this->index;
    }
    
    /**
     * @Desc 根据序号获取rowData，分块时会被销毁
     * @param $index
     * @return mixed
     * @Date 2023/6/14 22:38
     */
    public function getRowInChunkByIndex($index) {
        return $this->chunkData->where('index', $index)->first();
    }
    
    /**
     * @var Collection $chunkData 分块数据
     */
    public $chunkData;
    
    public function insertChunkData(Collection $data) {
        $this->chunkData = $data;
        $index = $this->getIndex();
        
        // 给每行数据绑定index
        foreach ($this->chunkData as $k => $rowData) {
            if ($rowData instanceof Model) {
                $rowData->index = $index;
            } else {
                $rowData['index'] = $index;
                $this->chunkData->put($k, $rowData);
            }
            $index++;
        }
        
        foreach ($this->chunkData as $rowData) {
            $this->setRowHeight();
            
            $rowArray = $this->eachRow($rowData);
            
            foreach ($rowArray as $column => $columnData) {
                $this->insertCell($this->currentLine, $column, $columnData);
            }
            
            // 行插入后回调，$this->chunkData是分块数据，绑定了index，$this->getCurrentLine()获取当前行数，$this->getRowByIndex($this->index）获取该行数据。
            $this->afterInsertEachRowInEachChunk($rowData);
            
            $this->index++;
            $this->currentLine++;
        }
        
        unset($rowArray, $column);
        
        return $this;
    }
    
    /**
     * @Desc 在分块数据插入每行后回调（到下一个分块，则上一分块被销毁）
     * @param $rowData
     * @Date 2023/6/14 22:55
     */
    public function afterInsertEachRowInEachChunk($rowData) {
    }
    
    public function getCellName(int $currentLine, int $column) {
        return $this->getColumn($column) . $currentLine;
    }
    
    public function insertCellHandle($currentLine, $column, $data, $format, $formatHandle) {
        return $this->excel->insertText($currentLine, $column, $data, $format, $formatHandle);
    }
    
    /**
     * Insert data on the cell
     * @param int               $currentLine
     * @param int               $column
     * @param int|string|double $data
     * @param string|null       $format
     * @param resource|null     $formatHandle
     * @return Excel
     */
    public function insertCell(int $currentLine, int $column, $data, ?string $format = null, $formatHandle = null) {
        try {
            if ($this->useGlobalStyle) {
                $formatHandle = $formatHandle ?? $this->getNormalStyle();
            }
            return $this->insertCellHandle($currentLine, $column, $data, $format, $formatHandle);
        } catch (Exception $e) {
            throw new Exception('行数为' . $this->getCurrentLine() . '的记录导出失败，原因：' . $e->getMessage());
        }
    }
    
    /**
     * @var array 定义静态数据合并
     */
    public $mergeCellsByStaticData;
    
    public function mergeCellsAfterInsertData() {
        if ($this->useTitle) {
            return [
                ['range' => "A1:{$this->end}1", 'value' => $this->getTableTitle(), 'formatHandle' => $this->titleStyle],
            ];
        }
        return [];
    }
    
    public function afterInsertData() {
        if (!empty($this->mergeCellsByStaticData = $this->mergeCellsAfterInsertData())) {
            foreach ($this->mergeCellsByStaticData as $i) {
                $this->excel->mergeCells($i['range'], $i['value'], $i['formatHandle'] ?? null);
            }
        }
        if ($this->debug) {
            dump('触发afterInsertData-耗时' . (number_format(microtime(true) - $this->time, 2)) . '秒' . "-" . '内存：' . memory_get_peak_usage() / 1024000);
            dd('数据插入已完成');
        }
        return $this;
    }
    
    public function beforeOutput() {
    }
    
    public function output() {
        $this->beforeOutput();
        return $this->excel->output();
    }
    
    public $columnMap = [];
    
    /**
     * @Desc 根据列数得到字母
     * 可以看做10进制转26进制，除26取余，逆序排列，把余数转成字母倒序拼接。
     * @param int $columnIndex
     * @return string
     * @Date 2023/6/15 17:51
     */
    public function getColumn(int $columnIndex) {
        if (array_key_exists($columnIndex, $this->columnMap)) {
            return $this->columnMap[$columnIndex];
        }
        
        // 由于循环条件为$divide>0，而且$columnIndex从0开始，所以+1
        $divide = $columnIndex + 1;
        $columnName = '';
        while ($divide > 0) {
            // $mod为0~25，对应26个字母，$divide初始最小为1，要-1才能得到正确的余数范围
            $mod = ($divide - 1) % 26;
            $columnName = chr(65 + $mod) . $columnName;
            $divide = (int)(($divide - $mod) / 26); // 减$mod，就是去掉末尾一位的数，除以26，相当于去掉这个数位，循环这个过程，直到取到最高位，也就是截取后的数，前面为0
        }
        return $this->columnMap[$columnIndex] = $columnName;
    }
    
    public $columnIndexMap = [];
    
    /**
     * @Desc 根据字母列名得到列数
     * @param string $columnName
     * @return float|int
     * @Date 2023/6/15 19:49
     */
    public function getColumnIndexByName(string $columnName) {
        if (array_key_exists($columnName, $this->columnIndexMap)) {
            return $this->columnIndexMap[$columnName];
        }
        // 将列名中的字母按顺序拆分成一个一个单独的字母，并进行倒序排列。
        $columnNameReverse = strrev($columnName);
        $arr = str_split($columnNameReverse);
        
        // 对每个字母进行转换，将其转换为对应的数字
        $columnIndex = 0;
        foreach ($arr as $key => $value) {
            $num = ord($value) - 64;
            $columnIndex += $num * (26 ** $key);
        }
        // 将最终计算出的列数值减去1，以得到以0为起点的列数值
        return $this->columnIndexMap[$columnName] = $columnIndex - 1;
    }
    
    public function shouldDelete($v = true) {
        $this->shouldDelete = $v;
        return $this;
    }
    
    public function download($filePath = null) {
        if ($filePath) {
            $this->filePath = $filePath;
        }
        response()->download($this->filePath)->deleteFileAfterSend($this->shouldDelete)->send();
        exit();
    }
    
    public function export() {
        $this->store()->shouldDelete()->download();
    }
    
    public $max = 500000;    // 最大一次导出50万条数据
    public $chunkSize = 5000;// 分块处理 5000查一次 ，数值越大，内存占用越大
    public $completed = 0;   // 已完成
    public $debug = false;
    
    /**
     * @param string $fontFamily
     */
    public function setFontFamily(string $fontFamily) {
        $this->fontFamily = $fontFamily;
        return $this;
    }
    
    /**
     * @param int $headerRowHeight
     */
    public function setHeaderRowHeight(int $headerRowHeight) {
        $this->headerRowHeight = $headerRowHeight;
        return $this;
    }
    
    /**
     * @param int $titleRowHeight
     */
    public function setTitleRowHeight(int $titleRowHeight) {
        $this->titleRowHeight = $titleRowHeight;
        return $this;
    }
    
    /**
     * @param bool $useTitle
     */
    public function setUseTitle(bool $useTitle) {
        $this->useTitle = $useTitle;
        return $this;
    }
    
    /**
     * @param int $max
     */
    public function setMax(int $max) {
        $this->max = $max;
        return $this;
    }
    
    /**
     * @param int $chunkSize
     */
    public function setChunkSize(int $chunkSize) {
        $this->chunkSize = $chunkSize;
        return $this;
    }
    
    /**
     * @param bool $debug
     */
    public function setDebug(bool $debug) {
        $this->debug = $debug;
        return $this;
    }
    
    public $time;
    
    public function chunk($callback = null) {
        $times = 1;
        $this->completed = 0;
        
        do {
            /** @var Collection $result */
            $result = $callback($times, $this->chunkSize);
            // dd($result->toArray());
            $count = count($result);
            $this->completed += $count;
            // dd($times,$result,$count);
            $this->insertChunkData($result);
            unset($this->chunkData, $result);
            if ($this->debug) {
                dump('已导出：' . $this->completed . '条，耗时' . (number_format(microtime(true) - $this->time, 2)) . '秒' . "-" . '内存：' . memory_get_peak_usage() / 1024000);
            }
            $times++;
        } while ($count === $this->chunkSize && $this->completed < $this->max);
        
        return $this;
    }
    
    /**
     * Get data with export query.
     * @param int $page    第几个分块
     * @param int $perPage 分块大小
     * @return Collection
     */
    public function buildData(?int $page = null, ?int $perPage = null) {
        switch ($this->dataSourceType) {
            case 'query':
                return $this->buildDataFromQuery($page, $perPage);
            case 'collection':
                return $this->buildDataFromCollection($page, $perPage);
            default :
                throw new Exception('无效的数据源类型');
        }
    }
    
    public function buildDataFromQuery(?int $page = null, ?int $perPage = null) {
        return $this->query->forPage($page, $perPage)->get();
    }
    
    public function buildDataFromCollection(?int $page = null, ?int $perPage = null) {
        return $this->data->forPage($page, $perPage);
    }
    
    /**
     * Create a instance.
     * @param mixed ...$params
     * @return $this
     */
    public static function make(...$params) {
        return new static(...$params);
    }
}
