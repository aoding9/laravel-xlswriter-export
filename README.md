### 简介

laravel扩展：xlswriter导出

之前用了laravel-excel做数据导出，太耗内存速度也慢，数据量大的时候容易超时卡死，换xlswriter这个扩展来搞，分块查询处理，及时销毁已插入的数据，减少内存占用。

**chunk=2000,导出1万条**

![laravel扩展：xlswriter导出，自定义复杂合并及样式](https://cdn.learnku.com/uploads/images/202306/21/78338/2ByrnkgGCh.png!large)

**chunk=50000 导出50万条**

![laravel扩展：xlswriter导出，自定义复杂合并及样式](https://cdn.learnku.com/uploads/images/202306/21/78338/4Vt41lzmc6.png!large)



**导出类简单示例**

![laravel扩展：xlswriter导出，自定义复杂合并及样式](https://cdn.learnku.com/uploads/images/202306/21/78338/Kt61PhQsZd.png!large)


![laravel扩展：xlswriter导出，自定义复杂合并及样式](https://cdn.learnku.com/uploads/images/202306/21/78338/9Dz0wi42rl.png!large)


**自定义组装数据集合（2种方法）**


![laravel扩展：xlswriter导出，自定义复杂合并及样式](https://cdn.learnku.com/uploads/images/202306/21/78338/BaobtwVQlL.png!large)

![laravel扩展：xlswriter导出，自定义复杂合并及样式](https://cdn.learnku.com/uploads/images/202306/21/78338/TjWY5sNSfK.png!large)

![laravel扩展：xlswriter导出，自定义复杂合并及样式](https://cdn.learnku.com/uploads/images/202306/21/78338/2X189w4zSP.png!large)


**复杂合并及指定单元格样式**


![laravel扩展：xlswriter导出，自定义复杂合并及样式](https://cdn.learnku.com/uploads/images/202306/21/78338/3EGobTIbWm.png!large)


### 安装

首先根据xlswriter文档安装扩展

https://xlswriter-docs.viest.me/

在phpinfo中确认安装成功后，进行下一步

`composer require aoding9/laravel-xlswriter-export`

国内composer镜像如果安装失败，请设置官方源

`composer config repo.packagist composer https://packagist.org`

因为官方源下载慢，国内镜像又有各种问题可能导致安装失败，也可以把以下代码添加到composer.json，直接从github安装

```json
{
  "repositories": [
    {
      "type": "vcs",
      "url": "https://github.com/aoding9/laravel-xlswriter-export"
    }
  ]
}
```

### 配置

暂无配置

### 使用

1、以用户导出为例，首先创建一个UserExport导出类，继承`Aoding9\Laravel\Xlswriter\Export\BaseExport`基类，一般放在app\Exports目录下

```php
<?php
namespace Aoding9\Laravel\Xlswriter\Export\Demo;
use Aoding9\Laravel\Xlswriter\Export\BaseExport;

class UserExport extends BaseExport {
    public $header = [
        ['column' => 'a', 'width' => 8, 'name' => '序号'],
        ['column' => 'b', 'width' => 8, 'name' => 'id'],
        ['column' => 'c', 'width' => 10, 'name' => '姓名'],
        ['column' => 'd', 'width' => 10, 'name' => '性别'],
        ['column' => 'e', 'width' => 20, 'name' => '注册时间'],
    
    ];
    
    public $fileName = '用户导出表'; // 导出的文件名
    public $tableTitle = '用户导出表'; // 第一行标题
    
    // 将模型字段与表头关联
    public function eachRow($row) {
        return [
            $this->index,
            $row->id,
            \Faker\Factory::create('zh_CN')->name,
            random_int(0, 1) ? '男' : '女',
            $row->created_at->toDateTimeString(),
        ];
    }
}

```
**使用自定义的数组或集合**

如果不希望使用查询构造器获取数据，比如从接口获取数据，有2种方式使用自己定义的数据集合。

1、将集合或数组传给构造函数，弊端是需要传入全部数据，无法分块；好处是写法简单，数据在外部定义，适合数据量小的导出
```php
       $data = [
            ['id' => 1, 'name' => '小白', 'created_at' => now()->toDateString()],
            ['id' => 2, 'name' => '小红', 'created_at' => now()->toDateString()],
        ];
        // $data = User::get()->toArray();
        \Aoding9\Laravel\Xlswriter\Export\Demo\UserExportFromCollection::make($data)->export();
		
		\Aoding9\Laravel\Xlswriter\Export\Demo\AreaExportFromCollection::make(\App\Models\Area::query()->limit(500000)->get())->export();
```

**不使用分页获取，直接导50万条数据的集合，因为要保存全部的数据，所以内存占用极高**

![laravel扩展：xlswriter导出，自定义复杂合并及样式](https://cdn.learnku.com/uploads/images/202306/21/78338/cXSEyYQNwb.png!large)

2、构造函数传参留空，在导出类中重写buildData方法，分页返回集合，适合数据量大的情况

```php
 \Aoding9\Laravel\Xlswriter\Export\Demo\UserExportFromCollection::make()->export();

<?php
namespace Aoding9\Laravel\Xlswriter\Export\Demo;

use Aoding9\Laravel\Xlswriter\Export\BaseExport;

class UserExportFromCollection extends BaseExport {
    public $header = [
        ['column' => 'a', 'width' => 8, 'name' => '序号'],
        ['column' => 'b', 'width' => 8, 'name' => 'id'],
        ['column' => 'c', 'width' => 10, 'name' => '姓名'],
        ['column' => 'd', 'width' => 10, 'name' => '性别'],
        ['column' => 'e', 'width' => 20, 'name' => '注册时间'],
    
    ];
    public $fileName = '用户导出表';   // 导出的文件名
    public $tableTitle = '用户导出表'; // 第一行标题
    
    // 将模型字段与表头关联
    public function eachRow($row) {
        return [
            $this->index,
            $row['id'],
            $row['name'],
            random_int(0, 1) ? '男' : '女',
            $row['created_at'],
        ];
    }
    
    // 方法2 可以分块获取数据
    public function buildData(?int $page = null, ?int $perPage = null) {
        return collect([
                           ['id' => 1, 'name' => '小白', 'created_at' => now()->toDateString()],
                           ['id' => 2, 'name' => '小红', 'created_at' => now()->toDateString()],
                       ]);
    }
}

```

***复杂合并单元格，指定单元格样式***

在每个分块插入之前，每行的数据会被绑定一个index值，在每行插入后，会回调`afterInsertEachRowInEachChunk`，在其中可以使用`getCurrentLine`获取当前行数，使用
`getRowByIndex`获取分块中index对应的rowData

`setHeaderData` 设置表头数据，重写可修改预定义的表头、标题等

`$this->excel`是xlswriter的Excel实例，可以使用`$this->excel->mergeCells`合并单元格，更多样式设置方法请参考官方文档。

`afterInsertData`是所有数据插入完成后的回调，默认在其中调用了`mergeCellsAfterInsertData`方法，合并标题，合并表头，或者对整个表格进行最后修改。

`insertCellHandle`是插入单元格数据的处理方法，重写后可实现单独设置某个单元格的样式

`getCellName`可以根据传入的行数和列数，返回单元格名称，配合insertCellHandle，可判断当前写入的单元格

```php
<?php

namespace Aoding9\Laravel\Xlswriter\Export\Demo;

use Aoding9\Laravel\Xlswriter\Export\BaseExport;
use Illuminate\Database\Eloquent\Builder;
use Illuminate\Support\Carbon;
use Vtiful\Kernel\Format;

// 要导出的模型，用于代码提示

class UserMergeExport extends BaseExport {
    public $header = [
        ['column' => 'a', 'width' => 10, 'name' => '序号'],
        ['column' => 'b', 'width' => 10, 'name' => 'id'],
        ['column' => 'c', 'width' => 10, 'name' => '姓名'],
        ['column' => 'd', 'width' => 10, 'name' => '性别'],
        ['column' => 'e', 'width' => 20, 'name' => '注册时间'],
    ];
    
    public function getGender() {
        return random_int(0, 1) ? '男' : '女';
    }
    
    // 处理每行的模型，使其对应到表头
    public function eachRow($row) {
        return [
            $this->index,      // 自增序号，绑定在模型中
            $row->id,
            \Faker\Factory::create('zh_CN')->name,
            $this->getGender(),
            $row->created_at,
        ];
    }
    
    public $fileName = '用户导出表';     // 导出的文件名
    public $tableTitle = '用户导出表';   // 第一行标题
    public $useFreezePanes = false; // 是否冻结表头
    public $fontFamily = '宋体';
    public $rowHeight = 30;       // 行高
    public $titleRowHeight = 40;  // 首行大标题行高 
    public $headerRowHeight = 50; // 表头行高 
    public $useGlobalStyle=false; // 是否用全局默认样式代替列默认样式（为ture时，数据末尾行下方没有边框，但是速度会慢一点点）
	
    /**
     * @Desc 在分块数据插入每行后回调（到下一个分块，则上一分块被销毁）
     * @param $row
     */
    public function afterInsertEachRowInEachChunk($row) {
        // 奇数行进行合并，且不合并到有效数据行之外
        if ($this->index % 2 === 1 && $this->getCurrentLine() < $this->completed + $this->startDataRow) {
            // 定义纵向合并范围，范围形如"B1:B2"
            $range1 = "B" . $this->getCurrentLine() . ":B" . ($this->getCurrentLine() + 1);
            $nextRow = $this->getRowInChunkByIndex($this->index + 1);
            
            $ids = $row->id . '---' . ($nextRow ? $nextRow->id : null);
            // mergeCells（范围, 数据, 样式） ，通过第三个参数可以设置合并单元格的字体颜色等
            $this->excel->mergeCells($range1, $ids, $this->getSpecialStyle());
            
            // 横向合并，形如"C3:D3"
            $range2 = "C" . $this->getCurrentLine() . ":D" . $this->getCurrentLine();
            $nameAndGender = $row->name . "---" . $this->getGender();
            $this->excel->mergeCells($range2, $nameAndGender);
        }
    }
    
    public function setHeaderData() {
        parent::setHeaderData();
        // 把表头放到第三行，第二行留空用于合并
        $this->headerData->put(2, $this->headerData->get(1));
        $this->headerData->put(1, []);
        return $this;
    }
    
    /**
     * @Desc 插入数据完成后进行合并
     * @return array[]
     */
    public function mergeCellsAfterInsertData() {
        // range是合并范围，$this->end是末尾的列名字母，formatHandle指定合并单元格的样式
        return [
            ['range' => "A1:{$this->end}1", 'value' => $this->getTableTitle(), 'formatHandle' => $this->titleStyle],
            ['range' => "A2:A3", 'value' => '序号', 'formatHandle' => $this->getSpecialStyle()],
            ['range' => "B2:B3", 'value' => 'id', 'formatHandle' => $this->headerStyle],
            ['range' => "C2:E2", 'value' => '基本资料', 'formatHandle' => $this->getSpecialStyle()],
        ];
    }
    
    public $specialStyle;
    
    /**
     * 定义个特别的表格样式
     * @return resource
     */
    public function getSpecialStyle() {
        return $this->specialStyle ?: $this->specialStyle = (new Format($this->fileHandle))
            ->background(Format::COLOR_YELLOW)
            ->fontSize(10)
            ->border(Format::BORDER_THIN)
            ->italic()
            ->font('微软雅黑')
            ->align(Format::FORMAT_ALIGN_CENTER, Format::FORMAT_ALIGN_VERTICAL_CENTER)
            ->wrap()
            ->toResource();
    }
    
    // public $specialStyle2;
    // public $specialStyle3;
    // public function getSpecialStyle2() {}
    // public function getSpecialStyle3() {}
    
    /**
     * @Desc 重写插入单元格数据的处理方法，可单独设置某个单元格的样式
     * @param int           $currentLine  单元格行数
     * @param int           $column       单元格列数
     * @param mixed         $data         插入的数据
     * @param string|null   $format       数据格式化
     * @param resource|null $formatHandle 表格样式
     * @return \Vtiful\Kernel\Excel
     * @Date 2023/6/21 23:37
     */
    public function insertCellHandle($currentLine, $column, $data, $format, $formatHandle) {
        // if($this->getCellName($currentLine,$column)==='A4'){ ... } // 根据单元格名称判断

        // 筛选出E列，且日期秒数为偶数的单元格
        if ($this->getColumn($column) === 'E' && $data instanceof Carbon) {
            if ($data->second % 2 === 0) {
				// 设置为上面定义好的样式（黄色背景，斜体，微软雅黑，水平垂直居中等）
                $formatHandle = $this->getSpecialStyle();
            }
            $data = $data->toDateTimeString();
        }
        return $this->excel->insertText($currentLine, $column, $data, $format, $formatHandle);
    }
}

```




2、在控制器中使用

```php
public function exportModels() {
    // 定义查询构造器，设置查询条件，如果有关联关系，使用with预加载以优化查询
    $query=\Aoding9\Laravel\Xlswriter\Export\Demo\User::query();
    
    // 将查询构造器传入构造函数，然后调用export即可触发下载 
    \Aoding9\Laravel\Xlswriter\Export\Demo\UserExport::make($query)->export();
    
    // 合并单元格的demo
    \Aoding9\Laravel\Xlswriter\Export\Demo\UserMergeExport::make($query)->export();
	
	// 用数据集合或数组
	// 方式1：如果给构造函数传数组或集合，必须把数据全部传入
    $data = [
            	['id' => 1, 'name' => '小白', 'created_at' => now()->toDateString()],
            	['id' => 2, 'name' => '小红', 'created_at' => now()->toDateString()],
    ];
	// $data = User::get()->toArray();
	\Aoding9\Laravel\Xlswriter\Export\Demo\UserExportFromCollection::make($data)->export();
	
	// 方式2：无需传参给构造函数，但需要重写buildData方法，分块返回数据
	\Aoding9\Laravel\Xlswriter\Export\Demo\UserExportByCollection::make()->export();
    
    // 地区导出的demo
	// 用于调试模式查看运行耗时
	$time =microtime(true);

	// 用查询构造器
	$query=\App\Models\Area::where('parent_code',0); // 查父级为0的地区，即查省份
	\Aoding9\Laravel\Xlswriter\Export\Demo\AreaExport::make($query,$time)->export();

	// 用数组或集合
	// 数据量大时占用很高，需要修改内存上限，不推荐
	ini_set('memory_limit', '2048M');
	set_time_limit(0);
	\Aoding9\Laravel\Xlswriter\Export\Demo\AreaExportFromCollection::make(\App\Models\Area::query()->limit(500000)->get(),$time)->export();
}
```

3、其他：

合并单元格的范围请使用大写字母，小写字母会报错。

如果eachRow中需要调用关联模型，请使用with预加载以优化查询。

仓库中包含几个导出类的demo,如果你已有users表或者areas表，可以尝试使用demo进行导出测试，设置`$debug=true;`即可查看导出的耗时和内存占用。

为了方便自定义排版和修改数据，基类属性和方法都为public，方便子类重写