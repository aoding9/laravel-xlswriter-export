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

    // public $debug=true;
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
