<?php
namespace Aoding9\Laravel\Xlswriter\Export\Demo;

use Aoding9\Laravel\Xlswriter\Export\BaseExport;

class UserExportByCollection extends BaseExport {
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
        /** @var User $row 用于代码提示 */
        return [
            $this->index,
            $row['id'],
            $row['name'],
            random_int(0, 1) ? '男' : '女',
            $row['created_at']->toDateString(),
        ];
    }
    
    public function buildData(?int $page = null, ?int $perPage = null) {
        //return $this->query->forPage($page, $perPage)->get();
        return collect([
                           ['id' => 1, 'name' => '小白', 'created_at' => now()],
                           ['id' => 2, 'name' => '小红', 'created_at' => now()],
                       ]);
    }
}
