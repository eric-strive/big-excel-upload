<?php

namespace excelUpload\readFilter;


use PhpOffice\PhpSpreadsheet\Reader\IReadFilter;

/**
 * 获取excel文件的总行数以及title行
 * Class TitleReadFilter
 *
 * @package common\helpers\excelUpload\readFilter
 */
class TitleReadFilter implements IReadFilter
{
    public  $record   = [];//总行数
    private $lastRow  = '';
    private $titleRow = '';//title所在的行

    public function __construct($titleRow = null)
    {
        $this->titleRow = $titleRow;
    }

    public function readCell($column, $row, $worksheetName = '')
    {
        if (isset($this->record[$worksheetName])) {
            if ($this->lastRow != $row) {
                $this->record[$worksheetName]++;
                $this->lastRow = $row;
            }
            //返回title数据
            if ($row == $this->titleRow) {
                return true;
            }
        } else {
            $this->record[$worksheetName] = 1;
            $this->lastRow                = $row;

            return true;
        }

        return false;
    }
}