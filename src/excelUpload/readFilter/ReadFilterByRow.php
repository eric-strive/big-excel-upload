<?php

namespace excelUpload\readFilter;

use PhpOffice\PhpSpreadsheet\Reader\IReadFilter;

/**
 * 根据指定的行来过滤
 */
class ReadFilterByRow implements IReadFilter
{
    public $startRow;
    public $endRow;
    public $worksheetName;

    public function readCell($column, $row, $worksheetName = '')
    {
        if ($row >= ($this->startRow + 1) && $row <= ($this->endRow + 1)) {
            return true;
        }

        return false;
    }
}