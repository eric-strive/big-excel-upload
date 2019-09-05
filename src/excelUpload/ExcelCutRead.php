<?php

namespace excelUpload;

use excelUpload\readFilter\ReadFilterByRow;
use excelUpload\readFilter\TitleReadFilter;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Exception;

/**
 * 文件切割类
 * Class ExcelCut
 *
 * @package common\helpers\excelUpload
 */
class ExcelCutRead
{
    public $averageNum       = 5000;//每页读取行数
    public $startRow         = 0;//开始接收数据的行
    public $invertedSubtract = 0;//文件倒减行
    public $titleRow         = null;//title所在的行
    public $sheet            = 0;//读取数据的sheet
    public $isReverse        = false;//是否倒序
    public $filePath         = '';//文件路径

    public function __construct($filePath, $initData = [])
    {
        $this->averageNum       = $initData['averageNum'] ?? 5000;
        $this->startRow         = $initData['startRow'] ?? 0;
        $this->invertedSubtract = $initData['invertedSubtract'] ?? 0;
        $this->titleRow         = $initData['titleRow'] ?? null;
        $this->sheet            = $initData['sheet'] ?? 0;
        $this->isReverse        = $initData['isReverse'] ?? false;
        $this->filePath         = $filePath;
        if (!file_exists($filePath)) {
            throw new Exception($filePath . ' 不存在');
        }
    }

    /**
     * 切割文件
     *
     * @param callable $callback 处理数据的函数
     *
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     * @throws \PhpOffice\PhpSpreadsheet\Reader\Exception
     */
    public function cutFromFile(callable $callback)
    {
        $filePath = $this->filePath;
        //文件预读，获取分割配置
        $cutRules = $this->readAheadFromFile($filePath);
        //初始化读
        $myFilter = new ReadFilterByRow();
        $reader   = self::excelBeforeLoadProcess($filePath, $myFilter);
        if ($this->isReverse) {
            $cutRules = array_reverse($cutRules);
        }
        foreach ($cutRules as $sheetName => $rowIndexRange) {
            list($myFilter->startRow, $myFilter->endRow) = $rowIndexRange;
            $this->obtainingProcessingData($reader, $myFilter, $callback);
        }
    }

    /**
     * @param $reader
     * @param $myFilter ReadFilterByRow
     * @param $callback 处理数据的函数
     */
    private function obtainingProcessingData($reader, $myFilter, $callback)
    {
        echo sprintf("处理文件第%s行到%s行开始.\n", $myFilter->startRow, $myFilter->endRow);
        $spreadsheetReader = $reader->load($this->filePath);
        //获取数据
        $sheetData = $spreadsheetReader->getSheet($this->sheet)
            ->toArray(null, false, false, false);
        //删除不需要的数据
        $realData = array_splice($sheetData, $myFilter->startRow,
            ($myFilter->endRow - $myFilter->startRow + 1), false);
        //释放内存
        $spreadsheetReader->disconnectWorksheets();
        unset($sheetData);
        unset($spreadsheetReader);

        call_user_func_array($callback, [$realData, $myFilter->startRow, $myFilter->endRow]);
        echo sprintf("处理文件第%s行到%s行结束......\n", $myFilter->startRow, $myFilter->endRow);
    }

    /**
     * 加载excel之前的处理
     *
     * @param string $filePath
     * @param object $myFilter
     * @param bool   $isOnlyRead
     *
     * @return \PhpOffice\PhpSpreadsheet\Reader\IReader
     * @throws \PhpOffice\PhpSpreadsheet\Reader\Exception
     */
    public static function excelBeforeLoadProcess($filePath, $myFilter, $isOnlyRead = true)
    {
        $inputFileType = IOFactory::identify($filePath);
        $reader        = IOFactory::createReader($inputFileType);
        $reader->setReadDataOnly($isOnlyRead); //只读数据
        $reader->setReadFilter($myFilter);

        return $reader;
    }

    /**
     * 获取文件行数以及title数据
     *
     * @param      $filePath
     * @param null $titleRow
     * @param int  $sheet
     *
     * @return array
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     * @throws \PhpOffice\PhpSpreadsheet\Reader\Exception
     */
    public static function getFileRows($filePath, $titleRow = null, $sheet = 0)
    {
        $myFilter    = new TitleReadFilter($titleRow);
        $reader      = self::excelBeforeLoadProcess($filePath, $myFilter);
        $spreadsheet = $reader->load($filePath)->getSheet($sheet);
        $titleData   = [];
        //如果$titleRow>0就需要获取title这一行
        if ($titleRow > 0) {
            $sheetData = $spreadsheet->toArray(null, false, false, false);
            $titleInfo = array_splice($sheetData, $titleRow - 1, $titleRow);
            if (empty($titleInfo[0])) {
                throw new Exception('获取title出错，请检查excel文件');
            }
            $titleData = $titleInfo[0];
        }

        return [
            'titleData' => $titleData,
            'recordRow' => $myFilter->record,
        ];
    }

    /**
     * 预读文件,获取文件分割
     *
     * @return array
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     * @throws \PhpOffice\PhpSpreadsheet\Reader\Exception
     */
    public function readAheadFromFile()
    {
        $filePath         = $this->filePath;
        $averageNum       = $this->averageNum;
        $startRow         = $this->startRow > 0 ? $this->startRow - 1 : 0;
        $invertedSubtract = $this->invertedSubtract;
        //获取统计数据
        $recordInfo = self::getFileRows($filePath, $this->titleRow);
        $totalRows  = current($recordInfo['recordRow']);//总行数
        $lastRow    = $totalRows - $invertedSubtract;//最后记录ID
        $recordRow  = $lastRow - $startRow;//需要统计记录的条数

        $i        = 1;
        $beginRow = $startRow;
        while ($recordRow > $averageNum) {
            $cutData[$i] = [
                $beginRow,
                ($i) * $averageNum + $startRow - 1,
            ];
            $beginRow    += $averageNum;
            $recordRow   -= $averageNum;
            $i++;
        }
        $cutData[$i] = [
            $beginRow,
            $lastRow - 1,
        ];

        return $cutData;
    }

    /**
     * 字母转为数字
     *
     * @param $abc
     *
     * @return float|int
     */
    public static function AlphabeticConversion($abc)
    {
        $len  = strlen($abc);
        $ten  = 26 * ($len - 1);
        $char = substr($abc, $len - 1, 1);//反向获取单个字符
        $int  = ord($char);
        $ten  += ($int - 65);


        return $ten;
    }
}