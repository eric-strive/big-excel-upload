<?php
namespace excelUpload;
use excelUpload\readFilter\ReadFilterByRow;
use excelUpload\readFilter\TitleReadFilter;
use PhpOffice\PhpSpreadsheet\Exception;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;

/**
 * 切割类
 */
class ExcelCutWrite
{
    public $cutNum = 5;
    public $returnType = 'Csv';
    public $fileDir = '/tmp/';
    public $log;

    /**
     * 切割字符串
     */
    public function cutFromStr($str)
    {
        try {
            $filePath = '/tmp/' . time() . mt_rand(1000, 9000) . $this->returnType;
            file_put_contents($filePath, $str);
            if (file_exists($filePath)) {
                $result =  $this->cutFromFile($filePath);
                unlink($filePath);
                return $result;
            } else {
                throw new Exception('文件写入错误');
            }
        } catch (Exception $e) {
            $this->log = $e->getMessage();
            return false;
        }

    }


    /**
     * 切割文件
     */
    public function cutFromFile($file)
    {
        try {
            $cutRules = $this->readaheadFromFile($file);
            $dir = $this->getFileDir($file);
            $returnType = $this->returnType ? $this->returnType : 'Csv';
            $results = array();
            //初始化读
            $myFilter = new ReadFilterByRow();
            $inputFileType = IOFactory::identify($file);
            $reader = IOFactory::createReader($inputFileType);
            $reader->setReadDataOnly(true);
            $reader->setReadFilter($myFilter);

            foreach ($cutRules as $sheetName => $rowIndexRange) {
                //读
                list($myFilter->startRow, $myFilter->endRow, $myFilter->worksheetName) = $rowIndexRange;
                $spreadsheetReader = $reader->load($file);
                $sheetData = $spreadsheetReader->setActiveSheetIndexByName($myFilter->worksheetName)->toArray(null, false, false, false);
                $realDatas = array_splice($sheetData, $myFilter->startRow, ($myFilter->endRow - $myFilter->startRow + 1));
                $spreadsheetReader->disconnectWorksheets();
                unset($sheetData);
                unset($spreadsheetReader);

                //写
                $saveFile = $dir . $sheetName . '.' . $returnType;
                $spreadsheetWriter = new Spreadsheet();
                foreach ($realDatas as $rowIndex => $row) {
                    foreach ($row as $colIndex => $col) {
                        $spreadsheetWriter->getActiveSheet()->setCellValueByColumnAndRow($colIndex+1, $rowIndex+1, $col);
                    }
                }
                $writer = IOFactory::createWriter($spreadsheetWriter, $returnType);
                $writer->save($saveFile);
                $spreadsheetWriter->disconnectWorksheets();
                unset($spreadsheetWriter);
                $results[] = $saveFile;
            }

            return $results;

        } catch (Exception $e) {
            $this->log = $e->getMessage();
            return false;
        }
    }

    /**
     * 预读文件
     */
    public  function readaheadFromFile($file)
    {
        if (file_exists($file)) {

            //获取统计数据
            $myFilter = new TitleReadFilter();
            $inputFileType = IOFactory::identify($file);
            $reader = IOFactory::createReader($inputFileType);
            $reader->setReadDataOnly(true); //只读数据
            $reader->setReadFilter($myFilter);
            $spreadsheet = $reader->load($file);
            //$sheetData = $spreadsheet->getActiveSheet()->toArray(null, false, false, false);
            list($fileName,) = explode('.', basename($file));

            $datas = array();
            $averageNum = ceil(array_sum($myFilter->record) / $this->cutNum);
            foreach ($myFilter->record as $sheetName => $count) {
                for ($i=0; $i<ceil($count/$averageNum); $i++) {
                    $datas[$fileName . '_' . $sheetName . '_' . $i] = array($i*$averageNum, ($i+1)*$averageNum-1, $sheetName);
                }
            }

            return $datas;
        } else {
            throw new Exception($file . ' not exists');
        }
    }
    
    /**
     * 创建目录
     */
    protected function getFileDir($file)
    {
        $baseName = basename($file);
        list($name) = explode('.', $baseName);
        $fullName = $name .'_'. time() . '_' . mt_rand(1000, 9999);
        $path = $this->fileDir . $fullName . '/';
        mkdir($path, 0777);
        chmod($path, 0777);

        if (is_dir($path)) {
            return $path;
        } else {
            $this->log = "mkdir {$path} failed";
            return false;
        }
    }



}