<?php

use PHPUnit\Framework\TestCase;

class ExcelCutTest extends TestCase
{
    /**
     * 大文件分割
     *
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     * @throws \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    public function testReadaheadFromFile()
    {
        $datas = [
            '工作表_1' => [
                [1, 2, 3, 4],
                ['a', 'b', 'c', 'd'],
                [1, 2, 3, 4],
            ],
            '工作表_2' => [
                ['a', 'b', 'c', 'd'],
                ['a', 'b', 'c', 'd'],
                [1, 2, 3, 4],
                ['a', 'b', 'c', 'd'],
                ['a', 'b', 'c', 'd'],
                [1, 2, 3, 4],
                [1, 2, 3, 4],
                ['a', 'b', 'c', 'd'],
                [1, 2, 3, 4],
                ['a', 'b', 'c', 'd'],
                ['a', 'b', 'c', 'd'],
                ['a', 'b', 'c', 'd'],
                ['a', 'b', 'c', 'd'],
                ['a', 'b', 'c', 'd'],
                ['a', 'b', 'c', 'd'],
                ['a', 'b', 'c', 'd'],
                ['a', 'b', 'c', 'd'],
                ['a', 'b', 'c', 'd'],
            ],
        ];

        require_once '../src/PhpSpreadsheet/autoload.php';
        $spreadsheet = new \PhpOffice\PhpSpreadsheet\Spreadsheet();
        $i           = 0;
        foreach ($datas as $sheetName => $sheetData) {
            if ($i > 0) {
                $spreadsheet->createSheet();
                $spreadsheet->setActiveSheetIndex($i);
            }
            $spreadsheet->getActiveSheet()->setTitle($sheetName);
            foreach ($sheetData as $rowIndex => $row) {
                foreach ($row as $colIndex => $col) {
                    $spreadsheet->getActiveSheet()->setCellValueByColumnAndRow($colIndex + 1, $rowIndex + 1, $col);
                }
            }
            $i++;
        }
        $writer = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet, 'Xlsx');
        $writer->save('/tmp/a.xlsx');

        $excel         = new \excelUpload\ExcelCutWrite();
        $excel->cutNum = 5;
        $result        = $excel->readaheadFromFile('/tmp/a.xlsx');

        $this->assertEquals($result, [
            'a_工作表_1_0' => [0, 4, '工作表_1'],
            'a_工作表_2_0' => [0, 4, '工作表_2'],
            'a_工作表_2_1' => [5, 9, '工作表_2'],
            'a_工作表_2_2' => [10, 14, '工作表_2'],
            'a_工作表_2_3' => [15, 19, '工作表_2'],
        ]);
        unlink('/tmp/a.xlsx');
    }

    /**
     * 分割大文件到多个小文件中
     *
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     * @throws \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    public function testCutFromFile()
    {
        $file  = '/tmp/a.xlsx';
        $datas = [
            '工作表_1' => [
                [1, 'a'],
                [2, 'b'],
                [3, 'c'],
                [4, 'd'],
            ],
            '工作表_2' => [
                [10, 'ab'],
                [20, 'bc'],
                [30, 'cd'],
            ],
        ];
        $spreadsheet = new \PhpOffice\PhpSpreadsheet\Spreadsheet();
        $i           = 0;
        foreach ($datas as $sheetName => $sheetData) {
            if ($i > 0) {
                $spreadsheet->createSheet();
                $spreadsheet->setActiveSheetIndex($i);
            }
            $spreadsheet->getActiveSheet()->setTitle($sheetName);
            foreach ($sheetData as $rowIndex => $row) {
                foreach ($row as $colIndex => $col) {
                    $spreadsheet->getActiveSheet()->setCellValueByColumnAndRow($colIndex + 1, $rowIndex + 1, $col);
                }
            }
            $i++;
        }
        $writer = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet, 'Xlsx');
        $writer->save($file);

        $excel             = new \excelUpload\ExcelCutWrite();
        $excel->cutNum     = 4;
        $excel->returnType = 'Csv';
        $result            = $excel->cutFromFile($file);

        $this->assertEquals(basename($result[0]), 'a_工作表_1_0.Csv');
        $this->assertEquals(basename($result[1]), 'a_工作表_1_1.Csv');
        $this->assertEquals(basename($result[2]), 'a_工作表_2_0.Csv');
        $this->assertEquals(basename($result[3]), 'a_工作表_2_1.Csv');


        $this->assertEquals(preg_split("/(\s|,)+/", file_get_contents($result[0])), ['"1"', '"a"', '"2"', '"b"', '']);
        $this->assertEquals(preg_split("/(\s|,)+/", file_get_contents($result[1])), ['"3"', '"c"', '"4"', '"d"', ""]);
        $this->assertEquals(preg_split("/(\s|,)+/", file_get_contents($result[2])),
            ['"10"', '"ab"', '"20"', '"bc"', ""]);
        $this->assertEquals(preg_split("/(\s|,)+/", file_get_contents($result[3])), ['"30"', '"cd"', ""]);
        unlink($file);
    }

    /**
     * 分割读数据
     *
     * @param $localFilePath
     *
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     * @throws \PhpOffice\PhpSpreadsheet\Reader\Exception
     */
    public static function fileCutRead($localFilePath)
    {
        $excelCutModel = new \excelUpload\ExcelCutRead($localFilePath,
            [
                'startRow'         => '',
                'averageNum'       => 5000,
                'invertedSubtract' => '',
                'isReverse'        => '',
            ]);
        //分批获取以及处理需要的数据
        $excelCutModel->cutFromFile(function ($excelData) {
            //处理$excelData
        });
    }


}