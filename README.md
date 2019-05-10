## 大文件数据分割读取
$excelCutModel = new \excelUpload\ExcelCutRead($localFilePath,
            [
                'startRow'         => '',//文件数据开始行
                'averageNum'       => 5000,//分割每页需要的数据
                'invertedSubtract' => '',//尾部不需要数据的行数
                'isReverse'        => '',//文件数据读取是否倒序
            ]);
        //分批获取以及处理需要的数据
        $excelCutModel->cutFromFile(function ($excelData) {
            //处理$excelData
        });

## 将大文件分割成小文件

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