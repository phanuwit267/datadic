<?php
require './vendor/autoload.php';

$configDb = [
    'type'     => 'mysql',
    'hostname' => '127.0.0.1',
    'port' => '3306',
    'database' => 'db',
    'username' => 'root',
    'password' => '',
];

try {
    $mysql = new PDO("{$configDb['type']}:host={$configDb['hostname']};dbname={$configDb['database']}", $configDb['username'], $configDb['password']);
    $mysql->query("SET NAMES utf8mb4");
} catch (PDOException $e) {
    exit('Failed to connect to database' . $e->getMessage());
}

$res = $mysql->query('SHOW TABLE STATUS');
$tables = [];
while ($row = $res->fetch()) {
    array_push($tables, [
        'name'      => $row['Name'],
        'engine'    => $row['Engine'],
        'collation' => $row['Collation'],
        'comment'   => $row['Comment'],
    ]);
}

foreach ($tables as &$val) {
    $res = $mysql->query("SHOW FULL FIELDS FROM `{$val['name']}`");
    $fields = [];
    while ($row = $res->fetch()) {
        array_push($fields, [
            'field'     => $row['Field'],
            'type'      => $row['Type'],
            'collation' => $row['Collation'],
            'null'      => $row['Null'],
            'key'       => getKeyName($row['Key']),
            'default'   => $row['Default'],
            'extra'     => $row['Extra'],
            'comment'   => $row['Comment'],
        ]);
    }
    $val['field'] = $fields;
}

function getKeyName($str)
{
    switch ($str) {
        case "PRI":
            $key = "PK";
            break;
        case "MUL":
            $key = "FK";
            break;
        default:
            $key = "";
    }
    return $key;
}

$excel = new PHPExcel();
$excel->getProperties()->setCreator('phanuwit.h@gmail.com');
$excel->getProperties()->setTitle($configDb['database']);

$excel->getDefaultStyle()->getFont()->setName('TH SarabunPSK')->setSize(16);

$excel->setActiveSheetIndex(0);
$excel->getActiveSheet()->setTitle('Data Dictionary');
$activeSheet = $excel->getActiveSheet();

$activeSheet->getColumnDimension('B')->setWidth(10);
$activeSheet->getColumnDimension('C')->setWidth(20);
$activeSheet->getColumnDimension('D')->setWidth(24);
$activeSheet->getColumnDimension('E')->setWidth(20);
$activeSheet->getColumnDimension('F')->setWidth(12);
$activeSheet->getColumnDimension('G')->setWidth(18);
$activeSheet->getColumnDimension('H')->setWidth(30);

$activeSheet->getDefaultStyle()->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$activeSheet->getDefaultStyle()->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
$activeSheet->getDefaultRowDimension()->setRowHeight(20);
$styleArray = [
    'borders' => [
        'allborders' => [
            //'style' => PHPExcel_Style_Border::BORDER_THICK,
            'style' => PHPExcel_Style_Border::BORDER_THIN,
            //'color' => ['argb' => 'FFFF0000'],
        ],
    ],
];
$styleTitleArray = [
    'font'  => [
        'bold'  => true,
        'size'  => 18,
        'name'  => 'TH SarabunPSK'
    ],
    'alignment' => array(
        'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_LEFT,
    )
];

$num = 1;
foreach ($tables as $key => $val) {
    $activeSheet->setCellValue('B' . $num, 'Table ' . ($key + 1) . ' : ' . $val['name']);
    $activeSheet->mergeCells('B' . $num . ':H' . $num);
    $activeSheet->getStyle('B' . $num)->applyFromArray($styleTitleArray);
    $num++;

    $start = $num;
    $activeSheet->setCellValue('B' . $num, 'No');
    $activeSheet->setCellValue('C' . $num, 'Column');
    $activeSheet->setCellValue('D' . $num, 'Data Type');
    $activeSheet->setCellValue('E' . $num, 'Nullable');
    $activeSheet->setCellValue('F' . $num, 'Key');
    $activeSheet->setCellValue('G' . $num, 'Extra');
    $activeSheet->setCellValue('H' . $num, 'Description');
    $activeSheet->getStyle('B' . $num . ':H' . $num)->applyFromArray($styleTitleArray);
    $activeSheet->getStyle('B' . $num . ':H' . $num)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
    $activeSheet->getStyle('B' . $num . ':H' . $num)->getFont()->setName('TH SarabunPSK')->setSize(16);
    $num++;
    foreach ($val['field'] as $k => $v) {
        $activeSheet->setCellValue('B' . $num, $k + 1);
        $activeSheet->setCellValue('C' . $num, $v['field']);
        $activeSheet->setCellValue('D' . $num, $v['type']);
        $activeSheet->setCellValue('E' . $num, $v['null']);
        $activeSheet->setCellValue('F' . $num, $v['key']);
        $activeSheet->setCellValue('G' . $num, $v['extra']);
        $activeSheet->setCellValue('H' . $num, $v['comment']);
        $num++;
    }
    $activeSheet->getStyle('B' . $start . ':H' . ($num - 1))->applyFromArray($styleArray);
    $num++;
}
$write = new PHPExcel_Writer_Excel2007($excel);
$write->save("data_dictionary_" . date('YmdHis') . ".xlsx");
