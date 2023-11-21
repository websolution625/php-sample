<?php
require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Style\NumberFormat;

$inputFileName = 'Excel.xls'; // Đường dẫn file Excel 
$outputFileName = 'Excel_processed.xlsx'; // Đường dẫn file Excel sau khi xử lý
$sheetname = 'Sheet1'; // Tên của sheet muốn đọc

$spreadsheet = IOFactory::load($inputFileName);
$worksheet = $spreadsheet->getSheetByName($sheetname);

$highestRow = $worksheet->getHighestRow(); // Lấy số dòng lớn nhất
$highestColumn = $worksheet->getHighestColumn(); // Lấy số cột lớn nhất

$columnsToFormat = ['Trị Giá', 'Đơn Giá']; // Các cột muốn đọc và định dạng
$columnIndexes = [];

// Tìm chỉ số cột dựa trên giá trị của hàng đầu tiên
for ($column = 'A'; $column !== $highestColumn; ++$column) {
    $cellValue = $worksheet->getCell($column . '1')->getValue();
    if (in_array($cellValue, $columnsToFormat)) {
        $columnIndexes[] = $column;
    }
}

foreach ($columnIndexes as $columnIndex) {
    for ($row = 1; $row <= $highestRow; $row++) {
        $cell = $worksheet->getCell($columnIndex . $row);
        if (is_numeric($cell->getValue())) {
            $worksheet->getStyle($columnIndex . $row)
                ->getNumberFormat()
                ->setFormatCode('#,##0');
        }
    }
}

$writer = new Xlsx($spreadsheet);
$writer->save($outputFileName);
?>
