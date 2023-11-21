<?php
require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Style\Font;

$inputFileName = 'excel.xlsx'; // Path to your Excel file
$outputFileName = 'excel_processed.xlsx'; // Path to the processed Excel file

$spreadsheet = IOFactory::load($inputFileName);
$worksheet = $spreadsheet->getActiveSheet();

$highestRow = $worksheet->getHighestRow(); // Get the highest row number
$highestColumn = $worksheet->getHighestColumn(); // Get the highest column letter

// Loop through all cells and change the font
for ($row = 1; $row <= $highestRow; $row++) {
    for ($column = 'A'; $column <= $highestColumn; $column++) {
        $worksheet->getStyle($column . $row)
            ->getFont()
            ->setName('Times New Roman');
    }
}

$writer = new Xlsx($spreadsheet);
$writer->save($outputFileName);
?>