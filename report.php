<?php
//menghubungkan dengan file autoload.php
require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

$spreadsheet = new Spreadsheet();
$sheet = $spreadsheet -> getActiveSheet();
//mengisi cell pada file excel
$sheet -> setCellValue('A1', 'Hello World !');

$writer = new Xlsx($spreadsheet);
//menyimpan dalam bentuk file excel hello world.xlsx
$writer -> save('hello world.xlsx'); 
?>