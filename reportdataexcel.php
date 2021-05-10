<?php  
//menghubungkan dengan file koneksi.php
include ('koneksi.php');
//menghubungkan dengan file autoload.php
require 'vendor/autoload.php';
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

$spreadsheet = new Spreadsheet();
$sheet = $spreadsheet -> getActiveSheet();
//mengisi cell pada file excel
$sheet -> setCellValue('A1','No');
$sheet -> setCellValue('B1','Nama');
$sheet -> setCellValue('C1','Kelas');
$sheet -> setCellValue('D1','Alamat');

//mengambil data dari tabel tb_siswa
$query = mysqli_query($koneksi, "SELECT * FROM tb_siswa");
$i = 2;
$no = 1;
while ($row = mysqli_fetch_array($query)) { //perulangan untuk mencetak data
	//mencetak data dari database ke file excel
	$sheet -> setCellValue('A'.$i, $no++);
	$sheet -> setCellValue('B'.$i, $row['nama']);
	$sheet -> setCellValue('C'.$i, $row['kelas']);
	$sheet -> setCellValue('D'.$i, $row['alamat']);
	$i++;
}

//pengaturan border
$styleArray = [
	'borders' => [
		'allBorders' =>[
			'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
		],
	],
];
$i = $i -1;
$sheet -> getStyle('A1:D'.$i) -> applyFromArray($styleArray);

$writer = new Xlsx($spreadsheet);
//menyimpan dalam bentuk file excel Report Data Siswa.xlsx
$writer -> save('Report Data Siswa.xlsx');
?>