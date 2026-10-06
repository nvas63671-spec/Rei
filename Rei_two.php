<?php
// Run this on your localhost (e.g. http://localhost/pixel_art.php) and the Excel file downloads.
require 'vendor/autoload.php';
 
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Style\Fill;
 
$spreadsheet = new Spreadsheet();
$activeWorksheet = $spreadsheet->getActiveSheet();
$activeWorksheet->setTitle('Pixel Art');
 
// Make every cell a square pixel
foreach (range('A', 'U') as $col) {
    $activeWorksheet->getColumnDimension($col)->setWidth(3.2);
}
for ($r = 1; $r <= 28; $r++) {
    $activeWorksheet->getRowDimension($r)->setRowHeight(18);
}
 
// ---- Row 1 ----
$activeWorksheet->getStyle('F1')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('G1')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('H1')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('I1')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('J1')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('K1')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('L1')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('M1')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('N1')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('O1')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('P1')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
 
// ---- Row 2 ----
$activeWorksheet->getStyle('E2')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('F2')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF8AB1CC');
$activeWorksheet->getStyle('G2')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF8AB1CC');
$activeWorksheet->getStyle('H2')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF8AB1CC');
$activeWorksheet->getStyle('I2')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF8AB1CC');
$activeWorksheet->getStyle('J2')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF8AB1CC');
$activeWorksheet->getStyle('K2')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF8AB1CC');
$activeWorksheet->getStyle('L2')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF8AB1CC');
$activeWorksheet->getStyle('M2')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF8AB1CC');
$activeWorksheet->getStyle('N2')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF8AB1CC');
$activeWorksheet->getStyle('O2')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF8AB1CC');
$activeWorksheet->getStyle('P2')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF8AB1CC');
$activeWorksheet->getStyle('Q2')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
 
// ---- Row 3 ----
$activeWorksheet->getStyle('D3')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('E3')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF8AB1CC');
$activeWorksheet->getStyle('F3')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF8AB1CC');
$activeWorksheet->getStyle('G3')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF8AB1CC');
$activeWorksheet->getStyle('H3')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF8AB1CC');
$activeWorksheet->getStyle('I3')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF8AB1CC');
$activeWorksheet->getStyle('J3')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF8AB1CC');
$activeWorksheet->getStyle('K3')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF8AB1CC');
$activeWorksheet->getStyle('L3')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF8AB1CC');
$activeWorksheet->getStyle('M3')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF8AB1CC');
$activeWorksheet->getStyle('N3')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF8AB1CC');
$activeWorksheet->getStyle('O3')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF8AB1CC');
$activeWorksheet->getStyle('P3')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF8AB1CC');
$activeWorksheet->getStyle('Q3')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF8AB1CC');
$activeWorksheet->getStyle('R3')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
 
// ---- Row 4 ----
$activeWorksheet->getStyle('C4')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('D4')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF8AB1CC');
$activeWorksheet->getStyle('E4')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF8AB1CC');
$activeWorksheet->getStyle('F4')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('G4')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFE8E4D9');
$activeWorksheet->getStyle('H4')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF8AB1CC');
$activeWorksheet->getStyle('I4')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF8AB1CC');
$activeWorksheet->getStyle('J4')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF8AB1CC');
$activeWorksheet->getStyle('K4')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF8AB1CC');
$activeWorksheet->getStyle('L4')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF8AB1CC');
$activeWorksheet->getStyle('M4')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF8AB1CC');
$activeWorksheet->getStyle('N4')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF8AB1CC');
$activeWorksheet->getStyle('O4')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFE8E4D9');
$activeWorksheet->getStyle('P4')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('Q4')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF8AB1CC');
$activeWorksheet->getStyle('R4')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF8AB1CC');
$activeWorksheet->getStyle('S4')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
 
// ---- Row 5 ----
$activeWorksheet->getStyle('B5')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('C5')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF8AB1CC');
$activeWorksheet->getStyle('D5')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF8AB1CC');
$activeWorksheet->getStyle('E5')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('F5')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFE8E4D9');
$activeWorksheet->getStyle('G5')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('H5')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF8AB1CC');
$activeWorksheet->getStyle('I5')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF8AB1CC');
$activeWorksheet->getStyle('J5')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF8AB1CC');
$activeWorksheet->getStyle('K5')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF8AB1CC');
$activeWorksheet->getStyle('L5')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF8AB1CC');
$activeWorksheet->getStyle('M5')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF8AB1CC');
$activeWorksheet->getStyle('N5')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF8AB1CC');
$activeWorksheet->getStyle('O5')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('P5')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFE8E4D9');
$activeWorksheet->getStyle('Q5')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('R5')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF8AB1CC');
$activeWorksheet->getStyle('S5')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF8AB1CC');
$activeWorksheet->getStyle('T5')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
 
// ---- Row 6 ----
$activeWorksheet->getStyle('B6')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('C6')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF8AB1CC');
$activeWorksheet->getStyle('D6')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF8AB1CC');
$activeWorksheet->getStyle('E6')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('F6')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('G6')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF8AB1CC');
$activeWorksheet->getStyle('H6')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF8AB1CC');
$activeWorksheet->getStyle('I6')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF8AB1CC');
$activeWorksheet->getStyle('J6')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF8AB1CC');
$activeWorksheet->getStyle('K6')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF8AB1CC');
$activeWorksheet->getStyle('L6')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF8AB1CC');
$activeWorksheet->getStyle('M6')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF8AB1CC');
$activeWorksheet->getStyle('N6')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF8AB1CC');
$activeWorksheet->getStyle('O6')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF8AB1CC');
$activeWorksheet->getStyle('P6')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('Q6')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('R6')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF8AB1CC');
$activeWorksheet->getStyle('S6')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF8AB1CC');
$activeWorksheet->getStyle('T6')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
 
// ---- Row 7 ----
$activeWorksheet->getStyle('A7')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('B7')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF8AB1CC');
$activeWorksheet->getStyle('C7')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF8AB1CC');
$activeWorksheet->getStyle('D7')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF8AB1CC');
$activeWorksheet->getStyle('E7')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF8AB1CC');
$activeWorksheet->getStyle('F7')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF8AB1CC');
$activeWorksheet->getStyle('G7')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF8AB1CC');
$activeWorksheet->getStyle('H7')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF8AB1CC');
$activeWorksheet->getStyle('I7')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF8AB1CC');
$activeWorksheet->getStyle('J7')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF8AB1CC');
$activeWorksheet->getStyle('K7')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF8AB1CC');
$activeWorksheet->getStyle('L7')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF8AB1CC');
$activeWorksheet->getStyle('M7')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF8AB1CC');
$activeWorksheet->getStyle('N7')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF8AB1CC');
$activeWorksheet->getStyle('O7')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF8AB1CC');
$activeWorksheet->getStyle('P7')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF8AB1CC');
$activeWorksheet->getStyle('Q7')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF8AB1CC');
$activeWorksheet->getStyle('R7')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF8AB1CC');
$activeWorksheet->getStyle('S7')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF8AB1CC');
$activeWorksheet->getStyle('T7')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF8AB1CC');
$activeWorksheet->getStyle('U7')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
 
// ---- Row 8 ----
$activeWorksheet->getStyle('A8')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('B8')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF8AB1CC');
$activeWorksheet->getStyle('C8')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF8AB1CC');
$activeWorksheet->getStyle('D8')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF8AB1CC');
$activeWorksheet->getStyle('E8')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF8AB1CC');
$activeWorksheet->getStyle('F8')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF8AB1CC');
$activeWorksheet->getStyle('G8')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF8AB1CC');
$activeWorksheet->getStyle('H8')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('I8')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF8AB1CC');
$activeWorksheet->getStyle('J8')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF8AB1CC');
$activeWorksheet->getStyle('K8')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF8AB1CC');
$activeWorksheet->getStyle('L8')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF8AB1CC');
$activeWorksheet->getStyle('M8')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF8AB1CC');
$activeWorksheet->getStyle('N8')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('O8')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF8AB1CC');
$activeWorksheet->getStyle('P8')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF8AB1CC');
$activeWorksheet->getStyle('Q8')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF8AB1CC');
$activeWorksheet->getStyle('R8')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF8AB1CC');
$activeWorksheet->getStyle('S8')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF8AB1CC');
$activeWorksheet->getStyle('T8')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF8AB1CC');
$activeWorksheet->getStyle('U8')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
 
// ---- Row 9 ----
$activeWorksheet->getStyle('A9')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('B9')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF8AB1CC');
$activeWorksheet->getStyle('C9')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF8AB1CC');
$activeWorksheet->getStyle('D9')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF8AB1CC');
$activeWorksheet->getStyle('E9')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF8AB1CC');
$activeWorksheet->getStyle('F9')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('G9')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('H9')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('I9')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF4B78AA');
$activeWorksheet->getStyle('J9')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF8AB1CC');
$activeWorksheet->getStyle('K9')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF8AB1CC');
$activeWorksheet->getStyle('L9')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF8AB1CC');
$activeWorksheet->getStyle('M9')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF4B78AA');
$activeWorksheet->getStyle('N9')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('O9')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('P9')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF4B78AA');
$activeWorksheet->getStyle('Q9')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF8AB1CC');
$activeWorksheet->getStyle('R9')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF8AB1CC');
$activeWorksheet->getStyle('S9')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF8AB1CC');
$activeWorksheet->getStyle('T9')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF8AB1CC');
$activeWorksheet->getStyle('U9')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
 
// ---- Row 10 ----
$activeWorksheet->getStyle('A10')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('B10')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF4B78AA');
$activeWorksheet->getStyle('C10')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF8AB1CC');
$activeWorksheet->getStyle('D10')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF8AB1CC');
$activeWorksheet->getStyle('E10')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('F10')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('G10')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFF0CFA5');
$activeWorksheet->getStyle('H10')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('I10')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('J10')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF4B78AA');
$activeWorksheet->getStyle('K10')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF8AB1CC');
$activeWorksheet->getStyle('L10')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF4B78AA');
$activeWorksheet->getStyle('M10')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('N10')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('O10')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFF0CFA5');
$activeWorksheet->getStyle('P10')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('Q10')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('R10')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF8AB1CC');
$activeWorksheet->getStyle('S10')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF8AB1CC');
$activeWorksheet->getStyle('T10')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF4B78AA');
$activeWorksheet->getStyle('U10')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
 
// ---- Row 11 ----
$activeWorksheet->getStyle('A11')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('B11')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF4B78AA');
$activeWorksheet->getStyle('C11')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF8AB1CC');
$activeWorksheet->getStyle('D11')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF8AB1CC');
$activeWorksheet->getStyle('E11')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('F11')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFF0CFA5');
$activeWorksheet->getStyle('G11')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('H11')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('I11')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFF0CFA5');
$activeWorksheet->getStyle('J11')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('K11')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF8AB1CC');
$activeWorksheet->getStyle('L11')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('M11')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFF0CFA5');
$activeWorksheet->getStyle('N11')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('O11')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('P11')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFF0CFA5');
$activeWorksheet->getStyle('Q11')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('R11')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF8AB1CC');
$activeWorksheet->getStyle('S11')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF8AB1CC');
$activeWorksheet->getStyle('T11')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF4B78AA');
$activeWorksheet->getStyle('U11')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
 
// ---- Row 12 ----
$activeWorksheet->getStyle('A12')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('B12')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF4B78AA');
$activeWorksheet->getStyle('C12')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF8AB1CC');
$activeWorksheet->getStyle('D12')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF8AB1CC');
$activeWorksheet->getStyle('E12')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('F12')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFF0CFA5');
$activeWorksheet->getStyle('G12')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFC4221A');
$activeWorksheet->getStyle('H12')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('I12')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFF0CFA5');
$activeWorksheet->getStyle('J12')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFF0CFA5');
$activeWorksheet->getStyle('K12')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('L12')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFF0CFA5');
$activeWorksheet->getStyle('M12')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFF0CFA5');
$activeWorksheet->getStyle('N12')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('O12')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFC4221A');
$activeWorksheet->getStyle('P12')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFF0CFA5');
$activeWorksheet->getStyle('Q12')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('R12')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF8AB1CC');
$activeWorksheet->getStyle('S12')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF8AB1CC');
$activeWorksheet->getStyle('T12')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF4B78AA');
$activeWorksheet->getStyle('U12')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
 
// ---- Row 13 ----
$activeWorksheet->getStyle('B13')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('C13')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF4B78AA');
$activeWorksheet->getStyle('D13')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF8AB1CC');
$activeWorksheet->getStyle('E13')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF4B78AA');
$activeWorksheet->getStyle('F13')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFF0CFA5');
$activeWorksheet->getStyle('G13')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFE84A3C');
$activeWorksheet->getStyle('H13')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFE84A3C');
$activeWorksheet->getStyle('I13')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFF0CFA5');
$activeWorksheet->getStyle('J13')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFF0CFA5');
$activeWorksheet->getStyle('K13')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFF0CFA5');
$activeWorksheet->getStyle('L13')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFF0CFA5');
$activeWorksheet->getStyle('M13')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFF0CFA5');
$activeWorksheet->getStyle('N13')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFE84A3C');
$activeWorksheet->getStyle('O13')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFE84A3C');
$activeWorksheet->getStyle('P13')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFF0CFA5');
$activeWorksheet->getStyle('Q13')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('R13')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF8AB1CC');
$activeWorksheet->getStyle('S13')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF4B78AA');
$activeWorksheet->getStyle('T13')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
 
// ---- Row 14 ----
$activeWorksheet->getStyle('B14')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('C14')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF4B78AA');
$activeWorksheet->getStyle('D14')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF4B78AA');
$activeWorksheet->getStyle('E14')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF8AB1CC');
$activeWorksheet->getStyle('F14')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('G14')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFE84A3C');
$activeWorksheet->getStyle('H14')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFE84A3C');
$activeWorksheet->getStyle('I14')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFF0CFA5');
$activeWorksheet->getStyle('J14')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFF0CFA5');
$activeWorksheet->getStyle('K14')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFF0CFA5');
$activeWorksheet->getStyle('L14')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFF0CFA5');
$activeWorksheet->getStyle('M14')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFF0CFA5');
$activeWorksheet->getStyle('N14')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFE84A3C');
$activeWorksheet->getStyle('O14')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFE84A3C');
$activeWorksheet->getStyle('P14')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('Q14')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF8AB1CC');
$activeWorksheet->getStyle('R14')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF4B78AA');
$activeWorksheet->getStyle('S14')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF4B78AA');
$activeWorksheet->getStyle('T14')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
 
// ---- Row 15 ----
$activeWorksheet->getStyle('C15')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('D15')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF4B78AA');
$activeWorksheet->getStyle('E15')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF4B78AA');
$activeWorksheet->getStyle('F15')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('G15')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFF0CFA5');
$activeWorksheet->getStyle('H15')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFF0CFA5');
$activeWorksheet->getStyle('I15')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFF0CFA5');
$activeWorksheet->getStyle('J15')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFF0CFA5');
$activeWorksheet->getStyle('K15')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('L15')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFF0CFA5');
$activeWorksheet->getStyle('M15')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFF0CFA5');
$activeWorksheet->getStyle('N15')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFF0CFA5');
$activeWorksheet->getStyle('O15')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('P15')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF4B78AA');
$activeWorksheet->getStyle('Q15')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF4B78AA');
$activeWorksheet->getStyle('R15')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF4B78AA');
$activeWorksheet->getStyle('S15')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
 
// ---- Row 16 ----
$activeWorksheet->getStyle('D16')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('E16')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF4B78AA');
$activeWorksheet->getStyle('F16')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF4B78AA');
$activeWorksheet->getStyle('G16')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('H16')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFF0CFA5');
$activeWorksheet->getStyle('I16')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFF0CFA5');
$activeWorksheet->getStyle('J16')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFF0CFA5');
$activeWorksheet->getStyle('K16')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFF0CFA5');
$activeWorksheet->getStyle('L16')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFF0CFA5');
$activeWorksheet->getStyle('M16')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFF0CFA5');
$activeWorksheet->getStyle('N16')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFF0CFA5');
$activeWorksheet->getStyle('O16')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('P16')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF4B78AA');
$activeWorksheet->getStyle('Q16')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF4B78AA');
$activeWorksheet->getStyle('R16')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
 
// ---- Row 17 ----
$activeWorksheet->getStyle('E17')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('F17')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF4B78AA');
$activeWorksheet->getStyle('G17')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF4B78AA');
$activeWorksheet->getStyle('H17')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('I17')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('J17')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('K17')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('L17')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('M17')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('N17')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('O17')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF4B78AA');
$activeWorksheet->getStyle('P17')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF4B78AA');
$activeWorksheet->getStyle('Q17')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
 
// ---- Row 18 ----
$activeWorksheet->getStyle('F18')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('G18')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('H18')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('I18')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFC4221A');
$activeWorksheet->getStyle('J18')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('K18')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('L18')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('M18')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFC4221A');
$activeWorksheet->getStyle('N18')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('O18')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('P18')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
 
// ---- Row 19 ----
$activeWorksheet->getStyle('E19')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('F19')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFE8E4D9');
$activeWorksheet->getStyle('G19')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF8B4A1C');
$activeWorksheet->getStyle('H19')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('I19')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFE8E4D9');
$activeWorksheet->getStyle('J19')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFC4221A');
$activeWorksheet->getStyle('K19')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('L19')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFC4221A');
$activeWorksheet->getStyle('M19')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFE8E4D9');
$activeWorksheet->getStyle('N19')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('O19')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF8B4A1C');
$activeWorksheet->getStyle('P19')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFE8E4D9');
$activeWorksheet->getStyle('Q19')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
 
// ---- Row 20 ----
$activeWorksheet->getStyle('D20')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('E20')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFE8E4D9');
$activeWorksheet->getStyle('F20')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFE8E4D9');
$activeWorksheet->getStyle('G20')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('H20')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('I20')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFE8E4D9');
$activeWorksheet->getStyle('J20')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFE8E4D9');
$activeWorksheet->getStyle('K20')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('L20')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFE8E4D9');
$activeWorksheet->getStyle('M20')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFE8E4D9');
$activeWorksheet->getStyle('N20')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('O20')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('P20')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFE8E4D9');
$activeWorksheet->getStyle('Q20')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFE8E4D9');
$activeWorksheet->getStyle('R20')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
 
// ---- Row 21 ----
$activeWorksheet->getStyle('C21')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('D21')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFE8E4D9');
$activeWorksheet->getStyle('E21')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFE8E4D9');
$activeWorksheet->getStyle('F21')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('G21')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('H21')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFE8E4D9');
$activeWorksheet->getStyle('I21')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFE8E4D9');
$activeWorksheet->getStyle('J21')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFE8E4D9');
$activeWorksheet->getStyle('K21')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFE8E4D9');
$activeWorksheet->getStyle('L21')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFE8E4D9');
$activeWorksheet->getStyle('M21')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFE8E4D9');
$activeWorksheet->getStyle('N21')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFE8E4D9');
$activeWorksheet->getStyle('O21')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('P21')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('Q21')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFE8E4D9');
$activeWorksheet->getStyle('R21')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFE8E4D9');
$activeWorksheet->getStyle('S21')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
 
// ---- Row 22 ----
$activeWorksheet->getStyle('C22')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('D22')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFE8E4D9');
$activeWorksheet->getStyle('E22')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('F22')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('G22')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('H22')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFE8E4D9');
$activeWorksheet->getStyle('I22')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFE8E4D9');
$activeWorksheet->getStyle('J22')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFE8E4D9');
$activeWorksheet->getStyle('K22')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF4DB5CB');
$activeWorksheet->getStyle('L22')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFE8E4D9');
$activeWorksheet->getStyle('M22')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFE8E4D9');
$activeWorksheet->getStyle('N22')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFE8E4D9');
$activeWorksheet->getStyle('O22')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('P22')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('Q22')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('R22')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFE8E4D9');
$activeWorksheet->getStyle('S22')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
 
// ---- Row 23 ----
$activeWorksheet->getStyle('C23')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('D23')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFE8E4D9');
$activeWorksheet->getStyle('E23')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFE8E4D9');
$activeWorksheet->getStyle('F23')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('G23')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFE8E4D9');
$activeWorksheet->getStyle('H23')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFE8E4D9');
$activeWorksheet->getStyle('I23')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFE8E4D9');
$activeWorksheet->getStyle('J23')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFE8E4D9');
$activeWorksheet->getStyle('K23')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFE8E4D9');
$activeWorksheet->getStyle('L23')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFE8E4D9');
$activeWorksheet->getStyle('M23')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFE8E4D9');
$activeWorksheet->getStyle('N23')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFE8E4D9');
$activeWorksheet->getStyle('O23')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFE8E4D9');
$activeWorksheet->getStyle('P23')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('Q23')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFE8E4D9');
$activeWorksheet->getStyle('R23')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFE8E4D9');
$activeWorksheet->getStyle('S23')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
 
// ---- Row 24 ----
$activeWorksheet->getStyle('C24')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('D24')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFE8E4D9');
$activeWorksheet->getStyle('E24')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('F24')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('G24')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF93AEC0');
$activeWorksheet->getStyle('H24')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFE8E4D9');
$activeWorksheet->getStyle('I24')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFE8E4D9');
$activeWorksheet->getStyle('J24')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFE8E4D9');
$activeWorksheet->getStyle('K24')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF93AEC0');
$activeWorksheet->getStyle('L24')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFE8E4D9');
$activeWorksheet->getStyle('M24')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFE8E4D9');
$activeWorksheet->getStyle('N24')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFE8E4D9');
$activeWorksheet->getStyle('O24')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF93AEC0');
$activeWorksheet->getStyle('P24')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('Q24')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('R24')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFE8E4D9');
$activeWorksheet->getStyle('S24')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
 
// ---- Row 25 ----
$activeWorksheet->getStyle('C25')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('D25')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('E25')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('F25')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('G25')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('H25')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF93AEC0');
$activeWorksheet->getStyle('I25')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF93AEC0');
$activeWorksheet->getStyle('J25')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF93AEC0');
$activeWorksheet->getStyle('K25')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('L25')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF93AEC0');
$activeWorksheet->getStyle('M25')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF93AEC0');
$activeWorksheet->getStyle('N25')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF93AEC0');
$activeWorksheet->getStyle('O25')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('P25')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('Q25')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('R25')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('S25')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
 
// ---- Row 26 ----
$activeWorksheet->getStyle('C26')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('D26')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('E26')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('F26')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('G26')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFE8E4D9');
$activeWorksheet->getStyle('H26')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('I26')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('J26')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFE8E4D9');
$activeWorksheet->getStyle('K26')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF8B4A1C');
$activeWorksheet->getStyle('L26')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFE8E4D9');
$activeWorksheet->getStyle('M26')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('N26')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('O26')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFE8E4D9');
$activeWorksheet->getStyle('P26')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('Q26')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('R26')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('S26')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
 
// ---- Row 27 ----
$activeWorksheet->getStyle('C27')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('D27')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('E27')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('F27')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('G27')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('H27')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFE8E4D9');
$activeWorksheet->getStyle('I27')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFE8E4D9');
$activeWorksheet->getStyle('J27')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFE8E4D9');
$activeWorksheet->getStyle('K27')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('L27')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFE8E4D9');
$activeWorksheet->getStyle('M27')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFE8E4D9');
$activeWorksheet->getStyle('N27')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFE8E4D9');
$activeWorksheet->getStyle('O27')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('P27')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('Q27')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('R27')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('S27')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
 
// ---- Row 28 ----
$activeWorksheet->getStyle('C28')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('D28')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('E28')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('F28')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('G28')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('H28')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFE8E4D9');
$activeWorksheet->getStyle('I28')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFE8E4D9');
$activeWorksheet->getStyle('J28')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFE8E4D9');
$activeWorksheet->getStyle('K28')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('L28')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFE8E4D9');
$activeWorksheet->getStyle('M28')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFE8E4D9');
$activeWorksheet->getStyle('N28')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFE8E4D9');
$activeWorksheet->getStyle('O28')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('P28')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('Q28')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('R28')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
$activeWorksheet->getStyle('S28')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FF110F0B');
 
// ---- Download in the browser ----
header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
header('Content-Disposition: attachment; filename="pixel_art.xlsx"');
header('Cache-Control: max-age=0');
 
$writer = new Xlsx($spreadsheet);
$writer->save('php://output');
exit;