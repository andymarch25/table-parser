<?php

require 'vendor/autoload.php';

use DOMWrap\Document;
use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

# Input Data
$awb_code = 235;
$awb_number = 51737512;

# CURL query
$ch = curl_init("https://www.turkishcargo.com.tr/en/online-services/shipment-tracking?quick=True&awbInput={$awb_code}-{$awb_number}");
curl_setopt($ch, CURLOPT_RETURNTRANSFER, true);
$page = curl_exec($ch);
// $page = file_get_contents("page-example.html");

# Create object from HTML
$doc = new Document();
$doc->html($page);

# Create spreadsheet
$spreadsheet = new Spreadsheet();

/**
 * Filling spreadsheets
 */

# Table #1
$table = $doc->find('.shipment-tracking')->eq(0)->find('table');

$spreadsheet->setActiveSheetIndex(0);
$sheet = $spreadsheet->getActiveSheet();
$sheet->setTitle('Table 1');

$data = $table->find('tr');

# Fill data
$index = 1;
foreach ($data as $node => $value) {
    $cells = $value->find('td');
    $title = $cells[0]->textContent;
    $value = $cells[1]->textContent;
    $sheet->setCellValue('A' . $index, $title);
    $sheet->setCellValue('B' . $index, $value);
    $index++;
};


# Table #2
$table = $doc->find('.shipment-tracking')->eq(1)->find('table');

$spreadsheet->createSheet();
$spreadsheet->setActiveSheetIndex(1);
$sheet = $spreadsheet->getActiveSheet();
$sheet->setTitle('Table 2');

$headings = $table->find('th');
$data = $table->find('tr');

# Fill headings
$index = 1;
foreach ($headings as $node) {
    $title = $node->textContent;
    $letter = Coordinate::stringFromColumnIndex($index);
    $sheet->setCellValue($letter . '1', $title);
    $index++;
};

# Fill data
$index = 1;
foreach ($data as $node => $value) {
    $cells = $value->find('td');
    $indexCell = 1;
    foreach ($cells as $cell) {
        $letter = Coordinate::stringFromColumnIndex($indexCell);

        # Test for checked value
        # In this case, it doesn't work because "checked" attribute
        # and class "active" assigned by javascript

        if (count($cell->find('label'))) {
            if ($cell->find('label')->eq(0)->hasClass('active')) {
                $value = '+';
            } else {
                $value = '-';
            }
        } else {
            $value = $cell->textContent;
        }

        $sheet->setCellValue($letter . $index, $value);
        $indexCell++;
    }
    $index++;
};


# Table #3
$table = $doc->find('.shipment-tracking')->eq(3)->find('table');

$spreadsheet->createSheet();
$spreadsheet->setActiveSheetIndex(2);
$sheet = $spreadsheet->getActiveSheet();
$sheet->setTitle('Table 3');

$headings = $table->find('th');
$data = $table->find('tr');

# Fill headings
$index = 1;
foreach ($headings as $node) {
    $title = $node->textContent;
    $letter = Coordinate::stringFromColumnIndex($index);
    $sheet->setCellValue($letter . '1', $title);
    $index++;
};

# Fill data
$index = 1;
foreach ($data as $node => $value) {
    $cells = $value->find('td');
    $indexCell = 1;
    foreach ($cells as $cell) {
        $letter = Coordinate::stringFromColumnIndex($indexCell);
        $value = $cell->textContent;
        $sheet->setCellValue($letter . $index, $value);
        $indexCell++;
    }
    $index++;
};

# Write and save file
$writer = new Xlsx($spreadsheet);
$writer->save('output.xlsx');