<?php
require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpWord\PhpWord;
use PhpOffice\PhpWord\Shared\Html;

if ($_SERVER['REQUEST_METHOD'] == 'POST') {
    // Get uploaded files
    $wordFile = $_FILES['wordFile']['tmp_name'];
    $excelFile = $_FILES['excelFile']['tmp_name'];

    // Load the Excel file
    $spreadsheet = IOFactory::load($excelFile);
    $sheet = $spreadsheet->getActiveSheet();

    // Extract data from Excel
    $data = [];
    foreach ($sheet->getRowIterator() as $row) {
        $cellIterator = $row->getCellIterator();
        $cellIterator->setIterateOnlyExistingCells(false);
        $rowData = [];
        foreach ($cellIterator as $cell) {
            $rowData[] = $cell->getValue();
        }
        $data[] = $rowData;
    }

    // Load the Word template
    $phpWord = new PhpWord();
    $templateProcessor = $phpWord->loadTemplate($wordFile);

    // Fill the Word template with data
    foreach ($data as $index => $row) {
        $templateProcessor->setValue('Sno#' . ($index + 1), $row[0]);
        $templateProcessor->setValue('AgentBroker_Name#' . ($index + 1), $row[1]);
        $templateProcessor->setValue('Address#' . ($index + 1), $row[2]);
        $templateProcessor->setValue('Insured_Name#' . ($index + 1), $row[3]);
        $templateProcessor->setValue('Policy_No#' . ($index + 1), $row[4]);
        $templateProcessor->setValue('Expiry_Date#' . ($index + 1), $row[5]);
        $templateProcessor->setValue('SumInsured#' . ($index + 1), $row[6]);
        $templateProcessor->setValue('Gross_Premium#' . ($index + 1), $row[7]);
        $templateProcessor->setValue('Class__of_Business#' . ($index + 1), $row[8]);
    }

    // Save the filled Word document
    $outputFile = 'filled_template.docx';
    $templateProcessor->saveAs($outputFile);

    echo "Word document has been filled and saved as '$outputFile'.";
}
?>
