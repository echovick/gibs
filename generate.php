<?php
require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Reader\Xlsx;
use PhpOffice\PhpWord\TemplateProcessor;

if ($_SERVER['REQUEST_METHOD'] == 'POST') {
    // Get uploaded files
    $wordFile  = $_FILES['wordFile']['tmp_name'];
    $excelFile = $_FILES['excelFile']['tmp_name'];

    // Load the Excel file
    $reader = new Xlsx();
    $reader->setReadDataOnly(true);
    $spreadsheet = $reader->load($excelFile);
    $sheet       = $spreadsheet->getActiveSheet();

    $rowCount = 0;
    foreach ($sheet->getRowIterator() as $index => $row) {
        if ($rowCount >= 5) {
            break;
        }

        $cellIterator = $row->getCellIterator();
        $cellIterator->setIterateOnlyExistingCells(false);
        $rowData = [];
        foreach ($cellIterator as $cell) {
            $rowData[] = $cell->getValue();
        }

        if (is_string($rowData[0])) {
            continue;
        }

        // Create a new TemplateProcessor instance for each row
        $templateProcessor = new TemplateProcessor($wordFile);
        $templateProcessor->setValue('Sno', $rowData[0] ?? '');
        $templateProcessor->setValue('AgentBroker_Name', $rowData[7] ?? '');
        $templateProcessor->setValue('Address', $rowData[14] ?? '');
        $templateProcessor->setValue('Insured_Name', $rowData[4] ?? '');
        $templateProcessor->setValue('Policy_No', $rowData[3] ?? '');
        $templateProcessor->setValue('Expiry_Date', $rowData[22] ?? '');
        $templateProcessor->setValue('SumInsured', $rowData[17] ?? '');
        $templateProcessor->setValue('Gross_Premium', $rowData[18] ?? '');
        $templateProcessor->setValue('Class__of_Business', $rowData[10] ?? '');

        // Generate a unique filename for each document
        $outputFile = 'filled_template_' . ($rowCount + 1) . '.docx';
        $templateProcessor->saveAs($outputFile);

        echo "Generated: $outputFile <br>";
        $rowCount++;
    }
}
