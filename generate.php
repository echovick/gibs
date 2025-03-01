<?php
// Enable error reporting for debugging
error_reporting(E_ALL);
ini_set('display_errors', 1);

// Include Composer's autoloader
require __DIR__ . '/vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\IOFactory as ExcelIOFactory;
use PhpOffice\PhpWord\IOFactory as WordIOFactory;
use PhpOffice\PhpWord\PhpWord;

if (!isset($_FILES['excelFile']) || $_FILES['excelFile']['error'] !== UPLOAD_ERR_OK) {
    die("Error uploading file.");
}

$tmpFilePath = $_FILES['excelFile']['tmp_name'];

try {
    // Load the Excel file
    $spreadsheet = ExcelIOFactory::load($tmpFilePath);
    $sheet = $spreadsheet->getActiveSheet();
    $data = $sheet->toArray();

    // Remove header row
    array_shift($data);

    // Create a ZIP archive
    $zip = new ZipArchive();
    $zipFileName = 'renewal_notices.zip';
    if ($zip->open($zipFileName, ZipArchive::CREATE | ZipArchive::OVERWRITE) !== TRUE) {
        die("Cannot create ZIP file");
    }

    foreach ($data as $row) {
        // Load the Word template
        $templatePath = __DIR__ . '/RENEWAL_NOTICE_STENCIL.docx';
        if (!file_exists($templatePath)) {
            die("Template file not found: $templatePath");
        }

        $phpWord = WordIOFactory::load($templatePath);

        // Define placeholders and their replacements
        $placeholders = [
            'Sno' => $row[0] ?? '',
            'AgentBroker_Name' => $row[1] ?? '',
            'Address' => $row[2] ?? '',
            'Insured_Name' => $row[3] ?? '',
            'Policy_No' => $row[4] ?? '',
            'Expiry_Date' => $row[5] ?? '',
            'SumInsured' => $row[6] ?? '',
            'Gross_Premium' => $row[7] ?? '',
            'Class_of_Business' => $row[8] ?? '',
        ];

        // Replace placeholders in the Word document
        foreach ($phpWord->getSections() as $section) {
            foreach ($section->getElements() as $element) {
                if ($element instanceof \PhpOffice\PhpWord\Element\TextRun) {
                    foreach ($element->getElements() as $text) {
                        if ($text instanceof \PhpOffice\PhpWord\Element\Text) {
                            $text->setText(str_replace(
                                array_map(fn($key) => "«$key»", array_keys($placeholders)),
                                array_values($placeholders),
                                $text->getText()
                            ));
                        }
                    }
                }
            }
        }

        // Save the generated document
        $outputPath = "renewal_notice_{$row[0]}.docx";
        $phpWord->save($outputPath);

        // Add the document to the ZIP archive
        $zip->addFile($outputPath, basename($outputPath));
    }

    $zip->close();

    // Send the ZIP file to the user
    header('Content-Type: application/zip');
    header('Content-Disposition: attachment; filename="' . $zipFileName . '"');
    header('Content-Length: ' . filesize($zipFileName));
    readfile($zipFileName);

    // Clean up temporary files
    foreach ($data as $row) {
        $outputPath = "renewal_notice_{$row[0]}.docx";
        if (file_exists($outputPath)) {
            unlink($outputPath);
        }
    }
    unlink($zipFileName);
} catch (Exception $e) {
    die("Error processing the file: " . $e->getMessage());
}
?>