<?php

// Author: Nermeen Shoman
// Date : 3 Jan 2017
// Company : QDRAH
// 
//  ************* URL *************************
//  
// localhost/aanaab-docs/?filename=testdocx.docx&precent=70
// filename: should be in same path of project
// precent: default is 5 if not sent in url
// 
//  ************* Installation *************************
//  
// should install imagemagick for image screen shot
// sudo apt-get install imagemagick
// should install libreoffice to convert word,ppt to pdf files
// sudo apt-get install libreoffice --no-install-recommends
//
// document should be in main folder except ppt files should be in pptx folder

$file = $_GET['filename']; // file

$precent = 5;
if (isset($_GET['precent'])) { // precent
    $precent = $_GET['precent'];
}

documentPreview($file, $precent);

function documentPreview($file, $precent) {
    $fileArr = explode('.', $file);
    $fileArrCount = count($fileArr);
    $fileExt = $fileArr[$fileArrCount - 1];
//    echo $fileExt; die; 
    switch ($fileExt) {
        case 'pdf':
            // Create and check permissions on end directory!
            splitPdf($file, 'pdf/', $precent);
            break;
        case 'docx':
        case 'doc':
            exec('libreoffice --headless --invisible --convert-to pdf output.pdf ' . $file);
            $file = substr($file, 0, strrpos($file, "."));
            splitPdf($file . ".pdf", 'pdf/', $precent);
            break;
        case 'pptx':
            exec('libreoffice --headless --invisible --convert-to pdf output.pdf testpptx/' . $file);
            $file = substr($file, 0, strrpos($file, "."));
            splitPdf($file . ".pdf", 'pdf/', $precent);
            
            break;
        case 'xlsx':
        case 'xls':
            readXls($file, $precent);
            break;
        default :
            echo "this file is not supported yet";
    }
}

/**
 * Split PDF file
 *
 * <p>Split all of the pages from a larger PDF files into
 * single-page PDF files.</p>
 *
 * @package FPDF required http://www.fpdf.org/
 * @package FPDI required http://www.setasign.de/products/pdf-php-solutions/fpdi/
 * @param string $filename The filename of the PDF to split
 * @param string $end_directory The end directory for split PDF (original PDF's directory by default)
 * @return void
 */
function splitPdf($filename, $end_directory = false, $precent = 5) {
    require_once('library/PDF/fpdf/fpdf.php');
    require_once('library/PDF/fpdi/fpdi.php');

    $end_directory = $end_directory ? $end_directory : './';
    $new_path = preg_replace('/[\/]+/', '/', $end_directory . '/' . substr($filename, 0, strrpos($filename, '/')));

    if (!is_dir($new_path)) {
        // Will make directories under end directory that don't exist
        // Provided that end directory exists and has the right permissions
        mkdir($new_path, 0777);
    }

    $pdf = new FPDI();
    $pagecount = $pdf->setSourceFile($filename); // How many pages?
    $divideBy = $precent / 100;
    $newCount = round($pagecount * $divideBy);

    if ($newCount < 1) {
        $newCount = 1;
    }
    $maxPages = 3;
    $limit = min($maxPages, $newCount);

    $new_pdf = new FPDI();
    // Split each $limit number of pages into a new PDF
    for ($i = 1; $i <= $limit; $i++) {
        $new_pdf->AddPage();
        $new_pdf->setSourceFile($filename);
        $new_pdf->useTemplate($new_pdf->importPage($i));
    }

    try {
        $new_filename = $end_directory . str_replace('.pdf', '', $filename) . "_" . $limit . ".pdf";
        $new_pdf->Output($new_filename, "F");
        exec('convert -density 300 -trim ' . $filename . '[0] -quality 100 ' . $new_filename . '.jpg');
        chmod($new_filename, 0777);  //changed to add the zero
        chmod($new_filename.".jpg", 0777);  //changed to add the zero
        echo "New PDF and thumbnail created " . $limit . " out of " . $pagecount . " pages.\n";
    } catch (Exception $e) {
        echo 'Caught exception: ', $e->getMessage(), "\n";
    }
}

function readXls($file, $precent = 5) {

    /** PHPExcel_IOFactory */
    include 'library/PHPExcelReader/Classes/PHPExcel/IOFactory.php';

    try {
        $objPHPExcel = PHPExcel_IOFactory::load($file);
    } catch (Exception $e) {
        die('Error loading file "' . pathinfo($file, PATHINFO_BASENAME) . '": ' . $e->getMessage());
    }

    $highestRow = 10;
    $lowestRow = 1;
    $sheet = 0;
// Iterate through each of the 20 worksheets
    foreach ($objPHPExcel->getWorksheetIterator() as $worksheet) {
        // Remove rows 3-60 (58 rows starting from row 3)
        // This will move the last row (61) up to row 3
        $highestRow = $worksheet->getHighestDataRow();
        $lowestRow = round($highestRow * $precent / 100);
        $worksheet->removeRow($lowestRow+1, $highestRow);
        echo "Excel sheet exported " . $lowestRow . " rows out of " . $highestRow . " rows in sheet" . $sheet . "<br/>";
        $sheet++;
    }

    $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
    $file = substr($file, 0, strrpos($file, "."));
    $objWriter->save($file . 'preview.xlsx');
    chmod($file . 'preview.xlsx', 0777);  //changed to add the zero
}
