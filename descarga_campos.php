<?php
  require_once 'lib/PHPExcel/IOFactory.php';
  $archivo = "documentIdsList.xlsx";
  $objPHPExcel = PHPExcel_IOFactory::load($archivo);
  $objPHPExcel->setActiveSheetIndex(0);
  $numRows = $objPHPExcel->setActiveSheetIndex(0)->getHighestRow();
  for ($i = 2; $i <= $numRows; $i++){

    $dato = $objPHPExcel->getActiveSheet()->getCell('A'.$i)->getCalculatedValue();
    $documents = 'documents/'.$dato.'.xlsx';

    if (!file_exists($documents)) {
      $objPHPExcel = PHPExcel_IOFactory::load($documents);
      $objPHPExcel->setActiveSheetIndex(0);
      $numRows = $objPHPExcel->setActiveSheetIndex(0)->getHighestRow();

    }
  }
 ?>
