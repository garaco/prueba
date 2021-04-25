<?php
require_once 'lib/PHPExcel/IOFactory.php';

    $estiloTitulo = array(
    'fill' => array(
      'type'  => PHPExcel_Style_Fill::FILL_SOLID,
      'color' => array('rgb' => 'BDD7EE'))
    );

    $estiloColumnas = array(
      'fill' => array(
        'type'  => PHPExcel_Style_Fill::FILL_SOLID,
        'color' => array('rgb' => 'F2F2F2'))
    );

    $objPHPExcel = new PHPExcel();
    $objPHPExcel->getProperties()->setCreator("")->setTitle("listOfField");
    $objPHPExcel->setActiveSheetIndex(0);
  	$objPHPExcel->getActiveSheet()->setTitle("Sheet1");

    $listExcel = PHPExcel_IOFactory::load("listOfField.xlsx");
    $listExcel->setActiveSheetIndex(0);
    $numRows = $listExcel->setActiveSheetIndex(0)->getHighestRow();

    $objPHPExcel->getActiveSheet()->getColumnDimension('A')->setWidth(2);

    $objPHPExcel->getActiveSheet()->getStyle('B2')->applyFromArray($estiloTitulo);
    $objPHPExcel->getActiveSheet()->setCellValue('B2', $listExcel->getActiveSheet()->getCell('B2')->getCalculatedValue());

    $objPHPExcel->getActiveSheet()->getStyle('C2')->applyFromArray($estiloTitulo);
    $objPHPExcel->getActiveSheet()->setCellValue('C2', $listExcel->getActiveSheet()->getCell('C2')->getCalculatedValue());

    $objPHPExcel->getActiveSheet()->getStyle('D2')->applyFromArray($estiloTitulo);
    $objPHPExcel->getActiveSheet()->setCellValue('D2', $listExcel->getActiveSheet()->getCell('D2')->getCalculatedValue());

    $objPHPExcel->getActiveSheet()->getStyle('E2')->applyFromArray($estiloTitulo);
    $objPHPExcel->getActiveSheet()->setCellValue('E2', $listExcel->getActiveSheet()->getCell('E2')->getCalculatedValue());

    $objPHPExcel->getActiveSheet()->getStyle('F2')->applyFromArray($estiloTitulo);
    $objPHPExcel->getActiveSheet()->setCellValue('F2', $listExcel->getActiveSheet()->getCell('F2')->getCalculatedValue());

    for ($j = 3; $j <= $numRows; $j++){

      $objPHPExcel->getActiveSheet()->getStyle('B'.$j)->applyFromArray($estiloColumnas);
      $objPHPExcel->getActiveSheet()->setCellValue('B'.$j, $listExcel->getActiveSheet()->getCell('B'.$j)->getCalculatedValue());

      $objPHPExcel->getActiveSheet()->getStyle('C'.$j)->applyFromArray($estiloColumnas);
      $objPHPExcel->getActiveSheet()->setCellValue('C'.$j, $listExcel->getActiveSheet()->getCell('C'.$j)->getCalculatedValue());

      $objPHPExcel->getActiveSheet()->getStyle('D'.$j)->applyFromArray($estiloColumnas);
      $objPHPExcel->getActiveSheet()->setCellValue('D'.$j, $listExcel->getActiveSheet()->getCell('D'.$j)->getCalculatedValue());

      $objPHPExcel->getActiveSheet()->getStyle('E'.$j)->applyFromArray($estiloColumnas);
      $objPHPExcel->getActiveSheet()->setCellValue('E'.$j, $listExcel->getActiveSheet()->getCell('E'.$j)->getCalculatedValue());

      $objPHPExcel->getActiveSheet()->getStyle('F'.$j)->applyFromArray($estiloColumnas);
      $objPHPExcel->getActiveSheet()->setCellValue('F'.$j, $listExcel->getActiveSheet()->getCell('F'.$j)->getCalculatedValue());

      $objPHPExcel->getActiveSheet()->setCellValue('G'.$j, $listExcel->getActiveSheet()->getCell('G'.$j)->getCalculatedValue());

    }

    $objPHPExcel->createSheet();
     $sheet = $objPHPExcel->setActiveSheetIndex(1);
     $sheet->setTitle("DB");

    $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
    $objWriter->save('listOfField1.xlsx');

// $archivo = "documentIdsList.xlsx";
// $documentIds = PHPExcel_IOFactory::load($archivo);
// $documentIds->setActiveSheetIndex(0);
// $numRows = $documentIds->setActiveSheetIndex(0)->getHighestRow();
// // for ($i = 2; $i <= $numRows; $i++){
// for ($i = 2; $i <= 3; $i++){
//
//   $dato = $documentIds->getActiveSheet()->getCell('A'.$i)->getCalculatedValue();
//   $documents = 'documents/'.$dato.'.xlsx';
//
//   $listExcel = PHPExcel_IOFactory::load("listOfField.xlsx");
//   $listExcel->setActiveSheetIndex(0);
//   $numRows = $listExcel->setActiveSheetIndex(0)->getHighestRow();
//
//
//   for ($j = 3; $j <= $numRows; $j++){
//     $celda = $listExcel->getActiveSheet()->getCell('D'.$j)->getCalculatedValue();
//     $temp = $listExcel->getActiveSheet()->getCell('D'.$j)->getCalculatedValue();
//
//     $cell = PHPExcel_IOFactory::load($documents);
//     $cell->setActiveSheetIndex(0);
//     $celdaDato = $cell->getActiveSheet()->getCell($celda)->getCalculatedValue();
//
//      $listExcel->getActiveSheet()->setCellValue('H'.$j, $celdaDato);
//      $temp="";
//
//     echo "$celdaDato <br>";
//
//   }
//


// }
?>
