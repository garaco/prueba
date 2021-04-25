<?php
set_time_limit(900000000000);
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

    $fila = array('A','B','C','D','E','F','G','H','I','J','K','L','M',
                  'N','O','P','Q','R','S','T','U','V','W','X','Y','Z');
    $complemento = array('','A','B','C');
    $objPHPExcel->createSheet();
     $sheet = $objPHPExcel->setActiveSheetIndex(1);
     $sheet->setTitle("DB");
     $cont=1;
     $comp=0;

     $objPHPExcel->getActiveSheet()->getStyle('A1')->applyFromArray($estiloTitulo);
     $objPHPExcel->getActiveSheet()->setCellValue('A1', "IdDocument");
     for ($j = 3; $j <= $numRows; $j++){

       if($cont>25){ $cont=0; $comp++; }
       $objPHPExcel->getActiveSheet()->getStyle($complemento[$comp].$fila[$cont].'1')->applyFromArray($estiloTitulo);
       $objPHPExcel->getActiveSheet()->setCellValue($complemento[$comp].$fila[$cont].'1', $listExcel->getActiveSheet()->getCell('B'.$j)->getCalculatedValue());
       $cont++;
     }
     $objPHPExcel->getActiveSheet()->getStyle($complemento[$comp].$fila[$cont].'1')->applyFromArray($estiloTitulo);


    $archivo = "documentIdsList.xlsx";
    $documentIds = PHPExcel_IOFactory::load($archivo);
    $documentIds->setActiveSheetIndex(0);
    $ids = $documentIds->setActiveSheetIndex(0)->getHighestRow();

    for ($i = 2; $i <= $ids; $i++){

      $dato = $documentIds->getActiveSheet()->getCell('A'.$i)->getCalculatedValue();
      $documents = 'documents/'.$dato.'.xlsx';

      $listExcel = PHPExcel_IOFactory::load("listOfField.xlsx");
      $listExcel->setActiveSheetIndex(0);
      $numRows = $listExcel->setActiveSheetIndex(0)->getHighestRow();

      $cont=1;
      $comp=0;

      $objPHPExcel->getActiveSheet()->setCellValue('A'.$i, $dato);
      for ($j = 3; $j <= $numRows; $j++){

          $celda = $listExcel->getActiveSheet()->getCell('D'.$j)->getCalculatedValue();

          $cell = PHPExcel_IOFactory::load($documents);
          $cell->setActiveSheetIndex(0);
          $celdaDato = $cell->getActiveSheet()->getCell($celda)->getCalculatedValue();

        if($cont>25){ $cont=0; $comp++; }
          $objPHPExcel->getActiveSheet()->setCellValue(
          $complemento[$comp].$fila[$cont].$i,
          $celdaDato
        );
        $cont++;
      }

      $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
      $objWriter->save('listOfField.xlsx');
    }

  $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
  $objWriter->save('listOfField.xlsx');

  echo "<p> Descarga de campos finalizada <p>";
?>
