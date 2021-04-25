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
      $url = 'https://downloads.regulations.gov/'.$dato.'/attachment_1.xlsx';

      if(!url_exists($url)){
          $url = 'https://downloads.regulations.gov/'.$dato.'/attachment_2.xlsx';
      }

      $file   = file($url);
      $result = file_put_contents($documents, $file);
    }
  }
  echo "Descarga finalizada";

  function url_exists($url) {

     $ch = @curl_init($url);
     @curl_setopt($ch, CURLOPT_HEADER, TRUE);
     @curl_setopt($ch, CURLOPT_NOBODY, TRUE);
     @curl_setopt($ch, CURLOPT_FOLLOWLOCATION, FALSE);
     @curl_setopt($ch, CURLOPT_RETURNTRANSFER, TRUE);
     $status = array(); preg_match('/HTTP\/.* ([0-9]+) .*/', @curl_exec($ch) , $status);
     return ($status[1] == 200);

  }
 ?>
