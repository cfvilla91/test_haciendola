<?php

// LIBRERÍA PARA EXPORTAR A EXCEL EN PHP
require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

// NOMBRES DE LAS CABECERAS DE EXCEL
$excelHeaders = array(
    'A',  'B',  'C',  'D',  'E',  'F',  'G',  'H',  'I',  'J',  'K',  'L',  'M',  'N',  'O',  'P',  'Q',
    'R',  'S',  'T',  'U',  'V',  'W',  'X',  'Y',  'Z',  'AA',  'AB',  'AC',  'AD',  'AE',  'AF',  'AG',
    'AH',  'AI',  'AJ',  'AK',  'AL',  'AM',  'AN',  'AO',  'AP',  'AQ',  'AR',  'AS',  'AT',  'AU',  'AV',
    'AW',  'AX',  'AY',  'AZ',  'BA',  'BB',  'BC',  'BD',  'BE',  'BF',  'BG',  'BH',  'BI',  'BJ',  'BK',
    'BL',  'BM',  'BN',  'BO',  'BP',  'BQ',  'BR',  'BS',  'BT',  'BU',  'BV',  'BW',  'BX',  'BY',  'BZ',
    'CA',  'CB',  'CC',  'CD',  'CE',  'CF',  'CG',  'CH',  'CI',  'CJ',  'CK',  'CL',  'CM',  'CN',  'CO',
    'CP',  'CQ',  'CR',  'CS',  'CT',  'CU',  'CV',  'CW',  'CX',  'CY',  'CZ'
);
error_reporting(E_ERROR | E_PARSE);
// LLAMADA HTTP AL SERVICIO DE SHOPIFY
$ch = curl_init();
curl_setopt($ch, CURLOPT_SSL_VERIFYPEER, false);
curl_setopt($ch, CURLOPT_RETURNTRANSFER, true);
curl_setopt($ch, CURLOPT_URL, 'https://cfba29612729d671122bc7594d7a0a6b:b31ed7b51a7b1338e06bcf0a0c5b6a90@testhaciendola.myshopify.com/admin/api/2020-01/orders.json');
$result = curl_exec($ch);
curl_close($ch);

$obj = json_decode($result, true);
// ORDERS OBTENIDAS DE LA API
$orders = $obj['orders'];

// INICIALIZO EL EXCEL
$spreadsheet = new Spreadsheet();
$spreadsheet->getProperties()
    ->setCreator('Cristóbal Félix Villa Rojas')
    ->setLastModifiedBy('Cristóbal Félix Villa Rojas')
    ->setTitle('Test para Haciéndola')
    ->setSubject('Test para Haciéndola')
    ->setDescription('Test para Haciéndola');

// INICIALIZO LA WORKSHEET DE EXCEL
$sheet = $spreadsheet->getActiveSheet();
$sheet->setTitle('Orders');

// PRIMERO ITERO LAS LLAVES DEL PRIMER OBJETO DEL ARRAY PARA CONSTRUIR LAS CABECERAS DEL EXCEL
$i = 0;
foreach ($orders[0] as $key => $value) {
    $sheet->setCellValue($excelHeaders[$i] . 1, $key);
    $i++;
}

// AHORA RECORREMOS LAS ORDERS OBTENIDAS DE LA API Y LAS ESCRIBIMOS EN EL EXCEL
$i = 2;
foreach ($orders as $order) {
    $j = 0;
    foreach ($order as $key => $value) {
        $sheet->setCellValue($excelHeaders[$j] . $i, $value);
        $j++;
    }
    $i++;
}

// ESTO CREA UN BUFFER PARA GUARDAR ESTA INFORMACIÓN EN UN ARCHIVO XLSX
$writer = new Xlsx($spreadsheet);

// ESTO GUARDA TODO EN UN ARCHIVO EXCEL EN EL SERVIDOR QUE EJECUTA ESTE SCRIPT
$writer->save('prueba_tecnica_haciendola_cfvilla.xlsx');

echo 'FIN DEL SCRIPT';
