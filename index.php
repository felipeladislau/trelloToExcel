<?php

require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

// Ler o conteúdo do arquivo data.json
$jsonData = file_get_contents('data.json');

// Nome do arquivo gerado
$filename = 'cards.xlsx';

// Decodificar o JSON em um array associativo
$data = json_decode($jsonData, true);

// Dados dos cards do Trello
$cards = array();

foreach ($data['cards'] as $card) {

    $labels = "";
    foreach($card['labels'] as $label){
        $labels .= $label['name'] . ", ";
    }

    $cards[] = array(
        'creation_date' => $card['dateLastActivity'],
        'name' => $card['name'],
        'description' => $card['desc'],
        'label' => $labels
    );
    
}

// Criar um novo objeto Spreadsheet
$spreadsheet = new Spreadsheet();
$sheet = $spreadsheet->getActiveSheet();

// Definir os cabeçalhos das colunas
$sheet->setCellValue('A1', 'Creation Date');
$sheet->setCellValue('B1', 'Name');
$sheet->setCellValue('C1', 'Description');
$sheet->setCellValue('D1', 'Label');

// Iterar sobre os cartões e gravar os elementos em cada linha
$row = 2;
foreach ($cards as $card) {
    $creationDate = $card['creation_date'];
    $name = $card['name'];
    $description = $card['description'];
    $label = implode(',', $card['label']);

    $sheet->setCellValue('A' . $row, $creationDate);
    $sheet->setCellValue('B' . $row, $name);
    $sheet->setCellValue('C' . $row, $description);
    $sheet->setCellValue('D' . $row, $label);

    $row++;
}

// Criar um novo objeto Writer para salvar o arquivo XLSX
$writer = new Xlsx($spreadsheet);
$writer->save($filename);

// Definir o cabeçalho HTTP para download do arquivo
header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
header('Content-Disposition: attachment;filename="' . $filename . '"');
header('Cache-Control: max-age=0');
header('Expires: 0');
header('Pragma: public');

// Enviar o arquivo para o cliente
readfile($filename);

echo 'Arquivo XLSX gerado com sucesso!';
