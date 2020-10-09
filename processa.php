<?php
set_time_limit(0);
ini_set('memory_limit', '-1');
error_reporting(E_ALL);

require 'vendor/autoload.php'; //autoload do projeto

function lerXlsx($arq)
{
  $reader = new \PhpOffice\PhpSpreadsheet\Reader\Xlsx();
  $spreadsheet = $reader->load($arq);
  $worksheet = $spreadsheet->getActiveSheet(); //retornando a aba ativa
  return $worksheet->toArray();
}

function createColumns($max = null)
{
  $column = range('A', 'Z');
  if (empty($max))
    return $column;
  $first = $max[0];
  if (strlen($max) > 1)
    foreach ($column as $letter) {
      foreach (range('A', 'Z') as $secondLetter) {
        $column[] = $letter . $secondLetter;
        if ($letter . $secondLetter == $max)
          break;
      }
      if ($first == $letter)
        return $column;
    }
  else return range('A', $max);
}

$arquivo1 = lerXlsx('integrados.xlsx');
$arquivo2x1 = lerXlsx('copia lista quero ser socio.xlsx');
$columns = createColumns();

$listaBusca = array();

$row = 0;
foreach ($arquivo1 as $linha) {
  if ($row != 0)
    $listaBusca[$row] = trim($linha[0]);
  $row++;
}

// Escrita do arquivo
$spreadsheet = new PhpOffice\PhpSpreadsheet\Spreadsheet; //instanciando uma nova planilha
$sheet = $spreadsheet->getActiveSheet(); //retornando a aba ativa

$numRow = 1;
$row = 2;
$export[] = $arquivo2x1[0];
$columns = range('A', 'Z');

for ($i = 0; $i < count($arquivo2x1[0]); $i++) {
  $sheet->setCellValue($columns[$i] . '1', $arquivo2x1[0][$i]);
}

foreach ($arquivo2x1 as $linha) {
  if ($numRow != 1)
    if (!array_search(trim($linha[0]), $listaBusca)) {
      for ($i = 0; $i < count($linha); $i++) {
        $sheet->setCellValue($columns[$i] . $row, $linha[$i]);
      }
      $row++;
    }
  $numRow++;
}

$writer = new \PhpOffice\PhpSpreadsheet\Writer\Xlsx($spreadsheet); //classe que salvará a planilha em .csv
$writer->save('/Users/vinifdiniz/Downloads/laporte.xlsx');
