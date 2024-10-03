<?php

declare(strict_types=1);

// require __DIR__ . '/../../vendor/autoload.php';
namespace Src\Classes;

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;


class Excel
{
    private $spreadsheet;
    private $productsheet;
    public function __construct($creator, $title, $subject, $description)
    {
        $this->spreadsheet = new Spreadsheet();
        $this->spreadsheet->getProperties()
            ->setCreator($creator)
            ->setTitle($title)
            ->setSubject($subject)
            ->setDescription($description);
    }

    public function getSpreadsheet($titleSheet, array $tablesNames)
    {
        $this->productsheet = $this->spreadsheet->createSheet();
        $this->productsheet->setTitle($titleSheet);
        //criar  as tabelas com especificas letras
        foreach ($tablesNames as $key => $value) {
            $this->productsheet->setCellValue($key . 1, $value);
        }

        // fazer o foreach para cria as dimenscoes das colunas
        $productsheets = range(array_key_first($tablesNames), array_key_last($tablesNames));
        foreach ($productsheets as $column) {
            $this->productsheet->getColumnDimension($column)->setAutoSize(true);
        }
    }

    public function getProperties(array $tableProperties)
    {
        $rowcounterProducts = 2;
        $dataInsert = $this->productsheet;
        foreach ($tableProperties as $key => $row) {

            foreach ($row as $key => $value) {
                $dataInsert->SetCellValue($key . $rowcounterProducts, $value);
            }
            $rowcounterProducts++;
        }
    }

    public function saveAs($filename)
    {
        $writer = new Xlsx($this->spreadsheet);
        // $filename = "Export_Tjek_Excel_" . $month . "_" . $year;
        header('Content-Type: application/vnd.ms-excel');
        header('Content-Disposition: attachment;filename=' . $filename . ".xlsx");
        header('Cache-Control: max-age=0');
        $writer->save('php://output');
    }
}
