<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <link rel="stylesheet" href="/index.css">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>FILTRO Y REPORTES</title>
</head>
<body>
    <section class="1">
    <form action="filtro.php" method="post">
        <label for="busqueda">QUE EQUIPOS CON CARACTERISTICA EN COMUN ESTAS BUSCANDO</label>
        <input type="search" name="busqueda">

<section class="seccion1" id='label1'>
    <input type="checkbox" name="asignado[]" value="EN USO" id="checkbox1">
        <label for="asignado">EN USO</label>
</section>

        <section class='seccion1'>
        <input type="checkbox" name="asignado[]" value="EN BODEGA"> 
        <label for="asignado">EN BODEGA</label>
</section>        
        <input type="submit" value="filtrar">
    </form>
    </section class="2">
</body>
</html>


<?php

use PhpOffice\PhpWord\Writer\Word2007;
use PhpOffice\PhpWord\SimpleType\Jc;
use PhpOffice\PhpWord\Style\Font;
use PhpOffice\PhpWord\Shared\Converter;

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpWord\Element\TextRun;
use Mikehaertl\ShellCommand\Command;
use Mpdf\Mpdf;
use PhpOffice\PhpWord\Writer\HTML;



if($_SERVER['REQUEST_METHOD']=== 'POST'){

    $caracteristica = $_POST['caracteristica'];
    $asignado[] = $_POST['asignado'];

if(empty($caracteristica)){
    $caracteristica = "";
}

if(empty($asignado)){
    $asignado = "";
}
    if(isset($caracteristica) or isset($asignado) ){

        $caracteristica = $_POST['busqueda'];

        $asignado[] = $_POST['asignado'];

    }elseif(empty($caracteristica) or empty($asignado)){

    $spreadsheet = new Spreadsheet();
    $hojaCalculo = IOFactory::load('C:/Users/Admin/Downloads/01_INVENTARIO PLANTA NORTE 2023.xlsx');

    $elemento = $hojaCalculo->getActiveSheet();

    $hojita = $hojaCalculo->getSheet(1);

    $cellIterator = $elemento->getRowIterator();



    foreach ($hojita->getRowIterator() as $row) {
        foreach ($row->getCellIterator() as $cell) {
            $cellValue = $cell->getValue();
                $foundCell = $cell->getCoordinate();
                list($columna, $fila) = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::coordinateFromString($foundCell);
            }

        }
    }

    if (!isset($foundCell)) {
echo "<h2>No se han encontrado dispositivos con estas caracteristicas </h2>";
    }
    }
    




?>