<?php

use PhpOffice\PhpWord\SimpleType\Jc;
use PhpOffice\PhpWord\Shared\Converter;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\IOFactory;

require '/Users/Admin/Documents/GitHub/prueba-codigo-actas/vendor/autoload.php';
include_once './ActaEntregaComputadores.php';

if($_SERVER['REQUEST_METHOD']=== 'POST'){
    if(isset($_POST['actualizarExcel'])){

        if($tipoEquipo === "escritorio"){

            $nombreUsuario = $_POST['nombre'];
            $cedulaUsuario = $_POST['cedula'];
            $tipoEquipo = $_POST['tipoEquipo'];
            $estadoEquipo = $_POST['usoEquipo'];
            $nombreEquipo = $_POST['nombreEquipo'];
            $nombreProcesador = $_POST['procesadorEquipo'];
            $almacenamientoEquipo = $_POST['almacenamientoEquipo'];
            $RAMEquipo = $_POST['memoriaRAM'];
            $marcaEquipo = $_POST['marcaEquipo'];
            $modeloEquipo = $_POST['modeloEquipo'];
            $serialEquipo = $_POST['serialEquipo'];
            $versionSO = $_POST['versionSO'];
        

            $columnaNombrePersona ="O" ;
    $columnaCedulaPersona ="Z" ;
    $columnaTipoDeEquipo ="Q" ;
    $columnaTipoDeEstado ="AB" ;
    $columnaNombreEquipo = "A" ;
    $columnaNombreProcesadorEquipo = "B" ;
    $columnaAlmacenamientoEquipo = "C" ;
    $columnaRAMEquipo = "D";
    $columnaMarcaEquipo = "E";
    $columnaModeloEquipo = "F";
    $columnaSerialEquipo = "G";
    $columnaVersionSO = "H";
    $columnaAsignado = "AC";

    $hojaCalculo = IOFactory::load('C:/Users/Admin/Downloads/01_INVENTARIO PLANTA NORTE 2023.xlsx');

    $elemento = $hojaCalculo->getActiveSheet();

    $cellIterator = $elemento->getRowIterator();

    foreach ($elemento->getRowIterator() as $row) {
        foreach ($row->getCellIterator() as $cell) {
            $cellValue = $cell->getValue();
        
            if ($cellValue == $serialEquipo) {
                $foundCell = $cell->getCoordinate();
                list($columna, $fila) = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::coordinateFromString($foundCell);
                break 2;  // Salir de ambos bucles
            }

        }
    }

    if (!isset($foundCell)) {
        $fila = 1;
        while (!empty($elemento->getCell($columnaNombreEquipo . $fila)->getValue())) {
        $fila++;
    }
    }
    

    $elemento->setCellValue($columnaNombrePersona . $fila, $nombreUsuario);
    $elemento->setCellValue($columnaCedulaPersona . $fila, $cedulaUsuario);
    $elemento->setCellValue($columnaNombreEquipo . $fila, $nombreEquipo);
    $elemento->setCellValue($columnaNombreProcesadorEquipo . $fila, $nombreProcesador);
    $elemento->setCellValue($columnaAlmacenamientoEquipo . $fila, $almacenamientoEquipo);
    $elemento->setCellValue($columnaRAMEquipo . $fila, $RAMEquipo);
    $elemento->setCellValue($columnaMarcaEquipo . $fila, $marcaEquipo);
    $elemento->setCellValue($columnaSerialEquipo . $fila, $serialEquipo);
    $elemento->setCellValue($columnaModeloEquipo . $fila, $modeloEquipo);
    $elemento->setCellValue($columnaVersionSO . $fila, $versionSO);
    $elemento->setCellValue($columnaTipoDeEquipo . $fila, $tipoEquipo);
    $elemento->setCellValue($columnaTipoDeEstado . $fila, $estadoEquipo);
    $elemento->setCellValue($columnaAsignado . $fila, "EN USO");

    $writer = IOFactory::createWriter($hojaCalculo, 'Xlsx');
    $writer->save('C:/Users/Admin/Downloads/01_INVENTARIO PLANTA NORTE 2023.xlsx');

        }elseif($tipoEquipo === "portatil"){


    $nombreUsuario = $_POST['nombre'];
    $cedulaUsuario = $_POST['cedula'];
    $tipoEquipo = $_POST['tipoEquipo'];
    $estadoEquipo = $_POST['usoEquipo'];
    $nombreEquipo = $_POST['nombreEquipo'];
    $nombreProcesador = $_POST['procesadorEquipo'];
    $almacenamientoEquipo = $_POST['almacenamientoEquipo'];
    $RAMEquipo = $_POST['memoriaRAM'];
    $marcaEquipo = $_POST['marcaEquipo'];
    $modeloEquipo = $_POST['modeloEquipo'];
    $serialEquipo = $_POST['serialEquipo'];
    $versionSO = $_POST['versionSO'];

    $columnaCedulaPersona ="Z" ;
    $columnaTipoDeEquipo ="Q" ;
    $columnaTipoDeEstado ="AB" ;
    $columnaNombreEquipo = "A" ;
    $columnaNombreProcesadorEquipo = "B" ;
    $columnaAlmacenamientoEquipo = "C" ;
    $columnaRAMEquipo = "D";
    $columnaMarcaEquipo = "E";
    $columnaModeloEquipo = "F";
    $columnaSerialEquipo = "G";
    $columnaVersionSO = "H";
    $columnaAsignado = "AC";

    $hojaCalculo = IOFactory::load('C:/Users/Admin/Downloads/01_INVENTARIO PLANTA NORTE 2023.xlsx');

    $elemento = $hojaCalculo->getActiveSheet();

    $cellIterator = $elemento->getRowIterator();

    foreach ($elemento->getRowIterator() as $row) {
        foreach ($row->getCellIterator() as $cell) {
            $cellValue = $cell->getValue();
        
            if ($cellValue == $serialEquipo) {
                $foundCell = $cell->getCoordinate();
                list($columna, $fila) = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::coordinateFromString($foundCell);
                break 2;  // Salir de ambos bucles
            }

        }
    }

    if (!isset($foundCell)) {
        $fila = 1;
        while (!empty($elemento->getCell($columnaNombreEquipo . $fila)->getValue())) {
        $fila++;
    }
    }
    

    $elemento->setCellValue($columnaNombrePersona . $fila, $nombreUsuario);
    $elemento->setCellValue($columnaCedulaPersona . $fila, $cedulaUsuario);
    $elemento->setCellValue($columnaNombreEquipo . $fila, $nombreEquipo);
    $elemento->setCellValue($columnaNombreProcesadorEquipo . $fila, $nombreProcesador);
    $elemento->setCellValue($columnaAlmacenamientoEquipo . $fila, $almacenamientoEquipo);
    $elemento->setCellValue($columnaRAMEquipo . $fila, $RAMEquipo);
    $elemento->setCellValue($columnaMarcaEquipo . $fila, $marcaEquipo);
    $elemento->setCellValue($columnaSerialEquipo . $fila, $serialEquipo);
    $elemento->setCellValue($columnaModeloEquipo . $fila, $modeloEquipo);
    $elemento->setCellValue($columnaVersionSO . $fila, $versionSO);
    $elemento->setCellValue($columnaTipoDeEquipo . $fila, $tipoEquipo);
    $elemento->setCellValue($columnaTipoDeEstado . $fila, $estadoEquipo);
    $elemento->setCellValue($columnaAsignado . $fila, "EN USO");

    $writer = IOFactory::createWriter($hojaCalculo, 'Xlsx');
    $writer->save('C:/Users/Admin/Downloads/01_INVENTARIO PLANTA NORTE 2023.xlsx');


        }


    }
}

?>