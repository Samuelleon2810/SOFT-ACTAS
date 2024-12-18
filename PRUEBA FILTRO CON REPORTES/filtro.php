<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <link rel="stylesheet" href="/index.css">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>FILTRO Y REPORTES</title>
</head>
<body>
    <section class="ae">
    <form action="filtro.php" method="post">
    <label for="carac">POR CUAL CARACTERISTICA VAS A BUSCAR</label>
    <select name="carac" id="carac">
        <option value="F">MODELO</option>
        <option value="G">SERIAL</option>
        <option value="B">PROCESADOR</option>
        <option value="E">MARCA</option>
        <option value="O">NOMBRE PROPIETARIO</option>
        <option value="A">NOMBRE EQUIPO</option>
        <option value="">SIN ESPECIFICACION</option>
    </select>

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
        
        <section class='seccion2'>
        <input type="checkbox" name="asignado[]" value="PORTATIL"> 
        <label for="asignado">ES PORTATIL</label>
        </section>  

        <section class='seccion2'>
        <input type="checkbox" name="asignado[]" value="ESCRITORIO" id="checkbox4"> 
        <label for="asignado">ES ESCRITORIO</label>
        </section>  

        <section class='seccion2'>
        <input type="checkbox" name="asignado[]" value="RENTADO"> 
        <label for="asignado">ES RENTADO</label>
        </section>  

        <section class='seccion2'>
        <input type="checkbox" name="asignado[]" value="PROPIO" id="checkbox5"> 
        <label for="asignado">ES PROPIO</label>
        </section>  
        
        <input type="submit" value="filtrar">
    </form>


<?php

require '/Users/Admin/Desktop/prueba codigo actas/vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\IOFactory;


if ($_SERVER['REQUEST_METHOD'] === 'POST') {

    // Inicializar variables de búsqueda
    $caracteristica = isset($_POST['busqueda']) ? $_POST['busqueda'] : "";
    $asignado = isset($_POST['asignado']) ? $_POST['asignado'] : "";
    $tipoCol = isset($_POST['carac']) ? $_POST['carac'] : "";

    // Cargar hoja de cálculo
    $hojaCalculo = IOFactory::load('C:/Users/Admin/Downloads/01_INVENTARIO PLANTA NORTE 2023.xlsx');

    // Seleccionar la hoja de trabajo
    $hojita = $hojaCalculo->getSheet(1);

    // Inicializar variables para datos encontrados
    $encontrado = false;
    $matriz = [];

    // Iterar sobre las filas de la hoja de cálculo

    /*
    foreach ($hojita->getRowIterator() as $row) {
        $datosFila = [];
        // Iterar sobre las celdas de la fila
        foreach ($row->getCellIterator() as $cell) {
            $cellValue = $cell->getValue();

            // Verificar si la celda coincide con la característica de búsqueda
            if ($cellValue === $caracteristica){
                $encontrado = true;
                $foundCell = $cell->getCoordinate();

                // Obtener coordenadas de la celda encontrada
                list($columna, $fila) = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::coordinateFromString($foundCell);

                // Obtener datos de la fila encontrada
                for ($col = $columna; $col <= $hojita->getHighestDataColumn(); $col++) {
                    $valorCelda = $hojita->getCell($col . $fila)->getValue();
                    $datosFila[] = $valorCelda;
                }

                // Almacenar datos de la fila encontrada en la matriz
                $matriz[] = $datosFila;
                break; // Salir del bucle después de encontrar la primera coincidencia
            }
        }
    }
*/

foreach ($hojita->getRowIterator() as $row) {
    $datosFila = [];

    // Obtener el valor de la celda en la columna específica
    $indiceCol = PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($tipoCol);
    $cellValue = $hojita->getCellByColumnAndRow($indiceCol, $row->getRowIndex())->getValue();

    // Verificar si la celda coincide con la característica de búsqueda
    if ($cellValue === $caracteristica) {
        $encontrado = true; 

        // Iterar sobre las celdas de la fila y almacenar en el arreglo
        foreach ($row->getCellIterator() as $cell) {
            $valorCelda = $cell->getValue();
            $datosFila[] = $valorCelda;
        }

        // Almacenar datos de la fila encontrada en la matriz
        $matriz[] = $datosFila;
    }
}


    // Verificar si se encontraron datos
    if ($encontrado) {
        // Imprimir o procesar la matriz de datos encontrados
        print_r($matriz);
    } else {
        echo "<h2>No se han encontrado dispositivos con estas características</h2>";
    }
}
?>

</section>
</body>
</html>