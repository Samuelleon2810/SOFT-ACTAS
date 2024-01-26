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
    <form action="filtro.php" method="post" class="formulario-filtro">
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
        
        <section class='seccion2'>
        <input type="checkbox" name="asignado[]" value="NUEVO" id="checkbox5"> 
        <label for="asignado">ES NUEVO</label>
        </section>  

        <section class='seccion2'>
        <input type="checkbox" name="asignado[]" value="USADO" id="checkbox5"> 
        <label for="asignado">ES USADO</label>
        </section>  
        
        <input type="submit" value="filtrar">
    </form>


<?php
session_start();
ini_set('memory_limit', '2048M');
set_time_limit(300);
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


if(empty($asignado)){
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
}elseif($asignado){

foreach ($asignado as $opcion) {
    switch ($opcion){
        case "EN USO":

                foreach ($hojita->getRowIterator() as $row) {
                    $datosFila = [];

                    // Obtener el valor de la celda en la columna específica
                    $indiceCol = PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($tipoCol);
                    $cellValue = $hojita->getCellByColumnAndRow($indiceCol, $row->getRowIndex())->getValue();
                    $indiceCol2 = PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString("AC");
                    $cellValue2 =$hojita->getCellByColumnAndRow($indiceCol2, $row->getRowIndex())->getValue();
                    // Verificar si la celda coincide con la característica de búsqueda
                    if ($cellValue === $caracteristica and $cellValue2 === $opcion) {
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

            break;

        case "EN BODEGA":

            foreach ($hojita->getRowIterator() as $row) {
                $datosFila = [];

                // Obtener el valor de la celda en la columna específica
                $indiceCol = PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($tipoCol);
                $cellValue = $hojita->getCellByColumnAndRow($indiceCol, $row->getRowIndex())->getValue();
                $indiceCol2 = PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString("AC");
                $cellValue2 =$hojita->getCellByColumnAndRow($indiceCol2, $row->getRowIndex())->getValue();
                // Verificar si la celda coincide con la característica de búsqueda
                if ($cellValue === $caracteristica and $cellValue2 === $opcion) {
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

            break;

        case "PORTATIL":
            
            foreach ($hojita->getRowIterator() as $row) {
                $datosFila = [];

                // Obtener el valor de la celda en la columna específica
                $indiceCol = PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($tipoCol);
                $cellValue = $hojita->getCellByColumnAndRow($indiceCol, $row->getRowIndex())->getValue();
                $indiceCol2 = PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString("Q");
                $cellValue2 =$hojita->getCellByColumnAndRow($indiceCol2, $row->getRowIndex())->getValue();
                // Verificar si la celda coincide con la característica de búsqueda
                if ($cellValue === $caracteristica and $cellValue2 === $opcion) {
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

            break;
            
        case "ESCRITORIO":

            foreach ($hojita->getRowIterator() as $row) {
                $datosFila = [];

                // Obtener el valor de la celda en la columna específica
                $indiceCol = PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($tipoCol);
                $cellValue = $hojita->getCellByColumnAndRow($indiceCol, $row->getRowIndex())->getValue();
                $indiceCol2 = PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString("Q");
                $cellValue2 =$hojita->getCellByColumnAndRow($indiceCol2, $row->getRowIndex())->getValue();
                // Verificar si la celda coincide con la característica de búsqueda
                if ($cellValue === $caracteristica and $cellValue2 === $opcion) {
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

            break;

        case "RENTADO":
            break;

        case "PROPIO":
            break;

        case "NUEVO":

            foreach ($hojita->getRowIterator() as $row) {
                $datosFila = [];

                // Obtener el valor de la celda en la columna específica
                $indiceCol = PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($tipoCol);
                $cellValue = $hojita->getCellByColumnAndRow($indiceCol, $row->getRowIndex())->getValue();
                $indiceCol2 = PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString("AB");
                $cellValue2 =$hojita->getCellByColumnAndRow($indiceCol2, $row->getRowIndex())->getValue();
                // Verificar si la celda coincide con la característica de búsqueda
                if ($cellValue === $caracteristica and $cellValue2 === $opcion) {
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

            break;
            
        case "USADO":

            foreach ($hojita->getRowIterator() as $row) {
                $datosFila = [];

                // Obtener el valor de la celda en la columna específica
                $indiceCol = PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($tipoCol);
                $cellValue = $hojita->getCellByColumnAndRow($indiceCol, $row->getRowIndex())->getValue();
                $indiceCol2 = PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString("AB");
                $cellValue2 =$hojita->getCellByColumnAndRow($indiceCol2, $row->getRowIndex())->getValue();
                // Verificar si la celda coincide con la característica de búsqueda
                if ($cellValue === $caracteristica and $cellValue2 === $opcion) {
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

            break;
    }
}
}
foreach ($matriz as &$fila) {
    foreach ($fila as &$valor) {
        if ($valor === null || $valor === '') {
            $valor = 'N/A'; 
        }
    }
}

$_SESSION['matriz'] = $matriz;
    // Verificar si se encontraron datos
    if ($encontrado) {
        // Imprimir o procesar la matriz de datos encontrados
        echo '<table class="tabla">';
        echo '<thead>
                <tr>
                    <th>Nombre Equipo</th>
                    <th>Procesador</th>
                    <th>Disco Duro</th>
                    <th>RAM</th>
                    <th>Marca</th>
                    <th>Modelo</th>
                    <th>Serial</th>
                    <th>SO</th>
                    <th>Actualizacion</th>
                    <th>Office</th>
                    <th>Cuenta Empresarial</th>
                    <th>Propio</th>
                    <th>Rentado</th>
                    <th>SoftWare</th>
                    <th>Usuario</th>
                    <th>Departamento</th>
                    <th>Equipo</th>
                    <th>Teclado</th>
                    <th>Monitor</th>
                    <th>Mouse</th>
                    <th>Validacion</th>
                    <th>Cortex Palo Alto</th>
                    <th>Estado</th>
                    <th>Asignado</th>
                </tr>
              </thead>';
              echo '<tbody>';
        foreach ($matriz as $valor) {
            echo '<tr>';
            echo '<td>' . $valor[0] . '</td>';
            echo '<td>' . $valor[1] . '</td>';
            echo '<td>' . $valor[2] . '</td>';
            echo '<td>' . $valor[3] . '</td>';
            echo '<td>' . $valor[4] . '</td>';
            echo '<td>' . $valor[5] . '</td>';
            echo '<td>' . $valor[6] . '</td>';
            echo '<td>' . $valor[7] . '</td>';
            echo '<td>' . $valor[8] . '</td>';
            echo '<td>' . $valor[9] . '</td>';
            echo '<td>' . $valor[10] . '</td>';
            echo '<td>' . $valor[11] . '</td>';
            echo '<td>' . $valor[12] . '</td>';
            echo '<td>' . $valor[13] . '</td>';
            echo '<td>' . $valor[14] . '</td>';
            echo '<td>' . $valor[15] . '</td>';
            echo '<td>' . $valor[16] . '</td>';
            echo '<td>' . $valor[19] . '</td>';
            echo '<td>' . $valor[20] . '</td>';
            echo '<td>' . $valor[21] . '</td>';
            echo '<td>' . $valor[22] . '</td>';
            echo '<td>' . $valor[23] . '</td>';
            echo '<td>' . $valor[27] . '</td>';
            echo '<td>' . $valor[28] . '</td>';

            // Añade más celdas según sea necesario
        
            echo '</tr>';
        }
        echo '</tbody>';
        echo '<tfoot>
        <tr>
        <td colspan="29">
          <section class="seccion-botones">
            <form action="descarga.php" method="post">
              <input type="submit" name="excel" value="Descargar Excel" class="botones">
              <input type="submit" name="pdf" value="Descargar PDF" class="botones">
            </form>
          </section>
        </td>
      </tr>
        </tfoot>';
        echo '</table>';

    } else {
        echo "<h2>No se han encontrado dispositivos con estas características</h2>";
    }
}
?>

</section>

</body>
</html>