<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <link rel="stylesheet" href="../index.css">
    <link rel="shortcut icon" href="/IMAGENES/logoElis.png" type="image/x-icon">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Descarga archivo</title>
</head>
<body>
    <section class="seccion-descarga">
    </section>



<?php
session_start();
require "/Users/Admin/Documents/GitHub/prueba-codigo-actas/vendor/autoload.php";

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\IOFactory;
require_once '/Users/Admin/Documents/GitHub/prueba-codigo-actas/vendor/tecnickcom/tcpdf/tcpdf.php';
use Mpdf\Mpdf;


if($_SERVER['REQUEST_METHOD'] === 'POST'){

$matriz = $_SESSION['matriz'];

if(isset($_POST['excel'])){

 
    $spreadsheet = new Spreadsheet();

    $hoja = $spreadsheet->getActiveSheet();

   $rows = count($matriz);
   $cols = count($matriz[0]);
 //  echo $rows .'<br>';
//echo $cols . '<br>';
//echo $matriz[0][1];

//echo '<pre>';
//var_dump($matriz);
//echo '</pre>';

for ($row = 1; $row <= $rows; $row++) {
    for ($col = 1; $col <= $cols; $col++) {
        $hoja->setCellValueByColumnAndRow($col, $row, ($matriz[$row - 1][$col - 1]));
    }
}

    $writer = new Xlsx($spreadsheet);

    $fecha = date('d_m_y');
    $hora = date('H:M:S');
$guardado ='Reporte_Equipos_'.$fecha.'_'.$_SESSION['caracteristica'].'.xlsx';
    $writer->save($guardado);

    header("Location:$guardado");

      //  header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
      //  header('Content-Disposition: attachment;filename="' . $archivo . '"');
      //  header('Cache-Control: max-age=0');
        exit();
    }elseif(isset($_POST['pdfe'])){

        $conteoMatriz = count($matriz);

        echo '<table class="tabla-descarga">';
        echo '<thead>
                <tr>
                    <th>Nombre Equipo</th>
                    <th>Procesador</th>
                    <th>Disco Duro</th>
                    <th>RAM</th>
                    <th>Marca</th>
                    <th>Modelo</th>
                    <th>Serial</th> 
                    <th>Propio</th>
                    <th>Rentado</th>
                    <th>SoftWare</th>
                    <th>Usuario</th>
                    <th>Departamento</th>
                    <th>Equipo</th>
                    <th>Cortex Palo Alto</th>
                    <th>Estado</th>
                    <th>Asignado</th>
                </tr>
                </thead>';
            echo '<tbody>';
            foreach ($matriz as $indice => $valor) {
                // Ignora el último elemento
            echo '<tr>';
            echo '<td>' . $valor[0] . '</td>';
            echo '<td>' . $valor[1] . '</td>';
            echo '<td>' . $valor[2] . '</td>';
            echo '<td>' . $valor[3] . '</td>';
            echo '<td>' . $valor[4] . '</td>';
            echo '<td>' . $valor[5] . '</td>';
            echo '<td>' . $valor[6] . '</td>';
            echo '<td>' . $valor[11] . '</td>';
            echo '<td>' . $valor[12] . '</td>';
            echo '<td>' . $valor[13] . '</td>';
            echo '<td>' . $valor[14] . '</td>';
            echo '<td>' . $valor[15] . '</td>';
            echo '<td>' . $valor[16] . '</td>';
            echo '<td>' . $valor[23] . '</td>';
            echo '<td>' . $valor[27] . '</td>';
            $valor[28] = 'N/A';
            echo '<td>' . $valor[28] . '</td>';
            echo '</tr>';
        }
            echo '</tbody>';
            echo '<tfoot>
            <tr>
            <td colspan="29">
              <section class="seccion-botones">
              <button onclick="location=`./filtro.php`">Volver</button>
              <button class="botones" onclick="window.print()">Descarga PDF</button>
              </section>
            </td>
          </tr>
            </tfoot>';
            echo '</table>';



?>
</body>
</html>

<?php


    }
}
?>