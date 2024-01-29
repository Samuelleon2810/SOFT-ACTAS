<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <link rel="stylesheet" href="../index.css">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Descarga archivo</title>
</head>
<body>
    <section class="seccion-descarga">
    <h1>LA DESCARGA COMENZARA PRONTO</h1>
    <h2>si no empieza la descarga presiona el boton</h2>
    <form action="descarga.php" action="post">
        <input type="submit" value="Volver a descargar" class="botones">
        
    </form>
    <button onclick="location='./filtro.php'">Volver</button>
    </section>



<?php
session_start();
require "/Users/Admin/Documents/GitHub/prueba-codigo-actas/vendor/autoload.php";

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

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
    }

}elseif(isset($_POST['pdf'])){

    $spreadsheet = new Spreadsheet();

    $hoja = $spreadsheet->getActiveSheet();
    
    // Agregar datos a la hoja desde la matriz
    $hoja->fromArray($_SESSION['matriz'], null, 'A1');
    
    $stream = fopen('php://temp', 'r+');
    // Crear un objeto Writer para Xlsx (Excel 2007 y versiones posteriores)
    $writer = new Xlsx($spreadsheet);
    
    // Guardar el contenido del Excel en el manejador de flujo temporal
    $writer->save($stream);
    
    // Volver al principio del manejador de flujo temporal
    rewind($stream);

    $excelContent = stream_get_contents($stream);

    $mpdf->WriteHTML($excelContent);

    $mpdf = new Mpdf();

    $mpdf->AddPage();
    
    $mpdf->WriteHTML($excelContent);
    
    header('Content-Type: application/pdf');
    header('Content-Disposition: attachment;filename="Reporte_Equipos_'.$fecha.'.pdf"');
    header('Cache-Control: max-age=0');
    
    // Salida directa al navegador
    $mpdf->Output();

}

?>

</body>
</html>