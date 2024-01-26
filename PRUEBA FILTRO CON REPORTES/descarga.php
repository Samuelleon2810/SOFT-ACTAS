<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Descarga archivo</title>
</head>
<body>
    <h1>SU REPORTE SE ESTA DESCARGANDO</h1>
    <h2>si no empieza la descarga presiona el boton</h2>
    <form action="descarga.php" action="post">
        <input type="submit" value="Volver a descargar" class="botones">
    </form>

    <button onclick="location='./filtro.php'">Volver</button>
</body>
</html>


<?php
session_start();
require "/Users/Admin/Documents/GitHub/prueba-codigo-actas/vendor/autoload.php";

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

use Mpdf\Mpdf;


if($_SERVER['REQUEST_METHOD'] === 'POST'){



if(isset($_POST['excel'])){

 
    $spreadsheet = new Spreadsheet();

    $hoja = $spreadsheet->getActiveSheet();

    $indice = 1;
    foreach($_SESSION['matriz'] as &$valor){
    $hoja->fromArray($valor, null, 'A'.$indice.'');
    $indice++;
    }

    $writer = new Xlsx($spreadsheet);

    $fecha = date('d_m_y');

    $writer->save('Reporte_Equipos_'.$fecha.'.xlsx');

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

}
?>