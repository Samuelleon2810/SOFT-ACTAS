<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=, initial-scale=1.0">
    <link rel="stylesheet" href="../index.css">
    <title>Document</title>
</head>
<body>
<form action="ActaEntregaCeculares.php" method="post">

<!-- datos persona -->
<label for="nombre">Ingrese el nombre del responsable</label>    
<input type="text" name="nombre" pattern="[A-Za-zÁÉÍÓÚáéíóúñÑ\s]+" title="Ingrese solo letras y espacios" required>
<label for="cedula">Ingrese el documento de a quien se entrega:</label>
<input type="text" name="cedula" pattern="\d+" title="Ingrese solo números" required>
<label for="Corporativo">Ingrese el número corporativo del celular</label>
<input type="number" placeholder="3015899630" name="Corporativo" required>
<label for="Asignado">Ingrese el nombre de la persona asignado</label>    
<input type="text" name="Asignado" pattern="[A-Za-zÁÉÍÓÚáéíóúñÑ\s]+" title="Ingrese solo letras y espacios" required>



<!-- caracteristicas equipo -->
<label for="usoEquipo">Ingrese la calidad del equipo:</label>
<select id="" name="usoEquipo" required>
    <option value="Nuevo">Nuevo</option>
    <option value="Usado">Usado</option>
</select>

<label for="serialEquipo">Ingrese el serial del equipo:</label>
<input type="text" placeholder="JN1TVV1" name="serialEquipo" required>
<label for="marcaEquipo">Ingrese la marca del equipo:</label>
<input type="text" placeholder=",motorola" name="marcaEquipo" required>
<label for="modeloEquipo">Ingresar el modelo del equipo:<label>
<input type="text" placeholder="g23" name="modeloEquipo" required>
<label for="Imei1">Ingresar el  primer IMEI del equipo:</label>
<input type="text" placeholder="15761248315" name="Imei1" required>
<label for="Imei2">Ingrese el segundo IMEI del equipo:</label>
<input type="text" placeholder="21587461223" name="Imei2" required>
<label for="Numero">Ingrese el nùmero identificador del celular</label>
<input type="number" placeholder="1" name="Numero" required>
<label for="Sim">Ingresar la SIM del equipo:</label>
<input type="text" placeholder="3104789564 " name="SIM" required>
<label for="Accesorios">Ingrese el correo del dispositivo:</label>
<input type="email" placeholder="example@gmail.com" name="EMAIL" required>
<label for="Contraseña">Ingresar la contraseña del equipo:</label>
<input type="password" placeholder="abc123 " name="Contraseña" required>
<label for="PIN">Ingresar el PIN del equipo:</label>
<input type="tel" placeholder="0000" name="PIN" required>
<input type="submit" value="enviar">
</form>
</body>
</html>

<?php
require '/Users/Admin/Desktop/prueba codigo actas/vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;   
use PhpOffice\PhpSpreadsheet\IOFactory;

if($_SERVER['REQUEST_METHOD'] === 'POST'){
    echo "SE HA ENVIADO EL FORMULARIO";

    //respuestas usuario

$nombreUsuario = $_POST['nombre'];
$nombreAsignado = $_POST['Asignado'];
$cedulaUsuario = $_POST['cedula'];
$Corporativo = $_POST [ 'Corporativo'];
$estadoEquipo = $_POST['usoEquipo'];
$imei1 = $_POST['Imei1'];
$imei2 = $_POST['Imei2'];
$marcaEquipo = $_POST['marcaEquipo'];
$modeloEquipo = $_POST['modeloEquipo'];
$serialEquipo = $_POST['serialEquipo'];
$contraseñaCorreo = $_POST['Contraseña'];
$numero = $_POST['Numero'];
$SIM = $_POST['SIM'];
$EMAIL = $_POST['EMAIL'];
$PIN = $_POST['PIN'];


    $spreadsheet = new Spreadsheet();
    $hojaCalculo = IOFactory::load('C:/Users/Admin/Downloads/CELULARES ELIS 2023.xlsx');

    $elemento = $hojaCalculo->getActiveSheet();

    $columnaNombrePersona ="H" ;
    $columnaCedulaPersona ="N" ;
    $columnaIMEI1 ="D" ;
    $columnaTipoDeEstado ="O" ;
    $columnaIMEI2 = "E" ;
    $columnaAsignado = "I" ;
    $columnaSIM = "G";
    $columnaMarcaEquipo = "B";
    $columnaModeloEquipo = "C";
    $columnaSerialEquipo = "A";
    $columnaNumero = "F";
    $columnaCorporativo = "J";
    $columnaEMAIl = "K";
    $columnaContraseña = "L";
    $columnaPIN = "M";
    

    $fila = 1;
        while (!empty($elemento->getCell($columnaSerialEquipo . $fila)->getValue())) {
    $fila++;
    }

    $elemento->setCellValue($columnaSerialEquipo . $fila, $serialEquipo);
    $elemento->setCellValue($columnaMarcaEquipo . $fila, $marcaEquipo);
    $elemento->setCellValue($columnaModeloEquipo . $fila, $modeloEquipo);
    $elemento->setCellValue($columnaIMEI1 . $fila, $imei1);
    $elemento->setCellValue($columnaIMEI2 . $fila, $imei2);
    $elemento->setCellValue($columnaNumero . $fila, $numero);
    $elemento->setCellValue($columnaSIM . $fila, $SIM);
    $elemento->setCellValue($columnaNombrePersona . $fila, $nombreUsuario);
    $elemento->setCellValue($columnaAsignado . $fila, $nombreAsignado);
    $elemento->setCellValue($columnaCorporativo . $fila, $Corporativo);
    $elemento->setCellValue($columnaEMAIl . $fila, $EMAIL);
    $elemento->setCellValue($columnaContraseña . $fila, $contraseñaCorreo);
    $elemento->setCellValue($columnaPIN . $fila, $PIN);
    $elemento->setCellValue($columnaCedulaPersona . $fila, $cedulaUsuario);
    $elemento->setCellValue($columnaTipoDeEstado . $fila, $estadoEquipo);

    $writer = IOFactory::createWriter($hojaCalculo, 'Xlsx');
    $writer->save('C:/Users/Admin/Downloads/CELULARES ELIS 2023.xlsx');
}
?>