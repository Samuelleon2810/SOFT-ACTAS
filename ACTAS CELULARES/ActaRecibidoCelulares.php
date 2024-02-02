<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=, initial-scale=1.0">
    <link rel="stylesheet" href="../index.css">
    <link rel="shortcut icon" href="/IMAGENES/logoElis.png" type="image/x-icon">
    <title>Acta recibido celulares</title>
</head>
<body>
<form action="ActaRecibidoCelulares.php" method="post">
<h1>ACTA DE RECIBIDO EQUIPO CELULAR</h1>
<!-- datos persona -->
<label for="nombre">Ingrese el nombre del que entrega el telefono:</label>    
<input type="text" name="nombre" pattern="[A-Za-zÁÉÍÓÚáéíóúñÑ\s]+" title="Ingrese solo letras y espacios" required>
<label for="color">Ingrese el color del telefono:</label>    
<input type="text" name="color" pattern="[A-Za-zÁÉÍÓÚáéíóúñÑ\s]+" title="Ingrese solo letras y espacios" required>
<label for="cedula">Ingrese el documento de quien entrega el telefono:</label>
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
<label for="PIN">Ingresar el departamento de a quien se le entrga el equipo:</label>
<input type="tel" placeholder="CONTABILIDAD" name="departamento" required>
<input type="submit" value="enviar">
</form>


<?php
require '/Users/Admin/Desktop/prueba codigo actas/vendor/autoload.php';


use PhpOffice\PhpSpreadsheet\Writer\Xlsx;   
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


    $phpWord = new \PhpOffice\PhpWord\PhpWord();

    $section = $phpWord->addSection();
    
    $imagePath = 'C:/Users/Admin/Desktop/prueba codigo actas/IMAGENES/logoElis.png';
    $section->addImage(
        $imagePath,
        array(
            'width' => Converter::cmToPixel(3),
            'height' => Converter::cmToPixel(2),      
            'marginTop' => Converter::cmToPixel(1), 
        )
    );
    
    
    $titleFontStyle = array('name' => 'Calibri', 'size' => 16, 'color' => '000000');
    

    $paragraphStyle = array('alignment' => Jc::CENTER);
    $justificar = array('algnment' => Jc::BOTH);

    $section->addText("ACTA DE RECIBIDO CELULAR", $titleFontStyle, $paragraphStyle);
    
    $normalFontStyle = array('name' => 'Century Gothic', 'size' => 10,5, 'color' => '1B2232' , 'alignment' => Jc::BOTH);
    $normalFontStyleConNegrita = array('name' => 'Century Gothic', 'size' => 10,5, 'color' => '1B2232' , 'alignment' => Jc::BOTH , 'bold' => true);
    
    // Texto de la primera parte del acta


    $meses = [
        'Enero',
        'Febrero',
        'Marzo',
        'Abril',
        'Mayo',
        'Junio',
        'Julio',
        'Agosto',
        'Septiembre',
        'Octubre',
        'Noviembre',
        'Diciembre'
    ];
    
    $fecha = "En la ciudad de Bogotá, a los " . date('d') . " días del mes de " . $meses[date('m') - 1] . " del año 20" . date('y') . ", se hace entrega de un CELULAR, al señor Julian Andres Ariza , Soporte Tecnico en sistemas IT Elis Colombia por parte de ";
    $textRun = $section->addTextRun($normalFontStyle);
    $textRun->addText($fecha , $normalFontStyle);
    
    $textRun->addText($nombreUsuario, $normalFontStyleConNegrita);
    
    $identificacion = " identificado con cédula de ciudadanía número ";
    $textRun->addText($identificacion, $normalFontStyle + array('spaceAfter' => 0));
    
    $textRun->addText($cedulaUsuario, $normalFontStyleConNegrita);
    
    $especificacion = " con las siguientes especificaciones:";
    $section->addText($especificacion, $normalFontStyle);
    
    
    // Lista de especificaciones
    $specifications = [
        "MARCA" => $marcaEquipo,
        "MODELO" => $modeloEquipo,
        "COLOR" => $colorEquipo,
        "NUMERO DE SERIE" => $nombreProcesador,
        "IMEI 1" => $imei1,
        "IMEI 2" => $imei2,
        "SIMCARD" => $SIM
    ];
    
    
    foreach ($specifications as $label => $value) {
        $section->addText("    • $label: $value " , $normalFontStyleConNegrita);
    }
    
    
    $section->addText("\nAl momento de recibir el equipo aquí especificado se realizaron las pruebas de funcionamiento y se encuentra en buen estado de funcionamiento.   " , $normalFontStyle , $justificar);
    $section->addText("\nDe acuerdo con lo anterior se hace constar que en el equipo se encuentra en las condiciones adecuadas para recibirlo sin ningunas salvedades." , $normalFontStyle , $justificar);
    
    $section->addText("\nRecibe el equipo                                                                                  Entrega" , $normalFontStyleConNegrita);
    
    $imagePathJul = 'C:/Users/Admin/Desktop/prueba codigo actas/IMAGENES/jul.png';
    $section->addImage(
        $imagePathJul,
        array(
            'width' => Converter::cmToPixel(3),
            'height' => Converter::cmToPixel(1.5),      
            'marginTop' => Converter::cmToPixel(1), 
        )
    );
    
    $section->addText("\nJulian Andres Ariza Pardo                                                             ".$nombreUsuario."" , $normalFontStyleConNegrita);
    $section ->addText("\nSoporte Tecnico de sistemas IT" , $normalFontStyleConNegrita);
    
    $imagePathInfo = 'C:/Users/Admin/Desktop/prueba codigo actas/IMAGENES/infoElis.png';
    $section->addImage(
        $imagePathInfo,
        array(
            'width' => Converter::cmToPixel(12),
            'height' => Converter::cmToPixel(1.5),      
            'marginTop' => Converter::cmToPixel(1), 
        )
    );
    
    $archivoWord = 'Acta_Recibido_Celular_' . $nombreUsuario . ' '. date('d/m/y') .' .docx';
    $objWriter = \PhpOffice\PhpWord\IOFactory::createWriter($phpWord, 'Word2007');
    $objWriter->save($archivoWord);
    
    // Redireccionar a la descarga del documento Word
    header("Location: $archivoWord");
    
    echo "SE DESCARGO SU WORD";
    
    $archivoWord = 'C:Users/Admin/Downloads/Acta_Recibido_Celular_' . $nombreUsuario . '.docx';
    
    // Ruta del archivo PDF de salida
    $archivoPdf = 'C:Users/Admin/Downloads/Acta_Recibido_Celular_' . $nombreUsuario . '.pdf';
    
    // Cargar el documento Word
    $phpWord = IOFactory::load($archivoWord);
    
    // Guardar el documento Word en HTML temporal
    $archivoHtml = 'Acta_Recibido_Celular_' . $nombreUsuario . '.html';
    $objWriter = IOFactory::createWriter($phpWord, 'HTML');
    $objWriter->save($archivoHtml);
    
    // Convertir el archivo HTML a PDF
    $command = new Command("wkhtmltopdf $archivoHtml $archivoPdf");
    $command->execute();
    
    // Redireccionar o hacer algo con el PDF generado
    header("Location: $archivoPdf");
    
    echo "SE DESCARGO SU PDF";
    
    exit();
}
?>

<form action='actualizarInventario.php'>
<input type='hidden' name='nombre' value='<?php echo $nombreUsuario?>'>
<input type='hidden' name='cedula' value='<?php echo $cedulaUsuario?>'>
<input type='hidden' name='Asignado' value='<?php echo $nombreAsignado?>'>
<input type='hidden' name='usoEquipo' value='<?php echo $estadoEquipo?>'>
<input type='hidden' name='corporativo' value='<?php echo $Corporativo?>'>
<input type='hidden' name='Imei1' value='<?php echo $imei1?>'>
<input type='hidden' name='Imei2' value='<?php echo $imei2?>'>
<input type='hidden' name='contraseña' value='<?php echo $contraseñaCorreo?>'>
<input type='hidden' name='marcaEquipo' value='<?php echo $marcaEquipo?>'>
<input type='hidden' name='modeloEquipo' value='<?php echo $modeloEquipo?>'>
<input type='hidden' name='serialEquipo' value='<?php echo $serialEquipo?>'>
<input type='hidden' name='Numero' value='<?php echo $numero?>'>
<input type='hidden' name='departamento' value='<?php echo $departamento?>'>
<input type='hidden' name='SIM' value='<?php echo $SIM?>'>
<input type='hidden' name='EMAIL' value='<?php echo $EMAIL?>'>
<input type='hidden' name='PIN' value='<?php echo $PIN?>'>
<input type='hidden' name='color' value='<?php echo $colorEquipo?>'>
<input type='hidden' name='claseEquipo' value='celular'>
<input type='submit' name='actualizarExcel' value="Subir al Inventario" class='botones'>
</form>

</body>
</html>