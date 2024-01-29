<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <link rel="stylesheet" href="/index.css">
    <link rel="shortcut icon" href="/IMAGENES/logoElis.png" type="image/x-icon">
    <title>Acta Entrega Computadores</title>
</head> 
<body>
<form action="ActaEntregaComputadores.php" method="post">
<h1>ACTA DE ENTREGA EQUIPO COMPUTADOR </h1>
<!-- datos persona -->
<label for="nombre">Ingrese el nombre del destinatario</label>    
<input type="text" name="nombre" pattern="[A-Za-zÁÉÍÓÚáéíóúñÑ\s]+" title="Ingrese solo letras y espacios" required>
<label for="cedula">Ingrese el documento de a quien se entrega:</label>
<input type="text" name="cedula" pattern="\d+" title="Ingrese solo números" required>

<!-- caracteristicas equipo -->
<label for="tipoEquipo">Ingrese el tipo de equipo:</label>
<select id="" name="tipoEquipo" required>
    <option value="portatil">Portatil</option>
    <option value="escritorio">Escritorio</option>
</select>

<label for="usoEquipo">Ingrese la calidad del equipo:</label>
<select id="" name="usoEquipo" required>
    <option value="Nuevo">Nuevo</option>
    <option value="Usado">Usado</option>
</select>
<Label for=" nombreEquipo">Ingresar el nombre del equipo:</Label>
<input type="text" placeholder="DESKTOP-76ILRV7" name="nombreEquipo" required>
<label for="procesardorEquipo">Ingresar el procesadoir del equipo:</label>
<input type="text" placeholder="INTEL CORE I7-8565U CPU " name="procesadorEquipo" required>
<label for="almacenamientoEquipo">Ingresar el almacenamiento del equipo:</label>
<input type="text" placeholder="GB" name="almacenamientoEquipo" required>
<label for="memoriaRAM">Ingrese la RAM del equipo:</label>
<input type="text" placeholder="GB RAM" name="memoriaRAM" required>
<label for="marcaEquipo">Ingrese la marca del equipo:</label>
<input type="text" placeholder="Lenovo" name="marcaEquipo" required>
<label for="modeloEqauipo">Ingresar el modelo del equipo:<label>
<input type="text" placeholder="RS-127644" name="modeloEquipo" required>
<label for="serialEquipo">Ingrese el serial del equipo:</label>
<input type="text" placeholder="JN1TVV1" name="serialEquipo" required>
<label for="versionSO">Ingrese la versiòn del sistema operativo:</label>
<input type="text" placeholder="10 PRO" name="versionSO" required>
<input type="submit" value="enviar" name="enviar">
</form>    
</body>
</html>

<?php
require '/Users/Admin/Desktop/prueba codigo actas/vendor/autoload.php';

//extensiones para excel
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

    $tipoEquipo = $_POST['tipoEquipo'];


    if($tipoEquipo === "escritorio"){

  //      echo "<label for='Perifericos'>Ingrese las especificaciones de los perifericos:</label>\n";
//echo "<input type='text' placeholder='Raton,Teclado....' name='periferico' required>\n";
//echo "<input type='submit' value='Descargar y Llenar' name='enviarInfo'>\n";

//if($_POST['enviarInfo']){
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
   // $periferico = $_POST['periferico']; 
//}else{
  //      echo "por favor llene los campos para generar su documento";
   // }

    //escritura en excel

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

    $spreadsheet = new Spreadsheet();
    $hojaCalculo = IOFactory::load('C:/Users/Admin/Downloads/01_INVENTARIO PLANTA NORTE 2023.xlsx');

    $elemento = $hojaCalculo->getActiveSheet();

    $hojita = $hojaCalculo->getSheet(1);

    $cellIterator = $elemento->getRowIterator();

    foreach ($hojita->getRowIterator() as $row) {
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

    $phpWord = new \PhpOffice\PhpWord\PhpWord();


// Agregar una sección al documento
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

// Configura el estilo de párrafo para centrar
$paragraphStyle = array('alignment' => Jc::CENTER);
$justificar = array('algnment' => Jc::BOTH);

// Agrega el título usando addText con el estilo de párrafo
$section->addText("ACTA DE ENTREGA DE EQUIPO DE ESCRITORIO", $titleFontStyle, $paragraphStyle);


$fechaActual = date('a los d/m/Y');
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

$fecha = "En la ciudad de Bogotá, a los " . date('d') . " días del mes de " . $meses[date('m') - 1] . " del año 20" . date('y') . ", se hace entrega de un equipo de escritorio, a ";
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
    "SERIAL" => $serialEquipo,
    "PROCESADOR" => $nombreProcesador,
    "DISCO DURO" => $almacenamientoEquipo,
    "MEMORIA RAM" => $RAMEquipo,
    "NOMBRE DEL EQUIPO" => $nombreEquipo,
    "TECLADO Y MOUSE" => $periferico, // Debes llenar este valor según tus necesidades
];


foreach ($specifications as $label => $value) {
    $section->addText("    • $label: $value " , $normalFontStyleConNegrita);
}


$section->addText("\nAl momento de recibir el equipo aquí especificado se realizaron las pruebas de funcionamiento y se encuentra en buen estado físico. $estadoEquipo, es responsable del computador de su información y manejo de la misma." , $normalFontStyle , $justificar);
$section->addText("\nEl equipo cuenta con el siguiente software instalado: Windows 10 Pro, Office 365 Empresas, Navegador web Chrome, Adobe Pdf, Microsoft Teams." , $normalFontStyle , $justificar);


$section->addText("\nDe acuerdo con lo anterior se hace constar que en el teclado y mouse se encuentran ESTADOPERIFERICOS y en las condiciones adecuadas para recibirlo sin ninguna salvedad. Después de entregado es responsabilidad de la persona brindar buen uso." , $normalFontStyle , $justificar);
$section->addText("En caso de retiro de la compañía, se debe reintegrar en buen estado de funcionamiento." , $normalFontStyle , $justificar);


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

$archivoWord = 'Acta_Entrega_Computador_Escritorio_' . $nombreUsuario . '.docx';
$objWriter = \PhpOffice\PhpWord\IOFactory::createWriter($phpWord, 'Word2007');
$objWriter->save($archivoWord);

// Redireccionar a la descarga del documento Word
header("Location: $archivoWord");

echo "SE DESCARGO SU WORD";

$archivoWord = 'C:Users/Admin/Downloads/Acta_Entrega_Computador_Escritorio_' . $nombreUsuario . '.docx';

// Ruta del archivo PDF de salida
$archivoPdf = 'C:Users/Admin/Downloads/Acta_Entrega_Computador_Escritorio_' . $nombreUsuario . '.pdf';

// Cargar el documento Word
$phpWord = IOFactory::load($archivoWord);

// Guardar el documento Word en HTML temporal
$archivoHtml = 'Acta_Entrega_Computadores_' . $nombreUsuario . '.html';
$objWriter = IOFactory::createWriter($phpWord, 'HTML');
$objWriter->save($archivoHtml);

// Convertir el archivo HTML a PDF
$command = new Command("wkhtmltopdf $archivoHtml $archivoPdf");
$command->execute();

// Redireccionar o hacer algo con el PDF generado
header("Location: $archivoPdf");

echo "SE DESCARGO SU PDF";

exit();



}elseif($tipoEquipo=== 'portatil'){



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
    //escritura en excel

    $columnaNombrePersona ="O" ;
    $columnaCedulaPersona ="Z" ;
    $columnaTipoDeEquipo ="AA" ;
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

    $spreadsheet = new Spreadsheet();
    $hojaCalculo = IOFactory::load('C:/Users/Admin/Downloads/01_INVENTARIO PLANTA NORTE 2023.xlsx');

    $elemento = $hojaCalculo->getActiveSheet();

    $hojita = $hojaCalculo->getSheet(0);

    $cellIterator = $elemento->getRowIterator();

    foreach ($hojita->getRowIterator() as $row) {
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


    //escritura en word
    
        //creamos un arreglo asociativo con los datos anteriores


$phpWord = new \PhpOffice\PhpWord\PhpWord();
//$command = new Mikehaertl\ShellCommand\Command();



// Agregar una sección al documento
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

// Configura el estilo de párrafo para centrar
$paragraphStyle = array('alignment' => Jc::CENTER);
$justificar = array('algnment' => Jc::BOTH);

// Agrega el título usando addText con el estilo de párrafo
$section->addText("ACTA DE ENTREGA DE EQUIPO PORTATIL", $titleFontStyle, $paragraphStyle);


$fechaActual = date('a los d/m/Y');
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

$fecha = "En la ciudad de Bogotá, a los " . date('d') . " días del mes de " . $meses[date('m') - 1] . " del año 20" . date('y') . ", se hace entrega de un equipo portatil, a ";
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
    "SERIAL" => $serialEquipo,
    "PROCESADOR" => $nombreProcesador,
    "DISCO DURO" => $almacenamientoEquipo,
    "MEMORIA RAM" => $RAMEquipo,
    "NOMBRE DEL EQUIPO" => $nombreEquipo,
];


foreach ($specifications as $label => $value) {
    $section->addText("    • $label: $value " , $normalFontStyleConNegrita);
}


$section->addText("\nAl momento de recibir el equipo aquí especificado se realizaron las pruebas de funcionamiento y se encuentra en buen estado físico. \n Equipo $estadoEquipo, usted  es responsable del computador de su información y manejo de la misma." , $normalFontStyle , $justificar);
$section->addText("\nEl equipo cuenta con el siguiente software instalado: Windows 10 Pro, Office 365 Empresas, Navegador web Chrome, Adobe Pdf, Microsoft Teams." , $normalFontStyle , $justificar);

$section->addText("En caso de retiro de la compañía, se debe reintegrar en buen estado de funcionamiento." , $normalFontStyle , $justificar);


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

$archivoWord = 'Acta_Entrega_Computadores_' . $nombreUsuario . '.docx';
$objWriter = \PhpOffice\PhpWord\IOFactory::createWriter($phpWord, 'Word2007');
$objWriter->save($archivoWord);

// Redireccionar a la descarga del documento Word
header("Location: $archivoWord");

echo "SE DESCARGO SU WORD";

$archivoWord = 'C:Users/Admin/Downloads/Acta_Entrega_Computador_Portatil' . $nombreUsuario . '.docx';

// Ruta del archivo PDF de salida
$archivoPdf = 'C:Users/Admin/Downloads/Acta_Entrega_Computador_Portatil' . $nombreUsuario . '.pdf';

// Cargar el documento Word
$phpWord = IOFactory::load($archivoWord);

// Guardar el documento Word en HTML temporal
$archivoHtml = 'Acta_Entrega_Computadores_' . $nombreUsuario . '.html';
$objWriter = IOFactory::createWriter($phpWord, 'HTML');
$objWriter->save($archivoHtml);

// Convertir el archivo HTML a PDF
$command = new Command("wkhtmltopdf $archivoHtml $archivoPdf");
$command->execute();

// Redireccionar o hacer algo con el PDF generado
header("Location: $archivoPdf");

echo "SE DESCARGO SU PDF";


exit();
}}

?>