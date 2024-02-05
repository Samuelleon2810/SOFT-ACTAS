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
<label for="departamento">Ingrese el departamento al que pertenece el responsable del equipo:</label>
<input type="text" placeholder="FINANCIERA, PRODUCCION" name="departamento" required>


<label for="cortex">¿El equipo cuenta con el cortex?:</label>
<select id="" name="cortex" required>
    <option value="ok">SI</option>
    <option value="no">NO</option>
</select>

    <label for="glpi">¿El equipo cuenta con el script de inventario?:</label>
<select id="" name="glpi" required>
    <option value="ok">SI</option>
    <option value="no">NO</option>
</select>


<input type="submit" value="enviar" name="enviar">
</form>    

<?php
require '/Users/Admin/Desktop/prueba codigo actas/vendor/autoload.php';


//extensiones para excel
use PhpOffice\PhpWord\SimpleType\Jc;
use PhpOffice\PhpWord\Shared\Converter;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\IOFactory;



if(isset($_POST['enviar'])){

    $tipoEquipo = $_POST['tipoEquipo'];


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
    $departamento = $_POST['departamento'];
    $cortex = $_POST['cortex'];
    $glpi = $_POST['glpi'];


    $phpWord = new \PhpOffice\PhpWord\PhpWord();


// Agregar una sección al documento
$section = $phpWord->addSection();

$imagePath = 'C:/Users/Admin/Desktop/prueba codigo actas/IMAGENES/logoElis.png';
$section->addImage(
    $imagePath,
    array(
        'width' => Converter::cmToPixel(3),
        'height' => Converter::cmToPixel(1.5),      
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
    "TECLADO Y MOUSE" => "SI",
];


foreach ($specifications as $label => $value) {
    $section->addText("    • $label: $value " , $normalFontStyleConNegrita);
}


$section->addText("\nAl momento de recibir el equipo aquí especificado se realizaron las pruebas de funcionamiento y se encuentra en buen estado físico. $estadoEquipo, es responsable del computador de su información y manejo de la misma."  , $normalFontStyle);
$section->addText("\nEl equipo cuenta con el siguiente software instalado: Windows 10 Pro, Office 365 Empresas, Navegador web Chrome, Adobe Pdf, Microsoft Teams." , $normalFontStyle);


$section->addText("\nDe acuerdo con lo anterior se hace constar que en el teclado y mouse se encuentran en buen estado y en las condiciones adecuadas para recibirlo sin ninguna salvedad. Después de entregado es responsabilidad de la persona brindar buen uso." , $normalFontStyle);
$section->addText("En caso de retiro de la compañía, se debe reintegrar en buen estado de funcionamiento." , $normalFontStyle);

$section->addText("\nSe deja en claro que el equipo no cuenta con nungun tipo de seguro contra robo perdida o cualquier daño, es total responsabilidad quien recibe y firma", $normalFontStyle , $justificar);

$section->addText("\nEntrega el equipo                                                                                 Recibe el equipo" , $normalFontStyleConNegrita);

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
        'height' => Converter::cmToPixel(1),      
        'marginTop' => Converter::cmToPixel(1), 
    )
);

$archivoWord = 'Acta_Entrega_Computador_Escritorio_' . $nombreUsuario . '.docx';

$objWriter = \PhpOffice\PhpWord\IOFactory::createWriter($phpWord, 'Word2007');

$objWriter->save($archivoWord);

// Redireccionar a la descarga del documento Word
header("Location: $archivoWord");

    unset($phpWord);
    unset($section);
    unset($textRun);
    unset($objWriter);

    $hojaCalculo->disconnectWorksheets();

    unset($spreadsheet);
    unset($hojaCalculo);
    unset($elemento);
    unset($hojita);
    unset($writer);

echo "SE DESCARGO SU WORD";

//exit();



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
    $departamento = $_POST['departamento'];
    $cortex = $_POST['cortex'];
    $glpi = $_POST['glpi'];


$phpWord = new \PhpOffice\PhpWord\PhpWord();
//$command = new Mikehaertl\ShellCommand\Command();



// Agregar una sección al documento
$section = $phpWord->addSection();

$imagePath = 'C:/Users/Admin/Desktop/prueba codigo actas/IMAGENES/logoElis.png';
$section->addImage(
    $imagePath,
    array(
        'width' => Converter::cmToPixel(3),
        'height' => Converter::cmToPixel(1.5),      
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


$section->addText("\nAl momento de recibir el equipo aquí especificado se realizaron las pruebas de funcionamiento y se encuentra en buen estado físico. \n Equipo $estadoEquipo, usted  es responsable del computador de su información y manejo de la misma." , $normalFontStyle);
$section->addText("\nEl equipo cuenta con el siguiente software instalado: Windows 10 Pro, Office 365 Empresas, Navegador web Chrome, Adobe Pdf, Microsoft Teams." , $normalFontStyle);

$section->addText("\nSe deja en claro que el equipo no cuenta con nungun tipo de seguro contra robo perdida o cualquier daño, es total responsabilidad quien recibe y firma", $normalFontStyle , $justificar);
$section->addText("En caso de retiro de la compañía, se debe reintegrar en buen estado de funcionamiento." , $normalFontStyle);


$section->addText("\nEntrega el equipo                                                                               Recibe el equipo" , $normalFontStyleConNegrita);

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
        'height' => Converter::cmToPixel(1),      
        'marginTop' => Converter::cmToPixel(1), 
    )
);

$archivoWord = 'Acta_Entrega_Computadores_' . $nombreUsuario . '.docx';
$objWriter = \PhpOffice\PhpWord\IOFactory::createWriter($phpWord, 'Word2007');
$objWriter->save($archivoWord);

// Redireccionar a la descarga del documento Word
header("Location: $archivoWord");

echo "SE DESCARGO SU WORD";

//exit();
}}

?>
<form action='actualizarInventario.php'>
<input type='hidden' name='nombre' value='<?php echo $nombreUsuario?>'>
<input type='hidden' name='cedula' value='<?php echo $cedulaUsuario?>'>
<input type='hidden' name='tipoEquipo' value='<?php echo $tipoEquipo?>'>
<input type='hidden' name='usoEquipo' value='<?php echo $estadoEquipo?>'>
<input type='hidden' name='nombreEquipo' value='<?php echo $nombreEquipo?>'>
<input type='hidden' name='procesadorEquipo' value='<?php echo $nombreProcesador?>'>
<input type='hidden' name='almacenamientoEquipo' value='<?php echo $almacenamientoEquipo?>'>
<input type='hidden' name='memoriaRAM' value='<?php echo $RAMEquipo?>'>
<input type='hidden' name='marcaEquipo' value='<?php echo $marcaEquipo?>'>
<input type='hidden' name='modeloEquipo' value='<?php echo $modeloEquipo?>'>
<input type='hidden' name='serialEquipo' value='<?php echo $serialEquipo?>'>
<input type='hidden' name='versionSO' value='<?php echo $versionSO?>'>
<input type='hidden' name='departamento' value='<?php echo $departamento?>'>
<input type='hidden' name='cortex' value='<?php echo $cortex?>'>
<input type='hidden' name='glpi' value='<?php echo $glpi?>'>
<input type="hidden" name='propiedadEquipo' value="uso">
<input type='submit' name='actualizarExcel' value="Subir al Inventario" class='botones'>
</form>

</body>
</html>
