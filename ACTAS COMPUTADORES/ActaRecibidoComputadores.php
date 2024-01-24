<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Acta Recibido Computadores</title>
</head>
<body>
    <form action="ActaRecibidoComputadores.php" method="post">
    <h1>ACTA DE RECIBIDO EQUIPO COMPUTADOR</h1>
    <label for="nombre">Ingrese el nombre de quien entrega el equipo </label>    
<input type="text" name="nombre" pattern="[A-Za-zÁÉÍÓÚáéíóúñÑ\s]+" title="Ingrese solo letras y espacios" required>
<label for="cedula">Ingrese el documento de quien lo entrega:</label>
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
<label for="entrega">Ingrese lo que se entrega con el equipo:</label>
<input type="text" placeholder="CARGADOR ORIGINAL" name="entrega" required>
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
    $periferico = $_POST['periferico'];
    $entrega = $_POST['entrega'];

    if($tipoEquipo === "escritorio"){

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
        $section->addText("ACTA DE RECIBIDO DE EQUIPO DE ESCRITORIO", $titleFontStyle, $paragraphStyle);
        
        
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
        
        $fecha = "En la ciudad de Bogotá, a los " . date('d') . " días del mes de " . $meses[date('m') - 1] . " del año 20" . date('y') . ", se hace entrega de un equipo portatil, al señor Julian Andres Ariza , Soporte Tecnico en sistemas IT Elis Colombia por parte de ";
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
            "TECLADO Y MOUSE" => $periferico
        ];
        
        
        foreach ($specifications as $label => $value) {
            $section->addText("    • $label: $value " , $normalFontStyleConNegrita);
        }
        
        
        $section->addText("\nAl momento de recibir el equipo aquí especificado se realizaron las pruebas de funcionamiento y se encuentra en buen estado de funcionamiento.  " , $normalFontStyle , $justificar);
        $section->addText("\nDe acuerdo con lo anterior se hace constar que en el equipo se encuentra en las condiciones adecuadas para recibirlo sin ningunas salvedades." , $normalFontStyle , $justificar);
        
        
        $section->addText("\nDe acuerdo con lo anterior se hace constar que en el teclado y mouse se encuentran ESTADOPERIFERICOS y en las condiciones adecuadas para recibirlo sin ninguna salvedad. Después de entregado es responsabilidad de la persona brindar buen uso." , $normalFontStyle , $justificar);
        
        
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
        
        $archivoWord = 'Acta_Recibido_Computador_Escritorio_' . $nombreUsuario . '.docx';
        $objWriter = \PhpOffice\PhpWord\IOFactory::createWriter($phpWord, 'Word2007');
        $objWriter->save($archivoWord);
        
        // Redireccionar a la descarga del documento Word
        header("Location: $archivoWord");
        
        echo "SE DESCARGO SU WORD";
        
        $archivoWord = 'C:Users/Admin/Downloads/Acta_Recibido_Computador_Escritorio_' . $nombreUsuario . '.docx';
        
        $archivoPdf = 'C:Users/Admin/Downloads/Acta_Recibido_Computador_Escritorio_' . $nombreUsuario . '.pdf';
        
        $phpWord = IOFactory::load($archivoWord);
        
        $archivoHtml = 'Acta_Recibido_Computadores_' . $nombreUsuario . '.html';
        $objWriter = IOFactory::createWriter($phpWord, 'HTML');
        $objWriter->save($archivoHtml);
        
        $command = new Command("wkhtmltopdf $archivoHtml $archivoPdf");
        $command->execute();
        
        header("Location: $archivoPdf");
        
        echo "SE DESCARGO SU PDF";
        
        exit();

    }else{

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

        $section->addText("ACTA DE RECIBIDO EQUIPO PORTATIL", $titleFontStyle, $paragraphStyle);
        
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
        
        $fecha = "En la ciudad de Bogotá, a los " . date('d') . " días del mes de " . $meses[date('m') - 1] . " del año 20" . date('y') . ", se hace entrega de un equipo portatil, al señor Julian Andres Ariza , Soporte Tecnico en sistemas IT Elis Colombia por parte de ";
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
            "ENTREGA" => $entrega
        ];
        
        
        foreach ($specifications as $label => $value) {
            $section->addText("    • $label: $value " , $normalFontStyleConNegrita);
        }
        
        
        $section->addText("\nAl momento de recibir el equipo aquí especificado se realizaron las pruebas de funcionamiento y se encuentra en buen estado de funcionamiento.  " , $normalFontStyle , $justificar);
        $section->addText("\nDe acuerdo con lo anterior se hace constar que en el equipo se encuentra en las condiciones adecuadas para recibirlo sin ningunas salvedades." , $normalFontStyle , $justificar);
        
        
        $section->addText("\nDe acuerdo con lo anterior se hace constar que en el teclado y mouse se encuentran en buen estado y en las condiciones adecuadas para recibirlo sin ninguna salvedad." , $normalFontStyle , $justificar);
        
        
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
        
        $archivoWord = 'Acta_Entrega_Computador_Portatil_' . $nombreUsuario . '.docx';
        $objWriter = \PhpOffice\PhpWord\IOFactory::createWriter($phpWord, 'Word2007');
        $objWriter->save($archivoWord);
        
        // Redireccionar a la descarga del documento Word
        header("Location: $archivoWord");
        
        echo "SE DESCARGO SU WORD";
        
        $archivoWord = 'C:Users/Admin/Downloads/Acta_Recibido_Computador_Portatil_' . $nombreUsuario . '.docx';
        
        // Ruta del archivo PDF de salida
        $archivoPdf = 'C:Users/Admin/Downloads/Acta_Recibido_Computador_Portatil_' . $nombreUsuario . '.pdf';
        
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

    }
}
?>