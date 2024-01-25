<?php
require '../ACTAS COMPUTADORES/ActaEntregaComputadores.php';
require '/Users/Admin/Desktop/prueba codigo actas/vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Cell\Coordinate;

use Picqer\Barcode\BarcodeGeneratorHTML;
use BaconQrCode\Encoder\QrCode;
use BaconQrCode\Renderer\Image\Png;
use BaconQrCode\Writer;

require 'path/to/qrlib.php';

if ($_SERVER['REQUEST_METHOD'] === 'POST') {
    $tipoEquipo = $_POST['tipoEquipo'];

    if ($tipoEquipo === "escritorio") {
        // Resto de tu código para el tipo de equipo escritorio...
    } elseif ($tipoEquipo === 'portatil') {
        // Resto de tu código para el tipo de equipo portátil...

        // Obtener información del equipo desde Excel u otra fuente
        $serialEquipo = $_POST['serialEquipo'];
        $informacionEquipo = obtenerInformacionDesdeExcel($serialEquipo);

        // Generar el código QR
        $rutaCodigoQR = generarCodigoQR($informacionEquipo);

        // Resto de tu código para generar el Acta de Entrega y guardar la información en Excel...

        // Redirigir o realizar otras acciones con el código QR generado
        header("Location: $rutaCodigoQR");
        exit();
    }
}

function obtenerInformacionDesdeExcel($serialEquipo) {
    // Tu código para leer desde Excel usando PhpSpreadsheet...
    // Retorna un array con la información del equipo.
}

function generarCodigoQR($informacionEquipo) {
    // Directorio donde se almacenarán los códigos QR
    $directorioQR = '/path/to/qr_codes/';

    // Asegurarse de que el directorio exista
    if (!file_exists($directorioQR)) {
        mkdir($directorioQR, 0777, true);
    }

    // Nombre del archivo QR (puedes personalizarlo según tus necesidades)
    $nombreArchivoQR = 'codigoqr_' . uniqid() . '.png';

    // Ruta completa del archivo QR
    $rutaCodigoQR = $directorioQR . $nombreArchivoQR;

    // Generar el código QR
    QRcode::png(json_encode($informacionEquipo), $rutaCodigoQR);

    return $rutaCodigoQR;
}


// https://chat.openai.com/share/c425e6ac-6508-4b74-8257-4a9a4ef29b90