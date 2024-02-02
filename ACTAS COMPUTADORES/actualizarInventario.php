<?php

use PhpOffice\PhpWord\SimpleType\Jc;
use PhpOffice\PhpWord\Shared\Converter;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\IOFactory;

require '/Users/Admin/Documents/GitHub/prueba-codigo-actas/vendor/autoload.php';
include_once './ActaEntregaComputadores.php';

if($_SERVER['REQUEST_METHOD']=== 'POST'){
    if(isset($_POST['actualizarExcel'])){

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
        

            require "./CONEXION BD/conexion.php";

            try{

                $consultaEquipo = "INSERT INTO laptops (`nombre_equipo`, `procesador` , `almacenamiento` , `memoria_ram` , `marca` , `modelo`, `serial`, `sistema_operativo` , `departamento` , `cortex` , `propiedad` , `glpi` , `estado`)
    VALUES('$nombreEquipo' , '$nombreProcesador' , '$almacenamientoEquipo' , '$RAM' , '$marcaEquipo' , '$modeloEquipo','$serialEquipo' , '$versionSO' , '$departamento' , '$cortex', '$propiedad', '$glpi', '$estado')";

            mysqli_query($conexion2 , $consultaEquipo);

            $confirmacionUsuario = "SELECT * FROM `ussers` WHERE 'usuario' = $nombreUsuario";

            $resultado = mysqli_query($conexion2 , $confirmacionUsuario);

            if($resultado){
                $consultaPersonas = "INSERT INTO ussers(`usuario` , `documento` , `departamento`)WHERE 'usuario' = $nombreUsuario VALUES( '$nombreUsuario' , '$cedulaUsuario' , '$departamento')";
            }else{
                $consultaPersonas = "INSERT INTO ussers(`usuario` , `documento` , `departamento`) VALUES( '$nombreUsuario' , '$cedulaUsuario' , '$departamento')";
            }


            mysqli_query($conexion2 , $consultaPersonas);

            if($conexion2 -> affected_rows > 1){
echo "Se ha actualizado el inventario";
$conexion2-> close();
            }else{
echo "Algo malo ha pasado, intente de nuevo...";
            }
$conexion2-> close();
        } catch (PDOException $e) {
            echo "Error de conexión: " . $e->getMessage();
        }

    // $columnaNombrePersona ="O" ;
    // $columnaCedulaPersona ="Z" ;
    // $columnaTipoDeEquipo ="Q" ;
    // $columnaTipoDeEstado ="AB" ;
    // $columnaNombreEquipo = "A" ;
    // $columnaNombreProcesadorEquipo = "B" ;
    // $columnaAlmacenamientoEquipo = "C" ;
    // $columnaRAMEquipo = "D";
    // $columnaMarcaEquipo = "E";
    // $columnaModeloEquipo = "F";
    // $columnaSerialEquipo = "G";
    // $columnaVersionSO = "H";
    // $columnaAsignado = "AC";

    // $hojaCalculo = IOFactory::load('C:/Users/Admin/Downloads/01_INVENTARIO PLANTA NORTE 2023.xlsx');

    // $elemento = $hojaCalculo->getActiveSheet();

    // $cellIterator = $elemento->getRowIterator();

    // foreach ($elemento->getRowIterator() as $row) {
    //     foreach ($row->getCellIterator() as $cell) {
    //         $cellValue = $cell->getValue();
        
    //         if ($cellValue == $serialEquipo) {
    //             $foundCell = $cell->getCoordinate();
    //             list($columna, $fila) = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::coordinateFromString($foundCell);
    //             break 2;  // Salir de ambos bucles
    //         }

    //     }
    // }

    // if (!isset($foundCell)) {
    //     $fila = 1;
    //     while (!empty($elemento->getCell($columnaNombreEquipo . $fila)->getValue())) {
    //     $fila++;
    // }
    // }
    

    // $elemento->setCellValue($columnaNombrePersona . $fila, $nombreUsuario);
    // $elemento->setCellValue($columnaCedulaPersona . $fila, $cedulaUsuario);
    // $elemento->setCellValue($columnaNombreEquipo . $fila, $nombreEquipo);
    // $elemento->setCellValue($columnaNombreProcesadorEquipo . $fila, $nombreProcesador);
    // $elemento->setCellValue($columnaAlmacenamientoEquipo . $fila, $almacenamientoEquipo);
    // $elemento->setCellValue($columnaRAMEquipo . $fila, $RAMEquipo);
    // $elemento->setCellValue($columnaMarcaEquipo . $fila, $marcaEquipo);
    // $elemento->setCellValue($columnaSerialEquipo . $fila, $serialEquipo);
    // $elemento->setCellValue($columnaModeloEquipo . $fila, $modeloEquipo);
    // $elemento->setCellValue($columnaVersionSO . $fila, $versionSO);
    // $elemento->setCellValue($columnaTipoDeEquipo . $fila, $tipoEquipo);
    // $elemento->setCellValue($columnaTipoDeEstado . $fila, $estadoEquipo);
    // $elemento->setCellValue($columnaAsignado . $fila, "EN USO");

    // $writer = IOFactory::createWriter($hojaCalculo, 'Xlsx');
    // $writer->save('C:/Users/Admin/Downloads/01_INVENTARIO PLANTA NORTE 2023.xlsx');

        }elseif($tipoEquipo === "portatil"){


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

    require "./CONEXION BD/conexion.php";

    try{

    $consultaEquipo = "INSERT INTO laptops (`nombre_equipo`, `procesador` , `almacenamiento` , `memoria_ram` , `marca` , `modelo`, `serial`, `sistema_operativo` , `departamento` , `cortex` , `propiedad` , `glpi` , `estado`)
    VALUES('$nombreEquipo' , '$nombreProcesador' , '$almacenamientoEquipo' , '$RAM' , '$marcaEquipo' , '$modeloEquipo','$serialEquipo' , '$versionSO' , '$departamento' , '$cortex', '$propiedad', '$glpi', '')";

    mysqli_query($conexion2 , $consultaEquipo);
    
    $confirmacionUsuario = "SELECT * FROM `ussers` WHERE 'usuario' = $nombreUsuario";

    $resultado = mysqli_query($conexion2 , $confirmacionUsuario);

    if($resultado){
        $consultaPersonas = "INSERT INTO ussers(`usuario` , `documento` , `departamento`)WHERE 'usuario' = $nombreUsuario VALUES( '$nombreUsuario' , '$cedulaUsuario' , '$departamento')";
    }else{
        $consultaPersonas = "INSERT INTO ussers(`usuario` , `documento` , `departamento`) VALUES( '$nombreUsuario' , '$cedulaUsuario' , '$departamento')";
    }
    mysqli_query($conexion2 , $consultaPersonas);

    if($conexion2 -> affected_rows > 1){
echo "Se ha actualizado el inventario";
$conexion2-> close();
    }else{
echo "Algo malo ha pasado, intente de nuevo...";
    }
$conexion2-> close();
} catch (PDOException $e) {
    echo "Error de conexión: " . $e->getMessage();
}

    // $columnaCedulaPersona ="Z" ;
    // $columnaTipoDeEquipo ="Q" ;
    // $columnaTipoDeEstado ="AB" ;
    // $columnaNombreEquipo = "A" ;
    // $columnaNombreProcesadorEquipo = "B" ;
    // $columnaAlmacenamientoEquipo = "C" ;
    // $columnaRAMEquipo = "D";
    // $columnaMarcaEquipo = "E";
    // $columnaModeloEquipo = "F";
    // $columnaSerialEquipo = "G";
    // $columnaVersionSO = "H";
    // $columnaAsignado = "AC";

    // $hojaCalculo = IOFactory::load('C:/Users/Admin/Downloads/01_INVENTARIO PLANTA NORTE 2023.xlsx');

    // $elemento = $hojaCalculo->getActiveSheet();

    // $cellIterator = $elemento->getRowIterator();

    // foreach ($elemento->getRowIterator() as $row) {
    //     foreach ($row->getCellIterator() as $cell) {
    //         $cellValue = $cell->getValue();
        
    //         if ($cellValue == $serialEquipo) {
    //             $foundCell = $cell->getCoordinate();
    //             list($columna, $fila) = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::coordinateFromString($foundCell);
    //             break 2;  // Salir de ambos bucles
    //         }

    //     }
    // }

    // if (!isset($foundCell)) {
    //     $fila = 1;
    //     while (!empty($elemento->getCell($columnaNombreEquipo . $fila)->getValue())) {
    //     $fila++;
    // }
    // }
    

    // $elemento->setCellValue($columnaNombrePersona . $fila, $nombreUsuario);
    // $elemento->setCellValue($columnaCedulaPersona . $fila, $cedulaUsuario);
    // $elemento->setCellValue($columnaNombreEquipo . $fila, $nombreEquipo);
    // $elemento->setCellValue($columnaNombreProcesadorEquipo . $fila, $nombreProcesador);
    // $elemento->setCellValue($columnaAlmacenamientoEquipo . $fila, $almacenamientoEquipo);
    // $elemento->setCellValue($columnaRAMEquipo . $fila, $RAMEquipo);
    // $elemento->setCellValue($columnaMarcaEquipo . $fila, $marcaEquipo);
    // $elemento->setCellValue($columnaSerialEquipo . $fila, $serialEquipo);
    // $elemento->setCellValue($columnaModeloEquipo . $fila, $modeloEquipo);
    // $elemento->setCellValue($columnaVersionSO . $fila, $versionSO);
    // $elemento->setCellValue($columnaTipoDeEquipo . $fila, $tipoEquipo);
    // $elemento->setCellValue($columnaTipoDeEstado . $fila, $estadoEquipo);
    // $elemento->setCellValue($columnaAsignado . $fila, "EN USO");

    // $writer = IOFactory::createWriter($hojaCalculo, 'Xlsx');
    // $writer->save('C:/Users/Admin/Downloads/01_INVENTARIO PLANTA NORTE 2023.xlsx');


        } elseif($tipoEquipo === "celular"){
    
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
            $colorEquipo = $_POST['color'];
            $departamento = $_POST['departamento'];

            require "./CONEXION BD/conexion.php";

    try{

    $consultaEquipo = "INSERT INTO cellphones (`departamento`, `color` , `PIN` , `EMAIL` , `SIM` , `Numero`, `Contraseña`, `serialEquipo` , `modeloEquipo` , `marcaEquipo` , `Imei1` , `Imei2` , `usoEquipo` , 'Corporativo' , 'cedula' , 'Asignado' , 'nombre')
    VALUES('$departamento' , '$color' , '$PIN' , '$EMAIL' , '$SIM' , '$Numero','$Contraseña' , '$serialEquipo' , '$modeloEquipo' , '$marcaEquipo', '$Imei1', '$Imei2', '$usoEquipo', '$Corporativo', '$cedula', '$Asignado', 'nombre')";
    
    mysqli_query($conexion2 , $consultaEquipo);


    $confirmacionUsuario = "SELECT * FROM `ussers` WHERE 'usuario' = $nombreUsuario";

    $resultado = mysqli_query($conexion2 , $confirmacionUsuario);

    if($resultado){
        $consultaPersonas = "INSERT INTO ussers(`usuario` , `documento` , `departamento`)WHERE 'usuario' = $nombreUsuario VALUES( '$nombreUsuario' , '$cedulaUsuario' , '$departamento')";
    }else{
        $consultaPersonas = "INSERT INTO ussers(`usuario` , `documento` , `departamento`) VALUES( '$nombreUsuario' , '$cedulaUsuario' , '$departamento')";
    }
    mysqli_query($conexion2 , $consultaPersonas);

    if($conexion2 -> affected_rows > 1){
echo "Se ha actualizado el inventario";
$conexion2-> close();
    }else{
echo "Algo malo ha pasado, intente de nuevo...";
    }
$conexion2-> close();

    mysqli_query($conexion2 , $consultaEquipo);
    } catch (PDOException $e) {
        echo "Error de conexión: " . $e->getMessage();
    }
}}
}

?>