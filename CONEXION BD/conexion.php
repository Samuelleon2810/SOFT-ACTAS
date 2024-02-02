<?php
require_once "config.php";

try{

$conexion2 = new mysqli($servidor,$nombre_usuario,$contraseña,$nombre_bd,'3306');

echo "Conexion Exitosa";

} catch (PDOException $e) {
    echo "Error de conexión: " . $e->getMessage();
}
?>