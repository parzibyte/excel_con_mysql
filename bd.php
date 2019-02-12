<?php

function obtenerBD(){
    # Mira a esquema.sql y también confgura tus credenciales
    # Recomiendo: https://parzibyte.me/blog/2018/02/12/mysql-php-pdo-crud/
    $nombre_base_de_datos = "sistema_ventas";
    $usuario = "root";
    $contraseña = "";
    try {

        $base_de_datos = new PDO('mysql:host=localhost;dbname=' . $nombre_base_de_datos, $usuario, $contraseña);
        $base_de_datos->query("set names utf8;");
        $base_de_datos->setAttribute(PDO::ATTR_EMULATE_PREPARES, false);
        $base_de_datos->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION);
        $base_de_datos->setAttribute(PDO::ATTR_DEFAULT_FETCH_MODE, PDO::FETCH_OBJ);
        return $base_de_datos;
    } catch (Exception $e) {
        # Nota: ¡en la vida real no imprimas errores!
        exit("Error obteniendo BD: " . $e->getMessage());
        return null;
    }
}