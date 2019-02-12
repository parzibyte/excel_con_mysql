<?php
/**
 * Ejemplo de cómo usar PDO y PHPSpreadSheet para
 * exportar datos de MySQL a Excel de manera
 * fácil, rápida y segura
 *
 * @author parzibyte
 * @see https://parzibyte.me/blog/2019/02/14/leer-archivo-excel-php-phpspreadsheet/
 * @see https://parzibyte.me/blog/2018/02/12/mysql-php-pdo-crud/
 * @see https://parzibyte.me/blog/2019/02/16/php-pdo-parte-2-iterar-cursor-comprobar-si-elemento-existe/
 * @see https://parzibyte.me/blog/2018/11/08/crear-archivo-excel-php-phpspreadsheet/
 * @see https://parzibyte.me/blog/2018/10/11/sintaxis-corta-array-php/
 *
 */
require_once "vendor/autoload.php";

# Nuestra base de datos
require_once "bd.php";

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

# Obtener base de datos
$bd = obtenerBD();

$documento = new Spreadsheet();
$documento
    ->getProperties()
    ->setCreator("Luis Cabrera Benito (parzibyte)")
    ->setLastModifiedBy('Parzibyte')
    ->setTitle('Archivo exportado desde MySQL')
    ->setDescription('Un archivo de Excel exportado desde MySQL por parzibyte');

# Como ya hay una hoja por defecto, la obtenemos, no la creamos
$hojaDeProductos = $documento->getActiveSheet();
$hojaDeProductos->setTitle("Productos");

# Escribir encabezado de los productos
$encabezado = ["Código de barras", "Descripción", "Precio de compra", "Precio de venta", "Existencia"];
# El último argumento es por defecto A1 pero lo pongo para que se explique mejor
$hojaDeProductos->fromArray($encabezado, null, 'A1');

$consulta = "select codigo, descripcion, precioCompra, precioVenta, existencia from productos";
$sentencia = $bd->prepare($consulta, [
    PDO::ATTR_CURSOR => PDO::CURSOR_SCROLL,
]);
$sentencia->execute();
# Comenzamos en la 2 porque la 1 es del encabezado
$numeroDeFila = 2;
while ($producto = $sentencia->fetchObject()) {
    # Obtener los datos de la base de datos
    $codigo = $producto->codigo;
    $descripcion = $producto->descripcion;
    $precioCompra = $producto->precioCompra;
    $precioVenta = $producto->precioVenta;
    $existencia = $producto->existencia;
    # Escribirlos en el documento
    $hojaDeProductos->setCellValueByColumnAndRow(1, $numeroDeFila, $codigo);
    $hojaDeProductos->setCellValueByColumnAndRow(2, $numeroDeFila, $descripcion);
    $hojaDeProductos->setCellValueByColumnAndRow(3, $numeroDeFila, $precioCompra);
    $hojaDeProductos->setCellValueByColumnAndRow(4, $numeroDeFila, $precioVenta);
    $hojaDeProductos->setCellValueByColumnAndRow(5, $numeroDeFila, $existencia);
    $numeroDeFila++;
}

# Ahora los clientes
# Ahora sí creamos una nueva hoja
$hojaDeClientes = $documento->createSheet();
$hojaDeClientes->setTitle("Clientes");

# Escribir encabezado
$encabezado = ["Nombre", "Correo electrónico"];
# El último argumento es por defecto A1 pero lo pongo para que se explique mejor
$hojaDeClientes->fromArray($encabezado, null, 'A1');
# Obtener clientes de BD
$consulta = "select nombre, correo from clientes";
$sentencia = $bd->prepare($consulta, [
    PDO::ATTR_CURSOR => PDO::CURSOR_SCROLL,
]);
$sentencia->execute();

# Comenzamos en la 2 porque la 1 es del encabezado
$numeroDeFila = 2;
while ($cliente = $sentencia->fetchObject()) {
    # Obtener los datos de la base de datos
    $nombre = $cliente->nombre;
    $correo = $cliente->correo;

    # Escribirlos en el documento
    $hojaDeClientes->setCellValueByColumnAndRow(1, $numeroDeFila, $nombre);
    $hojaDeClientes->setCellValueByColumnAndRow(2, $numeroDeFila, $correo);
    $numeroDeFila++;
}
# Crear un "escritor"
$writer = new Xlsx($documento);
# Le pasamos la ruta de guardado
$writer->save('Exportado.xlsx');
