<?php
/**
 * Ejemplo de cómo usar PDO y PHPSpreadSheet para
 * importar datos de Excel a MySQL de manera
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

# Cargar clases instaladas por Composer
require_once "vendor/autoload.php";

# Nuestra base de datos
require_once "bd.php";

# Indicar que usaremos el IOFactory
use PhpOffice\PhpSpreadsheet\IOFactory;

# Obtener conexión o salir en caso de error, mira bd.php
$bd = obtenerBD();

# El archivo a importar
# Recomiendo poner la ruta absoluta si no está junto al script
$rutaArchivo = "Ventas.xlsx";
$documento = IOFactory::load($rutaArchivo);

# Se espera que en la primera hoja estén los productos
$hojaDeProductos = $documento->getSheet(0);

# Preparar base de datos para que los inserts sean rápidos
$bd->beginTransaction();

# Preparar sentencia de productos
$sentencia = $bd->prepare("insert into productos
(codigo, descripcion, precioCompra, precioVenta, existencia) values
(?, ?, ? ,?, ? )");

# Calcular el máximo valor de la fila como entero, es decir, el
# límite de nuestro ciclo
$numeroMayorDeFila = $hojaDeProductos->getHighestRow(); // Numérico
$letraMayorDeColumna = $hojaDeProductos->getHighestColumn(); // Letra
# Convertir la letra al número de columna correspondiente
$numeroMayorDeColumna = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($letraMayorDeColumna);

// Recorrer filas; comenzar en la fila 2 porque omitimos el encabezado
for ($indiceFila = 2; $indiceFila <= $numeroMayorDeFila; $indiceFila++) {

    # Las columnas están en este orden:
    # Código de barras, Descripción, Precio de Compra, Precio de Venta, Existencia
    $codigoDeBarras = $hojaDeProductos->getCellByColumnAndRow(1, $indiceFila);
    $descripcion = $hojaDeProductos->getCellByColumnAndRow(2, $indiceFila);
    $precioCompra = $hojaDeProductos->getCellByColumnAndRow(3, $indiceFila);
    $precioVenta = $hojaDeProductos->getCellByColumnAndRow(4, $indiceFila);
    $existencia = $hojaDeProductos->getCellByColumnAndRow(5, $indiceFila);
    $sentencia->execute([$codigoDeBarras, $descripcion, $precioCompra, $precioVenta, $existencia]);
}

# Ahora vamos con los clientes
$sentencia = $bd->prepare("insert into clientes
(nombre, correo) values (?, ?)");

# Se espera que en la segunda hoja estén los clientes
$hojaDeClientes = $documento->getSheet(1);

# Calcular el máximo valor de la fila como entero, es decir, el
# límite de nuestro ciclo
$numeroMayorDeFila = $hojaDeClientes->getHighestRow(); // Numérico
$letraMayorDeColumna = $hojaDeClientes->getHighestColumn(); // Letra
# Convertir la letra al número de columna correspondiente
$numeroMayorDeColumna = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($letraMayorDeColumna);

// Recorrer filas; comenzar en la fila 2 porque omitimos el encabezado
for ($indiceFila = 2; $indiceFila <= $numeroMayorDeFila; $indiceFila++) {

    # Las columnas están en este orden:
    # Nombre, Correo electrónico
    $nombre = $hojaDeClientes->getCellByColumnAndRow(1, $indiceFila);
    $correoElectronico = $hojaDeClientes->getCellByColumnAndRow(2, $indiceFila);
    $sentencia->execute([$nombre, $correoElectronico]);
}

# Hacer commit para guardar cambios de la base de datos
$bd->commit();
