create database if not exists sistema_ventas;
use sistema_ventas;
create table if not exists productos(
    id bigint unsigned not null auto_increment,
    codigo varchar(255) not null,
    descripcion varchar(255) not null,
    precioCompra decimal(9, 2) not null,
    precioVenta decimal(9, 2) not null,
    existencia decimal(9, 2) not null,
    primary key(id)
);

create table if not exists clientes(
    id bigint unsigned not null auto_increment,
    nombre varchar(255) not null,
    correo varchar(255) not null,
    primary key(id)
);