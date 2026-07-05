CREATE DATABASE ElVergel;
GO

USE Ferreteria_ElVergel;
GO

CREATE TABLE Usuarios (
    id_usuario INT IDENTITY(1,1) PRIMARY KEY,
    usuario NVARCHAR(50) NOT NULL UNIQUE,
    password_hash VARBINARY(32) NOT NULL,
    password_salt VARBINARY(16) NOT NULL,
    rol NVARCHAR(30) NOT NULL,
    activo BIT NOT NULL DEFAULT 1
);
GO

CREATE TABLE Catalogo_Ampliado (
    sku NVARCHAR(30) PRIMARY KEY,
    categoria NVARCHAR(100),
    producto NVARCHAR(150),
    unidad NVARCHAR(50),
    costo_promedio DECIMAL(12,2),
    precio_venta DECIMAL(12,2),
    margen DECIMAL(10,4),
    iva_incluido NVARCHAR(20),
    proveedor NVARCHAR(100),
    lead_time_dias INT,
    stock_minimo DECIMAL(12,2),
    punto_reorden DECIMAL(12,2),
    stock_maximo DECIMAL(12,2),
    stock_actual DECIMAL(12,2),
    estado_stock NVARCHAR(30),
    ubicacion NVARCHAR(100),
    fecha_alta DATE,
    observacion NVARCHAR(255)
);
GO

CREATE TABLE Control_Stock (
    sku NVARCHAR(30) PRIMARY KEY,
    categoria NVARCHAR(100),
    producto NVARCHAR(150),
    unidad NVARCHAR(50),
    stock_inicial DECIMAL(12,2),
    entradas_compras DECIMAL(12,2),
    salidas_ventas DECIMAL(12,2),
    merma_ajuste DECIMAL(12,2),
    stock_calculado DECIMAL(12,2),
    stock_fisico DECIMAL(12,2),
    diferencia DECIMAL(12,2),
    stock_minimo DECIMAL(12,2),
    punto_reorden DECIMAL(12,2),
    stock_maximo DECIMAL(12,2),
    estado NVARCHAR(30),
    ultima_reposicion DATE,
    proveedor NVARCHAR(100),
    observacion NVARCHAR(255)
);
GO

CREATE TABLE Ventas_2026 (
    id_venta NVARCHAR(30) PRIMARY KEY,
    fecha DATE,
    mes NVARCHAR(20),
    dia_semana NVARCHAR(20),
    categoria NVARCHAR(100),
    sku NVARCHAR(30),
    producto NVARCHAR(150),
    unidad NVARCHAR(50),
    cantidad DECIMAL(12,2),
    precio_unitario DECIMAL(12,2),
    total_mxn DECIMAL(12,2),
    metodo_pago NVARCHAR(50),
    tipo_cliente NVARCHAR(80),
    vendedor NVARCHAR(80),
    tipo_registro NVARCHAR(80),
    observacion NVARCHAR(255)
);
GO

CREATE OR ALTER PROCEDURE sp_crear_usuario
    @usuario NVARCHAR(50),
    @contrasena NVARCHAR(100),
    @rol NVARCHAR(30)
AS
BEGIN
    DECLARE @salt VARBINARY(16);
    DECLARE @hash VARBINARY(32);

    SET @salt = CRYPT_GEN_RANDOM(16);
    SET @hash = HASHBYTES('SHA2_256', @salt + CONVERT(VARBINARY(MAX), @contrasena));

    INSERT INTO Usuarios(usuario, password_hash, password_salt, rol)
    VALUES(@usuario, @hash, @salt, @rol);
END;
GO

CREATE OR ALTER PROCEDURE sp_login
    @usuario NVARCHAR(50),
    @contrasena NVARCHAR(100)
AS
BEGIN
    SELECT 
        id_usuario,
        usuario,
        rol
    FROM Usuarios
    WHERE usuario = @usuario
      AND activo = 1
      AND password_hash = HASHBYTES('SHA2_256', password_salt + CONVERT(VARBINARY(MAX), @contrasena));
END;
GO

CREATE OR ALTER PROCEDURE sp_resumen_general
AS
BEGIN
    SELECT 
        COUNT(*) AS registros,
        SUM(total_mxn) AS venta_total,
        AVG(total_mxn) AS promedio_registro,
        MIN(fecha) AS fecha_inicio,
        MAX(fecha) AS fecha_corte,
        COUNT(DISTINCT sku) AS productos_vendidos
    FROM Ventas_2026;
END;
GO

CREATE OR ALTER PROCEDURE sp_ventas_por_mes
AS
BEGIN
    SELECT 
        MONTH(fecha) AS numero_mes,
        mes,
        COUNT(*) AS registros,
        SUM(total_mxn) AS venta_total,
        AVG(total_mxn) AS promedio
    FROM Ventas_2026
    GROUP BY MONTH(fecha), mes
    ORDER BY MONTH(fecha);
END;
GO

CREATE OR ALTER PROCEDURE sp_ventas_por_categoria
AS
BEGIN
    SELECT 
        categoria,
        COUNT(*) AS registros,
        SUM(total_mxn) AS venta_total
    FROM Ventas_2026
    GROUP BY categoria
    ORDER BY venta_total DESC;
END;
GO

CREATE OR ALTER PROCEDURE sp_buscar_ventas
    @producto NVARCHAR(100) = '',
    @categoria NVARCHAR(100) = ''
AS
BEGIN
    SELECT TOP 300
        id_venta,
        fecha,
        categoria,
        sku,
        producto,
        cantidad,
        precio_unitario,
        total_mxn,
        metodo_pago,
        tipo_cliente
    FROM Ventas_2026
    WHERE producto LIKE '%' + @producto + '%'
      AND categoria LIKE '%' + @categoria + '%'
    ORDER BY fecha DESC;
END;
GO

CREATE OR ALTER PROCEDURE sp_consultar_stock
    @estado NVARCHAR(30) = ''
AS
BEGIN
    SELECT 
        sku,
        producto,
        categoria,
        unidad,
        stock_fisico,
        stock_minimo,
        punto_reorden,
        estado,
        proveedor
    FROM Control_Stock
    WHERE estado LIKE '%' + @estado + '%'
    ORDER BY 
        CASE estado
            WHEN 'Crítico' THEN 1
            WHEN 'Reordenar' THEN 2
            WHEN 'Vigilar' THEN 3
            ELSE 4
        END;
END;
GO

EXEC sp_crear_usuario 
    @usuario = 'admin',
    @contrasena = '1234',
    @rol = 'Administrador';
GO

