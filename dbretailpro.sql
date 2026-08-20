/* =====================================================================================
   RetailPro_DB  ·  Base de datos del proyecto integrador
   Curso Data Analytics — Coderhouse (V2)
   Motor: MICROSOFT SQL SERVER  (2016 o superior)
   -------------------------------------------------------------------------------------
   QUE CREA ESTE SCRIPT
     1. La base de datos  RetailPro_DB
     2. Cinco tablas relacionadas: categorias, territorios, clientes, productos, ventas
     3. Todas las restricciones de integridad: PK, FK, NOT NULL, CHECK, UNIQUE
     4. Los datos: 5 categorias, 4 territorios, 40 clientes, 22 productos
        y aproximadamente 1.400 ventas distribuidas entre 2023 y 2024
     5. Una vista  vw_ventas_detalle  con el cruce completo (opcional)
     6. Un bloque de validacion que imprime el resultado de la carga

   COMO EJECUTARLO
     - Abrir en SQL Server Management Studio (SSMS) o Azure Data Studio.
     - Ejecutar el script COMPLETO (F5). Usa separadores GO, no funciona pegado
       en una herramienta que no los interprete.
     - Es REPETIBLE: se puede ejecutar las veces que haga falta. Borra y recrea
       las tablas respetando el orden inverso de dependencias.

   SI VES CARACTERES RAROS (Ã³, Ã±) AL ABRIRLO
     El archivo esta en UTF-8. En SSMS abrirlo con  Archivo > Abrir > Archivo...
     y en el desplegable del boton Abrir elegir "Abrir con..." > "Codificacion" > UTF-8.
     Todos los literales de texto usan el prefijo N'...' (Unicode) por esta razon.

   PARA QUE SIRVE CADA COSA EN EL CURSO
     M6  Power Query: conectar, perfilar, limpiar, tipar, renombrar a Dim_/Fact_ y hacer Merge
     M8  DAX: relaciones 1:N, tabla calendario y las 5 medidas core
     M10 Estadistica: media vs mediana del ticket y outliers con IQR
     M11 Proyecto final: es la base sobre la que se arma el reporte completo
   ===================================================================================== */


/* =====================================================================================
   1) CREAR LA BASE DE DATOS
   ===================================================================================== */
IF DB_ID(N'RetailPro_DB') IS NULL
BEGIN
    CREATE DATABASE RetailPro_DB;
    PRINT N'Base de datos RetailPro_DB creada.';
END
ELSE
    PRINT N'La base de datos RetailPro_DB ya existia. Se reutiliza.';
GO

USE RetailPro_DB;
GO


/* =====================================================================================
   2) LIMPIEZA PREVIA  (orden INVERSO a las dependencias)
      Primero la vista, despues los hechos, al final las dimensiones.
      Regla: se borra primero lo que TIENE claves foraneas.
   ===================================================================================== */
DROP VIEW  IF EXISTS dbo.vw_ventas_detalle;
DROP TABLE IF EXISTS dbo.ventas;
DROP TABLE IF EXISTS dbo.productos;
DROP TABLE IF EXISTS dbo.clientes;
DROP TABLE IF EXISTS dbo.territorios;
DROP TABLE IF EXISTS dbo.categorias;
GO


/* =====================================================================================
   3) DIMENSIONES  (se crean primero: no dependen de nadie)
   ===================================================================================== */

-- 3.1 categorias -----------------------------------------------------------------------
CREATE TABLE dbo.categorias (
    id_categoria      INT           NOT NULL,
    nombre_categoria  NVARCHAR(50)  NOT NULL,
    descripcion       NVARCHAR(200)     NULL,
    CONSTRAINT PK_categorias PRIMARY KEY (id_categoria),
    CONSTRAINT UQ_categorias_nombre UNIQUE (nombre_categoria)
);
GO

-- 3.2 territorios ----------------------------------------------------------------------
CREATE TABLE dbo.territorios (
    id_territorio  INT          NOT NULL,
    region         NVARCHAR(50) NOT NULL,
    pais           NVARCHAR(50) NOT NULL,
    zona           NVARCHAR(50)     NULL,
    CONSTRAINT PK_territorios PRIMARY KEY (id_territorio)
);
GO

-- 3.3 clientes -------------------------------------------------------------------------
--     email es NULLABLE a proposito: hay clientes sin correo cargado.
--     Por eso la unicidad se resuelve con un INDICE FILTRADO y no con UNIQUE,
--     ya que en SQL Server una restriccion UNIQUE solo admite UN valor NULL.
CREATE TABLE dbo.clientes (
    id_cliente      INT           NOT NULL,
    nombre          NVARCHAR(100) NOT NULL,
    email           NVARCHAR(120)     NULL,
    ciudad          NVARCHAR(80)      NULL,
    segmento        NVARCHAR(20)      NULL,
    fecha_registro  DATE          NOT NULL,
    CONSTRAINT PK_clientes PRIMARY KEY (id_cliente),
    CONSTRAINT CK_clientes_segmento
        CHECK (segmento IN (N'Consumo', N'PyME', N'Corporativo'))
);
GO

CREATE UNIQUE INDEX UX_clientes_email
    ON dbo.clientes (email)
    WHERE email IS NOT NULL;
GO

-- 3.4 productos ------------------------------------------------------------------------
--     costo es NULLABLE: hay productos nuevos sin costo cargado todavia.
--     Ese nulo es CRITICO (sin costo no hay margen) y en M6 hay que decidir
--     y justificar que hacer con el.
CREATE TABLE dbo.productos (
    id_producto      INT            NOT NULL,
    nombre_producto  NVARCHAR(120)  NOT NULL,
    id_categoria     INT            NOT NULL,
    subcategoria     NVARCHAR(50)       NULL,
    precio           DECIMAL(10,2)  NOT NULL,
    costo            DECIMAL(10,2)      NULL,
    stock            INT            NOT NULL CONSTRAINT DF_productos_stock  DEFAULT (0),
    activo           BIT            NOT NULL CONSTRAINT DF_productos_activo DEFAULT (1),
    CONSTRAINT PK_productos PRIMARY KEY (id_producto),
    CONSTRAINT FK_productos_categorias
        FOREIGN KEY (id_categoria) REFERENCES dbo.categorias (id_categoria),
    CONSTRAINT CK_productos_precio CHECK (precio >= 0),
    CONSTRAINT CK_productos_costo  CHECK (costo  >= 0),
    CONSTRAINT CK_productos_stock  CHECK (stock  >= 0)
);
GO


/* =====================================================================================
   4) TABLA DE HECHOS
      total_venta es una COLUMNA CALCULADA PERSISTIDA: se guarda fisicamente y el
      motor garantiza que siempre sea coherente con cantidad, precio y descuento.
      precio_unitario se almacena aparte de productos.precio porque registra el
      precio REAL al momento de la venta (el catalogo cambia con el tiempo).
   ===================================================================================== */
CREATE TABLE dbo.ventas (
    id_venta         INT            NOT NULL IDENTITY(1,1),
    fecha_venta      DATE           NOT NULL,
    id_cliente       INT            NOT NULL,
    id_producto      INT            NOT NULL,
    id_territorio    INT            NOT NULL,
    cantidad         INT            NOT NULL,
    precio_unitario  DECIMAL(10,2)  NOT NULL,
    descuento        DECIMAL(5,4)   NOT NULL CONSTRAINT DF_ventas_descuento DEFAULT (0),
    total_venta      AS CAST(cantidad * precio_unitario * (1 - descuento) AS DECIMAL(12,2)) PERSISTED,
    canal            NVARCHAR(20)   NOT NULL,
    CONSTRAINT PK_ventas PRIMARY KEY (id_venta),
    CONSTRAINT FK_ventas_clientes
        FOREIGN KEY (id_cliente)    REFERENCES dbo.clientes    (id_cliente),
    CONSTRAINT FK_ventas_productos
        FOREIGN KEY (id_producto)   REFERENCES dbo.productos   (id_producto),
    CONSTRAINT FK_ventas_territorios
        FOREIGN KEY (id_territorio) REFERENCES dbo.territorios (id_territorio),
    CONSTRAINT CK_ventas_cantidad  CHECK (cantidad > 0),
    CONSTRAINT CK_ventas_precio    CHECK (precio_unitario >= 0),
    CONSTRAINT CK_ventas_descuento CHECK (descuento >= 0 AND descuento < 1),
    CONSTRAINT CK_ventas_canal     CHECK (canal IN (N'Online', N'Presencial'))
);
GO

-- Indices sobre las claves foraneas: aceleran los JOIN y la importacion a Power BI.
CREATE INDEX IX_ventas_id_cliente    ON dbo.ventas (id_cliente);
CREATE INDEX IX_ventas_id_producto   ON dbo.ventas (id_producto);
CREATE INDEX IX_ventas_id_territorio ON dbo.ventas (id_territorio);
CREATE INDEX IX_ventas_fecha_venta   ON dbo.ventas (fecha_venta);
GO


/* =====================================================================================
   5) CARGA DE DIMENSIONES
   ===================================================================================== */

-- 5.1 categorias -----------------------------------------------------------------------
INSERT INTO dbo.categorias (id_categoria, nombre_categoria, descripcion) VALUES
    (1, N'Computación',    N'Notebooks, equipos de escritorio y monitores'),
    (2, N'Periféricos',    N'Teclados, mouses y accesorios de escritorio'),
    (3, N'Audio',          N'Auriculares, parlantes y micrófonos'),
    (4, N'Almacenamiento', N'Discos, unidades de estado sólido y memorias'),
    (5, N'Redes',          N'Routers, switches y equipamiento de conectividad');
GO

-- 5.2 territorios ----------------------------------------------------------------------
INSERT INTO dbo.territorios (id_territorio, region, pais, zona) VALUES
    (1, N'Norte',   N'Argentina', N'Noroeste'),
    (2, N'Centro',  N'Argentina', N'Pampeana'),
    (3, N'Litoral', N'Argentina', N'Mesopotamia'),
    (4, N'Sur',     N'Argentina', N'Patagonia');
GO

-- 5.3 clientes -------------------------------------------------------------------------
--     40 clientes. Los clientes 37 a 40 NO tienen ventas a proposito, para que la
--     consulta "clientes sin ventas" de M5 (LEFT JOIN ... IS NULL) devuelva filas.
--     Suciedad intencional para limpiar en Power Query (M6):
--       - id  7 y 23  -> email NULL
--       - id 15       -> ciudad NULL
--       - id 31       -> segmento NULL
--       - id  4, 12, 28 -> ciudad con espacios sobrantes o mayusculas inconsistentes
INSERT INTO dbo.clientes (id_cliente, nombre, email, ciudad, segmento, fecha_registro) VALUES
    ( 1, N'María López',        N'maria.lopez@mail.com',      N'Buenos Aires',  N'Consumo',     '2022-01-14'),
    ( 2, N'Carlos Ruiz',        N'carlos.ruiz@mail.com',      N'Córdoba',       N'PyME',        '2022-01-23'),
    ( 3, N'Ana Gómez',          N'ana.gomez@mail.com',        N'Rosario',       N'Consumo',     '2022-02-02'),
    ( 4, N'Pedro Sanz',         N'pedro.sanz@mail.com',       N'  Rosario ',    N'Corporativo', '2022-02-17'),
    ( 5, N'Laura Torres',       N'laura.torres@mail.com',     N'Tucumán',       N'Consumo',     '2022-03-05'),
    ( 6, N'Diego Fernández',    N'diego.fernandez@mail.com',  N'Salta',         N'PyME',        '2022-03-19'),
    ( 7, N'Sofía Ramírez',      NULL,                         N'Mendoza',       N'Consumo',     '2022-04-01'),
    ( 8, N'Martín Acosta',      N'martin.acosta@mail.com',    N'La Plata',      N'Corporativo', '2022-04-12'),
    ( 9, N'Valeria Suárez',     N'valeria.suarez@mail.com',   N'Neuquén',       N'PyME',        '2022-04-28'),
    (10, N'Javier Molina',      N'javier.molina@mail.com',    N'Bahía Blanca',  N'Consumo',     '2022-05-09'),
    (11, N'Camila Ferreyra',    N'camila.ferreyra@mail.com',  N'Posadas',       N'Consumo',     '2022-05-21'),
    (12, N'Nicolás Herrera',    N'nicolas.herrera@mail.com',  N'buenos aires',  N'PyME',        '2022-06-03'),
    (13, N'Lucía Domínguez',    N'lucia.dominguez@mail.com',  N'Santa Fe',      N'Corporativo', '2022-06-15'),
    (14, N'Federico Ibáñez',    N'federico.ibanez@mail.com',  N'Corrientes',    N'Consumo',     '2022-06-30'),
    (15, N'Paula Benítez',      N'paula.benitez@mail.com',    NULL,             N'PyME',        '2022-07-11'),
    (16, N'Gonzalo Vera',       N'gonzalo.vera@mail.com',     N'San Juan',      N'Consumo',     '2022-07-24'),
    (17, N'Florencia Ojeda',    N'florencia.ojeda@mail.com',  N'Resistencia',   N'Corporativo', '2022-08-05'),
    (18, N'Sebastián Ponce',    N'sebastian.ponce@mail.com',  N'Jujuy',         N'Consumo',     '2022-08-19'),
    (19, N'Mariana Cabrera',    N'mariana.cabrera@mail.com',  N'Río Gallegos',  N'PyME',        '2022-09-02'),
    (20, N'Emiliano Rojas',     N'emiliano.rojas@mail.com',   N'Paraná',        N'Consumo',     '2022-09-14'),
    (21, N'Agustina Núñez',     N'agustina.nunez@mail.com',   N'Córdoba',       N'Corporativo', '2022-09-27'),
    (22, N'Tomás Aguirre',      N'tomas.aguirre@mail.com',    N'Mar del Plata', N'Consumo',     '2022-10-08'),
    (23, N'Rocío Medina',       NULL,                         N'Santiago',      N'PyME',        '2022-10-20'),
    (24, N'Ignacio Peralta',    N'ignacio.peralta@mail.com',  N'Bariloche',     N'Consumo',     '2022-11-01'),
    (25, N'Julieta Sosa',       N'julieta.sosa@mail.com',     N'Formosa',       N'Corporativo', '2022-11-15'),
    (26, N'Matías Godoy',       N'matias.godoy@mail.com',     N'Catamarca',     N'Consumo',     '2022-11-28'),
    (27, N'Belén Cardozo',      N'belen.cardozo@mail.com',    N'Comodoro',      N'PyME',        '2022-12-06'),
    (28, N'Ramiro Vázquez',     N'ramiro.vazquez@mail.com',   N'CÓRDOBA',       N'Consumo',     '2022-12-14'),
    (29, N'Antonella Bruno',    N'antonella.bruno@mail.com',  N'San Luis',      N'Corporativo', '2022-12-21'),
    (30, N'Facundo Ledesma',    N'facundo.ledesma@mail.com',  N'La Rioja',      N'Consumo',     '2022-01-09'),
    (31, N'Micaela Ávila',      N'micaela.avila@mail.com',    N'Salta',         NULL,           '2022-02-24'),
    (32, N'Lautaro Quiroga',    N'lautaro.quiroga@mail.com',  N'Viedma',        N'PyME',        '2022-03-30'),
    (33, N'Guadalupe Silva',    N'guadalupe.silva@mail.com',  N'Concordia',     N'Consumo',     '2022-05-02'),
    (34, N'Bruno Castro',       N'bruno.castro@mail.com',     N'Trelew',        N'Corporativo', '2022-06-21'),
    (35, N'Delfina Moreno',     N'delfina.moreno@mail.com',   N'Rafaela',       N'Consumo',     '2022-08-09'),
    (36, N'Joaquín Maldonado',  N'joaquin.maldonado@mail.com',N'Villa María',   N'PyME',        '2022-10-03'),
    (37, N'Renata Villalba',    N'renata.villalba@mail.com',  N'Ushuaia',       N'Consumo',     '2022-11-07'),
    (38, N'Santiago Escobar',   N'santiago.escobar@mail.com', N'Goya',          N'PyME',        '2022-11-22'),
    (39, N'Abril Figueroa',     N'abril.figueroa@mail.com',   N'Pergamino',     N'Consumo',     '2022-12-02'),
    (40, N'Ezequiel Navarro',   N'ezequiel.navarro@mail.com', N'Tandil',        N'Corporativo', '2022-12-27');
GO

-- 5.4 productos ------------------------------------------------------------------------
--     22 productos. Los productos 21 y 22 NO tienen ventas a proposito, para que la
--     consulta "productos sin ventas" de M5 devuelva filas.
--     Suciedad intencional (M6):
--       - id 20 -> costo NULL  (nulo CRITICO: sin costo no se puede calcular margen)
--       - id 21 -> subcategoria NULL
INSERT INTO dbo.productos
    (id_producto, nombre_producto, id_categoria, subcategoria, precio, costo, stock, activo) VALUES
    ( 1, N'Notebook Pro 15"',          1, N'Notebooks',            1250.00,  890.00,  24, 1),
    ( 2, N'Notebook Air 13"',          1, N'Notebooks',             980.00,  700.00,  31, 1),
    ( 3, N'PC Escritorio Ryzen 7',     1, N'Equipos de escritorio',1420.00, 1010.00,  12, 1),
    ( 4, N'Monitor 27" 4K',            1, N'Monitores',             465.00,  320.00,  40, 1),
    ( 5, N'Monitor 24" Full HD',       1, N'Monitores',             215.00,  148.00,  65, 1),
    ( 6, N'Teclado Mecánico RGB',      2, N'Teclados',               95.00,   58.00,  88, 1),
    ( 7, N'Teclado Inalámbrico Slim',  2, N'Teclados',               42.00,   24.00, 120, 1),
    ( 8, N'Mouse Inalámbrico Ergo',    2, N'Mouses',                 38.00,   21.00, 145, 1),
    ( 9, N'Base Notebook Aluminio',    2, N'Accesorios',             55.00,   30.00,  72, 1),
    (10, N'Auriculares BT Pro',        3, N'Auriculares',           128.00,   82.00,  54, 1),
    (11, N'Auriculares In-Ear USB-C',  3, N'Auriculares',            34.00,   18.00, 160, 1),
    (12, N'Parlante Bluetooth 20W',    3, N'Parlantes',              89.00,   54.00,  47, 1),
    (13, N'Micrófono Condenser USB',   3, N'Micrófonos',            115.00,   72.00,  29, 1),
    (14, N'SSD NVMe 1TB',              4, N'Discos sólidos',        132.00,   88.00,  58, 1),
    (15, N'SSD SATA 480GB',            4, N'Discos sólidos',         61.00,   39.00,  95, 1),
    (16, N'Disco Externo 2TB',         4, N'Discos externos',        78.00,   51.00,  63, 1),
    (17, N'Memoria RAM 16GB DDR4',     4, N'Memorias',               54.00,   33.00, 110, 1),
    (18, N'Router WiFi 6 AX1800',      5, N'Routers',               118.00,   74.00,  41, 1),
    (19, N'Switch 8 puertos Gigabit',  5, N'Switches',               46.00,   27.00,  77, 1),
    (20, N'Repetidor WiFi Mesh',       5, N'Accesorios de red',      72.00,    NULL,  35, 1),
    (21, N'Webcam 4K Streaming',       2, NULL,                      96.00,   61.00,  18, 1),
    (22, N'Hub USB-C 7 en 1',          2, N'Accesorios',             49.00,   28.00,   0, 0);
GO


/* =====================================================================================
   6) CARGA DE LA TABLA DE HECHOS  (generador determinista)
      No usa numeros aleatorios: todos los valores se derivan aritmeticamente del
      numero de fila. Eso hace que TODOS obtengan exactamente los mismos datos.

      El plan de carga modela una historia de negocio concreta, la del brief de M1:
        - La empresa crece en 2024 en Centro, Litoral y Sur.
        - La empresa crece un 17,57% interanual...
        - ...pero la region NORTE cae un 13,92%.
        - La caida de Norte se explica por un desplome de la categoria Computacion
          (-30,05%, el producto de ticket alto), no por una baja general de unidades.
      Verificalo vos mismo con las consultas del bloque 8.
   ===================================================================================== */

-- 6.1 Plan de carga: cuantas ventas genera cada region en cada mes ----------------------
IF OBJECT_ID(N'tempdb..#plan_carga') IS NOT NULL DROP TABLE #plan_carga;

SELECT
    a.anio,
    m.mes,
    t.id_territorio,
    CAST(ROUND(
        t.base_mensual
        -- Estacionalidad: enero y febrero flojos, noviembre y diciembre pico.
      * CASE m.mes
            WHEN  1 THEN 0.80 WHEN  2 THEN 0.78 WHEN  3 THEN 0.92
            WHEN  4 THEN 0.95 WHEN  5 THEN 1.00 WHEN  6 THEN 1.05
            WHEN  7 THEN 1.10 WHEN  8 THEN 0.98 WHEN  9 THEN 1.02
            WHEN 10 THEN 1.08 WHEN 11 THEN 1.25 ELSE 1.40
        END
        -- Crecimiento 2024 por region. Norte casi no baja en unidades: lo que cae
        -- es su facturacion, por el cambio de mix de productos (ver 6.2).
      * CASE
            WHEN a.anio = 2024 AND t.id_territorio = 1 THEN 0.97
            WHEN a.anio = 2024 AND t.id_territorio = 2 THEN 1.22
            WHEN a.anio = 2024 AND t.id_territorio = 3 THEN 1.15
            WHEN a.anio = 2024 AND t.id_territorio = 4 THEN 1.10
            ELSE 1.00
        END
    , 0) AS INT) AS cant_ventas
INTO #plan_carga
FROM      (VALUES (2023), (2024)) AS a (anio)
CROSS JOIN (VALUES (1),(2),(3),(4),(5),(6),(7),(8),(9),(10),(11),(12)) AS m (mes)
CROSS JOIN (VALUES (1, 14), (2, 22), (3, 11), (4, 8)) AS t (id_territorio, base_mensual);
GO

-- 6.2 Expansion: una fila por venta, con su numero global y su producto ----------------
IF OBJECT_ID(N'tempdb..#ventas_gen') IS NOT NULL DROP TABLE #ventas_gen;

WITH tally AS (
    SELECT TOP (200) ROW_NUMBER() OVER (ORDER BY (SELECT NULL)) AS i
    FROM sys.all_objects
),
expandido AS (
    SELECT
        p.anio,
        p.mes,
        p.id_territorio,
        ROW_NUMBER() OVER (ORDER BY p.anio, p.mes, p.id_territorio, ta.i) AS n
    FROM #plan_carga p
    INNER JOIN tally ta ON ta.i <= p.cant_ventas
)
SELECT
    e.anio,
    e.mes,
    e.id_territorio,
    e.n,
    CASE
        -- Norte 2024: un tercio de las ventas se desplaza a productos de ticket bajo
        -- (ids 6 a 20), es decir FUERA de la categoria Computacion (ids 1 a 5).
        WHEN e.anio = 2024 AND e.id_territorio = 1 AND (e.n % 3) = 0
            THEN 6 + ((e.n * 13) % 15)
        ELSE     1 + ((e.n * 13) % 20)
    END AS id_producto
INTO #ventas_gen
FROM expandido e;
GO

-- 6.3 Insercion final ------------------------------------------------------------------
INSERT INTO dbo.ventas
    (fecha_venta, id_cliente, id_producto, id_territorio, cantidad, precio_unitario, descuento, canal)
SELECT
    -- Dia del mes derivado del numero de fila, siempre dentro del rango valido.
    DATEFROMPARTS(
        g.anio,
        g.mes,
        1 + ((g.n * 17) % DAY(EOMONTH(DATEFROMPARTS(g.anio, g.mes, 1))))
    )                                                              AS fecha_venta,

    -- Solo clientes 1 a 36: los clientes 37 a 40 quedan sin ventas.
    1 + ((g.n * 7) % 36)                                           AS id_cliente,

    g.id_producto,
    g.id_territorio,

    -- Los productos caros se venden de a 1 o 2; los baratos, de a 1 a 4.
    CASE WHEN g.id_producto <= 5
         THEN 1 + ((g.n * 5) % 2)
         ELSE 1 + ((g.n * 5) % 4)
    END                                                            AS cantidad,

    -- Precio historico: en 2023 la lista estaba un 8% por debajo de la actual.
    CAST(pr.precio * CASE WHEN g.anio = 2023 THEN 0.92 ELSE 1.00 END AS DECIMAL(10,2))
                                                                   AS precio_unitario,

    -- Descuentos escalonados: 55% sin descuento, luego 5%, 10% y 15%.
    CASE
        WHEN ((g.n * 29) % 100) < 55 THEN 0.0000
        WHEN ((g.n * 29) % 100) < 80 THEN 0.0500
        WHEN ((g.n * 29) % 100) < 94 THEN 0.1000
        ELSE                              0.1500
    END                                                            AS descuento,

    -- El canal Online crece de 2023 a 2024 (36% -> 48% de las operaciones).
    CASE WHEN ((g.n * 11) % 100) < CASE WHEN g.anio = 2023 THEN 36 ELSE 48 END
         THEN N'Online'
         ELSE N'Presencial'
    END                                                            AS canal
FROM #ventas_gen g
INNER JOIN dbo.productos pr ON pr.id_producto = g.id_producto
ORDER BY 1, g.n;
GO

DROP TABLE #plan_carga;
DROP TABLE #ventas_gen;
GO


/* =====================================================================================
   7) VISTA DE APOYO  (opcional)
      Es la "vista base del proyecto" de M5: el INNER JOIN de las cuatro tablas.
      Sirve para explorar los datos y para validar resultados.

      IMPORTANTE PARA M6 Y M8: para Power BI NO conviene importar esta vista.
      Hay que importar las CINCO TABLAS por separado y armar el esquema en estrella
      en el modelo. Esta vista es una tabla plana y anula el ejercicio de modelado.
   ===================================================================================== */
GO
CREATE VIEW dbo.vw_ventas_detalle AS
SELECT
    v.id_venta,
    v.fecha_venta,
    c.id_cliente,
    c.nombre           AS nombre_cliente,
    c.segmento,
    c.ciudad,
    t.id_territorio,
    t.region,
    t.pais,
    t.zona,
    p.id_producto,
    p.nombre_producto,
    p.subcategoria,
    cat.id_categoria,
    cat.nombre_categoria,
    v.cantidad,
    v.precio_unitario,
    v.descuento,
    v.total_venta,
    v.canal
FROM dbo.ventas       v
INNER JOIN dbo.clientes    c   ON c.id_cliente    = v.id_cliente
INNER JOIN dbo.productos   p   ON p.id_producto   = v.id_producto
INNER JOIN dbo.territorios t   ON t.id_territorio = v.id_territorio
INNER JOIN dbo.categorias  cat ON cat.id_categoria = p.id_categoria;
GO


/* =====================================================================================
   8) VALIDACION DE LA CARGA
      Ejecutar y revisar. Si algun conteo da 0, algo fallo.
   ===================================================================================== */

PRINT N'--- Conteo de filas por tabla ---';
SELECT N'categorias'  AS tabla, COUNT(*) AS filas FROM dbo.categorias
UNION ALL SELECT N'territorios', COUNT(*) FROM dbo.territorios
UNION ALL SELECT N'clientes',    COUNT(*) FROM dbo.clientes
UNION ALL SELECT N'productos',   COUNT(*) FROM dbo.productos
UNION ALL SELECT N'ventas',      COUNT(*) FROM dbo.ventas;

PRINT N'--- Rango de fechas y facturacion total ---';
SELECT
    MIN(fecha_venta)          AS primera_venta,
    MAX(fecha_venta)          AS ultima_venta,
    COUNT(*)                  AS total_ventas,
    SUM(total_venta)          AS facturacion_total,
    CAST(AVG(total_venta) AS DECIMAL(12,2)) AS ticket_promedio
FROM dbo.ventas;

PRINT N'--- Facturacion por anio y region  (la historia del brief M1) ---';
SELECT
    t.region,
    SUM(CASE WHEN YEAR(v.fecha_venta) = 2023 THEN v.total_venta ELSE 0 END) AS ventas_2023,
    SUM(CASE WHEN YEAR(v.fecha_venta) = 2024 THEN v.total_venta ELSE 0 END) AS ventas_2024,
    CAST(
        ( SUM(CASE WHEN YEAR(v.fecha_venta) = 2024 THEN v.total_venta ELSE 0 END)
        - SUM(CASE WHEN YEAR(v.fecha_venta) = 2023 THEN v.total_venta ELSE 0 END) )
        / NULLIF(SUM(CASE WHEN YEAR(v.fecha_venta) = 2023 THEN v.total_venta ELSE 0 END), 0) * 100
    AS DECIMAL(6,2)) AS variacion_pct
FROM dbo.ventas v
INNER JOIN dbo.territorios t ON t.id_territorio = v.id_territorio
GROUP BY t.region
ORDER BY variacion_pct;

PRINT N'--- Que categoria explica la caida del Norte ---';
SELECT
    cat.nombre_categoria,
    SUM(CASE WHEN YEAR(v.fecha_venta) = 2023 THEN v.total_venta ELSE 0 END) AS norte_2023,
    SUM(CASE WHEN YEAR(v.fecha_venta) = 2024 THEN v.total_venta ELSE 0 END) AS norte_2024,
    CAST(
        ( SUM(CASE WHEN YEAR(v.fecha_venta) = 2024 THEN v.total_venta ELSE 0 END)
        - SUM(CASE WHEN YEAR(v.fecha_venta) = 2023 THEN v.total_venta ELSE 0 END) )
        / NULLIF(SUM(CASE WHEN YEAR(v.fecha_venta) = 2023 THEN v.total_venta ELSE 0 END), 0) * 100
    AS DECIMAL(6,2)) AS variacion_pct
FROM dbo.ventas v
INNER JOIN dbo.territorios t   ON t.id_territorio = v.id_territorio
INNER JOIN dbo.productos   p   ON p.id_producto   = v.id_producto
INNER JOIN dbo.categorias  cat ON cat.id_categoria = p.id_categoria
WHERE t.region = N'Norte'
GROUP BY cat.nombre_categoria
ORDER BY variacion_pct;

PRINT N'--- Ventas por canal y anio ---';
SELECT
    YEAR(fecha_venta) AS anio,
    canal,
    COUNT(*)          AS operaciones,
    SUM(total_venta)  AS facturacion
FROM dbo.ventas
GROUP BY YEAR(fecha_venta), canal
ORDER BY anio, canal;

PRINT N'--- Resumen de nulos a resolver en Power Query (M6) ---';
SELECT N'clientes.email NULL'        AS hallazgo, COUNT(*) AS filas FROM dbo.clientes  WHERE email    IS NULL
UNION ALL SELECT N'clientes.ciudad NULL',         COUNT(*) FROM dbo.clientes  WHERE ciudad   IS NULL
UNION ALL SELECT N'clientes.segmento NULL',       COUNT(*) FROM dbo.clientes  WHERE segmento IS NULL
UNION ALL SELECT N'productos.costo NULL',         COUNT(*) FROM dbo.productos WHERE costo    IS NULL
UNION ALL SELECT N'productos.subcategoria NULL',  COUNT(*) FROM dbo.productos WHERE subcategoria IS NULL;

PRINT N'--- Las filas concretas que hay que limpiar (miralas antes de M6) ---';
SELECT id_cliente, nombre, email, N'[' + ciudad + N']' AS ciudad_entre_corchetes, segmento
FROM dbo.clientes
WHERE id_cliente IN (4, 7, 12, 15, 23, 28, 31)
ORDER BY id_cliente;
-- Los corchetes hacen visibles los espacios sobrantes de la ciudad del cliente 4.
-- Los clientes 12 y 28 tienen la misma ciudad que otros pero escrita distinto
-- ('buenos aires', 'CORDOBA'): en Power Query se normalizan con Recortar y Mayus. inicial.

SELECT id_producto, nombre_producto, subcategoria, precio, costo
FROM dbo.productos
WHERE costo IS NULL OR subcategoria IS NULL
ORDER BY id_producto;

PRINT N'--- Registros sin par (para las consultas de M5) ---';
SELECT N'clientes sin ventas' AS caso, COUNT(*) AS filas
FROM dbo.clientes c LEFT JOIN dbo.ventas v ON v.id_cliente = c.id_cliente
WHERE v.id_cliente IS NULL
UNION ALL
SELECT N'productos sin ventas', COUNT(*)
FROM dbo.productos p LEFT JOIN dbo.ventas v ON v.id_producto = p.id_producto
WHERE v.id_producto IS NULL;

PRINT N'';
PRINT N'====================================================================';
PRINT N'  RetailPro_DB lista. Siguiente paso: conectar Power BI Desktop.';
PRINT N'  Obtener datos > SQL Server > Servidor: (local) o el nombre de tu';
PRINT N'  instancia · Base de datos: RetailPro_DB · Modo: Importar';
PRINT N'  Seleccionar las 5 TABLAS (no la vista) y pulsar Transformar datos.';
PRINT N'====================================================================';
GO
