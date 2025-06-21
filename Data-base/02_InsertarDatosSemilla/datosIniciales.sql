IF EXISTS (SELECT name FROM sys.databases WHERE name = N'HubLectura')
BEGIN
    ALTER DATABASE HubLectura SET SINGLE_USER WITH ROLLBACK IMMEDIATE;
    DROP DATABASE HubLectura;
END
GO

CREATE DATABASE HubLectura;
GO
USE HubLectura;
GO

-- Crear tabla Usuarios
CREATE TABLE Usuarios (
    Id INT IDENTITY(1,1) PRIMARY KEY,
    Nombre NVARCHAR(100) NOT NULL
);
GO

-- Tabla de géneros
CREATE TABLE Generos (
    Id INT PRIMARY KEY IDENTITY,
    Nombre NVARCHAR(50) UNIQUE NOT NULL
);

-- Tabla de libros
CREATE TABLE Libros (
    Id INT PRIMARY KEY IDENTITY,
    Titulo NVARCHAR(200) NOT NULL,
    Autor NVARCHAR(100),
    GeneroId INT NOT NULL,
    UrlMega NVARCHAR(MAX),
    FOREIGN KEY (GeneroId) REFERENCES Generos(Id)
);

-- Tabla de preferencias de usuario
CREATE TABLE Preferencias (
    Id INT PRIMARY KEY IDENTITY,
    UsuarioId INT NOT NULL,
    GeneroId INT NOT NULL,
    FOREIGN KEY (UsuarioId) REFERENCES Usuarios(Id),
    FOREIGN KEY (GeneroId) REFERENCES Generos(Id),
    CONSTRAINT UQ_UsuarioGenero UNIQUE (UsuarioId, GeneroId)
);

-- Crear tabla ListasDeLectura primero (depende de Usuarios y Libros)
CREATE TABLE ListasDeLectura (
    Id INT IDENTITY(1,1) PRIMARY KEY,
    UsuarioId INT NOT NULL,
    LibroId INT NOT NULL,
    FechaRegistro DATETIME2 DEFAULT SYSDATETIME(),
    CONSTRAINT UQ_UsuarioLibro UNIQUE (UsuarioId, LibroId),
    CONSTRAINT FK_Lectura_Usuario FOREIGN KEY (UsuarioId) REFERENCES Usuarios(Id),
    CONSTRAINT FK_Lectura_Libro FOREIGN KEY (LibroId) REFERENCES Libros(Id)
);
GO

-- Luego crear tabla ListasDeLecturaEstados (depende de ListasDeLectura)
CREATE TABLE ListasDeLecturaEstados (
    Id INT IDENTITY(1,1) PRIMARY KEY,
    ListaLecturaId INT NOT NULL,
    Estado VARCHAR(20) NOT NULL, -- 'Leido', 'PorLeer', 'NoGusto', 'Favorito'
    Fecha DATETIME2 DEFAULT SYSDATETIME(),
    CONSTRAINT UQ_Estado_Unico UNIQUE (ListaLecturaId, Estado),
    CONSTRAINT FK_Estado_Lectura FOREIGN KEY (ListaLecturaId) REFERENCES ListasDeLectura(Id)
);
GO