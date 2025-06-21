-- Crear base de datos
CREATE DATABASE HubLectura;
GO

USE HubLectura;
GO

-- Tabla de usuarios
CREATE TABLE Usuarios (
    Id INT PRIMARY KEY IDENTITY,
    Nombre NVARCHAR(100) NOT NULL,
    Usuario NVARCHAR(50) UNIQUE NOT NULL,
    Contrasena NVARCHAR(100) NOT NULL
);

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

-- Tabla de listas de lectura
CREATE TABLE ListasLectura (
    Id INT PRIMARY KEY IDENTITY,
    UsuarioId INT NOT NULL,
    LibroId INT NOT NULL,
    Estado NVARCHAR(20) NOT NULL CHECK (Estado IN ('Leido', 'PorLeer', 'NoGusto')),
    FechaAgregado DATETIME DEFAULT SYSDATETIME(),
    FOREIGN KEY (UsuarioId) REFERENCES Usuarios(Id),
    FOREIGN KEY (LibroId) REFERENCES Libros(Id),
    CONSTRAINT UQ_UsuarioLibro UNIQUE (UsuarioId, LibroId)
);

CREATE TABLE ListasDeLecturaEstados (
    Id INT IDENTITY(1,1) PRIMARY KEY,
    ListaLecturaId INT NOT NULL,
    Estado VARCHAR(20) NOT NULL, -- Ej: 'Leido', 'PorLeer', 'Favorito', 'NoGusto'
    Fecha DATETIME2 DEFAULT SYSDATETIME(),
    CONSTRAINT UQ_Estado_Unico UNIQUE (ListaLecturaId, Estado),
    CONSTRAINT FK_Estado_Lectura FOREIGN KEY (ListaLecturaId) REFERENCES ListasDeLectura(Id)
);


-- Tabla recomendaciones automáticas
CREATE TABLE Recomendaciones (
    Id INT PRIMARY KEY IDENTITY,
    UsuarioId INT NOT NULL,
    LibroId INT NOT NULL,
    FechaRecomendada DATETIME DEFAULT SYSDATETIME(),
    FOREIGN KEY (UsuarioId) REFERENCES Usuarios(Id),
    FOREIGN KEY (LibroId) REFERENCES Libros(Id)
);

-- Insertar géneros iniciales
INSERT INTO Generos (Nombre) VALUES
('Ficción'), 
('Misterio'), 
('Fantasía'), 
('Ciencia Ficción'), 
('Historia');

-- Insertar libros de ejemplo
INSERT INTO Libros (Titulo, Autor, GeneroId, UrlMega) VALUES
('1984', 'George Orwell', 4, 'https://mega.nz/ejemplo1'),
('El nombre del viento', 'Patrick Rothfuss', 3, 'https://mega.nz/ejemplo2'),
('Sherlock Holmes', 'Arthur Conan Doyle', 2, 'https://mega.nz/ejemplo3'),
('Los pilares de la Tierra', 'Ken Follett', 5, 'https://mega.nz/ejemplo4');
