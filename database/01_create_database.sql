IF NOT EXISTS (SELECT name FROM sys.databases WHERE name = N'DB_MiniDespesas')
BEGIN
    CREATE DATABASE DB_MiniDespesas;
END
GO