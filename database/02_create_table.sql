USE DB_MiniDespesas;
GO

IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[Despesas]') AND type in (N'U'))
BEGIN
    CREATE TABLE dbo.Despesas (
        ID INT IDENTITY(1,1) PRIMARY KEY,
        DataDespesa DATE NOT NULL,
        Valor DECIMAL(10,2) NOT NULL,
        Descricao VARCHAR(100) NOT NULL,
        DataCriacao DATETIME DEFAULT GETDATE()
    );
END
GO