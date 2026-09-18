USE DB_MiniDespesas;
GO

-- 1. Procedure para Inserir Despesa
CREATE OR ALTER PROCEDURE dbo.sp_DespesaInserir
    @DataDespesa DATE,
    @Valor DECIMAL(10,2),
    @Descricao VARCHAR(100)
AS
BEGIN
    SET NOCOUNT ON;
    INSERT INTO dbo.Despesas (DataDespesa, Valor, Descricao)
    VALUES (@DataDespesa, @Valor, @Descricao);
END
GO

-- 2. Procedure para Listar Despesas do Mês Atual
CREATE OR ALTER PROCEDURE dbo.sp_DespesaListarMes
AS
BEGIN
    SET NOCOUNT ON;
    SELECT ID, DataDespesa, Valor, Descricao
    FROM dbo.Despesas
    WHERE MONTH(DataDespesa) = MONTH(GETDATE()) 
      AND YEAR(DataDespesa) = YEAR(GETDATE())
    ORDER BY DataDespesa DESC, ID DESC;
END
GO

-- 3. Procedure para Eliminar Despesa
CREATE OR ALTER PROCEDURE dbo.sp_DespesaExcluir
    @ID INT
AS
BEGIN
    SET NOCOUNT ON;
    DELETE FROM dbo.Despesas WHERE ID = @ID;
END
GO