USE DB_MiniDespesas;
GO
-- Consulta por intervalo: inclui o início e exclui o primeiro dia do mês seguinte.
-- Não altera nem exclui registros e mantém a listagem anterior disponível.
CREATE OR ALTER PROCEDURE dbo.usp_DespesasListarPeriodo
    @Inicio date,
    @Fim date
AS
BEGIN
    SET NOCOUNT ON;
    IF @Inicio IS NULL OR @Fim IS NULL OR @Fim <= @Inicio
        THROW 50004, 'Informe um período válido.', 1;
    SELECT ID, DataDespesa, Valor, Descricao
    FROM dbo.Despesas
    WHERE DataDespesa >= @Inicio AND DataDespesa < @Fim
    ORDER BY DataDespesa DESC, ID DESC;
END;
GO
