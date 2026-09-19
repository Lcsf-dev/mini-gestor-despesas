USE DB_MiniDespesas;
GO
-- Evolução aditiva: preserva tabela, registros e procedures anteriores.
CREATE OR ALTER PROCEDURE dbo.usp_DespesasListar
AS
BEGIN
    SET NOCOUNT ON;
    SELECT ID, DataDespesa, Valor, Descricao
    FROM dbo.Despesas ORDER BY DataDespesa DESC, ID DESC;
END;
GO
CREATE OR ALTER PROCEDURE dbo.usp_DespesasInserir
    @DataDespesa date, @Valor decimal(10,2), @Descricao varchar(100)
AS
BEGIN
    SET NOCOUNT ON;
    IF @DataDespesa IS NULL OR @Valor IS NULL OR @Valor <= 0
        THROW 50001, 'Data e valor positivo são obrigatórios.', 1;
    SET @Descricao = LTRIM(RTRIM(@Descricao));
    IF @Descricao IS NULL OR @Descricao = ''
        THROW 50002, 'Informe a descrição.', 1;
    INSERT dbo.Despesas(DataDespesa, Valor, Descricao)
    VALUES (@DataDespesa, @Valor, @Descricao);
END;
GO
CREATE OR ALTER PROCEDURE dbo.usp_DespesasEditar
    @ID int, @DataDespesa date, @Valor decimal(10,2), @Descricao varchar(100),
    @DescricaoOriginal varchar(100), @ValorOriginal decimal(10,2), @DataOriginal date
AS
BEGIN
    SET NOCOUNT ON;
    IF @DataDespesa IS NULL OR @Valor IS NULL OR @Valor <= 0
        THROW 50001, 'Data e valor positivo são obrigatórios.', 1;
    SET @Descricao = LTRIM(RTRIM(@Descricao));
    IF @Descricao IS NULL OR @Descricao = ''
        THROW 50002, 'Informe a descrição.', 1;
    UPDATE dbo.Despesas SET Descricao=@Descricao, Valor=@Valor, DataDespesa=@DataDespesa
    WHERE ID=@ID AND CONVERT(varbinary(100), Descricao)=CONVERT(varbinary(100), @DescricaoOriginal)
        AND Valor=@ValorOriginal AND DataDespesa=@DataOriginal;
    IF @@ROWCOUNT = 0
        THROW 50003, 'Registro removido ou alterado. Cancele a edição e atualize a lista.', 1;
END;
GO
CREATE OR ALTER PROCEDURE dbo.usp_DespesasExcluir
    @ID int, @DescricaoOriginal varchar(100), @ValorOriginal decimal(10,2), @DataOriginal date
AS
BEGIN
    SET NOCOUNT ON;
    DELETE dbo.Despesas
    WHERE ID=@ID AND CONVERT(varbinary(100), Descricao)=CONVERT(varbinary(100), @DescricaoOriginal)
        AND Valor=@ValorOriginal AND DataDespesa=@DataOriginal;
    IF @@ROWCOUNT = 0
        THROW 50003, 'Registro removido ou alterado. Atualize a lista antes de excluir.', 1;
END;
GO
