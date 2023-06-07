IF EXISTS (SELECT NAME FROM SYSOBJECTS
   WHERE NAME = 'sp_CalculoTestadaIdeal' AND TYPE = 'P')
   DROP PROCEDURE sp_CalculoTestadaIdeal
GO

-- sp_CalculoTestadaIdeal 4,0
CREATE PROCEDURE sp_CalculoTestadaIdeal(@PKIdImobiliario AS INT,
		               	        @dblValor	 AS MONEY OUTPUT)
AS
	DECLARE @dblAreaConstruida      MONEY,
		@dblTestadaPrincipal    MONEY,
		@dblAreaTotalConstruida MONEY,
		@strParametros		NVARCHAR(100)

	SET @strParametros = ''+ CONVERT(NVARCHAR(30),@PKIdImobiliario) +', @dblValor OUTPUT'	
	--AREA CONTRUÍDA
	EXECUTE sp_CalculoFormulaExecutada -18, @dblAreaConstruida OUTPUT, @strParametros
	--TESTADA PRINCIPAL
	EXECUTE sp_CalculoFormulaExecutada -20, @dblTestadaPrincipal OUTPUT, @strParametros
	--AREA TOTAL CONTRUÍDA
	EXECUTE sp_CalculoFormulaExecutada -19, @dblAreaTotalConstruida OUTPUT, @strParametros

	IF @dblAreaTotalConstruida = 0 --Não Existe Divisão por Zero
		SET @dblAreaTotalConstruida = 1 
	SET @dblValor = (ISNULL(@dblAreaConstruida,0)*ISNULL(@dblTestadaPrincipal,0)/ISNULL(@dblAreaTotalConstruida,1))
