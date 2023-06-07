IF EXISTS (SELECT NAME FROM SYSOBJECTS
   WHERE NAME = 'sp_CalculoFracaoIdeal' AND TYPE = 'P')
   DROP PROCEDURE sp_CalculoFracaoIdeal
GO

-- sp_CalculoFracaoIdeal 4,0
CREATE PROCEDURE sp_CalculoFracaoIdeal(@PKIdImobiliario AS INT,
		               	       @dblValor	AS MONEY OUTPUT)
AS
	DECLARE @dblAreaConstruida      MONEY,
		@dblAreaDoTerreno       MONEY,
		@dblAreaTotalConstruida MONEY,
		@strParametros		NVARCHAR(100)

	SET @strParametros = ''+ CONVERT(NVARCHAR(30),@PKIdImobiliario) +', @dblValor OUTPUT'	
	--AREA CONTRUÍDA
	EXECUTE sp_CalculoFormulaExecutada -18, @dblAreaConstruida OUTPUT, @strParametros
	--AREA DO TERRENO
	EXECUTE sp_CalculoFormulaExecutada -17, @dblAreaDoTerreno OUTPUT, @strParametros
	--AREA TOTAL CONTRUÍDA
	EXECUTE sp_CalculoFormulaExecutada -19, @dblAreaTotalConstruida OUTPUT, @strParametros

	IF (@dblAreaTotalConstruida = 0) -- Não Existe Divisão por Zero
		SET @dblAreaTotalConstruida = 1
	SET @dblValor = (ISNULL(@dblAreaConstruida,0)*ISNULL(@dblAreaDoTerreno,0)/ISNULL(@dblAreaTotalConstruida,1))
