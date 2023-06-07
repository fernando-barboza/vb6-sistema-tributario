IF EXISTS (SELECT NAME FROM SYSOBJECTS
   WHERE NAME = 'sp_CalculoValorTerreno' AND TYPE = 'P')
   DROP PROCEDURE sp_CalculoValorTerreno
GO

-- sp_CalculoFormulaExecutada 97, NULL , '4,@dblValor OUTPUT'
-- sp_CalculoValorTerreno 48,0
CREATE PROCEDURE sp_CalculoValorTerreno(@PKIdImobiliario AS INT,
		               		@dblValor	 AS MONEY OUTPUT)
AS
	DECLARE @dblMT2doTerreno MONEY,
		@dblPedologia	 MONEY,
		@dblTopografia   MONEY,
		@dblSituacao	 MONEY,
		@strParametros	 NVARCHAR(100)

	SET @strParametros = ''+ CONVERT(NVARCHAR(30),@PKIdImobiliario) +', @dblValor OUTPUT'
	IF (SELECT ISNULL((SELECT bytEdificado FROM tblImobiliario
			    WHERE PKId = @PKIdImobiliario),0)) = 1
		--FRAÇÃO IDEAL
		EXECUTE sp_CalculoFormulaExecutada -4, @dblValor OUTPUT, @strParametros
	ELSE
		--AREA DO TERRENO
		EXECUTE sp_CalculoFormulaExecutada -17, @dblValor OUTPUT, @strParametros
	--MT² DO TERRENO
	EXECUTE sp_CalculoFormulaExecutada -16, @dblMT2doTerreno OUTPUT, @strParametros
	--VALOR DA PEDOLOGIA
	EXECUTE sp_CalculoFormulaExecutada -15, @dblPedologia OUTPUT, @strParametros
	--VALOR DA TOPOGRAFIA
	EXECUTE sp_CalculoFormulaExecutada -13, @dblTopografia OUTPUT, @strParametros
	--VALOR DA SITUAÇÃO
	EXECUTE sp_CalculoFormulaExecutada -14, @dblSituacao OUTPUT, @strParametros

	SET @dblValor = ISNULL(@dblValor,0)*ISNULL(@dblMT2doTerreno,0)*ISNULL(@dblPedologia,0)
			* ISNULL(@dblTopografia,0)*ISNULL(@dblSituacao,0)

	
	