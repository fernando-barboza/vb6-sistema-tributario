IF EXISTS (SELECT NAME FROM SYSOBJECTS
   WHERE NAME = 'sp_ConservacaoCalcamento' AND TYPE = 'P')
   DROP PROCEDURE sp_ConservacaoCalcamento
GO

-- sp_ConservacaoCalcamento 4,0
CREATE PROCEDURE sp_ConservacaoCalcamento(@PKIdImobiliario AS INT,
				      	  @dblValor	   AS MONEY OUTPUT)
AS
	DECLARE @dblTestadaIdeal MONEY,
		@dblValorIndicado    MONEY,
		@strParametros	     NVARCHAR(100)

	SET @strParametros = ''+ CONVERT(NVARCHAR(30),@PKIdImobiliario) +', @dblValor OUTPUT'	
	--TESTADA IDEAL
	EXECUTE sp_CalculoFormulaExecutada -5, @dblTestadaIdeal OUTPUT, @strParametros

	SET @dblValorIndicado = (1.000)
	SET @dblValor = ISNULL(@dblTestadaIdeal,0) * ISNULL(@dblValorIndicado,0)
