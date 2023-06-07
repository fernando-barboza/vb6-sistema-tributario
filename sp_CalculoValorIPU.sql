IF EXISTS (SELECT NAME FROM SYSOBJECTS
   WHERE NAME = 'sp_CalculoValorIPU' AND TYPE = 'P')
   DROP PROCEDURE sp_CalculoValorIPU
GO

-- sp_CalculoValorIPU 4,0
CREATE PROCEDURE sp_CalculoValorIPU(@PKIdImobiliario AS INT,
		               	    @dblValor	     AS MONEY OUTPUT)
AS
	DECLARE @dblValorVenalEdificacao MONEY,
		@dblValorIndicado  	 MONEY,
		@strParametros		 NVARCHAR(100)

	SET @strParametros = ''+ CONVERT(NVARCHAR(30),@PKIdImobiliario) +', @dblValor OUTPUT'
	EXECUTE sp_CalculoFormulaExecutada -2, @dblValorVenalEdificacao OUTPUT, @strParametros

	SET @dblValorIndicado = (0.005)
	SET @dblValor = ISNULL(@dblValorVenalEdificacao,0) * ISNULL(@dblValorIndicado,0)
