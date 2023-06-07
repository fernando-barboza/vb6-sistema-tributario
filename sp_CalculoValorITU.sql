IF EXISTS (SELECT NAME FROM SYSOBJECTS
   WHERE NAME = 'sp_CalculoValorITU' AND TYPE = 'P')
   DROP PROCEDURE sp_CalculoValorITU
GO

-- sp_CalculoValorITU 4,0
CREATE PROCEDURE sp_CalculoValorITU(@PKIdImobiliario AS INT,
		               	    @dblValor	     AS MONEY OUTPUT)
AS
	DECLARE @dblValorVenalTerreno MONEY,
		@dblValorIndicado     MONEY,
		@strParametros	      NVARCHAR(100)

	SET @strParametros = ''+ CONVERT(NVARCHAR(30),@PKIdImobiliario) +', @dblValor OUTPUT'

	EXECUTE sp_CalculoFormulaExecutada -1, @dblValorVenalTerreno OUTPUT,@strParametros
	IF ((SELECT bytEdificado FROM tblImobiliario WHERE PKId = @PKIdImobiliario) = 1)
		SET @dblValorIndicado = (0.005)
	ELSE
		SET @dblValorIndicado = (0.01)

	SET @dblValor = ISNULL(@dblValorVenalTerreno,0) * ISNULL(@dblValorIndicado,0)

