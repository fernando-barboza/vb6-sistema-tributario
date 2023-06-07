IF EXISTS (SELECT NAME FROM SYSOBJECTS
   WHERE NAME = 'sp_CalculoValorImovel' AND TYPE = 'P')
   DROP PROCEDURE sp_CalculoValorImovel
GO

-- sp_CalculoValorImovel 4,0
CREATE PROCEDURE sp_CalculoValorImovel(@PKIdImobiliario AS INT,
		               	       @dblValor	AS MONEY OUTPUT)
AS
	DECLARE @dblValorVenalTerreno    MONEY,
		@dblValorVenalEdificacao MONEY,
		@strParametros		 NVARCHAR(100)

	SET @strParametros = ''+ CONVERT(NVARCHAR(30),@PKIdImobiliario) +', @dblValor OUTPUT'
	--VALOR VENAL DO TERRENO
	EXECUTE sp_CalculoFormulaExecutada -1, @dblValorVenalTerreno OUTPUT, @strParametros
	IF (ISNULL((SELECT bytEdificado FROM tblImobiliario WHERE PKId = @PKIdImobiliario),0) =1)
		--VALOR VENAL DA EDIFICAÇÃO(CONSTRUÇÃO)
		EXECUTE sp_CalculoFormulaExecutada -2, @dblValorVenalEdificacao OUTPUT, @strParametros

	SET @dblValor = ISNULL(@dblValorVenalTerreno,0) + ISNULL(@dblValorVenalEdificacao,0)
