IF EXISTS (SELECT NAME FROM SYSOBJECTS
   WHERE NAME = 'sp_CalculoIluminacaoPublica' AND TYPE = 'P')
   DROP PROCEDURE sp_CalculoIluminacaoPublica
GO

-- sp_CalculoIluminacaoPublica 4,0
CREATE PROCEDURE sp_CalculoIluminacaoPublica(@PKIdImobiliario	AS INT,
				      	     @dblValor		AS MONEY OUTPUT)
AS

	--SE EXISTE ILUMINA��O P�BLICA PARA O IM�VEL
	IF (SELECT ISNULL((SELECT 1 FROM tblSecaoDeLogradouro SL,tblLogradouro L, 
					 tblMelhoramentoPublico MP, 
					 tblMelhoramentoDaSecaoDeLogradouro MS  
		            WHERE L.PKId = SL.intLogradouro
			      AND MP.PKId = MS.intMelhoramento 
			      AND SL.PKId = MS.intSecaoDeLogradouro 
			      AND L.PKId = (SELECT intLogradouro FROM tblImobiliario 
					     WHERE PKId = @PKIdImobiliario)
			      AND intCodigoDoMelhoramento = 3),0)) = 1 
	BEGIN
		DECLARE @dblTestadaIdeal MONEY,
			@dblValorIndicado    MONEY,
			@strParametros	     NVARCHAR(100)

		SET @strParametros = ''+ CONVERT(NVARCHAR(30),@PKIdImobiliario) +', @dblValor OUTPUT'	
		--TESTADA IDEAL
		EXECUTE sp_CalculoFormulaExecutada -5, @dblTestadaIdeal OUTPUT, @strParametros
		SET @dblValorIndicado = (1.000)
		SET @dblValor = ISNULL(@dblTestadaIdeal,0) * ISNULL(@dblValorIndicado,0)
	END
	ELSE
		SET @dblValor = (0.000)
