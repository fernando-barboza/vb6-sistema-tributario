IF EXISTS (SELECT NAME FROM SYSOBJECTS
   WHERE NAME = 'sp_CalculoMT2doTerreno' AND TYPE = 'P')
   DROP PROCEDURE sp_CalculoMT2doTerreno
GO

-- sp_CalculoMT2doTerreno 11,0
CREATE PROCEDURE sp_CalculoMT2doTerreno(@PKIdImobiliario AS INT,
		               		@dblValor	 AS MONEY OUTPUT)
AS
	SET @dblValor = (SELECT ISNULL(TV.dblValor,0) 
			   FROM tblTabelaDeValor TV, tblSecaoDeLogradouro SL,
			        tblImobiliario IM
			  WHERE TV.PKId = SL.intValorDaSecao
			    AND IM.intSecoes = SL.PKId
			    AND IM.PKId = @PKIdImobiliario)
