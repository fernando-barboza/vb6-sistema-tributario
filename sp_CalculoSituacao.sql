IF EXISTS (SELECT NAME FROM SYSOBJECTS
   WHERE NAME = 'sp_CalculoSituacao' AND TYPE = 'P')
   DROP PROCEDURE sp_CalculoSituacao
GO
-- sp_CalculoSituacao 4,0
CREATE PROCEDURE sp_CalculoSituacao(@PKIdImobiliario AS INT,
		        	    @dblValor	     AS MONEY OUTPUT)
AS
	DECLARE @intCodigoDaCaracteristica 	AS INT,
		@intUtilizacaoDaCaracteristica	AS INT

	--CÓDIGO E UTILIZAÇÃO DA CARACTERISTICA------------
	SET @intUtilizacaoDaCaracteristica = (2) --TERRENO
	SET @intCodigoDaCaracteristica =(1)	 --SITUAÇÃO
	---------------------------------------------------

	SET @dblValor = (SELECT ISNULL(dblValor,0)  
			   FROM tblCaracteristicaGeral A , tblDetalheDaCaracteristica B, 
			        tblCaracteristicasDoImovel C,	tblImobiliario D, 
				tblTabelaDeValor E
		          WHERE A.PKId = B.intCaracteristica
			    AND A.PKId = C.intCodigoCaracteristicaGeral
			    AND B.PKId = C.intCodigoDetalheDaCaracteristica
			    AND D.PKId = C.intCodigoImobiliario
			    AND D.PKId = @PKIdImobiliario
			    AND E.PKId = B.intTabelaDeValores
			    AND A.intUtilizacaoDaCaracteristica = @intUtilizacaoDaCaracteristica
			    AND intCodigoCaracteristicaGeral = (SELECT PKId 
								  FROM tblCaracteristicaGeral A 
			        				 WHERE intutilizacaodacaracteristica = @intUtilizacaoDaCaracteristica  
							           AND intCodigoDaCaracteristica = @intCodigoDaCaracteristica))		
