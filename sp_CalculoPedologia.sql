IF EXISTS (SELECT NAME FROM SYSOBJECTS
   WHERE NAME = 'sp_CalculoPedologia' AND TYPE = 'P')
   DROP PROCEDURE sp_CalculoPedologia
GO
-- sp_CalculoPedologia 4,0
CREATE PROCEDURE sp_CalculoPedologia(@PKIdImobiliario AS INT,
		        	     @dblValor	      AS MONEY OUTPUT)
AS
	DECLARE @intCodigoDaCaracteristica 	AS INT,
		@intUtilizacaoDaCaracteristica	AS INT

	--C�DIGO E UTILIZA��O DA CARACTERISTICA---
	SET @intUtilizacaoDaCaracteristica = (2) --TERRENO
	SET @intCodigoDaCaracteristica =(2)	 --PEDOLOGIA
	------------------------------------------
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

