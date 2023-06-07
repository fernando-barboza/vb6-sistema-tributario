IF EXISTS (SELECT NAME FROM SYSOBJECTS
   WHERE NAME = 'sp_CalculoFatorDeCorrecao' AND TYPE = 'P')
   DROP PROCEDURE sp_CalculoFatorDeCorrecao
GO

-- sp_CalculoFatorDeCorrecao 4,0
CREATE PROCEDURE sp_CalculoFatorDeCorrecao(@PKIdImobiliario AS INT,
		              		   @dblValor	    AS MONEY OUTPUT)
AS
	DECLARE @PKIdTipo 	          INT,
		@intCodCaracteristica     INT,
		@dblAuxFC		  MONEY

	--INICIALIZAÇÃO DE VARIÁVEIS-------------------------------------------------------
	SET @dblValor = (1.000)
	SET @PKIdTipo = (SELECT PKId FROM tblCaracteristicaGeral  
			  WHERE intutilizacaodacaracteristica = 3 
			    AND intCodigoDaCaracteristica = 1)
	--FIM DA INICIALIZAÇÃO-----------------------------------------------------------
	
	DECLARE c_FatorDeCorrecao CURSOR 
		FOR
		SELECT PKId
		  FROM tblCaracteristicaGeral  
		 WHERE intutilizacaodacaracteristica = 3 
		  AND intCodigoDaCaracteristica <> 1
	OPEN	c_FatorDeCorrecao
	FETCH	c_FatorDeCorrecao INTO
		@intCodCaracteristica
	WHILE @@FETCH_STATUS = 0
	BEGIN	
		DECLARE c_FatorDeCorrecaoValor CURSOR 
			FOR	
			SELECT TV.dblValor
			  FROM tblCaracteristicaGeral A , tblDetalheDaCaracteristica B, 
			       tblCaracteristicasDoImovel C, tblImobiliario D, 
			       tblTabelaDeValor TV
			 WHERE A.PKId = B.intCaracteristica
			   AND A.PKId = C.intCodigoCaracteristicaGeral
			   AND B.PKId = C.intCodigoDetalheDaCaracteristica
			   AND D.PKId = C.intCodigoImobiliario
			   AND TV.PKId = B.intTabelaDeValores
			   AND D.PKId = @PKIdImobiliario
			   AND A.intUtilizacaoDaCaracteristica = 3
			   AND intCodigoCaracteristicaGeral = @intCodCaracteristica
		OPEN	c_FatorDeCorrecaoValor
		FETCH	c_FatorDeCorrecaoValor INTO
			@dblAuxFC
		WHILE @@FETCH_STATUS = 0
		BEGIN		
			SET @dblValor = @dblValor * @dblAuxFC
			FETCH c_FatorDeCorrecaoValor INTO
				@dblAuxFC
		END
		CLOSE c_FatorDeCorrecaoValor
		DEALLOCATE c_FatorDeCorrecaoValor					
		FETCH c_FatorDeCorrecao INTO
			@intCodCaracteristica
	END
	CLOSE c_FatorDeCorrecao
	DEALLOCATE c_FatorDeCorrecao

