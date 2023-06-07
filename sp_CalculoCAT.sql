IF EXISTS (SELECT NAME FROM SYSOBJECTS
   WHERE NAME = 'sp_CalculoCAT' AND TYPE = 'P')
   DROP PROCEDURE sp_CalculoCAT
GO

-- sp_CalculoCAT 4,0
CREATE PROCEDURE sp_CalculoCAT(@PKIdImobiliario AS INT,
		               @dblValor	AS MONEY OUTPUT)
AS
	DECLARE @PKIdTipo 	          INT,
		@PKIdTipoDET	          INT,
		@intCodCaracteristica     INT,
		@intDetalheCaracteristica INT,
		@dblAuxCAT		  MONEY


	--INICIALIZAÇÃO DE VARIÁVEIS-------------------------------------------------------
	SET @dblValor = (0.000)
	SET @PKIdTipo = (SELECT PKId FROM tblCaracteristicaGeral  
			  WHERE intutilizacaodacaracteristica = 3 
			    AND intCodigoDaCaracteristica = 1)
	SET @PKIdTipoDET = (SELECT C.intCodigoDetalheDaCaracteristica
			      FROM tblCaracteristicaGeral A , tblDetalheDaCaracteristica B,
				   tblCaracteristicasDoImovel C, tblImobiliario D
			     WHERE A.PKId = B.intCaracteristica
			       AND A.PKId = C.intCodigoCaracteristicaGeral
			       AND B.PKId = C.intCodigoDetalheDaCaracteristica
			       AND D.PKId = C.intCodigoImobiliario
			       AND D.PKId = @PKIdImobiliario
			       AND A.intUtilizacaoDaCaracteristica = 3
			       AND intCodigoCaracteristicaGeral = @PKIdTipo)
	--FIM DA INICIALIZAÇÃO-----------------------------------------------------------

	
	DECLARE c_CAT CURSOR 
		FOR
		SELECT C.intCodigoCaracteristicaGeral, C.intCodigoDetalheDaCaracteristica
		  FROM tblCaracteristicaGeral A , tblDetalheDaCaracteristica B, 
		       tblCaracteristicasDoImovel C, tblImobiliario D
		 WHERE A.PKId = B.intCaracteristica
		   AND A.PKId = C.intCodigoCaracteristicaGeral
		   AND B.PKId = C.intCodigoDetalheDaCaracteristica
		   AND D.PKId = C.intCodigoImobiliario
		   AND D.PKId = @PKIdImobiliario
		   AND A.intUtilizacaoDaCaracteristica = 3
		   AND intCodigoCaracteristicaGeral != @PKIdTipo
	OPEN	c_CAT
	FETCH	c_CAT INTO
		@intCodCaracteristica, @intDetalheCaracteristica
	WHILE @@FETCH_STATUS = 0
	BEGIN	
		DECLARE c_CATValor CURSOR 
			FOR	
			SELECT ISNULL(dblvalor,0.000) FROM tblFatorDeCorrecao 
			 WHERE intCaracteristicaHorizontal = @PKIdTipo
			   AND intDetalheCaracteristicaHorizontal = @PKIdTipoDET
			   AND intCaracteristicaVertical = @intCodCaracteristica
			   AND intDetalheCaracteristicaVertical = @intDetalheCaracteristica
		OPEN	c_CATValor
		FETCH	c_CATValor INTO
			@dblAuxCAT
		WHILE @@FETCH_STATUS = 0
		BEGIN		
			SET @dblValor = @dblValor + @dblAuxCAT
			FETCH c_CATValor INTO
				@dblAuxCAT
		END
		CLOSE c_CATValor
		DEALLOCATE c_CATValor					
		FETCH c_CAT INTO
			@intCodCaracteristica, @intDetalheCaracteristica
	END
	CLOSE c_CAT
	DEALLOCATE c_CAT
	SET @dblValor = @dblValor/100
