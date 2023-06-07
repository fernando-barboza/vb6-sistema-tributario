IF EXISTS (SELECT NAME FROM SYSOBJECTS
   WHERE NAME = 'sp_CalculoPKIdTipoDET' AND TYPE = 'P')
   DROP PROCEDURE sp_CalculoPKIdTipoDET
GO

-- sp_CalculoPKIdTipoDET 11,0
CREATE PROCEDURE sp_CalculoPKIdTipoDET(@PKIdImobiliario AS INT,
		               	       @PKIdTipoDET	AS INT OUTPUT)
AS
	DECLARE @PKIdTipo INT

	EXECUTE sp_CalculoPKIdTIPO @PKIdTipo OUTPUT

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