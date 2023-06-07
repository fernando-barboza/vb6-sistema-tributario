IF EXISTS (SELECT NAME FROM SYSOBJECTS
   WHERE NAME = 'sp_CalculoPKIdTIPO' AND TYPE = 'P')
   DROP PROCEDURE sp_CalculoPKIdTIPO
GO

-- sp_CalculoPKIdTIPO 0
CREATE PROCEDURE sp_CalculoPKIdTIPO(@PKIdTipo AS INT OUTPUT)
AS
	SET @PKIdTipo = (SELECT PKId FROM tblCaracteristicaGeral  
			  WHERE intutilizacaodacaracteristica = 3 
			    AND intCodigoDaCaracteristica = 24)