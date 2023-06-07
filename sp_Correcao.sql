IF EXISTS(SELECT NAME FROM SYSOBJECTS 
   WHERE NAME = 'sp_Correcao' AND TYPE = 'P')
   DROP PROCEDURE sp_Correcao

GO
--drop PROCEDURE sp_Correcao

   CREATE PROCEDURE sp_Correcao(@dtmDataParcela AS DateTime)
AS

   DECLARE @ValorInicial NUMERIC(12,8),
           @ValorFinal   NUMERIC(12,8)

   SET @ValorInicial = (SELECT dblValor FROM tblIndiceEconomico A, tblIndexadorEconomico B
                        WHERE B.PKId = A.intindexador
		        AND A.dtmData = @dtmDataParcela)

   SET @ValorFinal = (SELECT dblValor FROM tblIndiceEconomico A, tblIndexadorEconomico B
                      WHERE B.PKId = A.intindexador
        	      AND A.dtmData = GETDATE())



   SELECT ISNULL(@ValorInicial,0) - ISNULL(@ValorFinal,0)

--sp_Correcao '2001-01-01 00:00:00.000'
