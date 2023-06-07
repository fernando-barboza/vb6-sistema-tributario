IF EXISTS (SELECT NAME FROM SYSOBJECTS
   WHERE NAME = 'sp_CalculoAreaConstruida' AND TYPE = 'P')
   DROP PROCEDURE sp_CalculoAreaConstruida
GO

-- sp_CalculoAreaConstruida 4,0
CREATE PROCEDURE sp_CalculoAreaConstruida(@PKIdImobiliario AS INT,
		               		  @dblValor	   AS MONEY OUTPUT)
AS
	SET @dblValor = (SELECT ISNULL(intMedidaDaArea,0) 
			   FROM tblAreaImobiliario AI, tblTipoDeArea TA
		          WHERE AI.intImobiliario = @PKIdImobiliario
			    AND TA.PKId = AI.intTipoDeArea
		            AND TA.intCodigoDaArea = 1) --Área Construída da Unidade
