IF EXISTS (SELECT NAME FROM SYSOBJECTS
   WHERE NAME = 'sp_CalculoAreaTotalConstruida' AND TYPE = 'P')
   DROP PROCEDURE sp_CalculoAreaTotalConstruida
GO

-- sp_CalculoAreaTotalConstruida 4,0
CREATE PROCEDURE sp_CalculoAreaTotalConstruida(@PKIdImobiliario AS INT,
		               		       @dblValor	AS MONEY OUTPUT)
AS
	SET @dblValor = (SELECT ISNULL(intMedidaDaArea,0) 
			   FROM tblAreaImobiliario AI, tblTipoDeArea TA
		          WHERE AI.intImobiliario = @PKIdImobiliario
			    AND TA.PKId = AI.intTipoDeArea
		            AND TA.intCodigoDaArea = 2) --Área Total Contruída
	