IF EXISTS (SELECT NAME FROM SYSOBJECTS
   WHERE NAME = 'sp_CalculoAreaDoTerreno' AND TYPE = 'P')
   DROP PROCEDURE sp_CalculoAreaDoTerreno
GO

-- sp_CalculoAreaDoTerreno 4,0
CREATE PROCEDURE sp_CalculoAreaDoTerreno(@PKIdImobiliario AS INT,
		               		 @dblValor	  AS MONEY OUTPUT)
AS
	SET @dblValor = (SELECT CONVERT(NUMERIC,strArea) 
			   FROM tblImobiliario
		          WHERE PKId = @PKIdImobiliario)
