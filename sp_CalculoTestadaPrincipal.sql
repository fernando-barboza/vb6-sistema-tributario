IF EXISTS (SELECT NAME FROM SYSOBJECTS
   WHERE NAME = 'sp_CalculoTestadaPrincipal' AND TYPE = 'P')
   DROP PROCEDURE sp_CalculoTestadaPrincipal
GO

-- sp_CalculoTestadaPrincipal 4,0
CREATE PROCEDURE sp_CalculoTestadaPrincipal(@PKIdImobiliario AS INT,
		               		    @dblValor	     AS MONEY OUTPUT)
AS
	SET @dblValor = (SELECT CONVERT(MONEY,REPLACE(ISNULL(strMedidaDaTestada,0),',','.')) 
   			   FROM tblTestadaImobiliario TI, tblTipoDeTestada TT
		          WHERE TI.intImobiliario = @PKIdImobiliario
			    AND TT.PKId = TI.intTipoDeTestada 
			    AND TT.intCodigoDaTestada = 1)--Testada Principal*/

