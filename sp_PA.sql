IF EXISTS (SELECT NAME FROM SYSOBJECTS
   WHERE NAME = 'sp_PA' AND TYPE = 'P')
   DROP PROCEDURE sp_PA
GO
-- sp_PA 3,1,2,1,NULL

-- sp_CalculoParaUsuario '01010010015001','15,25,95',1,20,0,0,0,'3,1,2,1, @dblValor OUTPUT'

CREATE PROCEDURE sp_PA(@PKIdEdital 	  AS INT,
		       @PKIdSecao  	  AS INT,
		       @NReceitas  	  AS INT,
		       @intContribuintes  AS INT,
		       @dblValor   	  AS MONEY OUTPUT)
AS

	IF @intContribuintes  = 0	
		SELECT @intContribuintes = COUNT(PKId) FROM tblImobiliario
        	 WHERE intSecoes = @PKIdSecao

	IF @intContribuintes  != 0
		SET @dblValor = (SELECT (dblCustoDaParcela + dblCustoDeTerceiros) AS CUSTO 
				   FROM tblTabelaDeEdital
				  WHERE PKId = @PKIdEdital)/(@intContribuintes * @NReceitas)
	ELSE
		SET @dblValor = (0.000)