IF EXISTS (SELECT NAME FROM SYSOBJECTS
   WHERE NAME = 'sp_CEV' AND TYPE = 'P')
   DROP PROCEDURE sp_CEV
GO
-- sp_CEV 3,1,2,1, @dblValor OUTPUT
CREATE PROCEDURE sp_CEV(@PKIdEdital 	  AS INT,
			@PKIdSecao  	  AS INT,
			@NReceitas  	  AS INT,
			@intContribuintes AS INT,
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