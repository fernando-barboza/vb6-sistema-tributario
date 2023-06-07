IF EXISTS (SELECT NAME FROM SYSOBJECTS
   WHERE NAME = 'sp_VerificaIsencao' AND TYPE = 'P')
   DROP PROCEDURE sp_VerificaIsencao
GO
-- sp_VerificaIsencao 0,'01020010012001',16,15,1000
CREATE PROCEDURE sp_VerificaIsencao(@bytUtilizacao	   AS TINYINT,
				    @strInscricaoCadastral AS NVARCHAR(100),
				    @intComposicaoReceita  AS INT,
  				    @intReceita	    	   AS INT,
				    @dblValor 	    	   AS MONEY OUTPUT)

AS
	SET @dblValor = @dblValor - (@dblValor * ISNULL((SELECT intAliquota FROM tblIsencaoImunidade
					                  WHERE bitTipoDeInscricao = @bytUtilizacao
							    AND intReceita = @intReceita
						            AND intComposicaoDaReceita = @intComposicaoReceita
						            AND strInscricao = @strInscricaoCadastral
						            AND GETDATE() BETWEEN dtmInicial 
						            AND dtmFinal),0)/100)
