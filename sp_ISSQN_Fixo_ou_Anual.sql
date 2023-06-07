IF EXISTS (SELECT NAME FROM SYSOBJECTS
   WHERE NAME = 'sp_ISSQN_Fixo_ou_Anual' AND TYPE = 'P')
   DROP PROCEDURE sp_ISSQN_Fixo_ou_Anual
GO

-- sp_ISSQN_Fixo_ou_Anual 53, 1,0
CREATE PROCEDURE sp_ISSQN_Fixo_ou_Anual(@PKIdEconomico	   AS INT,
					@dblValorInformado AS MONEY,
					@dblValor 	   AS MONEY OUTPUT)
AS
	DECLARE @bytTipoDoValor TINYINT

	SELECT @dblValor = TV.dblValor, @bytTipoDoValor = TV.bytTipoDoValor
	  FROM tblAtividadeEC AEC, tblAtividadeDaEmpresa AE, 
	       tblEconomico E, tblAtividadeBasica AB, 
	       tblAtividade ATIV,tblTabelaDeValor TV
	 WHERE AEC.PKId = AE.intAtividade
	   AND E.PKId = AE.intEconomico
	   AND AB.PKId = E.intAtividadeBasica
	   AND AEC.PKID = ATIV.intAtividade
	   AND TV.PKId = ATIV.intTabelaDeValor
	   AND E.intAtividadeBasica = 5 --AUTÔNOMOS E AMBULANTES
	   AND ATIV.intReceita = 99 --CÓDIGO DA RECEITA DO ISSQN - FIXO - ANUAL
	   AND E.PKId = @PKIdEconomico 

	IF @bytTipoDoValor = 0	--PERCENTUAL
		SET @dblValor = @dblValorInformado * (@dblValor/100)
	ELSE 
		IF @bytTipoDoValor = 1 -- QUANTIDADE
			SET @dblValor = (ISNULL((SELECT IE.dblValor		 
					           FROM tblIndiceEconomico  IE, 
						        tblIndexadorEconomico E 
					          WHERE IE.intIndexador = E.PKId			    
						    AND IE.dtmData IN 
					        (SELECT MAX(IE.dtmData) 
 				                   FROM tblIndiceEconomico IE)),0) * @dblValor)   
		--ELSE  MOEDA
