IF EXISTS (SELECT NAME FROM SYSOBJECTS
   WHERE NAME = 'sp_TLLF' AND TYPE = 'P')
   DROP PROCEDURE sp_TLLF
GO
-- sp_CalculoLancamentoReceitas 3, -1 , -1,'13,15,99',' SELECT EC.intContribuinte, EC.strInscricaoCadastral  FROM tblEconomico AS EC,  tblTributoEmpresa AS TE,  tblAtividadeDaEmpresa AS AE,  tblAtividadeEC AS AEC  WHERE EC.PKId = TE.intEconomico  AND EC.PKID = AE.intEconomico AND EC.intAtividadeBasica = 5 AND AE.intEconomico = EC.PKId  AND AEC.PKId = AE.intAtividade  AND AE.blnPrincipal = 1  AND EC.PKId Between "12" AND "12" AND dtmDataBaixa IS NULL  AND TE.intTributo = 3 ',2001,3, NULL, '2001/01/01 00:00:00','2001/01/31 00:00:00',0,3,0,2,7,2,45,0,0,0,'10000.00'
-- sp_TLLF 53, 1,0
CREATE PROCEDURE sp_TLLF(@PKIdEconomico	    AS INT,
			 @dblValorInformado AS MONEY,
			 @dblValor 	    AS MONEY OUTPUT)
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
	   AND ATIV.intReceita = 13 --CÓDIGO DA RECEITA DA TLLF
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