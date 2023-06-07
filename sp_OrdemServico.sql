IF EXISTS (SELECT NAME FROM SYSOBJECTS
   WHERE NAME = 'sp_OrdemServico' AND TYPE = 'P')
   DROP PROCEDURE sp_OrdemServico
GO

CREATE PROCEDURE sp_OrdemServico(@strPKId	  AS NVARCHAR(1000) = '',
				 @strDtHorario    AS NVARCHAR(1000)= '',
				 @PKIdOPT	  AS INT , 		 
				 @txtstrRazao     AS NVARCHAR(500) = '',
				 @opt             AS TINYINT = 3,
				 @InscCadastral   AS INT = 0,
				 @glngCodUsr      AS INT = 0,
				 @strModoOperacao AS NVARCHAR(10))
AS						
		CREATE TABLE #t_OrdemServico
		(PKId		INT)

	DECLARE @PKIdAtual  	INT,
		@PKId	  	INT,
		@posicao  	INT,
		@dtmDtHorario	DATETIME,
		@strDtHora	NVARCHAR(19)

		SET @posicao = 1
		SET @strDtHora = ''

		IF @strModoOperacao = 'INCLUIR'
		BEGIN
			INSERT INTO tblOrdemServico 
			SELECT @txtstrRazao,@opt,@InscCadastral, GETDATE(), @glngCodUsr

			SET @PKIdAtual = (SELECT MAX(PKId) FROM tblOrdemServico)
		END		
		
		IF @strModoOperacao = 'ALTERAR'
		BEGIN
			UPDATE tblOrdemServico SET strRazaoDaFiscalizacao = @txtstrRazao,
							      bytOrigem = @opt,
						       dtmDtAtualizacao = GETDATE(),
							      lngCodUsr = @glngCodUsr
							     WHERE PKId = @PKIdOPT
			DELETE FROM tblOrdemServicoFiscal WHERE intOrdemServico = @PKIdOPT
			SET @PKIdAtual = @PKIdOPT
		END

		IF @strModoOperacao = 'DELETAR'
		BEGIN
			DELETE FROM tblOrdemServicoFiscal WHERE intOrdemServico = @PKIdOPT 		 
			DELETE FROM tblOrdemServico WHERE PKId = @PKIdOPT
		END

	IF @strModoOperacao = 'ALTERAR' OR @strModoOperacao = 'INCLUIR'
	BEGIN
		EXECUTE('INSERT INTO #t_OrdemServico
				SELECT PKid FROM tblFiscal
				WHERE PKId IN('+@strPKId+')')

		DECLARE	c_OrdemServico CURSOR 
			FOR
			SELECT PKId FROM #t_OrdemServico
		OPEN	c_OrdemServico
		FETCH	c_OrdemServico INTO
			@PKId
		WHILE @@FETCH_STATUS = 0
		BEGIN
			WHILE NOT ASCII(SUBSTRING(@strDtHorario, @posicao, 1)) = 44
			BEGIN
				SET @strDtHora = @strDtHora +  CHAR(ASCII(SUBSTRING(@strDtHorario, @posicao, 1)))
				SET @posicao = @posicao + 1
			END
			SET @posicao = @posicao + 1
			SET @dtmDtHorario = CONVERT(DATETIME, @strDtHora)
			PRINT @dtmDtHorario
			INSERT INTO tblOrdemServicoFiscal
			SELECT @PKIdAtual, @PKId, @dtmDtHorario, GETDATE(), @glngCodUsr
			SET @strDtHora = ''
			FETCH c_OrdemServico INTO  
				      @PKId
		END
		CLOSE c_OrdemServico
		DEALLOCATE c_OrdemServico
	END
	IF @strModoOperacao = 'IMPRIMIR'
	BEGIN
		DECLARE @strInscricaoCadastral NVARCHAR(50),
			@strNomeProprietario   NVARCHAR(100)

		CREATE TABLE #t_Proprietario
		(PKIdOS			INT,
		 strInscricaoCadastral	NVARCHAR(50),
		 strNomeProprietario	NVARCHAr(100))

		CREATE TABLE #t_RelatorioOS
		(strInscricaoCadastral	NVARCHAR(50),
		 strNomeProprietario	NVARCHAr(100),
		 strNomeFiscal		NVARCHAR(100),
		 dtmDtHorarioFiscal	DATETIME)
		
		IF @opt = 0
			INSERT INTO #t_Proprietario 
				SELECT OS.PKId, EC.strInscricaoCadastral, CC.strNome
				 FROM tblOrdemServico OS, tblEconomico EC, tblContribuinte CC
				WHERE CC.PKId = EC.intContribuinte
				  AND EC.PKId = OS.intInscricaoCadastral
				  AND OS.bytOrigem = 0
		ELSE
			IF @opt = 1
				INSERT INTO #t_Proprietario 
					SELECT OS.PKId, IU.strInscricaoAnterior, CC.strNome
					 FROM tblOrdemServico OS, tblImobiliario IU, tblContribuinte CC
					WHERE CC.PKId = IU.intContribuinte
					  AND IU.PKId = OS.intInscricaoCadastral
					  AND OS.bytOrigem = 1
			ELSE
				INSERT INTO #t_Proprietario 
					SELECT OS.PKId, IR.strInscricaoAnterior, CC.strNome
					 FROM tblOrdemServico OS, tblImobiliarioRural IR, tblContribuinte CC
					WHERE CC.PKId = IR.intContribuinte
					  AND IR.PKId = OS.intInscricaoCadastral
					  AND OS.bytOrigem = 1
		DECLARE c_RelatorioOS CURSOR 
			FOR
			SELECT PKIdOS,strInscricaoCadastral,strNomeProprietario FROM #t_Proprietario
		OPEN	c_RelatorioOS
		FETCH	c_RelatorioOS INTO
			@PKId, @strInscricaoCadastral, @strNomeProprietario
		WHILE @@FETCH_STATUS = 0
		BEGIN
			INSERT INTO #t_RelatorioOS 
				SELECT @strInscricaoCadastral, @strNomeProprietario ,
				       CC.strNome, OSF.dtmDtHorario 
				  FROM tblOrdemServicoFiscal OSF, tblFiscal FC,
				       tblContribuinte CC
				 WHERE CC.PKId = FC.intContribuinte
				   AND FC.PKId = OSF.intFiscal
				   AND OSF.intOrdemServico = @PKId

			FETCH c_RelatorioOS INTO  
				      @PKId, @strInscricaoCadastral, @strNomeProprietario
		END
		CLOSE c_RelatorioOS
		DEALLOCATE c_RelatorioOS
		SELECT * FROM #t_RelatorioOS ORDER BY strInscricaoCadastral
	END
	
-- sp_OrdemServico '1','2001/05/21 08:50:00,',0,'Testim',0,1,23,'INCLUIR'
-- sp_OrdemServico '1','2001/05/23 08:50:00,',13,'Testim',0,1,23,'ALTERAR'
-- sp_OrdemServico '1','2001/05/23 08:50:00,',13,'Testim',0,1,23,'DELETAR'
-- sp_OrdemServico '1','2001/05/23 08:50:00,',13,'Testim',1,2,23,'IMPRIMIR'