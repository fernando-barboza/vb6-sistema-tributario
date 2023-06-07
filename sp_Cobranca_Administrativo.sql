IF EXISTS (SELECT NAME FROM SYSOBJECTS
   WHERE NAME = 'sp_Cobranca_Administrativo' AND TYPE = 'P')
   DROP PROCEDURE sp_Cobranca_Administrativo
GO
-- SELECT * FROM tblParcelaReceita
-- SELECT * FROM tblDividaAtiva  
-- SELECT * FROM tblDetalheDividaAtiva
-- EXECUTE sp_Cobranca_Administrativo 2,1,'01010010015001','01010010015001',16,2001,0,2,'1','2001/01/01 00:00:00',45

CREATE PROCEDURE sp_Cobranca_Administrativo(@intFlag		  AS INT,
/* Detalhes do  @intFLag (Valor por Form)*/ @bytOrigem		  AS TINYINT,
/* 1 - Remissão de Débitos 		 */ @strInscricaoInicial  AS NVARCHAR(60),
/* 2 - Cancelamentos de Débitos 	 */ @strInscricaoFinal    AS NVARCHAR(60),
/* 3 - Prescrição de Débito		 */ @intComposicaoReceita AS INT,
/* 4 - Execução Fiscal  		 */ @intExercicio	  AS INT,
/* 5 - Cobrança Extra-Judicial           */ @intNumeroParcela	  AS INT,
					    --Opcionais em Alguns Forms--
					    @intParcelaFinal	  AS INT,
					    @strNumeroProcesso	  AS NVARCHAR(20),
					    @dtmData		  AS DATETIME,
					    -----------------------------
					    @glngCodUsr           AS INT)
AS			
	DECLARE @PKId              INT,
		@intNumeroProcesso INT,
		@intContribuinte   INT,
		@intQuantidade	   INT,
		@intRemido	   INT,
		@intAuxParcela     INT,
		@strInsert         NVARCHAR(4000),
		@strUpdate	   NVARCHAR(4000),
		@strsql		   NVARCHAR(4000)

	CREATE TABLE #t_LancamentoCalculo
		(PKId		  INT,
		 intContribuinte  INT,
		 intNumeroParcela INT)
	--*--
	IF @intFlag <> 4 AND @intFlag <> 5
	BEGIN
		SET @intNumeroProcesso = CONVERT(INT,@strNumeroProcesso)
		SET @strInsert = 'INSERT INTO #t_LancamentoCalculo
			     SELECT LC.PKId, LC.intContribuinte, 0
			       FROM tblLancamentoCalculo LC,
				    tblParcelaReceita PR
			      WHERE PR.intNumeroParcela = ' + CONVERT(NVARCHAR,@intNumeroParcela)  
		
		SET @strUpdate = ' AND DD.intNumeroParcela = '+ CONVERT(NVARCHAR,@intNumeroParcela) 
		SET @intAuxParcela = @intNumeroParcela
	END
	ELSE
	BEGIN
		SET @strInsert = 'INSERT INTO #t_LancamentoCalculo
			     SELECT LC.PKId, LC.intContribuinte, PR.intNumeroParcela
			       FROM tblLancamentoCalculo LC,
				    tblParcelaReceita PR
			      WHERE PR.intNumeroParcela BETWEEN ' + CONVERT(NVARCHAR,@intNumeroParcela) + ' 
				AND ' + CONVERT(NVARCHAR,@intParcelaFinal)  
		IF @intFlag = 4  
			SET @strInsert = @strInsert + ' AND PR.bytAtiva = 1 '
		ELSE
			SET @intNumeroProcesso = CONVERT(INT,@strNumeroProcesso)
		SET @strUpdate = ' AND DD.intNumeroParcela BETWEEN ' + CONVERT(NVARCHAR,@intNumeroParcela) + ' 
                                   AND ' + CONVERT(NVARCHAR,@intParcelaFinal) 
	END  -- UTILIZADO NO UPDATE
	SET @strsql = ('AND intDividaAtiva IN(
			SELECT DA.PKId FROM tblDividaAtiva DA,
			    tblDetalheDividaAtiva DD
		         WHERE DD.bytOrigem = '+ CONVERT(NVARCHAR,@bytOrigem) + '
			   AND DD.intComposicaoReceita = '+ CONVERT(NVARCHAR,@intComposicaoReceita) +'
			   AND DD.intExercicio = '+ CONVERT(NVARCHAR,@intExercicio) +'
			   AND DD.strInscricaoCadastral BETWEEN "'+ @strInscricaoInicial + '" AND "' + @strInscricaoFinal +'"
			   AND DA.PKId = DD.intDividaAtiva 								
			   AND DD.dtmPrescricao IS NULL
			   AND DD.dtmCancelamento IS NULL
			   AND DD.dtmAjuizamento IS NULL   
			   AND DD.bytSituacao > 0 ' + @strUpdate )

	SET @strInsert = @strInsert + ' AND LC.PKId = PR.intLancamentoCalculo
					AND LC.bytOrigem = ' + CONVERT(NVARCHAR,@bytOrigem) + '
					AND LC.intComposicaoReceita = ' + CONVERT(NVARCHAR,@intComposicaoReceita) + '
					AND LC.intExercicio = ' + CONVERT(NVARCHAR,@intExercicio) + '
					AND LC.strInscricaoCadastral BETWEEN "'+ @strInscricaoInicial + '" 
					AND "' + @strInscricaoFinal +'"
					AND PR.bytSuspensaoDeExigencia = 0
					AND LC.intContribuinte NOT IN
					(SELECT intContribuinte From tblSuspensaoDeExigencia)
					AND LC.intContribuinte NOT IN
                        	       	(SELECT intContribuinte From tblCobrancaExtraJudicial)'
	EXECUTE(@strInsert)
	DECLARE	c_Cobranca_Administrativo CURSOR FOR
		SELECT PKId, intContribuinte, intNumeroParcela FROM #t_LancamentoCalculo
	OPEN	c_Cobranca_Administrativo
	FETCH	c_Cobranca_Administrativo INTO
		@PKId, @intContribuinte, @intNumeroParcela
	WHILE @@FETCH_STATUS = 0
	BEGIN
		IF @intFlag <> 4 AND @intFlag <> 5
			SET @intNumeroParcela = @intAuxParcela
		IF @intFlag = 5
		BEGIN
			INSERT INTO tblCobrancaExtraJudicial
				SELECT (SELECT PKId FROM tblParcelaReceita 
				             WHERE intLancamentoCalculo = @PKId
					       AND intNumeroParcela = @intNumeroParcela) 
						    AS intParcelaReceita, 
	                           @intContribuinte AS Contribuinte, 
			                   @dtmData AS DataCobranca, GETDATE(),
					@glngCodUsr AS CodigoDeUsuario
			GOTO SAIFORA
		END

		SET @intQuantidade = (SELECT COUNT(DA.PKId) 
					FROM tblDividaAtiva DA,
			                     tblDetalheDividaAtiva DD
			 	       WHERE DD.bytOrigem = @bytOrigem
	        			 AND DD.intComposicaoReceita = @intComposicaoReceita
					 AND DD.intExercicio = @intExercicio
					 AND DD.intNumeroParcela = @intNumeroParcela
					 AND DD.strInscricaoCadastral BETWEEN  @strInscricaoInicial  AND  @strInscricaoFinal 
					 AND DA.PKId = DD.intDividaAtiva
					 AND DD.dtmPrescricao IS NULL
			                 AND DD.dtmCancelamento IS NULL
					 AND DD.dtmAjuizamento IS NULL
					 AND DD.bytSituacao > 0
					 AND DA.intContribuinte = @intContribuinte)
		IF @intQuantidade > 0		
		BEGIN
			DELETE FROM tblParcelaReceita 
			      WHERE intLancamentoCalculo = @PKId
				AND intNumeroParcela = @intNumeroParcela
			DELETE FROM tblParcelaTaxa 
			      WHERE intLancamentoCalculo = @PKId
				AND intNumeroParcela = @intNumeroParcela 
		 	IF @intFlag = 1
			BEGIN
				SET @intRemido = (SELECT PKId FROM tblOcorrencia WHERE bytRemido = 1)
				SET @strUpdate = 'UPDATE tblDetalheDividaAtiva SET bytSituacao = 0 ,
						                                 intOcorrencia = ' + CONVERT(NVARCHAR, @intRemido) + ',
 									      dtmDtAtualizacao = GETDATE() ,
				 					             lngCodUsr = '+ CONVERT(NVARCHAR,@glngCodUsr) + '
							                WHERE intNumeroParcela = ' + CONVERT(NVARCHAR,@intNumeroParcela) + 
						+ @strsql + 'AND DA.intContribuinte = ' + CONVERT(NVARCHAR,@intContribuinte) + ')'
			END
			ELSE IF @intFlag = 2
				SET @strUpdate = 'UPDATE tblDetalheDividaAtiva SET bytSituacao = 0,
					      		     intNumeroProcessoCancelamento  = ' + CONVERT(NVARCHAR,@intNumeroProcesso) + ',
							                   dtmCancelamento  = "' + CONVERT(NVARCHAR, @dtmData) + '",
 									   dtmDtAtualizacao = GETDATE() ,
				 					          lngCodUsr = '+ CONVERT(NVARCHAR,@glngCodUsr) + '
							             WHERE intNumeroParcela = ' + CONVERT(NVARCHAR,@intNumeroParcela) + 								
						+ @strsql + 'AND DA.intContribuinte = ' + CONVERT(NVARCHAR,@intContribuinte) + ')'
			ELSE IF @intFlag = 3
				SET @strUpdate = 'UPDATE tblDetalheDividaAtiva SET bytSituacao = 0,
					      		           intNumeroProcessoPrescricao = ' + CONVERT(NVARCHAR,@intNumeroProcesso) + ',
							                         dtmPrescricao = "' + CONVERT(NVARCHAR, @dtmData) + '",
 									      dtmDtAtualizacao = GETDATE() ,
				 					             lngCodUsr = '+ CONVERT(NVARCHAR,@glngCodUsr) + '
							                WHERE intNumeroParcela = ' + CONVERT(NVARCHAR,@intNumeroParcela) +
					     	+ @strsql + 'AND DA.intContribuinte = ' + CONVERT(NVARCHAR,@intContribuinte) + ')'
			ELSE IF @intFlag = 4 --*--
				BEGIN
				SET @strUpdate = 'UPDATE tblDetalheDividaAtiva SET bytSituacao = 0,
					      		          strNumeroProcessoAjuizamento = "' + @strNumeroProcesso + '",
							                        dtmAjuizamento = "' + CONVERT(NVARCHAR, @dtmData) + '",
 									      dtmDtAtualizacao = GETDATE() ,
				 					             lngCodUsr = '+ CONVERT(NVARCHAR,@glngCodUsr) + '
							                WHERE intNumeroParcela = ' + CONVERT(NVARCHAR,@intNumeroParcela) +
						+ @strsql + 'AND DA.intContribuinte = ' + CONVERT(NVARCHAR,@intContribuinte) + ')'
				END
			EXECUTE (@strUpdate)
			SELECT @strUpdate
		END
		SAIFORA:
		FETCH c_Cobranca_Administrativo INTO  
			      @PKId, @intContribuinte, @intNumeroParcela
	END
	CLOSE c_Cobranca_Administrativo
	DEALLOCATE c_Cobranca_Administrativo


