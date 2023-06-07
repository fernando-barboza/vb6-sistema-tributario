IF EXISTS (SELECT NAME FROM SYSOBJECTS
   WHERE NAME = 'sp_ExtratoIndividual' AND TYPE = 'P')
   DROP PROCEDURE sp_ExtratoIndividual
GO

-- sp_ExtratoIndividual 2739,2739,'2001/01/01 00:00:00','2002/03/04 00:00:00'
CREATE PROCEDURE sp_ExtratoIndividual(@PKIdInicial AS INT,
				      @PKIdFinal   AS INT,
		               	      @dtmInicial  AS DATETIME,
				      @dtmFinal	   AS DATETIME)
AS
	CREATE TABLE #t_Relatorio
	(strNome               NVARCHAR(100),
	 strInscricaoCadastral NVARCHAR(20),
	 strCNPJCPF            NVARCHAR(50),
	 dtmData               DATETIME,
	 strLancPagamento      NVARCHAR(400),
	 bytLancPagamento      TINYINT,
	 intExercicio          INT,
	 intNumeroParcela      INT,
	 dblValor              MONEY,
	 strComposicao	       NVARCHAR(200),
	 dtmDataPagamento      DATETIME,
	 dtmDataVencimento     DATETIME)
	
	IF @PKIdInicial = 0 
	BEGIN
		SET @PKIdInicial = (SELECT PKId FROM tblContribuinte 
				     WHERE PKId = (SELECT MIN(PKId) 
						     FROM tblContribuinte))
		SET @PKIdFinal = (SELECT PKId FROM tblContribuinte 
				   WHERE PKId = (SELECT MAX(PKId) 
						   FROM tblContribuinte))
	END
	INSERT INTO #t_Relatorio
		SELECT C.strNome, B.strInscricaoCadastral, C.strCNPJCPF, B.dtmLancamento, D.strSigla,0, B.intExercicio,
		       A.intNumeroParcela, A.dblValorParcela, D.strDescricao, A.dtmDataPagamento, A.dtmDataVencimento 
		  FROM tblparcelareceita A, tbllancamentocalculo B, tblContribuinte C,
		       tblComposicaoDaReceita D
		 WHERE A.intLancamentoCalculo = B.PKId
		   AND B.intContribuinte = C.PKId
		   AND A.intComposicaoDaReceita = D.PKId
		   AND B.dtmLancamento BETWEEN @dtmInicial AND @dtmFinal
		   AND C.PKId BETWEEN @PKIdInicial AND @PKIdFinal

/*	INSERT INTO #t_Relatorio
		SELECT B.strNome, A.strInscricaoCadastral, B.strCNPJCPF, A.dtmDataLancamento, C.strSigla,1, A.intExercicio,
		       A.intNumeroDaParcela, A.dblTotalPago,  C.strDescricao, A.dtmDataPagamento, 
		       A.dtmDataVencimento
		  FROM tblpagamentoparcela A, tblContribuinte B, tblComposicaoDaReceita C
		 WHERE A.intComposicaoDaReceita = C.PKId
		   AND A.intContribuinte = B.PKId
		   AND A.dtmDataLancamento BETWEEN @dtmInicial AND @dtmFinal
		   AND B.PKId BETWEEN @PKIdInicial AND @PKIdFinal
*/
	SELECT * FROM #t_Relatorio
	ORDER BY strNome,strInscricaoCadastral,strComposicao

--SELECT * FROM tblpagamentoparcela  dtmDataVencimento
--Select * FROM tblparcelareceita    dtmDataVencimento