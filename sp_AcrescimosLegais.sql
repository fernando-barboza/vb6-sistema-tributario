/*  PROCEDURE ACRÉSCIMOS LEGAIS - TRIBUTÁRIO
	                          	BY LUÍS HENRIQUE */

IF EXISTS(SELECT NAME FROM SYSOBJECTS
   WHERE NAME = 'sp_AcrescimosLegais' AND TYPE = 'P')
   DROP PROCEDURE sp_AcrescimosLegais

GO

	CREATE PROCEDURE sp_AcrescimosLegais (@InscricaoInicial 	AS NVARCHAR(20),
					      @InscricaoFinal   	AS NVARCHAR(20),
					      @dtmVencimento 		AS DATETIME,
					      @Selecionou		AS INT)	
AS 

       	CREATE TABLE #t_Acrescimos 
			(ValorParcela  	MONEY,
			 JurosTotal	MONEY,
			 MultaTotal	MONEY,
			 CorrecaoTotal	MONEY)
               	       

        DECLARE @strFormula       As NVARCHAR(100),
		@Juros            As MONEY,
		@JurosTotal       As MONEY,
		@Multa            As MONEY,
		@MultaTotal       As MONEY,
		@Correcao         As MONEY,
		@CorrecaoTotal    As MONEY,
		@NovoValorParcela As MONEY,
		@QuantDias        As INTEGER

--sp_AcrescimosLegais '00001', '43', '2001/10/01', 1
/*################################ BuscaDadosReceita ###########################################*/    
	IF @Selecionou <> 1 
	BEGIN
		SELECT   PR.PKId, LC.intContribuinte, LC.strInscricaoCadastral, LC.intExercicio, 
	       	         PR.dtmDataVencimento, PR.intNumeroParcela, LC.intOcorrencia, PR.dblValorParcela, LC.BytOrigem
		FROM     tblLancamentoCalculo LC, tblParcelaReceita PR
		WHERE    LC.PKId = PR.intLancamentoCalculo
	       	         AND PR.dtmDataVencimento <  @dtmvencimento
               	         AND LC.strInscricaoCadastral BETWEEN @InscricaoInicial AND @InscricaoFinal
 		ORDER BY intContribuinte
	END
	ELSE
	BEGIN
		SELECT   PR.PKId, LC.intContribuinte, LC.strInscricaoCadastral, LC.intExercicio, 
	       		 PR.dtmDataVencimento, PR.intNumeroParcela, LC.intOcorrencia, PR.dblValorParcela, LC.BytOrigem
		FROM     tblLancamentoCalculo LC, tblParcelaReceita PR
		WHERE    LC.PKId = PR.intLancamentoCalculo
 		ORDER BY intContribuinte 
	END

/*################################ BuscaDadosTaxa ###########################################*/ 
	IF @Selecionou <> 1 
	BEGIN
		SELECT   PT.PKId, LC.intContribuinte, LC.strInscricaoCadastral, LC.intExercicio, 
	       		 PT.dtmDataVencimento, PT.intNumeroParcela, LC.intOcorrencia, PT.dblValorParcela, LC.BytOrigem
		FROM     tblLancamentoCalculo LC, tblParcelaTaxa  PT 
		WHERE    LC.PKId = PT.intLancamentoCalculo
	       		 AND PT.dtmDataVencimento <  @dtmVencimento
			 AND LC.strInscricaoCadastral BETWEEN @InscricaoInicial AND @InscricaoFinal
		ORDER BY intContribuinte 
	END
	ELSE
	BEGIN
		SELECT   PT.PKId, LC.intContribuinte, LC.strInscricaoCadastral, LC.intExercicio, 
	       		 PT.dtmDataVencimento, PT.intNumeroParcela, LC.intOcorrencia, PT.dblValorParcela, LC.BytOrigem
		FROM     tblLancamentoCalculo LC, tblParcelaTaxa  PT 
		WHERE    LC.PKId = PT.intLancamentoCalculo
		ORDER BY intContribuinte 
	END
