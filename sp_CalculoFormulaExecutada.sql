IF EXISTS (SELECT NAME FROM SYSOBJECTS
   WHERE NAME = 'sp_CalculoFormulaExecutada' AND TYPE = 'P')
   DROP PROCEDURE sp_CalculoFormulaExecutada
GO
-- sp_CalculoFormulaExecutada -1, NULL , '4,@dblValor OUTPUT'
CREATE PROCEDURE sp_CalculoFormulaExecutada(@PKId_Flag     AS INT,
		               		    @dblValor	   AS MONEY OUTPUT,
					    @strParametros AS NVARCHAR(4000))--UTILIZADO EM ALGUNS CASOS
AS
	DECLARE @strNomeProcedimento AS NVARCHAR(500)

	IF @PKId_Flag > 0
		SELECT @strNomeProcedimento = strNome
		  FROM tblFormulaDeCalculo
		 WHERE intReceita = @PKId_Flag
	ELSE
		SELECT @strNomeProcedimento = strNome
		  FROM tblFormulaBasica
		 WHERE BytTipoDeFormula = (@PKId_Flag*(-1))

	--CRIA TABELA TEMPORÁRIA
	CREATE TABLE #t_Temporaria(dblValor MONEY)
	--CONCATENA PROCEDIMENTO NAS VARIÁVEIS
	SET @strNomeProcedimento = 'DECLARE @dblValor AS MONEY EXECUTE ' + @strNomeProcedimento
				   + ' ' + @strParametros  + ';' + 'INSERT INTO #t_Temporaria SELECT ISNULL(@dblValor,0.00) '		

	--EXECUTA O PROCEDIMENTO COM SEUS DEVIDOS PARÂMETROS E SUAS ATRIBUIÇÕES
	EXECUTE(@strNomeProcedimento)
	SET @dblValor = (SELECT MAX(dblValor) FROM #t_Temporaria)
	DROP TABLE #t_Temporaria

