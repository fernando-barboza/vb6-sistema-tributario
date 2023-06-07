IF EXISTS (SELECT NAME FROM SYSOBJECTS
   WHERE NAME = 'sp_RemoveLancamentos' AND TYPE = 'P')
   DROP PROCEDURE sp_RemoveLancamentos
GO		
CREATE PROCEDURE sp_RemoveLancamentos(@strQueryCursor AS NVARCHAR(4000))
AS
	DECLARE @PKId INT

	EXECUTE('DECLARE c_RemoveLancamentos CURSOR FOR ' + @strQueryCursor)

	OPEN	c_RemoveLancamentos
	FETCH	c_RemoveLancamentos INTO
		@PKId
	WHILE @@FETCH_STATUS = 0
	BEGIN	
		DELETE FROM tblParcelaReceita
		WHERE intLancamentoCalculo = @PKId
		DELETE FROM tblParcelaTaxa
		WHERE intLancamentoCalculo = @PKId
		DELETE FROM tblLancamentoCalculo
		WHERE PKId = @PKId
		FETCH c_RemoveLancamentos INTO
			@PKId
	END
	CLOSE c_RemoveLancamentos
	DEALLOCATE c_RemoveLancamentos