IF EXISTS (SELECT NAME FROM SYSOBJECTS
   WHERE NAME = 'sp_ParametroReceitas' AND TYPE = 'P')
   DROP PROCEDURE sp_ParametroReceitas
GO

CREATE PROCEDURE sp_ParametroReceitas(@intReceitaFormula   AS INT,
				      @PKIdImobilEconomico AS INT,
				      @strParametro 	   AS NVARCHAR(4000) OUTPUT)
AS
	IF @intReceitaFormula = 97 OR @intReceitaFormula = 101
		OR @intReceitaFormula = 65 OR @intReceitaFormula = 67
		OR @intReceitaFormula = 68
		SET @strParametro = CONVERT(NVARCHAR(30),@PKIdImobilEconomico) 
				    + ',' + ' @dblValor OUTPUT'

	ELSE IF @intReceitaFormula = 99 OR @intReceitaFormula = 13
		SET @strParametro = CONVERT(NVARCHAR(30),@PKIdImobilEconomico) 
				    + ',' + @strParametro + ', @dblValor OUTPUT'
		

	ELSE IF @intReceitaFormula != 9 AND @intReceitaFormula != 11 
		AND @intReceitaFormula != 12 AND @intReceitaFormula != 95
		AND @intReceitaFormula != 25
		SET @strParametro = ' @dblValor OUTPUT'


/*
97	Imposto Territorial Urbano                                                                           ITU        2       1             0           NULL             0                  NULL            NULL                13                2                         2001-10-01 00:00:00.000     64
99	Imposto Sobre Serviço de Qualquer Natureza - Anual                                                   ISSQN-ANO  2       1             0           NULL             0                  NULL            11                  13                9                         2001-10-01 00:00:00.000     64
101	Imposto Predial Urbano                                                                               IPU        2       1             0           NULL             0                  NULL            10                  13                2                         2001-10-01 00:00:00.000     64
7	Imposto Sobre Serviços de Qualquer Natureza - Mensal                                                 ISSQN-MES  2       0             0           .00              0                  NULL            NULL                13                9                         2001-11-23 00:00:00.000     40
9	Imposto Sobre Serviços de Qualquer Natureza - Arbitrado                                              ISSQN-ARB  2       1             0           NULL             0                  NULL            14                  13                9                         2001-10-01 00:00:00.000     64
10	Imposto Sobre Serviços de Qualquer Natureza - Estimado                                               ISSQN-EST  2       1             0           NULL             0                  NULL            NULL                13                9                         2001-10-01 00:00:00.000     64
11	Imposto Sobre Transmissão Inter-Vivos - Urbano                                                       ITBI-URB   2       1             0           NULL             0                  NULL            NULL                13                8                         2001-10-01 00:00:00.000     64
12	Imposto Sobre Transmissão Inter-Vivos - Rural                                                        ITBI-RURAL 2       1             0           NULL             0                  NULL            NULL                13                8                         2001-10-01 00:00:00.000     64
13	Taxa de Licença Localização e Funcionamento                                                          TLLF       3       1             0           NULL             0                  NULL            NULL                13                10                        2001-10-01 00:00:00.000     64
14	Taxa de Propaganda e Publicidade                                                                     TPP        3       1             0           NULL             0                  NULL            NULL                13                10                        2001-10-01 00:00:00.000     64
15	Taxa de Expediente
*/	
--	SELECT * FROM tblReceita
