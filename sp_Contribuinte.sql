/*********************************************************/
/*Cadastro de Contribuintes 			         */
/*14/03/2002 - Érico					 */
/*********************************************************/

IF EXISTS (SELECT NAME FROM SYSOBJECTS
   WHERE NAME = 'sp_Contribuinte' AND TYPE = 'P')
   DROP PROCEDURE sp_Contribuinte
GO

CREATE PROCEDURE sp_Contribuinte

AS

	/*Cria uma tabela temporária*/

       	CREATE TABLE #t_Contribuinte
		    (PKId                    INT,
		     strNome                 NVARCHAR(100),
		     bytNaturezaJuridica     TINYINT,
		     strCNPJCPF              NVARCHAR(14),
		     blnResidenteNoMunicipio BIT,
	 	     dtmDataCadastro         DATETIME)	
   
    INSERT INTO #t_Contribuinte
         SELECT PKId,
		strNome,
	        bytNaturezaJuridica,
                strCNPJCPF,
                blnResidenteNoMunicipio,
		dtmDataCadastro
	   FROM	tblContribuinte 

       	 SELECT PKId,
		strNome,
	        bytNaturezaJuridica,
	        CASE bytNaturezaJuridica WHEN 0 THEN 'Física' WHEN 1 THEN 'Jurídica' WHEN 2 THEN 'SC' WHEN 3 THEN 'Outros' END AS strNaturezaJuridica,
                strCNPJCPF,
		blnResidenteNoMunicipio,
                CASE blnResidenteNoMunicipio WHEN 1 THEN 'Sim' WHEN 2 THEN 'Não' END AS strResidenteNoMunicipio,
		dtmDataCadastro
	   FROM #t_Contribuinte ORDER BY PKId, strNome

-- sp_Contribuinte 