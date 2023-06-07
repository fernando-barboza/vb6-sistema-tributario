VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptSubFichaImobiliarioPredio 
   Caption         =   "rptSubFichaImobiliarioPredio (ActiveReport)"
   ClientHeight    =   10815
   ClientLeft      =   -15
   ClientTop       =   360
   ClientWidth     =   15360
   MDIChild        =   -1  'True
   _ExtentX        =   27093
   _ExtentY        =   19076
   SectionData     =   "rptSubFichaImobiliarioPredio.dsx":0000
End
Attribute VB_Name = "rptSubFichaImobiliarioPredio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Detail_Format()
    LinhaD.Y2 = Detail.Height
    LinhaE.Y2 = Detail.Height
    txt_dblValor = gstrConvVrDoSql(txt_dblValor, 1)
    
End Sub

Private Sub PreenchePadraoPredio(lngPkidPredio As Long)
    
    Dim strSQL As String
    Dim adoRec As New ADODB.Recordset

    strSQL = "SELECT "
    strSQL = strSQL & "FP.Strdescricao as DescricaoFaixaPontos, "
    strSQL = strSQL & "PadraoPontuacao.Pkid, "
    strSQL = strSQL & "PadraoPontuacao.Valor "
    strSQL = strSQL & "From " & gstrCategoriaConstrucao & " CC, "
    strSQL = strSQL & gstrFaixaPontosPredio & " FP, "
    strSQL = strSQL & gstrExercicioValorM2Predio & " EVM2, "
    strSQL = strSQL & gstrMoedas & " ME,"
    strSQL = strSQL & "(SELECT SUM(TV.DBLVALOR) As Valor, AI.Pkid PKID "
    strSQL = strSQL & "FROM " & gstrCaracteristicaGeral & " CG, "
    strSQL = strSQL & gstrImobiliario & " IU, "
    strSQL = strSQL & gstrDetalheDaCaracteristica & " DC, "
    strSQL = strSQL & gstrUtilizacaoDaTabelaDeValor & " UTV, "
    strSQL = strSQL & gstrCaracteristicaDoImovel & " CI, "
    strSQL = strSQL & gstrTabelaDeValor & " TV, "
    strSQL = strSQL & gstrAreaImobiliario & " AI, "
    strSQL = strSQL & gstrCategoriaConstrucao & " CC "
    strSQL = strSQL & "WHERE CG.Pkid  " & strOUTJOracle & " = CI.Intcodigocaracteristicageral "
    strSQL = strSQL & " AND IU.PKId  = AI.intImobiliario " 'IU.PKId  = CI.intCodigoImobiliario  Alterado Rafael 21/10/04
    strSQL = strSQL & " AND DC.pkid " & strOUTJOracle & " = CI.Intcodigodetalhedacaracteristi "
    strSQL = strSQL & " AND UTV.PKId " & strOUTJOracle & " = CG.intUtilizacaoDaCaracteristica " 'UTV.PKId (+) = CI.intCodigoUtilizacaoDaTabelaDeV Alterado Rafael 21/10/04
    strSQL = strSQL & " AND TV.Pkid " & strOUTJOracle & " = DC.Inttabeladevalores  "
    strSQL = strSQL & " AND CG.intUtilizacaoDaCaracteristica = 3 "
    strSQL = strSQL & " AND CI.Intarea = AI.pkid "
    strSQL = strSQL & " AND IU.Pkid = " & rptFichaCadastroImobiliario.txt_IDImob.Text
    strSQL = strSQL & " AND AI.intImobiliario = IU.Pkid "
    strSQL = strSQL & " AND CC.Pkid " & strOUTJOracle & " = AI.intCategoriaConstrucao "
    strSQL = strSQL & " Group by AI.Pkid ) PadraoPontuacao "
    strSQL = strSQL & " Where CC.Pkid = FP.intCategoriaConstrucao  "
    strSQL = strSQL & " And FP.Pkid = EVM2.Intfaixapontospredio "
    strSQL = strSQL & " And Me.Pkid = EVM2.Intmoeda "
    strSQL = strSQL & " And CC.Pkid = " & txt_Pkid_Construcao.Text
    strSQL = strSQL & " And EVM2.intExercicio = 2004"
'    strSql = strSql & " AND PadraoPontuacao.pkid = " &
    strSQL = strSQL & " And PadraoPontuacao.VALOR BETWEEN FP.Dblpontoinicial AND FP.Dblpontofinal"
    strSQL = strSQL & " And PadraoPontuacao.Pkid = " & lngPkidPredio
    
    Set gobjBanco = New clsBanco
    
    If gobjBanco.CriaADO(strSQL, 5, adoRec) Then
        With adoRec
            If Not .EOF Then
                txt_strPadrao.DataField = ""
                txt_strPadrao.Text = !DescricaoFaixaPontos
            End If
        End With
    End If

End Sub

Private Sub GroupFooter1_Format()
txt_dblTotal = Trim(txt_dblTotal.Text)
End Sub

Private Sub GroupHeader1_Format()
    PreenchePadraoPredio adoDataControl.Recordset("Pkid").Value
End Sub
