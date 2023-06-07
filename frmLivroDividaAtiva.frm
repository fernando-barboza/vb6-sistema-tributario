VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Begin VB.Form frmLivroDividaAtiva 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Livro de Inscrição em Divida Ativa"
   ClientHeight    =   2190
   ClientLeft      =   1290
   ClientTop       =   1965
   ClientWidth     =   3765
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2190
   ScaleWidth      =   3765
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab SSTab1 
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   3413
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Livro de Inscrição em Divida Ativa"
      TabPicture(0)   =   "frmLivroDividaAtiva.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fra_DividaAtiva"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin VB.Frame fra_DividaAtiva 
         Height          =   1215
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   3015
         Begin MSDataListLib.DataCombo dbcintLivro 
            Height          =   315
            Left            =   1560
            TabIndex        =   3
            Top             =   540
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Número do Livro:"
            Height          =   195
            Left            =   240
            TabIndex        =   2
            Top             =   600
            Width           =   1215
         End
      End
   End
End
Attribute VB_Name = "frmLivroDividaAtiva"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Sub MantemForm(ByVal strModoOperacao As String)
    Select Case strModoOperacao
        Case gstrPreencherLista
            dbcintLivro.Tag = "SELECT intLivro, intLivro FROM " & gstrDativa & " GROUP BY intLivro" & ";intLivro"
            PreencherListaDeOpcoes dbcintLivro
            'LeDaTabelaParaObj gstrDativa, dbcintLivro, "SELECT intLivro, intLivro FROM " & gstrDativa & " GROUP BY intLivro "
        Case gstrImprimir
            If blnDadosOK = True Then
               ImprimeRelatorio rptLivroDividaAtiva, strQueryRelatorio, "Livro de Inscrição em Dívida Ativa."
            End If
    End Select
End Sub

Private Function strQueryRelatorio2() As String
Dim strSql As String

    strSql = "SELECT DISTINCT DA.Intlivro, DA.Intfolha,LA.Intexercicio, LA.Strinscricao, DA.Dtmdtinscricao, " & _
             "LA.Strcomposicaodareceita, DA.Intcertidao ,LA.Strnumeroaviso, " & _
             "DA.strnomeproprietario, DA.strlogradouro, DA.strnumero, DA.strcomplemento, DA.strBairro , DA.strMunicipio, " & _
             "DA.strUF, DA.intcep, DAP.dtmdtVencimento, DAP.intTotalParcelas, " & _
             "LA.strLogradouro , LA.strnumero, LA.strBairro, LA.strComplemento, "
    strSql = strSql & gstrCASEWHEN("CR.intUtilizacao", "1,la.strnomeproprietario, 2, LE.Strnomefantasia") & " strProprietarioNomeFantasia, " & _
             gstrCASEWHEN("CR.intUtilizacao", "1,la.strpromissario , 2, EA.strDescricaoAtividade") & " strPromissarioAtividade "
    strSql = strSql & "FROM " & gstrDativa & " DA, " & _
                                gstrLancamentoAlfa & " LA, " & _
                                gstrDaParcel & " DP, " & _
                                gstrComposicaoDaReceita & " CR, " & _
                                gstrLancamentoEconomico & " LE, " & _
                                "(SELECT LEA.INTLANCAMENTOECONOMICO, LEA.strDescricaoAtividade " & _
                                "FROM " & gstrLctEconomicoAtividade & " LEA " & _
                                "WHERE blnPrincipal = 1) EA, " & _
                                "(SELECT DAP.Intdativa , " & _
                                         "COUNT(DAP.Intparcela) intTotalParcelas, " & _
                                         "MIN(DAP.Dtmdtvencimento) dtmdtVencimento " & _
                                "FROM " & gstrDaParcel & " DAP, " & _
                                          gstrDativa & " DAT " & _
                                "WHERE DAP.Intdativa = DAT.Pkid " & _
                                       "AND DAT.Intcertidao IS NOT NULL " & _
                                       "GROUP BY DAP.Intdativa) DAP "
    strSql = strSql & "Where DA.Intlancamentoalfa = LA.Pkid " & _
                      " AND DA.Intlivro = " & dbcintLivro & _
                      " AND DAP.intDativa = DA.pkid " & _
                      " AND DP.intDativa = DA.pkid " & _
                      " AND DA.intFolha = 1 " & _
                      " AND LA.Intcomposicaodareceita = CR.Pkid " & _
                      " AND LE.Intlancamentoalfa (+) =  LA.pkid " & _
                      " AND EA.Intlancamentoeconomico (+) = LE.Pkid "
    strSql = strSql & " ORDER BY DA.intFolha, DA.intCertidao "
                      
    strQueryRelatorio2 = strSql
End Function


Private Function strQueryRelatorio() As String
Dim strSql As String

    'Alterado em Pendencia tri0809
    strSql = "SELECT DISTINCT LA.Intutilizacao, DA.Intlivro, DA.Intfolha, LA.Intexercicio, " & _
             "LA.strInscricao strInscricao, " & _
             "DA.Dtmdtinscricao, LA.Strcomposicaodareceita, DA.Intcertidao, " & _
             gstrCONVERT(CDT_numeric, "LA.strNumeroAviso") & " strNumeroAviso, " & _
             "DA.strnomeproprietario, DA.strlogradouro, DA.strnumero, DA.strcomplemento, DA.strBairro , DA.strMunicipio, " & _
             "DA.strUF, DA.intcep, DAP.dtmdtVencimento, DAP.intTotalParcelas, LA.strLogradouro , LA.strnumero, LA.strBairro, LA.strComplemento, DA.Dblvalorimposto, DA.Dblvalortaxas, (" & gstrISNULL("DA.Dblvalorimposto", "0") & " + " & gstrISNULL("DA.Dblvalortaxas", "0") & ") as dblImpTaxas, " & _
             gstrISNULL("DAP.Dblcorrecaomonet", "0") & " Dblcorrecaomonet, " & gstrISNULL("DAP.Dblmulta", "0") & " Dblmulta, " & gstrISNULL("DAP.dbljuros", "0") & " dbljuros, " & gstrISNULL("DAP.Dblvalor", "0") & " Dblvalor, "
    
    strSql = strSql & "CASE "
    strSql = strSql & "WHEN CR.intUtilizacao IN (1, 4, 5) THEN "
    strSql = strSql & "LA.strNomeProprietario "
    strSql = strSql & "WHEN CR.intUtilizacao = 2 THEN "
    strSql = strSql & "CASE "
    strSql = strSql & "WHEN LE.StrNomeFantasia IS NULL OR LE.StrNomeFantasia = '' THEN "
    strSql = strSql & "LA.strNomeProprietario "
    strSql = strSql & "ELSE "
    strSql = strSql & "LE.StrNomeFantasia "
    strSql = strSql & "END END strProprietarioNomeFantasia, "
    
    strSql = strSql & gstrCASEWHEN("CR.intUtilizacao", "1,la.strpromissario , 2, ''") & " strPromissarioAtividade "
    strSql = strSql & "FROM (" & _
                      "SELECT  DAP.Intdativa, " & _
                              "COUNT(DAP.Intparcela) intTotalParcelas, " & _
                              "MIN(DAP.Dtmdtvencimento) dtmdtVencimento, " & _
                              "SUM(dap.dblmulta) dblmulta, " & _
                              "SUM(dap.dbljuros) dbljuros, " & _
                              "SUM(dap.dblcorrecaomonet) dblcorrecaomonet, " & _
                              "SUM(dap.Dblvalor) DblValor " & _
                      "FROM " & gstrDaParcel & " DAP, " & _
                             gstrDativa & " DAT " & _
                      "Where DAP.intDativa = DAT.Pkid " & _
                            "AND DAT.Intcertidao IS NOT NULL " & _
                      "GROUP BY DAP.Intdativa" & _
                            ") DAP "
    strSql = strSql & "RIGHT JOIN " & gstrDativa & " DA " & _
                      "ON DAP.intDativa = DA.pkid "
    strSql = strSql & "INNER JOIN " & gstrLancamentoAlfa & " LA " & _
                      "ON DA.Intlancamentoalfa = LA.Pkid "
    strSql = strSql & "INNER JOIN " & gstrComposicaoDaReceita & " CR " & _
                      "ON LA.Intcomposicaodareceita = CR.Pkid "
    strSql = strSql & "LEFT JOIN " & gstrLancamentoEconomico & " LE " & _
                      "ON LA.Pkid = LE.intlancamentoAlfa "
    strSql = strSql & "LEFT JOIN " & gstrLctEconomicoAtividade & " EA " & _
                      "ON LE.Pkid = EA.Intlancamentoeconomico "
    strSql = strSql & " WHERE  DA.intLivro = " & dbcintLivro.Text & " AND (EA.blnPrincipal = 1 or LE.pkid is null or EA.pkid is null)"
                            
                    
    strSql = strSql & " ORDER BY intFolha, intCertidao "
    
    strQueryRelatorio = strSql

End Function

Private Function blnDadosOK() As Boolean
  If dbcintLivro.MatchedWithList = False Then
     ExibeMensagem "A inscrição deve ser informada."
     dbcintLivro.SetFocus
     Exit Function
  End If
  
  blnDadosOK = True
End Function

