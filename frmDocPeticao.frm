VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmDocPeticao 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Petição"
   ClientHeight    =   2160
   ClientLeft      =   4170
   ClientTop       =   4530
   ClientWidth     =   4995
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2160
   ScaleWidth      =   4995
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab tab_3dPasta 
      Height          =   2085
      Left            =   60
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   45
      Width           =   4860
      _ExtentX        =   8573
      _ExtentY        =   3678
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Petição"
      TabPicture(0)   =   "frmDocPeticao.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lbl_Lote"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "txtintLote"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txtintVia"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "fra_Sequencia"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      Begin VB.Frame fra_Sequencia 
         Caption         =   "Seqüência"
         Height          =   765
         Left            =   180
         TabIndex        =   5
         Top             =   1020
         Width           =   4485
         Begin MSDataListLib.DataCombo dbcFinal 
            Height          =   315
            Left            =   2790
            TabIndex        =   9
            Top             =   300
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   556
            _Version        =   393216
            IntegralHeight  =   0   'False
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dbcInicial 
            Height          =   315
            Left            =   720
            TabIndex        =   8
            Top             =   300
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   556
            _Version        =   393216
            IntegralHeight  =   0   'False
            Text            =   ""
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Final"
            Height          =   195
            Left            =   2400
            TabIndex        =   7
            Top             =   360
            Width           =   330
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Inicial"
            Height          =   195
            Left            =   240
            TabIndex        =   6
            Top             =   360
            Width           =   405
         End
      End
      Begin VB.TextBox txtintVia 
         Height          =   285
         Left            =   3015
         MaxLength       =   4
         TabIndex        =   3
         Top             =   630
         Width           =   1065
      End
      Begin VB.TextBox txtintLote 
         Height          =   285
         Left            =   915
         MaxLength       =   10
         TabIndex        =   1
         Top             =   630
         Width           =   1065
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Nº de Vias"
         Height          =   195
         Left            =   2190
         TabIndex        =   4
         Top             =   690
         Width           =   750
      End
      Begin VB.Label lbl_Lote 
         AutoSize        =   -1  'True
         Caption         =   "Lote"
         Height          =   195
         Left            =   450
         TabIndex        =   2
         Top             =   690
         Width           =   315
      End
   End
End
Attribute VB_Name = "frmDocPeticao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public strOpcao As String


Private Sub dbcFinal_Click(Area As Integer)
    DropDownDataCombo dbcFinal, Me, Area
End Sub

Private Sub dbcFinal_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbcFinal, Me, , KeyCode, Shift
End Sub

Private Sub dbcInicial_Click(Area As Integer)
    DropDownDataCombo dbcInicial, Me, Area
    Dim objRecordsetAux As New Recordset
    If dbcInicial.MatchedWithList Then
        If dbcFinal.Text = "" Or Val(dbcInicial.Text) > Val(dbcFinal.Text) Then
            

'            dbcFinal.Refresh
            
            
            dbcFinal.Text = dbcInicial.Text
            PreencherListaDeOpcoes dbcFinal, dbcInicial.BoundText
            dbcFinal.BoundText = dbcInicial.BoundText
            'dbcFinal.MatchEntry = dblBasicMatching
        End If
    End If
    
End Sub

Private Sub dbcInicial_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbcInicial, Me, , KeyCode, Shift
End Sub

Private Sub Form_Activate()
    gintCodSeguranca = 1389
    HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrImprimir
End Sub

Public Sub MantemForm(ByVal strModoOperacao As String)
        
    Select Case UCase(strModoOperacao)
        Case Is = UCase(gstrImprimir)
            If blnDadosOK Then
                
                If strOpcao = "PET" Then
                    ImprimeRelatorio rptPeticao, strQuery
                ElseIf strOpcao = "CDA" Then
                    ImprimeRelatorio rptCertidaoDativa, strQueryCertidaoDividaAtiva
                End If
                    
            End If
        Case Is = UCase(gstrNovo)
            LimpaObjeto Me
        Case Is = UCase(gstrPreencherLista)
            ToolBarGeral strModoOperacao, gstrExecutivo, False, , Me
    End Select
    
    
                    
End Sub

Private Function blnDadosOK() As Boolean
Dim stpTemplate As String
Dim stpTemplate1 As String
Dim stpDocumentPath As String
Dim stpTemplatePath As String
Dim objFileSystem  As Object
Dim strNomeDocumento As String
    
    blnDadosOK = False
    
    If Len(Trim(txtintLote)) = 0 Then
        ExibeMensagem "Informe o Lote."
        txtintLote.SetFocus
        Exit Function
    End If
    
    If Not IsNumeric(txtintLote) Then
        ExibeMensagem "Lote inválido."
        txtintLote.SetFocus
        Exit Function
    End If
    
    If Len(Trim(txtintVia)) = 0 Then
        ExibeMensagem "Informe o Número de Vias."
        txtintVia.SetFocus
        Exit Function
    End If
    
    If Not IsNumeric(txtintVia) Then
        ExibeMensagem "Número de Vias inválido."
        txtintVia.SetFocus
        Exit Function
    End If
    
    If txtintVia < 1 Then
        ExibeMensagem "Número de Vias inválido."
        txtintVia.SetFocus
        Exit Function
    End If
    
    If Not dbcInicial.MatchedWithList Then
        ExibeMensagem "Informe o Seqüencial Inicial."
        dbcInicial.SetFocus
        Exit Function
    End If
    
    If Not IsNumeric(dbcInicial.Text) Then
        ExibeMensagem "Seqüencial Inicial inválido."
        dbcInicial.SetFocus
        Exit Function
    End If
    
    If Not dbcFinal.MatchedWithList Then
        ExibeMensagem "Informe o Seqüencial Final."
        dbcFinal.SetFocus
        Exit Function
    End If
    
    If Not IsNumeric(dbcFinal.Text) Then
        ExibeMensagem "Seqüencial Final inválido."
        dbcFinal.SetFocus
        Exit Function
    End If
    
    If Val(dbcFinal.Text) < Val(dbcInicial.Text) Then
        ExibeMensagem "Seqüencial Final precisa ser igual ou superior ao Seqüencial Inicial"
        dbcFinal.SetFocus
        Exit Function
    End If
    
    If strOpcao = "PET" Then
        strNomeDocumento = "Peticao"
    ElseIf strOpcao = "CDA" Then
        strNomeDocumento = "CertidaoDA"
    End If
    stpTemplate = gstrDirDocumentos & "Documentos\" & App.ProductName & "\WordModelos\" & strNomeDocumento & ".rtf"
    stpTemplate1 = gstrDirDocumentos & "Documentos\" & App.ProductName & "\WordModelos\" & strNomeDocumento & "Fim.rtf"
    stpTemplatePath = gstrDirDocumentos & "Documentos\" & App.ProductName & "\WordModelos\"
    stpDocumentPath = gstrDirDocumentos & "Documentos\" & App.ProductName & "\WordGravados\"
    
    Set objFileSystem = New Scripting.FileSystemObject
    
    If objFileSystem.FolderExists(stpTemplatePath) Then
        If objFileSystem.FileExists(stpTemplate) Then
            If objFileSystem.FileExists(stpTemplate1) Then
                If objFileSystem.FolderExists(stpDocumentPath) Then
                    blnDadosOK = True
                Else
                    ExibeMensagem "Pasta não encontrada: " & stpDocumentPath
                End If
            Else
                ExibeMensagem "Arquivo não encontrado: " & stpTemplate1
            End If
        Else
            ExibeMensagem "Arquivo não encontrado: " & stpTemplate
        End If
    Else
        ExibeMensagem "Pasta não encontrada: " & stpTemplatePath
    End If
    
    

End Function

Private Function strQuery() As String
Dim intCont As Integer
Dim strSQL As String
Dim strVias As String

    strVias = ""
    
    For intCont = 1 To CInt(txtintVia)
        strVias = strVias & " SELECT " & intCont & " X "
        strVias = strVias & IIf(bytDBType = Oracle, " FROM DUAL ", " ")
        strVias = strVias & " UNION ALL "
    Next
    If strVias <> "" Then
        strVias = " (" & Mid(strVias, 1, Len(strVias) - 10) & ") Vias "
    End If


    strSQL = ""
    strSQL = strSQL & "SELECT "
    strSQL = strSQL & "EX.pkId " & strCONCAT & " '-' " & strCONCAT & " X Grupo, "
    strSQL = strSQL & "DA.pkId Dativa, "
    strSQL = strSQL & "EX.pkId, "
    strSQL = strSQL & "EX.strExecutadoNome, "
    strSQL = strSQL & "EX.strExecutadoTPLogNotif, "
    strSQL = strSQL & "EX.strExecutadoTITLogNotif, "
    strSQL = strSQL & "EX.strExecutadoNomeLogNotif, "
    strSQL = strSQL & "EX.strExecutadoNumLogNotif, "
    strSQL = strSQL & "EX.strExecutadoComplNotif, "
    strSQL = strSQL & "EX.strExecutadoBairroNotif,   "
    strSQL = strSQL & "EX.strExecutadoCidNotif, "
    strSQL = strSQL & "EX.intExecutadoCepNotif, "
    strSQL = strSQL & "LA.strInscricao, "
    strSQL = strSQL & "CR.strSigla intComposicaoDaReceita, "
    strSQL = strSQL & "LA.intExercicio, "
    strSQL = strSQL & gstrCONVERT(cdt_numeric, gstrISNULL("LA.strNumeroAviso", "0")) & " strNumeroAviso, "
    strSQL = strSQL & "EX.dtmDtCalculoPeticao, "
    strSQL = strSQL & "EX.strNumDistribuidor, "
    strSQL = strSQL & "EX.strIndexadorDESCR, "
    strSQL = strSQL & "EX.dblQuantIndexador, "
    strSQL = strSQL & "EX.dblVLIndexador, "
    strSQL = strSQL & "LA.intUtilizacao, "
    strSQL = strSQL & gstrCONVERT(CDT_NVARCHAR, "EX.strNumDistribuidor") & strCONCAT & " '/' " & strCONCAT & gstrCONVERT(CDT_NVARCHAR, "EX.intNumSeq") & " Controle, "
    strSQL = strSQL & "SUM(EP.dblVLPrincipal) dblVLTOTPrincipal, "
    strSQL = strSQL & "SUM(EP.dblVLJuros) dblVLTOTJuros, "
    strSQL = strSQL & "SUM(EP.dblVLMulta) dblVLTOTMulta, "
    strSQL = strSQL & "SUM(EP.dblVLTotal) dblVLTOTTotal "
    strSQL = strSQL & "FROM "
    strSQL = strSQL & "tblExecutivo EX, "
    strSQL = strSQL & "tblDativa DA, "
    strSQL = strSQL & "tblLancamentoAlfa LA, "
    strSQL = strSQL & "tblComposicaoDaReceita CR, "
    strSQL = strSQL & "tblExecutivoParcela EP, "
    strSQL = strSQL & strVias
    strSQL = strSQL & "WHERE "
    strSQL = strSQL & "DA.intLancamentoAlfa = LA.pkId "
    strSQL = strSQL & "AND DA.intExecutivo = EX.pkId "
    strSQL = strSQL & "AND EP.intDAtiva " & strOUTJOracle & " =" & strOUTJSQLServer & " DA.pkId "
    strSQL = strSQL & "AND CR.pkId " & strOUTJOracle & " =" & strOUTJSQLServer & "  LA.intComposicaoDaReceita "
    strSQL = strSQL & "AND EX.intLoteExecutivo = " & txtintLote
    strSQL = strSQL & "AND EX.intNumSeq Between " & dbcInicial.Text & " AND " & dbcFinal.Text & " "
    strSQL = strSQL & "GROUP BY "
    strSQL = strSQL & "X, "
    strSQL = strSQL & "DA.pkId, "
    strSQL = strSQL & "EX.pkId, "
    strSQL = strSQL & "EX.strExecutadoNome, "
    strSQL = strSQL & "EX.strExecutadoTPLogNotif, "
    strSQL = strSQL & "EX.strExecutadoTITLogNotif, "
    strSQL = strSQL & "EX.intNumSeq, "
    strSQL = strSQL & "EX.strExecutadoNomeLogNotif, "
    strSQL = strSQL & "EX.strExecutadoNumLogNotif, "
    strSQL = strSQL & "EX.strExecutadoComplNotif, "
    strSQL = strSQL & "EX.strExecutadoBairroNotif,   "
    strSQL = strSQL & "EX.strExecutadoCidNotif, "
    strSQL = strSQL & "EX.intExecutadoCepNotif, "
    strSQL = strSQL & "LA.strInscricao, "
    strSQL = strSQL & "CR.strSigla, "
    strSQL = strSQL & "LA.intExercicio, "
    strSQL = strSQL & "LA.strNumeroAviso, "
    strSQL = strSQL & "EX.dtmDtCalculoPeticao, "
    strSQL = strSQL & "EX.strNumDistribuidor, "
    strSQL = strSQL & "EX.strIndexadorDESCR, "
    strSQL = strSQL & "EX.dblQuantIndexador, "
    strSQL = strSQL & "EX.dblVLIndexador, "
    strSQL = strSQL & "LA.intUtilizacao "
    strSQL = strSQL & "ORDER BY "
    strSQL = strSQL & "X, EX.pkId "
    
    strQuery = strSQL

End Function

Private Sub Form_Load()
    dbcInicial.Tag = "SELECT DISTINCT pkId," & gstrCONVERT(cdt_numeric, "intNumSeq ") & " intNumSeq FROM " & gstrExecutivo & "; INTNumSeq"
    dbcFinal.Tag = "SELECT DISTINCT pkId," & gstrCONVERT(cdt_numeric, "intNumSeq ") & " intNumSeq FROM " & gstrExecutivo & "; INTNumSeq"
End Sub

Private Sub dbcFinal_GotFocus()
    MarcaCampo dbcFinal
End Sub

Private Sub dbcFinal_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", dbcFinal, False
End Sub

Private Sub dbcInicial_GotFocus()
    MarcaCampo dbcInicial
End Sub

Private Sub dbcInicial_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", dbcInicial, False
End Sub

Private Sub txtintLote_GotFocus()
    MarcaCampo txtintLote
End Sub

Private Sub txtintLote_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtintLote, False
End Sub

Private Sub txtintVia_GotFocus()
    MarcaCampo txtintVia
End Sub

Private Sub txtintVia_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtintVia, False
End Sub

Private Function strQueryCertidaoDividaAtiva() As String
Dim intCont As Integer
Dim strSQL As String
Dim strVias As String

    strVias = ""
    
    For intCont = 1 To CInt(txtintVia)
        strVias = strVias & " SELECT " & intCont & " X "
        strVias = strVias & IIf(bytDBType = Oracle, " FROM DUAL ", " ")
        strVias = strVias & " UNION ALL "
    Next
    If strVias <> "" Then
        strVias = " (" & Mid(strVias, 1, Len(strVias) - 10) & ") Vias "
    End If

    strSQL = strSQL & " "
    strSQL = strSQL & "SELECT "
    strSQL = strSQL & gstrCONVERT(CDT_VARCHAR, "X") & strCONCAT & gstrCONVERT(CDT_VARCHAR, "DA.PKID") & " DATIVA, "
    strSQL = strSQL & "DA.intCertidao, "
    strSQL = strSQL & "DA.dtmDtInscricao, "
    strSQL = strSQL & "FL.strDescricao strFundamento, "
    
    strSQL = strSQL & "EX.strExecutadoNome strContribuinte, "
    strSQL = strSQL & "EX.STREXECUTADOCNPJCPF CPF, "
    strSQL = strSQL & "EX.STREXECUTADOIDENTIDADE RG, "
    strSQL = strSQL & "EX.Strexecutadotplognotif strTipoLogradouroC, "
    strSQL = strSQL & "EX.Strexecutadotitlognotif strTitLogradouroC, "
    strSQL = strSQL & "EX.Strexecutadonomelognotif strLogradouroC, "
    strSQL = strSQL & "EX.Strexecutadonumlognotif strNumeroC, "
    strSQL = strSQL & "EX.Strexecutadocomplnotif strComplementoC, "
    strSQL = strSQL & "EX.Strexecutadobairronotif strBairroC, "
    strSQL = strSQL & "EX.Strexecutadocidnotif strMunicipioC, "
    strSQL = strSQL & "EX.Strexecutadoufnotif strUFC, "
    strSQL = strSQL & "EX.Intexecutadocepnotif intCEPC, "
    
    strSQL = strSQL & "DA.Strlogradouro, "
    strSQL = strSQL & "DA.strNumero, "
    strSQL = strSQL & "DA.strComplemento, "
    strSQL = strSQL & "DA.strBairro, "
    strSQL = strSQL & "DA.strMunicipio, "
    strSQL = strSQL & "DA.strUF, "
    strSQL = strSQL & "DA.intCEP, "
    
    strSQL = strSQL & "LA.strInscricao, "
    strSQL = strSQL & "LA.intUtilizacao, "
    strSQL = strSQL & "LA.strNumeroAviso, "
    strSQL = strSQL & "LA.intExercicio, "
    strSQL = strSQL & "LA.Strcomposicaodareceita, "
    
    strSQL = strSQL & "EP.Intparcela, "
    strSQL = strSQL & "EP.Dtmdtvencimento, "
    strSQL = strSQL & "EP.Dblvloriginal, "
    strSQL = strSQL & "EP.Dblvlprincipal, "
    strSQL = strSQL & "EP.Dblvlcorrecao, "
    strSQL = strSQL & "EP.Dblvlmulta, "
    strSQL = strSQL & "EP.Dblvljuros, "
    strSQL = strSQL & "EP.Dblvltotal, "
    
    strSQL = strSQL & "EX.Dblvlindexador, "
    strSQL = strSQL & "EX.dblQuantIndexador dblQuantidadeIndexador, "
    strSQL = strSQL & "EX.Dtmdtcalculopeticao, "
    strSQL = strSQL & "EX.Strnumdistribuidor, "
    strSQL = strSQL & "EX.intNumSeq, "
    
    strSQL = strSQL & gstrCONVERT(CDT_NVARCHAR, "EX.strNumDistribuidor") & strCONCAT & " '/' " & _
                      strCONCAT & gstrCONVERT(CDT_NVARCHAR, "EX.intNumSeq") & " Controle "
    
    strSQL = strSQL & "FROM "
    strSQL = strSQL & gstrDativa & " DA, "
    strSQL = strSQL & gstrExecutivo & " EX, "
    strSQL = strSQL & gstrExecutivoParcela & " EP, "
    strSQL = strSQL & gstrFundamentoLegal & " FL, "
    strSQL = strSQL & gstrLancamentoAlfa & " LA, "
    strSQL = strSQL & strVias
    
    strSQL = strSQL & "WHERE "
    strSQL = strSQL & "DA.intLancamentoAlfa = LA.PkID "
    strSQL = strSQL & "AND DA.intExecutivo = EX.pkId "
    strSQL = strSQL & "AND EP.intDativa = DA.pkId "
    strSQL = strSQL & "AND FL.intComposicaoDaReceita " & strOUTJOracle & " =" & strOUTJSQLServer & " LA.intComposicaoDaReceita "
    strSQL = strSQL & "AND FL.intExercicio " & strOUTJOracle & " =" & strOUTJSQLServer & " LA.intExercicio "
    strSQL = strSQL & "AND EX.intLoteExecutivo = " & txtintLote
    strSQL = strSQL & "AND EX.intNumSeq BETWEEN " & dbcInicial.Text & " AND " & dbcFinal.Text & " "
    
    strSQL = strSQL & "ORDER BY X, EX.intNumSeq, DA.pkID, EP.intParcela"
    
    strQueryCertidaoDividaAtiva = strSQL

End Function

Private Function VerificaArquivos() As Boolean


End Function
