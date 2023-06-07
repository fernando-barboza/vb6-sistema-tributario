VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmCadExecucaoFiscal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Execução Fiscal"
   ClientHeight    =   3630
   ClientLeft      =   2250
   ClientTop       =   4035
   ClientWidth     =   8535
   Icon            =   "CadExecucaoFiscal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3630
   ScaleWidth      =   8535
   Begin TabDlg.SSTab tab_3dPasta 
      Height          =   3570
      Left            =   30
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   30
      Width           =   8475
      _ExtentX        =   14949
      _ExtentY        =   6297
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Execução Fiscal"
      TabPicture(0)   =   "CadExecucaoFiscal.frx":1042
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fra_Inscricao"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      Begin VB.Frame fra_Inscricao 
         Height          =   645
         Left            =   150
         TabIndex        =   16
         Top             =   390
         Width           =   8175
         Begin VB.OptionButton optbitTipoDeInscricao 
            Caption         =   "Imobiliário Urbano"
            CausesValidation=   0   'False
            Height          =   195
            Index           =   0
            Left            =   1680
            TabIndex        =   17
            Top             =   360
            Value           =   -1  'True
            Width           =   1725
         End
         Begin VB.OptionButton optbitTipoDeInscricao 
            Caption         =   "Econômico"
            Height          =   195
            Index           =   1
            Left            =   5460
            TabIndex        =   0
            Top             =   360
            Width           =   1605
         End
      End
      Begin VB.Frame Frame1 
         Height          =   2265
         Left            =   150
         TabIndex        =   10
         Top             =   1110
         Width           =   8175
         Begin VB.TextBox txt_intParcelaFinal 
            Height          =   285
            Left            =   7680
            MaxLength       =   2
            TabIndex        =   6
            Top             =   1410
            Width           =   345
         End
         Begin VB.TextBox txt_intLei 
            Height          =   285
            Left            =   6810
            MaxLength       =   8
            TabIndex        =   8
            Top             =   1800
            Width           =   1215
         End
         Begin VB.TextBox txt_dtmDataDaExecucao 
            Height          =   285
            Left            =   2100
            MaxLength       =   12
            TabIndex        =   7
            Top             =   1800
            Width           =   1035
         End
         Begin VB.TextBox txt_intExercicio 
            Height          =   285
            Left            =   2100
            MaxLength       =   4
            TabIndex        =   4
            Top             =   1425
            Width           =   525
         End
         Begin VB.TextBox txt_intParcelaInicial 
            Height          =   285
            Left            =   6930
            MaxLength       =   2
            TabIndex        =   5
            Top             =   1425
            Width           =   345
         End
         Begin MSDataListLib.DataCombo dbc_strInscricaoFinal 
            Height          =   315
            Left            =   2100
            TabIndex        =   2
            Top             =   630
            Width           =   5925
            _ExtentX        =   10451
            _ExtentY        =   556
            _Version        =   393216
            IntegralHeight  =   0   'False
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dbc_strInscricaoInicial 
            Height          =   315
            Left            =   2100
            TabIndex        =   1
            Top             =   240
            Width           =   5925
            _ExtentX        =   10451
            _ExtentY        =   556
            _Version        =   393216
            IntegralHeight  =   0   'False
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dbc_intComposicaoDaReceita 
            Height          =   315
            Left            =   2100
            TabIndex        =   3
            Top             =   1020
            Width           =   5925
            _ExtentX        =   10451
            _ExtentY        =   556
            _Version        =   393216
            IntegralHeight  =   0   'False
            Style           =   2
            Text            =   ""
         End
         Begin VB.Label lbl_intParcelaFinal 
            AutoSize        =   -1  'True
            Caption         =   "a"
            Height          =   195
            Left            =   7410
            TabIndex        =   20
            Top             =   1500
            Width           =   90
         End
         Begin VB.Label lbl_intLei 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "N° do Processo de Execução"
            Height          =   195
            Left            =   4560
            TabIndex        =   19
            Top             =   1875
            Width           =   2100
         End
         Begin VB.Label lbl_dtmDataDaPrescricao 
            AutoSize        =   -1  'True
            Caption         =   "Data da Execução"
            Height          =   195
            Left            =   585
            TabIndex        =   18
            Top             =   1875
            Width           =   1335
         End
         Begin VB.Label lbl_ComposicaoDaReceita 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Composição da Receita"
            Height          =   195
            Left            =   240
            TabIndex        =   15
            Top             =   1110
            Width           =   1695
         End
         Begin VB.Label lbl_InscricaoFinal 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Inscrição Cadastral Final"
            Height          =   195
            Left            =   225
            TabIndex        =   14
            Top             =   720
            Width           =   1725
         End
         Begin VB.Label lbl_InscricaoInicial 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Inscrição Cadastral Inicial"
            Height          =   195
            Left            =   150
            TabIndex        =   13
            Top             =   330
            Width           =   1800
         End
         Begin VB.Label lbl_intExercicio 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Exercício"
            Height          =   195
            Left            =   1275
            TabIndex        =   12
            Top             =   1500
            Width           =   675
         End
         Begin VB.Label lbl_intParcelaInicial 
            AutoSize        =   -1  'True
            Caption         =   "N° da parcela a ser executada de"
            Height          =   195
            Left            =   4440
            TabIndex        =   11
            Top             =   1500
            Width           =   2385
         End
      End
   End
End
Attribute VB_Name = "frmCadExecucaoFiscal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim mblnAlterando                   As Boolean
    Dim mobjAux                         As Object
    Dim mblnSelecionou                  As Boolean
    Dim mblnPrimeiraVez                 As Boolean
    Dim intCodigoInicial                As Integer
    Dim intCodigoFinal                  As Integer
    Dim CCInicial                       As Integer
    Dim CCFinal                         As Integer
    Dim TipoDeInscricao                 As Integer
Private Sub dbc_strInscricaoFinal_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbc_strInscricaoFinal
End Sub

Private Sub dbc_strInscricaoInicial_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbc_strInscricaoInicial
End Sub

Private Sub txt_dtmDataDaExecucao_LostFocus()
    txt_dtmDataDaExecucao = gstrDataFormatada(txt_dtmDataDaExecucao)
End Sub

Private Sub txt_intExercicio_GotFocus()
    MarcaCampo txt_intExercicio
End Sub

Private Sub txt_intExercicio_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txt_intExercicio
End Sub

Private Sub txt_intParcelaInicial_GotFocus()
    MarcaCampo txt_intParcelaInicial
End Sub

Private Sub txt_intParcelaInicial_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txt_intParcelaInicial
End Sub

Private Sub txt_intParcelaFinal_GotFocus()
    MarcaCampo txt_intParcelaFinal
End Sub

Private Sub txt_intParcelaFinal_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txt_intParcelaFinal
End Sub

Private Sub txt_dtmDataDaExecucao_GotFocus()
    MarcaCampo txt_dtmDataDaExecucao
End Sub

Private Sub txt_dtmDataDaExecucao_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "D", txt_dtmDataDaExecucao
End Sub

Private Sub txt_intLei_KeyPress(KeyAscii As Integer)
    If Chr(KeyAscii) <> "/" Then
        CaracterValido KeyAscii, "N", txt_intLei
    End If
End Sub

Private Sub dbc_intComposicaoDaReceita_Click(Area As Integer)
   DropDownDataCombo dbc_intComposicaoDaReceita, Me, Area
End Sub

Private Sub dbc_intComposicaoDaReceita_KeyDown(KeyCode As Integer, Shift As Integer)
   DropDownDataCombo dbc_intComposicaoDaReceita, Me, , KeyCode, Shift
End Sub

Private Sub dbc_intComposicaoDaReceita_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbc_intComposicaoDaReceita
End Sub

Private Sub dbc_strInscricaoFinal_Click(Area As Integer)
   DropDownDataCombo dbc_strInscricaoFinal, Me, Area
End Sub

Private Sub dbc_strInscricaoFinal_KeyDown(KeyCode As Integer, Shift As Integer)
   DropDownDataCombo dbc_strInscricaoFinal, Me, , KeyCode, Shift
End Sub

Private Sub dbc_strInscricaoInicial_Click(Area As Integer)
   DropDownDataCombo dbc_strInscricaoInicial, Me, Area
End Sub

Private Sub dbc_strInscricaoInicial_KeyDown(KeyCode As Integer, Shift As Integer)
   DropDownDataCombo dbc_strInscricaoInicial, Me, , KeyCode, Shift
End Sub

Private Sub Form_Activate()
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrSalvar, gstrAplicar, gstrDeletar, gstrImprimir
    HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrNovo, gstrFechar
    HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrCalcularReajuste
End Sub

Private Sub Form_Deactivate()
    HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrSalvar, gstrAplicar, gstrDeletar, gstrImprimir
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrNovo, gstrFechar
    HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrCalcularReajuste
End Sub

Private Sub Form_Load()
    CCInicial = 0
    CCFinal = 0
    dbc_strInscricaoInicial.Tag = strQueryInscricao(0) & ";A.strInscricaoAnterior"
    dbc_strInscricaoFinal.Tag = strQueryInscricao(0) & ";A.strInscricaoAnterior"
    LeDaTabelaParaObj gstrComposicaoDaReceita, dbc_intComposicaoDaReceita, strQuerryComposicao(0)
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrSalvar, gstrAplicar, gstrDeletar, gstrImprimir
    HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrNovo, gstrFechar
    HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrCalcularReajuste
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrCalcularReajuste
End Sub

Private Function blnValidaDados() As Boolean
    
    If Not dbc_strInscricaoInicial.MatchedWithList Then
        ExibeMensagem "A Inscrição Inicial tem que ser selecionada."
        dbc_strInscricaoInicial.SetFocus
        Exit Function
    End If
    
    If Not dbc_strInscricaoFinal.MatchedWithList Then
        ExibeMensagem "A Inscrição Final tem que ser selecionada."
        dbc_strInscricaoFinal.SetFocus
        Exit Function
    End If
    
    If Not dbc_intComposicaoDaReceita.MatchedWithList Then
        ExibeMensagem "A Composição da Receita tem que ser selecionada."
        dbc_intComposicaoDaReceita.SetFocus
        Exit Function
    End If
    
    If Trim(txt_intExercicio.Text) = "" Then
        ExibeMensagem "O exercício tem que ser digitado."
        txt_intExercicio.SetFocus
        Exit Function
    End If
    
    If Trim(txt_intParcelaInicial.Text) = "" Then
        ExibeMensagem "O Nº da parcela inicial tem que ser digitado."
        txt_intParcelaInicial.SetFocus
        Exit Function
    End If
    
    If Trim(txt_intParcelaFinal.Text) = "" Then
        ExibeMensagem "O Nº da parcela final tem que ser digitado."
        txt_intParcelaFinal.SetFocus
        Exit Function
    End If
    
    If Trim(txt_dtmDataDaExecucao.Text) = "" Then
        ExibeMensagem " A Data do Cancelamento tem que ser digitada."
        txt_dtmDataDaExecucao.SetFocus
        Exit Function
    ElseIf Not gblnDataValida(txt_dtmDataDaExecucao.Text, True) Then
        txt_dtmDataDaExecucao.SetFocus
        Exit Function
    End If
    
    If Trim(txt_intLei.Text) = "" Then
        ExibeMensagem "O Nº do processo tem que ser digitado."
        txt_intLei.SetFocus
        Exit Function
    End If
    
    blnValidaDados = True
End Function


Public Sub MantemForm(ByVal strModoOperacao As String)

    Dim strSQL As String
    Dim intAuxIndice As Integer
    On Error Resume Next

    If UCase(strModoOperacao) = gstrPreencherLista Then
        For intAuxIndice = 0 To optbitTipoDeInscricao.Count - 1
            If optbitTipoDeInscricao(intAuxIndice).Value = True Then
                Exit For
            End If
        Next
        
        dbc_strInscricaoInicial.Tag = strQueryInscricao(intAuxIndice) & ";A.strInscricaoAnterior"
        dbc_strInscricaoFinal.Tag = strQueryInscricao(intAuxIndice) & ";A.strInscricaoAnterior"
        
        PreencherListaDeOpcoes dbc_strInscricaoInicial
        PreencherListaDeOpcoes dbc_strInscricaoFinal
        Exit Sub
    End If

    If blnValidaDados Then
        If UCase(strModoOperacao) = UCase(gstrCalcularReajuste) Then
            Set gobjBanco = New clsBanco
            gobjBanco.ExecutaBeginTrans
            strSQL = ""
            strSQL = strSQL & gstrStoredProcedure("sp_Cobranca_Administrativo", "4," & TipoDeInscricao & ",'" & _
                            dbc_strInscricaoInicial.BoundText & "','" & dbc_strInscricaoFinal.BoundText & "'," & _
                            dbc_intComposicaoDaReceita.BoundText & "," & txt_intExercicio.Text & "," & _
                            txt_intParcelaInicial & "," & txt_intParcelaFinal & ",'" & txt_intLei & "'," & _
                            gstrConvDtParaSql(txt_dtmDataDaExecucao) & "," & glngCodUsr)
            Set gobjBanco = New clsBanco
            If gobjBanco.Execute(strSQL, False) Then
                gobjBanco.ExecutaCommitTrans
                ExibeMensagem "Cálculo efetuado com sucesso!"
            Else
                gobjBanco.ExecutaRollbackTrans
            End If
        End If
    End If
    If UCase(strModoOperacao) = UCase(gstrNovo) Then
        LimpaControlesDoFormulario
    End If
    
    If UCase(strModoOperacao) = UCase(gstrFechar) Then
        Unload Me
    End If
    
End Sub

Sub LimpaControlesDoFormulario()
    optbitTipoDeInscricao(0).Value = True
    dbc_strInscricaoInicial.BoundText = ""
    dbc_strInscricaoFinal.BoundText = ""
    dbc_intComposicaoDaReceita.BoundText = ""
    txt_intExercicio.Text = ""
    txt_intParcelaInicial.Text = ""
    txt_intParcelaFinal = ""
    txt_intLei.Text = ""
    txt_dtmDataDaExecucao.Text = ""
    optbitTipoDeInscricao(1).SetFocus
End Sub

Private Sub optbitTipoDeInscricao_Click(Index As Integer)
Dim strSQL As String
Dim intIndice As Integer

    TipoDeInscricao = 0
    TipoDeInscricao = Val(Index)

    optbitTipoDeInscricao(Index).CausesValidation = True

    For intIndice = 0 To 1
        If intIndice <> Index Then
            optbitTipoDeInscricao(intIndice).CausesValidation = False
        End If
    Next

    Set dbc_strInscricaoInicial.RowSource = Nothing
    dbc_strInscricaoInicial.Text = ""
    Set dbc_strInscricaoFinal.RowSource = Nothing
    dbc_strInscricaoFinal.Text = ""
    
    dbc_intComposicaoDaReceita.BoundText = 0
    LeDaTabelaParaObj gstrComposicaoDaReceita, dbc_intComposicaoDaReceita, strQuerryComposicao(Index)
    
End Sub

Private Function strQuerryComposicao(Index As Integer) As String
    Dim strSQL As String
    Dim Utilizacao As Integer
    
    Utilizacao = 0
    
    If Index = 0 Then
        Utilizacao = 1
    ElseIf Index = 1 Then
        Utilizacao = 2
    End If
    
    strSQL = ""
    strSQL = strSQL & " SELECT PKId, strDescricao "
    
    strSQL = strSQL & " FROM "
    strSQL = strSQL & gstrComposicaoDaReceita
    
    strSQL = strSQL & " WHERE "
    strSQL = strSQL & " intUtilizacao = " & Utilizacao
    
    strSQL = strSQL & " ORDER BY strDescricao "
    
    strQuerryComposicao = strSQL
End Function

Private Function strQueryInscricao(Index As Integer) As String
    Dim strSQL As String
    
    strSQL = ""
    
    If Index = 0 Then
        strSQL = strSQL & " SELECT A.strInscricao, LTRIM(RTRIM(A.strInscricao)) " & strCONCAT & " ' - ' " & strCONCAT & "  LTRIM(RTRIM(B.strNome)) AS Descricao "
    ElseIf Index = 1 Then
        strSQL = strSQL & " SELECT A.strInscricaoCadastral, LTRIM(RTRIM(A.strInscricaoCadastral)) " & strCONCAT & " ' - ' " & strCONCAT & "  LTRIM(RTRIM(B.strNome)) AS Descricao "
    End If
    
    strSQL = strSQL & " FROM "
    
    If Index = 0 Then
        strSQL = strSQL & gstrImobiliario & " A, "
        strSQL = strSQL & gstrContribuinte & " B "
    ElseIf Index = 1 Then
        strSQL = strSQL & gstrEconomico & " A, "
        strSQL = strSQL & gstrContribuinte & " B "
    End If
    
    strSQL = strSQL & " WHERE "
    
    If Index = 0 Then
        strSQL = strSQL & " A.intContribuinte = B.PKId "
        strSQL = strSQL & " ORDER BY " & gstrCONVERT(CDT_NUMERIC, "strInscricao")
    ElseIf Index = 1 Then
        strSQL = strSQL & " A.intContribuinte = B.PKId "
        strSQL = strSQL & " ORDER BY " & gstrCONVERT(CDT_NUMERIC, "strInscricaoCadastral")
    End If
    
    strQueryInscricao = strSQL
End Function

