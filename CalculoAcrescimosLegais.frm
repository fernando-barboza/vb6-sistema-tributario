VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmCalculoAcrescimosLegais 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cálculo de Acréscimos Legais"
   ClientHeight    =   4410
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9735
   Icon            =   "CalculoAcrescimosLegais.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4410
   ScaleWidth      =   9735
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab tab_3dPasta 
      Height          =   4035
      Left            =   210
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   180
      Width           =   9345
      _ExtentX        =   16484
      _ExtentY        =   7117
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Cálculo de Acréscimos Legais"
      TabPicture(0)   =   "CalculoAcrescimosLegais.frx":1042
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lbl_strInscricaoCadastralFinal"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lbl_strInscricaoCadastralInicial"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblIndexador"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lbl_dtmDataVencimento"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "dbc_Indexador"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "dbc_strInscricaoCadastralInicial"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "dbc_strInscricaoCadastralFinal"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "fra_Origem"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "chk_Selecionar"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "fra_formula"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txt_dtmDataVencimento"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).ControlCount=   11
      Begin VB.TextBox txt_dtmDataVencimento 
         Height          =   285
         Left            =   7200
         MaxLength       =   15
         TabIndex        =   9
         Top             =   2160
         Width           =   1035
      End
      Begin VB.Frame fra_formula 
         Caption         =   "Fórmulas de Cálculo"
         Height          =   675
         Left            =   240
         TabIndex        =   14
         Top             =   1320
         Width           =   8865
         Begin VB.CheckBox chk_Juros 
            Caption         =   "Juros"
            Height          =   255
            Left            =   1290
            TabIndex        =   5
            Top             =   270
            Width           =   735
         End
         Begin VB.CheckBox chk_Multa 
            Caption         =   "Multa"
            Height          =   255
            Left            =   3840
            TabIndex        =   6
            Top             =   270
            Width           =   915
         End
         Begin VB.CheckBox chk_Correcao 
            Caption         =   "Correção Monetária"
            Height          =   255
            Left            =   6120
            TabIndex        =   7
            Top             =   270
            Width           =   1875
         End
      End
      Begin VB.CheckBox chk_Selecionar 
         Caption         =   "Selecionar todas as Inscrições"
         Height          =   255
         Left            =   2400
         TabIndex        =   12
         Top             =   3330
         Width           =   2835
      End
      Begin VB.Frame fra_Origem 
         Caption         =   " Origem"
         Height          =   675
         Left            =   240
         TabIndex        =   13
         Top             =   480
         Width           =   8865
         Begin VB.OptionButton optOrigem 
            Caption         =   "Receitas Diversas"
            Height          =   195
            Index           =   4
            Left            =   7110
            TabIndex        =   4
            Top             =   300
            Width           =   1605
         End
         Begin VB.OptionButton optOrigem 
            Caption         =   "Contribuição de Melhorias"
            Height          =   195
            Index           =   3
            Left            =   4860
            TabIndex        =   3
            Top             =   300
            Width           =   2265
         End
         Begin VB.OptionButton optOrigem 
            Caption         =   "Econômico"
            Height          =   195
            Index           =   2
            Left            =   3510
            TabIndex        =   2
            Top             =   300
            Width           =   1155
         End
         Begin VB.OptionButton optOrigem 
            Caption         =   "Imobiliário Urbano"
            Height          =   195
            Index           =   1
            Left            =   1770
            TabIndex        =   1
            Top             =   300
            Width           =   1695
         End
         Begin VB.OptionButton optOrigem 
            Caption         =   "Imobiliário Rural"
            Height          =   195
            Index           =   0
            Left            =   210
            TabIndex        =   0
            Top             =   300
            Width           =   1575
         End
      End
      Begin MSDataListLib.DataCombo dbc_strInscricaoCadastralFinal 
         Height          =   315
         Left            =   2400
         TabIndex        =   11
         Top             =   2940
         Width           =   5865
         _ExtentX        =   10345
         _ExtentY        =   556
         _Version        =   393216
         IntegralHeight  =   0   'False
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dbc_strInscricaoCadastralInicial 
         Height          =   315
         Left            =   2400
         TabIndex        =   10
         Top             =   2550
         Width           =   5865
         _ExtentX        =   10345
         _ExtentY        =   556
         _Version        =   393216
         IntegralHeight  =   0   'False
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dbc_Indexador 
         Height          =   315
         Left            =   2400
         TabIndex        =   8
         Top             =   2160
         Width           =   2325
         _ExtentX        =   4101
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
      Begin VB.Label lbl_dtmDataVencimento 
         AutoSize        =   -1  'True
         Caption         =   "Data de Vencimento"
         Height          =   195
         Left            =   5610
         TabIndex        =   19
         Top             =   2220
         Width           =   1455
      End
      Begin VB.Label lblIndexador 
         AutoSize        =   -1  'True
         Caption         =   "Indexador"
         Height          =   225
         Left            =   1575
         TabIndex        =   18
         Top             =   2250
         Width           =   705
      End
      Begin VB.Label lbl_strInscricaoCadastralInicial 
         AutoSize        =   -1  'True
         Caption         =   "Inscrição Cadastral Inicial"
         Height          =   195
         Left            =   480
         TabIndex        =   17
         Top             =   2640
         Width           =   1800
      End
      Begin VB.Label lbl_strInscricaoCadastralFinal 
         AutoSize        =   -1  'True
         Caption         =   "Inscrição Cadastral Final"
         Height          =   195
         Left            =   555
         TabIndex        =   16
         Top             =   3030
         Width           =   1725
      End
   End
End
Attribute VB_Name = "frmCalculoAcrescimosLegais"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mblnAlterando           As Boolean
Dim mobjAux                 As Object
Dim mblnSelecionou          As Boolean
Dim mblnPrimeiraVez         As Boolean
Dim adoRecDadosReceita      As ADODB.Recordset
Dim adoRecDadosTaxa         As ADODB.Recordset

Private Sub chk_Selecionar_Click()
    If chk_Selecionar.Value = 1 Then
        dbc_strInscricaoCadastralInicial.BoundText = ""
        dbc_strInscricaoCadastralFinal.BoundText = ""
        dbc_strInscricaoCadastralInicial.Enabled = False
        TrocaCorObjeto dbc_strInscricaoCadastralInicial, True
        dbc_strInscricaoCadastralFinal.Enabled = False
        TrocaCorObjeto dbc_strInscricaoCadastralFinal, True
    Else
        dbc_strInscricaoCadastralInicial.Enabled = True
        TrocaCorObjeto dbc_strInscricaoCadastralInicial, False
        dbc_strInscricaoCadastralFinal.Enabled = True
        TrocaCorObjeto dbc_strInscricaoCadastralFinal, False
        dbc_strInscricaoCadastralInicial.SetFocus
    End If

End Sub

Private Sub dbc_Indexador_Click(Area As Integer)
    DropDownDataCombo dbc_Indexador, Me, Area
End Sub

Private Sub dbc_Indexador_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbc_Indexador, Me, , KeyCode, Shift
End Sub

Private Sub dbc_strInscricaoCadastralFinal_Click(Area As Integer)
    DropDownDataCombo dbc_strInscricaoCadastralFinal, Me, Area
End Sub

Private Sub dbc_strInscricaoCadastralFinal_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbc_strInscricaoCadastralFinal, Me, , KeyCode, Shift
End Sub

Private Sub dbc_strInscricaoCadastralInicial_Click(Area As Integer)
    DropDownDataCombo dbc_strInscricaoCadastralInicial, Me, Area
End Sub

Private Sub dbc_strInscricaoCadastralInicial_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbc_strInscricaoCadastralInicial, Me, , KeyCode, Shift
End Sub

Private Sub Form_Activate()
    gintCodSeguranca = 646
    VirificaGradeListView Me
    If mblnSelecionou Then
       HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrDeletar
    Else
       HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrDeletar
    End If
    If mobjAux Is Nothing Then
       HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar
    Else
       HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrAplicar
    End If
    HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrCalcularReajuste
End Sub

Private Sub Form_Deactivate()
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar
End Sub

Private Sub Form_Load()
    mblnAlterando = False
    LeDaTabelaParaObj gstrIndexadorEconomico, dbc_Indexador, "PKId, strSiglaIndexador "
    VerificaObjParaAplicar mobjAux
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar
    mblnSelecionou = False
    mblnPrimeiraVez = False
End Sub

Private Sub optOrigem_Click(Index As Integer)
    Dim strSql As String
    Dim intIndice As Integer

    optOrigem(Index).CausesValidation = True
    
    If Index = 4 Then
        lbl_strInscricaoCadastralInicial.Caption = "Contribuinte Inicial"
        lbl_strInscricaoCadastralFinal.Caption = "Contribuinte Final"
    Else
        lbl_strInscricaoCadastralInicial.Caption = "Inscrição Inicial"
        lbl_strInscricaoCadastralFinal.Caption = "Inscrição Final"
    End If


    For intIndice = 0 To 4
        If intIndice <> Index Then
            optOrigem(intIndice).CausesValidation = False
        End If
    Next

    Set dbc_strInscricaoCadastralInicial.RowSource = Nothing
    Set dbc_strInscricaoCadastralFinal.RowSource = Nothing
    dbc_strInscricaoCadastralInicial.Text = ""
    dbc_strInscricaoCadastralFinal.Text = ""
End Sub

Private Function strQueryInscricao(Index As Integer) As String

'******************************************************************************************
' Data: 02/05/2003
' Alteração: - Substituição do comando nativo CONVERT do SQL Server pela função gstrCONVERT
' Responsável: Everton Bianchini
'------------------------------------------------------------------------------------------
' Data: 05/05/2003
' Alteração: - Substituição do comando nativo de concatenação de strings (+) do SQL Server
'            pela variável strCONCAT.
' Responsável: Everton Bianchini
'******************************************************************************************
    
    Dim strSql As String
    
    strSql = ""
    If Index = 0 Or Index = 1 Then
'        strSQL = strSQL & " SELECT B.PKId, LTRIM(RTRIM(A.strInscricaoAnterior)) + ' - ' +  LTRIM(RTRIM(B.strNome)) AS Descricao " 'A.strInscricaoAnterior
        strSql = strSql & " SELECT B.PKId, LTRIM(RTRIM(A.strInscricaoAnterior)) " & strCONCAT & " ' - ' " & strCONCAT & "  LTRIM(RTRIM(B.strNome)) AS Descricao " 'A.strInscricaoAnterior
    ElseIf Index = 2 Then
'        strSQL = strSQL & " SELECT B.PKId, LTRIM(RTRIM(A.strInscricaoCadastral)) + ' - ' +  LTRIM(RTRIM(B.strNome)) AS Descricao " 'A.strInscricaoCadastral
        strSql = strSql & " SELECT B.PKId, LTRIM(RTRIM(A.strInscricaoCadastral)) " & strCONCAT & " ' - ' " & strCONCAT & "  LTRIM(RTRIM(B.strNome)) AS Descricao " 'A.strInscricaoCadastral
    ElseIf Index = 3 Then
'        strSQL = strSQL & " SELECT C.PKId, LTRIM(RTRIM(A.strInscricaoAnterior)) + ' - ' +  LTRIM(RTRIM(C.strNome)) AS Descricao "
        strSql = strSql & " SELECT C.PKId, LTRIM(RTRIM(A.strInscricaoAnterior)) " & strCONCAT & " ' - ' " & strCONCAT & "  LTRIM(RTRIM(C.strNome)) AS Descricao "
    End If
    
    strSql = strSql & " FROM "
    
    If Index = 1 Then
        strSql = strSql & gstrImobiliario & " A, "
        strSql = strSql & gstrContribuinte & " B "
    ElseIf Index = 0 Then
        strSql = strSql & gstrImobiliarioRural & " A, "
        strSql = strSql & gstrContribuinte & " B "
    ElseIf Index = 2 Then
        strSql = strSql & gstrEconomico & " A, "
        strSql = strSql & gstrContribuinte & " B "
    ElseIf Index = 3 Then
        strSql = strSql & gstrImobiliario & " A, "
        strSql = strSql & gstrContribuicaoMelhoria & " B, "
        strSql = strSql & gstrContribuinte & " C "
    End If
    strSql = strSql & " WHERE "
    If Index = 0 Or Index = 1 Then
        strSql = strSql & " A.intContribuinte = B.PKId "
'        strSql = strSql & " ORDER BY convert(numeric,strInscricaoAnterior) "
        strSql = strSql & " ORDER BY " & gstrCONVERT(CDT_NUMERIC, "strInscricaoAnterior")
    ElseIf Index = 2 Then
        strSql = strSql & " A.intContribuinte = B.PKId "
'        strSql = strSql & " ORDER BY convert(numeric,strInscricaoCadastral) "
        strSql = strSql & " ORDER BY " & gstrCONVERT(CDT_NUMERIC, "strInscricaoCadastral")
    ElseIf Index = 3 Then
        strSql = strSql & " B.intImobiliario = A.PKId "
        strSql = strSql & " AND A.intContribuinte = C.PKId "
'        strSql = strSql & " ORDER BY convert(numeric,strInscricaoAnterior) "
        strSql = strSql & " ORDER BY " & gstrCONVERT(CDT_NUMERIC, "strInscricaoAnterior")
    End If
    
strQueryInscricao = strSql
End Function

Public Sub MantemForm(ByVal strModoOperacao As String)
Dim strSql As String
Dim i As Integer
Dim j As Integer

    If UCase(strModoOperacao) = UCase(gstrFechar) Then
        Unload Me
    End If
    
    If strModoOperacao = gstrCalcularReajuste Then
        EfetuaCalculoAcrescimosLegais
    End If

    If strModoOperacao = gstrPreencherLista Then
        For i = 0 To 4
            If optOrigem(i).Value Then
                j = i
                Exit For
            End If
        Next i
        
        Select Case j
            
            Case 4
                strSql = ""
                strSql = "SELECT DISTINCT REC.intContribuinte, CON.strNome FROM " & gstrReceitaDiversa & " REC, " & gstrContribuinte & " CON WHERE CON.PKId = REC.intContribuinte ORDER BY CON.strNome "
            Case Else
                strSql = strQueryInscricao(j)
        End Select
        dbc_strInscricaoCadastralInicial.Tag = strSql & ";strNome"
        dbc_strInscricaoCadastralFinal.Tag = dbc_strInscricaoCadastralInicial.Tag

        PreencherListaDeOpcoes Me.ActiveControl

    End If

End Sub

Private Sub EfetuaCalculoAcrescimosLegais()

'******************************************************************************************
' Data: 12/05/2003
' Alteração: - Incluídos comandos de inicialização e finalização de bloco PL/SQL permitindo
'            , assim, a execução de múltiplos comandos SQL de uma única vez.
' Responsável: Everton Bianchini
'******************************************************************************************

Dim strSql           As String
Dim strFormula       As String
Dim adorecFormula    As ADODB.Recordset
Dim Juros            As Double
Dim JurosTotal       As Double
Dim Multa            As Double
Dim MultaTotal       As Double
Dim Correcao         As Double
Dim CorrecaoTotal    As Double
Dim NovoValorParcela As Double
Dim QuantDias        As Integer

If blnDadosOk Then
    If MsgBox("Confirma início do cálculo dos Acréscimos Legais?", vbYesNo + vbQuestion) = vbNo Then
        Exit Sub
    End If
    
    BuscaDadosReceita
    BuscaDadosTaxa

    If adoRecDadosReceita.RecordCount = 0 Or adoRecDadosTaxa.RecordCount = 0 Then
      ExibeMensagem "Não foram encontrados dados a serem calculados!"
      Exit Sub
    End If
    
    If chk_Juros.Value = 1 Then
        strFormula = ProcuraFormula(10)
        Set gobjBanco = New clsBanco
        If gobjBanco.CriaADO(strFormula, 5, adorecFormula) Then
            Juros = adorecFormula.Fields(0)
        Else
            Juros = 0
        End If
    Else
        Juros = 0
    End If
    If chk_Multa.Value = 1 Then
        strFormula = ProcuraFormula(11)
        Set gobjBanco = New clsBanco
        If gobjBanco.CriaADO(strFormula, 5, adorecFormula) Then
            Multa = adorecFormula.Fields(0)
        Else
            Multa = 0
        End If
    Else
        Multa = 0
    End If
    
    Set gobjBanco = New clsBanco
    gobjBanco.ExecutaBeginTrans
    Screen.MousePointer = vbHourglass

    adoRecDadosReceita.MoveFirst
    adoRecDadosTaxa.MoveFirst
    NovoValorParcela = 0
    QuantDias = 0
    strSql = ""
    
    strSql = strSql & IIf((bytDBType = EDatabases.Oracle), "BEGIN ", "")
    
    Do While Not adoRecDadosReceita.EOF
    
        QuantDias = Abs(adoRecDadosReceita!dtmDataVencimento - CDate(txt_dtmDataVencimento.Text))
        
        If chk_Correcao.Value = 1 Then
            strFormula = ProcuraFormula(12)
            strFormula = strFormula & " " & gstrConvDtParaSql(adoRecDadosReceita!dtmDataVencimento) ' & ", " & dbc_Indexador.BoundText
            Set gobjBanco = New clsBanco
            If gobjBanco.CriaADO(strFormula, 5, adorecFormula) Then
                Correcao = adorecFormula.Fields(0)
            Else
                Correcao = 0
            End If
        Else
            Correcao = 0
        End If
        
        If chk_Juros.Value = 1 Then
            JurosTotal = (adoRecDadosReceita!dblValorParcela * Juros / 100)
        Else
            JurosTotal = 0
        End If
        If chk_Multa.Value = 1 Then
            MultaTotal = (adoRecDadosReceita!dblValorParcela * Multa / 100 / 30) * QuantDias
        Else
            MultaTotal = 0
        End If
        CorrecaoTotal = adoRecDadosReceita!dblValorParcela * Correcao
        
        NovoValorParcela = adoRecDadosReceita!dblValorParcela + JurosTotal + MultaTotal + CorrecaoTotal
        
        strSql = strSql & " UPDATE " & gstrParcelaReceita & " SET "
        strSql = strSql & " dblValorParcela = " & gstrConvVrParaSql(NovoValorParcela)
        strSql = strSql & " WHERE PKID = " & adoRecDadosReceita!PKId
        
        strSql = strSql & IIf((bytDBType = EDatabases.Oracle), "; ", "")
        
        adoRecDadosReceita.MoveNext
        
    Loop
    
    Do While Not adoRecDadosTaxa.EOF
    
        QuantDias = Abs(adoRecDadosTaxa!dtmDataVencimento - CDate(txt_dtmDataVencimento.Text))
        
        If chk_Correcao.Value = 1 Then
            strFormula = ProcuraFormula(12)
            strFormula = strFormula & " " & gstrConvDtParaSql(adoRecDadosTaxa!dtmDataVencimento) ' & ", " & dbc_Indexador.BoundText
            Set gobjBanco = New clsBanco
            If gobjBanco.CriaADO(strFormula, 5, adorecFormula) Then
                Correcao = adorecFormula.Fields(0)
            Else
                Correcao = 0
            End If
        Else
            Correcao = 0
        End If
        
        If chk_Juros.Value = 1 Then
            JurosTotal = (adoRecDadosTaxa!dblValorParcela * Juros / 100)
        Else
            JurosTotal = 0
        End If
        If chk_Multa.Value = 1 Then
            MultaTotal = (adoRecDadosTaxa!dblValorParcela * Multa / 100 / 30) * QuantDias
        Else
            MultaTotal = 0
        End If
        CorrecaoTotal = adoRecDadosTaxa!dblValorParcela * Correcao
        
        NovoValorParcela = adoRecDadosTaxa!dblValorParcela + JurosTotal + MultaTotal + CorrecaoTotal
        
        strSql = strSql & " UPDATE " & gstrParcelaTaxa & " SET "
        strSql = strSql & " dblValorParcela = " & gstrConvVrParaSql(NovoValorParcela)
        strSql = strSql & " WHERE PKID = " & adoRecDadosTaxa!PKId
        
        strSql = strSql & IIf((bytDBType = EDatabases.Oracle), "; ", "")
        
        adoRecDadosTaxa.MoveNext
        
    Loop
    
    strSql = strSql & IIf((bytDBType = EDatabases.Oracle), "END; ", "")
    
    Set gobjBanco = New clsBanco
    If gobjBanco.Execute(strSql, False) Then
        gobjBanco.ExecutaCommitTrans
        Screen.MousePointer = vbNormal
        ExibeMensagem "Cálculo efetuado com sucesso!"
    Else
        Screen.MousePointer = vbNormal
        gobjBanco.ExecutaRollbackTrans
    End If
 
End If
End Sub

Private Sub BuscaDadosReceita()
Dim strSql          As String

strSql = ""
strSql = strSql & " SELECT PR.PKId, LC.intContribuinte, LC.strInscricaoCadastral, LC.intExercicio, "
strSql = strSql & " PR.dtmDataVencimento, PR.intNumeroParcela, LC.intOcorrencia, PR.dblValorParcela, "
strSql = strSql & " LC.BytOrigem "
strSql = strSql & " FROM " & gstrLancamentoCalculo & " LC, "
strSql = strSql & gstrParcelaReceita & " PR "
strSql = strSql & " WHERE LC.PKId = PR.intLancamentoCalculo "
strSql = strSql & " AND PR.dtmDataVencimento < " & gstrConvDtParaSql(txt_dtmDataVencimento.Text)
If chk_Selecionar.Value <> 1 Then
    strSql = strSql & " AND LC.strInscricaoCadastral BETWEEN '" & dbc_strInscricaoCadastralInicial.Text & "' AND '" & dbc_strInscricaoCadastralFinal.Text & "'"
End If
strSql = strSql & " ORDER BY intContribuinte "

Set gobjBanco = New clsBanco
gobjBanco.CriaADO strSql, 5, adoRecDadosReceita

End Sub

Private Sub BuscaDadosTaxa()
Dim strSql          As String

strSql = ""
strSql = strSql & " SELECT PT.PKId, LC.intContribuinte, LC.strInscricaoCadastral, LC.intExercicio, "
strSql = strSql & " PT.dtmDataVencimento, PT.intNumeroParcela, LC.intOcorrencia, PT.dblValorParcela, "
strSql = strSql & " LC.BytOrigem "
strSql = strSql & " FROM " & gstrLancamentoCalculo & " LC, "
strSql = strSql & gstrParcelaTaxa & " PT "
strSql = strSql & " WHERE LC.PKId = PT.intLancamentoCalculo "
strSql = strSql & " AND PT.dtmDataVencimento < " & gstrConvDtParaSql(txt_dtmDataVencimento.Text)
If chk_Selecionar.Value <> 1 Then
    strSql = strSql & " AND LC.strInscricaoCadastral BETWEEN '" & dbc_strInscricaoCadastralInicial.Text & "' AND '" & dbc_strInscricaoCadastralFinal.Text & "'"
End If
strSql = strSql & " ORDER BY intContribuinte "

Set gobjBanco = New clsBanco
gobjBanco.CriaADO strSql, 5, adoRecDadosTaxa

End Sub

Private Function ProcuraFormula(intCodigo As Integer) As String
Dim strSql As String
Dim adorecFormula As ADODB.Recordset

strSql = ""
strSql = strSql & " SELECT FB.strDescricao "
strSql = strSql & " FROM " & gstrFormulaBasica & " FB "
strSql = strSql & " WHERE FB.bytTipoDeFormula =  " & intCodigo

Set gobjBanco = New clsBanco
gobjBanco.CriaADO strSql, 5, adorecFormula
ProcuraFormula = adorecFormula!strDescricao

End Function

Private Function blnDadosOk() As Boolean
       blnDadosOk = False
    
    If chk_Juros.Value = 0 And chk_Multa.Value = 0 And chk_Correcao.Value = 0 Then
        ExibeMensagem "Selecione uma Fórmula de Cálculo!"
        Exit Function
    End If
    If dbc_Indexador.BoundText = "" Then
        ExibeMensagem "Selecione um Indexador para efetuar o cálculo!"
        dbc_Indexador.SetFocus
        Exit Function
    End If
    If txt_dtmDataVencimento.Text = "" Then
        ExibeMensagem "Digite uma data de Vencimento para efetuar o cálculo!"
        txt_dtmDataVencimento.SetFocus
        Exit Function
    End If
    If chk_Selecionar.Value <> 1 Then
        If dbc_strInscricaoCadastralInicial.BoundText = "" Then
            ExibeMensagem "Selecione uma Inscrição Cadastral Inicial para efetuar o cálculo! "
            dbc_strInscricaoCadastralInicial.SetFocus
            Exit Function
        End If
        If dbc_strInscricaoCadastralFinal.BoundText = "" Then
            ExibeMensagem "Selecione uma Inscrição Cadastral Final para efetuar o cálculo! "
            dbc_strInscricaoCadastralFinal.SetFocus
            Exit Function
        End If
    End If
    
    blnDadosOk = True
End Function

'############################ Caracter Válido e Marca Campo ###########################

Private Sub dbc_Indexador_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "", dbc_Indexador
End Sub

Private Sub dbc_strInscricaoCadastralInicial_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "", dbc_strInscricaoCadastralInicial
End Sub

Private Sub dbc_strInscricaoCadastralfinal_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "", dbc_strInscricaoCadastralFinal
End Sub

Private Sub txt_dtmDataVencimento_GotFocus()
    MarcaCampo txt_dtmDataVencimento
End Sub

Private Sub txt_dtmDataVencimento_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "D", txt_dtmDataVencimento
End Sub

Private Sub txt_dtmDataVencimento_LostFocus()
txt_dtmDataVencimento = gstrDataFormatada(txt_dtmDataVencimento)
End Sub
