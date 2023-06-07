VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmRelatorioDeIsencoesImunidadesNaoIncidente 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Beneficiados com Imunidade / Isenção / Não Incidência"
   ClientHeight    =   2850
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8775
   Icon            =   "RelatorioDeIsencoesImunidadesNaoIncidente.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2850
   ScaleWidth      =   8775
   Begin TabDlg.SSTab tab_3dPasta 
      Height          =   2535
      Left            =   150
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   150
      Width           =   8475
      _ExtentX        =   14949
      _ExtentY        =   4471
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Beneficiados com Imunidade / Isenção / Não Incidência"
      TabPicture(0)   =   "RelatorioDeIsencoesImunidadesNaoIncidente.frx":1042
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fra_Inscricao"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      Begin VB.Frame Frame1 
         Height          =   1305
         Left            =   150
         TabIndex        =   10
         Top             =   1050
         Width           =   8175
         Begin VB.CheckBox chk_Selecionar 
            Caption         =   "Selecionar todos os Contribuintes"
            Height          =   255
            Left            =   1590
            TabIndex        =   3
            Top             =   990
            Width           =   2835
         End
         Begin MSDataListLib.DataCombo dbcintContribuinteInicial 
            Height          =   315
            Left            =   1590
            TabIndex        =   1
            Top             =   210
            Width           =   6465
            _ExtentX        =   11404
            _ExtentY        =   556
            _Version        =   393216
            IntegralHeight  =   0   'False
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dbcintContribuinteFinal 
            Height          =   315
            Left            =   1590
            TabIndex        =   2
            Top             =   630
            Width           =   6465
            _ExtentX        =   11404
            _ExtentY        =   556
            _Version        =   393216
            IntegralHeight  =   0   'False
            Text            =   ""
         End
         Begin VB.Label lbl_label2 
            AutoSize        =   -1  'True
            Caption         =   "Contribuinte Final"
            Height          =   195
            Left            =   255
            TabIndex        =   12
            Top             =   735
            Width           =   1215
         End
         Begin VB.Label lbl_Label1 
            AutoSize        =   -1  'True
            Caption         =   "Contribuinte Inicial"
            Height          =   195
            Left            =   180
            TabIndex        =   11
            Top             =   315
            Width           =   1290
         End
      End
      Begin VB.Frame fra_Inscricao 
         Height          =   645
         Left            =   150
         TabIndex        =   5
         Top             =   390
         Width           =   8175
         Begin VB.OptionButton optbitTipoDeInscricao 
            Caption         =   "Imobiliário Urbano"
            Height          =   195
            Index           =   0
            Left            =   60
            TabIndex        =   0
            Top             =   270
            Value           =   -1  'True
            Width           =   1605
         End
         Begin VB.OptionButton optbitTipoDeInscricao 
            Caption         =   "Imobiliário Rural"
            CausesValidation=   0   'False
            Height          =   195
            Index           =   1
            Left            =   1680
            TabIndex        =   9
            Top             =   270
            Width           =   1425
         End
         Begin VB.OptionButton optbitTipoDeInscricao 
            Caption         =   "Econômico"
            CausesValidation=   0   'False
            Height          =   195
            Index           =   2
            Left            =   3150
            TabIndex        =   8
            Top             =   270
            Width           =   1155
         End
         Begin VB.OptionButton optbitTipoDeInscricao 
            Caption         =   "Contribuição de Melhorias"
            CausesValidation=   0   'False
            Height          =   195
            Index           =   3
            Left            =   4290
            TabIndex        =   7
            Top             =   270
            Width           =   2205
         End
         Begin VB.OptionButton optbitTipoDeInscricao 
            Caption         =   "Receitas Diversas"
            CausesValidation=   0   'False
            Height          =   195
            Index           =   4
            Left            =   6480
            TabIndex        =   6
            Top             =   270
            Width           =   1605
         End
      End
   End
End
Attribute VB_Name = "frmRelatorioDeIsencoesImunidadesNaoIncidente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mblnAlterando           As Boolean
Dim mobjAux                 As Object
Dim mblnSelecionou          As Boolean
Dim mblnPrimeiraVez         As Boolean
Dim intCodigoInicial        As Integer
Dim intCodigoFinal          As Integer
Dim CCInicial               As Integer
Dim CCFinal                 As Integer
Dim TipoDeInscricao         As Integer

Private Sub chk_Selecionar_Click()
    If chk_Selecionar.Value = 1 Then
        dbcintContribuinteInicial.BoundText = ""
        dbcintContribuinteFinal.BoundText = ""
        dbcintContribuinteInicial.Enabled = False
        TrocaCorObjeto dbcintContribuinteInicial, True
        dbcintContribuinteFinal.Enabled = False
        TrocaCorObjeto dbcintContribuinteFinal, True
    Else
        dbcintContribuinteInicial.Enabled = True
        TrocaCorObjeto dbcintContribuinteInicial, False
        dbcintContribuinteFinal.Enabled = True
        TrocaCorObjeto dbcintContribuinteFinal, False
    End If
End Sub

Private Sub dbcintContribuinteFinal_Click(Area As Integer)
    DropDownDataCombo dbcintContribuinteFinal, Me, Area
End Sub

Private Sub dbcintContribuinteFinal_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbcintContribuinteFinal, Me, , KeyCode, Shift
End Sub

Private Sub dbcintContribuinteInicial_Click(Area As Integer)
    DropDownDataCombo dbcintContribuinteInicial, Me, Area
End Sub

Private Sub dbcintContribuinteInicial_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbcintContribuinteInicial, Me, , KeyCode, Shift
End Sub

Private Sub Form_Activate()
    gintCodSeguranca = 458
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
End Sub

Private Sub Form_Deactivate()
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar
End Sub

Private Sub Form_Load()
    CCInicial = 0
    CCFinal = 0
    HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrSalvar, gstrDeletar
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar
    mblnSelecionou = False
    mblnPrimeiraVez = False
End Sub

Public Sub MantemForm(ByVal strModoOperacao As String)
On Error Resume Next
Dim strSQL As String
Dim Resultado As String
Dim adoResultado   As ADODB.Recordset
Dim i As Integer
Dim j As Integer


    If UCase(strModoOperacao) = UCase(gstrImprimir) Then
        If blnDadosOk Then
        If chk_Selecionar.Value = 0 Then
            If Val(dbcintContribuinteInicial.BoundText) < Val(dbcintContribuinteFinal.BoundText) Then
                CCInicial = Val(dbcintContribuinteInicial.BoundText)
                CCFinal = Val(dbcintContribuinteFinal.BoundText)
            Else
                CCInicial = Val(dbcintContribuinteFinal.BoundText)
                CCFinal = Val(dbcintContribuinteInicial.BoundText)
            End If
        End If
            ImprimeRelatorio rptRelatorioDeBenificiadosIsencaoImunidadeNaoIncidente, strQuerryRelatorio
        End If
    End If
    
    If UCase(strModoOperacao) = UCase(gstrNovo) Then
        LimpaObjetos
    End If
    
    If UCase(strModoOperacao) = UCase(gstrFechar) Then
        Unload Me
    End If
    
    If UCase(strModoOperacao) = UCase(gstrPreencherLista) Then
    
        For i = 0 To optbitTipoDeInscricao.Count - 1
            If optbitTipoDeInscricao(i).Value Then
                j = i
                Exit For
            End If
        Next i
    
        Select Case j
            Case 0
                strSQL = "SELECT DISTINCT IMO.intContribuinte, CON.strNome FROM " & gstrImobiliario & " IMO, " & gstrContribuinte & " CON WHERE CON.PKId = IMO.intContribuinte ORDER BY CON.strNome "
            Case 1
                strSQL = "SELECT DISTINCT IMORU.intContribuinte, CON.strNome FROM " & gstrImobiliarioRural & " IMORU, " & gstrContribuinte & " CON WHERE CON.PKId = IMORU.intContribuinte ORDER BY CON.strNome "
            Case 2
                strSQL = "SELECT DISTINCT ECO.intContribuinte, CON.strNome FROM " & gstrEconomico & " ECO, " & gstrContribuinte & " CON WHERE CON.PKId = ECO.intContribuinte ORDER BY CON.strNome "
            Case 3
                strSQL = "SELECT DISTINCT IMO.intContribuinte, CON.strNome FROM " & gstrContribuicaoMelhoria & " CM, " & gstrImobiliario & " IMO, " & gstrContribuinte & " CON WHERE IMO.PKId = CM.intImobiliario  AND CON.PKId = IMO.intContribuinte ORDER BY CON.strNome "
            Case 4
                strSQL = "SELECT DISTINCT REC.intContribuinte, CON.strNome FROM " & gstrReceitaDiversa & " REC, " & gstrContribuinte & " CON WHERE CON.PKId = REC.intContribuinte ORDER BY CON.strNome "
        End Select
        dbcintContribuinteInicial.Tag = strSQL & ";CON.strNome"
        dbcintContribuinteFinal.Tag = dbcintContribuinteInicial.Tag
        PreencherListaDeOpcoes Me.ActiveControl
        
    End If
    
End Sub

Private Sub LimpaObjetos()
    optbitTipoDeInscricao(0).Value = True
    dbcintContribuinteInicial.BoundText = ""
    dbcintContribuinteFinal.BoundText = ""
    dbcintContribuinteInicial.SetFocus
End Sub

Private Function blnDadosOk() As Boolean
blnDadosOk = False
On Error GoTo err_blnDadosOK
    If chk_Selecionar.Value = 0 Then
        If dbcintContribuinteInicial.MatchedWithList = False Then
           ExibeMensagem "O Contribuinte Incial tem que ser selecionado."
           dbcintContribuinteInicial.SetFocus
           Exit Function
        End If
        If dbcintContribuinteFinal.MatchedWithList = False Then
           ExibeMensagem "O Contribuinte Final tem que ser selecionado."
           dbcintContribuinteFinal.SetFocus
           Exit Function
        End If
    Else
        blnDadosOk = True
    End If
    blnDadosOk = True
err_blnDadosOK:
End Function

Private Function strQuerryRelatorio() As String

'******************************************************************************************
' Data: 02/05/2003
' Alteração: - Substituição da estrutura CASE do SQL Server pela função gstrCASEWHEN.
' Responsável: Everton Bianchini
'------------------------------------------------------------------------------------------
' Data: 05/05/2003
' Alteração: - Substituição do comando nativo de concatenação de strings (+) do SQL Server
'            pela variável strCONCAT.
' Responsável: Everton Bianchini
'******************************************************************************************
    
Dim strSQL As String
    
    strSQL = ""
    strSQL = strSQL & " SELECT A.PKId, A.intContribuinte, "
'    strSql = strSql & " CASE A.bitDefinicao WHEN 0 THEN 'Imunidade' WHEN 1 THEN 'Isenção' WHEN 2 THEN 'Não Incidência' END AS Definicao, "
    strSQL = strSQL & gstrCASEWHEN("A.bitDefinicao", "0, 'Imunidade', 1, 'Isenção', 2, 'Não Incidência'") & " AS Definicao, "
    strSQL = strSQL & " A.intComposicaoDaReceita, A.intReceita, B.strNome, "
'    strSQL = strSQL & " LTRIM(RTRIM(C.strDescricao)) + ' - ' + LTRIM(RTRIM(C.strSigla)) AS ComposicaoDaReceita, "
    strSQL = strSQL & " LTRIM(RTRIM(C.strDescricao)) " & strCONCAT & " ' - ' " & strCONCAT & " LTRIM(RTRIM(C.strSigla)) AS ComposicaoDaReceita, "
'    strSQL = strSQL & " LTRIM(RTRIM(D.strDescricao)) + ' - ' + LTRIM(RTRIM(D.strSigla)) AS Receita "
    strSQL = strSQL & " LTRIM(RTRIM(D.strDescricao)) " & strCONCAT & " ' - ' " & strCONCAT & " LTRIM(RTRIM(D.strSigla)) AS Receita "

    strSQL = strSQL & " FROM "
    strSQL = strSQL & gstrIsencaoImunidade & " A, "
    strSQL = strSQL & gstrContribuinte & " B, "
    strSQL = strSQL & gstrComposicaoDaReceita & " C, "
    strSQL = strSQL & gstrReceita & " D "

    strSQL = strSQL & " WHERE "
    strSQL = strSQL & " A.intContribuinte = B.PKId "
    strSQL = strSQL & " AND A.intComposicaoDaReceita = C.PKId "
    strSQL = strSQL & " AND A.intReceita = D.PKId "
    strSQL = strSQL & " AND A.bitTipoDeInscricao = " & TipoDeInscricao
    If chk_Selecionar.Value = 0 Then
        strSQL = strSQL & " AND A.intContribuinte BETWEEN " & CCInicial & " AND " & CCFinal
    End If
    strSQL = strSQL & " ORDER BY B.strNome "
    
strQuerryRelatorio = strSQL
End Function

Private Sub optbitTipoDeInscricao_Click(Index As Integer)
Dim strSQL As String
Dim intIndice As Integer
    
    TipoDeInscricao = 0
    TipoDeInscricao = Val(Index)
    
    optbitTipoDeInscricao(Index).CausesValidation = True
    
    For intIndice = 0 To 4
        If intIndice <> Index Then
            optbitTipoDeInscricao(intIndice).CausesValidation = False
        End If
    Next
        
    Set dbcintContribuinteInicial.RowSource = Nothing
    dbcintContribuinteInicial.Text = ""
    Set dbcintContribuinteFinal.RowSource = Nothing
    dbcintContribuinteFinal.Text = ""
    
End Sub

Private Sub dbcintContribuinteFinal_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbcintContribuinteFinal
End Sub

Private Sub dbcintContribuinteInicial_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbcintContribuinteInicial
End Sub
