VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmGuiaDeArrecadacao 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Guias de Arrecadação Municipal"
   ClientHeight    =   4845
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6285
   HelpContextID   =   455
   Icon            =   "GuiaDeArrecadacao.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4845
   ScaleWidth      =   6285
   Begin TabDlg.SSTab tab_3dPasta 
      Height          =   4605
      Left            =   157
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   120
      Width           =   5970
      _ExtentX        =   10530
      _ExtentY        =   8123
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Guias de Arrecadação Municipal"
      TabPicture(0)   =   "GuiaDeArrecadacao.frx":1042
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fra_Mensagem1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      Begin VB.Frame Frame1 
         Caption         =   "Mensagem 2                                "
         Height          =   1965
         Left            =   180
         TabIndex        =   8
         Top             =   2490
         Width           =   5595
         Begin VB.CheckBox chk_EmBranco2 
            Caption         =   "Em branco"
            Height          =   195
            Left            =   1350
            TabIndex        =   3
            Top             =   0
            Width           =   1095
         End
         Begin VB.TextBox txt_Mensagem2 
            Height          =   1185
            Left            =   120
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   660
            Width           =   5355
         End
         Begin MSDataListLib.DataCombo dbcintMensagem2 
            Height          =   315
            Left            =   1080
            TabIndex        =   4
            Top             =   270
            Width           =   4395
            _ExtentX        =   7752
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            Text            =   ""
         End
         Begin VB.Label lbl_Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Mensagem"
            Height          =   195
            Left            =   120
            TabIndex        =   9
            Top             =   300
            Width           =   780
         End
      End
      Begin VB.Frame fra_Mensagem1 
         Caption         =   "Mensagem 1                                "
         Height          =   1965
         Left            =   180
         TabIndex        =   7
         Top             =   450
         Width           =   5595
         Begin VB.CheckBox chk_EmBranco1 
            Caption         =   "Em branco"
            Height          =   195
            Left            =   1350
            TabIndex        =   0
            Top             =   0
            Width           =   1095
         End
         Begin VB.TextBox txt_Mensagem1 
            Height          =   1185
            Left            =   120
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   660
            Width           =   5355
         End
         Begin MSDataListLib.DataCombo dbcintMensagem1 
            Height          =   315
            Left            =   1080
            TabIndex        =   1
            Top             =   270
            Width           =   4395
            _ExtentX        =   7752
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            Text            =   ""
         End
         Begin VB.Label lbl_Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Mensagem"
            Height          =   195
            Left            =   120
            TabIndex        =   10
            Top             =   330
            Width           =   780
         End
      End
   End
End
Attribute VB_Name = "frmGuiaDeArrecadacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chk_EmBranco1_Click()
    If chk_EmBranco1.Value = 1 Then
        dbcintMensagem1.BoundText = ""
        dbcintMensagem1.Enabled = False
        TrocaCorObjeto dbcintMensagem1, True
        txt_Mensagem1.Text = ""
        txt_Mensagem1.Enabled = False
        TrocaCorObjeto txt_Mensagem1, True
    Else
        dbcintMensagem1.Enabled = True
        TrocaCorObjeto dbcintMensagem1, False
        txt_Mensagem1.Enabled = True
        TrocaCorObjeto txt_Mensagem1, False
    End If
End Sub

Private Sub chk_EmBranco1_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub chk_EmBranco2_Click()
    If chk_EmBranco2.Value = 1 Then
        dbcintMensagem2.BoundText = ""
        dbcintMensagem2.Enabled = False
        TrocaCorObjeto dbcintMensagem2, True
        txt_Mensagem2.Text = ""
        txt_Mensagem2.Enabled = False
        TrocaCorObjeto txt_Mensagem2, True
    Else
        dbcintMensagem2.Enabled = True
        TrocaCorObjeto dbcintMensagem2, False
        txt_Mensagem2.Enabled = True
        TrocaCorObjeto txt_Mensagem2, False
    End If
End Sub

Private Sub chk_EmBranco2_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub dbcintMensagem1_Click(Area As Integer)
    DropDownDataCombo dbcintMensagem1, Me, Area
    If Area = 2 Then
        LeDoComboParaTXT1
    End If
End Sub

Private Sub dbcintMensagem1_KeyDown(KeyCode As Integer, Shift As Integer)
   DropDownDataCombo dbcintMensagem1, Me, , KeyCode, Shift
End Sub

Private Sub dbcintMensagem2_Click(Area As Integer)
    DropDownDataCombo dbcintMensagem2, Me, Area
    If Area = 2 Then
        LeDoComboParaTXT2
    End If
End Sub

Private Function LeDoComboParaTXT1()
    Dim strSQL          As String
    Dim adoResultado    As ADODB.Recordset
    strSQL = ""
    strSQL = strSQL & " SELECT strMensagem "
    strSQL = strSQL & " FROM " & gstrMensagem
    strSQL = strSQL & " WHERE PKId = " & Val(dbcintMensagem1.BoundText)
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
        If adoResultado.EOF = False Then
            txt_Mensagem1.Text = adoResultado!strMensagem
            adoResultado.MoveNext
        Else
            txt_Mensagem1.Text = ""
        End If
    End If
End Function

Private Function LeDoComboParaTXT2()
    Dim strSQL          As String
    Dim adoResultado    As ADODB.Recordset
    strSQL = ""
    strSQL = strSQL & " SELECT strMensagem "
    strSQL = strSQL & " FROM " & gstrMensagem
    strSQL = strSQL & " WHERE PKId = " & Val(dbcintMensagem2.BoundText)
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
        If adoResultado.EOF = False Then
            txt_Mensagem2.Text = adoResultado!strMensagem
            adoResultado.MoveNext
        Else
            txt_Mensagem2.Text = ""
        End If
    End If
End Function

Private Sub dbcintMensagem2_KeyDown(KeyCode As Integer, Shift As Integer)
   DropDownDataCombo dbcintMensagem2, Me, , KeyCode, Shift
End Sub

Private Sub Form_Activate()
    gintCodSeguranca = 455
    HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrSalvar, gstrDeletar, gstrAplicar
End Sub

Private Sub Form_Deactivate()
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar, gstrAplicar
End Sub

Private Sub Form_Load()
    LeDaTabelaParaObj gstrMensagem, dbcintMensagem1, strQueryMensagem
    LeDaTabelaParaObj gstrMensagem, dbcintMensagem2, strQueryMensagem
End Sub

Private Function strQueryMensagem() As String

'******************************************************************************************
' Data: 06/03/2003
' Alteração: - Substituição dos comandos nativos do SQL Server pelas variáveis de comandos
'            nativos.
' Responsável: Everton Bianchini
'******************************************************************************************
    
    Dim strSQL As String
    strSQL = ""
'    strSql = strSql & "SELECT PKId, ltrim(rtrim(PKId)) + ' - ' + ltrim(rtrim(strDescricao)) as Descricao "
    strSQL = strSQL & "SELECT PKId, ltrim(rtrim(PKId))" & strCONCAT & "' - '" & strCONCAT & "ltrim(rtrim(strDescricao)) as Descricao "
    strSQL = strSQL & " FROM " & gstrMensagem
    strSQL = strSQL & " ORDER BY PKId "
    strQueryMensagem = strSQL
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrSalvar, gstrDeletar, gstrNovo
End Sub

Private Function blnDadosOk() As Boolean
    If chk_EmBranco1.Value = 0 Then
        If txt_Mensagem1.Text = "" Then
            ExibeMensagem "A mensagem 1 tem que ser selecionada"
            blnDadosOk = False
            Exit Function
        End If
    End If
    If chk_EmBranco2.Value = 0 Then
        If txt_Mensagem2.Text = "" Then
            ExibeMensagem "A mensagem 2 tem que ser selecionada"
            blnDadosOk = False
            Exit Function
        End If
    End If
    blnDadosOk = True
End Function

Public Sub MantemForm(ByVal strModoOperacao As String)

    If UCase(strModoOperacao) = UCase(gstrImprimir) Then
        If blnDadosOk = False Then
            Exit Sub
        End If
        
        'BY Power 17/01/03
        ImprimeRelatorio rptGuiaDeArrecadacaoMunicipal, strQueryRelatorio
        
    End If
   
    'BY Power 17/01/03
    If UCase(strModoOperacao) = UCase(gstrNovo) Then
        LimpaObjetos
    End If
    
    If UCase(strModoOperacao) = UCase(gstrFechar) Then
        Unload Me
    End If

'    ToolBarGeral strModoOperacao, "", False, , , , , , _
'                 rptGuiaDeArrecadacaoMunicipal, strQueryRelatorio
    
End Sub

Private Sub LimpaObjetos()
    dbcintMensagem1.BoundText = ""
    dbcintMensagem2.BoundText = ""
    txt_Mensagem1 = ""
    txt_Mensagem2 = ""
    dbcintMensagem1.SetFocus
End Sub

Private Function strQueryRelatorio() As String

'******************************************************************************************
' Data: 05/03/2003
' Alteração: - Substituição dos comandos nativos do SQL Server pelas variáveis de comandos
'            nativos.
' Responsável: Everton Bianchini
'------------------------------------------------------------------------------------------
' Data: 06/03/2003
' Alteração: - Adaptação dos outer joins.
' Responsável: Everton Bianchini
'------------------------------------------------------------------------------------------
' Data: 26/03/2003
' Alteração: - Substituição da variável strISNULL pela função gstrISNULL. Ver definição da
'            função para maiores esclarecimentos.
' Responsável: Everton Bianchini
'******************************************************************************************
    
Dim strSQL As String
    strSQL = ""
    strSQL = strSQL & " SELECT DISTINCT "
    strSQL = strSQL & " I.PKId, I.intExercicio, J.intNumeroParcela, "
    strSQL = strSQL & " I.strInscricaoCadastral, I.intComposicaoReceita CodReceita,"
    strSQL = strSQL & " J.dtmDataVencimento, G.strDescricao Municipio, "
    strSQL = strSQL & " H.strDescricao Bairro,C.strNome Contribuinte,"
    strSQL = strSQL & " C.intNumero, C.strComplemento, C.intCep,"
'    strSql = strSql & " LTRIM(RTRIM(ISNULL(E.strSigla,''))) + ' ' + "
'    strSQL = strSQL & " LTRIM(RTRIM(" & strISNULL & "(E.strSigla,'')))" & strCONCAT & "' '" & strCONCAT
    strSQL = strSQL & " LTRIM(RTRIM(" & gstrISNULL("E.strSigla", "''") & "))" & strCONCAT & "' '" & strCONCAT
'    strSql = strSql & " LTRIM(RTRIM(ISNULL(F.strSigla,''))) + ' ' + "
'    strSQL = strSQL & " LTRIM(RTRIM(" & strISNULL & "(F.strSigla,'')))" & strCONCAT & "' '" & strCONCAT
    strSQL = strSQL & " LTRIM(RTRIM(" & gstrISNULL("F.strSigla", "''") & "))" & strCONCAT & "' '" & strCONCAT
    strSQL = strSQL & " LTRIM(RTRIM(D.strDescricao)) AS Logradouro, "
    strSQL = strSQL & " J.PKId PKIdParcelaReceita "
    
    strSQL = strSQL & " FROM "
    strSQL = strSQL & gstrContribuinte & " C,"
    strSQL = strSQL & gstrLogradouro & " D,"
    strSQL = strSQL & gstrTipoLogradouro & " E,"
    strSQL = strSQL & gstrTituloLogradouro & " F,"
    strSQL = strSQL & gstrCidade & " G,"
    strSQL = strSQL & gstrBairro & " H, "
    strSQL = strSQL & gstrLancamentoCalculo & " I, "
    strSQL = strSQL & gstrParcelaReceita & " J "
    
    strSQL = strSQL & " WHERE "
    strSQL = strSQL & " J.intLancamentoCalculo = I.PKId"
    strSQL = strSQL & " AND I.bitUtilizacaoDebito = 4 "
    strSQL = strSQL & " AND I.intContribuinte = C.PKId "
'    strSql = strSql & " AND C.intMunicipio *= G.PKId "
    strSQL = strSQL & " AND C.intMunicipio " & strOUTJSQLServer & "= G.PKId " & strOUTJOracle
'    strSql = strSql & " AND C.intBairro *= H.PKId "
    strSQL = strSQL & " AND C.intBairro " & strOUTJSQLServer & "= H.PKId " & strOUTJOracle
    strSQL = strSQL & " AND C.intLogradouro = D.PKId "
'    strSql = strSql & " AND D.intTipoLogradouro *= E.PKId "
    strSQL = strSQL & " AND D.intTipoLogradouro " & strOUTJSQLServer & "= E.PKId " & strOUTJOracle
'    strSql = strSql & " AND D.intTituloLogradouro *= F.PKId "
    strSQL = strSQL & " AND D.intTituloLogradouro " & strOUTJSQLServer & "= F.PKId " & strOUTJOracle
    
strQueryRelatorio = strSQL
End Function

Private Sub dbcintMensagem1_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub dbcintMensagem2_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub tab_3dPasta_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub txt_Mensagem1_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub txt_Mensagem2_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub
