VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmDocumentosDiversos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Documentos Diversos"
   ClientHeight    =   4860
   ClientLeft      =   705
   ClientTop       =   3405
   ClientWidth     =   8655
   Icon            =   "DocumentosDiversos.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4860
   ScaleWidth      =   8655
   Begin TabDlg.SSTab tab_3dPasta 
      Height          =   4605
      Left            =   90
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   150
      Width           =   8475
      _ExtentX        =   14949
      _ExtentY        =   8123
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Documentos Diversos"
      TabPicture(0)   =   "DocumentosDiversos.frx":1042
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lbl_InscricaoInicial"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lbl_ComposicaoDaReceita"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "dbc_intDocumentosEmitidos"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "dbc_strInscricaoInicial"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "fra_Inscricao"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "chk_SelecionarTudo"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Frame2"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmd_Documento"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      Begin VB.CommandButton cmd_Documento 
         Height          =   300
         Left            =   7980
         Picture         =   "DocumentosDiversos.frx":105E
         Style           =   1  'Graphical
         TabIndex        =   14
         TabStop         =   0   'False
         ToolTipText     =   "Ativa Cadastro de Documentos"
         Top             =   1800
         Width           =   330
      End
      Begin VB.Frame Frame2 
         Caption         =   " Texto a ser impresso "
         Height          =   2235
         Left            =   120
         TabIndex        =   13
         Top             =   2220
         Width           =   8205
         Begin VB.TextBox txt_strTexto 
            Height          =   1935
            Left            =   60
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            OLEDragMode     =   1  'Automatic
            OLEDropMode     =   1  'Manual
            ScrollBars      =   2  'Vertical
            TabIndex        =   8
            Top             =   210
            Width           =   8055
         End
      End
      Begin VB.CheckBox chk_SelecionarTudo 
         Caption         =   "Selecionar todos"
         Height          =   255
         Left            =   1710
         TabIndex        =   6
         Top             =   1500
         Width           =   1725
      End
      Begin VB.Frame fra_Inscricao 
         Caption         =   " Tipo de Inscrição "
         Height          =   555
         Left            =   150
         TabIndex        =   10
         Top             =   480
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
            TabIndex        =   1
            Top             =   270
            Width           =   1425
         End
         Begin VB.OptionButton optbitTipoDeInscricao 
            Caption         =   "Econômico"
            CausesValidation=   0   'False
            Height          =   195
            Index           =   2
            Left            =   3150
            TabIndex        =   2
            Top             =   270
            Width           =   1155
         End
         Begin VB.OptionButton optbitTipoDeInscricao 
            Caption         =   "Contribuição de Melhorias"
            CausesValidation=   0   'False
            Height          =   195
            Index           =   3
            Left            =   4290
            TabIndex        =   3
            Top             =   270
            Width           =   2205
         End
         Begin VB.OptionButton optbitTipoDeInscricao 
            Caption         =   "Receitas Diversas"
            CausesValidation=   0   'False
            Height          =   195
            Index           =   4
            Left            =   6480
            TabIndex        =   4
            Top             =   270
            Width           =   1605
         End
      End
      Begin MSDataListLib.DataCombo dbc_strInscricaoInicial 
         Height          =   315
         Left            =   1710
         TabIndex        =   5
         Top             =   1140
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   556
         _Version        =   393216
         IntegralHeight  =   0   'False
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dbc_intDocumentosEmitidos 
         Height          =   315
         Left            =   1710
         TabIndex        =   7
         Top             =   1800
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   556
         _Version        =   393216
         IntegralHeight  =   0   'False
         Style           =   2
         Text            =   ""
      End
      Begin VB.Label lbl_ComposicaoDaReceita 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Documentos Emitidos"
         Height          =   195
         Left            =   90
         TabIndex        =   12
         Top             =   1920
         Width           =   1530
      End
      Begin VB.Label lbl_InscricaoInicial 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Inscrição Cadastral"
         Height          =   195
         Left            =   270
         TabIndex        =   11
         Top             =   1260
         Width           =   1350
      End
   End
End
Attribute VB_Name = "frmDocumentosDiversos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim mblnAlterando    As Boolean
    Dim mobjAux          As Object
    Dim mblnPrimeiraVez  As Boolean
    Dim TipoDeInscricao  As Integer
    Dim intCodSeguranca  As Integer
    
Private Sub chk_SelecionarTudo_Click()
    If chk_SelecionarTudo.Value = 1 Then
        dbc_strInscricaoInicial.BoundText = ""
        dbc_strInscricaoInicial.Enabled = False
        TrocaCorObjeto dbc_strInscricaoInicial, True
    Else
        dbc_strInscricaoInicial.Enabled = True
        TrocaCorObjeto dbc_strInscricaoInicial, False
    End If
End Sub

Private Sub chk_SelecionarTudo_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub cmd_Documento_Click()
    CarregaForm frmCadDocumentoEmitido, dbc_intDocumentosEmitidos
End Sub

Private Sub dbc_intDocumentosEmitidos_Click(Area As Integer)
    DropDownDataCombo dbc_intDocumentosEmitidos, Me, Area
    If Area = 2 And dbc_intDocumentosEmitidos.MatchedWithList = True Then
        MostraTexto
    End If
End Sub

Private Sub dbc_intDocumentosEmitidos_KeyDown(KeyCode As Integer, Shift As Integer)
   DropDownDataCombo dbc_intDocumentosEmitidos, Me, , KeyCode, Shift
End Sub

Private Sub dbc_intDocumentosEmitidos_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub dbc_strInscricaoInicial_Click(Area As Integer)
    DropDownDataCombo dbc_strInscricaoInicial, Me, Area
End Sub

Private Sub dbc_strInscricaoInicial_KeyDown(KeyCode As Integer, Shift As Integer)
   DropDownDataCombo dbc_strInscricaoInicial, Me, , KeyCode, Shift
End Sub

Private Sub Form_Activate()
    gintCodSeguranca = intCodSeguranca
    Me.HelpContextID = intCodSeguranca
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrSalvar, gstrAplicar, gstrDeletar, gstrImprimir
    HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrNovo, gstrFechar, gstrImprimir
End Sub

Private Sub Form_Deactivate()
    HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrSalvar, gstrAplicar, gstrDeletar, gstrImprimir
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrNovo, gstrFechar, gstrImprimir
End Sub

Private Sub Form_Load()
    TrocaCorObjeto txt_strTexto, True
    intCodSeguranca = gintCodSeguranca
    'LeDaTabelaParaObj gstrEconomico, dbc_strInscricaoInicial
    LeDaTabelaParaObj gstrDocumentoEmitido, dbc_intDocumentosEmitidos, "SELECT PKId, strDescricao FROM " & gstrDocumentoEmitido & " ORDER BY strDescricao "
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrSalvar, gstrAplicar, gstrDeletar, gstrImprimir
    HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrNovo, gstrFechar, gstrImprimir
End Sub

Private Function blnValidaDados() As Boolean
    If chk_SelecionarTudo.Value = 0 And dbc_strInscricaoInicial.MatchedWithList = False Then
        ExibeMensagem "A Inscrição Inicial tem que ser selecionada."
        dbc_strInscricaoInicial.SetFocus
    ElseIf dbc_strInscricaoInicial.MatchedWithList = False And chk_SelecionarTudo = 0 Then
        ExibeMensagem "A Inscrição Inicial tem que ser selecionada."
        dbc_strInscricaoInicial.SetFocus
    ElseIf dbc_intDocumentosEmitidos.MatchedWithList = False Then
        ExibeMensagem "O documento emitido tem que ser selecionado."
        dbc_intDocumentosEmitidos.SetFocus
    Else
        blnValidaDados = True
    End If
End Function

Public Sub MantemForm(ByVal strModoOperacao As String)
Dim i As Integer
Dim j As Integer
Dim strSql As String

If UCase(strModoOperacao) = UCase(gstrImprimir) Then
    If blnValidaDados Then
        ImprimeRelatorio rptDocumentoDiverso, strQuerryRelatorio
    End If
ElseIf UCase(strModoOperacao) = UCase(gstrNovo) Then
    LimpaControlesDoFormulario
ElseIf UCase(strModoOperacao) = UCase(gstrFechar) Then
    Unload Me
ElseIf UCase(strModoOperacao) = UCase(gstrPreencherLista) Then
    For i = 0 To optbitTipoDeInscricao.Count - 1
        If optbitTipoDeInscricao(i).Value Then
            j = i
            Exit For
        End If
    Next i
    
    If j < 4 Then
        strSql = strQueryInscricao(j)
    Else
        strSql = "SELECT DISTINCT REC.intContribuinte, CON.strNome FROM " & gstrReceitaDiversa & " REC, " & gstrContribuinte & " CON WHERE CON.PKId = REC.intContribuinte ORDER BY CON.strNome;strNome"
    End If
    
    dbc_strInscricaoInicial.Tag = strSql & ";strNome"
    PreencherListaDeOpcoes dbc_strInscricaoInicial
End If
End Sub

Sub LimpaControlesDoFormulario()
    optbitTipoDeInscricao(0).Value = True
    dbc_strInscricaoInicial.BoundText = ""
    dbc_intDocumentosEmitidos.BoundText = ""
    txt_strTexto.Text = ""
    chk_SelecionarTudo.Value = 0
    optbitTipoDeInscricao(0).SetFocus
End Sub

Private Sub optbitTipoDeInscricao_Click(Index As Integer)
    Dim strSql As String
    Dim intIndice As Integer
    TipoDeInscricao = 0
    TipoDeInscricao = Val(Index)
    optbitTipoDeInscricao(Index).CausesValidation = True
    
    lbl_InscricaoInicial.Caption = "Inscrição Cadastral"
    
    For intIndice = 0 To 4
        If intIndice <> Index Then
            optbitTipoDeInscricao(intIndice).CausesValidation = False
        End If
        
        If optbitTipoDeInscricao(4).Value Then
            lbl_InscricaoInicial.Caption = "Contribuinte"
        Else
            
        End If
    Next
    
    Set dbc_strInscricaoInicial.RowSource = Nothing
    dbc_strInscricaoInicial.Text = ""
    
    dbc_intDocumentosEmitidos.BoundText = 0
    txt_strTexto.Text = ""
End Sub

Private Function strQueryInscricao(Index As Integer) As String

'******************************************************************************************
' Data: 04/03/2003
' Alteração: - Substituição do comando nativo de concatenação de strings (+) do SQL Server
'        pela variável strCONCAT.
' Responsável: Everton Bianchini
'******************************************************************************************
    
    Dim strSql As String
    strSql = ""
    If Index = 0 Or Index = 1 Then
'        strSql = strSql & " SELECT A.strInscricaoAnterior, LTRIM(RTRIM(A.strInscricaoAnterior)) + ' - ' +  LTRIM(RTRIM(B.strNome)) AS Descricao "
        strSql = strSql & " SELECT A.strInscricaoAnterior, LTRIM(RTRIM(A.strInscricaoAnterior)) " & strCONCAT & " ' - ' " & strCONCAT & " LTRIM(RTRIM(B.strNome)) AS Descricao "
    ElseIf Index = 2 Then
'        strSql = strSql & " SELECT A.strInscricaoCadastral, LTRIM(RTRIM(A.strInscricaoCadastral)) + ' - ' +  LTRIM(RTRIM(B.strNome)) AS Descricao "
        strSql = strSql & " SELECT A.strInscricaoCadastral, LTRIM(RTRIM(A.strInscricaoCadastral)) " & strCONCAT & " ' - ' " & strCONCAT & " LTRIM(RTRIM(B.strNome)) AS Descricao "
    ElseIf Index = 3 Then
'        strSql = strSql & " SELECT A.strInscricaoAnterior, LTRIM(RTRIM(A.strInscricaoAnterior)) + ' - ' +  LTRIM(RTRIM(C.strNome)) AS Descricao "
        strSql = strSql & " SELECT A.strInscricaoAnterior, LTRIM(RTRIM(A.strInscricaoAnterior)) " & strCONCAT & " ' - ' " & strCONCAT & " LTRIM(RTRIM(C.strNome)) AS Descricao "
    End If
    strSql = strSql & " FROM "
    If Index = 0 Then
        strSql = strSql & gstrImobiliario & " A, "
        strSql = strSql & gstrContribuinte & " B "
    ElseIf Index = 1 Then
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
        strSql = strSql & " ORDER BY Descricao "
    ElseIf Index = 2 Then
        strSql = strSql & " A.intContribuinte = B.PKId "
        strSql = strSql & " ORDER BY Descricao "
    ElseIf Index = 3 Then
        strSql = strSql & " B.intImobiliario = A.PKId "
        strSql = strSql & " AND A.intContribuinte = C.PKId "
        strSql = strSql & " ORDER BY Descricao "
    End If
    strQueryInscricao = strSql
End Function

Private Sub dbc_strInscricaoInicial_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub MostraTexto()
    Dim strSql          As String
    Dim adoResultado    As ADODB.Recordset
    strSql = ""
    strSql = strSql & " SELECT strTexto "
    strSql = strSql & " FROM "
    strSql = strSql & gstrDocumentoEmitido
    strSql = strSql & " WHERE "
    strSql = strSql & " PKId = " & Val(dbc_intDocumentosEmitidos.BoundText)
    Set gobjBanco = New clsBanco
    gobjBanco.CriaADO strSql, 5, adoResultado
    With adoResultado
        If Not .EOF Then
            txt_strTexto.Text = gstrENulo(!strTexto)
        End If
    End With
End Sub

Private Function strQuerryRelatorio() As String
    Dim strSql  As String
    strSql = ""
    strSql = strSql & " SELECT "
    If TipoDeInscricao = 4 Then
        strSql = strSql & "  B.PKId AS intContribuinte, A.strDescricao, A.strTexto "
    Else
        strSql = strSql & "  B.intContribuinte, A.strDescricao, A.strTexto "
    End If
    strSql = strSql & " FROM "
    strSql = strSql & gstrDocumentoEmitido & " A, "
    If TipoDeInscricao = 0 Then
        strSql = strSql & gstrImobiliario & " B "
    ElseIf TipoDeInscricao = 1 Then
        strSql = strSql & gstrImobiliarioRural & " B "
    ElseIf TipoDeInscricao = 2 Then
        strSql = strSql & gstrEconomico & " B "
    ElseIf TipoDeInscricao = 3 Then
        strSql = strSql & gstrImobiliario & " B, "
        strSql = strSql & gstrContribuicaoMelhoria & " C "
    ElseIf TipoDeInscricao = 4 Then
        strSql = strSql & gstrContribuinte & " B "
    End If
    strSql = strSql & " WHERE "
    If TipoDeInscricao = 3 Then
        strSql = strSql & " C.intImobiliario = B.PKId "
        If chk_SelecionarTudo.Value = 0 Then
            strSql = strSql & " AND B.strInscricaoAnterior = '" & dbc_strInscricaoInicial.BoundText & "'  AND "
        End If
    End If
    If TipoDeInscricao = 2 Then
        If chk_SelecionarTudo.Value = 0 Then
            strSql = strSql & " B.strInscricaoCadastral = '" & dbc_strInscricaoInicial.BoundText & "'  AND "
        End If
    End If
    If TipoDeInscricao = 0 Or TipoDeInscricao = 1 Then
        If chk_SelecionarTudo.Value = 0 Then
            strSql = strSql & " B.strInscricaoAnterior = '" & dbc_strInscricaoInicial.BoundText & "'  AND "
        End If
    End If
    If TipoDeInscricao = 4 Then
        If chk_SelecionarTudo.Value = 0 Then
            strSql = strSql & " B.PKId = " & Val(dbc_strInscricaoInicial.BoundText) & " AND "
        End If
    End If
    strSql = strSql & " A.PKId = " & Val(dbc_intDocumentosEmitidos.BoundText)
    strSql = strSql & " ORDER BY A.strDescricao "
    strQuerryRelatorio = strSql
End Function

Private Sub optbitTipoDeInscricao_KeyPress(Index As Integer, KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub tab_3dPasta_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub
