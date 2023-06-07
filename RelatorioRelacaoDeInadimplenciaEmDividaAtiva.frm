VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmRelatorioRelacaoDeInadimplenciaEmDividaAtiva 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Relação de Inadimplência em Dívida Ativa"
   ClientHeight    =   2325
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7545
   Icon            =   "RelatorioRelacaoDeInadimplenciaEmDividaAtiva.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2325
   ScaleWidth      =   7545
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab tab_3dPasta 
      Height          =   1995
      Left            =   180
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   180
      Width           =   7155
      _ExtentX        =   12621
      _ExtentY        =   3519
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Relação de Inadimplência em Dívida Ativa"
      TabPicture(0)   =   "RelatorioRelacaoDeInadimplenciaEmDividaAtiva.frx":1042
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin VB.Frame Frame1 
         Height          =   1395
         Left            =   180
         TabIndex        =   1
         Top             =   390
         Width           =   6795
         Begin VB.CheckBox chk_Selecionar 
            Caption         =   "Selecionar todos os Contribuintes"
            Height          =   255
            Left            =   1590
            TabIndex        =   2
            Top             =   1020
            Width           =   2835
         End
         Begin MSDataListLib.DataCombo dbcintContribuinteInicial 
            Height          =   315
            Left            =   1590
            TabIndex        =   3
            Top             =   240
            Width           =   5055
            _ExtentX        =   8916
            _ExtentY        =   556
            _Version        =   393216
            IntegralHeight  =   0   'False
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dbcintContribuinteFinal 
            Height          =   315
            Left            =   1590
            TabIndex        =   4
            Top             =   660
            Width           =   5055
            _ExtentX        =   8916
            _ExtentY        =   556
            _Version        =   393216
            IntegralHeight  =   0   'False
            Text            =   ""
         End
         Begin VB.Label lbl_Label1 
            AutoSize        =   -1  'True
            Caption         =   "Contribuinte Inicial"
            Height          =   195
            Left            =   210
            TabIndex        =   6
            Top             =   345
            Width           =   1290
         End
         Begin VB.Label lbl_label2 
            AutoSize        =   -1  'True
            Caption         =   "Contribuinte Final"
            Height          =   195
            Left            =   285
            TabIndex        =   5
            Top             =   765
            Width           =   1215
         End
      End
   End
End
Attribute VB_Name = "frmRelatorioRelacaoDeInadimplenciaEmDividaAtiva"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mblnAlterando               As Boolean
Dim mobjAux                     As Object
Dim mblnSelecionou              As Boolean
Dim mblnPrimeiraVez             As Boolean
Dim intCodigoInicial            As Integer
Dim intCodigoFinal              As Integer
Dim CCInicial                   As Integer
Dim CCFinal                     As Integer
Dim TipoDeInscricao             As Integer

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
    dbcintContribuinteInicial.Tag = strQueryContribuinte & ";strNome"
    dbcintContribuinteFinal.Tag = dbcintContribuinteInicial.Tag
    HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrSalvar, gstrDeletar
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar
    mblnSelecionou = False
    mblnPrimeiraVez = False
End Sub

Public Sub MantemForm(ByVal strModoOperacao As String)
On Error Resume Next
Dim strSql As String
Dim Resultado As String
Dim adoResultado   As ADODB.Recordset

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
            ImprimeRelatorio rptRelacaoDeInadimplenciaEmDividaAtiva, strQueryRelatorio
        End If
    ElseIf UCase(strModoOperacao) = UCase(gstrNovo) Then
        LimpaObjetos
    ElseIf UCase(strModoOperacao) = UCase(gstrFechar) Then
        Unload Me
    ElseIf UCase(strModoOperacao) = UCase(gstrPreencherLista) Then
        PreencherListaDeOpcoes Me.ActiveControl
    End If
    
End Sub

Private Sub LimpaObjetos()
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

Private Function strQueryContribuinte() As String
Dim strSql As String
    strSql = ""
    strSql = strSql & " SELECT "
    strSql = strSql & " EC.intContribuinte, CO.strNome "
    strSql = strSql & " FROM "
    strSql = strSql & gstrEconomico & " EC, "
    strSql = strSql & gstrContribuinte & " CO "
    strSql = strSql & " WHERE "
    strSql = strSql & " EC.intContribuinte = CO.PKId "
    strSql = strSql & " ORDER BY CO.strNome "
strQueryContribuinte = strSql
End Function

Private Function strQueryRelatorio() As String

'******************************************************************************************
' Data: 05/05/2003
' Alteração: - Retirado comando "AS" utilizado para dar apelidos às tabelas, o qual não é
'            permitido pelo Oracle, da cláusula FROM.
' Responsável: Everton Bianchini
'******************************************************************************************

Dim strSql As String
    
    strSql = ""
    strSql = strSql & " SELECT DA.intContribuinte, CO.strNome, SUM(DDA.dblValorOriginal) as ValorDevido "
    strSql = strSql & " FROM "
'    strSql = strSql & gstrDividaAtiva & " AS DA, "
    strSql = strSql & gstrDividaAtiva & " DA, "
'    strSql = strSql & gstrDetalheDividaAtiva & " AS DDA, "
    strSql = strSql & gstrDetalheDividaAtiva & " DDA, "
'    strSql = strSql & gstrContribuinte & " AS CO "
    strSql = strSql & gstrContribuinte & " CO "
    strSql = strSql & " WHERE "
    strSql = strSql & " DA.PKId = DDA.intDividaAtiva "
    strSql = strSql & " AND DA.intContribuinte = CO.PKId "
    If chk_Selecionar.Value = 0 Then
        strSql = strSql & " AND DA.intContribuinte BETWEEN " & CCInicial & " AND " & CCFinal
    End If
    strSql = strSql & " AND DDA.bytSituacao <> 0 " 'Parcial ou Em Aberto
    strSql = strSql & " GROUP BY DA.intContribuinte, CO.strNome "
    strSql = strSql & " ORDER BY CO.strNome "
    
strQueryRelatorio = strSql
End Function

Private Sub dbcintContribuinteFinal_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbcintContribuinteFinal
End Sub

Private Sub dbcintContribuinteInicial_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbcintContribuinteInicial
End Sub



