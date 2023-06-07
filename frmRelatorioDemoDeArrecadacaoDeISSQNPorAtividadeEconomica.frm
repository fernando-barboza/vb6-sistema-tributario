VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmRelatorioDemoDeArrecadacaoDeISSQNPorAtividadeEconomica 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Demonstrativo de Arrecadação de ISSQN por Atividade Econômica"
   ClientHeight    =   2355
   ClientLeft      =   510
   ClientTop       =   465
   ClientWidth     =   6735
   Icon            =   "frmRelatorioDemoDeArrecadacaoDeISSQNPorAtividadeEconomica.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2355
   ScaleWidth      =   6735
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab tab_3dPasta 
      Height          =   2055
      Left            =   150
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   180
      Width           =   6480
      _ExtentX        =   11430
      _ExtentY        =   3625
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Demonstrativo de Arrecadação de ISSQN por Atividade Econômica"
      TabPicture(0)   =   "frmRelatorioDemoDeArrecadacaoDeISSQNPorAtividadeEconomica.frx":1042
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin VB.Frame Frame1 
         Height          =   1485
         Left            =   150
         TabIndex        =   1
         Top             =   390
         Width           =   6165
         Begin VB.TextBox txt_dtmInicial 
            Height          =   285
            Left            =   2130
            MaxLength       =   15
            TabIndex        =   3
            Top             =   1035
            Width           =   1035
         End
         Begin VB.TextBox txt_dtmFinal 
            Height          =   285
            Left            =   4980
            MaxLength       =   15
            TabIndex        =   2
            Top             =   1050
            Width           =   1035
         End
         Begin MSDataListLib.DataCombo dbc_AtividadeFinal 
            Height          =   315
            Left            =   2130
            TabIndex        =   4
            Top             =   630
            Width           =   3885
            _ExtentX        =   6853
            _ExtentY        =   556
            _Version        =   393216
            IntegralHeight  =   0   'False
            Style           =   2
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dbc_AtividadeInicial 
            Height          =   315
            Left            =   2130
            TabIndex        =   5
            Top             =   240
            Width           =   3885
            _ExtentX        =   6853
            _ExtentY        =   556
            _Version        =   393216
            IntegralHeight  =   0   'False
            Style           =   2
            Text            =   ""
         End
         Begin VB.Label lbl_AtividadeFinal 
            AutoSize        =   -1  'True
            Caption         =   "Atividade Econômica Final"
            Height          =   195
            Left            =   135
            TabIndex        =   9
            Top             =   720
            Width           =   1875
         End
         Begin VB.Label lbl_AtividadeInicial 
            AutoSize        =   -1  'True
            Caption         =   "Atividade Econômica Inicial"
            Height          =   195
            Left            =   60
            TabIndex        =   8
            Top             =   330
            Width           =   1950
         End
         Begin VB.Label lbl_dtmInicial 
            AutoSize        =   -1  'True
            Caption         =   "Data Inicial"
            Height          =   195
            Left            =   1215
            TabIndex        =   7
            Top             =   1125
            Width           =   795
         End
         Begin VB.Label lbl_dtmFinal 
            AutoSize        =   -1  'True
            Caption         =   "Data Final"
            Height          =   195
            Left            =   4140
            TabIndex        =   6
            Top             =   1140
            Width           =   720
         End
      End
   End
End
Attribute VB_Name = "frmRelatorioDemoDeArrecadacaoDeISSQNPorAtividadeEconomica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub dbc_AtividadeFinal_Click(Area As Integer)
    DropDownDataCombo dbc_AtividadeFinal, Me, Area
End Sub

Private Sub dbc_AtividadeFinal_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbc_AtividadeFinal, Me, , KeyCode, Shift
End Sub

Private Sub dbc_AtividadeInicial_Click(Area As Integer)
    DropDownDataCombo dbc_AtividadeInicial, Me, Area
End Sub

Private Sub dbc_AtividadeInicial_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbc_AtividadeInicial, Me, , KeyCode, Shift
End Sub

Private Sub Form_Activate()
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrDeletar
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrSalvar
End Sub

Private Sub Form_Deactivate()
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar
End Sub

Private Sub Form_Load()
    LeDaTabelaParaObj gstrEconomico, dbc_AtividadeInicial, strQueryAtividade
    dbc_AtividadeFinal.BoundColumn = dbc_AtividadeInicial.BoundColumn
    dbc_AtividadeFinal.ListField = dbc_AtividadeInicial.ListField
    Set dbc_AtividadeFinal.RowSource = dbc_AtividadeInicial.RowSource
    
    txt_dtmFinal.Text = gstrDataDoSistema
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar
End Sub

Public Sub MantemForm(ByVal strModoOperacao As String)
    
    If UCase(strModoOperacao) = UCase(gstrImprimir) Then
        If blnDadosOk Then
            ImprimeRelatorio rptRelatorioDemonstrativoDeArrecadacaoDeISSQNPorAtividadeEconomica, strQueryRelatorio
        End If
    End If
    
    If UCase(strModoOperacao) = UCase(gstrNovo) Then
        LimpaObjetos
    End If
    
    If UCase(strModoOperacao) = UCase(gstrFechar) Then
        Unload Me
    End If


End Sub

Private Function blnDadosOk() As Boolean
blnDadosOk = False
On Error GoTo err_blnDadosOK
    If dbc_AtividadeInicial.BoundText = "" Then
        ExibeMensagem " O campo " & lbl_AtividadeInicial.Caption & " não pode ser nulo."
        dbc_AtividadeInicial.SetFocus
        Exit Function
    End If
    If dbc_AtividadeFinal.BoundText = "" Then
        ExibeMensagem " O campo " & lbl_AtividadeFinal.Caption & " não pode ser nulo."
        dbc_AtividadeFinal.SetFocus
        Exit Function
        End If
    If dbc_AtividadeInicial.BoundText > dbc_AtividadeFinal.BoundText Then
        ExibeMensagem " A Atividade Inicial não pode ser superior a Atividade Final."
        dbc_AtividadeInicial.SetFocus
        Exit Function
        End If
    If gblnDataValida(txt_dtmInicial.Text) = False Then
        ExibeMensagem "A data inicial não é uma data válida."
        txt_dtmInicial.SetFocus
        Exit Function
        End If
    If gblnDataValida(txt_dtmFinal.Text) = False Then
        ExibeMensagem "A data final não é uma data válida."
        txt_dtmFinal.SetFocus
        Exit Function
    End If
    If CVDate(txt_dtmInicial.Text) > CVDate(txt_dtmFinal.Text) Then
        ExibeMensagem " A data Inicial tem que ser anterior a data Final."
        txt_dtmInicial.SetFocus
        Exit Function
        End If
    blnDadosOk = True
err_blnDadosOK:
End Function

Private Sub LimpaObjetos()
    dbc_AtividadeInicial.BoundText = ""
    dbc_AtividadeFinal.BoundText = ""
    txt_dtmInicial.Text = ""
    txt_dtmFinal.Text = ""
    dbc_AtividadeInicial.SetFocus
End Sub

Private Function strQueryAtividade() As String

'******************************************************************************************
' Data: 02/05/2003
' Alteração: - Substituição do comando nativo CONVERT do SQL Server pela função gstrCONVERT
' Responsável: Everton Bianchini
'------------------------------------------------------------------------------------------
' Data: 05/05/2003
' Alteração: - Substituição do comando nativo de concatenação de strings (+) do SQL Server
'            pela variável strCONCAT.
' Responsável: Everton Bianchini
'------------------------------------------------------------------------------------------
' Data: 05/05/2003
' Alteração: - Retirado comando "AS" utilizado para dar apelidos às tabelas, o qual não é
'            permitido pelo Oracle, da cláusula FROM.
' Responsável: Everton Bianchini
'******************************************************************************************
    
Dim strSql As String
    
    strSql = ""
'    strSql = strSql & " SELECT DISTINCT AEC.intCodigo, CONVERT(VARCHAR, AEC.intCodigo) + ' - ' + AEC.strDescricao AS strAtividadePrincipal "
    strSql = strSql & " SELECT DISTINCT AEC.intCodigo, " & gstrCONVERT(CDT_VARCHAR, "AEC.intCodigo") & strCONCAT & " ' - ' " & strCONCAT & " AEC.strDescricao AS strAtividadePrincipal "
    strSql = strSql & " FROM "
'    strSql = strSql & gstrAtividadeDaEmpresa & " AS AE, "
    strSql = strSql & gstrAtividadeDaEmpresa & " AE, "
'    strSql = strSql & gstrAtividadeEC & " AS AEC "
    strSql = strSql & gstrAtividadeEC & " AEC "
    strSql = strSql & " WHERE "
    strSql = strSql & " AEC.PKId = AE.intAtividade "
    strSql = strSql & " AND AE.blnPrincipal = 1 "

strQueryAtividade = strSql
End Function

Private Function strQueryRelatorio() As String

'******************************************************************************************
' Data: 02/05/2003
' Alteração: - Substituição do comando nativo CONVERT do SQL Server pela função gstrCONVERT
' Responsável: Everton Bianchini
'------------------------------------------------------------------------------------------
' Data: 05/05/2003
' Alteração: - Substituição do comando nativo de concatenação de strings (+) do SQL Server
'            pela variável strCONCAT.
' Responsável: Everton Bianchini
'------------------------------------------------------------------------------------------
' Data: 05/05/2003
' Alteração: - Retirado comando "AS" utilizado para dar apelidos às tabelas, o qual não é
'            permitido pelo Oracle, da cláusula FROM.
' Responsável: Everton Bianchini
'******************************************************************************************
    
Dim strSql As String
    
    strSql = ""
'    strSql = strSql & " SELECT EC.PKId, CONVERT(VARCHAR, AEC.intCodigo) + ' - ' + AEC.strDescricao AS strAtividadePrincipal, "
    strSql = strSql & " SELECT EC.PKId, " & gstrCONVERT(CDT_VARCHAR, "AEC.intCodigo") & strCONCAT & " ' - ' " & strCONCAT & " AEC.strDescricao AS strAtividadePrincipal, "
    strSql = strSql & " PP.dblTotalPago, PP.dtmDataPagamento "
    strSql = strSql & " FROM "
'    strSql = strSql & gstrEconomico & " AS EC, "
    strSql = strSql & gstrEconomico & " EC, "
'    strSql = strSql & gstrContribuinte & " AS CO, "
    strSql = strSql & gstrContribuinte & " CO, "
'    strSql = strSql & gstrPagamentoParcela & " AS PP, "
    strSql = strSql & gstrPagamentoParcela & " PP, "
'    strSql = strSql & gstrAtividadeDaEmpresa & " AS AE, "
    strSql = strSql & gstrAtividadeDaEmpresa & " AE, "
'    strSql = strSql & gstrAtividadeEC & " AS AEC "
    strSql = strSql & gstrAtividadeEC & " AEC "
    strSql = strSql & " WHERE "
    strSql = strSql & " AE.intEconomico = EC.PKId AND AE.blnPrincipal = 1 "
    strSql = strSql & " AND AEC.PKId = AE.intAtividade "
    strSql = strSql & " AND CO.PKId = EC.intContribuinte "
    strSql = strSql & " AND CO.PKId = PP.intContribuinte "
    strSql = strSql & " AND EC.PKId BETWEEN " & dbc_AtividadeInicial.BoundText & " AND " & dbc_AtividadeFinal.BoundText
    strSql = strSql & " AND PP.dtmDataPagamento BETWEEN " & gstrConvDtParaSql(txt_dtmInicial.Text) & " AND " & gstrConvDtParaSql(txt_dtmFinal.Text)

strQueryRelatorio = strSql
End Function

Private Sub dbc_AtividadeInicial_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbc_AtividadeInicial
End Sub

Private Sub dbc_AtividadeFinal_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbc_AtividadeFinal
End Sub

Private Sub txt_dtmInicial_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "D", txt_dtmInicial
End Sub

Private Sub txt_dtmInicial_GotFocus()
    MarcaCampo txt_dtmInicial
End Sub

Private Sub txt_dtmInicial_LostFocus()
    txt_dtmInicial.Text = gstrDataFormatada(txt_dtmInicial.Text)
End Sub

Private Sub txt_dtmFinal_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "D", txt_dtmFinal
End Sub

Private Sub txt_dtmFinal_GotFocus()
    MarcaCampo txt_dtmFinal
End Sub

Private Sub txt_dtmFinal_LostFocus()
    txt_dtmFinal.Text = gstrDataFormatada(txt_dtmFinal.Text)
End Sub

