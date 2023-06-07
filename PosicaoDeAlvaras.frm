VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmPosicaoDeAlvaras 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Posição de Alvarás"
   ClientHeight    =   2670
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8400
   Icon            =   "PosicaoDeAlvaras.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2670
   ScaleWidth      =   8400
   Begin TabDlg.SSTab tab_3dPasta 
      Height          =   2385
      Left            =   150
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   150
      Width           =   8115
      _ExtentX        =   14314
      _ExtentY        =   4207
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Posição de Alvarás"
      TabPicture(0)   =   "PosicaoDeAlvaras.frx":1042
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin VB.Frame Frame1 
         Height          =   1785
         Left            =   180
         TabIndex        =   5
         Top             =   390
         Width           =   7755
         Begin VB.TextBox txt_Data 
            Height          =   285
            Left            =   2550
            MaxLength       =   11
            TabIndex        =   3
            Top             =   1350
            Width           =   1095
         End
         Begin VB.CheckBox chk_Selecionar 
            Caption         =   "Selecionar todas as Atividades"
            Height          =   255
            Left            =   2550
            TabIndex        =   2
            Top             =   1020
            Width           =   2835
         End
         Begin MSDataListLib.DataCombo dbcintAtividadeInicial 
            Height          =   315
            Left            =   2550
            TabIndex        =   0
            Top             =   240
            Width           =   5055
            _ExtentX        =   8916
            _ExtentY        =   556
            _Version        =   393216
            IntegralHeight  =   0   'False
            Style           =   2
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dbcintAtividadeFinal 
            Height          =   315
            Left            =   2550
            TabIndex        =   1
            Top             =   660
            Width           =   5055
            _ExtentX        =   8916
            _ExtentY        =   556
            _Version        =   393216
            IntegralHeight  =   0   'False
            Style           =   2
            Text            =   ""
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Com data de emissão a partir de"
            Height          =   195
            Left            =   165
            TabIndex        =   8
            Top             =   1440
            Width           =   2265
         End
         Begin VB.Label lbl_label2 
            AutoSize        =   -1  'True
            Caption         =   "Atividade Final"
            Height          =   195
            Left            =   1395
            TabIndex        =   7
            Top             =   780
            Width           =   1035
         End
         Begin VB.Label lbl_Label1 
            AutoSize        =   -1  'True
            Caption         =   "Atividade Inicial"
            Height          =   195
            Left            =   1320
            TabIndex        =   6
            Top             =   360
            Width           =   1110
         End
      End
   End
End
Attribute VB_Name = "frmPosicaoDeAlvaras"
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
        dbcintAtividadeInicial.BoundText = ""
        dbcintAtividadeFinal.BoundText = ""
        dbcintAtividadeInicial.Enabled = False
        TrocaCorObjeto dbcintAtividadeInicial, True
        dbcintAtividadeFinal.Enabled = False
        TrocaCorObjeto dbcintAtividadeFinal, True
    Else
        dbcintAtividadeInicial.Enabled = True
        TrocaCorObjeto dbcintAtividadeInicial, False
        dbcintAtividadeFinal.Enabled = True
        TrocaCorObjeto dbcintAtividadeFinal, False
    End If
End Sub

Private Sub dbcintAtividadeFinal_Click(Area As Integer)
    DropDownDataCombo dbcintAtividadeFinal, Me, Area
End Sub

Private Sub dbcintAtividadeFinal_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbcintAtividadeFinal, Me, , KeyCode, Shift
End Sub

Private Sub dbcintAtividadeInicial_Click(Area As Integer)
    DropDownDataCombo dbcintAtividadeInicial, Me, Area
End Sub

Private Sub dbcintAtividadeInicial_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbcintAtividadeInicial, Me, , KeyCode, Shift
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
    LeDaTabelaParaObj gstrAtividadeBasica, dbcintAtividadeInicial, strQuerryAtividade
    LeDaTabelaParaObj gstrAtividadeBasica, dbcintAtividadeFinal, strQuerryAtividade
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
Dim adoResultado   As adodb.Recordset

    If UCase(strModoOperacao) = UCase(gstrImprimir) Then
        If blnDadosOk Then
        If chk_Selecionar.Value = 0 Then
            If Val(dbcintAtividadeInicial.BoundText) < Val(dbcintAtividadeFinal.BoundText) Then
                CCInicial = Val(dbcintAtividadeInicial.BoundText)
                CCFinal = Val(dbcintAtividadeFinal.BoundText)
            Else
                CCInicial = Val(dbcintAtividadeFinal.BoundText)
                CCFinal = Val(dbcintAtividadeInicial.BoundText)
            End If
        End If
            ImprimeRelatorio rptPosicaoDeAlvaras, strQuerryRelatorio
        End If
    End If
    
    If UCase(strModoOperacao) = UCase(gstrNovo) Then
        LimpaObjetos
    End If
    
    If UCase(strModoOperacao) = UCase(gstrFechar) Then
        Unload Me
    End If
    
End Sub

Private Sub LimpaObjetos()
    dbcintAtividadeInicial.BoundText = ""
    dbcintAtividadeFinal.BoundText = ""
    txt_Data.Text = ""
    dbcintAtividadeInicial.SetFocus
End Sub

Private Function blnDadosOk() As Boolean
blnDadosOk = False
On Error GoTo err_blnDadosOK
    If chk_Selecionar.Value = 0 Then
        If dbcintAtividadeInicial.MatchedWithList = False Then
           ExibeMensagem "A Atividade Incial tem que ser selecionada."
           dbcintAtividadeInicial.SetFocus
           Exit Function
        End If
        If dbcintAtividadeFinal.MatchedWithList = False Then
           ExibeMensagem "A Atividade Final tem que ser selecionada."
           dbcintAtividadeFinal.SetFocus
           Exit Function
        End If
    End If
    If gblnDataValida(txt_Data.Text) = False Then
        ExibeMensagem "Esta não é uma data válida."
        txt_Data.SetFocus
        Exit Function
    End If
    blnDadosOk = True
err_blnDadosOK:
End Function

Private Function strQuerryAtividade() As String
Dim strSql As String
    strSql = ""
    strSql = strSql & " SELECT "
    strSql = strSql & " PKId, strDescricao "
    strSql = strSql & " FROM "
    strSql = strSql & gstrAtividadeBasica
    strSql = strSql & " ORDER BY strDescricao "
strQuerryAtividade = strSql
End Function

Private Function strQuerryRelatorio() As String
Dim strSql As String
    
    strSql = ""
    strSql = strSql & " SELECT "
    strSql = strSql & " A.strNome, "
    strSql = strSql & " PP.strCodigo " & strCONCAT & "'/'" & strCONCAT
    strSql = strSql & " PP.intExercicio " & strCONCAT & " '-' " & strCONCAT
    strSql = strSql & " PP.bitDigito AS strNumeroProcessoAlvara,"
    strSql = strSql & " B.dtmDataAlvara,"
    strSql = strSql & " B.strVigenciaAlvara, C.strDescricao As Atividade "
    strSql = strSql & " FROM "
    strSql = strSql & gstrContribuinte & " A, "
    strSql = strSql & gstrEconomico & " B, "
    strSql = strSql & gstrAtividadeBasica & " C, "
    strSql = strSql & gstrProtocolizacaoProcesso & " PP"
    strSql = strSql & " WHERE "
    strSql = strSql & " B.intContribuinte = A.PKId "
    strSql = strSql & " AND B.intAtividadeBasica = C.PKId "
    strSql = strSql & " AND B.intProcAlvara " & strOUTJOracle & "=" & strOUTJSQLServer & " PP.Pkid"
    
    If chk_Selecionar.Value = 0 Then
        strSql = strSql & " AND B.intAtividadeBasica BETWEEN " & CCInicial & " AND " & CCFinal
    End If
    
    strSql = strSql & " AND B.dtmDataAlvara > " & gstrConvDtParaSql(txt_Data.Text)
    
    strSql = strSql & " ORDER BY C.strDescricao, A.strNome "
    
    
strQuerryRelatorio = strSql
End Function

Private Sub dbcintAtividadeFinal_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbcintAtividadeFinal
End Sub

Private Sub dbcintAtividadeInicial_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbcintAtividadeInicial
End Sub

Private Sub txt_Data_GotFocus()
    MarcaCampo txt_Data
End Sub

Private Sub txt_Data_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "D", txt_Data
End Sub
