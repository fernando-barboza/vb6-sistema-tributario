VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmRelatorioDeInscritosAtivosInativosBaixados 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Inscritos Ativos/Inativos/Baixados"
   ClientHeight    =   3285
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8820
   Icon            =   "RelatorioDeInscritosAtivosInativosBaixados.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3285
   ScaleWidth      =   8820
   Begin TabDlg.SSTab tab_3dPasta 
      Height          =   3015
      Left            =   210
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   120
      Width           =   8475
      _ExtentX        =   14949
      _ExtentY        =   5318
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Inscritos Ativos/Inativos/Baixados"
      TabPicture(0)   =   "RelatorioDeInscritosAtivosInativosBaixados.frx":1042
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin VB.Frame Frame1 
         Height          =   2445
         Left            =   150
         TabIndex        =   7
         Top             =   390
         Width           =   8175
         Begin VB.CheckBox chk_TodasAsOcorrencias 
            Caption         =   "Selecionar todas as Ocorrências"
            Height          =   255
            Left            =   1590
            TabIndex        =   5
            Top             =   2070
            Width           =   2835
         End
         Begin VB.CheckBox chk_Selecionar 
            Caption         =   "Selecionar todos os Contribuintes"
            Height          =   255
            Left            =   1590
            TabIndex        =   2
            Top             =   990
            Width           =   2835
         End
         Begin MSDataListLib.DataCombo dbcintContribuinteInicial 
            Height          =   315
            Left            =   1590
            TabIndex        =   0
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
            TabIndex        =   1
            Top             =   630
            Width           =   6465
            _ExtentX        =   11404
            _ExtentY        =   556
            _Version        =   393216
            IntegralHeight  =   0   'False
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dbcintOcorrenciaInicial 
            Height          =   315
            Left            =   1590
            TabIndex        =   3
            Top             =   1290
            Width           =   6465
            _ExtentX        =   11404
            _ExtentY        =   556
            _Version        =   393216
            IntegralHeight  =   0   'False
            Style           =   2
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dbcintOcorrenciaFinal 
            Height          =   315
            Left            =   1590
            TabIndex        =   4
            Top             =   1710
            Width           =   6465
            _ExtentX        =   11404
            _ExtentY        =   556
            _Version        =   393216
            IntegralHeight  =   0   'False
            Style           =   2
            Text            =   ""
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Ocorrência Final"
            Height          =   195
            Left            =   315
            TabIndex        =   11
            Top             =   1800
            Width           =   1155
         End
         Begin VB.Label lblintOcorrencia 
            AutoSize        =   -1  'True
            Caption         =   "Ocorrência Inicial"
            Height          =   195
            Left            =   240
            TabIndex        =   10
            Top             =   1380
            Width           =   1230
         End
         Begin VB.Label lbl_Label1 
            AutoSize        =   -1  'True
            Caption         =   "Inscrição Inicial"
            Height          =   195
            Left            =   375
            TabIndex        =   9
            Top             =   315
            Width           =   1095
         End
         Begin VB.Label lbl_label2 
            AutoSize        =   -1  'True
            Caption         =   "Inscrição Final"
            Height          =   195
            Left            =   450
            TabIndex        =   8
            Top             =   735
            Width           =   1020
         End
      End
   End
End
Attribute VB_Name = "frmRelatorioDeInscritosAtivosInativosBaixados"
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
Dim OCOInicial                      As Integer
Dim OCOFinal                        As Integer
Dim TipoDeInscricao                 As Integer

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

Private Sub chk_TodasAsOcorrencias_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", chk_TodasAsOcorrencias
End Sub

Private Sub chk_Selecionar_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", chk_Selecionar
End Sub

Private Sub chk_TodasAsOcorrencias_Click()
    If chk_TodasAsOcorrencias.Value = 1 Then
        dbcintOcorrenciaInicial.BoundText = ""
        dbcintOcorrenciaFinal.BoundText = ""
        dbcintOcorrenciaInicial.Enabled = False
        TrocaCorObjeto dbcintOcorrenciaInicial, True
        dbcintOcorrenciaFinal.Enabled = False
        TrocaCorObjeto dbcintOcorrenciaFinal, True
    Else
        dbcintOcorrenciaInicial.Enabled = True
        TrocaCorObjeto dbcintOcorrenciaInicial, False
        dbcintOcorrenciaFinal.Enabled = True
        TrocaCorObjeto dbcintOcorrenciaFinal, False
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

Private Sub dbcintOcorrenciaFinal_Click(Area As Integer)
    DropDownDataCombo dbcintOcorrenciaFinal, Me, Area
End Sub

Private Sub dbcintOcorrenciaFinal_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbcintOcorrenciaFinal, Me, , KeyCode, Shift
End Sub

Private Sub dbcintOcorrenciaFinal_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbcintOcorrenciaFinal
End Sub

Private Sub dbcintOcorrenciaInicial_Click(Area As Integer)
    DropDownDataCombo dbcintOcorrenciaInicial, Me, Area
End Sub

Private Sub dbcintOcorrenciaInicial_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbcintOcorrenciaInicial, Me, , KeyCode, Shift
End Sub

Private Sub dbcintOcorrenciaInicial_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbcintOcorrenciaInicial
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

'******************************************************************************************
' Data: 05/05/2003
' Alteração: - Substituição do comando nativo de concatenação de strings (+) do SQL Server
'            pela variável strCONCAT.
' Responsável: Everton Bianchini
'******************************************************************************************

    CCInicial = 0
    CCFinal = 0
    OCOFinal = 0
    OCOInicial = 0
    
'    dbcintContribuinteInicial.Tag = "SELECT ECO.intContribuinte, (ECO.strInscricaoCadastral + ' = ' + CON.strNome) AS strNome FROM " & gstrEconomico & " ECO, " & gstrContribuinte & " CON WHERE CON.PKId = ECO.intContribuinte ORDER BY ECO.strInscricaoCadastral;strNome"
    dbcintContribuinteInicial.Tag = "SELECT ECO.intContribuinte, (ECO.strInscricaoCadastral " & strCONCAT & " ' = ' " & strCONCAT & " CON.strNome) AS strNome FROM " & gstrEconomico & " ECO, " & gstrContribuinte & " CON WHERE CON.PKId = ECO.intContribuinte ORDER BY ECO.strInscricaoCadastral;strNome"
    dbcintContribuinteFinal.Tag = dbcintContribuinteInicial.Tag

    LeDaTabelaParaObj gstrOcorrencia, dbcintOcorrenciaInicial, strQueryOcorrencia
    LeDaTabelaParaObj gstrOcorrencia, dbcintOcorrenciaFinal, strQueryOcorrencia
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
        If chk_TodasAsOcorrencias.Value = 0 Then
            If Val(dbcintOcorrenciaInicial.BoundText) < Val(dbcintOcorrenciaFinal.BoundText) Then
                OCOInicial = Val(dbcintOcorrenciaInicial.BoundText)
                OCOFinal = Val(dbcintOcorrenciaFinal.BoundText)
            Else
                OCOInicial = Val(dbcintOcorrenciaFinal.BoundText)
                OCOFinal = Val(dbcintOcorrenciaInicial.BoundText)
            End If
        End If
            ImprimeRelatorio rptRelatorioDeInscritosAtivosInativosBaixados, strQuerryRelatorio
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
    dbcintOcorrenciaFinal.BoundText = ""
    dbcintOcorrenciaInicial.BoundText = ""
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
    End If
    If chk_TodasAsOcorrencias.Value = 0 Then
        If dbcintOcorrenciaInicial.MatchedWithList = False Then
           ExibeMensagem "A Ocorrência Incial tem que ser selecionada."
           dbcintOcorrenciaInicial.SetFocus
           Exit Function
        End If
        If dbcintOcorrenciaFinal.MatchedWithList = False Then
           ExibeMensagem "A Ocorrência Final tem que ser selecionada."
           dbcintOcorrenciaFinal.SetFocus
           Exit Function
        End If
    End If
    blnDadosOk = True
err_blnDadosOK:
End Function

Function strQueryOcorrencia() As String
    Dim strSQL As String
    strSQL = ""
    strSQL = strSQL & "SELECT PKId, strDescricao "
    strSQL = strSQL & "FROM " & gstrOcorrencia & " "
    strSQL = strSQL & "WHERE intUtilizacaoDaOcorrencia = 5 "
    strSQL = strSQL & "ORDER BY strDescricao"
    strQueryOcorrencia = strSQL
End Function

Private Function strQuerryRelatorio() As String
Dim strSQL As String
    
    strSQL = ""
    strSQL = strSQL & " SELECT "
    strSQL = strSQL & " B.intOcorrencia, C.strDescricao, B.intContribuinte, A.strNome, A.strCNPJCPF, B.strInscricaoCadastral "
    
    strSQL = strSQL & " FROM "
    strSQL = strSQL & gstrContribuinte & " A, "
    strSQL = strSQL & gstrEconomico & " B, "
    strSQL = strSQL & gstrOcorrencia & " C "
    
    strSQL = strSQL & " WHERE "
    strSQL = strSQL & " B.intContribuinte = A.PKId "
    strSQL = strSQL & " AND B.intOcorrencia = C.PKId "
    
    If chk_TodasAsOcorrencias.Value = 0 Then
        strSQL = strSQL & " AND B.intOcorrencia BETWEEN " & OCOInicial & " AND " & OCOFinal
    End If
    If chk_Selecionar.Value = 0 Then
        strSQL = strSQL & " AND B.intContribuinte BETWEEN " & CCInicial & " AND " & CCFinal
    End If
    strSQL = strSQL & " ORDER BY B.intOcorrencia "
    
strQuerryRelatorio = strSQL
End Function

Private Sub dbcintContribuinteFinal_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbcintContribuinteFinal
End Sub

Private Sub dbcintContribuinteInicial_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbcintContribuinteInicial
End Sub

