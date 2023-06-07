VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmRelatorioDeContadoresArrecadacaoPeriodo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Contadores e Arrecadação no Período"
   ClientHeight    =   3000
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7500
   Icon            =   "RelatorioDeContadoresArrecadacaoPeriodo.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3000
   ScaleWidth      =   7500
   Begin TabDlg.SSTab tab_3dPasta 
      Height          =   2715
      Left            =   180
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   150
      Width           =   7155
      _ExtentX        =   12621
      _ExtentY        =   4789
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Contadores e Arrecadação no Período"
      TabPicture(0)   =   "RelatorioDeContadoresArrecadacaoPeriodo.frx":1042
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin VB.Frame Frame1 
         Height          =   2175
         Left            =   150
         TabIndex        =   6
         Top             =   360
         Width           =   6825
         Begin VB.Frame Frame2 
            Caption         =   "Data de Pagamento"
            Height          =   705
            Left            =   1590
            TabIndex        =   9
            Top             =   1320
            Width           =   5055
            Begin VB.TextBox txt_DataFinal 
               Height          =   285
               Left            =   3780
               MaxLength       =   11
               TabIndex        =   4
               Top             =   300
               Width           =   1065
            End
            Begin VB.TextBox txt_DataInicial 
               Height          =   285
               Left            =   900
               MaxLength       =   11
               TabIndex        =   3
               Top             =   300
               Width           =   1065
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "Final"
               Height          =   195
               Left            =   3210
               TabIndex        =   11
               Top             =   390
               Width           =   330
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Inicial"
               Height          =   195
               Left            =   270
               TabIndex        =   10
               Top             =   390
               Width           =   405
            End
         End
         Begin VB.CheckBox chk_Selecionar 
            Caption         =   "Selecionar todos os Contadores"
            Height          =   255
            Left            =   1590
            TabIndex        =   2
            Top             =   1020
            Width           =   2835
         End
         Begin MSDataListLib.DataCombo dbcintContribuinteInicial 
            Height          =   315
            Left            =   1590
            TabIndex        =   0
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
            TabIndex        =   1
            Top             =   660
            Width           =   5055
            _ExtentX        =   8916
            _ExtentY        =   556
            _Version        =   393216
            IntegralHeight  =   0   'False
            Text            =   ""
         End
         Begin VB.Label lbl_label2 
            AutoSize        =   -1  'True
            Caption         =   "Contador Final"
            Height          =   195
            Left            =   405
            TabIndex        =   8
            Top             =   765
            Width           =   1020
         End
         Begin VB.Label lbl_Label1 
            AutoSize        =   -1  'True
            Caption         =   "Contador Inicial"
            Height          =   195
            Left            =   330
            TabIndex        =   7
            Top             =   345
            Width           =   1095
         End
      End
   End
End
Attribute VB_Name = "frmRelatorioDeContadoresArrecadacaoPeriodo"
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

Private Sub chk_Selecionar_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", chk_Selecionar
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
    
    dbcintContribuinteInicial.Tag = strQuerryContador & ";STRNOME"
    dbcintContribuinteFinal.Tag = dbcintContribuinteInicial.Tag
    
    txt_DataFinal.Text = gstrDataDoSistema
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
            If Val(dbcintContribuinteInicial.BoundText) < Val(dbcintContribuinteFinal.BoundText) Then
                CCInicial = Val(dbcintContribuinteInicial.BoundText)
                CCFinal = Val(dbcintContribuinteFinal.BoundText)
            Else
                CCInicial = Val(dbcintContribuinteFinal.BoundText)
                CCFinal = Val(dbcintContribuinteInicial.BoundText)
            End If
        End If
            ImprimeRelatorio rptRelatorioDeContadoresArrecadacaoPeriodo, strQuerryRelatorio
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
    txt_DataInicial.Text = ""
    txt_DataFinal.Text = ""
    dbcintContribuinteInicial.SetFocus
End Sub

Private Function blnDadosOk() As Boolean
blnDadosOk = False
On Error GoTo err_blnDadosOK
    If chk_Selecionar.Value = 0 Then
        If dbcintContribuinteInicial.MatchedWithList = False Then
           ExibeMensagem "O Contador Incial tem que ser selecionado."
           dbcintContribuinteInicial.SetFocus
           Exit Function
        End If
        If dbcintContribuinteFinal.MatchedWithList = False Then
           ExibeMensagem "O Contador Final tem que ser selecionado."
           dbcintContribuinteFinal.SetFocus
           Exit Function
        End If
    End If
    If gblnDataValida(txt_DataInicial.Text) = False Then
        ExibeMensagem "A data inicial não é uma data válida."
        txt_DataInicial.SetFocus
        Exit Function
    End If
    If gblnDataValida(txt_DataFinal.Text) = False Then
        ExibeMensagem "A data final não é uma data válida."
        txt_DataFinal.SetFocus
        Exit Function
    End If
    If CVDate(txt_DataFinal.Text) < CVDate(txt_DataInicial.Text) Then
        ExibeMensagem "A data inicial tem que ser anterior à data final."
        txt_DataInicial.SetFocus
        Exit Function
    End If
blnDadosOk = True
err_blnDadosOK:
End Function

Private Function strQuerryContador() As String
Dim strSql As String
    strSql = ""
    strSql = strSql & " SELECT "
    strSql = strSql & " A.intContribuinte, B.strNome "
    strSql = strSql & " FROM "
    strSql = strSql & gstrContador & " A, "
    strSql = strSql & gstrContribuinte & " B "
    strSql = strSql & " WHERE "
    strSql = strSql & " A.intContribuinte = B.PKId "
    strSql = strSql & " ORDER BY B.strNome "
strQuerryContador = strSql
End Function

Private Function strQuerryRelatorio() As String
Dim strSql As String
    
    strSql = ""
    strSql = strSql & " SELECT "
    strSql = strSql & " B.strNome, A.strCrc, SUM(D.dblTotalPago) AS TotalArrecadado "
    
    strSql = strSql & " FROM "
    strSql = strSql & gstrContador & " A, "
    strSql = strSql & gstrContribuinte & " B, "
    strSql = strSql & gstrEconomico & " C, "
    strSql = strSql & gstrPagamentoParcela & " D "

    strSql = strSql & " WHERE "
    strSql = strSql & " A.intContribuinte = B.PKId "
    strSql = strSql & " AND C.intContador = A.PKId "
    strSql = strSql & " AND D.intContribuinte = C.intContribuinte "
    strSql = strSql & " AND D.dtmDataPagamento BETWEEN " & gstrConvDtParaSql(txt_DataInicial.Text) & " AND " & gstrConvDtParaSql(txt_DataFinal.Text)
    If chk_Selecionar.Value = 0 Then
        strSql = strSql & " AND A.intContribuinte BETWEEN " & CCInicial & " AND " & CCFinal
    End If
    strSql = strSql & " GROUP BY C.intContador, B.strNome, A.strCRC "
    strSql = strSql & " ORDER BY B.strNome "
    
strQuerryRelatorio = strSql
End Function

Private Sub dbcintContribuinteFinal_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbcintContribuinteFinal
End Sub

Private Sub dbcintContribuinteInicial_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbcintContribuinteInicial
End Sub

Private Sub txt_DataFinal_GotFocus()
    MarcaCampo txt_DataFinal
End Sub

Private Sub txt_DataInicial_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "D", txt_DataInicial
End Sub

Private Sub txt_DataInicial_GotFocus()
    MarcaCampo txt_DataInicial
End Sub

Private Sub txt_DataFinal_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "D", txt_DataFinal
End Sub
