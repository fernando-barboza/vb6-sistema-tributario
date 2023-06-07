VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmRelatorioDeOperacaoDoUsuario 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relatório de Operações do Usuário"
   ClientHeight    =   3810
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6345
   Icon            =   "RelatorioDeOperacaoDoUsuario.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3810
   ScaleWidth      =   6345
   Begin TabDlg.SSTab SSTab1 
      Height          =   3525
      Left            =   150
      TabIndex        =   6
      Top             =   150
      Width           =   6075
      _ExtentX        =   10716
      _ExtentY        =   6218
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Operações do Usuário"
      TabPicture(0)   =   "RelatorioDeOperacaoDoUsuario.frx":1042
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin VB.Frame Frame1 
         Height          =   2895
         Left            =   150
         TabIndex        =   7
         Top             =   450
         Width           =   5775
         Begin VB.Frame Frame2 
            Caption         =   "Data da operação"
            Height          =   885
            Left            =   900
            TabIndex        =   10
            Top             =   1830
            Width           =   4725
            Begin VB.TextBox txt_Final 
               Height          =   285
               Left            =   3510
               MaxLength       =   11
               TabIndex        =   5
               Top             =   360
               Width           =   975
            End
            Begin VB.TextBox txt_Inicial 
               Height          =   285
               Left            =   750
               MaxLength       =   11
               TabIndex        =   4
               Top             =   360
               Width           =   975
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               Caption         =   "Final"
               Height          =   195
               Left            =   3060
               TabIndex        =   12
               Top             =   450
               Width           =   330
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Inicial"
               Height          =   195
               Left            =   210
               TabIndex        =   11
               Top             =   450
               Width           =   405
            End
         End
         Begin VB.CheckBox chk_Usuario 
            Caption         =   "Selecionar todos os usuários"
            Height          =   225
            Left            =   900
            TabIndex        =   1
            Top             =   690
            Width           =   2475
         End
         Begin VB.CheckBox chk_Funcao 
            Caption         =   "Selecionar todas as funções"
            Height          =   225
            Left            =   900
            TabIndex        =   3
            Top             =   1530
            Width           =   2385
         End
         Begin MSDataListLib.DataCombo dbcintUsuario 
            Height          =   315
            Left            =   900
            TabIndex        =   0
            Top             =   240
            Width           =   4725
            _ExtentX        =   8334
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dbcintFuncao 
            Height          =   315
            Left            =   900
            TabIndex        =   2
            Top             =   1110
            Width           =   4725
            _ExtentX        =   8334
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Função"
            Height          =   195
            Left            =   240
            TabIndex        =   9
            Top             =   1200
            Width           =   540
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Usuário"
            Height          =   195
            Left            =   240
            TabIndex        =   8
            Top             =   360
            Width           =   540
         End
      End
   End
End
Attribute VB_Name = "frmRelatorioDeOperacaoDoUsuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'>>> Valores para o campo tblHistoricoOperacao.bytModulo
'''1
''' COMPRAS
'''
'''2
''' CONCURSO PÚBLICO
'''
'''3
''' CRIANÇA E ADOLESCENTE
'''
'''4
''' ESCOLAR
'''
'''5
''' FROTA
'''
'''6
''' LEGISLAÇÃO
'''
'''7
''' MATERIAIS
'''
'''8
''' ORÇAMENTÁRIO
'''
'''9
''' OUVIDORIA
'''
'''10
''' PATRIMÔNIO
'''
'''11
''' PROTOCOLO
'''
'''12
''' RECURSOS HUMANOS
'''
'''13
''' SEGURANÇA
'''
'''14
''' TRIBUTÁRIO
'>>> Valores para o campo tblHistoricoOperacao.bytModulo

Option Explicit
    Dim mblnAlterando    As Boolean
    Dim mobjAux          As Object
    Dim mblnSelecionou   As Boolean
    Dim mblnPrimeiraVez  As Boolean
    
Private Sub chk_Funcao_Click()
    If chk_Funcao.Value = 1 Then
        dbcintFuncao.BoundText = ""
        dbcintFuncao.Enabled = False
        TrocaCorObjeto dbcintFuncao, True
    Else
        dbcintFuncao.Enabled = True
        TrocaCorObjeto dbcintFuncao, False
    End If
End Sub

Private Sub chk_Usuario_Click()
    If chk_Usuario.Value = 1 Then
        dbcintUsuario.BoundText = ""
        dbcintUsuario.Enabled = False
        TrocaCorObjeto dbcintUsuario, True
    Else
        dbcintUsuario.Enabled = True
        TrocaCorObjeto dbcintUsuario, False
    End If
End Sub

Private Sub chk_Funcao_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", chk_Funcao
End Sub

Private Sub chk_Usuario_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", chk_Usuario
End Sub

Private Sub dbcintFuncao_Click(Area As Integer)
    DropDownDataCombo dbcintFuncao, Me, Area
End Sub

Private Sub dbcintFuncao_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbcintFuncao, Me, , KeyCode, Shift
End Sub

Private Sub dbcintUsuario_Click(Area As Integer)
    DropDownDataCombo dbcintUsuario, Me, Area
End Sub

Private Sub dbcintUsuario_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbcintUsuario, Me, , KeyCode, Shift
End Sub

Private Sub Form_Activate()
    gintCodSeguranca = 864
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
    LeDaTabelaParaObj gstrUsuarios, dbcintUsuario, strQuerryUsuario
    LeDaTabelaParaObj gstrCatalogoTabela, dbcintFuncao, strQuerryFuncao
    txt_Final.Text = gstrDataDoSistema
    HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrSalvar, gstrDeletar
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar
    mblnSelecionou = False
    mblnPrimeiraVez = False
End Sub

Public Sub MantemForm(ByVal strModoOperacao As String)
On Error Resume Next
Dim adoResultado   As ADODB.Recordset

    If UCase(strModoOperacao) = UCase(gstrImprimir) Then
        If blnDadosOk Then
            ImprimeRelatorio rptRelatorioDeOperacaoDoUsuario, strQuerryRelatorio
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
    dbcintUsuario.BoundText = ""
    dbcintFuncao.BoundText = ""
    txt_Inicial.Text = ""
    txt_Final.Text = ""
    If dbcintUsuario.Enabled = True Then
        dbcintUsuario.SetFocus
    End If
End Sub

Private Function blnDadosOk() As Boolean
blnDadosOk = False
On Error GoTo err_blnDadosOK
    If chk_Usuario.Value = 0 Then
        If dbcintUsuario.MatchedWithList = False Then
           ExibeMensagem "O usuário tem que ser selecionado."
           dbcintUsuario.SetFocus
           Exit Function
        End If
    End If
    If chk_Funcao.Value = 0 Then
        If dbcintFuncao.MatchedWithList = False Then
           ExibeMensagem "A função tem que ser selecionada."
           dbcintFuncao.SetFocus
           Exit Function
        End If
    End If
    If gblnDataValida(txt_Inicial.Text) = False Then
        ExibeMensagem "A data incial não é uma data válida."
        txt_Inicial.SetFocus
        Exit Function
    End If
    If gblnDataValida(txt_Final.Text) = False Then
        ExibeMensagem "A data final não é uma data válida."
        txt_Final.SetFocus
        Exit Function
    End If
    If CVDate(txt_Inicial.Text) > CVDate(txt_Final.Text) Then
        ExibeMensagem "A data inicial tem que ser inferior à data final."
        txt_Inicial.SetFocus
        Exit Function
    End If
    blnDadosOk = True
err_blnDadosOK:
End Function

Private Function strQuerryUsuario() As String
Dim strSql As String
    strSql = ""
    strSql = strSql & " SELECT "
    strSql = strSql & " PKId, strNome "
    strSql = strSql & " FROM "
    strSql = strSql & gstrUsuarios
    strSql = strSql & " ORDER BY strNome "
strQuerryUsuario = strSql
End Function

Private Function strQuerryFuncao() As String
Dim strSql As String
    strSql = ""
    strSql = strSql & " SELECT "
    strSql = strSql & " PKId, strDescricao "
    strSql = strSql & " FROM "
    strSql = strSql & gstrCatalogoTabela
    strSql = strSql & " ORDER BY strDescricao "
strQuerryFuncao = strSql
End Function

Private Function strQuerryRelatorio() As String

'******************************************************************************************
' Data: 23/04/2003
' Alteração: - Substituição da estrutura CASE do SQL Server pela função gstrCASEWHEN.
' Responsável: Everton Bianchini
'******************************************************************************************
    
Dim strSql As String
    strSql = ""
    strSql = strSql & " SELECT "
    strSql = strSql & " C.PKId AS PKIdUsuario, C.strNome, "
    strSql = strSql & " B.PKId AS PKIdTabela, B.strDescricao AS Funcao, A.bytModulo, "
    strSql = strSql & " A.dtmData, A.strValor, A.strOperacao, "
    
'    strSql = strSql & " CASE A.bytModulo "
'    strSql = strSql & " WHEN 1 THEN 'Compras' "
'    strSql = strSql & " WHEN 2 THEN 'Concurso Público' "
'    strSql = strSql & " WHEN 3 THEN 'Menor' "
'    strSql = strSql & " WHEN 4 THEN 'Escolar' "
'    strSql = strSql & " WHEN 5 THEN 'Frota' "
'    strSql = strSql & " WHEN 6 THEN 'Legislação' "
'    strSql = strSql & " WHEN 7 THEN 'Material' "
'    strSql = strSql & " WHEN 8 THEN 'Orçamentário' "
'    strSql = strSql & " WHEN 9 THEN 'Ouvidoria' "
'    strSql = strSql & " WHEN 10 THEN 'Patrimônio' "
'    strSql = strSql & " WHEN 11 THEN 'Protocolo' "
'    strSql = strSql & " WHEN 12 THEN 'Recursos Humanos' "
'    strSql = strSql & " WHEN 13 THEN 'Segurança' "
'    strSql = strSql & " WHEN 14 THEN 'Tributário' "

    strSql = strSql & gstrCASEWHEN("A.bytModulo", _
        "1, 'Compras', " & _
        "2, 'Concurso Público', " & _
        "3, 'Menor', " & _
        "4, 'Escolar', " & _
        "5, 'Frota', " & _
        "6, 'Legislação', " & _
        "7, 'Material', " & _
        "8, 'Orçamentário', " & _
        "9, 'Ouvidoria', " & _
        "10, 'Patrimônio', " & _
        "11, 'Protocolo', " & _
        "12, 'Recursos Humanos', " & _
        "13, 'Segurança', " & _
        "14, 'Tributário'")

'    strSql = strSql & " END AS Modulo, "
    strSql = strSql & " AS Modulo, "
 
'    strSql = strSql & "Case A.strOperacao "
'    strSql = strSql & "WHEN 'A' THEN 'Alteração' "
'    strSql = strSql & "WHEN 'I' THEN 'Inclusão' "
'    strSql = strSql & "WHEN 'E' THEN 'Exclusão' "
    strSql = strSql & gstrCASEWHEN("A.strOperacao", _
        "'A', 'Alteração', " & _
        "'I', 'Inclusão', " & _
        "'E', 'Exclusão'")
    
'    strSql = strSql & "END AS Operacao "
    strSql = strSql & " AS Operacao "

    strSql = strSql & " FROM "
    strSql = strSql & gstrHistoricoOperacao & " A, "
    strSql = strSql & gstrCatalogoTabela & " B, "
    strSql = strSql & gstrUsuarios & " C "
    
    strSql = strSql & " WHERE "
    strSql = strSql & " B.PKId = A.intCatalogoTabela "
    strSql = strSql & " AND C.PKId = A.intUsuario "
    If chk_Usuario.Value = 0 Then
        strSql = strSql & " AND A.intUsuario = " & Val(dbcintUsuario.BoundText)
    End If
    If chk_Funcao.Value = 0 Then
        strSql = strSql & " AND A.intCatalogoTabela = " & Val(dbcintFuncao.BoundText)
    End If
    strSql = strSql & " AND A.dtmData BETWEEN " & gstrConvDtParaSql(txt_Inicial.Text) & " AND " & gstrConvDtParaSql(txt_Final.Text & " " & Time)
    strSql = strSql & " ORDER BY C.strNome, A.bytModulo, B.strDescricao,  A.strOperacao "
    
strQuerryRelatorio = strSql
End Function

Private Sub dbcintUsuario_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbcintUsuario
End Sub

Private Sub dbcintFuncao_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbcintFuncao
End Sub

Private Sub txt_Inicial_GotFocus()
    MarcaCampo txt_Inicial
End Sub

Private Sub txt_Inicial_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "D", txt_Inicial
End Sub

Private Sub txt_Final_GotFocus()
    MarcaCampo txt_Final
End Sub

Private Sub txt_Final_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "D", txt_Final
End Sub
