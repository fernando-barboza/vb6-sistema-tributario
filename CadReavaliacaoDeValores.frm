VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmCadReavaliacaoDeValores 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reavaliação de Valores"
   ClientHeight    =   5295
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6285
   HelpContextID   =   13
   Icon            =   "CadReavaliacaoDeValores.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5295
   ScaleWidth      =   6285
   Begin TabDlg.SSTab tab_3DPasta 
      Height          =   5145
      Left            =   120
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   60
      Width           =   6075
      _ExtentX        =   10716
      _ExtentY        =   9075
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Reavaliação de Valores"
      TabPicture(0)   =   "CadReavaliacaoDeValores.frx":1042
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lbl_Valor"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lbl_Final"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lbl_Porc"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lbl_Inicial"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lbl_CodigoDaUtilizacao"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txt_CodInicial"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txt_Valor"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txt_CodFinal"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "fra_ValoresReajustados"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Fra_TipoAjuste"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "chk_Todos"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "dbcUtilizacao"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).ControlCount=   12
      Begin MSDataListLib.DataCombo dbcUtilizacao 
         Height          =   315
         Left            =   1170
         TabIndex        =   16
         Top             =   450
         Width           =   4785
         _ExtentX        =   8440
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
      Begin VB.CheckBox chk_Todos 
         Caption         =   "Todos"
         Height          =   315
         Left            =   5190
         TabIndex        =   4
         Top             =   1500
         Width           =   825
      End
      Begin VB.Frame Fra_TipoAjuste 
         Height          =   615
         Left            =   1170
         TabIndex        =   11
         Top             =   780
         Width           =   2835
         Begin VB.OptionButton opt_Porcentagem 
            Caption         =   "Decréssimo"
            Height          =   195
            Index           =   1
            Left            =   1500
            TabIndex        =   12
            Top             =   270
            Width           =   1155
         End
         Begin VB.OptionButton opt_Porcentagem 
            Caption         =   "Acréssimo"
            Height          =   195
            Index           =   0
            Left            =   210
            TabIndex        =   0
            Top             =   270
            Value           =   -1  'True
            Width           =   1095
         End
      End
      Begin VB.Frame fra_ValoresReajustados 
         Caption         =   "Valores"
         Height          =   3105
         Left            =   90
         TabIndex        =   10
         Top             =   1920
         Width           =   5835
         Begin MSComctlLib.ListView lvw_Lista 
            Height          =   2775
            Left            =   120
            TabIndex        =   5
            Top             =   240
            Width           =   5595
            _ExtentX        =   9869
            _ExtentY        =   4895
            View            =   3
            LabelEdit       =   1
            Sorted          =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   1
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Descrição"
               Object.Width           =   52917
            EndProperty
         End
      End
      Begin VB.TextBox txt_CodFinal 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3600
         MaxLength       =   15
         TabIndex        =   3
         Top             =   1500
         Width           =   1335
      End
      Begin VB.TextBox txt_Valor 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4620
         MaxLength       =   15
         MultiLine       =   -1  'True
         TabIndex        =   1
         Top             =   1080
         Width           =   1125
      End
      Begin VB.TextBox txt_CodInicial 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1170
         MaxLength       =   15
         TabIndex        =   2
         Top             =   1500
         Width           =   1335
      End
      Begin VB.Label lbl_CodigoDaUtilizacao 
         AutoSize        =   -1  'True
         Caption         =   "Utilização"
         Height          =   195
         Left            =   390
         TabIndex        =   15
         Top             =   570
         Width           =   690
      End
      Begin VB.Label lbl_Inicial 
         AutoSize        =   -1  'True
         Caption         =   "Código Inicial "
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   1590
         Width           =   990
      End
      Begin VB.Label lbl_Porc 
         AutoSize        =   -1  'True
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5805
         TabIndex        =   13
         Top             =   1170
         Width           =   150
      End
      Begin VB.Label lbl_Final 
         AutoSize        =   -1  'True
         Caption         =   "Código Final"
         Height          =   195
         Left            =   2595
         TabIndex        =   9
         Top             =   1590
         Width           =   870
      End
      Begin VB.Label lbl_Valor 
         AutoSize        =   -1  'True
         Caption         =   "Índice"
         Height          =   195
         Left            =   4080
         TabIndex        =   8
         Top             =   1170
         Width           =   435
      End
   End
   Begin MSComctlLib.ListView lvw_ValorTemporario 
      Height          =   1725
      Left            =   2400
      TabIndex        =   7
      Top             =   2100
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   3043
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Codigo"
         Object.Width           =   52917
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Valor"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmCadReavaliacaoDeValores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mblnAlterando               As Boolean
Dim mobjAux                     As Object
Dim objList                     As Object
Dim objListTemp                 As Object
Dim Cancela                     As Boolean

Private Sub dbcUtilizacao_Click(Area As Integer)
    DropDownDataCombo dbcUtilizacao, Me, Area
    If Area = 2 And dbcUtilizacao.MatchedWithList = True Then
        LeDaTabelaParaObj gstrTabelaDeValor, lvw_Lista, strQueryValor
        fra_ValoresReajustados.Caption = "Valores"
        HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrCalcularReajuste
        HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrSalvar
    End If
End Sub

Private Function strQueryValor()
    'Traz todos os valores da Utilização
    Dim strSql As String
    strSql = ""
    strSql = strSql & "Select PKId, PKId intCodigo, dblValor "
    strSql = strSql & "From " & gstrTabelaDeValor & " "
    strSql = strSql & "Where intCodigoDaUtilizacao = " & dbcUtilizacao.BoundText
    strQueryValor = strSql
End Function

Private Function strQueryValorEscolido() As String
    'Traz todos os valores do código escolhido
    Dim strSql   As String
    Dim Inicial  As Double
    Dim Final    As Double
    
    Inicial = 0
    Final = 0
    Cancela = False
    If dbcUtilizacao.BoundText <= 0 Then
        ExibeMensagem "Selecione uma Utilização."
        dbcUtilizacao.SetFocus
        Cancela = True
        Screen.MousePointer = 0
        Exit Function
    End If
    If txt_Valor = "" Then
        ExibeMensagem "O Índice tem que ser preenchido."
        txt_Valor.SetFocus
        Cancela = True
        Screen.MousePointer = 0
        Exit Function
    End If
    If chk_Todos.Value = 0 Then
        If txt_CodInicial = "" Then
            ExibeMensagem "O código Inicial tem que ser preenchido."
            txt_CodInicial.SetFocus
            Cancela = True
            Screen.MousePointer = 0
            Exit Function
        ElseIf txt_CodFinal = "" Then
            ExibeMensagem "O código Final tem que ser preenchido."
            txt_CodFinal.SetFocus
            Cancela = True
            Screen.MousePointer = 0
            Exit Function
        End If
    End If
    If Val(txt_CodInicial) > Val(txt_CodFinal) Then
        ExibeMensagem "O código Inicial tem que ser menor que o Final."
        txt_CodInicial.SetFocus
        Cancela = True
        Screen.MousePointer = 0
        Exit Function
    End If
    
    Inicial = Val(txt_CodInicial)
    Final = Val(txt_CodFinal)
    
    strSql = ""
    strSql = strSql & "Select PKId, PKId intCodigo, dblValor "
    strSql = strSql & "From " & gstrTabelaDeValor & " "
    strSql = strSql & "Where intCodigoDaUtilizacao = " & dbcUtilizacao.BoundText
    If chk_Todos.Value = 1 Then
        strQueryValorEscolido = strSql
        Cancela = False
        Screen.MousePointer = 0
        Exit Function
    Else
        strSql = strSql & " AND PKId BETWEEN " & Inicial & " AND " & Final
    End If
    strQueryValorEscolido = strSql
    Cancela = False
End Function

Private Sub chk_Todos_Click()
    If chk_Todos.Value = 1 Then
        txt_CodFinal.Enabled = False
        txt_CodInicial.Enabled = False
        txt_CodFinal.BackColor = &H80000004
        txt_CodInicial.BackColor = &H80000004
        Else
        txt_CodFinal.Enabled = True
        txt_CodInicial.Enabled = True
        txt_CodFinal.BackColor = &H80000005
        txt_CodInicial.BackColor = &H80000005
    End If
End Sub

Private Sub dbcUtilizacao_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbcUtilizacao, Me, , KeyCode, Shift
End Sub

Private Sub Form_Activate()
    gintCodSeguranca = 634
    VirificaGradeListView Me
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
 On Error Resume Next
  If KeyCode = vbKeyF1 Then
     Call_HtmlHelp Me.HelpContextID
  End If
End Sub

Private Sub Form_Load()
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrSalvar, gstrDeletar
    HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrCalcularReajuste
    
    MontaColumnHeaders
    LeDaTabelaParaObj gstrUtilizacaoDaTabelaDeValor, dbcUtilizacao
    
    fra_ValoresReajustados.Caption = "Valores"
    lvw_ValorTemporario.Visible = False
End Sub

Sub MontaColumnHeaders()
    With lvw_Lista
        .ColumnHeaders.Clear
        .ColumnHeaders.Add 1, , "Código", 1600
        .ColumnHeaders.Add 2, , "Valor", 3600
    End With
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrCalcularReajuste
End Sub

Private Sub lvw_Lista_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    OrdenaColunaClicada lvw_Lista, ColumnHeader
End Sub

Private Sub LimpaObjetos()
    dbcUtilizacao.BoundText = ""
    txt_Valor = ""
    txt_CodInicial = ""
    txt_CodFinal = ""
    chk_Todos.Value = 0
    opt_Porcentagem(0).Value = True
    fra_ValoresReajustados.Caption = "Valores"
    lvw_Lista.ListItems.Clear
    lvw_ValorTemporario.ListItems.Clear
    dbcUtilizacao.SetFocus
    HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrCalcularReajuste
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrSalvar
End Sub

Private Function blnGravaValoresReajustados() As Boolean
    Dim strSql As String
    Dim i      As Integer
    
    On Error GoTo err_Reajuste
    
    Set gobjBanco = New clsBanco
    gobjBanco.ExecutaBeginTrans
    
    For i = 1 To lvw_Lista.ListItems.Count
        strSql = ""
        strSql = strSql & " UPDATE " & gstrTabelaDeValor & " SET dblValor = " & gstrConvVrParaSql(lvw_Lista.ListItems(i).SubItems(1))
        strSql = strSql & " WHERE PKId = " & Val(lvw_Lista.ListItems(i).Text)
        Set gobjBanco = New clsBanco
        gobjBanco.Execute strSql
    Next
    
    gobjBanco.ExecutaCommitTrans
    blnGravaValoresReajustados = True
    
Exit Function
err_Reajuste:
    ExibeDetalheErro ""
    blnGravaValoresReajustados = False
    gobjBanco.ExecutaRollbackTrans
End Function

Private Function blnCalcularReajuste() As Boolean
    Screen.MousePointer = 11
    
    LeDaTabelaParaObj gstrTabelaDeValor, lvw_ValorTemporario, strQueryValorEscolido
    
    If Cancela = True Then
        Screen.MousePointer = 0
        Exit Function
    End If
    
    Dim i           As Integer
    Dim Acressimo   As Double
    Dim Decressimo  As Double
    Dim MaisOuMenos As Boolean
    
    If opt_Porcentagem(0).Value = True Then
       Acressimo = CDbl(txt_Valor)
       MaisOuMenos = True
    ElseIf opt_Porcentagem(0).Value = False Then
       Decressimo = "-" & CDbl(txt_Valor)
       MaisOuMenos = False
    End If
    'pega os valores da lvw_ValorTemporario, reajusta e ...
    lvw_Lista.ListItems.Clear
    For i = 1 To lvw_ValorTemporario.ListItems.Count
    '...enche a lvw_Lista com os novos valores
        If MaisOuMenos = True Then
            Set objList = lvw_Lista.ListItems.Add(, , lvw_ValorTemporario.ListItems(i).Text)
            Set objListTemp = lvw_ValorTemporario.ListItems(i)
            objList.SubItems(1) = objListTemp.SubItems(1) * (1 + (Acressimo / 100))
        Else
            Set objList = lvw_Lista.ListItems.Add(, , (lvw_ValorTemporario.ListItems(i).Text))
            Set objListTemp = lvw_ValorTemporario.ListItems(i)
            objList.SubItems(1) = objListTemp.SubItems(1) * (1 + (Decressimo / 100))
        End If
    Next
    
    blnCalcularReajuste = True
    fra_ValoresReajustados.Caption = "Valores Reajustados"
    Screen.MousePointer = 0
    lvw_ValorTemporario.Visible = False
End Function

Private Sub txt_Valor_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "V", txt_Valor
End Sub

Private Sub txt_Valor_LostFocus()
    txt_Valor = gvntConvVrDoSql(txt_Valor)
End Sub

Private Sub txt_Valor_GotFocus()
    MarcaCampo txt_Valor
End Sub

Private Sub txt_CodFinal_GotFocus()
    MarcaCampo txt_CodFinal
End Sub

Private Sub txt_CodInicial_GotFocus()
    MarcaCampo txt_CodInicial
End Sub

Private Sub txt_CodFinal_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txt_CodFinal
End Sub

Private Sub txt_CodInicial_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txt_CodInicial
End Sub

Public Sub MantemForm(ByVal strModoOperacao As String)
    lvw_ValorTemporario.Visible = False
    Select Case UCase(strModoOperacao)
        Case gstrCalcularReajuste
            If blnCalcularReajuste = True Then
                HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrSalvar
            End If
            
        Case gstrNovo
            LimpaObjetos
            
        Case gstrSalvar
            If MsgBox("Confirma a gravação dos novos valores?", vbYesNo) = vbYes Then
               If blnGravaValoresReajustados = True Then
                  LimpaObjetos
               End If
            End If
            
        Case gstrFechar
            Unload Me
            
    End Select
End Sub

