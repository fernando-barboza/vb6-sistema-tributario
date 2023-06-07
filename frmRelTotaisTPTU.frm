VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmRelTotaisTPTU 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relatório de Baixas"
   ClientHeight    =   2100
   ClientLeft      =   3075
   ClientTop       =   2655
   ClientWidth     =   6420
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2100
   ScaleWidth      =   6420
   Begin TabDlg.SSTab tab_3dPasta 
      Height          =   1995
      Left            =   60
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   60
      Width           =   6315
      _ExtentX        =   11139
      _ExtentY        =   3519
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Totalização de Lancamento de IPTU"
      TabPicture(0)   =   "frmRelTotaisTPTU.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lbl_Emissao"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lbl_Exercicio"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lbl_Composicao"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "dbc_intComposicao"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txt_intExercicio"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txt_strEmissao"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      Begin VB.TextBox txt_strEmissao 
         Height          =   315
         Left            =   1860
         MaxLength       =   3
         TabIndex        =   3
         Top             =   1425
         Width           =   525
      End
      Begin VB.TextBox txt_intExercicio 
         Height          =   315
         Left            =   1860
         MaxLength       =   4
         TabIndex        =   2
         Top             =   960
         Width           =   525
      End
      Begin MSDataListLib.DataCombo dbc_intComposicao 
         Height          =   315
         Left            =   1860
         TabIndex        =   1
         Top             =   510
         Width           =   4365
         _ExtentX        =   7699
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin VB.Label lbl_Composicao 
         AutoSize        =   -1  'True
         Caption         =   "Composição da Receita"
         Height          =   195
         Left            =   90
         TabIndex        =   6
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label lbl_Exercicio 
         AutoSize        =   -1  'True
         Caption         =   "Exercício"
         Height          =   195
         Left            =   1110
         TabIndex        =   5
         Top             =   1050
         Width           =   675
      End
      Begin VB.Label lbl_Emissao 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Emissão"
         Height          =   195
         Left            =   1200
         TabIndex        =   4
         Top             =   1515
         Width           =   585
      End
   End
End
Attribute VB_Name = "frmRelTotaisTPTU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim mobjAux          As Object
    Dim mblnSelecionou   As Boolean


Private Sub dbc_intComposicao_GotFocus()
    MarcaCampo dbc_intComposicao
End Sub

Private Sub dbc_intComposicao_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbc_intComposicao, Me, , KeyCode, Shift
End Sub

Private Sub dbc_intComposicao_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbc_intComposicao
End Sub

Private Sub Form_Load()
    HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrSalvar, gstrDeletar
    dbc_intComposicao.Tag = strQueryComposicao & ";strDescricao"
End Sub
Private Sub Form_Activate()
'    gintCodSeguranca = 1133
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

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar
    mblnSelecionou = False
End Sub

Public Sub MantemForm(ByVal strModoOperacao As String)
    If UCase(strModoOperacao) = UCase(gstrImprimir) Then
        If blnDadosOk Then
            ImprimeRelatorio rptTotaisIPTU, strQueryRelatorio, Trim(dbc_intComposicao.Text) & " - Totalização do Lançamento exercício: " & txt_intExercicio & " Emissão: " & txt_strEmissao
            rptTotaisIPTU.strComposicao = dbc_intComposicao.BoundText
            rptTotaisIPTU.strExercicio = Trim(txt_intExercicio)
            rptTotaisIPTU.strEmissao = Trim(txt_strEmissao)
        End If
    End If
    
    If UCase(strModoOperacao) = UCase(gstrNovo) Then
        Limpa_Controles frmRelTotaisTPTU, True, False, True, True, False
    End If
    
    If UCase(strModoOperacao) = UCase(gstrFechar) Then
        Unload Me
    End If
    
    If UCase(strModoOperacao) = UCase(gstrPreencherLista) Then
        PreencherListaDeOpcoes Me.ActiveControl
    End If
    
End Sub

Private Function strQueryRelatorio() As String
    Dim strSql As String
    
    strSql = ""
    strSql = strSql & "Select "
    strSql = strSql & "1 as Tipo, "
    strSql = strSql & "Count(La.pkid) as TotLancamento, "
    strSql = strSql & "Sum(" & gstrISNULL("LI.Dblareaterreno", "0") & ") As AreaTerreno, "
    strSql = strSql & "Sum(" & gstrISNULL("LI.Dblvalorvenalterreno", "0") & ") AreaVenalTerreno, "
    strSql = strSql & "Sum(" & gstrISNULL("LI.Dblimpostoterreno", "0") & ") ValorImposto, "
    strSql = strSql & "Sum(" & gstrISNULL("LI.Dblareaexcedente", "0") & ") AreaExcedente, "
    strSql = strSql & "Sum(" & gstrISNULL("LI.Dblvalorterrenoexcedente", "0") & ") AreaTerrenoExcedente, "
    strSql = strSql & "Sum(" & gstrISNULL("LI.Dblimpostoexcedente", "0") & ") ValorImpostoExcedente "
    strSql = strSql & "From "
    strSql = strSql & gstrComposicaoDaReceita & " CR, "
    strSql = strSql & gstrLancamentoAlfa & " LA, "
    strSql = strSql & gstrLancamentoIPTU & " LI "
    strSql = strSql & "Where "
    strSql = strSql & "CR.Pkid = LA.Intcomposicaodareceita AND "
    strSql = strSql & "LA.Pkid = LI.Intlancamentoalfa AND "
    strSql = strSql & "CR.Pkid = " & dbc_intComposicao.BoundText & " And "
    strSql = strSql & "LA.strEmissao = " & Format$(Trim(txt_strEmissao), "000") & " And "
    strSql = strSql & "LA.Intexercicio = " & Trim(txt_intExercicio) & " "
    
    strSql = strSql & "Union "
    
    strSql = strSql & "Select "
    strSql = strSql & "2 as Tipo, "
    strSql = strSql & "Count(La.pkid) as TotLancamento, "
    strSql = strSql & "Sum(" & gstrISNULL("LI.Dblareaterreno", "0") & ") As AreaTerreno, "
    strSql = strSql & "Sum(" & gstrISNULL("LI.Dblvalorvenalterreno", "0") & ") AreaVenalTerreno, "
    strSql = strSql & "Sum(" & gstrISNULL("LI.Dblimpostoterreno", "0") & ") ValorImposto, "
    strSql = strSql & "Sum(" & gstrISNULL("LI.Dblareaexcedente", "0") & ") AreaExcedente, "
    strSql = strSql & "Sum(" & gstrISNULL("LI.Dblvalorterrenoexcedente", "0") & ") AreaTerrenoExcedente, "
    strSql = strSql & "Sum(" & gstrISNULL("LI.Dblimpostoexcedente", "0") & ") ValorImpostoExcedente "
    strSql = strSql & "From "
    strSql = strSql & gstrComposicaoDaReceita & " CR, "
    strSql = strSql & gstrLancamentoAlfa & " LA, "
    strSql = strSql & gstrLancamentoIPTU & " LI "
    strSql = strSql & "Where "
    strSql = strSql & "CR.Pkid = LA.Intcomposicaodareceita AND "
    strSql = strSql & "LA.Pkid = LI.Intlancamentoalfa AND "
    strSql = strSql & "CR.Pkid = " & dbc_intComposicao.BoundText & " And "
    strSql = strSql & "LA.strEmissao = " & Format$(Trim(txt_strEmissao), "000") & " And "
    strSql = strSql & "LA.Intexercicio = " & Trim(txt_intExercicio) & " "
    strSql = strSql & "Order By Tipo"
    
    strQueryRelatorio = strSql
End Function

Private Function blnDadosOk() As Boolean
    blnDadosOk = False
    If Not dbc_intComposicao.MatchedWithList Then
        ExibeMensagem "O campo de composição da receita não foi preenchido corretamente."
        dbc_intComposicao.SetFocus
        Exit Function
    ElseIf Trim(txt_intExercicio) = "" Then
        ExibeMensagem "O campo de exercício não foi preenchido corretamente."
        txt_intExercicio.SetFocus
        Exit Function
    ElseIf Trim(txt_strEmissao) = "" Then
        ExibeMensagem "O campo de emissão não foi preenchido corretamente."
        txt_strEmissao.SetFocus
        Exit Function
    End If
    blnDadosOk = True
End Function

Private Function strQueryComposicao() As String
    Dim strSql As String
    
    strSql = "SELECT Pkid,"
    strSql = strSql & gstrCONVERT(CDT_VARCHAR, "intCodigo") & strCONCAT & "' - '" & strCONCAT & " strDescricao Descricao"
    strSql = strSql & " FROM "
    strSql = strSql & gstrComposicaoDaReceita
    strSql = strSql & " WHERE"
    strSql = strSql & " intUtilizacao in (" & TYP_IMOBILIARIA & "," & TYP_ISS_CONSTRUCAO & ") "
    strSql = strSql & " ORDER BY intCodigo"

    strQueryComposicao = strSql

End Function

Private Sub txt_intExercicio_GotFocus()
    MarcaCampo txt_intExercicio
End Sub

Private Sub txt_intExercicio_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txt_intExercicio
End Sub

Private Sub txt_strEmissao_GotFocus()
    MarcaCampo txt_strEmissao
End Sub

Private Sub txt_strEmissao_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txt_strEmissao
End Sub
