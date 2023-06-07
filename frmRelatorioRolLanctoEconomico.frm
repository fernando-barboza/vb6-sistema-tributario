VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Begin VB.Form frmRelatorioRolLanctoEconomico 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Rol de Lançamentos Econômico"
   ClientHeight    =   1290
   ClientLeft      =   1965
   ClientTop       =   1755
   ClientWidth     =   6120
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1290
   ScaleWidth      =   6120
   Begin VB.Frame fra_ComposicaoDaReceita 
      Height          =   960
      Left            =   150
      TabIndex        =   0
      Top             =   120
      Width           =   5790
      Begin VB.CommandButton cmd_Composicao 
         Height          =   300
         Left            =   4050
         Picture         =   "frmRelatorioRolLanctoEconomico.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   1
         TabStop         =   0   'False
         Tag             =   "590"
         ToolTipText     =   "Ativa Cadastro de Composição da Receita"
         Top             =   480
         Width           =   360
      End
      Begin MSDataListLib.DataCombo dbc_intComposicao 
         Height          =   315
         Left            =   150
         TabIndex        =   2
         Top             =   480
         Width           =   3915
         _ExtentX        =   6906
         _ExtentY        =   556
         _Version        =   393216
         IntegralHeight  =   0   'False
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dbc_intExercicioInicial 
         Height          =   315
         Left            =   4710
         TabIndex        =   4
         Top             =   480
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin VB.Label lbl_ExercicioInicial 
         AutoSize        =   -1  'True
         Caption         =   "Exercício"
         Height          =   195
         Left            =   4740
         TabIndex        =   5
         Top             =   240
         Width           =   675
      End
      Begin VB.Label lbl_Composicao 
         AutoSize        =   -1  'True
         Caption         =   "Composição da Receita"
         Height          =   195
         Left            =   165
         TabIndex        =   3
         Top             =   240
         Width           =   1695
      End
   End
End
Attribute VB_Name = "frmRelatorioRolLanctoEconomico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub dbc_intComposicao_Change()

    If dbc_intComposicao.MatchedWithList Then
        PreencheExercicio dbc_intExercicioInicial
    End If

End Sub

Private Sub dbc_intComposicao_GotFocus()
    MarcaCampo dbc_intComposicao
End Sub
    
Private Sub dbc_intComposicao_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbc_intComposicao, Me, , KeyCode, Shift
End Sub

Private Sub dbc_intComposicao_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbc_intComposicao
End Sub

Private Sub dbc_intExercicioInicial_Click(Area As Integer)
    DropDownDataCombo dbc_intExercicioInicial, Me, Area
End Sub

Private Sub dbc_intExercicioInicial_GotFocus()
    MarcaCampo dbc_intExercicioInicial
End Sub

Private Sub dbc_intExercicioInicial_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbc_intExercicioInicial, Me, , KeyCode, Shift
End Sub

Private Sub dbc_intExercicioInicial_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", dbc_intExercicioInicial
End Sub

Private Sub Form_Activate()
    gintCodSeguranca = 1431
    HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrImprimir, gstrSalvar, gstrDeletar
End Sub

Private Sub Form_Load()

    dbc_intComposicao.Tag = strQueryComposicao & ";strDescricao"

End Sub

Public Sub MantemForm(ByVal strModoOperacao As String)

    Select Case UCase(strModoOperacao)
        
        Case Is = UCase(gstrPreencherLista)
                PreencherListaDeOpcoes Me.ActiveControl
                
        Case Is = UCase(gstrImprimir)
            If blnDadosOk Then
                ImprimeRelatorio rptRelatorioRolLanctoEconomico, strQueryRelatorio, "Rol de Lançamentos do Cadastro Econômico"
            End If
        
    End Select

End Sub

Private Function strQueryComposicao() As String
Dim strSql As String

    strSql = "SELECT Pkid,"
    strSql = strSql & gstrCONVERT(CDT_VARCHAR, "intCodigo") & strCONCAT & "' - '" & strCONCAT & " strDescricao Descricao"
    strSql = strSql & " FROM "
    strSql = strSql & gstrComposicaoDaReceita
    strSql = strSql & " WHERE"
    strSql = strSql & " intUtilizacao in (2) "
    strSql = strSql & " ORDER BY intCodigo"

    strQueryComposicao = strSql

End Function

Private Function strQueryRelatorio() As String
Dim strSql As String
Dim strSQLValorParcela As String

    strSQLValorParcela = "(SELECT " & gstrTOPnSQLServer(1) & " dblValor "
    strSQLValorParcela = strSQLValorParcela & " FROM " & gstrLancamentoValor
    strSQLValorParcela = strSQLValorParcela & " WHERE intLancamentoAlfa = LA.Pkid " & IIf(bytDBType = EDatabases.Oracle, " AND Rownum = 1 ", "") & ") as dblValorParcela "

    strSql = "SELECT LA.Pkid PkidAlfa, LA.intExercicio, LA.strComposicaoDaReceita, LA.intComposicaoDaReceita, LA.strInscricao, LA.strNumeroAviso, LA.strNomeProprietario, "
    strSql = strSql & " LA.strLogradouro, LA.strNumero, LA.strComplemento, LA.strBairro, LA.strMunicipio, LA.strUF, LA.strLogradouroC , LA.strNumeroC, LA.strComplementoC, LA.strBairroC, LA.strMunicipioC, LA.strUFC, "
    strSql = strSql & " SUM(LV.dblValor) dblTotalLancto, "
    strSql = strSql & strSQLValorParcela & ", "
    strSql = strSql & " LE.dblAreaOcupada, LE.dblNumeroEmpregados, LEA.strDescricaoAtividade "
    strSql = strSql & " FROM "
    strSql = strSql & gstrLancamentoAlfa & " LA, " & gstrLancamentoValor & " LV, " & gstrLancamentoEconomico & " LE, " & gstrLctEconomicoAtividade & " LEA "
    strSql = strSql & " WHERE"
    strSql = strSql & " LA.intComposicaoDaReceita = " & dbc_intComposicao.BoundText & " AND LA.intExercicio = " & dbc_intExercicioInicial.BoundText
    strSql = strSql & " AND LV.intLancamentoAlfa = LA.Pkid "
    strSql = strSql & " AND LE.intLancamentoAlfa = LA.Pkid "
    strSql = strSql & " AND LEA.intLancamentoEconomico = LE.Pkid "
    strSql = strSql & " AND LEA.blnPrincipal = 1 "
    strSql = strSql & " GROUP BY LA.Pkid, LA.intExercicio, LA.strComposicaoDaReceita, LA.intComposicaoDaReceita, LA.strInscricao, LA.strNumeroAviso, LA.strNomeProprietario, "
    strSql = strSql & " LA.strLogradouro, LA.strNumero, LA.strComplemento, LA.strBairro, LA.strMunicipio, LA.strUF, LA.strLogradouroC , LA.strNumeroC, LA.strComplementoC, LA.strBairroC, LA.strMunicipioC, LA.strUFC, "
    strSql = strSql & " LE.dblAreaOcupada, LE.dblNumeroEmpregados, LEA.strDescricaoAtividade "
    strSql = strSql & " ORDER BY LA.strInscricao"

    strQueryRelatorio = strSql

End Function

Private Sub PreencheExercicio(dbcExercicio As DataCombo)
Dim strSql As String
Dim adoResultado As ADODB.Recordset
    
    Set dbcExercicio.RowSource = Nothing
    
    strSql = "SELECT DISTINCT intExercicio"
    strSql = strSql & " FROM "
    strSql = strSql & gstrParametroAtualizacao & " " & strREADPAST
    strSql = strSql & " WHERE "
    strSql = strSql & " intComposicaoReceita = " & dbc_intComposicao.BoundText
    strSql = strSql & " ORDER BY intExercicio"

    Set gobjBanco = New clsBanco

    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        If Not adoResultado.EOF Then
            dbcExercicio.ListField = adoResultado.Fields(0).Name
            Set dbcExercicio.RowSource = adoResultado
        End If
    End If

End Sub

Private Function blnDadosOk() As Boolean
    
    blnDadosOk = False
    
    If Not dbc_intComposicao.MatchedWithList Then
        ExibeMensagem "É preciso selecionar alguma Composição da Receita válida."
        dbc_intComposicao.SetFocus
    ElseIf Not dbc_intExercicioInicial.MatchedWithList Then
        ExibeMensagem "É preciso selecionar algum Exercício válido."
        dbc_intExercicioInicial.SetFocus
    End If
    
    blnDadosOk = True
    
End Function
