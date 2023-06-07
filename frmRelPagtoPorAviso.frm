VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmRelPagtoPorAviso 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relatóriode Pagamentos por Aviso"
   ClientHeight    =   2100
   ClientLeft      =   3465
   ClientTop       =   3510
   ClientWidth     =   6630
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2100
   ScaleWidth      =   6630
   Begin TabDlg.SSTab tab_3dPasta 
      Height          =   2025
      Left            =   60
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   30
      Width           =   6525
      _ExtentX        =   11509
      _ExtentY        =   3572
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Totalização de Lancamento de IPTU"
      TabPicture(0)   =   "frmRelPagtoPorAviso.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Fra_Pagamentos"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin VB.Frame Fra_Pagamentos 
         Height          =   1545
         Left            =   90
         TabIndex        =   1
         Top             =   360
         Width           =   6315
         Begin VB.TextBox txt_intExercicio 
            Height          =   315
            Left            =   1860
            MaxLength       =   4
            TabIndex        =   3
            Top             =   660
            Width           =   525
         End
         Begin VB.TextBox txt_strNumeroAviso 
            Height          =   315
            Left            =   1860
            MaxLength       =   10
            TabIndex        =   4
            Top             =   1005
            Width           =   1035
         End
         Begin MSDataListLib.DataCombo dbc_intComposicao 
            Height          =   315
            Left            =   1860
            TabIndex        =   2
            Top             =   300
            Width           =   4365
            _ExtentX        =   7699
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin VB.Label lbl_Aviso 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Aviso"
            Height          =   195
            Left            =   1395
            TabIndex        =   7
            Top             =   1125
            Width           =   390
         End
         Begin VB.Label lbl_Exercicio 
            AutoSize        =   -1  'True
            Caption         =   "Exercício"
            Height          =   195
            Left            =   1110
            TabIndex        =   6
            Top             =   750
            Width           =   675
         End
         Begin VB.Label lbl_Composicao 
            AutoSize        =   -1  'True
            Caption         =   "Composição da Receita"
            Height          =   195
            Left            =   90
            TabIndex        =   5
            Top             =   390
            Width           =   1695
         End
      End
   End
End
Attribute VB_Name = "frmRelPagtoPorAviso"
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
    gintCodSeguranca = 1267
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
        If blnDadosOK Then
            ImprimeRelatorio rptPagtoPorAviso, strQueryRelatorio, "Consulta de Recebimentos"
        End If
    End If
    
    If UCase(strModoOperacao) = UCase(gstrNovo) Then
        Limpa_Controles frmRelTotaisTPTU, True, False, True, True, False
        dbc_intComposicao.SetFocus
    End If
    
    If UCase(strModoOperacao) = UCase(gstrFechar) Then
        Unload Me
    End If
    
    If UCase(strModoOperacao) = UCase(gstrPreencherLista) Then
        PreencherListaDeOpcoes Me.ActiveControl
    End If
    
End Sub

Private Function strQueryRelatorio() As String
    Dim strSQL As String
    
    strSQL = ""
    strSQL = strSQL & "SELECT "
    strSQL = strSQL & "CB.Pkid As IntBanco, "
    strSQL = strSQL & gstrCONVERT(CDT_VARCHAR, "CB.Intnumeroconta") & strCONCAT & "' '" & strCONCAT & " CB.strdescricao " & strCONCAT & "' '" & strCONCAT & " CB.strConta " & strCONCAT & " CB.strdigitoverificador" & strCONCAT & "' ' AS StrContaBancaria,"
    strSQL = strSQL & "MB.Dtmdtpagamento, "
    strSQL = strSQL & "MB.Dtmdtmovimento, "
    strSQL = strSQL & "CR.Pkid IntComposicao, "
    strSQL = strSQL & "CR.Strdescricao As strComposicao, "
    strSQL = strSQL & "LA.Intexercicio, "
    strSQL = strSQL & "LA.Strnumeroaviso AS strAviso, "
    strSQL = strSQL & "LV.Intparcela, "
    strSQL = strSQL & "B.Strabreviatura As strTipoBaixa, "
    strSQL = strSQL & "(" & gstrISNULL("MB.Dblprincipal", "0") & " + " & gstrISNULL("MB.Dblmulta", "0") & " + " & gstrISNULL("MB.Dbljuros", "0") & " + " & gstrISNULL("MB.Dblcorrecao", "0") & ")  as DblTotal "
    strSQL = strSQL & "FROM "
    strSQL = strSQL & gstrComposicaoDaReceita & " CR, "
    strSQL = strSQL & gstrLancamentoAlfa & " LA, "
    strSQL = strSQL & gstrLancamentoValor & " LV, "
    strSQL = strSQL & gstrLancamentoPagamento & " LP, "
    strSQL = strSQL & gstrCodigoDeBaixa & " B, "
    strSQL = strSQL & gstrMovimentoBancario & " MB, "
    strSQL = strSQL & gstrContaBancaria & " CB "
    strSQL = strSQL & "WHERE "
    strSQL = strSQL & "CR.Pkid = LA.Intcomposicaodareceita  AND "
    strSQL = strSQL & "LA.Pkid = LV.Intlancamentoalfa       AND "
    strSQL = strSQL & "LV.Pkid = LP.Intlancamentovalor      AND "
    strSQL = strSQL & "B.Pkid  = LP.Intcodigobaixa          AND "
    strSQL = strSQL & "LV.Pkid = MB.Intlancamentovalor" & strOUTJOracle & " AND "
    strSQL = strSQL & "MB.Intcontabancaria " & strOUTJSQLServer & "= CB.Pkid" & strOUTJOracle & " AND "
    strSQL = strSQL & "CR.Pkid = " & dbc_intComposicao.BoundText & " AND "
    strSQL = strSQL & "LA.Intexercicio = " & txt_intExercicio & " AND "
    strSQL = strSQL & "LA.Strnumeroaviso = '" & String(gintLenNumAviso - Len(txt_strNumeroAviso), "0") & txt_strNumeroAviso.Text & "'"
    strSQL = strSQL & "GROUP BY  "
    strSQL = strSQL & "CB.Pkid,  "
    strSQL = strSQL & "CB.Intnumeroconta, "
    strSQL = strSQL & "CB.strdescricao, "
    strSQL = strSQL & "CB.strConta, "
    strSQL = strSQL & "CB.strdigitoverificador, "
    strSQL = strSQL & "MB.Dblprincipal, "
    strSQL = strSQL & "MB.DblMulta, "
    strSQL = strSQL & "MB.DblJuros, "
    strSQL = strSQL & "MB.DblCorrecao, "
    strSQL = strSQL & "MB.Dtmdtpagamento, "
    strSQL = strSQL & "MB.Dtmdtmovimento, "
    strSQL = strSQL & "CR.Pkid, "
    strSQL = strSQL & "CR.Strdescricao, "
    strSQL = strSQL & "LA.Intexercicio, "
    strSQL = strSQL & "LA.Strnumeroaviso, "
    strSQL = strSQL & "LV.Intparcela, "
    strSQL = strSQL & "B.Strabreviatura "
    strSQL = strSQL & "Order By MB.dtmdtmovimento"
    
    strQueryRelatorio = strSQL
End Function

Private Function blnDadosOK() As Boolean
    blnDadosOK = False
    If Not dbc_intComposicao.MatchedWithList Then
        ExibeMensagem "O campo de composição da receita não foi preenchido corretamente."
        dbc_intComposicao.SetFocus
        Exit Function
    ElseIf Trim(txt_intExercicio) = "" Then
        ExibeMensagem "O campo de exercício não foi preenchido corretamente."
        txt_intExercicio.SetFocus
        Exit Function
    ElseIf Trim(txt_strNumeroAviso) = "" Then
        ExibeMensagem "O campo do número do aviso não foi preenchido corretamente."
        txt_strNumeroAviso.SetFocus
        Exit Function
    End If
    blnDadosOK = True
End Function

Private Function strQueryComposicao() As String
    Dim strSQL As String
    
    strSQL = "SELECT Pkid,"
    strSQL = strSQL & gstrCONVERT(CDT_VARCHAR, "intCodigo") & strCONCAT & "' - '" & strCONCAT & " strDescricao Descricao"
    strSQL = strSQL & " FROM "
    strSQL = strSQL & gstrComposicaoDaReceita
    strSQL = strSQL & " ORDER BY intCodigo"

    strQueryComposicao = strSQL

End Function

Private Sub txt_intExercicio_GotFocus()
    MarcaCampo txt_intExercicio
End Sub

Private Sub txt_intExercicio_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txt_intExercicio
End Sub

Private Sub txt_strNumeroAviso_GotFocus()
    MarcaCampo txt_strNumeroAviso
End Sub

Private Sub txt_strNumeroAviso_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txt_strNumeroAviso
End Sub

