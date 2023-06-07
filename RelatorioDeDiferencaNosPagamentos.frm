VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmRelatorioDeDiferencaNosPagamentos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relatório de Diferença nos Pagamentos"
   ClientHeight    =   3105
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7410
   HelpContextID   =   682
   Icon            =   "RelatorioDeDiferencaNosPagamentos.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3105
   ScaleWidth      =   7410
   Begin TabDlg.SSTab tab_3dPasta 
      Height          =   2865
      Left            =   120
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   7155
      _ExtentX        =   12621
      _ExtentY        =   5054
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Relatório de Diferença nos Pagamentos"
      TabPicture(0)   =   "RelatorioDeDiferencaNosPagamentos.frx":1042
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin VB.Frame Frame1 
         Height          =   2265
         Left            =   195
         TabIndex        =   1
         Top             =   360
         Width           =   6795
         Begin VB.TextBox txt_strCodigo 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#.##0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   1
            EndProperty
            Height          =   285
            Left            =   1590
            MaxLength       =   6
            TabIndex        =   5
            Top             =   780
            Width           =   975
         End
         Begin VB.TextBox txt_strInscricaoCadastral 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#.##0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   1
            EndProperty
            Height          =   285
            Left            =   4095
            MaxLength       =   15
            TabIndex        =   4
            Top             =   780
            Width           =   1935
         End
         Begin VB.ComboBox cbo_intUtilizacaoDebito 
            Height          =   315
            ItemData        =   "RelatorioDeDiferencaNosPagamentos.frx":105E
            Left            =   1590
            List            =   "RelatorioDeDiferencaNosPagamentos.frx":1060
            TabIndex        =   3
            Top             =   420
            Width           =   5055
         End
         Begin VB.TextBox txtDtInicial 
            Height          =   285
            Left            =   1590
            MaxLength       =   4
            TabIndex        =   2
            Top             =   1470
            Width           =   1065
         End
         Begin MSDataListLib.DataCombo dbc_intContribuinte 
            Height          =   315
            Left            =   1590
            TabIndex        =   6
            Top             =   1110
            Width           =   5055
            _ExtentX        =   8916
            _ExtentY        =   556
            _Version        =   393216
            IntegralHeight  =   0   'False
            Text            =   ""
         End
         Begin VB.Label lbl_Label1 
            AutoSize        =   -1  'True
            Caption         =   "Código/Contribuinte"
            Height          =   195
            Left            =   120
            TabIndex        =   11
            Top             =   1230
            Width           =   1410
         End
         Begin VB.Label lbl_strCodigo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Código"
            Height          =   195
            Left            =   1035
            TabIndex        =   10
            Top             =   870
            Width           =   495
         End
         Begin VB.Label lbl_strInscricao 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Inscrição Cadastral"
            Height          =   195
            Left            =   2640
            TabIndex        =   9
            Top             =   870
            Width           =   1350
         End
         Begin VB.Label lbl_intUtilizacaoDebito 
            AutoSize        =   -1  'True
            Caption         =   "Utilização"
            Height          =   195
            Left            =   840
            TabIndex        =   8
            Top             =   540
            Width           =   690
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Exercício"
            Height          =   195
            Left            =   855
            TabIndex        =   7
            Top             =   1560
            Width           =   675
         End
      End
   End
End
Attribute VB_Name = "frmRelatorioDeDiferencaNosPagamentos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub dbc_intContribuinte_Click(Area As Integer)
    DropDownDataCombo dbc_intContribuinte, Me, Area
    If Area = 0 Then
        If Trim(dbc_intContribuinte.Text) <> "" And Not dbc_intContribuinte.MatchedWithList Then
            MantemForm gstrPreencherLista
        End If
    End If
End Sub

Private Sub dbc_intContribuinte_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbc_intContribuinte, Me, , KeyCode, Shift
End Sub

Private Sub Form_Activate()
    gintCodSeguranca = 682
    
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrDeletar
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrSalvar
    HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrImprimir
End Sub

Private Sub Form_Deactivate()
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar
End Sub

Private Sub Form_Load()
cbo_intUtilizacaoDebito.AddItem "Imobiliárias "
cbo_intUtilizacaoDebito.ItemData(cbo_intUtilizacaoDebito.NewIndex) = "1"
cbo_intUtilizacaoDebito.AddItem "Econômicas"
cbo_intUtilizacaoDebito.ItemData(cbo_intUtilizacaoDebito.NewIndex) = "2"
cbo_intUtilizacaoDebito.AddItem "Fiscalização"
cbo_intUtilizacaoDebito.ItemData(cbo_intUtilizacaoDebito.NewIndex) = "3"
cbo_intUtilizacaoDebito.AddItem "Outras Receitas"
cbo_intUtilizacaoDebito.ItemData(cbo_intUtilizacaoDebito.NewIndex) = "4"

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar
End Sub

Public Sub MantemForm(ByVal strModoOperacao As String)
Dim strSQL As String

If strModoOperacao = gstrImprimir Then
    strSQL = strQueryRelatorio
    
    If cbo_intUtilizacaoDebito.ListIndex = -1 Then
        ImprimeRelatorio rptRelatorioDeDiferencaNosPagamentos, strSQL, "Relação de Diferença nos Pagamentos"
    ElseIf cbo_intUtilizacaoDebito.ListIndex = 0 Then
        ImprimeRelatorio rptRelatorioDeDiferencaNosPagamentos, strSQL, "Relação de Diferença nos Pagamentos dos Imóveis"
    ElseIf cbo_intUtilizacaoDebito.ListIndex = 1 Then
        ImprimeRelatorio rptRelatorioDeDiferencaNosPagamentos, strSQL, "Relação de Diferença nos Pagamentos dos Apêndices"
    End If
ElseIf strModoOperacao = gstrPreencherLista Then
    strSQL = ""
    strSQL = strSQL & " SELECT PKId, strNome FROM " & gstrContribuinte
    If IsNumeric(dbc_intContribuinte.Text) Then
        strSQL = strSQL & " WHERE strCodigoAnterior = '" & dbc_intContribuinte.Text & "'"
    ElseIf Not dbc_intContribuinte.MatchedWithList Then
        strSQL = strSQL & " WHERE strNome LIKE '" & dbc_intContribuinte.Text & "%'"
    End If
    
    LeDaTabelaParaObj "", dbc_intContribuinte, strSQL
End If
End Sub

Private Function strQueryRelatorio() As String

'******************************************************************************************
' Data: 02/05/2003
' Alteração: - Substituição do comando nativo ISNULL() do SQL Server pela função gstrISNULL
' Responsável: Everton Bianchini
'------------------------------------------------------------------------------------------
' Data: 02/05/2003
' Alteração: - Substituição do comando nativo CONVERT do SQL Server pela função gstrCONVERT
' Responsável: Everton Bianchini
'------------------------------------------------------------------------------------------
' Data: 05/05/2003
' Alteração: - Substituição do comando nativo de concatenação de strings (+) do SQL Server
'            pela variável strCONCAT.
' Responsável: Everton Bianchini
'------------------------------------------------------------------------------------------
' Data: 15/05/2003
' Alteração: - Foi comentado o comando SELECT devido à incompatibilidade entre o SQL Server
'            e o Oracle. O Oracle não permite a abertura de múltiplos recordsets em um
'            único objeto ADODB.Recordset.
' Responsável: Everton Bianchini
'******************************************************************************************
    
Dim strSQL As String

If cbo_intUtilizacaoDebito.ListIndex = -1 Then
    strSQL = ""
    strSQL = strSQL & " SELECT C.strCodigoAnterior, IM.strCodigo, A.intNumeroParcela, A.dtmDataVencimento, "
    strSQL = strSQL & " A.dtmDataPagamento, 'Manual' AS strTipoPagamento, A.dblTotalPago,"
'    strSql = strSql & " (A.dblValorParcela+0+ISNULL(A.dblJuros,0)+ISNULL(A.dblMulta,0)) AS dblDevido"
    strSQL = strSQL & " (A.dblValorParcela+0+ " & gstrISNULL("A.dblJuros", "0") & "+" & gstrISNULL("A.dblMulta", "0") & ") AS dblDevido, "
    strSQL = strSQL & " (A.dblTotalPago - (A.dblValorParcela+0+A.dblJuros+A.dblMulta)) AS dblDiferenca, "
    strSQL = strSQL & " IM.strInscricaoAnterior AS strInscricao, C.strNome, "
'    strSql = strSql & " (CONVERT(NVARCHAR,D.intCodigo) + ' - ' + D.strSigla) AS strSigla "
    strSQL = strSQL & " (" & gstrCONVERT(CDT_NVARCHAR, "D.intCodigo") & strCONCAT & " ' - ' " & strCONCAT & " D.strSigla) AS strSigla "
    strSQL = strSQL & " FROM "
    strSQL = strSQL & gstrParcelaReceita & " A, "
    strSQL = strSQL & gstrLancamentoCalculo & " B, "
    strSQL = strSQL & gstrContribuinte & " C,"
    strSQL = strSQL & gstrComposicaoDaReceita & " D, "
    strSQL = strSQL & gstrImobiliario & " IM "
    strSQL = strSQL & " WHERE  A.intLancamentoCalculo = B.PKId"
    strSQL = strSQL & " AND C.PKId = IM.intContribuinte"
    strSQL = strSQL & " AND IM.strInscricaoAnterior = B.strInscricaoCadastral"
    strSQL = strSQL & " AND B.intContribuinte = C.PKId"
    strSQL = strSQL & " AND A.intComposicaoDaReceita = D.PKId"
    strSQL = strSQL & " AND A.strSituacao = 'P'"
    
    If Trim(txt_strCodigo.Text) <> "" Then
        strSQL = strSQL & " AND IM.strCodigo = '" & txt_strCodigo.Text & "'"
    End If
    If Trim(txt_strInscricaoCadastral.Text) <> "" Then
        strSQL = strSQL & " AND IM.strInscricaoAnterior = '" & txt_strInscricaoCadastral.Text & "'"
    End If

'    strSQL = strSQL & " SELECT C.strCodigoAnterior, C.strCodigoAnterior AS strCodigo, A.intNumeroParcela, A.dtmDataVencimento, "
'    strSQL = strSQL & " B.dtmLancamento, A.dtmDataPagamento, 'Normal' AS strTipoPagamento, A.dblValorParcela, "
'    strSQL = strSQL & " 0 AS dblTotalCorrecao, A.dblJuros, A.dblMulta, A.dblTotalPago,"
'    strSQL = strSQL & " (A.dblTotalPago - (A.dblValorParcela+0+A.dblJuros+A.dblMulta)) AS dblDiferenca, C.strNome "
'    strSQL = strSQL & " FROM "
'    strSQL = strSQL & gstrParcelaReceita & " A, "
'    strSQL = strSQL & gstrLancamentoCalculo & " B, "
'    strSQL = strSQL & gstrContribuinte & " C,"
'    strSQL = strSQL & gstrComposicaoDaReceita & " D, "
'    strSQL = strSQL & gstrEconomico & " EC "
'    strSQL = strSQL & " WHERE  A.intLancamentoCalculo = B.PKId"
'    strSQL = strSQL & " AND C.PKId = EC.intContribuinte"
'    strSQL = strSQL & " AND EC.strInscricaoCadastral = B.strInscricaoCadastral"
'    strSQL = strSQL & " AND B.intContribuinte = C.PKId"
'    strSQL = strSQL & " AND A.intComposicaoDaReceita = D.PKId"
'    strSQL = strSQL & " AND A.strSituacao = 'P'"
'
'    If Trim(txt_strCodigo.Text) <> "" Then
'        strSQL = strSQL & " AND EC.strInscricaoCadastral = '" & txt_strCodigo.Text & "'"
'    End If
'    If Trim(txt_strInscricaoCadastral.Text) <> "" Then
'        strSQL = strSQL & " AND B.strInscricaoCadastral = '" & txt_strInscricaoCadastral.Text & "'"
'    End If

ElseIf cbo_intUtilizacaoDebito.ListIndex = 0 Then 'Imobiliaria

    strSQL = ""
    strSQL = strSQL & " SELECT C.strCodigoAnterior, IM.strCodigo, A.intNumeroParcela, A.dtmDataVencimento, "
    strSQL = strSQL & " A.dtmDataPagamento, 'Manual' AS strTipoPagamento, A.dblTotalPago,"
'    strSql = strSql & " (A.dblValorParcela+0+ISNULL(A.dblJuros,0)+ISNULL(A.dblMulta,0)) AS dblDevido, "
    strSQL = strSQL & " (A.dblValorParcela+0+" & gstrISNULL("A.dblJuros", "0") & "+" & gstrISNULL("A.dblMulta", "0") & ") AS dblDevido, "
    strSQL = strSQL & " (A.dblTotalPago - (A.dblValorParcela+0+A.dblJuros+A.dblMulta)) AS dblDiferenca, "
    strSQL = strSQL & " IM.strInscricaoAnterior AS strInscricao, C.strNome, "
'    strSql = strSql & " (CONVERT(NVARCHAR,D.intCodigo) + ' - ' + D.strSigla) AS strSigla "
    strSQL = strSQL & " (" & gstrCONVERT(CDT_NVARCHAR, "D.intCodigo") & strCONCAT & " ' - ' " & strCONCAT & " D.strSigla) AS strSigla "
    strSQL = strSQL & " FROM "
    strSQL = strSQL & gstrParcelaReceita & " A, "
    strSQL = strSQL & gstrLancamentoCalculo & " B, "
    strSQL = strSQL & gstrContribuinte & " C,"
    strSQL = strSQL & gstrComposicaoDaReceita & " D, "
    strSQL = strSQL & gstrImobiliario & " IM "
    strSQL = strSQL & " WHERE  A.intLancamentoCalculo = B.PKId"
    strSQL = strSQL & " AND C.PKId = IM.intContribuinte"
    strSQL = strSQL & " AND IM.strInscricaoAnterior = B.strInscricaoCadastral"
    strSQL = strSQL & " AND B.intContribuinte = C.PKId"
    strSQL = strSQL & " AND A.intComposicaoDaReceita = D.PKId"
    strSQL = strSQL & " AND A.strSituacao = 'P'"
    strSQL = strSQL & " AND (A.dblTotalPago - (A.dblValorParcela+0+A.dblJuros+A.dblMulta)) <> 0"
    If Trim(txt_strCodigo.Text) <> "" Then
        strSQL = strSQL & " AND IM.strCodigo = '" & txt_strCodigo.Text & "'"
    End If
    If Trim(txt_strInscricaoCadastral.Text) <> "" Then
        strSQL = strSQL & " AND IM.strInscricaoAnterior = '" & txt_strInscricaoCadastral.Text & "'"
    End If

ElseIf cbo_intUtilizacaoDebito.ListIndex = 1 Then 'Econômica
    strSQL = ""
    strSQL = strSQL & " SELECT C.strCodigoAnterior, C.strCodigoAnterior AS strCodigo, A.intNumeroParcela, A.dtmDataVencimento, "
    strSQL = strSQL & " B.dtmLancamento, A.dtmDataPagamento, 'Normal' AS strTipoPagamento, A.dblValorParcela, "
    strSQL = strSQL & " 0 AS dblTotalCorrecao, A.dblJuros, A.dblMulta, A.dblTotalPago,"
    strSQL = strSQL & " (A.dblTotalPago - (A.dblValorParcela+0+A.dblJuros+A.dblMulta)) AS dblDiferenca, C.strNome "
    strSQL = strSQL & " FROM "
    strSQL = strSQL & gstrParcelaReceita & " A, "
    strSQL = strSQL & gstrLancamentoCalculo & " B, "
    strSQL = strSQL & gstrContribuinte & " C,"
    strSQL = strSQL & gstrComposicaoDaReceita & " D, "
    strSQL = strSQL & gstrEconomico & " EC "
    strSQL = strSQL & " WHERE  A.intLancamentoCalculo = B.PKId"
    strSQL = strSQL & " AND C.PKId = EC.intContribuinte"
    strSQL = strSQL & " AND EC.strInscricaoCadastral = B.strInscricaoCadastral"
    strSQL = strSQL & " AND B.intContribuinte = C.PKId"
    strSQL = strSQL & " AND A.intComposicaoDaReceita = D.PKId"
    strSQL = strSQL & " AND A.strSituacao = 'P'"

    If Trim(txt_strCodigo.Text) <> "" Then
        strSQL = strSQL & " AND EC.strInscricaoCadastral = '" & txt_strCodigo.Text & "'"
    End If
    If Trim(txt_strInscricaoCadastral.Text) <> "" Then
        strSQL = strSQL & " AND B.strInscricaoCadastral = '" & txt_strInscricaoCadastral.Text & "'"
    End If
End If

If Trim(txtDtInicial.Text) <> "" Then
    strSQL = strSQL & " AND B.intExercicio = " & txtDtInicial.Text
End If

If dbc_intContribuinte.MatchedWithList Then
    strSQL = strSQL & " AND C.PKId = " & dbc_intContribuinte.BoundText
End If

strSQL = strSQL & " ORDER BY strNome, B.strInscricaoCadastral"

strQueryRelatorio = strSQL
End Function

Private Sub txtDtInicial_KeyPress(KeyAscii As Integer)
CaracterValido KeyAscii, "N", txtDtInicial
End Sub
