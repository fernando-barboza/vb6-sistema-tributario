VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmRelatorioParaControleDeInadimplencia 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Controle de Inadimplência Analítico/Sintético"
   ClientHeight    =   2895
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7185
   HelpContextID   =   769
   Icon            =   "RelatorioParaControleDeInadimplencia.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   7185
   Begin TabDlg.SSTab tab_3dPasta 
      Height          =   2865
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   7155
      _ExtentX        =   12621
      _ExtentY        =   5054
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Relatório das Parcelas Lançadas"
      TabPicture(0)   =   "RelatorioParaControleDeInadimplencia.frx":1042
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
         Begin VB.TextBox txtDtInicial 
            Height          =   285
            Left            =   1590
            MaxLength       =   4
            TabIndex        =   5
            Top             =   1470
            Width           =   1065
         End
         Begin VB.ComboBox cbo_intUtilizacaoDebito 
            Height          =   315
            ItemData        =   "RelatorioParaControleDeInadimplencia.frx":105E
            Left            =   1590
            List            =   "RelatorioParaControleDeInadimplencia.frx":1060
            TabIndex        =   4
            Top             =   420
            Width           =   5055
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
            TabIndex        =   3
            Top             =   780
            Width           =   1935
         End
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
            TabIndex        =   2
            Top             =   780
            Width           =   975
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
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Exercício"
            Height          =   195
            Left            =   855
            TabIndex        =   11
            Top             =   1560
            Width           =   675
         End
         Begin VB.Label lbl_intUtilizacaoDebito 
            AutoSize        =   -1  'True
            Caption         =   "Utilização"
            Height          =   195
            Left            =   840
            TabIndex        =   10
            Top             =   540
            Width           =   690
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
         Begin VB.Label lbl_strCodigo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Código"
            Height          =   195
            Left            =   1035
            TabIndex        =   8
            Top             =   870
            Width           =   495
         End
         Begin VB.Label lbl_Label1 
            AutoSize        =   -1  'True
            Caption         =   "Código/Contribuinte"
            Height          =   195
            Left            =   120
            TabIndex        =   7
            Top             =   1230
            Width           =   1410
         End
      End
   End
End
Attribute VB_Name = "frmRelatorioParaControleDeInadimplencia"
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
    gintCodSeguranca = 769
    
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
Dim strSql As String

If strModoOperacao = gstrImprimir Then
    strSql = strQueryRelatorio
    
    If cbo_intUtilizacaoDebito.ListIndex = -1 Then
        ImprimeRelatorio rptControleInadimplenciaSintetico, strSql
    ElseIf cbo_intUtilizacaoDebito.ListIndex = 0 Then
        ImprimeRelatorio rptControleInadimplenciaSintetico, strSql
    ElseIf cbo_intUtilizacaoDebito.ListIndex = 1 Then
        ImprimeRelatorio rptControleInadimplenciaSintetico, strSql
    End If
ElseIf strModoOperacao = gstrPreencherLista Then
    strSql = ""
    strSql = strSql & " SELECT PKId, strNome FROM " & gstrContribuinte
    If IsNumeric(dbc_intContribuinte.Text) Then
        strSql = strSql & " WHERE strCodigoAnterior = '" & dbc_intContribuinte.Text & "'"
    ElseIf Not dbc_intContribuinte.MatchedWithList Then
        strSql = strSql & " WHERE strNome LIKE '" & dbc_intContribuinte.Text & "'"
    End If
    
    LeDaTabelaParaObj "", dbc_intContribuinte, strSql
End If
End Sub

Private Function strQueryRelatorio() As String

'******************************************************************************************
' Data: 02/05/2003
' Alteração: - Substituição do comando nativo CONVERT do SQL Server pela função gstrCONVERT
' Responsável: Everton Bianchini
'******************************************************************************************
    
Dim strSql As String

If cbo_intUtilizacaoDebito.ListIndex = -1 Then
    ImprimeRelatorio rptRelatorioDeParcelasArrecadas, strSql, "Relação das Parcelas Lançadas"
ElseIf cbo_intUtilizacaoDebito.ListIndex = 0 Then 'Imobiliaria
    strSql = ""
    strSql = strSql & " SELECT D.intCodigo, D.strDescricao, C.strCodigoAnterior, C.strNome, SUM(A.dblValorParcela) as dblValor "
    strSql = strSql & " FROM "
    strSql = strSql & gstrParcelaReceita & " A, "
    strSql = strSql & gstrLancamentoCalculo & " B, "
    strSql = strSql & gstrContribuinte & " C,"
    strSql = strSql & gstrComposicaoDaReceita & " D, "
    strSql = strSql & gstrImobiliario & " IM "
    strSql = strSql & " WHERE  A.intLancamentoCalculo = B.PKId"
    strSql = strSql & " AND C.PKId = IM.intContribuinte"
    strSql = strSql & " AND IM.strInscricaoAnterior = B.strInscricaoCadastral"
    strSql = strSql & " AND B.intContribuinte = C.PKId"
    strSql = strSql & " AND A.intComposicaoDaReceita = D.PKId"
    strSql = strSql & " AND A.strSituacao = 'A'"
    strSql = strSql & " AND A.intNumeroParcela <> 0"
    
    If Trim(txt_strCodigo.Text) <> "" Then
        strSql = strSql & " AND IM.strCodigo = '" & txt_strCodigo.Text & "'"
    End If
    If Trim(txt_strInscricaoCadastral.Text) <> "" Then
        strSql = strSql & " AND IM.strInscricaoAnterior = '" & txt_strInscricaoCadastral.Text & "'"
    End If

ElseIf cbo_intUtilizacaoDebito.ListIndex = 1 Then 'Econômica
    strSql = ""
    strSql = strSql & " SELECT D.intCodigo, D.strDescricao, C.strCodigoAnterior, C.strNome, SUM(A.dblValorParcela) as dblValor "
    strSql = strSql & " FROM "
    strSql = strSql & gstrParcelaReceita & " A, "
    strSql = strSql & gstrLancamentoCalculo & " B, "
    strSql = strSql & gstrContribuinte & " C,"
    strSql = strSql & gstrComposicaoDaReceita & " D, "
    strSql = strSql & gstrEconomico & " EC "
    strSql = strSql & " WHERE  A.intLancamentoCalculo = B.PKId"
    strSql = strSql & " AND C.PKId = EC.intContribuinte"
    strSql = strSql & " AND EC.strInscricaoCadastral = B.strInscricaoCadastral"
    strSql = strSql & " AND B.intContribuinte = C.PKId"
    strSql = strSql & " AND A.intComposicaoDaReceita = D.PKId"
    strSql = strSql & " AND A.strSituacao = 'A'"
    strSql = strSql & " AND A.intNumeroParcela <> 0"

    If Trim(txt_strCodigo.Text) <> "" Then
        strSql = strSql & " AND EC.strInscricaoCadastral = '" & txt_strCodigo.Text & "'"
    End If
    If Trim(txt_strInscricaoCadastral.Text) <> "" Then
        strSql = strSql & " AND EC.strInscricaoCadastral = '" & txt_strInscricaoCadastral.Text & "'"
    End If
End If

If Trim(txtDtInicial.Text) <> "" Then
    strSql = strSql & " AND B.intExercicio = " & txtDtInicial.Text
End If

If dbc_intContribuinte.MatchedWithList Then
    strSql = strSql & " AND C.PKId = " & dbc_intContribuinte.BoundText
End If

strSql = strSql & " GROUP BY D.intCodigo, D.strDescricao, C.strCodigoAnterior, C.strNome"
'strSql = strSql & " ORDER BY D.intCodigo, CONVERT(NUMERIC,C.strCodigoAnterior)"
strSql = strSql & " ORDER BY D.intCodigo, " & gstrCONVERT(CDT_NUMERIC, "C.strCodigoAnterior")

strQueryRelatorio = strSql
End Function

Private Sub txtDtInicial_KeyPress(KeyAscii As Integer)
CaracterValido KeyAscii, "N", txtDtInicial
End Sub

