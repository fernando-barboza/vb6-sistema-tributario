VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCadqtdeLancamentoValorTipo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Quantidade de lançamentos, Valor e Tipo"
   ClientHeight    =   1920
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5775
   Icon            =   "CadqtdeLancamentoValorTipo.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1920
   ScaleWidth      =   5775
   Begin TabDlg.SSTab tab_3dPasta 
      Height          =   1725
      Left            =   120
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   60
      Width           =   5520
      _ExtentX        =   9737
      _ExtentY        =   3043
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Lançamentos, Valor e Tipo"
      TabPicture(0)   =   "CadqtdeLancamentoValorTipo.frx":1042
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lbl_Devolucao"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "txt_Ate"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txt_Inicial"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "fra_Tipoo"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      Begin VB.Frame fra_Tipoo 
         Height          =   645
         Left            =   2160
         TabIndex        =   5
         Top             =   360
         Width           =   2865
         Begin VB.OptionButton opt_Tipo 
            Caption         =   "Analítico"
            Height          =   195
            Index           =   0
            Left            =   270
            TabIndex        =   0
            Top             =   270
            Width           =   1035
         End
         Begin VB.OptionButton opt_Tipo 
            Caption         =   "Sintético"
            Height          =   195
            Index           =   1
            Left            =   1620
            TabIndex        =   1
            Top             =   270
            Width           =   1035
         End
      End
      Begin VB.TextBox txt_Inicial 
         Height          =   285
         Left            =   2160
         MaxLength       =   10
         TabIndex        =   2
         Top             =   1200
         Width           =   975
      End
      Begin VB.TextBox txt_Ate 
         Height          =   285
         Left            =   4050
         MaxLength       =   10
         TabIndex        =   3
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label lbl_Devolucao 
         AutoSize        =   -1  'True
         Caption         =   "Data Inicial"
         Height          =   195
         Left            =   1230
         TabIndex        =   7
         Top             =   1290
         Width           =   795
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Até"
         Height          =   195
         Left            =   3690
         TabIndex        =   6
         Top             =   1290
         Width           =   240
      End
   End
End
Attribute VB_Name = "frmCadqtdeLancamentoValorTipo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim mblnAlterando    As Boolean
    Dim mobjAux          As Object
    Dim mblnSelecionou   As Boolean
    Dim mblnPrimeiraVez  As Boolean

Private Sub Form_Activate()
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
    HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrSalvar, gstrDeletar
    txt_Ate = gstrDataFormatada(gstrDataDoSistema)
    opt_Tipo(0).Value = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar
    mblnSelecionou = False
    mblnPrimeiraVez = False
End Sub
Public Sub MantemForm(ByVal strModoOperacao As String)
Dim adoRelatorio   As ADODB.Recordset
On Error GoTo ErroImprimeRelatorio
    
    If UCase(strModoOperacao) = UCase(gstrImprimir) Then
        If blnDadosOK = False Then
            Exit Sub
        End If
    End If
    
    'SÍNTËTÌCÔ ®
    If UCase(strModoOperacao) = UCase(gstrImprimir) And opt_Tipo(1).Value = True Then
        ImprimeRelatorio rptQtdLancamentoSintetico, strQuerrySintetico
    End If
    
    'ÁNÁLÌTÌCÔ ®
    If UCase(strModoOperacao) = UCase(gstrImprimir) And opt_Tipo(0).Value = True Then
        Screen.MousePointer = vbHourglass
        Set gobjBanco = New clsBanco
        If gobjBanco.CriaADO(strQuerryAnalitico, 5, adoRelatorio) Then
            Set rptQtdLancamentoAnalitico.adoDataControl.Recordset = adoRelatorio
'            adoRelatorio.Close
'            gobjBanco.CriaADO strQuerryAnalitico2, 5, adoRelatorio
'            Set rptDevolucaoAnalitico.adoDataControl1.Recordset = adoRelatorio
            
            rptQtdLancamentoAnalitico.Show
        End If
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    If UCase(strModoOperacao) = UCase(gstrNovo) Then
        LimpaObjetos
    End If
    
    If UCase(strModoOperacao) = UCase(gstrFechar) Then
        Unload Me
    End If
    
Screen.MousePointer = vbDefault
ErroImprimeRelatorio:
Resume FimImprimeRelatorio
            
FimImprimeRelatorio:
End Sub

Private Function blnDadosOK() As Boolean
blnDadosOK = False

    If txt_Inicial = "" Then
        ExibeMensagem "A data inicial tem que ser digitada."
        txt_Inicial.SetFocus
        Exit Function
    Else
        If gblnDataValida(txt_Inicial.Text) = False Then
            ExibeMensagem "A data inicial não é válida."
            txt_Inicial.SetFocus
            Exit Function
        End If
    End If
    
    If txt_Ate = "" Then
        ExibeMensagem "A data final tem que ser digitada."
        txt_Ate.SetFocus
        Exit Function
    Else
        If gblnDataValida(txt_Ate.Text) = False Then
            ExibeMensagem "A data final não é válida."
            txt_Ate.SetFocus
            Exit Function
        End If
    End If
    
    If CVDate(txt_Inicial) > CVDate(txt_Ate) Then
        ExibeMensagem "A data inicial tem que ser anterior à data final."
        txt_Inicial.SetFocus
        Exit Function
    End If
    
blnDadosOK = True
End Function

Sub LimpaObjetos()
    txt_Inicial = ""
    txt_Ate = ""
    opt_Tipo(0).Value = True
End Sub

Private Function strQuerryAnalitico2() As String
Dim strSql As String
Dim dtInicial  As Date
Dim dtFinal    As Date
dtInicial = CVDate(txt_Inicial)
dtFinal = CVDate(txt_Ate)

'    strSql = ""
'    strSql = strSql & " SELECT COUNT(*) as TotalDocs , DV.intDocumentosEmitidos Inteiro, DE.strDescricao DocNome "
'    strSql = strSql & " FROM " & gstrDevolucao & " DV, "
'    strSql = strSql & gstrDocumentoEmitido & " DE "
'    strSql = strSql & " WHERE DV.intDocumentosEmitidos = DE.PKId "
'    strSql = strSql & " AND DV.dtmDevolucao BETWEEN " & gstrConvDtParaSql(dtInicial) & " AND " & gstrConvDtParaSql(dtFinal)
'    strSql = strSql & " AND DV.intContribuinte BETWEEN " & codInicial & " AND " & codFinal
'    strSql = strSql & " GROUP BY DV.intDocumentosEmitidos, DE.strDescricao "

strQuerryAnalitico2 = strSql
End Function

Private Function strQuerryAnalitico() As String
Dim strSql     As String
Dim dtInicial  As Date
Dim dtFinal    As Date
dtInicial = CVDate(txt_Inicial)
dtFinal = CVDate(txt_Ate)
    strSql = ""
    
'    strSql = strSql & " SELECT "
'    strSql = strSql & " "
'    strSql = strSql & " "
'
'    strSql = strSql & " FROM "
'    strSql = strSql & gstr & " ,"
'    strSql = strSql & gstr & " ,"
'    strSql = strSql & gstr & " ,"
'    strSql = strSql & gstr & "  "
    
'    strSql = strSql & " WHERE dtmDataLancamento BETWEEN " & gstrConvDtParaSql(dtInicial) & " AND " & gstrConvDtParaSql(dtFinal)
'    strSql = strSql & " "
'    strSql = strSql & " "
'    strSql = strSql & " "
'    strSql = strSql & " "
'
'    strSql = strSql & " GROUP BY "
'    strSql = strSql & " "
'    strSql = strSql & " "
'    strSql = strSql & " "
    
strQuerryAnalitico = strSql
End Function

Private Function strQuerrySintetico() As String
Dim strSql    As String
Dim Inicial   As Date
Dim Final     As Date
Inicial = CVDate(txt_Inicial)
Final = CVDate(txt_Ate)

strSql = ""
strSql = strSql & "SELECT CR.PKID AS CodReceita, CR.strDescricao AS Receita, PR.dblValorParcela, "
strSql = strSql & " OC.PKID AS CodOcorrencia, OC.strDescricao AS Ocorrencia, PR.intNumeroParcela "
strSql = strSql & " FROM " & gstrParcelaReceita & " PR, "
strSql = strSql & gstrComposicaoDaReceita & " CR,"
strSql = strSql & gstrOcorrencia & " OC "
strSql = strSql & " WHERE PR.intComposicaoReceita = CR.PKId "
strSql = strSql & " AND PR.intOcorrencia = OC.PKID "
strSql = strSql & " AND dtmDataLancamento BETWEEN " & gstrConvDtParaSql(Inicial) & " AND " & gstrConvDtParaSql(Final)
strSql = strSql & " ORDER BY CodReceita, CodOcorrencia"

'    select RC.strDescricao, sum(PR.dblValorParcela)
'    from tblParcelaReceita  PR,
'             tblComposicaoDaReceita RC
'
'
'    Where PR.intComposicaoReceita = RC.PKId
'
'    group by RC.strDescricao


strQuerrySintetico = strSql
End Function

Private Sub opt_Tipo_Click(Index As Integer)
    If Index = 0 Then
    
    
    End If
End Sub

Private Sub opt_Tipo_KeyPress(Index As Integer, KeyAscii As Integer)
    CaracterValido KeyAscii, "A", opt_Tipo(Index)
End Sub

Private Sub txt_Ate_GotFocus()
    MarcaCampo txt_Ate
End Sub

Private Sub txt_Ate_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "D", txt_Ate
End Sub

Private Sub txt_Inicial_GotFocus()
    MarcaCampo txt_Inicial
End Sub

Private Sub txt_Inicial_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "D", txt_Inicial
End Sub

