VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmRelatorioDeQuantidadeDeLancamentosValorTipo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Quantidade de Lançamento, Valor e Tipo"
   ClientHeight    =   1725
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5865
   Icon            =   "RelatorioDeQuantidadeDeLancamentosValorTipo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1725
   ScaleWidth      =   5865
   Begin TabDlg.SSTab tab_3dPasta 
      Height          =   1455
      Left            =   180
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   150
      Width           =   5505
      _ExtentX        =   9710
      _ExtentY        =   2566
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Quantidade de Lançamento, Valor e Tipo"
      TabPicture(0)   =   "RelatorioDeQuantidadeDeLancamentosValorTipo.frx":1042
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin VB.Frame Frame2 
         Caption         =   "Data para pesquisa"
         Height          =   765
         Left            =   210
         TabIndex        =   1
         Top             =   450
         Width           =   5055
         Begin VB.TextBox txt_DataFinal 
            Height          =   285
            Left            =   3780
            MaxLength       =   11
            TabIndex        =   3
            Top             =   300
            Width           =   1065
         End
         Begin VB.TextBox txt_DataInicial 
            Height          =   285
            Left            =   900
            MaxLength       =   11
            TabIndex        =   2
            Top             =   300
            Width           =   1065
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Final"
            Height          =   195
            Left            =   3210
            TabIndex        =   5
            Top             =   390
            Width           =   330
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Inicial"
            Height          =   195
            Left            =   270
            TabIndex        =   4
            Top             =   390
            Width           =   405
         End
      End
   End
End
Attribute VB_Name = "frmRelatorioDeQuantidadeDeLancamentosValorTipo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim mblnAlterando    As Boolean
    Dim mobjAux          As Object
    Dim mblnSelecionou   As Boolean
    Dim mblnPrimeiraVez  As Boolean
    Dim intCodigoInicial As Integer
    Dim intCodigoFinal   As Integer
    Dim TipoDeInscricao  As Integer

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
Dim adoResultado   As ADODB.Recordset

    If UCase(strModoOperacao) = UCase(gstrImprimir) Then
        If blnDadosOK Then
            ImprimeRelatorio rptRelatorioDeQuantidadeDeLancamentoValorTipo, strQuerryRelatorio
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
    txt_DataInicial.Text = ""
    txt_DataFinal.Text = ""
    txt_DataInicial.SetFocus
End Sub

Private Function blnDadosOK() As Boolean
blnDadosOK = False
On Error GoTo err_blnDadosOK
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
blnDadosOK = True
err_blnDadosOK:
End Function

Private Function strQuerryRelatorio() As String
Dim strSql As String
    
    strSql = ""
    strSql = strSql & " SELECT "
    strSql = strSql & " A.intOcorrencia, C.strDescricao AS OCORRENCIA, "
    strSql = strSql & " COUNT(distinct A.PKId) AS QTDLancamento, "
    strSql = strSql & " SUM(b.dblValorParcela) As SomaPorLancamento "
                            
    strSql = strSql & " FROM "
    strSql = strSql & gstrLancamentoCalculo & " A, "
    strSql = strSql & gstrParcelaReceita & " B, "
    strSql = strSql & gstrOcorrencia & " C "

    strSql = strSql & " WHERE "
    strSql = strSql & " B.intLancamentoCalculo = A.PKId "
    strSql = strSql & " AND A.intOcorrencia = C.PKId "
    strSql = strSql & " AND C.intUtilizacaoDaOcorrencia = 1 "
    strSql = strSql & " AND A.dtmLancamento BETWEEN " & gstrConvDtParaSql(txt_DataInicial.Text) & " AND " & gstrConvDtParaSql(txt_DataFinal.Text)
    
    strSql = strSql & " GROUP BY "
    strSql = strSql & " A.intOcorrencia , c.strDescricao "
    
    strSql = strSql & " ORDER BY "
    strSql = strSql & " C.strDescricao "
    
strQuerryRelatorio = strSql
End Function

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

