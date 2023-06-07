VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmRelatorioDeDebitosNaoInscritosDV 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Débitos Não Inscritos em Dívida Ativa"
   ClientHeight    =   1635
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3780
   Icon            =   "RelatorioDeDebitosNaoInscritosDV.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1635
   ScaleWidth      =   3780
   Begin TabDlg.SSTab SSTab1 
      Height          =   1395
      Left            =   180
      TabIndex        =   0
      Top             =   150
      Width           =   3435
      _ExtentX        =   6059
      _ExtentY        =   2461
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Débitos Não Inscritos"
      TabPicture(0)   =   "RelatorioDeDebitosNaoInscritosDV.frx":1042
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin VB.Frame Frame1 
         Height          =   735
         Left            =   270
         TabIndex        =   1
         Top             =   450
         Width           =   2895
         Begin VB.TextBox txt_Exercicio 
            Height          =   285
            Left            =   1500
            MaxLength       =   4
            TabIndex        =   3
            Top             =   270
            Width           =   555
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Exercício"
            Height          =   195
            Left            =   660
            TabIndex        =   2
            Top             =   360
            Width           =   675
         End
      End
   End
End
Attribute VB_Name = "frmRelatorioDeDebitosNaoInscritosDV"
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
    txt_Exercicio = gintExercicio
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
        If txt_Exercicio.Text = "" Then
            ExibeMensagem "O exercício tem que ser digitado."
            txt_Exercicio.SetFocus
            Exit Sub
        End If
        ImprimeRelatorio rptRelatorioDeDebitosNaoInscritosDV, strQuerryRelatorio
    End If
    
    If UCase(strModoOperacao) = UCase(gstrNovo) Then
        LimpaObjetos
    End If
    
    If UCase(strModoOperacao) = UCase(gstrFechar) Then
        Unload Me
    End If
    
End Sub

Private Sub LimpaObjetos()
    txt_Exercicio.Text = ""
    txt_Exercicio.SetFocus
End Sub

Private Function strQuerryRelatorio() As String
Dim strSql As String
    strSql = ""
    strSql = strSql & " SELECT "
    strSql = strSql & " C.strNome, D.strDescricao, A.dtmDataVencimento, A.dblValorParcela "

    strSql = strSql & " FROM "
    strSql = strSql & gstrParcelaReceita & " A, "
    strSql = strSql & gstrLancamentoCalculo & " B, "
    strSql = strSql & gstrContribuinte & " C, "
    strSql = strSql & gstrComposicaoDaReceita & " D "

    strSql = strSql & " WHERE "
    strSql = strSql & " A.intLancamentoCalculo = B.PKId "
    strSql = strSql & " AND B.intContribuinte = C.PKId "
    strSql = strSql & " AND A.intComposicaoDaReceita = D.PKId "
    strSql = strSql & " AND A.bytAtiva = 0 "
    strSql = strSql & " AND B.intExercicio = " & Val(txt_Exercicio.Text)
    
    strSql = strSql & " ORDER BY "
    strSql = strSql & " C.strNome "
    
strQuerryRelatorio = strSql
End Function


