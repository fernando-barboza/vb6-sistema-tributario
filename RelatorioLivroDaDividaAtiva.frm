VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmRelatorioLivroDaDividaAtiva 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Livro da Dívida Ativa"
   ClientHeight    =   2280
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9690
   Icon            =   "RelatorioLivroDaDividaAtiva.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2280
   ScaleWidth      =   9690
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab tab_3dPasta 
      Height          =   1905
      Left            =   180
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   180
      Width           =   9345
      _ExtentX        =   16484
      _ExtentY        =   3360
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Livro da Dívida Ativa"
      TabPicture(0)   =   "RelatorioLivroDaDividaAtiva.frx":1042
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lbl_NumeroLivroInscricao"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lbl_Exercicio"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fra_Origem"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txtNumeroLivroInscricao"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtExercicio"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      Begin VB.TextBox txtExercicio 
         Height          =   285
         Left            =   1005
         MaxLength       =   4
         TabIndex        =   1
         Top             =   1290
         Width           =   705
      End
      Begin VB.TextBox txtNumeroLivroInscricao 
         Height          =   285
         Left            =   2895
         MaxLength       =   8
         TabIndex        =   2
         Top             =   1290
         Width           =   1155
      End
      Begin VB.Frame fra_Origem 
         Caption         =   " Origem"
         Height          =   675
         Left            =   240
         TabIndex        =   0
         Top             =   480
         Width           =   8865
         Begin VB.OptionButton optOrigem 
            Caption         =   "Imobiliário Rural"
            Height          =   195
            Index           =   0
            Left            =   210
            TabIndex        =   8
            Top             =   300
            Width           =   1575
         End
         Begin VB.OptionButton optOrigem 
            Caption         =   "Imobiliário Urbano"
            Height          =   195
            Index           =   1
            Left            =   1770
            TabIndex        =   7
            Top             =   300
            Width           =   1695
         End
         Begin VB.OptionButton optOrigem 
            Caption         =   "Econômico"
            Height          =   195
            Index           =   2
            Left            =   3510
            TabIndex        =   6
            Top             =   300
            Width           =   1155
         End
         Begin VB.OptionButton optOrigem 
            Caption         =   "Contribuição de Melhorias"
            Height          =   195
            Index           =   3
            Left            =   4860
            TabIndex        =   5
            Top             =   300
            Width           =   2265
         End
         Begin VB.OptionButton optOrigem 
            Caption         =   "Receitas Diversas"
            Height          =   195
            Index           =   4
            Left            =   7110
            TabIndex        =   4
            Top             =   300
            Width           =   1605
         End
      End
      Begin VB.Label lbl_Exercicio 
         AutoSize        =   -1  'True
         Caption         =   "Exercício"
         Height          =   195
         Left            =   240
         TabIndex        =   10
         Top             =   1380
         Width           =   675
      End
      Begin VB.Label lbl_NumeroLivroInscricao 
         AutoSize        =   -1  'True
         Caption         =   "Nº do Livro"
         Height          =   195
         Left            =   2025
         TabIndex        =   9
         Top             =   1380
         Width           =   795
      End
   End
End
Attribute VB_Name = "frmRelatorioLivroDaDividaAtiva"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrDeletar
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrSalvar
End Sub
Private Sub Form_Deactivate()
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar
End Sub

Public Sub MantemForm(ByVal strModoOperacao As String)
    
    If UCase(strModoOperacao) = UCase(gstrImprimir) Then
        If blnDadosOk Then
            ImprimeRelatorio rptRelatorioLivroDaDividaAtiva, strQueryRelatorio
        End If
    End If
    
    If UCase(strModoOperacao) = UCase(gstrNovo) Then
        LimpaObjetos
    End If
    
    If UCase(strModoOperacao) = UCase(gstrFechar) Then
        Unload Me
    End If


End Sub

Private Function blnDadosOk() As Boolean
blnDadosOk = False
On Error GoTo err_blnDadosOK
    If optOrigem(0).Value = False And optOrigem(1).Value = False _
    And optOrigem(2).Value = False And optOrigem(3).Value = False _
    And optOrigem(4).Value = False Then
        ExibeMensagem "A origem tem que ser informada."
        Exit Function
    End If
    If txtExercicio.Text = "" Then
        ExibeMensagem " O campo " & lbl_Exercicio.Caption & " não pode ser nulo."
        txtExercicio.SetFocus
        Exit Function
        End If
    If txtNumeroLivroInscricao.Text = "" Then
        ExibeMensagem " O campo " & lbl_NumeroLivroInscricao.Caption & " não pode ser nulo."
        txtNumeroLivroInscricao.SetFocus
        Exit Function
    End If
    blnDadosOk = True
err_blnDadosOK:
End Function

Private Sub LimpaObjetos()
    txtExercicio.Text = ""
    txtNumeroLivroInscricao.Text = ""
    For giContador = 0 To 4
        If optOrigem(giContador).Value = True Then
            optOrigem(giContador).Value = False
        End If
    Next
End Sub

Private Function strQueryRelatorio() As String

'******************************************************************************************
' Data: 05/05/2003
' Alteração: - Retirado comando "AS" utilizado para dar apelidos às tabelas, o qual não é
'            permitido pelo Oracle, da cláusula FROM.
' Responsável: Everton Bianchini
'******************************************************************************************

Dim strTabela   As String
Dim strCampo    As String
Dim strSql      As String
    
'    If optOrigem(0).Value = True Then 'Imobiliario rural
'        strTabela = gstrImobiliarioRural
'        strCampo = "strInscricaoAnterior"
'    ElseIf optOrigem(1).Value = True Then 'Imobiliario
'        strTabela = gstrImobiliario
'        strCampo = "strInscricaoAnterior"
'    ElseIf optOrigem(2).Value = True Then 'Economico
'        strTabela = gstrEconomico
'        strCampo = "strInscricaoCadastral"
'    ElseIf optOrigem(3).Value = True Then 'Contribuicao de melhorias
'        strTabela = gstrImobiliario
'        strCampo = "strInscricaoCadastral"
'    ElseIf optOrigem(4).Value = True Then 'Receitas diversas
'        strTabela = gstrReceitasDiversas
'        strCampo = ""
'    End If
    
    strSql = ""
    strSql = strSql & " SELECT DDA.*, CO.strNome "
    strSql = strSql & " FROM "
'    strSql = strSql & gstrDividaAtiva & " AS DA, "
    strSql = strSql & gstrDividaAtiva & " DA, "
'    strSql = strSql & gstrDetalheDividaAtiva & " AS DDA, "
    strSql = strSql & gstrDetalheDividaAtiva & " DDA, "
'    strSql = strSql & gstrContribuinte & " AS CO "
    strSql = strSql & gstrContribuinte & " CO "
    strSql = strSql & " WHERE "
    strSql = strSql & " DA.PKId = DDA.intDividaAtiva "
    strSql = strSql & " AND DA.intContribuinte = CO.PKId "
    strSql = strSql & " AND DDA.intExercicio = " & txtExercicio.Text
    strSql = strSql & " AND DDA.intNumeroLivroInscricao = " & txtNumeroLivroInscricao.Text
    For giContador = 0 To 4
        If optOrigem(giContador).Value = True Then
            strSql = strSql & " AND bytOrigem = " & giContador
            Exit For
        End If
    Next
    
strQueryRelatorio = strSql
End Function

Private Sub txtExercicio_GotFocus()
    MarcaCampo txtExercicio
End Sub

Private Sub txtExercicio_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtExercicio
End Sub

Private Sub txtNumeroLivroInscricao_GotFocus()
    MarcaCampo txtNumeroLivroInscricao
End Sub

Private Sub txtNumeroLivroInscricao_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtNumeroLivroInscricao
End Sub


