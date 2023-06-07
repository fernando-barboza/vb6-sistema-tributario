VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmIntervaloDataArrecadacao 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Arrecadação da Receita - Intervalo de Datas"
   ClientHeight    =   1395
   ClientLeft      =   2025
   ClientTop       =   1935
   ClientWidth     =   5040
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1395
   ScaleWidth      =   5040
   Begin TabDlg.SSTab tab_DatasArrecadacao 
      Height          =   1050
      Left            =   270
      TabIndex        =   0
      Top             =   150
      Width           =   4425
      _ExtentX        =   7805
      _ExtentY        =   1852
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Intervalo de Datas"
      TabPicture(0)   =   "frmIntervaloDataArrecadacao.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblDataInicial"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblDataFinal"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "txtDataInicial"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txtDataFinal"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      Begin VB.TextBox txtDataFinal 
         Height          =   285
         Left            =   3105
         TabIndex        =   4
         Top             =   600
         Width           =   960
      End
      Begin VB.TextBox txtDataInicial 
         Height          =   285
         Left            =   1095
         TabIndex        =   2
         Top             =   600
         Width           =   960
      End
      Begin VB.Label lblDataFinal 
         AutoSize        =   -1  'True
         Caption         =   "Data Final"
         Height          =   195
         Left            =   2325
         TabIndex        =   3
         Top             =   645
         Width           =   720
      End
      Begin VB.Label lblDataInicial 
         AutoSize        =   -1  'True
         Caption         =   "Data Inicial"
         Height          =   195
         Left            =   225
         TabIndex        =   1
         Top             =   630
         Width           =   795
      End
   End
End
Attribute VB_Name = "frmIntervaloDataArrecadacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
   frmArrecadacaoReceita.blnAtivaFormImprime = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
   frmArrecadacaoReceita.blnAtivaFormImprime = False
End Sub

Private Sub txtDataInicial_GotFocus()
    MarcaCampo txtDataInicial
End Sub

Private Sub txtDataInicial_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "D", txtDataInicial
End Sub

Private Sub txtDataInicial_LostFocus()
    txtDataInicial = gstrDataFormatada(txtDataInicial)
End Sub
Private Sub txtDataFinal_GotFocus()
    MarcaCampo txtDataFinal
End Sub

Private Sub txtDataFinal_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "D", txtDataFinal
End Sub

Private Sub txtDataFinal_LostFocus()
    txtDataFinal = gstrDataFormatada(txtDataFinal)
End Sub
Public Sub MantemForm(ByVal strModoOperacao As String)
   If UCase(strModoOperacao) = UCase(gstrImprimir) Then
      If blnDadosOK Then
         frmArrecadacaoReceita.strDataInicial = txtDataInicial
         frmArrecadacaoReceita.strDataFinal = txtDataFinal
         frmArrecadacaoReceita.MantemForm gstrImprimir
         Unload Me
      Else
         frmArrecadacaoReceita.blnAtivaFormImprime = True
      End If
   End If
End Sub

Private Function blnDadosOK()
   If gblnDataValida(txtDataInicial) = False Then
        ExibeMensagem "A data inicial tem que ser informada corretamente."
        txtDataInicial.SetFocus
        Exit Function
   End If
   If gblnDataValida(txtDataFinal) = False Then
      ExibeMensagem "A data final tem que ser informada corretamente."
      txtDataFinal.SetFocus
      Exit Function
   End If
   If (CVDate(txtDataInicial) > CVDate(txtDataFinal)) Then
      ExibeMensagem "A data inicial não pode ser maior que a data final."
      txtDataInicial.SetFocus
      Exit Function
   End If
   blnDadosOK = True
End Function
