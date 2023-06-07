VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmDicas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Dicas do Dia"
   ClientHeight    =   3180
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5265
   HelpContextID   =   100
   Icon            =   "Dicas.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Dicas.frx":1042
   ScaleHeight     =   3180
   ScaleWidth      =   5265
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdDicaAnterior 
      Caption         =   "&Dica Anterior"
      Height          =   375
      Left            =   3990
      TabIndex        =   6
      Top             =   945
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3990
      TabIndex        =   4
      Top             =   15
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C0FFFF&
      Height          =   2715
      Left            =   15
      Picture         =   "Dicas.frx":190C
      ScaleHeight     =   2655
      ScaleWidth      =   3705
      TabIndex        =   0
      Top             =   15
      Width           =   3765
      Begin Threed.SSPanel panDicas 
         Height          =   495
         Left            =   465
         TabIndex        =   1
         Top             =   15
         Width           =   3225
         _Version        =   65536
         _ExtentX        =   5689
         _ExtentY        =   873
         _StockProps     =   15
         Caption         =   "Dicas do Dia"
         ForeColor       =   0
         BackColor       =   12648447
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
         BevelInner      =   1
      End
      Begin VB.Label lblTextoDicas 
         BackColor       =   &H00C0FFFF&
         Height          =   1860
         Left            =   135
         TabIndex        =   2
         Top             =   735
         Width           =   3435
      End
   End
   Begin VB.CommandButton cmdProximaDica 
      Caption         =   "&Próxima Dica"
      Height          =   375
      Left            =   3990
      TabIndex        =   5
      Top             =   480
      Width           =   1215
   End
   Begin VB.CheckBox chkMostraDicaInicio 
      Caption         =   "&Mostra dicas ao iniciar"
      Height          =   315
      Left            =   15
      TabIndex        =   3
      Top             =   2835
      Width           =   2055
   End
End
Attribute VB_Name = "frmDicas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim strRecebeTexto As String
    Dim vntAux         As Variant
    Dim intContDicas   As Integer
    Dim strSql         As String

Private Sub chkMostraDicaInicio_Click()
    gblnMostraDicas = chkMostraDicaInicio.Value
End Sub

Private Sub cmdDicaAnterior_Click()
    'Localiza a proxima dica a ser apresentada
    intContDicas = intContDicas - 1
    If cmdProximaDica.Enabled = False Then
        cmdProximaDica.Enabled = True
    End If
    If intContDicas < 2 Then
        intContDicas = 1
        cmdDicaAnterior.Enabled = False
    End If
    strRecebeTexto = Space(255)
    lblTextoDicas.Caption = ""
    vntAux = GetPrivateProfileString("Dicas", CStr(intContDicas), "", strRecebeTexto, Len(strRecebeTexto), App.Path & "\Dicas.Txt")
    lblTextoDicas.Caption = lblTextoDicas.Caption & gstrDiaDaSemana(CStr(Date)) & _
                               ", " & Day(Date) & " de " & _
                               gstrNomeDoMes(Month(Date)) & " de " & _
                               Year(Date) & vbCr & vbCr
    lblTextoDicas.Caption = lblTextoDicas.Caption & "Você sabia que ..." & vbCr & vbCr
    lblTextoDicas.Caption = lblTextoDicas.Caption & strRecebeTexto
End Sub

Private Sub cmdOK_Click()
    strSql = ""
    strSql = strSql & "UPDATE " & gstrUsuarios & " SET blnMostraDicas = " & _
                      Abs(gblnMostraDicas) & " WHERE PKId = " & glngCodUsr
    Set gobjBanco = New clsBanco
    gcncADOMain.Execute strSql
    Unload Me
End Sub

Private Sub cmdProximaDica_Click()
    'Localiza a proxima dica a ser apresentada
    intContDicas = intContDicas + 1
    If cmdDicaAnterior.Enabled = False Then
        cmdDicaAnterior.Enabled = True
    End If
    If intContDicas > 30 Then
        intContDicas = 31
        cmdProximaDica.Enabled = False
    End If
    strRecebeTexto = Space(255)
    lblTextoDicas.Caption = ""
    vntAux = GetPrivateProfileString("Dicas", CStr(intContDicas), "", strRecebeTexto, Len(strRecebeTexto), App.Path & "\Dicas.Txt")
    lblTextoDicas.Caption = lblTextoDicas.Caption & gstrDiaDaSemana(CStr(Date)) & _
                               ", " & Day(Date) & " de " & _
                               gstrNomeDoMes(Month(Date)) & " de " & _
                               Year(Date) & vbCr & vbCr
    lblTextoDicas.Caption = lblTextoDicas.Caption & "Você sabia que ..." & vbCr & vbCr
    lblTextoDicas.Caption = lblTextoDicas.Caption & strRecebeTexto
End Sub

Private Sub Form_Load()
    chkMostraDicaInicio.Value = Abs(gblnMostraDicas)
    'Seta a variavel para o valor do dia da semana
    intContDicas = Day(Date)
    strRecebeTexto = Space(255)
    'Retorna o dia a semana e a dica referente a este dia o arquivo ADICAS.TXT esta localizado
    'na pasta Windows
    vntAux = GetPrivateProfileString("Dicas", CStr(intContDicas), "", strRecebeTexto, Len(strRecebeTexto), App.Path & "\Dicas.Txt")
    lblTextoDicas.Caption = lblTextoDicas.Caption & gstrDiaDaSemana(CStr(Date)) & _
                               ", " & Day(Date) & " de " & _
                               gstrNomeDoMes(Month(Date)) & " de " & _
                               Year(Date) & vbCr & vbCr
    lblTextoDicas.Caption = lblTextoDicas.Caption & " Você sabia que ... " & vbCr & vbCr
    lblTextoDicas.Caption = lblTextoDicas.Caption & strRecebeTexto
End Sub
