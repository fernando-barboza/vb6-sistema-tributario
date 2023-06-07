VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmImprimirGuiaPositiva 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Certidão Positiva"
   ClientHeight    =   1365
   ClientLeft      =   3810
   ClientTop       =   3270
   ClientWidth     =   3660
   Icon            =   "frmImprimirGuiaPositiva.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1365
   ScaleWidth      =   3660
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab tab_3dPasta 
      Height          =   1275
      Left            =   60
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   60
      Width           =   3555
      _ExtentX        =   6271
      _ExtentY        =   2249
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Parâmetros"
      TabPicture(0)   =   "frmImprimirGuiaPositiva.frx":1042
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "cmd_Imprimir"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fra_Processo"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      Begin VB.Frame fra_Processo 
         Caption         =   " Processo "
         Height          =   615
         Left            =   210
         TabIndex        =   5
         Top             =   390
         Width           =   2085
         Begin VB.TextBox txtstrCodigo 
            CausesValidation=   0   'False
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "#.##0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
            EndProperty
            Height          =   315
            HideSelection   =   0   'False
            Left            =   90
            MaxLength       =   15
            MultiLine       =   -1  'True
            TabIndex        =   0
            Top             =   210
            Width           =   885
         End
         Begin VB.TextBox txtintExercicio 
            CausesValidation=   0   'False
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "#.##0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
            EndProperty
            Height          =   315
            Left            =   1020
            MaxLength       =   4
            TabIndex        =   1
            Top             =   210
            Width           =   525
         End
         Begin VB.TextBox txtbitDigito 
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "#.##0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
            EndProperty
            Height          =   315
            Left            =   1620
            MaxLength       =   2
            TabIndex        =   2
            Top             =   210
            Width           =   345
         End
      End
      Begin VB.CommandButton cmd_Imprimir 
         Caption         =   "Imprimir"
         Height          =   375
         Left            =   2370
         TabIndex        =   3
         Top             =   570
         Width           =   1035
      End
   End
End
Attribute VB_Name = "frmImprimirGuiaPositiva"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public bytTipoPositiva As Byte

Private Sub cmd_Imprimir_Click()
    
    If Trim(txtstrCodigo) <> "" Or Trim(txtintExercicio) <> "" Or Trim(txtbitDigito) <> "" Then
        If Trim(txtstrCodigo) = "" Then
            ExibeMensagem "O campo Código deve ser preenchido corretamente."
            txtstrCodigo.SetFocus
            Exit Sub
        ElseIf Trim(txtintExercicio) = "" Then
            ExibeMensagem "O campo Exercício deve ser preenchido corretamente."
            txtintExercicio.SetFocus
            Exit Sub
        ElseIf Trim(txtbitDigito) = "" Then
            ExibeMensagem "O campo Dígito deve ser preenchido corretamente."
            txtbitDigito.SetFocus
            Exit Sub
        End If
    
        If Not VerificaEmpenhoProcesso(Trim(txtstrCodigo), Val(txtbitDigito), Val(txtintExercicio)) Then
            ExibeMensagem "Processo não localizado."
            txtstrCodigo.SetFocus
        End If
    End If
    
    frmAtualizacaoDebitos.strNumeroProcesso = txtstrCodigo & " / " & txtintExercicio & " - " & txtbitDigito
    If bytTipoPositiva = 0 Then
        frmAtualizacaoDebitos.ImprimiPositiva
    Else
        frmAtualizacaoDebitos.ImprimiPositivaNegativo
    End If
    Unload Me
    
End Sub

Private Sub txtbitDigito_GotFocus()
    MarcaCampo txtbitDigito
End Sub

Private Sub txtbitDigito_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtbitDigito
End Sub

Private Sub txtintExercicio_GotFocus()
    MarcaCampo txtintExercicio
End Sub

Private Sub txtintExercicio_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtintExercicio
End Sub

Private Sub txtstrCodigo_GotFocus()
    MarcaCampo txtstrCodigo
End Sub

Private Sub txtstrCodigo_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrCodigo
End Sub

