VERSION 5.00
Begin VB.Form frmPageSetup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configuração da Página"
   ClientHeight    =   1185
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3270
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1185
   ScaleWidth      =   3270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Caption         =   "Margens"
      Height          =   1155
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   1935
      Begin VB.TextBox txtLeft 
         Height          =   285
         Left            =   900
         TabIndex        =   4
         Top             =   315
         Width           =   855
      End
      Begin VB.TextBox txtRight 
         Height          =   285
         Left            =   900
         TabIndex        =   3
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "&Esquerda"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   675
      End
      Begin VB.Label Label2 
         Caption         =   "&Direita"
         Height          =   315
         Left            =   300
         TabIndex        =   5
         Top             =   780
         Width           =   495
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      Top             =   75
      Width           =   1200
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   2040
      TabIndex        =   0
      Top             =   525
      Width           =   1200
   End
End
Attribute VB_Name = "frmPageSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Portions of the code are copyright 1995, Microsoft Corporation
Option Explicit

'*************************************************
' Purpose:  Unload the form
'*************************************************
Private Sub cmdCancel_Click()
    Unload Me
End Sub

'*************************************************
' Purpose:  Update the formatting of the text in
'           the active RTF control.
'*************************************************
Private Sub cmdOK_Click()
    ' validate margin values in text boxes
    If Val(txtLeft.Text) < 0 Then
        Beep 'beep..beep..beep
        ' show error
        MsgBox "Left margin must be no less than zero inches.", 16, "Margin Out Of Range"
        'select box with error
        txtLeft.SelStart = 0
        txtLeft.SelLength = Len(txtLeft.Text)
        'set focus to box
        txtLeft.SetFocus
    ElseIf Val(txtRight.Text) < 0 Then
        Beep
        ' show error message
        MsgBox "Right margin must be no less than zero inches.", 16, "Margin Out Of Range"
        'select text with  error
        txtRight.SelStart = 0
        txtRight.SelLength = Len(txtRight.Text)
        'set focus to text
        txtRight.SetFocus
    Else
        'passed checks
        ' variables to store old selection
        Dim lngOldStart As Long
        Dim lngOldLength As Long
        With MDIMenu.ActiveForm.rtf
            ' save old text selection
            lngOldStart = .SelStart
            lngOldLength = .SelLength
            ' select entire document
            .SelStart = 0
            .SelLength = Len(.Text)
            ' set new margins
            ' the value needs to be converted to twips
            ' for acuracy.  There are 1440 Twips/Inch.
            .SelIndent = CInt(Val(txtLeft.Text) * 1440)
            .SelRightIndent = CInt(Val(txtRight.Text) * 1440)
            ' restore old selection
            .SelStart = lngOldStart
            .SelLength = lngOldLength
        End With
        'unload form
        Unload Me
    End If
End Sub

'*************************************************
' Purpose:  Initialize form with margin values
'*************************************************
Private Sub Form_Load()
    With MDIMenu.ActiveForm.rtf
        ' variables to store new values
        Dim sglLeft As Single
        Dim sglRight As Single
        ' here the variables get set.
        ' I go through some extra steps to make sure
        ' I do not lose precision when converting
        ' integers to floatig-point values.
        sglLeft = .SelIndent
        sglLeft = sglLeft / 1440#
        sglLeft = CInt(sglLeft * 100#)
        sglLeft = sglLeft / 100#
        txtLeft.Text = Trim(Str(sglLeft))
        sglRight = .SelRightIndent
        sglRight = sglRight / 1440#
        sglRight = CInt(sglRight * 100#)
        sglRight = sglRight / 100#
        txtRight.Text = Trim(Str(sglRight))
    End With
End Sub

Private Sub txtLeft_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "V", txtLeft
End Sub

Private Sub txtRight_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "V", txtRight
End Sub
