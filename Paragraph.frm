VERSION 5.00
Begin VB.Form frmParagraph 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Paragrafo"
   ClientHeight    =   1050
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3315
   Icon            =   "Paragraph.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1050
   ScaleWidth      =   3315
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WhatsThisHelp   =   -1  'True
   Begin VB.Frame Frame1 
      Caption         =   "Alinhamento"
      Height          =   1035
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   1935
      Begin VB.PictureBox picAlign 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   540
         Left            =   1140
         ScaleHeight     =   510
         ScaleWidth      =   570
         TabIndex        =   6
         Top             =   360
         Width           =   600
      End
      Begin VB.OptionButton opnAlign 
         Caption         =   "Esquerda"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton opnAlign 
         Caption         =   "Centro"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Width           =   1455
      End
      Begin VB.OptionButton opnAlign 
         Caption         =   "Direita"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Width           =   1275
      End
   End
   Begin VB.PictureBox picAlignArr 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   0
      Left            =   120
      ScaleHeight     =   585
      ScaleWidth      =   645
      TabIndex        =   4
      Top             =   1500
      Width           =   675
   End
   Begin VB.PictureBox picAlignArr 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   1
      Left            =   840
      ScaleHeight     =   585
      ScaleWidth      =   645
      TabIndex        =   3
      Top             =   1500
      Width           =   675
   End
   Begin VB.PictureBox picAlignArr 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   2
      Left            =   1560
      ScaleHeight     =   585
      ScaleWidth      =   645
      TabIndex        =   2
      Top             =   1500
      Width           =   675
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   2100
      TabIndex        =   1
      Top             =   75
      Width           =   1155
   End
   Begin VB.CommandButton cmdCxl 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   390
      Left            =   2100
      TabIndex        =   0
      Top             =   525
      Width           =   1155
   End
End
Attribute VB_Name = "frmParagraph"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Portions of the code are copyright 1995, Microsoft Corporation
Option Explicit

'*************************************************
' Purpose:  Unloads form
'*************************************************
Private Sub cmdCxl_Click()
    Unload Me
End Sub


'*************************************************
' Purpose:  Sets the alignment of the selected
'           text in the active RTF control.
'*************************************************
Private Sub cmdOK_Click()
    Dim intSelected As Integer ' index of selected item
    With MDIMenu.ActiveForm.rtf
    If opnAlign(0).Value Then
        .SelAlignment = 0
    ElseIf opnAlign(1).Value Then
        .SelAlignment = 2
    ElseIf opnAlign(2).Value Then
        .SelAlignment = 1
    End If
    End With
    Unload Me
End Sub

'*************************************************
' Purpose:  Initialize form with values from selected
'           text in the active RTF control.
'*************************************************
Private Sub Form_Load()
    With MDIMenu.ActiveForm.rtf
        Select Case .SelAlignment
            Case Null, rtfLeft
                opnAlign(0).Value = True
            Case rtfCenter
                opnAlign(1).Value = True
            Case rtfRight
                opnAlign(2).Value = True
        End Select
    End With
End Sub

'*************************************************
' Purpose:  Set the picture of the selected option.
'*************************************************
Private Sub opnAlign_Click(intIndex As Integer)
    picAlign.Picture = picAlignArr(intIndex).Picture
End Sub
