VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmProgressBar 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   975
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   8310
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   975
   ScaleWidth      =   8310
   Begin VB.Frame fraProgressao 
      Caption         =   " Aguarde, carregando dados... "
      ClipControls    =   0   'False
      Height          =   735
      Left            =   135
      TabIndex        =   0
      Top             =   105
      Width           =   8055
      Begin VB.CommandButton cmd_Cancelar 
         Cancel          =   -1  'True
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   6825
         TabIndex        =   4
         Top             =   240
         Width           =   1110
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   345
         Left            =   660
         TabIndex        =   3
         Top             =   255
         Width           =   5565
         _Version        =   65536
         _ExtentX        =   9816
         _ExtentY        =   609
         _StockProps     =   15
         BackColor       =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         BevelInner      =   1
         FloodType       =   1
      End
      Begin VB.Label lblPorCento 
         AutoSize        =   -1  'True
         Caption         =   "0%"
         Height          =   195
         Index           =   0
         Left            =   390
         TabIndex        =   2
         Top             =   345
         Width           =   210
      End
      Begin VB.Label lblPorCento 
         AutoSize        =   -1  'True
         Caption         =   "100%"
         Height          =   195
         Index           =   1
         Left            =   6255
         TabIndex        =   1
         Top             =   330
         Width           =   390
      End
   End
End
Attribute VB_Name = "frmProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_Cancelar_Click()
    gblnCancelar = True
End Sub

Private Sub cmd_Cancelar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Screen.MousePointer = 0
End Sub

Private Sub Form_Load()
    With MDIMenu
      Me.Left = Int((.Width - Me.Width) / 2)
      Me.Top = Int((.Height - Me.Height) / 2) - 1000
      If Me.Top < 0 Then
         Me.Top = 0
      End If
    End With
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Screen.MousePointer = 11
End Sub

