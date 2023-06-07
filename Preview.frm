VERSION 5.00
Begin VB.Form frmPreview 
   Caption         =   "Visualização"
   ClientHeight    =   4440
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5145
   Icon            =   "Preview.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4440
   ScaleWidth      =   5145
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picParent 
      Height          =   2835
      Left            =   150
      ScaleHeight     =   2775
      ScaleWidth      =   4035
      TabIndex        =   1
      Top             =   750
      Width           =   4095
      Begin VB.PictureBox picPreview 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   2055
         Index           =   0
         Left            =   60
         ScaleHeight     =   2055
         ScaleWidth      =   3075
         TabIndex        =   0
         Top             =   0
         Visible         =   0   'False
         Width           =   3075
      End
      Begin VB.VScrollBar vscPreview 
         Height          =   2475
         LargeChange     =   2000
         Left            =   3720
         SmallChange     =   500
         TabIndex        =   4
         Top             =   120
         Width           =   195
      End
      Begin VB.HScrollBar hscPreview 
         Height          =   195
         LargeChange     =   2000
         Left            =   300
         SmallChange     =   500
         TabIndex        =   3
         Top             =   2280
         Width           =   2895
      End
      Begin VB.PictureBox imgCorner 
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         Height          =   240
         Left            =   3360
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   2
         Top             =   2340
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image picChild 
         Height          =   1875
         Left            =   720
         Stretch         =   -1  'True
         Top             =   360
         Width           =   2655
      End
   End
End
Attribute VB_Name = "frmPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
