VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmRelSaldoDividaAtiva 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Saldo Dívida Ativa"
   ClientHeight    =   1140
   ClientLeft      =   2910
   ClientTop       =   6210
   ClientWidth     =   9210
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1140
   ScaleWidth      =   9210
   Begin VB.Frame fra_Relatorio 
      Height          =   900
      Left            =   7560
      TabIndex        =   4
      Top             =   75
      Width           =   1530
      Begin VB.OptionButton opt_Detalhado 
         Caption         =   "Detalhado"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   525
         Width           =   1245
      End
      Begin VB.OptionButton opt_Resumido 
         Caption         =   "Resumido"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   225
         Value           =   -1  'True
         Width           =   1185
      End
      Begin MSDataListLib.DataCombo dbc_intExercicioFinal 
         Height          =   315
         Left            =   3480
         TabIndex        =   5
         Top             =   255
         Visible         =   0   'False
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin VB.Label lbl_ExercicioFinal 
         AutoSize        =   -1  'True
         Caption         =   "Exercício Final"
         Height          =   195
         Left            =   2370
         TabIndex        =   6
         Top             =   330
         Visible         =   0   'False
         Width           =   1050
      End
   End
   Begin VB.Frame fra_Dados 
      Height          =   900
      Left            =   90
      TabIndex        =   0
      Top             =   75
      Width           =   7275
      Begin VB.CommandButton cmd_Composicao 
         Height          =   300
         Left            =   4965
         Picture         =   "frmRelSaldoDividaAtiva.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   1
         TabStop         =   0   'False
         Tag             =   "590"
         ToolTipText     =   "AtivaCadastro de Composição da Receita"
         Top             =   360
         Width           =   360
      End
      Begin MSDataListLib.DataCombo dbc_intComposicao 
         Height          =   315
         Left            =   1140
         TabIndex        =   2
         Top             =   360
         Width           =   3795
         _ExtentX        =   6694
         _ExtentY        =   556
         _Version        =   393216
         IntegralHeight  =   0   'False
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dbc_intExercicioInicial 
         Height          =   315
         Left            =   6195
         TabIndex        =   7
         Top             =   360
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin VB.Label lbl_ExercicioInicial 
         AutoSize        =   -1  'True
         Caption         =   "Exercício"
         Height          =   195
         Left            =   5460
         TabIndex        =   8
         Top             =   405
         Width           =   675
      End
      Begin VB.Label lbl_Composicao 
         AutoSize        =   -1  'True
         Caption         =   "Composição"
         Height          =   195
         Left            =   165
         TabIndex        =   3
         Top             =   405
         Width           =   870
      End
   End
End
Attribute VB_Name = "frmRelSaldoDividaAtiva"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

