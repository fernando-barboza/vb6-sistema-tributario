VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCadImagens 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cálculo das Áreas das Principais Figuras Planas"
   ClientHeight    =   6660
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11220
   HelpContextID   =   5
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6660
   ScaleWidth      =   11220
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab tab_3DPasta 
      Height          =   6555
      Left            =   60
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   30
      Width           =   11115
      _ExtentX        =   19606
      _ExtentY        =   11562
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Figuras Planas"
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "shpBorda(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "shpBorda(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "shpBorda(2)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "shpBorda(3)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "shpBorda(4)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "shpBorda(5)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "shpBorda(6)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "shpBorda(7)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "shpBorda(8)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "shpBorda(9)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "shpBorda(10)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "optFigura(3)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "optFigura(4)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "optFigura(0)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "optFigura(1)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "optFigura(2)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "optFigura(5)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "optFigura(6)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "optFigura(7)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "optFigura(8)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "optFigura(10)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "optFigura(9)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "fraCalculo"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "fraformulas"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Frame2"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "cmdCancelar"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "cmdAplicar"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).ControlCount=   27
      Begin VB.CommandButton cmdAplicar 
         Caption         =   "&Aplicar"
         Enabled         =   0   'False
         Height          =   405
         Left            =   9780
         TabIndex        =   63
         Top             =   6000
         Width           =   1155
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   405
         Left            =   8535
         TabIndex        =   62
         Top             =   6000
         Width           =   1155
      End
      Begin VB.Frame Frame2 
         Caption         =   " Unidade de Medida "
         ForeColor       =   &H8000000D&
         Height          =   600
         Left            =   315
         TabIndex        =   26
         Top             =   5790
         Width           =   2085
         Begin VB.OptionButton optKm 
            Caption         =   "Km"
            Height          =   195
            Left            =   1185
            TabIndex        =   28
            Top             =   270
            Width           =   780
         End
         Begin VB.OptionButton optMetros 
            Caption         =   "Metros"
            Height          =   195
            Left            =   225
            TabIndex        =   27
            Top             =   270
            Value           =   -1  'True
            Width           =   1500
         End
      End
      Begin VB.Frame fraformulas 
         Caption         =   " Fórmula "
         ForeColor       =   &H8000000D&
         Height          =   1830
         Left            =   2460
         TabIndex        =   25
         Top             =   4050
         Width           =   1650
         Begin VB.Frame fraFormulaVazio 
            Height          =   1275
            Left            =   135
            TabIndex        =   60
            Top             =   435
            Width           =   1410
         End
         Begin VB.Frame fraFormulasFiguras 
            Height          =   1275
            Index           =   10
            Left            =   135
            TabIndex        =   57
            Top             =   435
            Width           =   1410
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "2"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   7
               Left            =   795
               TabIndex        =   59
               Top             =   795
               Width           =   120
            End
            Begin VB.Line Line1 
               Index           =   7
               X1              =   585
               X2              =   1110
               Y1              =   735
               Y2              =   735
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "S = b x h"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   10
               Left            =   285
               TabIndex        =   58
               Top             =   495
               Width           =   780
            End
         End
         Begin VB.Frame fraFormulasFiguras 
            Height          =   1275
            Index           =   9
            Left            =   135
            TabIndex        =   54
            Top             =   435
            Width           =   1410
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "2"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   6
               Left            =   795
               TabIndex        =   56
               Top             =   795
               Width           =   120
            End
            Begin VB.Line Line1 
               Index           =   6
               X1              =   585
               X2              =   1110
               Y1              =   735
               Y2              =   735
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "S = b x h"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   9
               Left            =   285
               TabIndex        =   55
               Top             =   495
               Width           =   780
            End
         End
         Begin VB.Frame fraFormulasFiguras 
            Height          =   1275
            Index           =   8
            Left            =   135
            TabIndex        =   51
            Top             =   435
            Width           =   1410
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "2"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   5
               Left            =   795
               TabIndex        =   53
               Top             =   795
               Width           =   120
            End
            Begin VB.Line Line1 
               Index           =   5
               X1              =   585
               X2              =   1110
               Y1              =   735
               Y2              =   735
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "S = b x h"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   8
               Left            =   285
               TabIndex        =   52
               Top             =   495
               Width           =   780
            End
         End
         Begin VB.Frame fraFormulasFiguras 
            Height          =   1275
            Index           =   7
            Left            =   135
            TabIndex        =   48
            Top             =   435
            Width           =   1410
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "2"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   4
               Left            =   795
               TabIndex        =   50
               Top             =   795
               Width           =   120
            End
            Begin VB.Line Line1 
               Index           =   4
               X1              =   585
               X2              =   1110
               Y1              =   735
               Y2              =   735
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "S = b x h"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   7
               Left            =   285
               TabIndex        =   49
               Top             =   495
               Width           =   780
            End
         End
         Begin VB.Frame fraFormulasFiguras 
            Height          =   1275
            Index           =   6
            Left            =   135
            TabIndex        =   45
            Top             =   435
            Width           =   1410
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "2"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   3
               Left            =   795
               TabIndex        =   47
               Top             =   795
               Width           =   120
            End
            Begin VB.Line Line1 
               Index           =   3
               X1              =   585
               X2              =   1110
               Y1              =   735
               Y2              =   735
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "S = b x h"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   6
               Left            =   285
               TabIndex        =   46
               Top             =   495
               Width           =   780
            End
         End
         Begin VB.Frame fraFormulasFiguras 
            Height          =   1275
            Index           =   5
            Left            =   135
            TabIndex        =   42
            Top             =   435
            Width           =   1410
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "S = b x h"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   5
               Left            =   285
               TabIndex        =   44
               Top             =   495
               Width           =   780
            End
            Begin VB.Line Line1 
               Index           =   2
               X1              =   585
               X2              =   1110
               Y1              =   735
               Y2              =   735
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "2"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   2
               Left            =   795
               TabIndex        =   43
               Top             =   795
               Width           =   120
            End
         End
         Begin VB.Frame fraFormulasFiguras 
            Height          =   1275
            Index           =   4
            Left            =   135
            TabIndex        =   39
            Top             =   435
            Width           =   1410
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "S = bL + b2 x h"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   4
               Left            =   45
               TabIndex        =   41
               Top             =   495
               Width           =   1320
            End
            Begin VB.Line Line1 
               Index           =   1
               X1              =   375
               X2              =   1065
               Y1              =   735
               Y2              =   735
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "2"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   1
               Left            =   675
               TabIndex        =   40
               Top             =   795
               Width           =   120
            End
         End
         Begin VB.Frame fraFormulasFiguras 
            Height          =   1275
            Index           =   3
            Left            =   135
            TabIndex        =   36
            Top             =   435
            Width           =   1410
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "2"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   0
               Left            =   675
               TabIndex        =   38
               Top             =   795
               Width           =   120
            End
            Begin VB.Line Line1 
               Index           =   0
               X1              =   375
               X2              =   1065
               Y1              =   735
               Y2              =   735
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "S = bL + b2 x h"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   3
               Left            =   45
               TabIndex        =   37
               Top             =   495
               Width           =   1320
            End
         End
         Begin VB.Frame fraFormulasFiguras 
            Height          =   1275
            Index           =   2
            Left            =   135
            TabIndex        =   34
            Top             =   435
            Width           =   1410
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "S = bh"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   2
               Left            =   435
               TabIndex        =   35
               Top             =   615
               Width           =   570
            End
         End
         Begin VB.Frame fraFormulasFiguras 
            Height          =   1275
            Index           =   1
            Left            =   135
            TabIndex        =   32
            Top             =   435
            Width           =   1410
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "S = bh"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   1
               Left            =   435
               TabIndex        =   33
               Top             =   615
               Width           =   570
            End
         End
         Begin VB.Frame fraFormulasFiguras 
            Height          =   1275
            Index           =   0
            Left            =   135
            TabIndex        =   30
            Top             =   435
            Width           =   1410
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "S = L²"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   0
               Left            =   435
               TabIndex        =   31
               Top             =   615
               Width           =   525
            End
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Área = S (superfície)"
            Height          =   195
            Left            =   120
            TabIndex        =   29
            Top             =   240
            Width           =   1455
         End
      End
      Begin VB.Frame fraCalculo 
         Caption         =   " Cálculo da área "
         ForeColor       =   &H8000000D&
         Height          =   1830
         Left            =   4245
         TabIndex        =   12
         Top             =   4050
         Width           =   6690
         Begin VB.CommandButton cmdCalcular 
            Caption         =   "Calc&ular"
            Height          =   405
            Left            =   5445
            TabIndex        =   64
            Top             =   1290
            Width           =   1155
         End
         Begin VB.TextBox txtL 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000A&
            Enabled         =   0   'False
            Height          =   315
            Left            =   4920
            TabIndex        =   24
            Top             =   315
            Width           =   1470
         End
         Begin VB.TextBox txtArea 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   1005
            Locked          =   -1  'True
            TabIndex        =   22
            Top             =   1335
            Width           =   1755
         End
         Begin VB.TextBox txtB2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000A&
            Enabled         =   0   'False
            Height          =   315
            Left            =   2880
            TabIndex        =   19
            Top             =   705
            Width           =   1470
         End
         Begin VB.TextBox txtBL 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000A&
            Enabled         =   0   'False
            Height          =   315
            Left            =   2880
            TabIndex        =   15
            Top             =   315
            Width           =   1470
         End
         Begin VB.TextBox txtH 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000A&
            Enabled         =   0   'False
            Height          =   315
            Left            =   735
            TabIndex        =   14
            Top             =   705
            Width           =   1470
         End
         Begin VB.TextBox txtB 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000A&
            Enabled         =   0   'False
            Height          =   315
            Left            =   735
            TabIndex        =   13
            Top             =   315
            Width           =   1470
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "m²"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   2895
            TabIndex        =   61
            Top             =   1365
            Width           =   210
         End
         Begin VB.Label lblL 
            AutoSize        =   -1  'True
            Caption         =   "L = "
            Height          =   195
            Left            =   4515
            TabIndex        =   23
            Top             =   375
            Width           =   270
         End
         Begin VB.Label lblarea 
            AutoSize        =   -1  'True
            Caption         =   "Área:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   360
            TabIndex        =   21
            Top             =   1380
            Width           =   465
         End
         Begin VB.Label lblB2 
            AutoSize        =   -1  'True
            Caption         =   "b2 ="
            Height          =   195
            Left            =   2460
            TabIndex        =   20
            Top             =   750
            Width           =   315
         End
         Begin VB.Label lblBL 
            AutoSize        =   -1  'True
            Caption         =   "bL ="
            Height          =   195
            Left            =   2460
            TabIndex        =   18
            Top             =   375
            Width           =   315
         End
         Begin VB.Label lblH 
            AutoSize        =   -1  'True
            Caption         =   "h = "
            Height          =   195
            Left            =   315
            TabIndex        =   17
            Top             =   750
            Width           =   270
         End
         Begin VB.Label lblB 
            AutoSize        =   -1  'True
            Caption         =   "b = "
            Height          =   195
            Left            =   315
            TabIndex        =   16
            Top             =   375
            Width           =   270
         End
      End
      Begin VB.OptionButton optFigura 
         Height          =   1635
         Index           =   9
         Left            =   8925
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   2250
         Width           =   2040
      End
      Begin VB.OptionButton optFigura 
         Height          =   1635
         Index           =   10
         Left            =   315
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   4005
         Width           =   2040
      End
      Begin VB.OptionButton optFigura 
         Height          =   1635
         Index           =   8
         Left            =   8925
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   510
         Width           =   2040
      End
      Begin VB.OptionButton optFigura 
         Height          =   1635
         Index           =   7
         Left            =   6765
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   2250
         Width           =   2040
      End
      Begin VB.OptionButton optFigura 
         Height          =   1635
         Index           =   6
         Left            =   4620
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   2250
         Width           =   2040
      End
      Begin VB.OptionButton optFigura 
         Height          =   1635
         Index           =   5
         Left            =   2460
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   2250
         Width           =   2040
      End
      Begin VB.OptionButton optFigura 
         Height          =   1635
         Index           =   2
         Left            =   4620
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   510
         Width           =   2040
      End
      Begin VB.OptionButton optFigura 
         Height          =   1635
         Index           =   1
         Left            =   2460
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   510
         Width           =   2040
      End
      Begin VB.OptionButton optFigura 
         Height          =   1635
         Index           =   0
         Left            =   315
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   510
         UseMaskColor    =   -1  'True
         Width           =   2040
      End
      Begin VB.OptionButton optFigura 
         Height          =   1635
         Index           =   4
         Left            =   315
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   2250
         Width           =   2040
      End
      Begin VB.OptionButton optFigura 
         Height          =   1635
         Index           =   3
         Left            =   6765
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   510
         Width           =   2040
      End
      Begin VB.Shape shpBorda 
         Height          =   1665
         Index           =   10
         Left            =   300
         Top             =   3990
         Width           =   2070
      End
      Begin VB.Shape shpBorda 
         Height          =   1665
         Index           =   9
         Left            =   8910
         Top             =   2220
         Width           =   2070
      End
      Begin VB.Shape shpBorda 
         Height          =   1665
         Index           =   8
         Left            =   8910
         Top             =   495
         Width           =   2070
      End
      Begin VB.Shape shpBorda 
         Height          =   1665
         Index           =   7
         Left            =   6750
         Top             =   2235
         Width           =   2070
      End
      Begin VB.Shape shpBorda 
         Height          =   1665
         Index           =   6
         Left            =   4605
         Top             =   2235
         Width           =   2070
      End
      Begin VB.Shape shpBorda 
         Height          =   1665
         Index           =   5
         Left            =   2445
         Top             =   2235
         Width           =   2070
      End
      Begin VB.Shape shpBorda 
         Height          =   1665
         Index           =   4
         Left            =   285
         Top             =   2235
         Width           =   2085
      End
      Begin VB.Shape shpBorda 
         Height          =   1665
         Index           =   3
         Left            =   6750
         Top             =   495
         Width           =   2070
      End
      Begin VB.Shape shpBorda 
         Height          =   1665
         Index           =   2
         Left            =   4605
         Top             =   495
         Width           =   2070
      End
      Begin VB.Shape shpBorda 
         Height          =   1665
         Index           =   1
         Left            =   2445
         Top             =   495
         Width           =   2070
      End
      Begin VB.Shape shpBorda 
         BorderColor     =   &H00000000&
         Height          =   1665
         Index           =   0
         Left            =   300
         Top             =   495
         Width           =   2070
      End
   End
End
Attribute VB_Name = "frmCadImagens"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim intFigura As Integer
Dim intFlagClick As Integer
Dim dbll As Double
Dim dblB As Double
Dim dblBL As Double
Dim dblB2 As Double
Dim dblH As Double
Dim dblArea As Double

Private Sub cmdAplicar_Click()
    On Error Resume Next
    If txtArea.Text <> "" Then
        'glngAreaTerreno = txtArea.Text
'        frmCadImobiliario.txtAreaTerreno = glngAreaTerreno
    End If
    Unload Me
End Sub

Private Sub cmdCalcular_Click()
    CalculaArea
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    With Me
        Me.Left = 0
        Me.Top = 0
        Me.Height = tab_3DPasta.Height + 500
        Me.Width = tab_3DPasta.Width + 250
    End With
End Sub

Private Sub optFigura_Click(Index As Integer)
    intFlagClick = 1
    For giContador = 0 To 10
        If optFigura(giContador).Value = True Then
            optFigura(giContador).BackColor = &HC0C000 ' azul
            fraFormulasFiguras(giContador).Visible = True
            fraFormulaVazio.Visible = False
            intFigura = giContador
        Else
            optFigura(giContador).BackColor = &H8000000F 'normal
            fraFormulasFiguras(giContador).Visible = False
        End If
    Next
    DesabilitaCampos Index
End Sub

Private Sub optFigura_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Screen.MousePointer = 99
    For giContador = 0 To 10
        If giContador = Index Then
            shpBorda(Index).BorderColor = 16711680
            shpBorda(Index).BorderWidth = 3
        Else
            shpBorda(giContador).BorderColor = &H0&
            shpBorda(giContador).BorderWidth = 1
        End If
    Next
End Sub

Private Sub optKm_Click()
    On Error GoTo Err_Unidade
    If optKm.Value = True Then
        If txtB.Enabled = True Then
            txtB = txtB / 1000
        End If
        If txtL.Enabled = True Then
            txtL = txtL / 1000
        End If
        If txtBL.Enabled = True Then
            txtBL = txtBL / 1000
        End If
        If txtB2.Enabled = True Then
            txtB2 = txtB2 / 1000
        End If
        If txtH.Enabled = True Then
            txtH = txtH / 1000
        End If
    End If
    txtArea = ""
Exit Sub
Err_Unidade:
End Sub

Private Sub optMetros_Click()
    On Error GoTo Err_Unidade
    If optMetros.Value = True Then
        If txtB.Enabled = True Then
            txtB = txtB * 1000
        End If
        If txtL.Enabled = True Then
            txtL = txtL * 1000
        End If
        If txtBL.Enabled = True Then
            txtBL = txtBL * 1000
        End If
        If txtB2.Enabled = True Then
            txtB2 = txtB2 * 1000
        End If
        If txtH.Enabled = True Then
            txtH = txtH * 1000
        End If
    End If
    txtArea = ""

Exit Sub
Err_Unidade:
End Sub

Private Sub SSTab1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    For giContador = 0 To 10
        shpBorda(giContador).BorderColor = &H0&
        shpBorda(giContador).BorderWidth = 1
    Next
End Sub

Sub CalculaArea()
    On Error GoTo Err_Area
    
    If intFlagClick = 0 Then Exit Sub
    
    dblArea = 0
    
    Select Case intFigura
    Case 0  'Quadrado
        AreaQuadrado
    Case 1  'Retangulo
        AreaRetangulo
    Case 2  'Paralelogramo
        AreaParalelogramo
    Case 3  'Trapezio Retangulo
        AreaTrapezioRetangulo
    Case 4  'Trapezio Isoceles
        AreaTrapezioIsoceles
    Case 5  'Triangulo Equilatero
        AreaTrianguloEquilatero
    Case 6  'Triangulo Isoceles
        AreaTrianguloIsoceles
    Case 7  'Triangulo Escaleno
        AreaTrianguloEscaleno
    Case 8  'Triangulo Retangulo
        AreaTrianguloRetangulo
    Case 9  'Triangulo Acutangulo
        AreaTrianguloAcutangulo
    Case 10 'Triangulo Obtusangulo
        AreaTrianguloObtusangulo
    End Select


    If optKm.Value = True Then
        txtArea = gvntConvVrDoSql(dblArea * 1000000)
    Else
        txtArea = gvntConvVrDoSql(dblArea, "V")
    End If

Exit Sub
Err_Area:
    ExibeDetalheErro ""
End Sub

Sub AreaQuadrado()
    If Trim(txtL) = "" Then
        MsgBox "Valor L inválido."
        Exit Sub
    End If
    dbll = CDbl(txtL)
    If optKm.Value = True Then
    End If
    dblArea = dbll * dbll
End Sub

Sub AreaRetangulo()
    If Trim(txtB) = "" Or Trim(txtH) = "" Then
        MsgBox "Valores b e h devem ser digitados."
        Exit Sub
    End If
    dblArea = CDbl(txtB) * CDbl(txtH)
End Sub

Sub AreaParalelogramo()
    If Trim(txtB) = "" Or Trim(txtH) = "" Then
        MsgBox "Valores b e h devem ser digitados."
        Exit Sub
    End If
    dblArea = CDbl(txtB) * CDbl(txtH)
End Sub

Sub AreaTrapezioRetangulo()
    If Trim(txtBL) = "" Or Trim(txtH) = "" Or Trim(txtB2) = "" Then
        MsgBox "Valores bL e bL2 e h devem ser digitados."
        Exit Sub
    End If
    dblArea = ((CDbl(txtBL) + CDbl(txtB2)) / 2) * CDbl(txtH)
End Sub

Sub AreaTrapezioIsoceles()
    If Trim(txtBL) = "" Or Trim(txtH) = "" Or Trim(txtB2) = "" Then
        MsgBox "Valores bL e bL2 e h devem ser digitados."
        Exit Sub
    End If
    dblArea = ((CDbl(txtBL) + CDbl(txtB2)) / 2) * CDbl(txtH)
End Sub

Sub AreaTrianguloEquilatero()
    If Trim(txtB) = "" Or Trim(txtH) = "" Then
        MsgBox "Valores b e h devem ser digitados."
        Exit Sub
    End If
    dblArea = (CDbl(txtB) * CDbl(txtH)) / 2
End Sub

Sub AreaTrianguloIsoceles()
    If Trim(txtB) = "" Or Trim(txtH) = "" Then
        MsgBox "Valores b e h devem ser digitados."
        Exit Sub
    End If
    dblArea = (CDbl(txtB) * CDbl(txtH)) / 2
End Sub

Sub AreaTrianguloEscaleno()
    If Trim(txtB) = "" Or Trim(txtH) = "" Then
        MsgBox "Valores b e h devem ser digitados."
        Exit Sub
    End If
    dblArea = (CDbl(txtB) * CDbl(txtH)) / 2
End Sub

Sub AreaTrianguloRetangulo()
    If Trim(txtB) = "" Or Trim(txtH) = "" Then
        MsgBox "Valores b e h devem ser digitados."
        Exit Sub
    End If
    dblArea = (CDbl(txtB) * CDbl(txtH)) / 2
End Sub

Sub AreaTrianguloAcutangulo()
    If Trim(txtB) = "" Or Trim(txtH) = "" Then
        MsgBox "Valores b e h devem ser digitados."
        Exit Sub
    End If
    dblArea = (CDbl(txtB) * CDbl(txtH)) / 2
End Sub

Sub AreaTrianguloObtusangulo()
    If Trim(txtB) = "" Or Trim(txtH) = "" Then
        MsgBox "Valores b e h devem ser digitados."
        Exit Sub
    End If
    dblArea = (CDbl(txtB) * CDbl(txtH)) / 2
End Sub

Sub DesabilitaCampos(intFlag As Integer)
    Limpa_Controles Me, True, False, False, False, False
    Select Case intFlag
            Case 0  'L
                txtB.Enabled = False: txtB.BackColor = &H8000000A
                txtB2.Enabled = False: txtB2.BackColor = &H8000000A
                txtBL.Enabled = False: txtBL.BackColor = &H8000000A
                txtH.Enabled = False: txtH.BackColor = &H8000000A
                txtL.Enabled = True: txtL.BackColor = &H80000005
            Case 1, 2, 5, 6, 7, 8, 9, 10 'b h
                txtB.Enabled = True: txtB.BackColor = &H80000005
                txtB2.Enabled = False: txtB2.BackColor = &H8000000A
                txtBL.Enabled = False: txtBL.BackColor = &H8000000A
                txtH.Enabled = True: txtH.BackColor = &H80000005
                txtL.Enabled = False: txtL.BackColor = &H8000000A
            Case 3, 4 'bl b2 h
                txtB.Enabled = False: txtB.BackColor = &H8000000A
                txtB2.Enabled = True: txtB2.BackColor = &H80000005
                txtBL.Enabled = True: txtBL.BackColor = &H80000005
                txtH.Enabled = True: txtH.BackColor = &H80000005
                txtL.Enabled = False: txtL.BackColor = &H8000000A
            Case Else
                txtB.Enabled = False: txtB.BackColor = &H8000000A
                txtB2.Enabled = False: txtB2.BackColor = &H8000000A
                txtBL.Enabled = False: txtBL.BackColor = &H8000000A
                txtH.Enabled = False: txtH.BackColor = &H8000000A
                txtL.Enabled = False: txtL.BackColor = &H8000000A
    End Select
End Sub

Private Sub txtArea_Change()
    With txtArea
        If .Text <> "" Then
            cmdAplicar.Enabled = True
        Else
            cmdAplicar.Enabled = False
        End If
    End With
End Sub

Private Sub txtArea_GotFocus()
    MarcaCampo txtArea
End Sub

Private Sub txtB_GotFocus()
    MarcaCampo txtB
End Sub

Private Sub txtB_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "V", txtB
End Sub

Private Sub txtB2_GotFocus()
    MarcaCampo txtB2
End Sub

Private Sub txtB2_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "V", txtB2
End Sub

Private Sub txtBL_GotFocus()
    MarcaCampo txtBL
End Sub

Private Sub txtBL_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "V", txtBL
End Sub

Private Sub txtH_GotFocus()
    MarcaCampo txtH
End Sub

Private Sub txtH_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "V", txtH
End Sub

Private Sub txtL_GotFocus()
    MarcaCampo txtL
End Sub

Private Sub txtL_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "V", txtL
End Sub
