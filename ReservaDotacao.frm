VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frmReservaDotacao 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reservas de Dotação"
   ClientHeight    =   6945
   ClientLeft      =   1545
   ClientTop       =   2655
   ClientWidth     =   10740
   HelpContextID   =   40
   Icon            =   "ReservaDotacao.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6945
   ScaleWidth      =   10740
   Begin VB.Frame fra_ProgramaDeTrabalho 
      Caption         =   " Dotação "
      Height          =   2655
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   10575
      Begin VB.CommandButton cmd_historico 
         Height          =   300
         Left            =   10110
         Picture         =   "ReservaDotacao.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   42
         TabStop         =   0   'False
         ToolTipText     =   "Clique para cadastar Tipo de Crédito"
         Top             =   2220
         Width           =   360
      End
      Begin VB.CommandButton cmd_ProgramaTrabalho 
         Height          =   300
         Left            =   7890
         Picture         =   "ReservaDotacao.frx":0396
         Style           =   1  'Graphical
         TabIndex        =   41
         TabStop         =   0   'False
         ToolTipText     =   "Clique para cadastar Tipo de Recurso"
         Top             =   640
         Width           =   360
      End
      Begin VB.TextBox txt_TotalBloqueado 
         BackColor       =   &H80000000&
         Enabled         =   0   'False
         Height          =   285
         Left            =   5225
         TabIndex        =   39
         Top             =   1005
         Width           =   1005
      End
      Begin VB.TextBox txtintExercicio 
         Height          =   285
         Left            =   4830
         MaxLength       =   4
         OLEDragMode     =   1  'Automatic
         TabIndex        =   5
         Top             =   255
         Width           =   555
      End
      Begin VB.TextBox txt_ValorProgramaTrabalho 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000000&
         Enabled         =   0   'False
         Height          =   285
         Left            =   796
         MaxLength       =   25
         MultiLine       =   -1  'True
         OLEDragMode     =   1  'Automatic
         OLEDropMode     =   1  'Manual
         TabIndex        =   16
         Top             =   1005
         Width           =   1005
      End
      Begin VB.TextBox txt_Saldo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000000&
         Enabled         =   0   'False
         Height          =   285
         Left            =   2973
         MaxLength       =   25
         MultiLine       =   -1  'True
         OLEDragMode     =   1  'Automatic
         OLEDropMode     =   1  'Manual
         TabIndex        =   18
         Top             =   1005
         Width           =   1005
      End
      Begin VB.ComboBox cbo_intProgramaTrabalho 
         Height          =   315
         ItemData        =   "ReservaDotacao.frx":0720
         Left            =   2610
         List            =   "ReservaDotacao.frx":0722
         OLEDragMode     =   1  'Automatic
         TabIndex        =   11
         ToolTipText     =   "Código do programa de trabalho"
         Top             =   630
         Width           =   1035
      End
      Begin VB.ComboBox cbointProgramaTrabalho 
         Height          =   315
         Left            =   3660
         OLEDragMode     =   1  'Automatic
         Sorted          =   -1  'True
         TabIndex        =   12
         ToolTipText     =   "Código do programa de trabalho"
         Top             =   630
         Width           =   4185
      End
      Begin VB.Frame fra_Historico 
         Caption         =   " Histórico "
         Height          =   795
         Left            =   120
         TabIndex        =   23
         Top             =   1335
         Width           =   10380
         Begin VB.TextBox txtstrHistorico 
            Height          =   495
            Left            =   90
            MaxLength       =   500
            MultiLine       =   -1  'True
            OLEDragMode     =   1  'Automatic
            ScrollBars      =   2  'Vertical
            TabIndex        =   24
            Top             =   210
            Width           =   10215
         End
      End
      Begin VB.ComboBox cbo_Historico 
         Height          =   315
         Left            =   120
         OLEDragMode     =   1  'Automatic
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   25
         ToolTipText     =   "Histórico padrão"
         Top             =   2205
         Width           =   9945
      End
      Begin VB.TextBox txt_TotalReserva 
         BackColor       =   &H80000000&
         Enabled         =   0   'False
         Height          =   285
         Left            =   9480
         TabIndex        =   22
         Top             =   1005
         Width           =   1005
      End
      Begin VB.TextBox txt_TotalEmpenho 
         BackColor       =   &H80000000&
         Enabled         =   0   'False
         Height          =   285
         Left            =   7387
         TabIndex        =   20
         Top             =   1005
         Width           =   1005
      End
      Begin VB.TextBox txtstrSolicitacao 
         Height          =   285
         Left            =   3645
         MaxLength       =   10
         OLEDragMode     =   1  'Automatic
         TabIndex        =   4
         Top             =   255
         Width           =   1155
      End
      Begin VB.TextBox txtstrSolicitante 
         Height          =   285
         Left            =   6285
         MaxLength       =   50
         OLEDragMode     =   1  'Automatic
         TabIndex        =   7
         Top             =   255
         Width           =   4200
      End
      Begin VB.TextBox txtdblValor 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   8760
         MaxLength       =   25
         MultiLine       =   -1  'True
         OLEDragMode     =   1  'Automatic
         TabIndex        =   14
         Top             =   630
         Width           =   1725
      End
      Begin VB.TextBox txtdtmData 
         Height          =   285
         Left            =   765
         OLEDragMode     =   1  'Automatic
         TabIndex        =   9
         Top             =   630
         Width           =   1095
      End
      Begin VB.TextBox txtintNumero 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   765
         MaxLength       =   25
         MultiLine       =   -1  'True
         OLEDragMode     =   1  'Automatic
         TabIndex        =   2
         Top             =   255
         Width           =   1095
      End
      Begin VB.Label lbl_Bloqueado 
         AutoSize        =   -1  'True
         Caption         =   "Tot. Bloqueado"
         Height          =   195
         Left            =   4054
         TabIndex        =   40
         Top             =   1095
         Width           =   1095
      End
      Begin VB.Label lbl_ValorProgramaTrabalho 
         AutoSize        =   -1  'True
         Caption         =   "Dotação"
         Height          =   195
         Left            =   105
         TabIndex        =   15
         Top             =   1095
         Width           =   615
      End
      Begin VB.Label lbl_Saldo 
         AutoSize        =   -1  'True
         Caption         =   "Sld. Dotação"
         Height          =   195
         Left            =   1877
         TabIndex        =   17
         Top             =   1095
         Width           =   1020
      End
      Begin VB.Label lbl_ProgramaTrabalho 
         AutoSize        =   -1  'True
         Caption         =   "Dotação"
         Height          =   195
         Left            =   1935
         TabIndex        =   10
         Top             =   720
         Width           =   615
      End
      Begin VB.Label lbl_TotalReserva 
         AutoSize        =   -1  'True
         Caption         =   "Tot. Reserva"
         Height          =   195
         Left            =   8468
         TabIndex        =   21
         Top             =   1095
         Width           =   930
      End
      Begin VB.Label lbl_TotalEmpenho 
         AutoSize        =   -1  'True
         Caption         =   "Tot. Empenho"
         Height          =   195
         Left            =   6306
         TabIndex        =   19
         Top             =   1095
         Width           =   1005
      End
      Begin VB.Label lblstrNumeroSolicitacao 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Solicitação de Compras"
         Height          =   195
         Left            =   1920
         TabIndex        =   3
         Top             =   345
         Width           =   1665
      End
      Begin VB.Label lblstrSolicitante 
         AutoSize        =   -1  'True
         Caption         =   "Solicitante"
         Height          =   195
         Left            =   5505
         TabIndex        =   6
         Top             =   345
         Width           =   735
      End
      Begin VB.Label lbldblValor 
         AutoSize        =   -1  'True
         Caption         =   "Valor"
         Height          =   195
         Left            =   8340
         TabIndex        =   13
         Top             =   720
         Width           =   360
      End
      Begin VB.Label lbldtmDataReserva 
         AutoSize        =   -1  'True
         Caption         =   "Data "
         Height          =   195
         Left            =   300
         TabIndex        =   8
         Top             =   720
         Width           =   390
      End
      Begin VB.Label lblintNumero 
         AutoSize        =   -1  'True
         Caption         =   "Reserva"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   345
         Width           =   600
      End
   End
   Begin VB.TextBox txt_tmp 
      Height          =   285
      Left            =   6120
      TabIndex        =   38
      Text            =   "txt_tmp"
      Top             =   0
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtPKId 
      Height          =   270
      Left            =   6990
      TabIndex        =   37
      Top             =   30
      Visible         =   0   'False
      Width           =   645
   End
   Begin TabDlg.SSTab tab_3DReservar 
      Height          =   2055
      Left            =   60
      TabIndex        =   26
      Top             =   2760
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   3625
      _Version        =   393216
      Style           =   1
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "Reservar"
      TabPicture(0)   =   "ReservaDotacao.frx":0724
      Tab(0).ControlEnabled=   0   'False
      Tab(0).ControlCount=   0
      TabCaption(1)   =   "Cancelamentos"
      TabPicture(1)   =   "ReservaDotacao.frx":0740
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "lbl_Data"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "lbl_ValorCancelado"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "tdb_Cancelado"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "txt_DataCancelamento"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "txt_ValorCancelado"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "fra_Hist"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "cbo_HistoricoCancelamento"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "cmd_HistoricoCancelamento"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).ControlCount=   8
      TabCaption(2)   =   "Empenhos"
      TabPicture(2)   =   "ReservaDotacao.frx":075C
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "tdb_Empenho"
      Tab(2).ControlCount=   1
      Begin VB.CommandButton cmd_HistoricoCancelamento 
         Height          =   300
         Left            =   3960
         Picture         =   "ReservaDotacao.frx":0778
         Style           =   1  'Graphical
         TabIndex        =   43
         TabStop         =   0   'False
         ToolTipText     =   "Clique para cadastar Tipo de Crédito"
         Top             =   1650
         Width           =   360
      End
      Begin VB.ComboBox cbo_HistoricoCancelamento 
         Height          =   315
         Left            =   120
         OLEDragMode     =   1  'Automatic
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   30
         ToolTipText     =   "Histórico padrão"
         Top             =   1650
         Width           =   3825
      End
      Begin VB.Frame fra_Hist 
         Caption         =   " Histórico "
         Height          =   885
         Left            =   120
         TabIndex        =   36
         Top             =   690
         Width           =   4185
         Begin VB.TextBox txt_HistoricoCancelamento 
            Height          =   615
            Left            =   90
            MaxLength       =   500
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   29
            Top             =   180
            Width           =   3990
         End
      End
      Begin VB.TextBox txt_ValorCancelado 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2160
         MaxLength       =   25
         TabIndex        =   28
         Top             =   390
         Width           =   1725
      End
      Begin VB.TextBox txt_DataCancelamento 
         Height          =   285
         Left            =   600
         TabIndex        =   27
         Top             =   390
         Width           =   975
      End
      Begin TrueOleDBGrid70.TDBGrid tdb_Cancelado 
         Height          =   1605
         Left            =   4620
         TabIndex        =   31
         Top             =   360
         Width           =   5685
         _ExtentX        =   10028
         _ExtentY        =   2831
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "PKID"
         Columns(0).DataField=   "PKId"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Reserva"
         Columns(1).DataField=   "intReservaDotacao"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Número Cancto."
         Columns(2).DataField=   "intNumero"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Data"
         Columns(3).DataField=   "dtmData"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "Valor"
         Columns(4).DataField=   "dblValor"
         Columns(4).NumberFormat=   "Standard"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "Historico"
         Columns(5).DataField=   "Historico"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   6
         Splits(0)._UserFlags=   0
         Splits(0).MarqueeStyle=   3
         Splits(0).RecordSelectors=   0   'False
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).ScrollBars=   2
         Splits(0).DividerColor=   12632256
         Splits(0).FilterBar=   -1  'True
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=6"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0).AllowSizing=0"
         Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
         Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(8)=   "Column(1).Width=2725"
         Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=2646"
         Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(13)=   "Column(2).Width=2514"
         Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=2434"
         Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(17)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(18)=   "Column(3).Width=1773"
         Splits(0)._ColumnProps(19)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(20)=   "Column(3)._WidthInPix=1693"
         Splits(0)._ColumnProps(21)=   "Column(3)._EditAlways=0"
         Splits(0)._ColumnProps(22)=   "Column(3)._ColStyle=516"
         Splits(0)._ColumnProps(23)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(24)=   "Column(4).Width=2461"
         Splits(0)._ColumnProps(25)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(26)=   "Column(4)._WidthInPix=2381"
         Splits(0)._ColumnProps(27)=   "Column(4)._EditAlways=0"
         Splits(0)._ColumnProps(28)=   "Column(4)._ColStyle=514"
         Splits(0)._ColumnProps(29)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(30)=   "Column(5).Width=2725"
         Splits(0)._ColumnProps(31)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(32)=   "Column(5)._WidthInPix=2646"
         Splits(0)._ColumnProps(33)=   "Column(5)._EditAlways=0"
         Splits(0)._ColumnProps(34)=   "Column(5).Visible=0"
         Splits(0)._ColumnProps(35)=   "Column(5).Order=6"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   3
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         DefColWidth     =   0
         HeadLines       =   1
         FootLines       =   1
         MultipleLines   =   0
         CellTips        =   1
         CellTipsWidth   =   0
         DeadAreaBackColor=   12632256
         RowDividerColor =   12632256
         RowSubDividerColor=   12632256
         DirectionAfterEnter=   1
         MaxRows         =   250000
         ViewColumnCaptionWidth=   0
         ViewColumnWidth =   0
         _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
         _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
         _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
         _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=18,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
         _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
         _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bgcolor=&H8000000F&"
         _StyleDefs(11)  =   ":id=2,.fgcolor=&H80000012&,.borderColor=&H80000009&"
         _StyleDefs(12)  =   "FooterStyle:id=3,.parent=1,.namedParent=35"
         _StyleDefs(13)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(14)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
         _StyleDefs(15)  =   "EditorStyle:id=7,.parent=1"
         _StyleDefs(16)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38,.bgcolor=&H80000014&"
         _StyleDefs(17)  =   ":id=8,.fgcolor=&H80000012&"
         _StyleDefs(18)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
         _StyleDefs(19)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
         _StyleDefs(20)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(21)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
         _StyleDefs(22)  =   "Splits(0).Style:id=13,.parent=1"
         _StyleDefs(23)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
         _StyleDefs(24)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
         _StyleDefs(25)  =   "Splits(0).FooterStyle:id=15,.parent=3"
         _StyleDefs(26)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
         _StyleDefs(27)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
         _StyleDefs(28)  =   "Splits(0).EditorStyle:id=17,.parent=7"
         _StyleDefs(29)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
         _StyleDefs(30)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
         _StyleDefs(31)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
         _StyleDefs(32)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(33)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(34)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
         _StyleDefs(35)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
         _StyleDefs(36)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
         _StyleDefs(37)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
         _StyleDefs(38)  =   "Splits(0).Columns(1).Style:id=58,.parent=13"
         _StyleDefs(39)  =   "Splits(0).Columns(1).HeadingStyle:id=55,.parent=14"
         _StyleDefs(40)  =   "Splits(0).Columns(1).FooterStyle:id=56,.parent=15"
         _StyleDefs(41)  =   "Splits(0).Columns(1).EditorStyle:id=57,.parent=17"
         _StyleDefs(42)  =   "Splits(0).Columns(2).Style:id=32,.parent=13"
         _StyleDefs(43)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
         _StyleDefs(44)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
         _StyleDefs(45)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
         _StyleDefs(46)  =   "Splits(0).Columns(3).Style:id=46,.parent=13"
         _StyleDefs(47)  =   "Splits(0).Columns(3).HeadingStyle:id=43,.parent=14,.alignment=2"
         _StyleDefs(48)  =   "Splits(0).Columns(3).FooterStyle:id=44,.parent=15"
         _StyleDefs(49)  =   "Splits(0).Columns(3).EditorStyle:id=45,.parent=17"
         _StyleDefs(50)  =   "Splits(0).Columns(4).Style:id=50,.parent=13,.alignment=1"
         _StyleDefs(51)  =   "Splits(0).Columns(4).HeadingStyle:id=47,.parent=14,.alignment=2"
         _StyleDefs(52)  =   "Splits(0).Columns(4).FooterStyle:id=48,.parent=15"
         _StyleDefs(53)  =   "Splits(0).Columns(4).EditorStyle:id=49,.parent=17"
         _StyleDefs(54)  =   "Splits(0).Columns(5).Style:id=54,.parent=13"
         _StyleDefs(55)  =   "Splits(0).Columns(5).HeadingStyle:id=51,.parent=14"
         _StyleDefs(56)  =   "Splits(0).Columns(5).FooterStyle:id=52,.parent=15"
         _StyleDefs(57)  =   "Splits(0).Columns(5).EditorStyle:id=53,.parent=17"
         _StyleDefs(58)  =   "Named:id=33:Normal"
         _StyleDefs(59)  =   ":id=33,.parent=0"
         _StyleDefs(60)  =   "Named:id=34:Heading"
         _StyleDefs(61)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(62)  =   ":id=34,.wraptext=-1"
         _StyleDefs(63)  =   "Named:id=35:Footing"
         _StyleDefs(64)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(65)  =   "Named:id=36:Selected"
         _StyleDefs(66)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(67)  =   "Named:id=37:Caption"
         _StyleDefs(68)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(69)  =   "Named:id=38:HighlightRow"
         _StyleDefs(70)  =   ":id=38,.parent=33,.bgcolor=&H80000014&,.fgcolor=&H80000012&"
         _StyleDefs(71)  =   "Named:id=39:EvenRow"
         _StyleDefs(72)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(73)  =   "Named:id=40:OddRow"
         _StyleDefs(74)  =   ":id=40,.parent=33"
         _StyleDefs(75)  =   "Named:id=41:RecordSelector"
         _StyleDefs(76)  =   ":id=41,.parent=34"
         _StyleDefs(77)  =   "Named:id=42:FilterBar"
         _StyleDefs(78)  =   ":id=42,.parent=33"
      End
      Begin TrueOleDBGrid70.TDBGrid tdb_Empenho 
         Height          =   1515
         Left            =   -74895
         TabIndex        =   32
         Top             =   420
         Width           =   10350
         _ExtentX        =   18256
         _ExtentY        =   2672
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "PKId"
         Columns(0).DataField=   "PKId"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Número"
         Columns(1).DataField=   "intNumero"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Dt Empenho"
         Columns(2).DataField=   "dtmData"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Empenhado"
         Columns(3).DataField=   "dblEmpenhado"
         Columns(3).NumberFormat=   "Standard"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "Anulado"
         Columns(4).DataField=   "dblAnulado"
         Columns(4).NumberFormat=   "Standard"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "Saldo"
         Columns(5).DataField=   "dblValor"
         Columns(5).NumberFormat=   "Standard"
         Columns(5).EditMaskUpdate=   -1  'True
         Columns(5).EditMaskRight=   -1  'True
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "Programa de Trabalho"
         Columns(6).DataField=   "strCodigo"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   7
         Splits(0)._UserFlags=   0
         Splits(0).MarqueeStyle=   3
         Splits(0).RecordSelectors=   0   'False
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).ScrollBars=   2
         Splits(0).DividerColor=   13160664
         Splits(0).FilterBar=   -1  'True
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=7"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0).AllowSizing=0"
         Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
         Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(8)=   "Column(1).Width=1508"
         Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=1429"
         Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(13)=   "Column(2).Width=1826"
         Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=1746"
         Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(17)=   "Column(2)._ColStyle=1"
         Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(19)=   "Column(3).Width=1826"
         Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=1746"
         Splits(0)._ColumnProps(22)=   "Column(3)._EditAlways=0"
         Splits(0)._ColumnProps(23)=   "Column(3)._ColStyle=2"
         Splits(0)._ColumnProps(24)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(25)=   "Column(4).Width=1826"
         Splits(0)._ColumnProps(26)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(27)=   "Column(4)._WidthInPix=1746"
         Splits(0)._ColumnProps(28)=   "Column(4)._EditAlways=0"
         Splits(0)._ColumnProps(29)=   "Column(4)._ColStyle=2"
         Splits(0)._ColumnProps(30)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(31)=   "Column(5).Width=1826"
         Splits(0)._ColumnProps(32)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(33)=   "Column(5)._WidthInPix=1746"
         Splits(0)._ColumnProps(34)=   "Column(5)._EditAlways=0"
         Splits(0)._ColumnProps(35)=   "Column(5)._ColStyle=2"
         Splits(0)._ColumnProps(36)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(37)=   "Column(6).Width=8255"
         Splits(0)._ColumnProps(38)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(39)=   "Column(6)._WidthInPix=8176"
         Splits(0)._ColumnProps(40)=   "Column(6)._EditAlways=0"
         Splits(0)._ColumnProps(41)=   "Column(6).Order=7"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   3
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         AllowUpdate     =   0   'False
         DefColWidth     =   0
         HeadLines       =   1
         FootLines       =   1
         MultipleLines   =   0
         CellTips        =   1
         CellTipsWidth   =   0
         MultiSelect     =   0
         DeadAreaBackColor=   13160664
         RowDividerColor =   13160664
         RowSubDividerColor=   13160664
         DirectionAfterEnter=   1
         MaxRows         =   250000
         ViewColumnCaptionWidth=   0
         ViewColumnWidth =   0
         _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
         _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
         _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
         _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
         _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
         _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(12)  =   ":id=2,.fontname=MS Sans Serif"
         _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(15)  =   ":id=3,.fontname=MS Sans Serif"
         _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36,.bgcolor=&H80000002&"
         _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1"
         _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38,.bgcolor=&H80000014&"
         _StyleDefs(20)  =   ":id=8,.fgcolor=&H80000012&"
         _StyleDefs(21)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
         _StyleDefs(22)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
         _StyleDefs(23)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(24)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
         _StyleDefs(25)  =   "Splits(0).Style:id=13,.parent=1"
         _StyleDefs(26)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
         _StyleDefs(27)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
         _StyleDefs(28)  =   "Splits(0).FooterStyle:id=15,.parent=3"
         _StyleDefs(29)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
         _StyleDefs(30)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
         _StyleDefs(31)  =   "Splits(0).EditorStyle:id=17,.parent=7"
         _StyleDefs(32)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
         _StyleDefs(33)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
         _StyleDefs(34)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
         _StyleDefs(35)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(36)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(37)  =   "Splits(0).Columns(0).Style:id=58,.parent=13"
         _StyleDefs(38)  =   "Splits(0).Columns(0).HeadingStyle:id=55,.parent=14"
         _StyleDefs(39)  =   "Splits(0).Columns(0).FooterStyle:id=56,.parent=15"
         _StyleDefs(40)  =   "Splits(0).Columns(0).EditorStyle:id=57,.parent=17"
         _StyleDefs(41)  =   "Splits(0).Columns(1).Style:id=54,.parent=13"
         _StyleDefs(42)  =   "Splits(0).Columns(1).HeadingStyle:id=51,.parent=14"
         _StyleDefs(43)  =   "Splits(0).Columns(1).FooterStyle:id=52,.parent=15"
         _StyleDefs(44)  =   "Splits(0).Columns(1).EditorStyle:id=53,.parent=17"
         _StyleDefs(45)  =   "Splits(0).Columns(2).Style:id=50,.parent=13,.alignment=2"
         _StyleDefs(46)  =   "Splits(0).Columns(2).HeadingStyle:id=47,.parent=14"
         _StyleDefs(47)  =   "Splits(0).Columns(2).FooterStyle:id=48,.parent=15"
         _StyleDefs(48)  =   "Splits(0).Columns(2).EditorStyle:id=49,.parent=17"
         _StyleDefs(49)  =   "Splits(0).Columns(3).Style:id=32,.parent=13,.alignment=1"
         _StyleDefs(50)  =   "Splits(0).Columns(3).HeadingStyle:id=29,.parent=14"
         _StyleDefs(51)  =   "Splits(0).Columns(3).FooterStyle:id=30,.parent=15"
         _StyleDefs(52)  =   "Splits(0).Columns(3).EditorStyle:id=31,.parent=17"
         _StyleDefs(53)  =   "Splits(0).Columns(4).Style:id=62,.parent=13,.alignment=1"
         _StyleDefs(54)  =   "Splits(0).Columns(4).HeadingStyle:id=59,.parent=14"
         _StyleDefs(55)  =   "Splits(0).Columns(4).FooterStyle:id=60,.parent=15"
         _StyleDefs(56)  =   "Splits(0).Columns(4).EditorStyle:id=61,.parent=17"
         _StyleDefs(57)  =   "Splits(0).Columns(5).Style:id=46,.parent=13,.alignment=1"
         _StyleDefs(58)  =   "Splits(0).Columns(5).HeadingStyle:id=43,.parent=14"
         _StyleDefs(59)  =   "Splits(0).Columns(5).FooterStyle:id=44,.parent=15"
         _StyleDefs(60)  =   "Splits(0).Columns(5).EditorStyle:id=45,.parent=17"
         _StyleDefs(61)  =   "Splits(0).Columns(6).Style:id=28,.parent=13"
         _StyleDefs(62)  =   "Splits(0).Columns(6).HeadingStyle:id=25,.parent=14"
         _StyleDefs(63)  =   "Splits(0).Columns(6).FooterStyle:id=26,.parent=15"
         _StyleDefs(64)  =   "Splits(0).Columns(6).EditorStyle:id=27,.parent=17"
         _StyleDefs(65)  =   "Named:id=33:Normal"
         _StyleDefs(66)  =   ":id=33,.parent=0"
         _StyleDefs(67)  =   "Named:id=34:Heading"
         _StyleDefs(68)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(69)  =   ":id=34,.wraptext=-1"
         _StyleDefs(70)  =   "Named:id=35:Footing"
         _StyleDefs(71)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(72)  =   "Named:id=36:Selected"
         _StyleDefs(73)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(74)  =   "Named:id=37:Caption"
         _StyleDefs(75)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(76)  =   "Named:id=38:HighlightRow"
         _StyleDefs(77)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(78)  =   "Named:id=39:EvenRow"
         _StyleDefs(79)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(80)  =   "Named:id=40:OddRow"
         _StyleDefs(81)  =   ":id=40,.parent=33"
         _StyleDefs(82)  =   "Named:id=41:RecordSelector"
         _StyleDefs(83)  =   ":id=41,.parent=34"
         _StyleDefs(84)  =   "Named:id=42:FilterBar"
         _StyleDefs(85)  =   ":id=42,.parent=33"
      End
      Begin VB.Label lbl_ValorCancelado 
         AutoSize        =   -1  'True
         Caption         =   "Valor"
         Height          =   195
         Left            =   1740
         TabIndex        =   35
         Top             =   480
         Width           =   360
      End
      Begin VB.Label lbl_Data 
         AutoSize        =   -1  'True
         Caption         =   "Data "
         Height          =   195
         Left            =   180
         TabIndex        =   34
         Top             =   480
         Width           =   390
      End
   End
   Begin TrueOleDBGrid70.TDBGrid tdb_ReservaDotacao 
      Height          =   1965
      Left            =   60
      TabIndex        =   33
      Top             =   4890
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   3466
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).DataField=   "PKId"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Reserva"
      Columns(1).DataField=   "intNumero"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Data"
      Columns(2).DataField=   "dtmData"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Reservado"
      Columns(3).DataField=   "dblValor"
      Columns(3).NumberFormat=   "Standard"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "Cancelado"
      Columns(4).DataField=   "dblCanceledo"
      Columns(4).NumberFormat=   "Standard"
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "Empenhado"
      Columns(5).DataField=   "dblEmpenhado"
      Columns(5).NumberFormat=   "Standard"
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "Saldo"
      Columns(6).DataField=   "dblSaldo"
      Columns(6).NumberFormat=   "Standard"
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "Solicitante"
      Columns(7).DataField=   "strSolicitante"
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   8
      Splits(0)._UserFlags=   0
      Splits(0).MarqueeStyle=   3
      Splits(0).RecordSelectors=   0   'False
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).ScrollBars=   2
      Splits(0).DividerColor=   12632256
      Splits(0).FilterBar=   -1  'True
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=8"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0).AllowSizing=0"
      Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
      Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(8)=   "Column(1).Width=1535"
      Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=1455"
      Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(13)=   "Column(2).Width=1773"
      Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=1693"
      Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
      Splits(0)._ColumnProps(17)=   "Column(2)._ColStyle=516"
      Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(19)=   "Column(3).Width=2461"
      Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=2381"
      Splits(0)._ColumnProps(22)=   "Column(3)._EditAlways=0"
      Splits(0)._ColumnProps(23)=   "Column(3)._ColStyle=514"
      Splits(0)._ColumnProps(24)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(25)=   "Column(4).Width=2461"
      Splits(0)._ColumnProps(26)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(27)=   "Column(4)._WidthInPix=2381"
      Splits(0)._ColumnProps(28)=   "Column(4)._EditAlways=0"
      Splits(0)._ColumnProps(29)=   "Column(4)._ColStyle=514"
      Splits(0)._ColumnProps(30)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(31)=   "Column(5).Width=2725"
      Splits(0)._ColumnProps(32)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(33)=   "Column(5)._WidthInPix=2646"
      Splits(0)._ColumnProps(34)=   "Column(5)._EditAlways=0"
      Splits(0)._ColumnProps(35)=   "Column(5)._ColStyle=2"
      Splits(0)._ColumnProps(36)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(37)=   "Column(6).Width=2461"
      Splits(0)._ColumnProps(38)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(39)=   "Column(6)._WidthInPix=2381"
      Splits(0)._ColumnProps(40)=   "Column(6)._EditAlways=0"
      Splits(0)._ColumnProps(41)=   "Column(6)._ColStyle=514"
      Splits(0)._ColumnProps(42)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(43)=   "Column(7).Width=5239"
      Splits(0)._ColumnProps(44)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(45)=   "Column(7)._WidthInPix=5159"
      Splits(0)._ColumnProps(46)=   "Column(7)._EditAlways=0"
      Splits(0)._ColumnProps(47)=   "Column(7).Order=8"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      RowDividerStyle =   4
      MultipleLines   =   0
      CellTips        =   1
      CellTipsWidth   =   0
      DeadAreaBackColor=   12632256
      RowDividerColor =   12632256
      RowSubDividerColor=   12632256
      DirectionAfterEnter=   1
      MaxRows         =   250000
      ViewColumnCaptionWidth=   0
      ViewColumnWidth =   0
      _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
      _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
      _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
      _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=18,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bgcolor=&H8000000F&"
      _StyleDefs(11)  =   ":id=2,.fgcolor=&H80000012&,.borderColor=&H80000009&"
      _StyleDefs(12)  =   "FooterStyle:id=3,.parent=1,.namedParent=35"
      _StyleDefs(13)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(14)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
      _StyleDefs(15)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(16)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38,.bgcolor=&H80000014&"
      _StyleDefs(17)  =   ":id=8,.fgcolor=&H80000012&"
      _StyleDefs(18)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
      _StyleDefs(19)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
      _StyleDefs(20)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
      _StyleDefs(21)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
      _StyleDefs(22)  =   "Splits(0).Style:id=13,.parent=1"
      _StyleDefs(23)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
      _StyleDefs(24)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
      _StyleDefs(25)  =   "Splits(0).FooterStyle:id=15,.parent=3"
      _StyleDefs(26)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
      _StyleDefs(27)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
      _StyleDefs(28)  =   "Splits(0).EditorStyle:id=17,.parent=7"
      _StyleDefs(29)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
      _StyleDefs(30)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
      _StyleDefs(31)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
      _StyleDefs(32)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(33)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(34)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
      _StyleDefs(35)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
      _StyleDefs(36)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
      _StyleDefs(37)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
      _StyleDefs(38)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
      _StyleDefs(39)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
      _StyleDefs(40)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
      _StyleDefs(41)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
      _StyleDefs(42)  =   "Splits(0).Columns(2).Style:id=46,.parent=13"
      _StyleDefs(43)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14,.alignment=2"
      _StyleDefs(44)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
      _StyleDefs(45)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
      _StyleDefs(46)  =   "Splits(0).Columns(3).Style:id=50,.parent=13,.alignment=1"
      _StyleDefs(47)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14,.alignment=2"
      _StyleDefs(48)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
      _StyleDefs(49)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
      _StyleDefs(50)  =   "Splits(0).Columns(4).Style:id=58,.parent=13,.alignment=1"
      _StyleDefs(51)  =   "Splits(0).Columns(4).HeadingStyle:id=55,.parent=14,.alignment=2"
      _StyleDefs(52)  =   "Splits(0).Columns(4).FooterStyle:id=56,.parent=15"
      _StyleDefs(53)  =   "Splits(0).Columns(4).EditorStyle:id=57,.parent=17"
      _StyleDefs(54)  =   "Splits(0).Columns(5).Style:id=66,.parent=13,.alignment=1"
      _StyleDefs(55)  =   "Splits(0).Columns(5).HeadingStyle:id=63,.parent=14"
      _StyleDefs(56)  =   "Splits(0).Columns(5).FooterStyle:id=64,.parent=15"
      _StyleDefs(57)  =   "Splits(0).Columns(5).EditorStyle:id=65,.parent=17"
      _StyleDefs(58)  =   "Splits(0).Columns(6).Style:id=62,.parent=13,.alignment=1"
      _StyleDefs(59)  =   "Splits(0).Columns(6).HeadingStyle:id=59,.parent=14,.alignment=2"
      _StyleDefs(60)  =   "Splits(0).Columns(6).FooterStyle:id=60,.parent=15"
      _StyleDefs(61)  =   "Splits(0).Columns(6).EditorStyle:id=61,.parent=17"
      _StyleDefs(62)  =   "Splits(0).Columns(7).Style:id=54,.parent=13"
      _StyleDefs(63)  =   "Splits(0).Columns(7).HeadingStyle:id=51,.parent=14"
      _StyleDefs(64)  =   "Splits(0).Columns(7).FooterStyle:id=52,.parent=15"
      _StyleDefs(65)  =   "Splits(0).Columns(7).EditorStyle:id=53,.parent=17"
      _StyleDefs(66)  =   "Named:id=33:Normal"
      _StyleDefs(67)  =   ":id=33,.parent=0"
      _StyleDefs(68)  =   "Named:id=34:Heading"
      _StyleDefs(69)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(70)  =   ":id=34,.wraptext=-1"
      _StyleDefs(71)  =   "Named:id=35:Footing"
      _StyleDefs(72)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(73)  =   "Named:id=36:Selected"
      _StyleDefs(74)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(75)  =   "Named:id=37:Caption"
      _StyleDefs(76)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(77)  =   "Named:id=38:HighlightRow"
      _StyleDefs(78)  =   ":id=38,.parent=33,.bgcolor=&H80000014&,.fgcolor=&H80000012&"
      _StyleDefs(79)  =   "Named:id=39:EvenRow"
      _StyleDefs(80)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(81)  =   "Named:id=40:OddRow"
      _StyleDefs(82)  =   ":id=40,.parent=33"
      _StyleDefs(83)  =   "Named:id=41:RecordSelector"
      _StyleDefs(84)  =   ":id=41,.parent=34"
      _StyleDefs(85)  =   "Named:id=42:FilterBar"
      _StyleDefs(86)  =   ":id=42,.parent=33"
   End
End
Attribute VB_Name = "frmReservaDotacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mblnAlterando                           As Boolean
Dim mobjAux                                 As Object
Dim mblnselecionou                          As Boolean
Dim mblnClickOk                             As Boolean
Dim mstrNumero                              As String
Dim strFiltro                               As String
Dim blnSalvando                             As Boolean
Dim strGuardaValorTxt_DataCancelamento      As String
Dim strGuardaValorTxt_ValorCancelado        As String
Dim strGuardaValorTxt_HistoricoCancelamento As String
Dim intGuardaValorCbo_HistoricoCancelamento As Integer
Dim intGuardaValorTdb_ReservaDotacao        As Integer
Dim intGuardaValorCbo_IntPrograma           As Integer
Dim strGuardaValorCboIntPrograma            As String
Dim strGuardaValorExercicio                 As String
Dim strGuardaValorSolicitacao               As String
Dim dataPedido                              As Date

Private Sub cbo_Historico_Change()
    txtstrHistorico = Trim(cbo_Historico)
End Sub

Private Sub cbo_Historico_Click()
    cbo_Historico_Change
End Sub

Private Sub cbo_Historico_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", cbo_Historico
End Sub

Private Sub cbo_HistoricoCancelamento_Change()
    txt_HistoricoCancelamento = cbo_HistoricoCancelamento
End Sub

Private Sub cbo_HistoricoCancelamento_Click()
    cbo_HistoricoCancelamento_Change
End Sub

Private Sub cbo_HistoricoCancelamento_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub cbo_intProgramaTrabalho_DropDown()
    '    If Len(cbo_intProgramaTrabalho.Text) > 0 Then
    '        LeProgramaTrabalhoComReduzido cbo_intProgramaTrabalho, cbointProgramaTrabalho, gintExercicio, "SELECT PKId, intCodigoReduzido, strCodigo FROM " & gstrProgramaDeTrabalho & " WHERE intExercicio=" & gintExercicio & " AND " & gstrCONVERT(CDT_VARCHAR, "intCodigoReduzido") & " LIKE '" & cbo_intProgramaTrabalho.Text & "%' ORDER BY strCodigo"
    '    End If
End Sub

Private Sub cbo_intProgramaTrabalho_GotFocus()
    MarcaCampo cbo_intProgramaTrabalho
End Sub

Private Sub cbo_intProgramaTrabalho_LostFocus()
    'ORC1532
    ''preencheDotacaoByCodigo cbo_intProgramaTrabalho, cbointProgramaTrabalho
    'cbointProgramaTrabalho.ListIndex = cbo_intProgramaTrabalho.ListIndex
    'ORC1532
    
    'If cbo_intProgramaTrabalho.ListIndex = -1 Then LimpaDados
End Sub

Private Sub cbointProgramaTrabalho_Click()
    
    If cbointProgramaTrabalho.ListIndex <> -1 And Trim(txtDTMDATA.Text) <> "" Then
        LeProgramaTrabalho cbointProgramaTrabalho, cbo_intProgramaTrabalho, txt_tmp, _
        txt_tmp, txt_tmp, txt_tmp, _
        txt_tmp, txt_tmp, _
        txt_tmp, txt_tmp, txt_tmp, _
        txt_tmp, txt_Saldo, txt_TotalEmpenho, _
        txt_ValorProgramaTrabalho, txt_TotalBloqueado, , , txt_TotalReserva, , , , , , txtDTMDATA
    End If
    
    'LeTotalReserva
End Sub

Private Sub cbointProgramaTrabalho_DropDown()
    '    If Len(cbointProgramaTrabalho.Text) > 0 Then
    '        LeProgramaTrabalhoComReduzido cbo_intProgramaTrabalho, cbointProgramaTrabalho, gintExercicio, "SELECT PKId, intCodigoReduzido, strCodigo FROM " & gstrProgramaDeTrabalho & " WHERE intExercicio=" & gintExercicio & " AND " & gstrCONVERT(CDT_VARCHAR, "strCodigo") & " LIKE '" & cbointProgramaTrabalho.Text & "%' ORDER BY strCodigo"
    '    End If
End Sub

Private Sub cbointProgramaTrabalho_GotFocus()
    MarcaCampo cbointProgramaTrabalho
End Sub

Private Sub cbointProgramaTrabalho_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", cbointProgramaTrabalho
End Sub

Private Sub cmd_Historico_Click()
    CarregaForm frmCadHistorico, cbo_Historico
End Sub

Private Sub cmd_HistoricoCancelamento_Click()
    CarregaForm frmCadHistorico, cbo_Historico
End Sub

Private Sub cmd_ProgramaTrabalho_Click()
    frmCadProgramaDeTrabalho.blnOrcamento = True
    CarregaForm frmCadProgramaDeTrabalho, cbointProgramaTrabalho
End Sub

Private Sub cbo_intProgramaTrabalho_Click()
    
    If Val(cbo_intProgramaTrabalho.Tag) <> 2 Then
        cbointProgramaTrabalho.ListIndex = gintIndiceCBO(cbointProgramaTrabalho, _
        gstrItemData(cbo_intProgramaTrabalho))
    End If
    
End Sub

Private Sub cbo_intProgramaTrabalho_KeyPress(KeyAscii As Integer)
'ORC1532
'    If KeyAscii = 13 Then
'        'preencheDotacaoByCodigo cbo_intProgramaTrabalho, cbointProgramaTrabalho
'        CaracterValido KeyAscii
'        Exit Sub
'    End If
'ORC1532
    CaracterValido KeyAscii, "N", cbo_intProgramaTrabalho
    'ProcuraTextoDigitado KeyAscii, cbo_intProgramaTrabalho, 1
End Sub


Private Sub Form_Activate()
    gintCodSeguranca = 240
    'VirificaGradeListView Me, mblnAlterando
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar
    If mobjAux Is Nothing Then
        HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar
    Else
        HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrAplicar
    End If
    
    If mblnAlterando Then
        HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrImprimir, gstrSalvar
        HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrDeletar, gstrCancelar
    Else
        HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrImprimir, gstrSalvar
        HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrCancelar
    End If
    
    If Len(Trim$(txtDTMDATA)) = 0 Then DataAutomatica
    
End Sub

Private Sub Form_Deactivate()
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = vbKeyF1 Then
        Call_HtmlHelp Me.HelpContextID
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    mstrNumero = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar
    
End Sub

Private Sub tdb_Cancelado_Click()
    If glngQtdLinhaTDBGrid(tdb_Cancelado) = 1 Then
        tdb_Cancelado_RowColChange 0, 0
    End If
End Sub

Private Sub tdb_Cancelado_KeyPress(KeyAscii As Integer)
    'CaracterValido KeyAscii
End Sub

Private Sub tdb_Cancelado_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    With tdb_Cancelado
        If (Not .EOF And Not .BOF) Then
            mblnAlterando = True
            gCorLinhaSelecionada tdb_Cancelado
            LeCancelamento
        End If
    End With
End Sub

Private Sub tdb_ReservaDotacao_Click()
    If glngQtdLinhaTDBGrid(tdb_ReservaDotacao) = 1 Then
        tdb_ReservaDotacao_RowColChange 0, 0
        
        If cbo_intProgramaTrabalho.Text = "" Then
            mblnClickOk = True
            tdb_ReservaDotacao_RowColChange 0, 0
        End If
    End If
    
End Sub

Private Sub tdb_ReservaDotacao_DblClick()
    MantemForm gstrAplicar
End Sub

Private Sub tdb_ReservaDotacao_FilterChange()
    gblnFilraCampos tdb_ReservaDotacao
End Sub
Private Sub tdb_Empenho_FilterChange()
    gblnFilraCampos tdb_Empenho
End Sub


Private Sub tdb_ReservaDotacao_KeyDown(KeyCode As Integer, Shift As Integer)
    mblnClickOk = True
End Sub

Private Sub tdb_ReservaDotacao_KeyPress(KeyAscii As Integer)
    If tdb_ReservaDotacao.Col = 2 Then
        CaracterValido KeyAscii, "D", tdb_ReservaDotacao
    End If
End Sub

Private Sub tdb_ReservaDotacao_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mblnClickOk = True
End Sub

Private Sub tdb_ReservaDotacao_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    
    Dim strSQL As String
    Dim adoResultado As ADODB.Recordset
    
    With tdb_ReservaDotacao
        If (Not .EOF And Not .BOF) And mblnClickOk Then
            mblnClickOk = False
            mblnAlterando = True
            txtPKId = .Columns("PKId").Value
            
            LeReserva
            
            TrocaCorDeFundoObjeto True, cbo_intProgramaTrabalho
            TrocaCorDeFundoObjeto True, cbointProgramaTrabalho
            TrocaCorDeFundoObjeto True, cmd_ProgramaTrabalho
            TrocaCorDeFundoObjeto True, txtintNumero
            TrocaCorDeFundoObjeto True, txtDTMDATA
            TrocaCorDeFundoObjeto True, txtdblValor
            TrocaCorDeFundoObjeto True, txtstrSolicitacao
            TrocaCorDeFundoObjeto True, txtintExercicio
            
            If Not VerificaLancamentos(txtPKId) Then
                
                If Trim(txtstrSolicitacao) <> "" And Trim(txtintExercicio) <> "" Then
                    If IsNumeric(txtstrSolicitacao) And IsNumeric(txtintExercicio) Then
                        strSQL = ""
                        strSQL = strSQL & " SELECT intAutorizacaoDeCompra FROM " & gstrRequisicaoCompras
                        strSQL = strSQL & "  WHERE intReserva = " & txtPKId
                        strSQL = strSQL & "    AND intCodigo = " & txtstrSolicitacao
                        strSQL = strSQL & "    AND intExercicio = " & txtintExercicio
                        strSQL = strSQL & "    AND intAutorizacaoDeCompra IS NOT NULL "
                        
                        Set gobjBanco = New clsBanco
                        If gobjBanco.CriaADO(strSQL, 20, adoResultado) Then
                            If adoResultado.EOF And adoResultado.BOF Then
                                TrocaCorDeFundoObjeto False, txtstrSolicitacao
                                TrocaCorDeFundoObjeto False, txtintExercicio
                                TrocaCorDeFundoObjeto False, cbointProgramaTrabalho
                                TrocaCorDeFundoObjeto False, cmd_ProgramaTrabalho
                                TrocaCorDeFundoObjeto False, cbo_intProgramaTrabalho
                            End If
                        End If
                        'ORC1532
                        adoResultado.Close
                        Set adoResultado = Nothing
                        Set gobjBanco = Nothing
                        'ORC1532
                    End If
                Else
                    TrocaCorDeFundoObjeto False, txtstrSolicitacao
                    TrocaCorDeFundoObjeto False, txtintExercicio
                    TrocaCorDeFundoObjeto False, cbointProgramaTrabalho
                    TrocaCorDeFundoObjeto False, cmd_ProgramaTrabalho
                    TrocaCorDeFundoObjeto False, cbo_intProgramaTrabalho
                End If
            End If
            
            'TrocaCorDeFundoObjeto True, txtStrhistorico
            'TrocaCorDeFundoObjeto True, cbo_Historico
            TrocaCorDeFundoObjeto True, txtstrSolicitante
            'TrocaCorDeFundoObjeto True, cmd_Historico
            If mobjAux Is Nothing Then
                HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar
            Else
                HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrAplicar
            End If
            gCorLinhaSelecionada tdb_ReservaDotacao
            tab_3DReservar.TabEnabled(1) = True
            tab_3DReservar.TabEnabled(2) = True
            LimpaDataCancelamento
            
            LeTabelaEmpenho
            LeMovimentosCancelamentos
            
            mostraCamposCalculo False
            HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrCancelar
            'HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrSalvar
            'ORC1532
            '            If cbo_intProgramaTrabalho.Text = "" Then
            '                mblnClickOk = True
            '                tdb_ReservaDotacao_RowColChange 0, 0
            '            End If
            'ORC1532
        End If
    End With
    
End Sub

Private Sub mostraCamposCalculo(ByVal blnMostrar As Boolean)
    lbl_ValorProgramaTrabalho.Visible = blnMostrar
    txt_ValorProgramaTrabalho.Visible = blnMostrar
    lbl_Saldo.Visible = blnMostrar
    txt_Saldo.Visible = blnMostrar
    lbl_TotalEmpenho.Visible = blnMostrar
    txt_TotalEmpenho.Visible = blnMostrar
    txt_TotalBloqueado.Visible = blnMostrar
    lbl_TotalReserva.Visible = blnMostrar
    txt_TotalReserva.Visible = blnMostrar
End Sub


Private Function blnDadosOk() As Boolean
    Dim strSQL As String
    Dim adoResultado As ADODB.Recordset
    Dim dtmDtEncerramento As Date
    Dim intProgramaTemp As Long
    
    'NUMERO RESERVA
    If Len(Trim(txtintNumero)) = 0 Then
        ExibeMensagem "O Número da reserva tem que ser informado corretamente."
        If txtintNumero.Enabled Then txtintNumero.SetFocus
        Exit Function
    'SOLICITACAO DE COMPRAS
    'ORC1532
    '    ElseIf Trim(txtstrSolicitacao) = "" Then
    '        ExibeMensagem "A solicitação de compras deve ser informada."
    '        txtstrSolicitacao.SetFocus
    '        Exit Function
    'ORC1532
    'EXERCICIO
    ElseIf txtintExercicio = "" Then
        ExibeMensagem "Informe o exercìcio" '"Digite o exercício no campo que está entre os campos solicitação de compras e solicitante."
        txtintExercicio.SetFocus
        Exit Function
    'SOLICITANTE
    
    'DATA
    ElseIf gblnDataValida(txtDTMDATA) = False Then
        ExibeMensagem "Data da reserva tem que ser informada corretamente."
        If txtDTMDATA.Enabled Then txtDTMDATA.SetFocus
        Exit Function
    'DOTACAO
    
    'FUNCIONAL PROGRAMATICA
    ElseIf cbointProgramaTrabalho.ListIndex = -1 Then
        ExibeMensagem "A dotação tem que ser informada corretamente."
        If cbointProgramaTrabalho.Enabled Then cbointProgramaTrabalho.SetFocus
        Exit Function
    'VALOR
    ElseIf Val(gstrConvVrParaSql(txtdblValor)) <= 0 Then
        ExibeMensagem "O valor da reserva tem que ser informado corretamente."
        If txtdblValor.Enabled Then txtdblValor.SetFocus
        Exit Function
                
    ElseIf Val(gstrConvVrParaSql(txt_Saldo)) < Val(gstrConvVrParaSql(txtdblValor)) Then
        ExibeMensagem "O valor da reserva não pode ser superior ao valor da dotação."
        If txtdblValor.Enabled Then txtdblValor.SetFocus
        Exit Function
        
    ElseIf SaldoDotacaoAtual(gstrItemData(cbointProgramaTrabalho), Val(Month(CDate(txtDTMDATA))), gintExercicio, CDbl(txtdblValor)) = Empty Then
        If txtdblValor.Enabled Then txtdblValor.SetFocus
        Exit Function
        
    ElseIf Val(gstrConvVrParaSql(txtdblValor)) > Val(gstrConvVrParaSql(txt_Saldo)) Then
        ExibeMensagem "O valor da reserva não pode ser superior ao saldo da dotação."
        If cbointProgramaTrabalho.Enabled Then cbointProgramaTrabalho.SetFocus
        Exit Function
    ElseIf Trim(txtstrSolicitacao) <> "" Then
        If CarregaDadosSolicitacaoCompras(txtstrSolicitacao.Text, Val(txtintExercicio.Text), True) = False Then
            ExibeMensagem "Esta solicitação de compras é inválida."
            txtstrSolicitacao.SetFocus
            Exit Function
        End If
    ElseIf dataPedido > CDate(txtDTMDATA) Then
        ExibeMensagem "A data da reserva não pode ser menor que a data da solicitação de compras (" & gstrDataFormatada(dataPedido) & ")."
        txtDTMDATA.SetFocus
        Exit Function
    End If
    
    If mblnAlterando Then
        If Not (Trim(txtstrSolicitacao) = "" And Trim(txtintExercicio) = "") Then
            If cbointProgramaTrabalho.ListIndex > -1 Then
                intProgramaTemp = gstrItemData(cbointProgramaTrabalho)
            End If
            If CarregaDadosSolicitacaoCompras(Val(txtstrSolicitacao), txtintExercicio) = False And (txtstrSolicitacao <> strGuardaValorSolicitacao Or txtintExercicio <> strGuardaValorExercicio) Then
                ExibeMensagem "Solicitação inválida"
                Exit Function
            End If
            If Len(Trim(intProgramaTemp)) > 0 Then
                LeProgramaTrabalhoComReduzidoReservaDotacao cbo_intProgramaTrabalho, cbointProgramaTrabalho, gintExercicio
                cbointProgramaTrabalho.ListIndex = gintIndiceCBO(cbointProgramaTrabalho, intProgramaTemp)
            End If
        End If
        If VerificaLancamentos(txtPKId) Then
            If (txtstrSolicitacao <> strGuardaValorSolicitacao Or txtintExercicio <> strGuardaValorExercicio) Then
                ExibeMensagem "A Reserva de Dotação sofreu movimentações."
                Exit Function
            End If
            If gstrItemData(cbointProgramaTrabalho) <> strGuardaValorCboIntPrograma Then
                ExibeMensagem "A Reserva de Dotação sofreu movimentações."
                Exit Function
            End If
        End If
    End If
    
    cbointProgramaTrabalho_Click
    'If tab_3DReservar.Tab = 0 Then

    dtmDtEncerramento = VerificaDataEncerramento("EO", gintExercicio)
    
    If dtmDtEncerramento = Empty Then
        Exit Function
    Else
        If CDate(txtDTMDATA) <= dtmDtEncerramento Then
            ExibeMensagem "A data da reserva deve ser maior que a data de último encerramento (" & dtmDtEncerramento & ")."
            If txtDTMDATA.Enabled Then txtDTMDATA.SetFocus
            Exit Function
        End If
    End If
    
    If Year(CDate(txtDTMDATA)) <> CInt(gintExercicio) Then
        ExibeMensagem "A data da Reserva tem que estar no exercício de " & gintExercicio & "."
        If txtDTMDATA.Enabled Then txtDTMDATA.SetFocus
        Exit Function
    End If
   
    
    If Trim(txtstrSolicitacao) <> "" And Trim(txtintExercicio) <> "" Then
        If IsNumeric(txtstrSolicitacao) And IsNumeric(txtintExercicio) Then
            strSQL = ""
            strSQL = strSQL & " SELECT intRequisicaoComprasSituacoes,intReserva FROM " & gstrRequisicaoCompras
            strSQL = strSQL & "  WHERE intCodigo = " & txtstrSolicitacao
            strSQL = strSQL & "    AND intExercicio = " & txtintExercicio
            
            
            Set gobjBanco = New clsBanco
            If gobjBanco.CriaADO(strSQL, 20, adoResultado) Then
                If (Not adoResultado.EOF) Or (Not adoResultado.BOF) Then
                    If adoResultado!intrequisicaoComprasSituacoes <> 2 And adoResultado!intReserva <> Null Then
                        ExibeMensagem "A solicitação já está associada à uma autorização de compras."
                        If txtstrSolicitacao.Enabled Then txtstrSolicitacao.SetFocus
                        Exit Function
                    End If
                    
                End If
            End If
        End If
    End If
    
    
    
    blnDadosOk = True
    
    'End If
    
End Function

Private Function blnDadosCancelarOk() As Boolean
    Dim strValor          As String
    Dim strSaldo          As String
    Dim strData           As String
    Dim dtmDtEncerramento As Date
    strSaldo = tdb_ReservaDotacao.Columns("dblSaldo")
    strValor = tdb_ReservaDotacao.Columns("dblValor")
    strData = tdb_ReservaDotacao.Columns("dtmData")
    If gblnDataValida(txt_DataCancelamento) = False Then
        ExibeMensagem "A data do cancelamento está incorreta."
        txt_DataCancelamento.SetFocus
        Exit Function
    ElseIf CVDate(txt_DataCancelamento) < CVDate(strData) Then
        ExibeMensagem "A data do cancelamento não pode menor que a data da reserva."
        txt_DataCancelamento.SetFocus
        Exit Function
    ElseIf Val(gstrConvVrParaSql(txt_ValorCancelado)) <= 0 Then
        ExibeMensagem "O valor cancelado está incorreto."
        txt_ValorCancelado.SetFocus
        Exit Function
    ElseIf Val(gstrConvVrParaSql(txt_ValorCancelado)) > Val(gstrConvVrParaSql(strValor)) Then
        ExibeMensagem "O valor cancelado não pode ser superior ao valor reservado."
        txt_ValorCancelado.SetFocus
        Exit Function
    ElseIf Val(gstrConvVrParaSql(txt_ValorCancelado)) > Val(gstrConvVrParaSql(strSaldo)) Then
        ExibeMensagem "O valor cancelado não pode ser superior ao saldo da reserva."
        txt_ValorCancelado.SetFocus
        Exit Function
    End If
    
    
    'ORC677
    If Year(CDate(txt_DataCancelamento)) <> CInt(gintExercicio) Then
        ExibeMensagem "A data de cancelamento tem que estar no exercício de " & gintExercicio & "."
        If txt_DataCancelamento.Enabled Then txt_DataCancelamento.SetFocus
        Exit Function
    End If
    
    dtmDtEncerramento = VerificaDataEncerramento("EO", gintExercicio)
    
    If dtmDtEncerramento = Empty Then
        Exit Function
    Else
        If CDate(txt_DataCancelamento) <= dtmDtEncerramento Then
            ExibeMensagem "A data do cancelamento deve ser maior que a data de último encerramento (" & dtmDtEncerramento & ")."
            txt_DataCancelamento.SetFocus
            Exit Function
        End If
    End If
    
    blnDadosCancelarOk = True
    
End Function

Private Function strQueryCancela() As String
    Dim strSQL  As String
    strSQL = ""
    strSQL = strSQL & "SELECT DL.PKId, DL.intNumero, DL.dtmData, DL.dblValor, HI.PKId AS Historico FROM "
    strSQL = strSQL & gstrReservaDotacaoLiberada & " DL, " & gstrHistorico & " HI "
    strSQL = strSQL & "WHERE HI.strDescricao=DL.strHistorico AND intReservaDotacao = " & Val(txtPKId)
    strQueryCancela = strSQL
End Function

Private Sub CancelaReservaDotacao()
    
    '******************************************************************************************
    ' Data: 09/06/2003
    ' Alteração: - Substituição do comando nativo ISNULL() do SQL Server pela função gstrISNULL
    ' Responsável: Everton Bianchini
    '------------------------------------------------------------------------------------------
    ' Data: 09/06/2003
    ' Alteração: - Substituição da chamada direta à stored procedure pela função
    '            gstrStoredProcedure.
    ' Responsável: Everton Bianchini
    '******************************************************************************************
    
    
    Dim strSQL  As String
    If mblnAlterando = False Or tdb_Cancelado.ApproxCount = 0 Then
        If blnDadosCancelarOk Then
            If gblnExclusaoGravacaoOk("I", "Confirma cancelamento?", True) Then
                strSQL = ""
                strSQL = strSQL & "INSERT INTO " & gstrReservaDotacaoLiberada & " ("
                strSQL = strSQL & "intReservaDotacao, intNumero, dtmData, dblValor, "
                strSQL = strSQL & "strHistorico, dtmDtAtualizacao, lngCodUsr, intFlag"
                strSQL = strSQL & ") (SELECT "
                '            strSql = strSql & Val(txtPKId) & ", ISNULL(MAX(intNumero), 0) + 1, "
                strSQL = strSQL & Val(txtPKId) & ", " & gstrISNULL("MAX(intNumero)", "0") & " + 1, "
                strSQL = strSQL & gstrConvDtParaSql(txt_DataCancelamento) & ", "
                strSQL = strSQL & gstrConvVrParaSql(txt_ValorCancelado) & ", "
                strSQL = strSQL & "'" & txt_HistoricoCancelamento & "', "
                strSQL = strSQL & gstrConvDtParaSql(gstrDataDoSistema) & ", "
                strSQL = strSQL & glngCodUsr & ", 0 "
                strSQL = strSQL & "FROM " & gstrReservaDotacaoLiberada & " "
                strSQL = strSQL & "WHERE intReservaDotacao = " & Val(txtPKId) & ")"
                Set gobjBanco = New clsBanco
                If gobjBanco.Execute(strSQL) Then
                    
                    If Len(Trim$(txtstrSolicitacao)) > 0 And Len(Trim$(txtintExercicio)) > 0 Then
                        
                        strSQL = "UPDATE " & gstrRequisicaoCompras & " SET "
                        strSQL = strSQL & "intRequisicaoComprasSituacoes = (SELECT PKId FROM " & gstrRequisicaoComprasSituacoes & " WHERE bitTipo = 1), "
                        strSQL = strSQL & "intReserva = NULL "
                        strSQL = strSQL & "WHERE intCodigo = " & txtstrSolicitacao & " AND intExercicio = " & txtintExercicio
                        
                        gobjBanco.Execute strSQL
                        
                    End If
                    
                    GuardaValoresCancelamento True
                    
                    LeDaTabelaParaObj "", tdb_Cancelado, strQueryCancela
                    '                VerificaListaAutomatica "", tdb_ReservaDotacao, "sp_ReservaDotacao"
                    'VerificaListaAutomatica "", tdb_ReservaDotacao, gstrStoredProcedure("sp_ReservaDotacao", , True)
                    'LeTabelaReservaDotacao
                    'AtualizaGridCancelamentos
                    
                    LeMovimentosCancelamentos
                    'LimpaDataCancelamento
                    'MantemForm gstrNovo
                    
                    GuardaValoresCancelamento False
                    
                End If
            End If
        End If
    End If
End Sub

Private Sub GravaReservaDotacao()
    
    '******************************************************************************************
    ' Data: 09/06/2003
    ' Alteração: - Substituição do comando nativo ISNULL() do SQL Server pela função gstrISNULL
    ' Responsável: Everton Bianchini
    '------------------------------------------------------------------------------------------
    ' Data: 10/06/2003
    ' Alteração: - Adicionados os nomes dos atributos nos quais estavam sendo inseridos valores
    '            no comando INSERT INTO SELECT.
    ' Responsável: Everton Bianchini
    '******************************************************************************************
    
    Dim strSQL     As String
    Dim mstrCodigo As String
    Dim adoResultado As ADODB.Recordset
    
    'implementar com rotina de verificação de operacao
    
    If blnDadosOk Then
        If mblnAlterando = False Then
ProximoCodigo:
            
            If gblnExisteCodigo(2, gstrReservaDotacao, "intNumero", txtintNumero, gstrDATEPART(strYEAR, "dtmData"), "'" & Val(gintExercicio) & "'") Then
                strSQL = " Select (Max(intNumero) + 1) Codigo From " & gstrReservaDotacao
                strSQL = strSQL & " WHERE " & gstrDATEPART("yyyy", "dtmData") & " = " & gintExercicio
                
                Set gobjBanco = New clsBanco
                
                If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
                    mstrCodigo = adoResultado!Codigo
                End If
                'mstrCodigo = (gstrProximoCodigo(txtintnumero, gstrReservaDotacao, "intNumero", gintCodSeguranca, gstrDATEPART(strYEAR, "dtmData"), Val(gintExercicio), , True))
                If MsgBox("O número de reserva informado já se encontra cadastrado. Deseja usar o número " & mstrCodigo & "?", vbYesNo + vbQuestion) = vbNo Then
                    If txtintNumero.Enabled Then txtintNumero.SetFocus
                    Exit Sub
                Else
                    txtintNumero = mstrCodigo
                    GoTo ProximoCodigo
                End If
            End If
            
            If gblnExclusaoGravacaoOk("I", "Confirma gravação?", True) Then
                
                mstrNumero = mstrNumero & txtintNumero.Text & ","
                
                strSQL = ""
                strSQL = strSQL & txtintNumero & ", "
                strSQL = strSQL & gstrConvDtParaSql(txtDTMDATA) & ", "
                strSQL = strSQL & gstrConvVrParaSql(txtdblValor) & ", "
                strSQL = strSQL & gstrItemData(cbointProgramaTrabalho) & ", "
                strSQL = strSQL & "'" & Trim(txtstrHistorico) & "', "
                strSQL = strSQL & "'" & Trim(txtstrSolicitante) & "', "
                strSQL = strSQL & "'" & Trim(txtstrSolicitacao) & "', "
                strSQL = strSQL & gstrENulo(Trim(txtintExercicio), , True) & ", "
                strSQL = strSQL & gstrConvDtParaSql(gstrDataDoSistema) & ", "
                strSQL = strSQL & glngCodUsr
                
                strSQL = gstrStoredProcedure("sp_GeraReservaDotacao", strSQL, True)
                
                
                '                strSql = strSql & "INSERT INTO " & gstrReservaDotacao & " "
                '
                '                strSql = strSql & "(intNumero, dtmData, dblValor, intProgramaTrabalho, "
                '                strSql = strSql & "strHistorico, strSolicitante, strSolicitacao, intExercicio, "
                '                strSql = strSql & "dtmDtAtualizacao, lngCodUsr) "
                '
                ''                strSql = strSql & "SELECT ISNULL(MAX(intNumero), 0) + 1, "
                '                strSql = strSql & "VALUES (" & txtintNumero & ", "
                '                strSql = strSql & gstrConvDtParaSql(txtdtmData) & ", "
                '                strSql = strSql & gstrConvVrParaSql(txtdblValor) & ", "
                '                strSql = strSql & gstrItemData(cbointProgramaTrabalho) & ", "
                '                strSql = strSql & "'" & Trim(txtstrHistorico) & "', "
                '                strSql = strSql & "'" & Trim(txtstrSolicitante) & "', "
                '                strSql = strSql & "'" & Trim(txtstrSolicitacao) & "', "
                '                strSql = strSql & "'" & Trim(txtintExercicio) & "', "
                '                strSql = strSql & gstrConvDtParaSql(gstrDataDoSistema) & ", "
                '                strSql = strSql & glngCodUsr & " )"
                Set gobjBanco = New clsBanco
                
                
                'If gobjBanco.Execute(strSql) Then
                If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
                    
                    If adoResultado!resultado = 1 Then
                        If Len(Trim$(txtstrSolicitacao)) > 0 And Len(Trim$(txtintExercicio)) > 0 Then
                            
                            'If CDbl(txtdblValor) <= CDbl(txt_Saldo) Then
                            strSQL = "UPDATE " & gstrRequisicaoCompras & " SET "
                            strSQL = strSQL & "intRequisicaoComprasSituacoes = (SELECT PKId FROM " & gstrRequisicaoComprasSituacoes & " WHERE bitTipo = 3), "
                            strSQL = strSQL & "intReserva = (SELECT PKId FROM " & gstrReservaDotacao & " WHERE intNumero = " & txtintNumero & " AND INTEXERCICIORESERVA = " & gintExercicio & " ), "
                            strSQL = strSQL & "strRequisitante = '" & Trim(txtstrSolicitante) & "', "
                            strSQL = strSQL & "strObjetoCompra = '" & Trim(txtstrHistorico) & "', "
                            strSQL = strSQL & "intProgramaDeTrabalho = " & gstrItemData(cbointProgramaTrabalho) & " "
                            strSQL = strSQL & "WHERE intCodigo = " & txtstrSolicitacao & " AND intExercicio = " & txtintExercicio
                            
                            gobjBanco.Execute strSQL
                            '                        Else
                            '                            strSql = "UPDATE " & gstrRequisicaoCompras & " SET "
                            '                            strSql = strSql & "intRequisicaoComprasSituacoes = (SELECT PKId FROM " & gstrRequisicaoComprasSituacoes & " WHERE bitTipo = 1), "
                            '                            strSql = strSql & "intReserva = (SELECT PKId FROM " & gstrReservaDotacao & " WHERE intNumero = " & txtintNumero & ") "
                            '                            strSql = strSql & "strRequisitante = '" & Trim(txtstrSolicitante) & "', "
                            '                            strSql = strSql & "strObjetoCompra = '" & Trim(txtstrHistorico) & "', "
                            '                            strSql = strSql & "intProgramaDeTrabalho = " & gstrItemData(cbointProgramaTrabalho) & " "
                            '                            strSql = strSql & "WHERE intCodigo = " & txtstrSolicitacao & " AND intExercicio = " & txtintExercicio
                            
                            'End If
                            
                            
                            
                        End If
                        
                        If tdb_ReservaDotacao.Text <> "" Then
                            strFiltro = "  AND RD.intNumero IN (" & Mid(mstrNumero, 1, Len(mstrNumero) - 1) & ")  "
                        Else
                            strFiltro = "  AND RD.intNumero = " & txtintNumero.Text
                        End If
                        
                        LimpaDados
                        
                        'VerificaListaAutomatica "", tdb_ReservaDotacao, "sp_ReservaDotacao"
                        LeTabelaReservaDotacao strFiltro
                        'HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrSalvar
                        HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrCancelar
                    Else
                        GoTo ProximoCodigo
                    End If
                End If
            End If
        Else
            If gblnExclusaoGravacaoOk("A", "Confirma gravação?", True) Then
                strSQL = "UPDATE " & gstrReservaDotacao
                strSQL = strSQL & " SET dblValor = " & gstrConvVrParaSql(txtdblValor) & ", "
                strSQL = strSQL & " strHistorico = '" & Trim(txtstrHistorico) & "', "
                strSQL = strSQL & " strSolicitante = '" & Trim(txtstrSolicitante) & "', "
                strSQL = strSQL & " strSolicitacao = '" & Trim(txtstrSolicitacao) & "', "
                strSQL = strSQL & " intExercicio = '" & Trim(txtintExercicio) & "', "
                strSQL = strSQL & " dtmDtAtualizacao = " & gstrConvDtParaSql(gstrDataDoSistema) & ", "
                strSQL = strSQL & " lngCodUsr = " & glngCodUsr
                strSQL = strSQL & " WHERE PKID=" & txtPKId
                
                Set gobjBanco = New clsBanco
                
                gobjBanco.Execute strSQL
                
                strSQL = ""
                strSQL = strSQL & "UPDATE " & gstrRequisicaoCompras
                strSQL = strSQL & " SET intReserva = Null, "
                strSQL = strSQL & " intRequisicaoComprasSituacoes=(SELECT PKId FROM " & gstrRequisicaoComprasSituacoes & " WHERE bitTipo = 1)"
                strSQL = strSQL & " WHERE intReserva = " & txtPKId
                
                
                
                Set gobjBanco = New clsBanco
                
                If gobjBanco.Execute(strSQL) Then
                    If txtstrSolicitacao <> "" And txtintExercicio <> "" Then
                        strSQL = ""
                        strSQL = strSQL & " Select * FROM " & gstrRequisicaoCompras
                        strSQL = strSQL & " WHERE intReserva =" & txtPKId
                        strSQL = strSQL & " AND intCodigo =" & txtstrSolicitacao
                        strSQL = strSQL & " AND intExercicio=" & txtintExercicio
                        
                        If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
                            If adoResultado.EOF And adoResultado.BOF Then
                                
                                strSQL = "UPDATE " & gstrRequisicaoCompras & " SET "
                                strSQL = strSQL & "intRequisicaoComprasSituacoes = (SELECT PKId FROM " & gstrRequisicaoComprasSituacoes & " WHERE bitTipo = 3), "
                                strSQL = strSQL & "intReserva = (SELECT PKId FROM " & gstrReservaDotacao & " WHERE intNumero = " & txtintNumero & " AND INTEXERCICIORESERVA = " & gintExercicio & " ), "
                                strSQL = strSQL & "strRequisitante = '" & Trim(txtstrSolicitante) & "', "
                                strSQL = strSQL & "strObjetoCompra = '" & Trim(txtstrHistorico) & "', "
                                strSQL = strSQL & "intProgramaDeTrabalho = " & gstrItemData(cbointProgramaTrabalho) & " "
                                strSQL = strSQL & "WHERE intCodigo = " & txtstrSolicitacao & " AND intExercicio = " & txtintExercicio
                                
                                gobjBanco.Execute strSQL
                            End If
                        End If
                    End If
                    
                End If
            End If
            
        End If
    End If
End Sub

Private Sub LimpaDataCancelamento()
    txt_DataCancelamento = ""
    txt_ValorCancelado = ""
    txt_HistoricoCancelamento = ""
    cbo_HistoricoCancelamento.ListIndex = -1
    'tdb_Cancelado.DataSource = Nothing
End Sub

Private Sub LimpaDados()
    txtPKId.Text = ""
    cbo_intProgramaTrabalho.ListIndex = -1
    cbointProgramaTrabalho.ListIndex = -1
    txt_ValorProgramaTrabalho = ""
    txt_TotalBloqueado = ""
    txt_Saldo = ""
    'txt_Orgao = ""
    'txt_UnidadeOrcamentaria = ""
    'txt_Subunidade = ""
    'txt_TipoCredito = ""
    'txt_Funcao = ""
    'txt_Subfuncao = ""
    'txt_Programa = ""
    'txt_SubPrograma = ""
    'txt_Projetoatividade = ""
    'txt_ElementoDespesa = ""
    LimpaDataCancelamento
    tdb_Cancelado.DataSource = Nothing
    txtintNumero = ""
    txtDTMDATA = ""
    txtdblValor = ""
    txtstrSolicitacao = ""
    txtintExercicio = ""
    txtstrHistorico = ""
    'cbo_Historico.ListIndex = -1
    'TrocaCorDeFundoObjeto False, cmd_Historico
    TrocaCorDeFundoObjeto False, cmd_ProgramaTrabalho
    txtstrSolicitante = ""
    tab_3DReservar.TabEnabled(1) = False
    tab_3DReservar.TabEnabled(2) = False
    tdb_Empenho.DataSource = Nothing
    
    cbo_intProgramaTrabalho.Enabled = True
    cbointProgramaTrabalho.Enabled = True
    cbo_intProgramaTrabalho.ListIndex = -1
    cbointProgramaTrabalho.ListIndex = -1
    txtintNumero.Enabled = True
    txtDTMDATA.Enabled = True
    txtdblValor.Enabled = True
    txtstrSolicitacao.Enabled = True
    txtintExercicio.Enabled = True
    txtstrHistorico.Enabled = True
    cbo_Historico.Enabled = True
    txtstrSolicitante.Enabled = True
    TrocaCorDeFundoObjeto False, txtintNumero
    TrocaCorDeFundoObjeto False, txtDTMDATA
    TrocaCorDeFundoObjeto False, txtdblValor
    TrocaCorDeFundoObjeto False, txtstrSolicitacao
    TrocaCorDeFundoObjeto False, txtintExercicio
    TrocaCorDeFundoObjeto False, txtstrHistorico
    TrocaCorDeFundoObjeto False, cbo_Historico
    TrocaCorDeFundoObjeto False, txtstrSolicitante
    TrocaCorDeFundoObjeto False, cbo_intProgramaTrabalho
    TrocaCorDeFundoObjeto False, cbointProgramaTrabalho
    txt_TotalReserva = ""
    txt_TotalEmpenho = ""
    TrocaCorObjeto txtdblValor, False
    mostraCamposCalculo True
    'cbo_intProgramaTrabalho.SetFocus
    mblnAlterando = False
    'HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrSalvar
    HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrCancelar
    If txtintNumero.Enabled Then txtintNumero.SetFocus
    
    
End Sub

Public Sub MantemForm(ByVal strModoOperacao As String)
    
    Select Case UCase(strModoOperacao)
    Case gstrNovo
        mblnAlterando = False
        If tab_3DReservar.Tab = 1 Then
            If txt_DataCancelamento.Text <> "" And txt_ValorCancelado.Text <> "" Then
                LimpaDataCancelamento
            Else
                LimpaDados
            End If
            Exit Sub
        ElseIf tab_3DReservar.Tab = 2 Then
            LimpaDados
            Exit Sub
        End If
        LimpaDados
    Case gstrCancelar
        CancelaReservaDotacao
        DoEvents
        'LimpaDataCancelamento
        Exit Sub
    Case gstrSalvar
        blnSalvando = True
        'If mblnAlterando Then Exit Sub
        GravaReservaDotacao
        blnSalvando = False
        Exit Sub
    Case gstrPreencherLista
        
        If Me.ActiveControl.Name = cbo_intProgramaTrabalho.Name Or Me.ActiveControl.Name = cbointProgramaTrabalho.Name Then
            'LeProgramaTrabalhoComReduzido cbo_intProgramaTrabalho, cbointProgramaTrabalho, gintExercicio, "SELECT PKId, intCodigoReduzido, strCodigo FROM " & gstrProgramaDeTrabalho & " WHERE intExercicio=" & gintExercicio & " AND " & gstrCONVERT(CDT_VARCHAR, "intCodigoReduzido") & " LIKE '" & cbo_intProgramaTrabalho.Text & "%' ORDER BY strCodigo"
            LeProgramaTrabalhoComReduzidoReservaDotacao cbo_intProgramaTrabalho, cbointProgramaTrabalho, CInt(gintExercicio)  ' Year(!dtmDataRequisicao)
            
        ''ElseIf Me.ActiveControl.Name = cbointProgramaTrabalho.Name Then
            'LeProgramaTrabalhoComReduzido cbo_intProgramaTrabalho, cbointProgramaTrabalho, gintExercicio, "SELECT PKId, intCodigoReduzido, strCodigo FROM " & gstrProgramaDeTrabalho & " WHERE intExercicio=" & gintExercicio & " AND " & gstrCONVERT(CDT_VARCHAR, "strCodigo") & " LIKE '" & cbointProgramaTrabalho.Text & "%' ORDER BY strCodigo"
        ''    LeProgramaTrabalhoComReduzidoReservaDotacao cbo_intProgramaTrabalho, cbointProgramaTrabalho, CInt(gintExercicio) 'Year(!dtmDataRequisicao)
            
        ElseIf Me.ActiveControl.Name = cbo_Historico.Name Then
            LeDaTabelaParaObj gstrHistorico, cbo_Historico
            
        ElseIf Me.ActiveControl.Name = cbo_HistoricoCancelamento.Name Then
            LeDaTabelaParaObj gstrHistorico, cbo_HistoricoCancelamento
        End If
        
    Case gstrLocalizar
        strFiltro = ""
        ToolBarGeral strModoOperacao, gstrReservaDotacao, mblnAlterando, tdb_ReservaDotacao, Me, mobjAux, , strQueryDotacao
    Case gstrAplicar
        ToolBarGeral strModoOperacao, gstrReservaDotacao, mblnAlterando, tdb_ReservaDotacao, Me, mobjAux, , strQueryAplicar
    Case gstrImprimir
        If Trim(txtPKId) = "" Then
            ExibeMensagem "É necessário selecionar uma reserva para impressão."
            Exit Sub
        End If
    End Select
    ToolBarGeral strModoOperacao, gstrReservaDotacao, _
    mblnAlterando, tdb_ReservaDotacao, Me, _
    mobjAux, strQueryDotacao, , rptReservaDotacao, _
    strQueryRelatorio
End Sub

Private Sub Form_Load()
    
    '******************************************************************************************
    ' Data: 09/06/2003
    ' Alteração: - Substituição da chamada direta à stored procedure pela função
    '            gstrStoredProcedure.
    ' Responsável: Everton Bianchini
    '******************************************************************************************
    
    mblnAlterando = False
    blnSalvando = False
    tab_3DReservar.TabEnabled(1) = False
    tab_3DReservar.TabEnabled(2) = False
    tab_3DReservar.TabVisible(0) = False
    VerificaObjParaAplicar mobjAux
    '    VerificaListaAutomatica gstrReservaDotacao, tdb_ReservaDotacao, strQueryDotacao  'gstrStoredProcedure("sp_ReservaDotacao", , True)
    
    
    gstrProximoCodigo txtintNumero, gstrReservaDotacao, "intNumero", gintCodSeguranca, gstrDATEPART(strYEAR, "dtmData"), Val(gintExercicio)
    
    'MarcaCampo txtintNumero
End Sub

Private Sub txt_DataCancelamento_GotFocus()
    MarcaCampo txt_DataCancelamento
    'AtivaPastaDeObjeto tab_3DReservar, 1
End Sub

Private Sub txt_DataCancelamento_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "D", txt_DataCancelamento
End Sub

Private Sub txt_DataCancelamento_LostFocus()
    txt_DataCancelamento = gstrDataFormatada(txt_DataCancelamento)
    
    'ORC677
    If IsDate(txt_DataCancelamento) Then
        If Year(CDate(txt_DataCancelamento)) <> CInt(gintExercicio) Then
            ExibeMensagem "A data de cancelamento tem que estar no exercício de " & gintExercicio & "."
            If txt_DataCancelamento.Enabled Then txt_DataCancelamento.SetFocus
            Exit Sub
        End If
    End If
    
End Sub

Private Sub txt_HistoricoCancelamento_GotFocus()
    MarcaCampo txt_HistoricoCancelamento
End Sub

Private Sub txt_HistoricoCancelamento_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub txt_ValorCancelado_GotFocus()
    MarcaCampo txt_ValorCancelado
    AtivaPastaDeObjeto tab_3DReservar, 1
End Sub

Private Sub txt_ValorCancelado_KeyPress(KeyAscii As Integer)
    'CaracterValido KeyAscii, "V", txt_ValorCancelado
    gstrLimitaCampoValor txt_ValorCancelado, KeyAscii, 14, 2 'ORC1532
End Sub

Private Sub txt_ValorCancelado_LostFocus()
    txt_ValorCancelado = gstrConvVrDoSql(txt_ValorCancelado)
End Sub

Private Sub txtdtmData_GotFocus()
    'AtivaPastaDeObjeto tab_3DReservar, 0
    If Trim(txtDTMDATA) = "" Then
        txtDTMDATA = DataAutomatica 'RetornaSugestaoData
    End If
    MarcaCampo txtDTMDATA
End Sub

Private Sub txtdtmData_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "D", txtDTMDATA
End Sub

Private Sub txtdtmData_LostFocus()
    txtDTMDATA = gstrDataFormatada(txtDTMDATA)
    
    'ORC677
    If IsDate(txtDTMDATA) Then
        If Year(CDate(txtDTMDATA)) <> CInt(gintExercicio) Then
            ExibeMensagem "A data tem que estar no exercício de " & gintExercicio & "."
            If txtDTMDATA.Enabled Then txtDTMDATA.SetFocus
            Exit Sub
        End If
    End If
    
    cbointProgramaTrabalho_Click
End Sub

Private Sub txtintExercicio_GotFocus()
    MarcaCampo txtintExercicio
End Sub

Private Sub txtintExercicio_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtintExercicio
End Sub

Private Sub txtintExercicio_LostFocus()
    If Len(txtstrSolicitacao) > 0 And Len(txtintExercicio) > 0 Then
        CarregaDadosSolicitacaoCompras txtstrSolicitacao, txtintExercicio
    Else
        If Not mblnAlterando Then
            LimpaDadosSolicitacaoCompras
        End If
    End If
End Sub

Private Sub txtintNumero_GotFocus()
    txtintNumero = gstrProximoCodigo(txtintNumero, gstrReservaDotacao, "intNumero", gintCodSeguranca, "intExercicioReserva", Val(gintExercicio), , True, , , "intExercicioReserva", Val(gintExercicio))
    'gstrProximoCodigo txtintNumero, gstrReservaDotacao, "intNumero", gintCodSeguranca, gstrDATEPART(strYEAR, "dtmData"), Val(gintExercicio) ', , True, , , gstrDATEPART(strYEAR, "dtmData"), Val(gintExercicio)
    If Trim(txtDTMDATA) = "" Then
        txtDTMDATA = DataAutomatica 'RetornaSugestaoData
    End If
    MarcaCampo txtintNumero
End Sub

Private Sub txtintNumero_KeyPress(KeyAscii As Integer)
    'CaracterValido KeyAscii
    gstrLimitaCampoValor txtintNumero, KeyAscii, 9, 0
End Sub

Private Sub txtstrHistorico_GotFocus()
    MarcaCampo txtstrHistorico
End Sub

Private Sub txtstrHistorico_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub txtstrSolicitacao_LostFocus()
    If Len(txtstrSolicitacao) > 0 And Len(txtintExercicio) > 0 Then
        CarregaDadosSolicitacaoCompras txtstrSolicitacao, txtintExercicio
    Else
        If Not mblnAlterando Then
            LimpaDadosSolicitacaoCompras
        End If
    End If
End Sub

Private Sub txtstrSolicitante_GotFocus()
    MarcaCampo txtstrSolicitante
End Sub

Private Sub txtstrSolicitante_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrSolicitante
End Sub

Private Sub txtdblValor_GotFocus()
    MarcaCampo txtdblValor
    'AtivaPastaDeObjeto tab_3DReservar, 0
End Sub

Private Sub txtdblValor_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "V", txtdblValor
End Sub

Private Sub txtdblValor_LostFocus()
    txtdblValor = gstrConvVrDoSql(txtdblValor)
End Sub

Private Sub txtstrSolicitacao_GotFocus()
    MarcaCampo txtstrSolicitacao
End Sub

Private Sub txtstrSolicitacao_KeyPress(KeyAscii As Integer)
    'CaracterValido KeyAscii, "N", txtstrSolicitacao
    gstrLimitaCampoValor txtstrSolicitacao, KeyAscii, 9, 0
End Sub

Private Sub LeReserva()
    Dim strSQL  As String
    Dim adoResultado As ADODB.Recordset
    Dim i As Integer

    strSQL = ""
    strSQL = strSQL & "SELECT RD.intNumero, RD.dtmData, RD.dblValor, RD.intProgramaTrabalho, "
    strSQL = strSQL & "RD.strHistorico, RD.strSolicitante, RD.strSolicitacao, RD.intExercicio "
    strSQL = strSQL & "FROM "
    strSQL = strSQL & gstrReservaDotacao & " RD "
    strSQL = strSQL & "WHERE RD.PKId = " & Val(tdb_ReservaDotacao.Columns(0))
    
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
        With adoResultado
            If .EOF = False Then
                
                txtintNumero = !INTNUMERO
                txtDTMDATA = gstrDataFormatada(!DTMDATA)
                txtdblValor = gstrConvVrDoSql(!dblValor)
                txtstrSolicitacao = gstrENulo(!strSolicitacao)
                
                If Trim(gstrENulo(!intExercicio)) = "0" Then
                    txtintExercicio = ""
                Else
                    txtintExercicio = gstrENulo(!intExercicio)
                End If
                
                If Not blnSalvando Then
                    txtstrHistorico = gstrENulo(!STRHISTORICO)
                End If
                
                txtstrSolicitante = gstrENulo(!strSolicitante)
                LeProgramaTrabalhoComReduzidoReservaDotacao cbo_intProgramaTrabalho, cbointProgramaTrabalho, gintExercicio
                cbointProgramaTrabalho.ListIndex = gintIndiceCBO(cbointProgramaTrabalho, !intProgramaTrabalho, False)
                intGuardaValorCbo_IntPrograma = cbo_intProgramaTrabalho.ListIndex
                strGuardaValorCboIntPrograma = gstrENulo(!intProgramaTrabalho)
                
                If Trim(gstrENulo(!intExercicio)) = "0" Then
                    strGuardaValorExercicio = ""
                Else
                    strGuardaValorExercicio = gstrENulo(!intExercicio)
                End If
                
                If Trim(gstrENulo(!strSolicitacao)) = "0" Then
                    strGuardaValorSolicitacao = ""
                Else
                    strGuardaValorSolicitacao = gstrENulo(!strSolicitacao)
                End If
                
            End If
        End With
    End If
    LeDaTabelaParaObj "", tdb_Cancelado, strQueryCancela
    'HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrSalvar
    HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrCancelar
End Sub

'Private Sub LeValoresbyDotacao(ByRef DotacaoAtual As Object, ByRef SaldoDotacao As Object, ByRef totalEmpenhadoAno As Object)
'    Dim strSQL  As String
'    Dim adoResultado As ADODB.Recordset
'
'
'    If cbo_intProgramaTrabalho.ListIndex = -1 And Trim(txtdtmData) = "" Then
'
'        Exit Sub
'    End If
'
'    strSQL = ""
'    strSQL = strSQL & " SELECT saldoIni, SUM(Empenhado) Empenhado , "
'    strSQL = strSQL & " SUM(Suplementado) Suplementado,SUM(Anulado) Anulado"
'    strSQL = strSQL & " FROM " & gstrContaValoresAcumulados & " "
'    strSQL = strSQL & " WHERE intProgramaTrabalho = " & gstrItemData(cbo_intProgramaTrabalho)
'    strSQL = strSQL & " AND Exercicio = " & gintExercicio
'    strSQL = strSQL & " AND Mes <= " & Month(CDate(DTMDATA))
'    strSQL = strSQL & "GROUP BY saldoIni"
'
'    Set gobjBanco = New clsBanco
'
'
'
'   'objSaldo = gstrConvVrDoSql(SaldoDotacaoAtual(gstrItemData(cboPrograma), Val(Month(CDate(strData))), gintExercicio))
'
'
'
'    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
'        With adoResultado
'            If .EOF = False Then
'                txt_ValorProgramaTrabalho = (!saldoIni + !suplementado) - !anulado
'                txt_Saldo = CDbl(txt_ValorProgramaTrabalho)
'                txt_TotalEmpenho
'
'           End If
'        End With
'    End If
'End Sub
'


Private Function LeDataUltimaReserva() As Date
    
    Dim strSQL  As String
    Dim adoResultado As ADODB.Recordset
    
    strSQL = ""
    strSQL = strSQL & " SELECT dtmData FROM "
    strSQL = strSQL & gstrReservaDotacao
    strSQL = strSQL & " WHERE PKID = (SELECT MAX(PKID) FROM "
    strSQL = strSQL & gstrReservaDotacao
    strSQL = strSQL & " WHERE intExercicio =" & gintExercicio & " )"
    
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
        With adoResultado
            If .EOF = False Then
                LeDataUltimaReserva = !DTMDATA
            Else
                LeDataUltimaReserva = Empty
            End If
        End With
    End If
End Function

Private Sub LeCancelamento()
    Dim intCont As Integer
    txt_DataCancelamento = tdb_Cancelado.Columns("Data")
    txt_ValorCancelado = tdb_Cancelado.Columns("Valor")
    
    
    If cbo_HistoricoCancelamento.ListCount = 0 Then
        LeDaTabelaParaObj gstrHistorico, cbo_HistoricoCancelamento
    End If
    
    'cbo_HistoricoCancelamento.Text = tdb_Cancelado.Columns("Historico")
    'cbo_HistoricoCancelamento.ListIndex = gintIndiceCBO(cbo_HistoricoCancelamento, gstrItemData(cbo_intProgramaTrabalho))
    
    With cbo_HistoricoCancelamento
        For intCont = 0 To .ListCount - 1
            If .list(intCont) = tdb_Cancelado.Columns("Historico") Then .ListIndex = intCont
        Next
    End With
    
    txt_HistoricoCancelamento.Text = tdb_Cancelado.Columns("Historico")
End Sub

Public Function strQuery()
    Dim strSQL  As String
    strSQL = ""
    strSQL = strSQL & "SELECT RD.PKId, RD.intNumero, "
    strSQL = strSQL & "RD.dtmData, RD.dblValor, RD.strSolicitante "
    strSQL = strSQL & "FROM "
    strSQL = strSQL & gstrReservaDotacao & " RD "
    strSQL = strSQL & "ORDER BY RD.intNumero"
    strQuery = strSQL
End Function

Private Function strQueryRelatorio() As String
    
    '******************************************************************************************
    ' Data: 09/06/2003
    ' Alteração: - Substituição do comando nativo de outer join (*= ou =*) do SQL Server pela
    '            variável strOUTJSQLServer e inclusão do comando de outer join ((+)) do Oracle,
    '            representado pela variável strOUTJOracle.
    ' Responsável: Everton Bianchini
    '******************************************************************************************
    
    Dim strSQL As String
    strSQL = ""
    strSQL = strSQL & " SELECT "
    strSQL = strSQL & " A.PKId, A.strHistorico, C.intCodigoReduzido, R.strCodigo AS intProgramaTrabalho, A.intNumero AS NumeroDaReserva,"
    strSQL = strSQL & " A.dtmData AS DataDaReserva, A.dblValor AS ValorReservado, B.intReservaDotacao AS PKIdReserva,"
    strSQL = strSQL & " B.intNumero AS NumeroDoCancelamento, B.dtmData AS DataDoCancelamento, B.dblValor AS ValorDoCancelamento,"
    strSQL = strSQL & "F.strDescricao AS strFuncao,S.strDescricao AS strSubFuncao, P.strDescricao AS strProjeto, R.strdescricao AS strPrograma,"
    strSQL = strSQL & " U.strCodigo AS intUnidadeOrcamentaria,"
    If bytDBType = SQLServer Then
        strSQL = strSQL & " SUBSTRING(E.strCodigoElementoDespesa,1,6) AS intElementoDespesa,"
    Else
        strSQL = strSQL & " LPAD(E.strCodigoElementoDespesa,6,1) AS intElementoDespesa,"
    End If
    strSQL = strSQL & " F.strCodigo AS intfuncao, S.strCodigo AS intSubFuncao, P.strCodigo AS intProjetoAtividade,"
    strSQL = strSQL & " E.strDescricao AS strElementoDespesa, U.strDescricao AS strUnidadeOrcamentaria"
    
    strSQL = strSQL & " FROM "
    strSQL = strSQL & gstrReservaDotacao & " A, "
    strSQL = strSQL & gstrReservaDotacaoLiberada & " B, "
    strSQL = strSQL & gstrProgramaDeTrabalho & " C, "
    strSQL = strSQL & gstrElementoDespesa & " E, "
    strSQL = strSQL & gstrFuncaoDoGoverno & " F, "
    strSQL = strSQL & gstrProjeto & " P, "
    strSQL = strSQL & gstrPrograma & " R, "
    strSQL = strSQL & gstrSubFuncaoGoverno & " S, "
    strSQL = strSQL & gstrUnidadeOrcamentaria & " U "
    
    strSQL = strSQL & " WHERE "
    strSQL = strSQL & " A.intProgramaTrabalho = C.PKId "
    '   strSql = strSql & " AND B.intReservaDotacao =* A.PKId "
    strSQL = strSQL & " AND B.intReservaDotacao " & strOUTJOracle & "=" & strOUTJSQLServer & " A.PKId"
    strSQL = strSQL & " AND C.intExercicio = " & gintExercicio
    strSQL = strSQL & " AND A.pkid = " & txtPKId
    strSQL = strSQL & " AND E.Pkid = C.intElementoDespesa"
    strSQL = strSQL & " AND F.Pkid = C.intFuncao"
    strSQL = strSQL & " AND P.pkid = C.intProjetoAtividade"
    strSQL = strSQL & " AND R.Pkid = C.intPrograma"
    strSQL = strSQL & " AND S.Pkid = C.intSubFuncao"
    strSQL = strSQL & " AND U.Pkid = C.intUnidadeOrcamentaria"
    
    strSQL = strSQL & " ORDER BY "
    strSQL = strSQL & " A.intProgramaTrabalho, A.PKId "
    strQueryRelatorio = strSQL
End Function
Private Function strQueryDotacao(Optional PkidReserva As Long) As String
    
    '******************************************************************************************
    ' Data: 09/06/2003s
    ' Alteração: - Substituição do comando nativo ISNULL() do SQL Server pela função gstrISNULL
    ' Responsável: Everton Bianchini
    '------------------------------------------------------------------------------------------
    ' Data: 09/06/2003
    ' Alteração: - Substituição do comando nativo de outer join (*= ou =*) do SQL Server pela
    '            variável strOUTJSQLServer e inclusão do comando de outer join ((+)) do Oracle,
    '            representado pela variável strOUTJOracle.
    ' Responsável: Everton Bianchini
    '******************************************************************************************
    
    Dim strSQL As String
    'strSql = strSql & "SELECT RD.PKId, RD.intNumero, RD.dtmData, RD.dblValor, RD.strSolicitante, ISNULL(SUM(RDL.dblValor),0) AS dblCanceledo, RD.dblValor - ISNULL(SUM(RDL.dblValor),0) AS dblSaldo "
    
    'alterado em 18/06/2004 09:45:
    'Troca dos nomes explícitos das tabelas pelas constantes usadas no sistema
    'por Wagner
    strSQL = strSQL & "SELECT RD.PKId, RD.intNumero, RD.dtmData, RD.dblValor, RD.strSolicitante, "
    strSQL = strSQL & "(SELECT " & gstrISNULL("SUM(RDL.dblValor)", "0") & " FROM " & gstrReservaDotacaoLiberada & " RDL WHERE RDL.intFlag = 0 AND RD.PKId " & strOUTJSQLServer & "= RDL.intReservaDotacao " & strOUTJOracle & ") AS dblCanceledo,"
    strSQL = strSQL & "((SELECT " & gstrISNULL("SUM(RDL.dblValor)", "0") & " FROM " & gstrReservaDotacaoLiberada & "  RDL WHERE RDL.intFlag = 1 AND RDL.dblValor >=0 AND RD.PKId " & strOUTJSQLServer & "= RDL.intReservaDotacao " & strOUTJOracle & ")"
    strSQL = strSQL & "+ (SELECT " & gstrISNULL("SUM(RDL.dblValor)", "0") & " FROM " & gstrReservaDotacaoLiberada & " RDL WHERE RDL.intFlag = 1 AND RDL.dblValor < 0 AND RD.PKId " & strOUTJSQLServer & "= RDL.intReservaDotacao " & strOUTJOracle & ")) AS dblEmpenhado,"
    strSQL = strSQL & " RD.dblValor - "
    strSQL = strSQL & "(SELECT " & gstrISNULL("SUM(RDL.dblValor)", "0") & " FROM " & gstrReservaDotacaoLiberada & " RDL WHERE RDL.intFlag = 0 AND RD.PKId " & strOUTJSQLServer & "= RDL.intReservaDotacao " & strOUTJOracle & ") "
    strSQL = strSQL & " - ((SELECT " & gstrISNULL("SUM(RDL.dblValor)", "0") & " FROM " & gstrReservaDotacaoLiberada & " RDL WHERE RDL.intFlag = 1 AND RDL.dblValor >=0 AND RD.PKId " & strOUTJSQLServer & "= RDL.intReservaDotacao " & strOUTJOracle & ")"
    strSQL = strSQL & "+ (SELECT " & gstrISNULL("SUM(RDL.dblValor)", "0") & " FROM " & gstrReservaDotacaoLiberada & " RDL WHERE RDL.intFlag = 1 AND RDL.dblValor < 0 AND RD.PKId " & strOUTJSQLServer & "= RDL.intReservaDotacao " & strOUTJOracle & "))"
    strSQL = strSQL & " AS dblSaldo "
    strSQL = strSQL & "FROM " & gstrReservaDotacao & " RD, " & gstrReservaDotacaoLiberada & " RDL "
    '    strSql = strSql & "WHERE RD.PKId*=RDL.intReservaDotacao "
    
    strSQL = strSQL & "WHERE RD.PKId " & strOUTJSQLServer & "= RDL.intReservaDotacao " & strOUTJOracle & " AND "
    strSQL = strSQL & gstrDATEPART("YYYY", "RD.dtmData") & " = " & gintExercicio & " "
    
    If cbo_intProgramaTrabalho.ListIndex <> -1 Then
        strSQL = strSQL & "AND RD.intProgramaTrabalho =" & gstrItemData(cbo_intProgramaTrabalho)
    End If
    
    If Val(PkidReserva) > 0 Then
        strSQL = strSQL & "AND RD.Pkid =" & PkidReserva
    End If
    
    strSQL = strSQL & " GROUP BY RD.PKId, RD.intNumero, RD.dtmData, RD.dblValor, RD.strSolicitante "
    strSQL = strSQL & "ORDER BY RD.PKId"
    strQueryDotacao = strSQL
    
End Function

Private Sub LeTabelaReservaDotacao(Optional strFiltro As String)
    Dim strSQL          As String
    Dim adoResultado    As ADODB.Recordset
    strSQL = ""
    'strSql = strSql & "SELECT RD.PKId, RD.intNumero, RD.dtmData, RD.dblValor, RD.strSolicitante, ISNULL(SUM(RDL.dblValor),0) AS dblCanceledo, RD.dblValor - ISNULL(SUM(RDL.dblValor),0) AS dblSaldo "
    strSQL = strSQL & "SELECT RD.PKId, RD.intNumero, RD.dtmData, RD.dblValor, RD.strSolicitante, "
    strSQL = strSQL & "(SELECT " & gstrISNULL("SUM(RDL.dblValor)", "0") & " FROM  " & gstrReservaDotacaoLiberada & " RDL WHERE RDL.intFlag = 0 AND RD.PKId " & strOUTJSQLServer & "= RDL.intReservaDotacao " & strOUTJOracle & ") AS dblCanceledo,"
    strSQL = strSQL & "((SELECT " & gstrISNULL("SUM(RDL.dblValor)", "0") & " FROM " & gstrReservaDotacaoLiberada & " RDL WHERE RDL.intFlag = 1 AND RDL.dblValor >=0 AND RD.PKId " & strOUTJSQLServer & "= RDL.intReservaDotacao " & strOUTJOracle & ")"
    strSQL = strSQL & "+ (SELECT " & gstrISNULL("SUM(RDL.dblValor)", "0") & " FROM " & gstrReservaDotacaoLiberada & " RDL WHERE RDL.intFlag = 1 AND RDL.dblValor < 0 AND RD.PKId " & strOUTJSQLServer & "= RDL.intReservaDotacao " & strOUTJOracle & ")) AS dblEmpenhado,"
    strSQL = strSQL & " RD.dblValor - "
    strSQL = strSQL & " (SELECT " & gstrISNULL("SUM(RDL.dblValor)", "0") & " FROM " & gstrReservaDotacaoLiberada & " RDL WHERE RDL.intFlag = 0 AND RD.PKId " & strOUTJSQLServer & "= RDL.intReservaDotacao " & strOUTJOracle & ")"
    strSQL = strSQL & " - ((SELECT " & gstrISNULL("SUM(RDL.dblValor)", "0") & " FROM " & gstrReservaDotacaoLiberada & " RDL WHERE RDL.intFlag = 1 AND RDL.dblValor >=0 AND RD.PKId " & strOUTJSQLServer & "= RDL.intReservaDotacao " & strOUTJOracle & ")"
    strSQL = strSQL & "+ (SELECT " & gstrISNULL("SUM(RDL.dblValor)", "0") & " FROM " & gstrReservaDotacaoLiberada & " RDL WHERE RDL.intFlag = 1 AND RDL.dblValor < 0 AND RD.PKId " & strOUTJSQLServer & "= RDL.intReservaDotacao " & strOUTJOracle & " ))"
    strSQL = strSQL & " AS dblSaldo "
    strSQL = strSQL & " FROM " & gstrReservaDotacao & " RD, " & gstrReservaDotacaoLiberada & " RDL "
    '    strSql = strSql & "WHERE RD.PKId*=RDL.intReservaDotacao "
    strSQL = strSQL & " WHERE RD.PKId " & strOUTJSQLServer & "= RDL.intReservaDotacao " & strOUTJOracle & " "
    strSQL = strSQL & " AND " & gstrDATEPART(strYEAR, "RD.dtmData") & " = " & gintExercicio
    strSQL = strSQL & strFiltro
    strSQL = strSQL & " GROUP BY RD.PKId, RD.intNumero, RD.dtmData, RD.dblValor, RD.strSolicitante "
    strSQL = strSQL & " ORDER BY RD.PKId"
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
        Set tdb_ReservaDotacao.DataSource = adoResultado
        tdb_ReservaDotacao.Refresh
    End If
End Sub

Private Function LeTabelaEmpenho()
    Dim strSQL       As String
    Dim adoResultado As ADODB.Recordset
    
    strSQL = ""
    strSQL = strSQL & "SELECT EP.PKId, EP.intNumero , "
    strSQL = strSQL & "EP.dtmData,"
    strSQL = strSQL & "EP.dblValor dblEmpenhado,"
    strSQL = strSQL & "(SELECT " & gstrISNULL("SUM(dblValor)", "0") & " FROM " & gstrSubempenho & " WHERE intEmpenho = EP.PKId AND intNumero = 0 AND bytSituacao = 4) AS dblAnulado, "
    strSQL = strSQL & " (EP.dblValor - (SELECT " & gstrISNULL("SUM(dblValor)", "0") & " FROM " & gstrSubempenho & " WHERE intEmpenho = EP.PKId AND intNumero = 0 AND bytSituacao = 4)) AS dblValor, "
    strSQL = strSQL & " PT.strCodigo "
    strSQL = strSQL & "FROM "
    strSQL = strSQL & gstrEmpenho & " EP, "
    strSQL = strSQL & gstrProgramaDeTrabalho & " PT "
    strSQL = strSQL & "WHERE EP.intProgramaTrabalho = PT.PKId AND PT.intExercicio =  " & gintExercicio
    strSQL = strSQL & " AND intReservaDotacao = " & Val(tdb_ReservaDotacao.Columns(0))
    strSQL = strSQL & " AND " & gstrDATEPART(strYEAR, "EP.dtmData") & " = " & gintExercicio
    strSQL = strSQL & " ORDER BY EP.intNumero, EP.dtmData, PT.strCodigo "
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
        Set tdb_Empenho.DataSource = adoResultado
        tdb_Empenho.Refresh
    End If
End Function

Private Function LeMovimentosCancelamentos()
    Dim strSQL       As String
    Dim adoResultado As ADODB.Recordset
    
    strSQL = ""
    strSQL = strSQL & "SELECT RDL.PKId, RD.intNumero intReservaDotacao , RDL.intNumero , "
    strSQL = strSQL & "RDL.dtmData, RDL.dblValor, RDL.strHistorico AS Historico"
    strSQL = strSQL & " FROM " & gstrReservaDotacaoLiberada & " RDL, "
    strSQL = strSQL & gstrReservaDotacao & " RD"
    'strSql = strSql & " WHERE intReservaDotacao = " & Val(tdb_ReservaDotacao.Columns(0))
    strSQL = strSQL & " WHERE RDL.intReservaDotacao = " & Val(txtPKId.Text)
    strSQL = strSQL & " AND RDL.intFlag = 0 AND " & gstrDATEPART(strYEAR, "RDL.dtmData") & " = " & gintExercicio
    strSQL = strSQL & " AND RD.PKID = RDL.intReservaDotacao "
    strSQL = strSQL & " ORDER BY RD.intNumero,RDL.intNumero, RDL.dtmData"
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
        Set tdb_Cancelado.DataSource = adoResultado
        tdb_Cancelado.Refresh
    End If
End Function

Private Function AtualizaGridCancelamentos()
    Dim strSQL       As String
    Dim adoResultado As ADODB.Recordset
    
    strSQL = ""
    strSQL = strSQL & "SELECT PKId,intReservaDotacao, intNumero , "
    strSQL = strSQL & "dtmData, dblValor, strHistorico AS Historico"
    strSQL = strSQL & " FROM " & gstrReservaDotacaoLiberada
    strSQL = strSQL & " WHERE intFlag = 0 AND " & gstrDATEPART(strYEAR, "dtmData") & " = " & gintExercicio
    strSQL = strSQL & " ORDER BY intReservaDotacao,intNumero, dtmData"
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
        Set tdb_Cancelado.DataSource = adoResultado
        tdb_Cancelado.Refresh
    End If
End Function


Private Sub LeTotalReserva()
    Dim strSQL       As String
    Dim strSubSQL    As String
    Dim adoResultado As ADODB.Recordset
    Dim dblValor     As Double
    
    If txtDTMDATA.Text = "" Then
        txt_TotalReserva.Text = ""
        Exit Sub
    End If
    
    If cbo_intProgramaTrabalho.ListIndex = -1 Then
        txt_TotalReserva.Text = ""
        Exit Sub
    End If
    
    strSQL = ""
    strSQL = strSQL & "SELECT SUM(TMP.ValorReservaDotacao) ValorReservaDotacao, SUM(TMP.ValorReservaDotacaoLiberada) ValorReservaDotacaoLiberada FROM ("
    strSQL = strSQL & "SELECT " & gstrISNULL(" SUM(RD.dblValor) / COUNT (*)", "0") & " ValorReservaDotacao,"
    strSQL = strSQL & gstrISNULL("SUM(DL.dblValor)", "0") & " ValorReservaDotacaoLiberada "
    strSQL = strSQL & "FROM " & gstrReservaDotacaoLiberada & " DL ,"
    strSQL = strSQL & "(SELECT PKId, dblValor  FROM " & gstrReservaDotacao & "  "
    strSQL = strSQL & " WHERE intProgramaTrabalho=" & gstrItemData(cbo_intProgramaTrabalho)
    strSQL = strSQL & " AND " & gstrDATEPART(strMONTH, "DTMDATA") & " <= " & Month(CDate(txtDTMDATA)) & ")RD "
    strSQL = strSQL & "WHERE RD.Pkid " & strOUTJSQLServer & "= DL.intReservaDotacao " & strOUTJOracle & " "
    strSQL = strSQL & " AND " & gstrDATEPART(strMONTH, "DTMDATA") & " <= " & Month(CDate(txtDTMDATA))
    strSQL = strSQL & " GROUP BY RD.pkid)TMP"
    
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
        If Not adoResultado.EOF Then
            txt_TotalReserva.Text = gstrConvVrDoSql(adoResultado!ValorReservaDotacao - adoResultado!ValorReservaDotacaoLiberada)
        End If
    End If
    
End Sub

Private Sub GuardaValoresCancelamento(ByVal blnGrava As Boolean)
    'Guarda valores dos campos de cancelamento
    If blnGrava Then
        strGuardaValorTxt_DataCancelamento = txt_DataCancelamento
        strGuardaValorTxt_ValorCancelado = txt_ValorCancelado
        strGuardaValorTxt_HistoricoCancelamento = txt_HistoricoCancelamento
        intGuardaValorCbo_HistoricoCancelamento = cbo_HistoricoCancelamento.ListIndex
        intGuardaValorTdb_ReservaDotacao = tdb_ReservaDotacao.Bookmark
    Else
        'devolve valores dos campos de cancelamento
        txt_DataCancelamento = strGuardaValorTxt_DataCancelamento
        txt_ValorCancelado = strGuardaValorTxt_ValorCancelado
        txt_HistoricoCancelamento = strGuardaValorTxt_HistoricoCancelamento
        cbo_HistoricoCancelamento.ListIndex = intGuardaValorCbo_HistoricoCancelamento
        tdb_ReservaDotacao.Bookmark = intGuardaValorTdb_ReservaDotacao
        If tdb_Cancelado.ApproxCount > 0 Then
            tdb_Cancelado.Bookmark = tdb_Cancelado.ApproxCount
        End If
    End If
End Sub


Private Function strQueryAplicar() As String
    Dim strSQL As String
    strSQL = ""
    strSQL = strSQL & "SELECT PKId, intNumero "
    strSQL = strSQL & "FROM " & gstrReservaDotacao
    strSQL = strSQL & " ORDER BY intNumero "
    strQueryAplicar = strSQL
    
End Function

Private Function CarregaDadosSolicitacaoCompras(lngSolicitacao As Long, intExercicio As Integer, Optional blnVerificar As Boolean) As Boolean
    Dim strSQL       As String
    Dim adoResultado As ADODB.Recordset
    
    
    strSQL = ""
    strSQL = strSQL & "SELECT REQ.strRequisitante, REQ.strObjetoCompra, SUM(REQ.DBLQuantidade * REQ.dblValorEstimado) ValorEstimado, Max(dblValorParaReserva) ValorParaReserva, REQ.dtmDataRequisicao, REQ.intReserva, PRO.intCodigoReduzido, PRO.Pkid "
    strSQL = strSQL & "FROM " & gstrRequisicaoCompras & " REQ, " & gstrProgramaDeTrabalho & " PRO "
    strSQL = strSQL & "WHERE REQ.intCodigo = " & lngSolicitacao
    If intExercicio > 0 Then
        strSQL = strSQL & " AND REQ.intExercicio = " & intExercicio
    End If
    strSQL = strSQL & " AND REQ.intProgramaDeTrabalho " & strOUTJSQLServer & "= PRO.Pkid " & strOUTJOracle & " "
    strSQL = strSQL & " GROUP BY REQ.strRequisitante, REQ.strObjetoCompra, REQ.dtmDataRequisicao, REQ.intReserva, PRO.intCodigoReduzido, PRO.Pkid "
    Set gobjBanco = New clsBanco
    
    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
        With adoResultado
            If .EOF = False Then
                If blnVerificar = True Then
                    CarregaDadosSolicitacaoCompras = True
                    Exit Function
                End If
                
                If IsNull(!intCodigoReduzido) Then
                    strGuardaValorCboIntPrograma = ""
                    cbo_intProgramaTrabalho.ListIndex = -1
                    cbointProgramaTrabalho.ListIndex = -1
                End If
                
                'Caso ja exista reserva vamos posicionar no registro
                If Not IsNull(!intReserva) Then
                    If Not mblnAlterando Then
                        LeDaTabelaParaObj gstrReservaDotacao, tdb_ReservaDotacao, strQueryDotacao(!intReserva)
                        If tdb_ReservaDotacao.ApproxCount > 0 Then
                            mblnClickOk = True
                            tdb_ReservaDotacao_Click
                        Else
                            ExibeMensagem "Esta Solicitação possui uma Reserva que não pode ser encontrada."
                            CarregaDadosSolicitacaoCompras = False
                            Exit Function
                        End If
                    Else
                        If Not (!intReserva = txtPKId) Then
                            ExibeMensagem "Esta Solicitação já possui uma reserva."
                            CarregaDadosSolicitacaoCompras = False
                            Exit Function
                        Else
                            LeDaTabelaParaObj gstrReservaDotacao, tdb_ReservaDotacao, strQueryDotacao(!intReserva)
                            If tdb_ReservaDotacao.ApproxCount > 0 Then
                                mblnClickOk = True
                                tdb_ReservaDotacao_Click
                                CarregaDadosSolicitacaoCompras = True
                            End If
                        End If
                    End If
                Else
                    If Not mblnAlterando Then
                        DataAutomatica 'RetornaSugestaoData
                        dataPedido = gstrENulo(!dtmDataRequisicao)
                        txtdblValor = gstrConvVrDoSql(!ValorParaReserva)
                        txtstrHistorico = gstrENulo(!strObjetoCompra)
                        txtstrSolicitante = gstrENulo(!strRequisitante)
                        LeProgramaTrabalhoComReduzidoReservaDotacao cbo_intProgramaTrabalho, cbointProgramaTrabalho, CInt(gintExercicio) 'Year(!dtmDataRequisicao)
                        cbointProgramaTrabalho.ListIndex = gintIndiceCBO(cbointProgramaTrabalho, !Pkid)
                        TrocaCorObjeto txtdblValor, True
                        TrocaCorObjeto txtstrSolicitante, True
                        TrocaCorObjeto cbo_intProgramaTrabalho, False
                        TrocaCorObjeto cbointProgramaTrabalho, False
                    Else
                        LeProgramaTrabalhoComReduzidoReservaDotacao cbo_intProgramaTrabalho, cbointProgramaTrabalho, Val(strGuardaValorExercicio)  'Year(!dtmDataRequisicao)
                        'cbointProgramaTrabalho.ListIndex = cbointProgramaTrabalho.ListIndex = gintIndiceCBO(cbointProgramaTrabalho, strGuardaValorCboIntPrograma)
                        'cbointProgramaTrabalho.Text = adoResultado!intCodigoReduzido
                        
                        cbointProgramaTrabalho.ListIndex = gintIndiceCBO(cbointProgramaTrabalho, !Pkid)
                        
                        
                        txtdblValor = gstrConvVrDoSql(!ValorParaReserva)
                        dataPedido = gstrENulo(!dtmDataRequisicao)
                        txtstrHistorico = gstrENulo(!strObjetoCompra)
                        txtstrSolicitante = gstrENulo(!strRequisitante)
                        TrocaCorObjeto txtdblValor, True
                        CarregaDadosSolicitacaoCompras = True
                        TrocaCorObjeto txtstrSolicitante, True
                        TrocaCorObjeto cbo_intProgramaTrabalho, False
                        TrocaCorObjeto cbointProgramaTrabalho, False
                    End If
                    
                End If
            Else
                If Not blnVerificar Then
                    If Not mblnAlterando Then
                        LimpaDadosSolicitacaoCompras
                    End If
                End If
            End If
        End With
    End If
    
End Function

Private Function RetornaSugestaoData() As String
    Dim dtmDtEncerramento As Date
    Dim DtmUltimaReserva As Date
    
    dtmDtEncerramento = VerificaDataEncerramento("EO", gintExercicio)
    DtmUltimaReserva = LeDataUltimaReserva
    
    If dtmDtEncerramento = Empty And DtmUltimaReserva = Empty Then
        RetornaSugestaoData = gstrDataDoSistema
    Else
        txtDTMDATA = gstrDataFormatada(dtmDtEncerramento + 1)
        If Not IsEmpty(DtmUltimaReserva) And dtmDtEncerramento < DtmUltimaReserva Then
            RetornaSugestaoData = gstrDataFormatada(DtmUltimaReserva)
        End If
    End If
End Function

Private Sub LimpaDadosSolicitacaoCompras()
    
    DataAutomatica 'RetornaSugestaoData
    txtdblValor = Space$(0)
    txtstrHistorico = Space$(0)
    txtstrSolicitante = Space$(0)
    
    If mblnAlterando Then
        
        cbointProgramaTrabalho.ListIndex = -1
        cbo_intProgramaTrabalho.ListIndex = -1
        
        '        LeProgramaTrabalhoComReduzido cbo_intProgramaTrabalho, cbointProgramaTrabalho, Val(strGuardaValorExercicio)
        '        cbointProgramaTrabalho.ListIndex = cbointProgramaTrabalho.ListIndex = gintIndiceCBO(cbointProgramaTrabalho, strGuardaValorCboIntPrograma)
        '        TrocaCorObjeto txtdblValor, True
        '        preencheDotacaoByCodigo cbo_intProgramaTrabalho, cbointProgramaTrabalho
        
    Else
        cbo_intProgramaTrabalho.ListIndex = -1
        cbointProgramaTrabalho.ListIndex = -1
    End If
    
    TrocaCorObjeto txtdblValor, False
    TrocaCorObjeto txtstrSolicitante, False
    TrocaCorObjeto cbo_intProgramaTrabalho, False
    TrocaCorObjeto cbointProgramaTrabalho, False
    
End Sub

Private Function DataAutomatica()
    Dim adoResultado As ADODB.Recordset
    Dim strSQL       As String
    
    Set gobjBanco = New clsBanco
    
    If (bytDBType = EDatabases.Oracle) Then
        strSQL = "SELECT dtmData FROM " & gstrReservaDotacao & " WHERE TO_CHAR(dtmData,'YYYY') = " & gintExercicio & " AND PKID = (SELECT MAX(PKID) FROM tblReservaDotacao WHERE TO_CHAR(dtmData,'YYYY') = " & gintExercicio & ")"
    Else
        strSQL = "SELECT dtmData FROM " & gstrReservaDotacao & " WHERE YEAR(dtmData) = " & gintExercicio & " AND PKID = (SELECT MAX(PKID) FROM tblReservaDotacao WHERE YEAR(dtmData) = " & gintExercicio & ")"
    End If
    
    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
        
        If Not adoResultado.EOF Then
            
            txtDTMDATA = CDate(gstrDataFormatada(adoResultado!DTMDATA))
            
            DataAutomatica = CDate(gstrDataFormatada(adoResultado!DTMDATA))
            
            '            If Weekday(CDate(gstrDataFormatada(adoResultado!dtmFechamento)) + 1) = 7 Then
            '                txtdtmData = CDate(gstrDataFormatada(adoResultado!dtmFechamento)) + 2
            '            ElseIf Weekday(CDate(gstrDataFormatada(adoResultado!dtmFechamento)) + 1) = 1 Then
            '                txtdtmData = CDate(gstrDataFormatada(adoResultado!dtmFechamento)) + 1
            '            End If
            '            TrocaCorObjeto txtdtmData, True
            
        End If
        
    End If
    adoResultado.Close
    Set adoResultado = Nothing
    Set gobjBanco = Nothing
    
End Function

Private Function VerificaLancamentos(lngPkid As Long) As Boolean
    Dim strSQL As String
    Dim adoResultado As ADODB.Recordset
    
    VerificaLancamentos = True
    
    strSQL = ""
    strSQL = strSQL & " SELECT * FROM " & gstrEmpenho
    strSQL = strSQL & " WHERE intReservaDotacao = " & lngPkid
    
    Set gobjBanco = New clsBanco
    
    If Not gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
        ExibeMensagem "Ocorreram erros ao ler os movimentos da reserva."
        Exit Function
    Else
        
        If adoResultado.RecordCount = 0 Then
            strSQL = ""
            strSQL = strSQL & " SELECT * FROM " & gstrReservaDotacaoLiberada
            strSQL = strSQL & " WHERE intReservaDotacao = " & lngPkid
            
            Set gobjBanco = New clsBanco
            
            If Not gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
                ExibeMensagem "Ocorreram erros ao ler os movimentos da reserva."
                Exit Function
            Else
                If adoResultado.RecordCount = 0 Then
                    VerificaLancamentos = False
                End If
            End If
            
        End If
    End If
End Function

Public Sub LeProgramaTrabalhoComReduzidoReservaDotacao(cboCodigoReduzido As ComboBox, _
    cboProgramaTrabalho As ComboBox, _
    Optional intExercicio As Integer, _
    Optional strQuery As String)
    Dim strSQL          As String
    Dim adoResultado    As ADODB.Recordset
    cboProgramaTrabalho.Clear
    cboCodigoReduzido.Clear
    
    If strQuery = "" Then
        strSQL = ""
        
        strSQL = strSQL & "SELECT "
        strSQL = strSQL & " PT.PKId, "
        strSQL = strSQL & " PT.intCodigoReduzido, "
        strSQL = strSQL & " PT.strCodigo"
        strSQL = strSQL & " FROM "
        strSQL = strSQL & gstrProgramaDeTrabalho & " PT "
        strSQL = strSQL & " INNER JOIN "
        strSQL = strSQL & gstrReservaDotacao & " RD "
        strSQL = strSQL & " ON RD.INTPROGRAMATRABALHO = PT.PKID "
        strSQL = strSQL & "WHERE "
        strSQL = strSQL & " PT.intExercicio = " & gintExercicio
        strSQL = strSQL & " AND PT.intCodigoReduzido > 0"
        strSQL = strSQL & " AND PT.STRCODIGO <> ' '"
        strSQL = strSQL & " GROUP BY PT.PKId,  PT.intCodigoReduzido,  PT.strCodigo"
        strSQL = strSQL & " ORDER BY PT.strCodigo  "
        
    Else
        strSQL = strQuery
    End If
    
    
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
        With adoResultado
            Do While Not .EOF
                    cboProgramaTrabalho.AddItem !strCodigo
                    cboProgramaTrabalho.ItemData(cboProgramaTrabalho.NewIndex) = !Pkid
                .MoveNext
            Loop
        End With
    End If
    
    If strQuery = "" Then
        strSQL = Mid(strSQL, 1, InStr(1, strSQL, "ORDER") - 1) + " ORDER BY PT.intCodigoReduzido"
    End If
    
    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
        With adoResultado
            Do While Not .EOF
                    cboCodigoReduzido.AddItem !intCodigoReduzido
                    cboCodigoReduzido.ItemData(cboCodigoReduzido.NewIndex) = !Pkid
                .MoveNext
            Loop
        End With
    End If
End Sub
