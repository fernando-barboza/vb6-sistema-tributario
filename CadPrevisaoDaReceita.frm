VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frmCadPrevisaoDaReceita 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Previsão da Receita"
   ClientHeight    =   8040
   ClientLeft      =   1500
   ClientTop       =   2415
   ClientWidth     =   11325
   HelpContextID   =   27
   Icon            =   "CadPrevisaoDaReceita.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8040
   ScaleWidth      =   11325
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd_intOrgao 
      Height          =   300
      Left            =   10460
      Picture         =   "CadPrevisaoDaReceita.frx":1042
      Style           =   1  'Graphical
      TabIndex        =   22
      TabStop         =   0   'False
      ToolTipText     =   "Clique para cadastar o orgão"
      Top             =   2300
      Width           =   330
   End
   Begin VB.CommandButton cmd_intModalidade 
      Height          =   300
      Left            =   10440
      Picture         =   "CadPrevisaoDaReceita.frx":13CC
      Style           =   1  'Graphical
      TabIndex        =   18
      TabStop         =   0   'False
      ToolTipText     =   "Clique para cadastar a modaliade de aplicação"
      Top             =   1920
      Width           =   330
   End
   Begin TabDlg.SSTab tab_3dPasta 
      Height          =   7875
      Left            =   120
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   60
      Width           =   11115
      _ExtentX        =   19606
      _ExtentY        =   13891
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "Previsão da Receita"
      TabPicture(0)   =   "CadPrevisaoDaReceita.frx":1756
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lbldblValor"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblintFonteRecurso"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblintCodigoorcamentario"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblintCodigoReduzido"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lbl_Total"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblstrLegislacao"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lbl_Evento"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lblModalidade"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label1"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cmd_CodigoOrcamentario"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "cmd_intFonteRecurso"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "tdb_Lista"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtdblValor"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtPKId"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txt_Total"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txtintCodigoReduzido"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "dbcintFonteRecurso"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "dbcintCodigoOrcamentario"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "cbo_intCodigoOrcamentario"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "txtintExercicio"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "txtstrLegislacao"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "txt_dblValor"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "dbcintEvento"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "txt_intEvento"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "cmd_Evento"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "dbcintModalidade"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "txt_intOrgao"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "dbcintOrgao"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "fra_Integrante"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "txt_intModalidade"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "txt_intConvenio"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).ControlCount=   31
      Begin VB.TextBox txt_intConvenio 
         Height          =   315
         Left            =   2790
         MaxLength       =   2
         TabIndex        =   14
         Top             =   1860
         Width           =   705
      End
      Begin VB.TextBox txt_intModalidade 
         Height          =   315
         Left            =   2040
         MaxLength       =   3
         TabIndex        =   13
         Top             =   1860
         Width           =   705
      End
      Begin VB.Frame fra_Integrante 
         Caption         =   "Integrante"
         Height          =   555
         Left            =   2040
         TabIndex        =   33
         Top             =   3540
         Width           =   4095
         Begin VB.CheckBox chkbytEducacao 
            Caption         =   "Educação"
            Height          =   255
            Left            =   2040
            TabIndex        =   35
            Top             =   240
            Width           =   1335
         End
         Begin VB.CheckBox chkbytSaude 
            Caption         =   "Saúde"
            Height          =   255
            Left            =   600
            TabIndex        =   34
            Top             =   240
            Width           =   1455
         End
      End
      Begin MSDataListLib.DataCombo dbcintOrgao 
         Height          =   315
         Left            =   3540
         TabIndex        =   21
         Top             =   2235
         Width           =   6765
         _ExtentX        =   11933
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin VB.TextBox txt_intOrgao 
         Height          =   315
         Left            =   2040
         TabIndex        =   20
         Top             =   2230
         Width           =   1455
      End
      Begin MSDataListLib.DataCombo dbcintModalidade 
         Height          =   315
         Left            =   3540
         TabIndex        =   17
         Top             =   1860
         Width           =   6765
         _ExtentX        =   11933
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin VB.CommandButton cmd_Evento 
         Height          =   315
         Left            =   10320
         Picture         =   "CadPrevisaoDaReceita.frx":1772
         Style           =   1  'Graphical
         TabIndex        =   6
         TabStop         =   0   'False
         Tag             =   "200"
         ToolTipText     =   "Clique aqui para cadastar Evento Contábil"
         Top             =   750
         Width           =   335
      End
      Begin VB.TextBox txt_intEvento 
         Height          =   315
         Left            =   2040
         TabIndex        =   4
         Top             =   765
         Width           =   1455
      End
      Begin VB.ComboBox dbcintEvento 
         Height          =   315
         Left            =   3540
         TabIndex        =   5
         Top             =   750
         Width           =   6765
      End
      Begin VB.TextBox txt_dblValor 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   9600
         TabIndex        =   31
         Top             =   4560
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.TextBox txtstrLegislacao 
         Height          =   525
         Left            =   2040
         MaxLength       =   500
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   23
         Top             =   2610
         Width           =   8835
      End
      Begin VB.TextBox txtintExercicio 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000000&
         CausesValidation=   0   'False
         Enabled         =   0   'False
         Height          =   285
         Left            =   7560
         MultiLine       =   -1  'True
         OLEDragMode     =   1  'Automatic
         OLEDropMode     =   1  'Manual
         TabIndex        =   30
         Top             =   30
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.ComboBox cbo_intCodigoOrcamentario 
         Height          =   315
         Left            =   2040
         Sorted          =   -1  'True
         TabIndex        =   8
         Top             =   1110
         Width           =   1455
      End
      Begin VB.ComboBox dbcintCodigoOrcamentario 
         Height          =   315
         Left            =   3540
         Sorted          =   -1  'True
         TabIndex        =   9
         Top             =   1110
         Width           =   6765
      End
      Begin MSDataListLib.DataCombo dbcintFonteRecurso 
         Height          =   315
         Left            =   2040
         TabIndex        =   12
         Top             =   1485
         Width           =   8265
         _ExtentX        =   14579
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin VB.TextBox txtintCodigoReduzido 
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         Height          =   285
         Left            =   2040
         OLEDragMode     =   1  'Automatic
         OLEDropMode     =   1  'Manual
         TabIndex        =   2
         Top             =   420
         Width           =   1665
      End
      Begin VB.TextBox txt_Total 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000000&
         CausesValidation=   0   'False
         Enabled         =   0   'False
         Height          =   285
         Left            =   4200
         MultiLine       =   -1  'True
         OLEDragMode     =   1  'Automatic
         OLEDropMode     =   1  'Manual
         TabIndex        =   26
         Top             =   3195
         Width           =   1485
      End
      Begin VB.TextBox txtPKId 
         Enabled         =   0   'False
         Height          =   285
         Left            =   8640
         OLEDragMode     =   1  'Automatic
         OLEDropMode     =   1  'Manual
         TabIndex        =   28
         Top             =   30
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.TextBox txtdblValor 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2040
         MaxLength       =   15
         MultiLine       =   -1  'True
         TabIndex        =   25
         Top             =   3195
         Width           =   1485
      End
      Begin TrueOleDBGrid70.TDBGrid tdb_Lista 
         Height          =   3555
         Left            =   90
         TabIndex        =   27
         Top             =   4200
         Width           =   10905
         _ExtentX        =   19235
         _ExtentY        =   6271
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "PKID"
         Columns(0).DataField=   "PKID"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Cod. Reduzido"
         Columns(1).DataField=   "intCodigoReduzido"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Código"
         Columns(2).DataField=   "strCodigoOrcamentario"
         Columns(2).NumberFormat=   "FormatText Event"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Descrição"
         Columns(3).DataField=   "strDescricao"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "Orgão"
         Columns(4).DataField=   "Orgao"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "Valor"
         Columns(5).DataField=   "dblValor"
         Columns(5).NumberFormat=   "Standard"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "Total"
         Columns(6).DataField=   "dblTotal"
         Columns(6).NumberFormat=   "Standard"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   0
         Columns(7)._MaxComboItems=   5
         Columns(7).Caption=   "bytSituacao"
         Columns(7).DataField=   "bytSituacao"
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
         Splits(0)._ColumnProps(8)=   "Column(1).Width=2037"
         Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=1958"
         Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(13)=   "Column(2).Width=2461"
         Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=2381"
         Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(17)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(18)=   "Column(3).Width=7276"
         Splits(0)._ColumnProps(19)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(20)=   "Column(3)._WidthInPix=7197"
         Splits(0)._ColumnProps(21)=   "Column(3)._EditAlways=0"
         Splits(0)._ColumnProps(22)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(23)=   "Column(4).Width=4392"
         Splits(0)._ColumnProps(24)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(25)=   "Column(4)._WidthInPix=4313"
         Splits(0)._ColumnProps(26)=   "Column(4)._EditAlways=0"
         Splits(0)._ColumnProps(27)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(28)=   "Column(5).Width=2461"
         Splits(0)._ColumnProps(29)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(30)=   "Column(5)._WidthInPix=2381"
         Splits(0)._ColumnProps(31)=   "Column(5)._EditAlways=0"
         Splits(0)._ColumnProps(32)=   "Column(5)._ColStyle=514"
         Splits(0)._ColumnProps(33)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(34)=   "Column(6).Width=2725"
         Splits(0)._ColumnProps(35)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(36)=   "Column(6)._WidthInPix=2646"
         Splits(0)._ColumnProps(37)=   "Column(6)._EditAlways=0"
         Splits(0)._ColumnProps(38)=   "Column(6).AllowSizing=0"
         Splits(0)._ColumnProps(39)=   "Column(6).Visible=0"
         Splits(0)._ColumnProps(40)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(41)=   "Column(7).Width=2725"
         Splits(0)._ColumnProps(42)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(43)=   "Column(7)._WidthInPix=2646"
         Splits(0)._ColumnProps(44)=   "Column(7)._EditAlways=0"
         Splits(0)._ColumnProps(45)=   "Column(7).AllowSizing=0"
         Splits(0)._ColumnProps(46)=   "Column(7).Visible=0"
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
         MultipleLines   =   0
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
         _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
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
         _StyleDefs(37)  =   "Splits(0).Columns(0).Style:id=54,.parent=13"
         _StyleDefs(38)  =   "Splits(0).Columns(0).HeadingStyle:id=51,.parent=14"
         _StyleDefs(39)  =   "Splits(0).Columns(0).FooterStyle:id=52,.parent=15"
         _StyleDefs(40)  =   "Splits(0).Columns(0).EditorStyle:id=53,.parent=17"
         _StyleDefs(41)  =   "Splits(0).Columns(1).Style:id=46,.parent=13"
         _StyleDefs(42)  =   "Splits(0).Columns(1).HeadingStyle:id=43,.parent=14"
         _StyleDefs(43)  =   "Splits(0).Columns(1).FooterStyle:id=44,.parent=15"
         _StyleDefs(44)  =   "Splits(0).Columns(1).EditorStyle:id=45,.parent=17"
         _StyleDefs(45)  =   "Splits(0).Columns(2).Style:id=28,.parent=13"
         _StyleDefs(46)  =   "Splits(0).Columns(2).HeadingStyle:id=25,.parent=14"
         _StyleDefs(47)  =   "Splits(0).Columns(2).FooterStyle:id=26,.parent=15"
         _StyleDefs(48)  =   "Splits(0).Columns(2).EditorStyle:id=27,.parent=17"
         _StyleDefs(49)  =   "Splits(0).Columns(3).Style:id=32,.parent=13"
         _StyleDefs(50)  =   "Splits(0).Columns(3).HeadingStyle:id=29,.parent=14"
         _StyleDefs(51)  =   "Splits(0).Columns(3).FooterStyle:id=30,.parent=15"
         _StyleDefs(52)  =   "Splits(0).Columns(3).EditorStyle:id=31,.parent=17"
         _StyleDefs(53)  =   "Splits(0).Columns(4).Style:id=66,.parent=13"
         _StyleDefs(54)  =   "Splits(0).Columns(4).HeadingStyle:id=63,.parent=14"
         _StyleDefs(55)  =   "Splits(0).Columns(4).FooterStyle:id=64,.parent=15"
         _StyleDefs(56)  =   "Splits(0).Columns(4).EditorStyle:id=65,.parent=17"
         _StyleDefs(57)  =   "Splits(0).Columns(5).Style:id=50,.parent=13,.alignment=1"
         _StyleDefs(58)  =   "Splits(0).Columns(5).HeadingStyle:id=47,.parent=14,.alignment=2"
         _StyleDefs(59)  =   "Splits(0).Columns(5).FooterStyle:id=48,.parent=15"
         _StyleDefs(60)  =   "Splits(0).Columns(5).EditorStyle:id=49,.parent=17"
         _StyleDefs(61)  =   "Splits(0).Columns(6).Style:id=58,.parent=13"
         _StyleDefs(62)  =   "Splits(0).Columns(6).HeadingStyle:id=55,.parent=14"
         _StyleDefs(63)  =   "Splits(0).Columns(6).FooterStyle:id=56,.parent=15"
         _StyleDefs(64)  =   "Splits(0).Columns(6).EditorStyle:id=57,.parent=17"
         _StyleDefs(65)  =   "Splits(0).Columns(7).Style:id=62,.parent=13"
         _StyleDefs(66)  =   "Splits(0).Columns(7).HeadingStyle:id=59,.parent=14"
         _StyleDefs(67)  =   "Splits(0).Columns(7).FooterStyle:id=60,.parent=15"
         _StyleDefs(68)  =   "Splits(0).Columns(7).EditorStyle:id=61,.parent=17"
         _StyleDefs(69)  =   "Named:id=33:Normal"
         _StyleDefs(70)  =   ":id=33,.parent=0"
         _StyleDefs(71)  =   "Named:id=34:Heading"
         _StyleDefs(72)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(73)  =   ":id=34,.wraptext=-1"
         _StyleDefs(74)  =   "Named:id=35:Footing"
         _StyleDefs(75)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(76)  =   "Named:id=36:Selected"
         _StyleDefs(77)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(78)  =   "Named:id=37:Caption"
         _StyleDefs(79)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(80)  =   "Named:id=38:HighlightRow"
         _StyleDefs(81)  =   ":id=38,.parent=33,.bgcolor=&H80000014&,.fgcolor=&H80000012&"
         _StyleDefs(82)  =   "Named:id=39:EvenRow"
         _StyleDefs(83)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(84)  =   "Named:id=40:OddRow"
         _StyleDefs(85)  =   ":id=40,.parent=33"
         _StyleDefs(86)  =   "Named:id=41:RecordSelector"
         _StyleDefs(87)  =   ":id=41,.parent=34"
         _StyleDefs(88)  =   "Named:id=42:FilterBar"
         _StyleDefs(89)  =   ":id=42,.parent=33"
      End
      Begin VB.CommandButton cmd_intFonteRecurso 
         Height          =   300
         Left            =   10320
         Picture         =   "CadPrevisaoDaReceita.frx":1AFC
         Style           =   1  'Graphical
         TabIndex        =   15
         TabStop         =   0   'False
         ToolTipText     =   "Clique para cadastar fonte de recurso"
         Top             =   1485
         Width           =   330
      End
      Begin VB.CommandButton cmd_CodigoOrcamentario 
         Height          =   300
         Left            =   10320
         Picture         =   "CadPrevisaoDaReceita.frx":1E86
         Style           =   1  'Graphical
         TabIndex        =   10
         TabStop         =   0   'False
         ToolTipText     =   "Clique para cadastar código orçamentário"
         Top             =   1110
         Width           =   330
      End
      Begin VB.Label Label1 
         Caption         =   "Orgão"
         Height          =   255
         Left            =   1440
         TabIndex        =   32
         Top             =   2280
         Width           =   495
      End
      Begin VB.Label lblModalidade 
         Caption         =   "Modalidade de Aplicação"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   1920
         Width           =   1815
      End
      Begin VB.Label lbl_Evento 
         AutoSize        =   -1  'True
         Caption         =   "Evento Contábil"
         Height          =   195
         Left            =   795
         TabIndex        =   3
         Top             =   795
         Width           =   1125
      End
      Begin VB.Label lblstrLegislacao 
         AutoSize        =   -1  'True
         Caption         =   "Legislação"
         Height          =   195
         Left            =   1200
         TabIndex        =   19
         Top             =   2745
         Width           =   765
      End
      Begin VB.Label lbl_Total 
         AutoSize        =   -1  'True
         Caption         =   "Total"
         Height          =   195
         Left            =   3600
         TabIndex        =   29
         Top             =   3285
         Width           =   360
      End
      Begin VB.Label lblintCodigoReduzido 
         AutoSize        =   -1  'True
         Caption         =   "Código Reduzido"
         Height          =   195
         Left            =   750
         TabIndex        =   1
         Top             =   450
         Width           =   1215
      End
      Begin VB.Label lblintCodigoorcamentario 
         AutoSize        =   -1  'True
         Caption         =   "Código Orçamentário"
         Height          =   195
         Left            =   480
         TabIndex        =   7
         Top             =   1170
         Width           =   1485
      End
      Begin VB.Label lblintFonteRecurso 
         AutoSize        =   -1  'True
         Caption         =   "Fonte de Recurso"
         Height          =   195
         Left            =   690
         TabIndex        =   11
         Top             =   1545
         Width           =   1515
      End
      Begin VB.Label lbldblValor 
         AutoSize        =   -1  'True
         Caption         =   "Valor"
         Height          =   195
         Left            =   1605
         TabIndex        =   24
         Top             =   3285
         Width           =   360
      End
   End
End
Attribute VB_Name = "frmCadPrevisaoDaReceita"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mblnProgramaAAprovar    As Boolean
Dim mblnClickOk             As Boolean
Dim mblnAlterando           As Boolean
Dim mobjAux                 As Object
Dim mblnselecionou          As Boolean
Dim mstrQueryAplicar        As String
Public blnOrcamento         As Boolean

Dim intPkIdRow              As Variant
Dim blnFilterBar            As Boolean
Dim dblValorAnt             As Double

Dim strCodAnt               As String
Dim strDescrAnt             As String

Dim itemAnterior            As String 'ORC1543 - FERNANDO

Private Sub cbo_intCodigoOrcamentario_Change()
    If Len(cbo_intCodigoOrcamentario.Text) = 8 Then
        LeCodigoOrcamentarioGeral cbo_intCodigoOrcamentario, dbcintCodigoOrcamentario, "SELECT * FROM " & gstrCodigoOrcamentario & " WHERE strCodigoOrcamentario LIKE '" & Trim(cbo_intCodigoOrcamentario.Text) & "%' AND intExercicio = " & IIf(blnOrcamento, gintExercicio, gintExercicio + 1)
        If cbo_intCodigoOrcamentario.ListCount > 0 Then
            cbo_intCodigoOrcamentario.ListIndex = 0
        End If
    End If
End Sub

Private Sub cbo_intCodigoOrcamentario_Click()
    
    dbcintCodigoOrcamentario.ListIndex = gintIndiceCBO(dbcintCodigoOrcamentario, _
                                                        gstrItemData(cbo_intCodigoOrcamentario))
   
End Sub

Private Sub cbo_intCodigoOrcamentario_GotFocus()
    MarcaCampo cbo_intCodigoOrcamentario
End Sub

Private Sub cbo_intCodigoOrcamentario_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", cbo_intCodigoOrcamentario
    'gstrLimitaCampoValor cbo_intCodigoOrcamentario, KeyAscii, 8, 0
End Sub

Private Sub cbo_intCodigoOrcamentario_LostFocus()
    If cbo_intCodigoOrcamentario.Text = "" Then dbcintCodigoOrcamentario.ListIndex = -1
End Sub

Private Sub cmd_CodigoOrcamentario_Click()
    
    If blnOrcamento = True Then
       gbytMenu = gbytMenuCadastro
    Else
       gbytMenu = gbytMenuProposta
    End If
    
    CarregaForm frmCadCodigoOrcamentario, dbcintCodigoOrcamentario
    
End Sub

Private Sub cmd_Evento_Click()
   CarregaForm frmCadEvento, dbcintEvento
End Sub

Private Sub cmd_intFonteRecurso_Click()
    If blnOrcamento = True Then
       gbytMenu = gbytMenuCadastro
    Else
       gbytMenu = gbytMenuProposta
    End If
    CarregaForm frmCadFonteRecurso, dbcintFonteRecurso
End Sub

Private Function strQueryAplicar() As String
    Dim strSQL As String
    If Trim(mstrQueryAplicar) = "" Then
        strSQL = ""
        strSQL = strSQL & "SELECT CO.PKId, CO.strDescricao FROM "
        strSQL = strSQL & gstrCodigoOrcamentario & " CO, "
        strSQL = strSQL & gstrPrevisaoDaReceita & " PR "
        strSQL = strSQL & "WHERE CO.PKId = PR.intCodigoOrcamentario "
        strSQL = strSQL & "ORDER BY CO.strDescricao"
        strQueryAplicar = strSQL
    Else
        strQueryAplicar = mstrQueryAplicar
    End If
End Function

Private Function strQueryCO() As String
Dim strSQL As String
'******************************************************************************************
' Data: 09/06/2003
' Alteração: - Substituição do comando nativo ISNULL() do SQL Server pela função gstrISNULL
' Responsável: Everton Bianchini
'******************************************************************************************

    strSQL = ""
    strSQL = strSQL & "SELECT PR.PKId, PR.intCodigoReduzido, CO.strCodigoOrcamentario, "
    strSQL = strSQL & "CO.strDescricao, PR.dblValor, PR.bytSituacao, "
    strSQL = strSQL & "OG.strDescricao Orgao, "
    strSQL = strSQL & "(SELECT " & gstrISNULL("SUM(dblValor)", "0") & " FROM "
    strSQL = strSQL & gstrPrevisaoDaReceita & " "
    strSQL = strSQL & "WHERE intExercicio = " & IIf(blnOrcamento, gintExercicio, gintExercicio + 1) & ") AS dblTotal "
    strSQL = strSQL & "FROM "
    strSQL = strSQL & gstrPrevisaoDaReceita & " PR, "
    strSQL = strSQL & gstrCodigoOrcamentario & " CO, "
    strSQL = strSQL & gstrOrgao & " OG "
    strSQL = strSQL & "WHERE PR.intCodigoOrcamentario = CO.PKId "
    strSQL = strSQL & "AND PR.intOrgao " & strOUTJSQLServer & "= OG.Pkid " & strOUTJOracle
    strSQL = strSQL & "AND PR.intExercicio = " & IIf(blnOrcamento, gintExercicio, gintExercicio + 1) & " "
    strSQL = strSQL & "ORDER BY CO.strCodigoOrcamentario"
    
    strQueryCO = strSQL
    
End Function

Private Sub cmd_intModalidade_Click()
        CarregaForm frmCadModalidadeAplicacao, dbcintModalidade
End Sub

Private Sub cmd_intOrgao_Click()
        CarregaForm frmCadOrgao, dbcintOrgao
End Sub

Private Sub dbcintModalidade_Change()
    If dbcintModalidade.MatchedWithList Then
        CarregaCodModalidade
    End If
End Sub

Private Sub dbcintModalidade_LostFocus()
    If Not dbcintModalidade.MatchedWithList Then
        txt_intModalidade.Text = ""
    End If
End Sub

'Private Sub dbcintOrgao_Change()
''ORC1543 - FERNANDO
''    If dbcintOrgao.Text = "" Then
''        'If Not strCodAnt <> txt_intOrgao Then
''        If Not strDescrAnt <> dbcintOrgao.Text Then
''            txt_intOrgao = ""
''            strCodAnt = txt_intOrgao
''        End If
''    End If
''ORC1543 - FERNANDO
'End Sub

Private Sub dbcintOrgao_Click(Area As Integer)
'ORC1543 - FERNANDO
'   If Len(dbcintOrgao.Text) > 0 Then
'    PreencheOrgao dbcintOrgao.BoundText, 1
'   End If
'    strDescrAnt = dbcintOrgao.Text
'ORC1543 - FERNANDO
    DropDownDataCombo dbcintOrgao, Me, Area
    
    If dbcintOrgao.MatchedWithList Then
        
        txt_intOrgao = LeCodOrgao(dbcintOrgao.BoundText)
        If itemAnterior = dbcintOrgao.BoundText Then Exit Sub
        itemAnterior = dbcintOrgao.BoundText
        
    End If
End Sub

Private Sub dbcintCodigoOrcamentario_Click()
        cbo_intCodigoOrcamentario.ListIndex = gintIndiceCBO(cbo_intCodigoOrcamentario, _
                                                        gstrItemData(dbcintCodigoOrcamentario))
End Sub

Private Sub dbcintCodigoOrcamentario_GotFocus()
    MarcaCampo dbcintCodigoOrcamentario
End Sub

Private Sub dbcintCodigoOrcamentario_KeyPress(KeyAscii As Integer)
   CaracterValido KeyAscii
End Sub

Private Sub dbcintCodigoOrcamentario_LostFocus()
    If dbcintCodigoOrcamentario.Text = "" Then cbo_intCodigoOrcamentario.ListIndex = -1
End Sub

Private Sub dbcintEvento_Click()
   leCodigoEvento txt_intEvento, dbcintEvento
   If txt_intEvento = "0" Then txt_intEvento = ""
End Sub

Private Sub dbcintEvento_KeyPress(KeyAscii As Integer)
   CaracterValido KeyAscii
End Sub

Private Sub dbcintEvento_LostFocus()
    If dbcintEvento.ListIndex = -1 Or Trim(dbcintEvento.Text) = "" Then txt_intEvento.Text = ""
End Sub

Private Sub dbcintFonteRecurso_Click(Area As Integer)
    DropDownDataCombo dbcintFonteRecurso, Me, Area
End Sub

Private Sub dbcintFonteRecurso_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbcintFonteRecurso, Me, , KeyCode, Shift
End Sub

Private Sub dbcintFonteRecurso_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, , dbcintFonteRecurso
End Sub

Private Sub dbcintOrgao_GotFocus()
    MarcaCampo dbcintOrgao
    If dbcintOrgao.BoundText <> "" Then
    dbcintOrgao_Click 0
    End If
End Sub

Private Sub dbcintOrgao_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbcintOrgao, Me, , KeyCode, Shift
End Sub

Private Sub dbcintOrgao_KeyPress(KeyAscii As Integer)
CaracterValido KeyAscii, "A", dbcintOrgao
End Sub

Private Sub dbcintOrgao_LostFocus()
    If Not dbcintOrgao.MatchedWithList Then
        txt_intOrgao = Space$(0)
    End If
End Sub

Private Sub Form_Activate()
    If blnOrcamento Then
        gintCodSeguranca = 837
    Else
        gintCodSeguranca = 761
    End If
    
    VirificaGradeListView Me
    If blnOrcamento = True Then
       txtintExercicio = gintExercicio
       txtintCodigoReduzido.OLEDropMode = 0
       TrocaCorObjeto txtdblValor, True
       txtdblValor = "0,00"
       TrocaCorObjeto txtintCodigoReduzido, False
       txtintCodigoReduzido.SetFocus
    Else
       txtintExercicio = gintExercicio + 1
    End If
    
    HabilitaDesabilitaBotao1 mblnProgramaAAprovar, gstrBtnArquivo, gstrGeraCodigoReduzido
    
    If mblnselecionou Then
        HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrDeletar
    Else
        HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrDeletar
    End If
    If mobjAux Is Nothing Then
        HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar
    Else
        HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrAplicar
    End If
    
    If blnOrcamento = True Then
       HabilitaDesabilitaBotao1 (gbytMenu = gbytMenuCadastro), gstrBtnArquivo, gstrGeraFundef
    Else
       HabilitaDesabilitaBotao1 (gbytMenu = gbytMenuProposta), gstrBtnArquivo, gstrGeraFundef
    End If
    
End Sub

Private Sub Form_Deactivate()
    HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrGeraCodigoReduzido, gstrGeraFundef
End Sub

Private Sub Form_GotFocus()
   dbcintCodigoOrcamentario.SetFocus
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
  If KeyCode = vbKeyF1 Then
     Call_HtmlHelp Me.HelpContextID
  End If
End Sub

Private Sub Form_Load()
    
    LeCodigoOrcamentarioGeral cbo_intCodigoOrcamentario, dbcintCodigoOrcamentario, "SELECT * FROM " & gstrCodigoOrcamentario & " WHERE intExercicio = " & IIf(blnOrcamento, gintExercicio, gintExercicio + 1)
    LeDaTabelaParaObj "", dbcintEvento, strQueryEventoInicial
    
    dbcintFonteRecurso.Tag = "SELECT PKId, strDescricao FROM " & gstrFonteRecurso & " WHERE intExercicio = " & IIf(blnOrcamento, gintExercicio, gintExercicio + 1) & ";strDescricao"
    dbcintModalidade.Tag = "SELECT PKID, strDescricao FROM " & gstrModalidade & " ORDER BY strDescricao ;strDescricao"
    dbcintOrgao.Tag = "SELECT PKID, strDescricao FROM " & gstrOrgao & " WHERE INTEXERCICIO =  " & IIf(blnOrcamento, gintExercicio, gintExercicio + 1) & " ORDER BY STRDESCRICAO ;strDescricao "
    
    VerificaListaAutomatica "", tdb_Lista, strQueryCO
    VerificaObjParaAplicar mobjAux, mstrQueryAplicar
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Form_Deactivate
End Sub



Private Sub tdb_Lista_BeforeRowColChange(Cancel As Integer)
    blnFilterBar = False
End Sub

Private Sub tdb_Lista_Click()
    If glngQtdLinhaTDBGrid(tdb_Lista) = 1 Then
        tdb_Lista_RowColChange 0, 0
    End If
End Sub

Private Sub tdb_Lista_DataSourceChanged()
    With tdb_Lista
        txt_total = .Columns("dblTotal")
        mblnProgramaAAprovar = Abs(Val(.Columns("bytSituacao"))) - 1
        
    If txtintCodigoReduzido.Text <> "" Then
        HabilitaDesabilitaBotao1 mblnProgramaAAprovar, _
                                 gstrBtnArquivo, gstrNovo, gstrSalvar, _
                                 gstrGeraCodigoReduzido, gstrDeletar
    End If
    End With
End Sub

Private Sub tdb_Lista_DblClick()
    MantemForm gstrAplicar
End Sub

Private Sub tdb_Lista_FilterChange()
    blnFilterBar = True
    gblnFilraCampos tdb_Lista
End Sub

Private Sub tdb_Lista_FormatText(ByVal ColIndex As Integer, Value As Variant, Bookmark As Variant)
    Value = gvntFormatacaoEspecifica(Value, 2)
End Sub

Private Sub tdb_Lista_HeadClick(ByVal ColIndex As Integer)
    gOrdenaGrid tdb_Lista, ColIndex
End Sub

Private Sub tdb_Lista_KeyDown(KeyCode As Integer, Shift As Integer)
    mblnClickOk = True
End Sub

Private Sub tdb_Lista_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        EditaValor tdb_Lista.Columns("PkId").Value, 9, tdb_Lista.Row
        KeyAscii = 0
    Else
        CaracterValido KeyAscii
    End If
End Sub

Private Sub tdb_Lista_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mblnClickOk = True
End Sub

Private Sub tdb_Lista_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    Dim blnAprovado As Boolean
    With tdb_Lista
    
        If (Not .EOF And Not .BOF) And mblnClickOk Then
            If blnFilterBar Then Exit Sub
            mblnClickOk = False
            txtPKID.Text = .Columns("PKID").Value
            gCorLinhaSelecionada tdb_Lista
            blnAprovado = Val(.Columns("bytSituacao"))
            
            LeDaTabelaParaObj "", dbcintEvento, strQueryEventoInicial
            LeCodigoOrcamentarioGeral cbo_intCodigoOrcamentario, dbcintCodigoOrcamentario, "SELECT * FROM " & gstrCodigoOrcamentario & " WHERE intExercicio = " & IIf(blnOrcamento, gintExercicio, gintExercicio + 1)
            LeDaTabelaParaObj gstrFonteRecurso, dbcintFonteRecurso, "SELECT PKId, strDescricao FROM " & gstrFonteRecurso & " WHERE intExercicio = " & IIf(blnOrcamento, gintExercicio, gintExercicio + 1)
            PreencherListaDeOpcoes dbcintModalidade
            
            LeDaTabelaParaObj gstrPrevisaoDaReceita, Me, blnAprovado
             PreencheOrgao IIf(dbcintOrgao.BoundText <> "", dbcintOrgao.BoundText, 0), 1
            
            If mobjAux Is Nothing Then
                HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar
            Else
                HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrAplicar
            End If
            HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrDeletar
            
            EditaValor tdb_Lista.Columns("PkId").Value, tdb_Lista.Col, tdb_Lista.Row
            intPkIdRow = .Columns("PkId").Value
            mblnAlterando = True
        End If
    End With
End Sub

Public Sub MantemForm(ByVal strModoOperacao As String)

    If UCase(strModoOperacao) = gstrDeletar And blnOrcamento = True Then
        ExibeMensagem "Não é possível excluir uma Receita já cadastrada."
        Exit Sub
    End If
    
    If UCase(strModoOperacao) = gstrSalvar And mblnAlterando = True And blnOrcamento = True Then
        If txtdblValor = "0,00" Then
            If blnDadosOk Then
                ToolBarGeral strModoOperacao, gstrPrevisaoDaReceita, mblnAlterando, _
                             tdb_Lista, Me, mobjAux, strQueryCO, strQueryAplicar, _
                             rptPrevisaoDaReceita, strQueryRelatorio
            End If
            LeDaTabelaParaObj "", tdb_Lista, strQueryCO
            txt_total = tdb_Lista.Columns("dblTotal")
            txt_intOrgao.Text = ""
            Exit Sub
        Else
            ExibeMensagem "Alteração permitida somente para Receitas com Saldo Inicial igual a zero."
            Exit Sub
        End If
    End If
    
    If UCase(strModoOperacao) = UCase(gstrAplicar) Then
        txtPKID = gstrItemData(dbcintCodigoOrcamentario)
        DoEvents
    End If
    
    If UCase(strModoOperacao) = gstrGeraCodigoReduzido Then
        If gblnGerouReduzidoReceita(Val(txtintExercicio)) Then
            VerificaListaAutomatica "", tdb_Lista, strQueryCO
        End If
    
    ElseIf UCase(strModoOperacao) = gstrSalvar Then
        If blnDadosOk Then
            ToolBarGeral strModoOperacao, gstrPrevisaoDaReceita, mblnAlterando, _
                         tdb_Lista, Me, mobjAux, strQueryCO, strQueryAplicar, _
                         rptPrevisaoDaReceita, strQueryRelatorio
            
            
            LeDaTabelaParaObj "", tdb_Lista, strQueryCO
            txt_intOrgao = ""
            txt_intModalidade.Text = ""
            txt_intConvenio.Text = ""
            txtdblValor = "0,00"
            txt_intEvento.SetFocus
            txt_total = tdb_Lista.Columns("dblTotal")
        End If
    
    ElseIf UCase(strModoOperacao) = gstrGeraFundef Then
        If MsgBox("Deseja atualizar os valores do FUNDEF?", vbYesNo) = vbYes Then
            CalculaGeraFundef
        End If
    
    ElseIf UCase(strModoOperacao) = UCase(gstrPreencherLista) Then
    
        If Me.ActiveControl.Name = cbo_intCodigoOrcamentario.Name Then
            LeCodigoOrcamentarioGeral cbo_intCodigoOrcamentario, dbcintCodigoOrcamentario, "SELECT * FROM " & gstrCodigoOrcamentario & " WHERE strCodigoOrcamentario LIKE '" & Trim(cbo_intCodigoOrcamentario.Text) & "%' AND intExercicio = " & IIf(blnOrcamento, gintExercicio, gintExercicio + 1)
        ElseIf Me.ActiveControl.Name = dbcintCodigoOrcamentario.Name Then
            LeCodigoOrcamentarioGeral cbo_intCodigoOrcamentario, dbcintCodigoOrcamentario, "SELECT * FROM " & gstrCodigoOrcamentario & " WHERE strDescricao LIKE '" & Trim(dbcintCodigoOrcamentario.Text) & "%' AND intExercicio = " & IIf(blnOrcamento, gintExercicio, gintExercicio + 1)
        ElseIf Me.ActiveControl.Name = dbcintEvento.Name Then
            LeDaTabelaParaObj "", dbcintEvento, strQueryEventoInicial
        ElseIf Me.ActiveControl.Name = dbcintFonteRecurso.Name Then
            PreencherListaDeOpcoes dbcintFonteRecurso
            'LeDaTabelaParaObj gstrFonteRecurso, dbcintFonteRecurso, "SELECT PKId, strDescricao FROM " & gstrFonteRecurso & " WHERE intExercicio = " & IIf(blnOrcamento, gintExercicio, gintExercicio + 1)
        ElseIf Me.ActiveControl.Name = dbcintModalidade.Name Then
            PreencherListaDeOpcoes dbcintModalidade
        ElseIf Me.ActiveControl.Name = dbcintOrgao.Name Then
            PreencherListaDeOpcoes dbcintOrgao
        End If
        
    Else
    
        ToolBarGeral strModoOperacao, gstrPrevisaoDaReceita, mblnAlterando, _
                     tdb_Lista, Me, mobjAux, strQueryCO, strQueryAplicar, _
                     rptPrevisaoDaReceita, strQueryRelatorio
        
        If UCase(strModoOperacao) = gstrDeletar Then
            txt_intConvenio.Text = ""
            txt_intModalidade.Text = ""
            txt_intOrgao.Text = ""
        End If
        If UCase(strModoOperacao) = gstrSalvar Then
            VerificaListaAutomatica "", tdb_Lista, strQueryCO
        End If
          
        VerificaObjParaAplicar mobjAux, mstrQueryAplicar
        
        If UCase(strModoOperacao) = UCase(gstrNovo) Then
            If blnOrcamento Then
                LimpaObjetos
                txtintCodigoReduzido.SetFocus
            'ORC1517 - FERNANDO
            Else
                txt_intOrgao.Text = ""
                txt_intModalidade.Text = ""
                txt_intConvenio.Text = ""
            'ORC1517 - FERNANDO
            End If
        End If
       
    End If
    
End Sub

Private Sub tdb_Lista_Scroll(Cancel As Integer)
    EditaValor tdb_Lista.Columns("PkId").Value, 0, tdb_Lista.Row
End Sub

Private Sub txt_dblValor_GotFocus()
    MarcaCampo txt_dblValor
End Sub

Private Sub txt_DBLVALOR_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        EditaValor tdb_Lista.Columns("PkId").Value, 0, tdb_Lista.Row
        tdb_Lista.SetFocus
    Else
        CaracterValido KeyAscii, "V", txt_dblValor
    End If
End Sub

Private Sub txt_intConvenio_GotFocus()
    MarcaCampo txt_intConvenio
End Sub

Private Sub txt_intConvenio_LostFocus()
    If Len(Trim(txt_intModalidade.Text)) > 0 Then
        PreencherListaDeOpcoes dbcintModalidade, CarregaModalidade(txt_intModalidade.Text, txt_intConvenio.Text)
    End If
End Sub

Private Sub txt_intEvento_KeyPress(KeyAscii As Integer)
   'CaracterValido KeyAscii, "N", txt_intEvento
   gstrLimitaCampoValor txt_intEvento, KeyAscii, 10, 0
End Sub

Private Sub txt_intEvento_LostFocus()
    PreencheEventobyCodigo txt_intEvento, dbcintEvento, "0"
End Sub
'Private Sub txt_intOrgao_Change()
''ORC1543 - FERNANDO
''    If Len(Trim$(txt_intOrgao)) = 0 Then
''        dbcintOrgao.Text = ""
''        Set dbcintOrgao.RowSource = Nothing
''    End If
''ORC1543 - FERNANDO
'End Sub

Private Sub txt_intModalidade_GotFocus()
    MarcaCampo txt_intModalidade
End Sub

Private Sub txt_intModalidade_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txt_intModalidade
End Sub

Private Sub txt_intOrgao_GotFocus()
    MarcaCampo txt_intOrgao
End Sub

Private Sub txt_intOrgao_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txt_intOrgao
End Sub

Private Sub txt_intOrgao_LostFocus()
'ORC1543 - FERNANDO
'    If Len(txt_intOrgao) > 0 Then
'        If strCodAnt <> Trim$(txt_intOrgao) Then
'            PreencheOrgao txt_intOrgao, 0
'        End If
'    End If
'    strCodAnt = txt_intOrgao
'ORC1543 - FERNANDO
    Dim strPKId As String
    
    If Trim(txt_intOrgao.Text) = "" Then Exit Sub
    
    strPKId = LeCodOrgao(, txt_intOrgao)
    If Trim(strPKId) = "" Then
    txt_intOrgao.Text = ""
    dbcintOrgao.BoundText = ""
    txt_intOrgao.SetFocus
    Exit Sub
    End If
    LeDaTabelaParaObj gstrOrgao, dbcintOrgao, "SELECT PKID, strDescricao FROM " & gstrOrgao & " WHERE INTEXERCICIO =  " & IIf(blnOrcamento, gintExercicio, gintExercicio + 1) & " ORDER BY STRDESCRICAO ", strPKId
    dbcintOrgao.BoundText = strPKId
        
End Sub

Private Sub txtintCodigoReduzido_GotFocus()
   
   gstrProximoCodigo txtintCodigoReduzido, gstrPrevisaoDaReceita, "intCodigoReduzido", gintCodSeguranca, "intExercicio", Val(gintExercicio), , , , , "intExercicio", Val(gintExercicio)
   MarcaCampo txtintCodigoReduzido
   
End Sub

Private Sub txtintCodigoReduzido_KeyPress(KeyAscii As Integer)
   'CaracterValido KeyAscii, "N"
   gstrLimitaCampoValor txtintCodigoReduzido, KeyAscii, 9, 0
End Sub

Private Sub txtintCodigoReduzido_LostFocus()
   Dim adoResultado As New ADODB.Recordset
   Dim strSQL       As String
   
   If Len(Trim(txtintCodigoReduzido)) > 0 Then
      strSQL = "SELECT * FROM " & gstrPrevisaoDaReceita
      strSQL = strSQL & " WHERE intCodigoReduzido = " & Trim(txtintCodigoReduzido)
      strSQL = strSQL & " AND intExercicio = " & txtintExercicio
      Set gobjBanco = New clsBanco
      If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
         With adoResultado
            If .EOF = False Then
               
               txtPKID = (!Pkid)
               cbo_intCodigoOrcamentario.ListIndex = gintIndiceCBO(cbo_intCodigoOrcamentario, !INTCODIGOORCAMENTARIO)
               dbcintCodigoOrcamentario.ListIndex = gintIndiceCBO(dbcintCodigoOrcamentario, !INTCODIGOORCAMENTARIO)
               dbcintEvento.ListIndex = gintIndiceCBO(dbcintEvento, !intEvento)
               PreencherListaDeOpcoes dbcintFonteRecurso, gstrVerificaCampoNulo(!INTFONTERECURSO)
               txtstrLegislacao = gstrVerificaCampoNulo(!strLegislacao)
               txtdblValor = gstrConvVrDoSql(!dblValor)
               If Len(Trim(txtdblValor)) = 0 Then
                  txtdblValor = "0,00"
               End If
               
               mblnAlterando = True
            End If
         End With
      End If
   End If
End Sub

Private Sub txtPKId_Change()
    txtPKID = txtPKID
End Sub

Private Sub txtdblValor_GotFocus()
    MarcaCampo txt_total
End Sub

Private Sub txtdblValor_KeyPress(KeyAscii As Integer)
    'CaracterValido KeyAscii, "V", txtdblValor
    gstrLimitaCampoValor txtdblValor, KeyAscii, 9, 2
End Sub

Private Sub txtdblValor_LostFocus()
    txtdblValor = gstrConvVrDoSql(txtdblValor)
End Sub

Public Function strQueryRelatorio()

'******************************************************************************************
' Data: 09/06/2003
' Alteração: - Substituição do comando nativo de outer join (*= ou =*) do SQL Server pela
'            variável strOUTJSQLServer e inclusão do comando de outer join ((+)) do Oracle,
'            representado pela variável strOUTJOracle.
' Responsável: Everton Bianchini
'******************************************************************************************

    Dim strSQL  As String
    strSQL = ""
    strSQL = strSQL & "SELECT PR.intCodigoreduzido AS CodigoReduzido, "
    strSQL = strSQL & "CO.strCodigoOrcamentario, "
    strSQL = strSQL & "CO.strDescricao AS DesCodigoOrcamentario, "
    strSQL = strSQL & "FR.strDescricao AS FonteRecurso, PR.dblValor, "
    strSQL = strSQL & "MO.strDescricao AS ModalidadeAplicacao "
    strSQL = strSQL & "FROM "
    strSQL = strSQL & gstrPrevisaoDaReceita & " PR, "
    strSQL = strSQL & gstrCodigoOrcamentario & " CO, "
    strSQL = strSQL & gstrFonteRecurso & " FR, "
    strSQL = strSQL & gstrModalidade & " MO "
'    strSql = strSql & "WHERE PR.intCodigoOrcamentario *= CO.PKId "
    strSQL = strSQL & "WHERE PR.intCodigoOrcamentario " & strOUTJSQLServer & "= CO.PKId " & strOUTJOracle
'    strSql = strSql & "AND PR.intFonteRecurso *= FR.PKId "
    strSQL = strSQL & "AND PR.intFonteRecurso " & strOUTJSQLServer & "= FR.PKId " & strOUTJOracle & " AND PR.intExercicio = " & txtintExercicio
    strSQL = strSQL & " AND PR.intModalidade " & strOUTJSQLServer & "= MO.PKID " & strOUTJOracle & ""
    strSQL = strSQL & " ORDER BY CO.strCodigoOrcamentario"
    strQueryRelatorio = strSQL
End Function

Private Function blnDadosOk() As Boolean
    
    If dbcintEvento.ListIndex = -1 Then
       ExibeMensagem "É necessário informar o Evento Contábil."
       dbcintEvento.SetFocus
       Exit Function
    ElseIf cbo_intCodigoOrcamentario.ListIndex = -1 Then
        ExibeMensagem "O código orçamentário é inexistente neste exercício." & vbCrLf & "Informe um código válido."
        cbo_intCodigoOrcamentario.SetFocus
        Exit Function
    ElseIf dbcintCodigoOrcamentario.ListIndex = -1 Then
        ExibeMensagem "A descricao do código orçamentário não é válida."
        dbcintCodigoOrcamentario.SetFocus
        Exit Function
    ElseIf Not dbcintFonteRecurso.MatchedWithList Then
        ExibeMensagem "Informe a fonte de recurso."
        dbcintFonteRecurso.SetFocus
        Exit Function
    ElseIf Not dbcintModalidade.MatchedWithList Then
        ExibeMensagem "Informe a Modalidade."
        dbcintModalidade.SetFocus
        Exit Function
    ElseIf txt_intOrgao = "" Or dbcintOrgao.Text = "" Then
        ExibeMensagem "É necessário informar um Orgão."
        txt_intOrgao.SetFocus
        Exit Function
    ElseIf Len(Trim(txtdblValor.Text)) = 0 Then
        ExibeMensagem "Informe o valor."
        If txtdblValor.Enabled Then txtdblValor.SetFocus
        Exit Function
    ElseIf Not blnVerificaCodigo Then
        ExibeMensagem "O código já está cadastrado para essa fonte de recursos/orgão/modalidade."
        cbo_intCodigoOrcamentario.SetFocus
        Exit Function
    
'    ElseIf Not VerificaEvento Then
'       ExibeMensagem "Este Evento Contábil não é compatível com o Código Orçamentário."
'       txt_intEvento.Text = ""
'       dbcintEvento.ListIndex = -1
'       txt_intEvento.SetFocus
'       Exit Function
    Else
        blnDadosOk = True
    End If
    
End Function

Private Sub AtualizaCodigoOrcamentario(cboCodigo As ComboBox, Optional strQuery As String)
    Dim strSQL          As String
    Dim adoResultado    As ADODB.Recordset
    
    cboCodigo.Clear
    If Trim(strQuery) = "" Then
        strSQL = ""
        strSQL = strSQL & " SELECT CO.PKId, CO.strCodigoOrcamentario"
        strSQL = strSQL & " FROM "
        strSQL = strSQL & gstrCodigoOrcamentario & " CO "
        If blnOrcamento = True Then
           strSQL = strSQL & " WHERE intExercicio = " & gintExercicio
        Else
           strSQL = strSQL & " WHERE intExercicio = " & gintExercicio + 1
        End If
    Else
        strSQL = strQuery
    End If
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
        With adoResultado
            Do While .EOF = False
                cboCodigo.AddItem gvntFormatacaoEspecifica(!strCodigoOrcamentario)
                cboCodigo.ItemData(cboCodigo.NewIndex) = !Pkid
                .MoveNext
            Loop
        End With
    End If
    
End Sub

Private Sub EditaValor(Pkid As Variant, Col As Integer, Row As Variant)
Dim adoResultado As ADODB.Recordset

On Error GoTo TrataErroLocal

PosGravaValor:

    DoEvents
    
    'Vamos tornar o campo valor editavel
    If Col = 4 And Not txt_dblValor.Visible And Not tdb_Lista.FilterActive And Not blnOrcamento Then
        
        'Vamos passar os atributos a caixa de texto de edicao
        txt_dblValor.Width = tdb_Lista.Columns("Valor").Width
        txt_dblValor.Height = tdb_Lista.RowHeight
        txt_dblValor.Top = tdb_Lista.Top + tdb_Lista.RowTop(Row)
        txt_dblValor.Left = tdb_Lista.Left + tdb_Lista.Columns("Valor").Left
        
        txt_dblValor.Text = gstrConvVrDoSql(tdb_Lista.Columns("Valor").Value)
        
        txt_dblValor.Visible = True
        
        txt_dblValor.SetFocus
        
        dblValorAnt = tdb_Lista.Columns("Valor").Value 'Alfred 18/07/2003

    ElseIf txt_dblValor.Visible Then
        
        txt_dblValor.Visible = False
        
        If Val(gstrConvVrParaSql(txt_dblValor.Text)) <> dblValorAnt Then 'Alfred 18/07/2003
        
'            If MsgBox("Deseja Realmente alterar o valor da Proposta Orçamentária do Exercício de " & IIf(blnOrcamento, gintExercicio, gintExercicio + 1) & "?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
'                Exit Sub
'            End If
        
            'Vamos atualizar o valor na tabela de nprograma de trabalho
            Set gobjBanco = New clsBanco
            gobjBanco.ExecutaBeginTrans
            gobjBanco.Execute "UPDATE " & gstrPrevisaoDaReceita & " SET dblValor = " & gstrConvVrParaSql(txt_dblValor.Text) & " , " & _
                              "dtmDtAtualizacao = " & gstrConvDtParaSql(gstrDataDoSistema) & ", " & _
                              "lngCodUsr = " & glngCodUsr & _
                              " WHERE PkId = " & intPkIdRow
            gobjBanco.ExecutaCommitTrans
            Set gobjBanco = Nothing
            
            LeDaTabelaParaObj "", tdb_Lista, strQueryCO
            
        End If
        
        txt_total = tdb_Lista.Columns("dblTotal")
        txtdblValor = txt_dblValor
        
        If intPkIdRow > 0 Then
            Set adoResultado = tdb_Lista.DataSource
            adoResultado.Find "PKId = '" & Pkid & "'"
            tdb_Lista.MarqueeStyle = dbgHighlightRow
        End If
                        
        GoTo PosGravaValor
        
    End If
    
Exit Sub

TrataErroLocal:
    Resume Next
    
End Sub

Private Sub CalculaGeraFundef()
Dim strSQL      As String
Dim adoRec      As ADODB.Recordset
Dim blnGerou    As Boolean
Dim ScrMouse    As Integer
    
    ScrMouse = Screen.MousePointer
    Screen.MousePointer = vbHourglass
    
    DoEvents
    If blnOrcamento = True Then
       strSQL = "UPDATE " & gstrPrevisaoDaReceita & " SET dblValor = 0 " & _
                " WHERE intCodigoOrcamentario IN " & _
                " (SELECT CO.PKId FROM " & gstrCodigoOrcamentario & " CO, " & gstrGerarFundef & " GF" & _
                " WHERE CO.strCodigoOrcamentario = GF.strCodigoOrcDestino AND GF.intExercicio = CO.intExercicio AND GF.intExercicio =" & gintExercicio & ")"
    Else
       strSQL = "UPDATE " & gstrPrevisaoDaReceita & " SET dblValor = 0 " & _
                " WHERE intCodigoOrcamentario IN " & _
                " (SELECT CO.PKId FROM " & gstrCodigoOrcamentario & " CO, " & gstrGerarFundef & " GF" & _
                " WHERE CO.strCodigoOrcamentario = GF.strCodigoOrcDestino AND GF.intExercicio = CO.intExercicio AND GF.intExercicio =" & gintExercicio + 1 & ")"
    End If
    Set gobjBanco = New clsBanco
    gobjBanco.Execute strSQL
    
    strSQL = "SELECT " & _
            " GF.PKId, " & _
            " CD.PKId CodigoDestino," & _
            " (PR.dblValor * GF.intPorcentagem / 100) PercentualSoma"
    strSQL = strSQL & " FROM " & _
            gstrGerarFundef & " GF, " & _
            gstrCodigoOrcamentario & " CD, " & _
            gstrCodigoOrcamentario & " CO, " & _
            gstrPrevisaoDaReceita & " PR"
    strSQL = strSQL & " WHERE" & _
            " GF.strCodigoOrcDestino = CD.strCodigoOrcamentario AND" & _
            " GF.strCodigoOrcOrigem = CO.strCodigoOrcamentario AND" & _
            " PR.intCodigoOrcamentario = CO.PKId AND" & _
            " CO.intExercicio = GF.intExercicio AND" & _
            " CD.intExercicio = GF.intExercicio AND"
            If blnOrcamento = True Then
               strSQL = strSQL & " GF.intExercicio = " & gintExercicio
            Else
               strSQL = strSQL & " GF.intExercicio = " & gintExercicio + 1
            End If

    If gobjBanco.CriaADO(strSQL, 30, adoRec) Then
        Do While Not adoRec.EOF
            If blnOrcamento = True Then
               strSQL = "UPDATE " & gstrPrevisaoDaReceita & " SET dblValor = dblValor + " & gstrConvVrParaSql(adoRec!PercentualSoma * -1) & _
                        " WHERE intExercicio=" & gintExercicio & " AND intCodigoOrcamentario=" & adoRec!CodigoDestino
            Else
               strSQL = "UPDATE " & gstrPrevisaoDaReceita & " SET dblValor = dblValor + " & gstrConvVrParaSql(adoRec!PercentualSoma * -1) & _
                        " WHERE intExercicio=" & gintExercicio + 1 & " AND intCodigoOrcamentario=" & adoRec!CodigoDestino
            End If
            
            gobjBanco.Execute strSQL
            adoRec.MoveNext
            blnGerou = True
        Loop

    End If
    
    If blnGerou = True Then ExibeMensagem "Registros atualizados com sucesso.", 64
    
    Screen.MousePointer = ScrMouse
    
End Sub


Private Sub LimpaObjetos()
   
   mblnAlterando = False
   cbo_intCodigoOrcamentario.Text = ""
   dbcintCodigoOrcamentario.Text = ""
   dbcintFonteRecurso.Text = ""
   txtstrLegislacao.Text = ""
   txt_intEvento.Text = ""
   dbcintEvento.Text = ""
   txt_intOrgao = ""
   dbcintOrgao = ""
   txtdblValor.Text = "0,00"
   txt_intModalidade.Text = ""
   txt_intConvenio.Text = ""
End Sub

Private Function VerificaEvento() As Boolean
   Dim strSQL       As String
   Dim adoResultado As New ADODB.Recordset
   
   VerificaEvento = True
   
   strSQL = "SELECT EV.PKID, EV.strDescricao, EVD.intContaContabil, PC.strContaContabil"
   strSQL = strSQL & " FROM " & gstrEvento & " EV, " & gstrEventoContaContabilDebito & " EVD,"
   strSQL = strSQL & gstrPlanoConta & " PC"
   strSQL = strSQL & " WHERE  EV.Pkid = EVD.intEvento AND EV.PKID = " & gstrItemData(dbcintEvento) & " AND "
   strSQL = strSQL & " EVD.intContaContabil = PC.PKid AND "
   strSQL = strSQL & strSUBSTRING & "(PC.strContaContabil,1,1) = '4' AND "
   strSQL = strSQL & strSUBSTRING & "(PC.strContaContabil,2,1) = '" & Mid(Trim(cbo_intCodigoOrcamentario.Text), 1, 1) & "' AND "
   strSQL = strSQL & " EV.intTipoEvento = 0"
   
   Set gobjBanco = New clsBanco
   If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
      With adoResultado
         If .EOF = True Then
            VerificaEvento = False
         End If
     End With
   End If
End Function

Private Function strQueryEventoInicial() As String
Dim strSQL As String
   
   strSQL = "SELECT EV.PKID, EV.strDescricao, EVD.intContaContabil, PC.strContaContabil"
   strSQL = strSQL & " FROM " & gstrEvento & " EV, " & gstrEventoContaContabilDebito & " EVD,"
   strSQL = strSQL & gstrPlanoConta & " PC"
   strSQL = strSQL & " WHERE EV.Pkid = EVD.intEvento AND "
   strSQL = strSQL & " EVD.intContaContabil = PC.PKid AND "
   If blnOrcamento Then
        If gintExercicio <= 2006 Then
            strSQL = strSQL & strSUBSTRING & "(PC.strContaContabil,1," & Len(gstrDigitoReceita) & ") = '" & gstrDigitoReceita & "' AND "
        End If
   Else
        If gintExercicio + 1 <= 2006 Then
            strSQL = strSQL & strSUBSTRING & "(PC.strContaContabil,1," & Len(gstrDigitoReceita) & ") = '" & gstrDigitoReceita & "' AND "
        End If
   End If
   strSQL = strSQL & " EV.intTipoEvento = 0"
   
   strQueryEventoInicial = strSQL
   
End Function

Private Sub txtstrLegislacao_GotFocus()
    MarcaCampo txtstrLegislacao
End Sub

Private Sub txtstrLegislacao_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrLegislacao
End Sub

Private Function PreencheOrgao(intCodigo As Double, Optional bitcampo As Integer) As String
Dim strSQL As String
Dim adoResultado As New ADODB.Recordset


'QUANDO O BITCAMPO FOR IGUAL A 0, O CAMPO DESCRICAO DA COMBO SERA PREENCHIDO
' "     "      "     "   "   " 1, "   "   INTCODIGO  DO TEXT    "     "
    
    
    strSQL = "SELECT PKID, "
    strSQL = strSQL & " STRCODIGO, "
    strSQL = strSQL & " STRDESCRICAO"
    strSQL = strSQL & " FROM "
    strSQL = strSQL & gstrOrgao
    strSQL = strSQL & " WHERE "
    If bitcampo = 0 Then
        strSQL = strSQL & " strcodigo = " & "'" & intCodigo & "'"
    Else
        strSQL = strSQL & " pkid = " & intCodigo
    End If

    Set gobjBanco = New clsBanco
    
    If gobjBanco.CriaADO(strSQL, 30, adoResultado) Then
         
       If adoResultado.RecordCount > 0 Then
         With adoResultado
            If .EOF = False Then
               
               If bitcampo = 1 Then
                    txt_intOrgao = (!strCodigo)
               Else
                    PreencherListaDeOpcoes dbcintOrgao, (!Pkid)
                    
                End If
             
             End If
         End With
       Else
            txt_intOrgao = ""
            txt_intOrgao.SetFocus
       End If
    End If

End Function

Private Function LeCodOrgao(Optional strPKId As String, Optional strCod As String)
    
    Dim strSQL          As String
    Dim adoResultado    As ADODB.Recordset
    
    If Trim(strPKId) = "" And Trim(strCod) = "" Then
        LeCodOrgao = ""
        Exit Function
    End If
    
    strSQL = ""
    strSQL = strSQL & "SELECT "
    strSQL = strSQL & " OG.strCodigo ,"
    strSQL = strSQL & " OG.PKID"
    strSQL = strSQL & " FROM    "
    strSQL = strSQL & gstrOrgao & " OG"
    strSQL = strSQL & " Where "
    
    If strPKId <> "" Then
        strSQL = strSQL & " OG.PKID = " & strPKId
    ElseIf strCod <> "" Then
        strSQL = strSQL & " OG.strCodigo = '" & strCod & "'"
    End If
    
    strSQL = strSQL & " AND OG.intExercicio = " & IIf(blnOrcamento, gintExercicio, gintExercicio + 1)
    strSQL = strSQL & " ORDER BY OG.STRDESCRICAO "
    
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
    
        If Not adoResultado.EOF Then
            If strPKId <> "" Then
                LeCodOrgao = gstrENulo(adoResultado!strCodigo)
            ElseIf strCod <> "" Then
                LeCodOrgao = gstrENulo(adoResultado!Pkid)
            End If
        End If
        
    End If
    
End Function

Private Function CarregaModalidade(strCodModalidade As String, strCodConvenio As String) As String
    Dim strSQL As String
    Dim adoResultado    As ADODB.Recordset
    strSQL = ""
    strSQL = "SELECT MO.PKID "
    strSQL = strSQL & " FROM " & gstrModalidade & " MO, "
    strSQL = strSQL & gstrConvenio & " CV"
    strSQL = strSQL & " WHERE " & gstrCONVERT(CDT_INT, "MO.strCodigo") & " = '" & Val(strCodModalidade) & "'"
    If Len(Trim(txt_intConvenio.Text)) > 0 Then
        If txt_intConvenio.Text <> "00" Then
            strSQL = strSQL & " AND " & gstrCONVERT(CDT_INT, "CV.strCodigo") & " = '" & Val(strCodConvenio) & "'"
            strSQL = strSQL & " AND MO.intConvenio = CV.PKID"
        End If
    End If
    strSQL = strSQL & " GROUP BY MO.PKID"
    Set gobjBanco = New clsBanco
    
    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
        If adoResultado.RecordCount > 0 Then
            CarregaModalidade = gstrENulo(adoResultado.Fields("PKID").Value)
        End If
    End If

End Function

Private Sub CarregaCodModalidade()
    Dim strSQL As String
    Dim adoResultado    As ADODB.Recordset
    strSQL = ""
    strSQL = strSQL & "SELECT  MO.strCodigo as CodModalidade, "
    strSQL = strSQL & " CV.strCodigo as CodConvenio "
    strSQL = strSQL & " FROM " & gstrModalidade & " MO, "
    strSQL = strSQL & gstrConvenio & " CV "
    strSQL = strSQL & " WHERE MO.PKID = " & dbcintModalidade.BoundText
    strSQL = strSQL & " AND MO.intConvenio " & strOUTJSQLServer & "= CV.PKID" & strOUTJOracle
    Set gobjBanco = New clsBanco
    
    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
        If adoResultado.RecordCount > 0 Then
            txt_intModalidade.Text = Format(gstrENulo(adoResultado.Fields("CodModalidade").Value), "000")
            If Len(Trim(gstrENulo(adoResultado.Fields("CodConvenio").Value))) > 0 Then
                txt_intConvenio.Text = Format(gstrENulo(adoResultado.Fields("CodConvenio").Value), "00")
            Else
                txt_intConvenio.Text = "00"
            End If
        End If
    End If
End Sub

Private Function blnVerificaCodigo() As Boolean
    Dim strSQL As String
    Dim adoResultado As ADODB.Recordset
    
    blnVerificaCodigo = True
    strSQL = ""
    strSQL = " SELECT  PKID "
    strSQL = strSQL & " From " & gstrPrevisaoDaReceita
    strSQL = strSQL & " WHERE   intCodigoOrcamentario = " & dbcintCodigoOrcamentario.ItemData(dbcintCodigoOrcamentario.ListIndex)
    strSQL = strSQL & " AND intFonteRecurso = " & dbcintFonteRecurso.BoundText
    strSQL = strSQL & " AND intOrgao = " & dbcintOrgao.BoundText
    strSQL = strSQL & " AND intModalidade = " & dbcintModalidade.BoundText
    strSQL = strSQL & " AND intExercicio = " & IIf(blnOrcamento, gintExercicio, gintExercicio + 1)
    Set gobjBanco = New clsBanco
    
    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
        If adoResultado.RecordCount > 0 Then
            With adoResultado
            If mblnAlterando Then
                While Not .EOF
                    If .Fields("PKID").Value = txtPKID.Text Then
                        blnVerificaCodigo = True
                        Exit Function
                    Else
                        blnVerificaCodigo = False
                    End If
                    .MoveNext
                Wend
            Else
                blnVerificaCodigo = False
            End If
            End With
        End If
    End If
        
End Function
