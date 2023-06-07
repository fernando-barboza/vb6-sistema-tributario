VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "TODG7.OCX"
Begin VB.Form frmCadFormulaDeCalculos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fórmulas de Cálculo"
   ClientHeight    =   5730
   ClientLeft      =   1665
   ClientTop       =   3930
   ClientWidth     =   9000
   HelpContextID   =   41
   Icon            =   "frmCadFormulaDeCalculos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   9000
   Begin VB.TextBox txtPKId 
      Height          =   270
      Left            =   5340
      TabIndex        =   6
      Top             =   90
      Visible         =   0   'False
      Width           =   645
   End
   Begin TabDlg.SSTab tab_3DPasta 
      Height          =   5565
      Left            =   75
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   90
      Width           =   8835
      _ExtentX        =   15584
      _ExtentY        =   9816
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Fórmulas de Cálculo"
      TabPicture(0)   =   "frmCadFormulaDeCalculos.frx":1042
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lbldblCodigo"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblstrNome"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblintCidade"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblintReceita"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblstrDescricao"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "dbcintCidade"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "dbcintReceita"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtdblCodigo"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtstrNome"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cmd_intReceita"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtstrDescricao"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "tdb_FormulaCalculo"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).ControlCount=   12
      TabCaption(1)   =   "Editor de Fórmula"
      TabPicture(1)   =   "frmCadFormulaDeCalculos.frx":105E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "dlg_BuscaFormula"
      Tab(1).Control(1)=   "cmd_BuscaFormula"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "txtstrFormula"
      Tab(1).Control(3)=   "fra_PalavraChave"
      Tab(1).Control(4)=   "fra_Agregacao"
      Tab(1).Control(5)=   "fra_aux"
      Tab(1).ControlCount=   6
      Begin MSComDlg.CommonDialog dlg_BuscaFormula 
         Left            =   -66945
         Top             =   5010
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton cmd_BuscaFormula 
         Height          =   330
         Left            =   -66690
         Picture         =   "frmCadFormulaDeCalculos.frx":107A
         Style           =   1  'Graphical
         TabIndex        =   18
         TabStop         =   0   'False
         ToolTipText     =   "Clique para buscar Fórmula"
         Top             =   3795
         Width           =   360
      End
      Begin VB.TextBox txtstrFormula 
         Height          =   3330
         Left            =   -74865
         MaxLength       =   4000
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   17
         Top             =   405
         Width           =   8535
      End
      Begin VB.Frame fra_PalavraChave 
         Caption         =   "Palavras Chaves"
         Height          =   1335
         Left            =   -74865
         TabIndex        =   16
         Top             =   4080
         Width           =   2775
         Begin VB.CommandButton cmd_Create 
            Caption         =   "Create"
            Height          =   315
            Left            =   1800
            TabIndex        =   27
            Top             =   960
            Width           =   855
         End
         Begin VB.CommandButton cmd_Group 
            Caption         =   "Group by"
            Height          =   315
            Left            =   960
            TabIndex        =   26
            Top             =   960
            Width           =   855
         End
         Begin VB.CommandButton cmd_Order 
            Caption         =   "Order By"
            Height          =   315
            Left            =   120
            TabIndex        =   25
            Top             =   960
            Width           =   855
         End
         Begin VB.CommandButton cmd_From 
            Caption         =   "From"
            Height          =   315
            Left            =   960
            TabIndex        =   23
            Top             =   600
            Width           =   855
         End
         Begin VB.CommandButton cmd_Where 
            Caption         =   "Where"
            Height          =   315
            Left            =   1800
            TabIndex        =   24
            Top             =   600
            Width           =   855
         End
         Begin VB.CommandButton cmd_Update 
            Caption         =   "Update"
            Height          =   315
            Left            =   1800
            TabIndex        =   21
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton cmd_Delete 
            Caption         =   "Delete"
            Height          =   315
            Left            =   120
            TabIndex        =   22
            Top             =   600
            Width           =   855
         End
         Begin VB.CommandButton cmd_Insert 
            Caption         =   "Insert"
            Height          =   315
            Left            =   960
            TabIndex        =   20
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton cmd_Select 
            Caption         =   "Select"
            Height          =   315
            Left            =   120
            TabIndex        =   19
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.Frame fra_Agregacao 
         Caption         =   "Agregação"
         Height          =   1335
         Left            =   -71985
         TabIndex        =   15
         Top             =   4080
         Width           =   2775
         Begin VB.CommandButton cmd_Min 
            Caption         =   "MIN"
            Height          =   315
            Left            =   960
            TabIndex        =   32
            Top             =   600
            Width           =   855
         End
         Begin VB.CommandButton cmd_Max 
            Caption         =   "MAX"
            Height          =   315
            Left            =   120
            TabIndex        =   31
            Top             =   600
            Width           =   855
         End
         Begin VB.CommandButton cmd_AVG 
            Caption         =   "AVG"
            Height          =   315
            Left            =   1800
            TabIndex        =   30
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton cmd_Count 
            Caption         =   "COUNT"
            Height          =   315
            Left            =   960
            TabIndex        =   29
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton cmd_Sum 
            Caption         =   "SUM"
            Height          =   315
            Left            =   120
            TabIndex        =   28
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.Frame fra_aux 
         Caption         =   "Auxiliares"
         Height          =   1335
         Left            =   -69105
         TabIndex        =   14
         Top             =   4080
         Width           =   2775
         Begin VB.CommandButton cmd_Porcentagem 
            Caption         =   "%"
            Height          =   315
            Left            =   1800
            TabIndex        =   38
            Top             =   600
            Width           =   855
         End
         Begin VB.CommandButton cmd_Mod 
            Caption         =   "Mod"
            Height          =   315
            Left            =   960
            TabIndex        =   37
            Top             =   600
            Width           =   855
         End
         Begin VB.CommandButton cmd_Divisao 
            Caption         =   "/"
            Height          =   315
            Left            =   120
            TabIndex        =   36
            Top             =   600
            Width           =   855
         End
         Begin VB.CommandButton cmd_Multiplicacao 
            Caption         =   "*"
            Height          =   315
            Left            =   1800
            TabIndex        =   35
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton cmd_menos 
            Caption         =   "-"
            Height          =   315
            Left            =   960
            TabIndex        =   34
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton cmd_Mais 
            Caption         =   "+"
            Height          =   315
            Left            =   120
            TabIndex        =   33
            Top             =   240
            Width           =   855
         End
      End
      Begin TrueOleDBGrid70.TDBGrid tdb_FormulaCalculo 
         Height          =   2715
         Left            =   105
         TabIndex        =   5
         Top             =   2730
         Width           =   7605
         _ExtentX        =   13414
         _ExtentY        =   4789
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
         Columns(1).Caption=   "Código"
         Columns(1).DataField=   "dblCodigo"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Nome"
         Columns(2).DataField=   "strNome"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Descrição"
         Columns(3).DataField=   "strDescricao"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   4
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
         Splits(0)._ColumnProps(0)=   "Columns.Count=4"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0).AllowSizing=0"
         Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
         Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(8)=   "Column(1).Width=1270"
         Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=1191"
         Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(13)=   "Column(2).Width=5715"
         Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=5636"
         Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(17)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(18)=   "Column(3).Width=5847"
         Splits(0)._ColumnProps(19)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(20)=   "Column(3)._WidthInPix=5768"
         Splits(0)._ColumnProps(21)=   "Column(3)._EditAlways=0"
         Splits(0)._ColumnProps(22)=   "Column(3).Order=4"
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
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=144,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
         _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
         _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34"
         _StyleDefs(11)  =   "FooterStyle:id=3,.parent=1,.namedParent=35"
         _StyleDefs(12)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(13)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
         _StyleDefs(14)  =   "EditorStyle:id=7,.parent=1"
         _StyleDefs(15)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
         _StyleDefs(16)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
         _StyleDefs(17)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
         _StyleDefs(18)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(19)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
         _StyleDefs(20)  =   "Splits(0).Style:id=13,.parent=1"
         _StyleDefs(21)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
         _StyleDefs(22)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
         _StyleDefs(23)  =   "Splits(0).FooterStyle:id=15,.parent=3"
         _StyleDefs(24)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
         _StyleDefs(25)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
         _StyleDefs(26)  =   "Splits(0).EditorStyle:id=17,.parent=7"
         _StyleDefs(27)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
         _StyleDefs(28)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
         _StyleDefs(29)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
         _StyleDefs(30)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(31)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(32)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
         _StyleDefs(33)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
         _StyleDefs(34)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
         _StyleDefs(35)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
         _StyleDefs(36)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
         _StyleDefs(37)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
         _StyleDefs(38)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
         _StyleDefs(39)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
         _StyleDefs(40)  =   "Splits(0).Columns(2).Style:id=46,.parent=13"
         _StyleDefs(41)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
         _StyleDefs(44)  =   "Splits(0).Columns(3).Style:id=50,.parent=13"
         _StyleDefs(45)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
         _StyleDefs(46)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
         _StyleDefs(47)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
         _StyleDefs(48)  =   "Named:id=33:Normal"
         _StyleDefs(49)  =   ":id=33,.parent=0"
         _StyleDefs(50)  =   "Named:id=34:Heading"
         _StyleDefs(51)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(52)  =   ":id=34,.wraptext=-1"
         _StyleDefs(53)  =   "Named:id=35:Footing"
         _StyleDefs(54)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(55)  =   "Named:id=36:Selected"
         _StyleDefs(56)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(57)  =   "Named:id=37:Caption"
         _StyleDefs(58)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(59)  =   "Named:id=38:HighlightRow"
         _StyleDefs(60)  =   ":id=38,.parent=33,.bgcolor=&H80000014&,.fgcolor=&H80000012&"
         _StyleDefs(61)  =   "Named:id=39:EvenRow"
         _StyleDefs(62)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(63)  =   "Named:id=40:OddRow"
         _StyleDefs(64)  =   ":id=40,.parent=33"
         _StyleDefs(65)  =   "Named:id=41:RecordSelector"
         _StyleDefs(66)  =   ":id=41,.parent=34"
         _StyleDefs(67)  =   "Named:id=42:FilterBar"
         _StyleDefs(68)  =   ":id=42,.parent=33"
      End
      Begin VB.TextBox txtstrDescricao 
         Height          =   780
         Left            =   1065
         MaxLength       =   4000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   1875
         Width           =   6630
      End
      Begin VB.CommandButton cmd_intReceita 
         Height          =   315
         Left            =   5415
         Picture         =   "frmCadFormulaDeCalculos.frx":1198
         Style           =   1  'Graphical
         TabIndex        =   12
         TabStop         =   0   'False
         Tag             =   "596"
         ToolTipText     =   "Ativa Cadastro de Receitas"
         Top             =   780
         Width           =   360
      End
      Begin VB.TextBox txtstrNome 
         Height          =   285
         Left            =   1065
         MaxLength       =   100
         TabIndex        =   3
         Top             =   1515
         Width           =   6630
      End
      Begin VB.TextBox txtdblCodigo 
         Height          =   285
         Left            =   1065
         MaxLength       =   9
         TabIndex        =   2
         Top             =   1170
         Width           =   1140
      End
      Begin MSDataListLib.DataCombo dbcintReceita 
         Height          =   315
         Left            =   1065
         TabIndex        =   1
         Top             =   780
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dbcintCidade 
         Height          =   315
         Left            =   1065
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   405
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Style           =   2
         Text            =   ""
      End
      Begin VB.Label lblstrDescricao 
         AutoSize        =   -1  'True
         Caption         =   "Descrição"
         Height          =   195
         Left            =   255
         TabIndex        =   13
         Top             =   1890
         Width           =   720
      End
      Begin VB.Label lblintReceita 
         AutoSize        =   -1  'True
         Caption         =   "Receita"
         Height          =   195
         Left            =   420
         TabIndex        =   11
         Top             =   900
         Width           =   555
      End
      Begin VB.Label lblintCidade 
         AutoSize        =   -1  'True
         Caption         =   "Município"
         Height          =   195
         Left            =   270
         TabIndex        =   10
         Top             =   510
         Width           =   705
      End
      Begin VB.Label lblstrNome 
         AutoSize        =   -1  'True
         Caption         =   "Nome"
         Height          =   195
         Left            =   555
         TabIndex        =   9
         Top             =   1590
         Width           =   420
      End
      Begin VB.Label lbldblCodigo 
         AutoSize        =   -1  'True
         Caption         =   "Código"
         Height          =   195
         Left            =   480
         TabIndex        =   8
         Top             =   1245
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmCadFormulaDeCalculos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mblnAlterando   As Boolean
Dim mobjAux         As Object
Dim adoResultado    As ADODB.Recordset
Dim strSql          As String
Dim objList         As Object
Dim mblnSelecionou  As Boolean
Dim mblnPrimeiraVez As Boolean

Private Sub cmd_AVG_Click()
Dim strAux As String
    strAux = txtstrFormula
    strAux = strAux & " AVG() "
    txtstrFormula = strAux
    txtstrFormula.SetFocus
    txtstrFormula.SelStart = Len(txtstrFormula.Text) + 1
End Sub

Private Sub cmd_Count_Click()
Dim strAux As String
    strAux = txtstrFormula
    strAux = strAux & " COUNT() "
    txtstrFormula = strAux
    txtstrFormula.SetFocus
    txtstrFormula.SelStart = Len(txtstrFormula.Text) + 1
End Sub

Private Sub cmd_Create_Click()
Dim strAux As String
strAux = txtstrFormula
strAux = strAux & " CREATE "
txtstrFormula = strAux
txtstrFormula.SetFocus
txtstrFormula.SelStart = Len(txtstrFormula.Text) + 1
End Sub

Private Sub cmd_Delete_Click()
Dim strAux As String
    strAux = txtstrFormula
    strAux = strAux & " DELETE "
    txtstrFormula = strAux
    txtstrFormula.SetFocus
    txtstrFormula.SelStart = Len(txtstrFormula.Text) + 1
End Sub

Private Sub cmd_Divisao_Click()
Dim strAux As String
    strAux = txtstrFormula
    strAux = strAux & " / "
    txtstrFormula = strAux
    txtstrFormula.SetFocus
    txtstrFormula.SelStart = Len(txtstrFormula.Text) + 1
End Sub

Private Sub cmd_From_Click()
Dim strAux As String
    strAux = txtstrFormula
    strAux = strAux & " FROM "
    txtstrFormula = strAux
    txtstrFormula.SetFocus
    txtstrFormula.SelStart = Len(txtstrFormula.Text) + 1
End Sub

Private Sub cmd_Group_Click()
Dim strAux As String
    strAux = txtstrFormula
    strAux = strAux & " GROUP BY "
    txtstrFormula = strAux
    txtstrFormula.SetFocus
    txtstrFormula.SelStart = Len(txtstrFormula.Text) + 1
End Sub

Private Sub cmd_Insert_Click()
Dim strAux As String
    strAux = txtstrFormula
    strAux = strAux & " INSERT "
    txtstrFormula = strAux
    txtstrFormula.SetFocus
    txtstrFormula.SelStart = Len(txtstrFormula.Text) + 1
End Sub

Private Sub cmd_intReceita_Click()
    ChamaFormCadastro frmCadReceita, dbcintReceita, "PKId, strDescricao"
End Sub

Private Sub cmd_Mais_Click()
Dim strAux As String
    strAux = txtstrFormula
    strAux = strAux & " + "
    txtstrFormula = strAux
    txtstrFormula.SetFocus
    txtstrFormula.SelStart = Len(txtstrFormula.Text) + 1
End Sub

Private Sub cmd_Max_Click()
Dim strAux As String
    strAux = txtstrFormula
    strAux = strAux & " MAX() "
    txtstrFormula = strAux
    txtstrFormula.SetFocus
    txtstrFormula.SelStart = Len(txtstrFormula.Text) + 1
End Sub

Private Sub cmd_menos_Click()
Dim strAux As String
    strAux = txtstrFormula
    strAux = strAux & " - "
    txtstrFormula = strAux
    txtstrFormula.SetFocus
    txtstrFormula.SelStart = Len(txtstrFormula.Text) + 1
End Sub

Private Sub cmd_Min_Click()
Dim strAux As String
    strAux = txtstrFormula
    strAux = strAux & " MIN() "
    txtstrFormula = strAux
    txtstrFormula.SetFocus
    txtstrFormula.SelStart = Len(txtstrFormula.Text) + 1
End Sub

Private Sub cmd_Mod_Click()

'******************************************************************************************
' Data: 11/03/2003
' Alteração: - Alteração da string de comando MOD para o Oracle.
' Responsável: Everton Bianchini
'******************************************************************************************
    
Dim strAux As String
    strAux = txtstrFormula
'    strAux = strAux & " Mod "
    If bytDBType = EDatabases.SQLServer Then
        strAux = strAux & " Mod "
    ElseIf bytDBType = EDatabases.Oracle Then
        strAux = strAux & " Mod( , ) "
    End If
    txtstrFormula = strAux
    txtstrFormula.SetFocus
    txtstrFormula.SelStart = Len(txtstrFormula.Text) + 1
End Sub

Private Sub cmd_Multiplicacao_Click()
Dim strAux As String
    strAux = txtstrFormula
    strAux = strAux & " * "
    txtstrFormula = strAux
    txtstrFormula.SetFocus
    txtstrFormula.SelStart = Len(txtstrFormula.Text) + 1
End Sub

Private Sub cmd_Order_Click()
Dim strAux As String
    strAux = txtstrFormula
    strAux = strAux & " ORDER BY "
    txtstrFormula = strAux
    txtstrFormula.SetFocus
    txtstrFormula.SelStart = Len(txtstrFormula.Text) + 1
End Sub

Private Sub cmd_Porcentagem_Click()
Dim strAux As String
    strAux = txtstrFormula
    strAux = strAux & " % "
    txtstrFormula = strAux
    txtstrFormula.SetFocus
    txtstrFormula.SelStart = Len(txtstrFormula.Text) + 1
End Sub

Private Sub cmd_Select_Click()
Dim strAux As String
    strAux = txtstrFormula
    strAux = strAux & " SELECT "
    txtstrFormula = strAux
    txtstrFormula.SetFocus
    txtstrFormula.SelStart = Len(txtstrFormula.Text) + 1
End Sub

Private Sub cmd_Sum_Click()
Dim strAux As String
    strAux = txtstrFormula
    strAux = strAux & " SUM() "
    txtstrFormula = strAux
    txtstrFormula.SetFocus
    txtstrFormula.SelStart = Len(txtstrFormula.Text) + 1
End Sub

Private Sub cmd_Update_Click()
Dim strAux As String
    strAux = txtstrFormula
    strAux = strAux & " UPDATE "
    txtstrFormula = strAux
    txtstrFormula.SetFocus
    txtstrFormula.SelStart = Len(txtstrFormula.Text) + 1
End Sub

Private Sub cmd_Where_Click()
Dim strAux As String
    strAux = txtstrFormula
    strAux = strAux & " WHERE "
    txtstrFormula = strAux
    txtstrFormula.SetFocus
    txtstrFormula.SelStart = Len(txtstrFormula.Text) + 1
End Sub

Private Sub dbcintCidade_Click(Area As Integer)
   DropDownDataCombo dbcintCidade, Me, Area
End Sub

Private Sub dbcintCidade_KeyDown(KeyCode As Integer, Shift As Integer)
   DropDownDataCombo dbcintCidade, Me, , KeyCode, Shift
End Sub

Private Sub dbcintCidade_KeyPress(KeyAscii As Integer)
   CaracterValido KeyAscii, "N", dbcintCidade
End Sub

Private Sub dbcintReceita_Click(Area As Integer)
   DropDownDataCombo dbcintReceita, Me, Area
End Sub

Private Sub dbcintReceita_KeyDown(KeyCode As Integer, Shift As Integer)
   DropDownDataCombo dbcintReceita, Me, , KeyCode, Shift
End Sub

Private Sub dbcintReceita_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbcintReceita
End Sub

Private Sub Form_Activate()
    gintCodSeguranca = 578
    VirificaGradeListView Me
    If mblnSelecionou Then
        HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrDeletar
    Else
        HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrDeletar
    End If
    If mobjAux Is Nothing Then
        HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar
    Else
        HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrAplicar
    End If
End Sub

Private Sub Form_Deactivate()
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar
End Sub

Private Sub Form_Load()
    txtstrFormula.Text = ""
    mblnAlterando = False
    'txtstrCidade = Trim(gstrNomeEmpresa)
    LeDaTabelaParaObj gstrEmpresa, dbcintCidade, "PKId, strNomeFantasia"
    dbcintCidade.BoundText = 1
    
    dbcintReceita.Tag = strQueryDataComboReceita & ";strDescricao"
    
    LeDaTabelaParaObj gstrFormulaDeCalculo, tdb_FormulaCalculo, strQuery
    VerificaObjParaAplicar mobjAux
End Sub

Private Function strQueryDataComboReceita()
    Dim strSql As String
    strSql = ""
    strSql = strSql & "SELECT PKId, strDescricao "
    strSql = strSql & "FROM " & gstrReceita & " "
    strSql = strSql & "ORDER BY strDescricao"
    strQueryDataComboReceita = strSql
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar, gstrDeletar
    mblnSelecionou = False
    mblnPrimeiraVez = False
End Sub


Private Sub tdb_FormulaCalculo_Click()
    mblnPrimeiraVez = True
    With tdb_FormulaCalculo
        If Not .BOF And Not .EOF Then
           If .Bookmark = 1 Then
               tdb_FormulaCalculo_RowColChange 0, 0
           End If
       End If
    End With
End Sub

Private Sub tdb_FormulaCalculo_DblClick()
    MantemForm gstrAplicar
End Sub

Private Sub tdb_FormulaCalculo_FilterChange()
    gblnFilraCampos tdb_FormulaCalculo
End Sub

Private Function strQuery() As String
    Dim strSql As String
    strSql = ""
    strSql = strSql & " SELECT PKID, dblCodigo, strNome, strDescricao FROM " & gstrFormulaDeCalculo
    strQuery = strSql
End Function

Private Function blnProcedimentoOK(strFormula As String) As Boolean

'******************************************************************************************
' Data: 11/03/2003
' Alteração: - Adaptação da instrução de exclusão de procedure ao Oracle.
' Responsável: Everton Bianchini
'******************************************************************************************
    
Dim strMsg As String
Dim strSql As String

On Error GoTo err_blnProcedimentoOK

    blnProcedimentoOK = False
    
    strSql = ""
'    strSql = strSql & " IF EXISTS (SELECT NAME FROM SYSOBJECTS"
'    strSql = strSql & " WHERE NAME = '" & txtstrNome & "' AND TYPE = 'P')"
'    strSql = strSql & " DROP PROCEDURE " & txtstrNome
    If bytDBType = EDatabases.SQLServer Then
        strSql = strSql & " IF EXISTS (SELECT NAME FROM SYSOBJECTS"
        strSql = strSql & " WHERE NAME = '" & txtstrNome & "' AND TYPE = 'P')"
        strSql = strSql & " DROP PROCEDURE " & txtstrNome
        
    ElseIf bytDBType = EDatabases.Oracle Then
        strSql = strSql & "DECLARE "
        strSql = strSql & "varSQL VARCHAR2(100); "
        strSql = strSql & "numCursor NUMBER; "
        strSql = strSql & "numReturn NUMBER; "
        strSql = strSql & "excNoPriveleges EXCEPTION; "
        strSql = strSql & "PRAGMA EXCEPTION_INIT (excNoPriveleges, -20040); "
        strSql = strSql & "BEGIN "
        strSql = strSql & "SELECT COUNT(*) INTO numReturn FROM ALL_OBJECTS "
        strSql = strSql & "WHERE UPPER(OBJECT_NAME) = '" & UCase(txtstrNome) & "' AND OBJECT_TYPE = 'PROCEDURE';"
        strSql = strSql & " IF numReturn > 0 THEN "
        strSql = strSql & " varSQL := 'DROP PROCEDURE " & txtstrNome & "'; "
        strSql = strSql & " numCursor := DBMS_SQL.OPEN_CURSOR; "
        strSql = strSql & " DBMS_SQL.PARSE(numCursor, varSQL, DBMS_SQL.V7); "
        strSql = strSql & " numReturn := DBMS_SQL.EXECUTE(numCursor); "
        strSql = strSql & " DBMS_SQL.CLOSE_CURSOR(numCursor); "
        strSql = strSql & " END IF;"
        strSql = strSql & " END;"
    
    End If
    
    Set gobjBanco = New clsBanco
    gobjBanco.Execute strSql
    
    gcncADOMain.BeginTrans
    
    
    strFormula = Replace(strFormula, Chr(207), "'")
    gcncADOMain.Execute strFormula, , adCmdText
    gcncADOMain.CommitTrans
    
    MsgBox "Procedimento efetuado com sucesso!"
    
    blnProcedimentoOK = True
    
    Exit Function
err_blnProcedimentoOK:
    
    gcncADOMain.RollbackTrans
    strMsg = ""
    strMsg = strMsg & "Ocorreu um erro na gravação do procedimento no Banco."
    ExibeDetalheErro strMsg
End Function

Public Sub MantemForm(ByVal strModoOperacao As String)
Dim strSql As String
Dim strMsg As String
Dim strFormula As String
Dim blnAlterando As Boolean
'Dim strNome As String
strSql = strQuery
'strNome = txtstrNome

strFormula = Replace(txtstrFormula, "'", Chr(207))
blnAlterando = mblnAlterando

If strModoOperacao = UCase(gstrImprimir) Then
    ToolBarGeral strModoOperacao, gstrFormulaDeCalculo, mblnAlterando, tdb_FormulaCalculo, Me, mobjAux, strSql, , rptCadFormulaDeCalculos, strQuery
    Exit Sub
End If


If ToolBarGeral(strModoOperacao, gstrFormulaDeCalculo, mblnAlterando, tdb_FormulaCalculo, Me, mobjAux, strSql, , , , False) Then
    If UCase(strModoOperacao) = gstrNovo Then
        dbcintCidade.BoundText = 1
    End If
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrDeletar
    mblnPrimeiraVez = False
    
    If UCase(strModoOperacao) = gstrSalvar Then
        If Trim(strFormula) <> "" Then
            If blnAlterando Then
                strMsg = ""
 '               txtstrNome = strNome
                strMsg = strMsg & "Deseja atualizar o procedimento armazenado?"
                If gblnExclusaoGravacaoOk("", strMsg, True) Then
                    If Not blnProcedimentoOK(strFormula) Then
                        mblnPrimeiraVez = True
                        tdb_FormulaCalculo_RowColChange 0, 0
                        mblnPrimeiraVez = False
                    End If
                End If
            Else
                If Not blnProcedimentoOK(strFormula) Then
                
                End If
            End If
        End If
    End If
    
End If
    
End Sub

Private Sub tdb_FormulaCalculo_HeadClick(ByVal ColIndex As Integer)
    gOrdenaGrid tdb_FormulaCalculo, ColIndex '
End Sub

Private Sub tdb_FormulaCalculo_KeyPress(KeyAscii As Integer)
    Select Case tdb_FormulaCalculo.Col
        Case 0
            CaracterValido KeyAscii, "N", tdb_FormulaCalculo
        Case Else
            CaracterValido KeyAscii, "A", tdb_FormulaCalculo
    End Select
End Sub

Private Sub tdb_FormulaCalculo_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    With tdb_FormulaCalculo
        If Not .EOF And Not .BOF Then
            If mblnPrimeiraVez Then
                txtPKID.Text = .Columns("PKID").Value
                LeDaTabelaParaObj gstrFormulaDeCalculo, Me
                'Busca Fórmula Catalogada no Banco
                BuscaFormulaCatalogada
                gCorLinhaSelecionada tdb_FormulaCalculo
                
                HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrDeletar
                
                If mobjAux Is Nothing Then
                    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar
                Else
                    HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrAplicar
                End If
                mblnSelecionou = True
                mblnAlterando = True
            End If
        End If
    End With
End Sub
Private Sub BuscaFormulaCatalogada()

'******************************************************************************************
' Data: 14/03/2003
' Alteração: - Adaptação da leitura da estrutura das stored procedures no Banco.
' Responsável: Everton Bianchini
'******************************************************************************************
    
    Dim strSql As String
    Dim adoResultado As ADODB.Recordset
    txtstrFormula.Text = ""
'    strSQL = "EXECUTE sp_helpText '" & txtstrNome.Text & "'"
    If (bytDBType = EDatabases.SQLServer) Then
        strSql = "EXECUTE sp_helpText '" & txtstrNome.Text & "'"
    
    ElseIf (bytDBType = EDatabases.Oracle) Then
        strSql = "SELECT TEXT text FROM ALL_SOURCE WHERE UPPER(NAME) = '" & UCase(txtstrNome.Text) & "' AND "
        strSql = strSql & "TYPE = 'PROCEDURE' ORDER BY LINE"
    End If
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        If (Not (adoResultado.EOF)) And (bytDBType = EDatabases.Oracle) Then
            txtstrFormula.Text = txtstrFormula.Text & "CREATE OR REPLACE "
        End If
        Do While Not adoResultado.EOF
'            txtstrFormula.Text = txtstrFormula.Text & gstrENulo(adoResultado("text"))
            If (bytDBType = EDatabases.SQLServer) Then
                txtstrFormula.Text = txtstrFormula.Text & gstrENulo(adoResultado("text"))
            ElseIf (bytDBType = EDatabases.Oracle) Then
                txtstrFormula.Text = txtstrFormula.Text & Replace(gstrENulo(adoResultado("text")), Chr(10), "") & vbCrLf
            End If
            adoResultado.MoveNext
        Loop
    End If
End Sub

Private Sub txtdblCodigo_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtdblCodigo
End Sub

Private Sub txtstrDescricao_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrDescricao
End Sub



Private Sub txtstrNome_GotFocus()
    MarcaCampo txtstrNome
End Sub

Private Sub txtstrDescricao_GotFocus()
    MarcaCampo txtstrDescricao
End Sub

Private Sub txtdblCodigo_GotFocus()
    MarcaCampo txtdblCodigo
End Sub

Function strQueryRelatorio() As String
Dim strSql As String
    strSql = ""
    strSql = strSql & "SELECT FC.*, RC.strDescricao Receita "
    strSql = strSql & " FROM " & gstrFormulaDeCalculo & " FC, "
    strSql = strSql & gstrReceita & " RC "
    If mblnAlterando = True Then
        strSql = strSql & " WHERE FC.intReceita = RC.PKId and FC.PKId = " & tdb_FormulaCalculo.Columns("PKId").Value
        Else
        strSql = strSql & " WHERE FC.intReceita = RC.PKId "
    End If
    strSql = strSql & " ORDER BY FC.strNome"
strQueryRelatorio = strSql
End Function

Private Sub txtstrNome_KeyPress(KeyAscii As Integer)
    If gblnValidaNome(KeyAscii) Then
        CaracterValido KeyAscii, "A", txtstrNome
    Else: KeyAscii = 0
    End If
End Sub

Private Function gblnValidaNome(intCaracter As Integer) As Boolean
    Select Case intCaracter
        Case vbKeyBack, vbKeyDelete
        Case 48 To 59, 65 To 90, 95, 97 To 122
        Case Else
            gblnValidaNome = False
            Exit Function
    End Select
    gblnValidaNome = True
End Function

Private Sub cmd_BuscaFormula_Click()
    With dlg_BuscaFormula
        .DialogTitle = "Abrir arquivo com Fórmula"
        .DefaultExt = "*.*"
        .Filter = "*.*"
        .InitDir = App.Path
        .flags = &H4
        .Filename = ""
        .ShowOpen
        If .Filename <> "" Then
            BuscaFormula (.Filename)
        End If
    End With
End Sub

Private Sub BuscaFormula(strArquivoFormula As String)
    Dim strLinha As String
    Screen.MousePointer = 11
    On Error GoTo ErroNaAbertura
    Open strArquivoFormula For Input As #1
        txtstrFormula.SetFocus
        txtstrFormula.Text = ""
        While Not EOF(1)
            Line Input #1, strLinha
            txtstrFormula.Text = txtstrFormula.Text & strLinha & Chr(13) & Chr(10)
        Wend
    Close #1
    On Error GoTo 0
ErroNaAbertura:
    Screen.MousePointer = vbDefault
End Sub
