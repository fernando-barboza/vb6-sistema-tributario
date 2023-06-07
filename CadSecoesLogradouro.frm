VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "TODG7.OCX"
Begin VB.Form frmCadSecoesLogradouro 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Seções de Logradouro"
   ClientHeight    =   6315
   ClientLeft      =   330
   ClientTop       =   2385
   ClientWidth     =   7800
   HelpContextID   =   37
   Icon            =   "CadSecoesLogradouro.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6315
   ScaleWidth      =   7800
   Begin TabDlg.SSTab tab_3DPasta 
      Height          =   6105
      Left            =   105
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   105
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   10769
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Seções de Logradouro"
      TabPicture(0)   =   "CadSecoesLogradouro.frx":1042
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblintLargura"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblintLogradouro"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblstrInscricaoCadastral"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblintUtilizacao"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblintValorDaSecao"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "dbcintLogradouro"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "tdb_Secao"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtintLargura"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cmd_CadLogradouro"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cbointUtilizacao"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "mskstrInscricaoCadastral"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "cbointValorDaSecao"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).ControlCount=   12
      TabCaption(1)   =   "Melhoramentos Públicos "
      TabPicture(1)   =   "CadSecoesLogradouro.frx":105E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lbl_Ano"
      Tab(1).Control(1)=   "lbl_Melhoramentos"
      Tab(1).Control(2)=   "lbl_MelhoramentosExistentes"
      Tab(1).Control(3)=   "lvw_Melhoramentos"
      Tab(1).Control(4)=   "lvw_MelhoramentosCadastrados"
      Tab(1).Control(5)=   "cmd_Remover"
      Tab(1).Control(6)=   "cmd_Adicionar"
      Tab(1).Control(7)=   "txt_intAno"
      Tab(1).ControlCount=   8
      Begin VB.ComboBox cbointValorDaSecao 
         Height          =   315
         Left            =   1695
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1665
         Width           =   2790
      End
      Begin VB.TextBox txt_intAno 
         Height          =   285
         Left            =   -72390
         MaxLength       =   4
         TabIndex        =   17
         Top             =   4560
         Width           =   645
      End
      Begin VB.CommandButton cmd_Adicionar 
         Height          =   465
         Left            =   -71670
         MouseIcon       =   "CadSecoesLogradouro.frx":107A
         MousePointer    =   99  'Custom
         Picture         =   "CadSecoesLogradouro.frx":1384
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Adicionar"
         Top             =   900
         Width           =   465
      End
      Begin VB.CommandButton cmd_Remover 
         Height          =   465
         Left            =   -71670
         MouseIcon       =   "CadSecoesLogradouro.frx":17C6
         MousePointer    =   99  'Custom
         Picture         =   "CadSecoesLogradouro.frx":1AD0
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Remover"
         Top             =   1440
         Width           =   465
      End
      Begin MSMask.MaskEdBox mskstrInscricaoCadastral 
         Height          =   285
         Left            =   1695
         TabIndex        =   1
         Top             =   900
         Width           =   2790
         _ExtentX        =   4921
         _ExtentY        =   503
         _Version        =   393216
         PromptInclude   =   0   'False
         PromptChar      =   "_"
      End
      Begin VB.ComboBox cbointUtilizacao 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1695
         Locked          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1260
         Width           =   2790
      End
      Begin VB.CommandButton cmd_CadLogradouro 
         Height          =   315
         Left            =   7065
         Picture         =   "CadSecoesLogradouro.frx":1F12
         Style           =   1  'Graphical
         TabIndex        =   9
         TabStop         =   0   'False
         Tag             =   "584"
         ToolTipText     =   "Ativa Cadastro de Logradouros"
         Top             =   495
         Width           =   360
      End
      Begin VB.TextBox txtintLargura 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1695
         MaxLength       =   12
         TabIndex        =   4
         Top             =   2055
         Width           =   1095
      End
      Begin MSComctlLib.ListView lvw_MelhoramentosCadastrados 
         Height          =   3645
         Left            =   -74475
         TabIndex        =   14
         Top             =   870
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   6429
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin MSComctlLib.ListView lvw_Melhoramentos 
         Height          =   3645
         Left            =   -71100
         TabIndex        =   15
         Top             =   885
         Width           =   3285
         _ExtentX        =   5794
         _ExtentY        =   6429
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin TrueOleDBGrid70.TDBGrid tdb_Secao 
         Height          =   3450
         Left            =   135
         TabIndex        =   21
         Top             =   2460
         Width           =   7290
         _ExtentX        =   12859
         _ExtentY        =   6085
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).DataField=   "PKID"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Número da Seção"
         Columns(1).DataField=   "Inscricao"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Logradouro"
         Columns(2).DataField=   "Logradouro"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Largura"
         Columns(3).DataField=   "Largura"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "Valor"
         Columns(4).DataField=   "ValorSL"
         Columns(4).NumberFormat=   "Standard"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   5
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
         Splits(0)._ColumnProps(0)=   "Columns.Count=5"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2064"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1984"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0).AllowSizing=0"
         Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
         Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(8)=   "Column(1).Width=2540"
         Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=2461"
         Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(13)=   "Column(2).Width=6482"
         Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=6403"
         Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(17)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(18)=   "Column(3).Width=1482"
         Splits(0)._ColumnProps(19)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(20)=   "Column(3)._WidthInPix=1402"
         Splits(0)._ColumnProps(21)=   "Column(3)._EditAlways=0"
         Splits(0)._ColumnProps(22)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(23)=   "Column(4).Width=1826"
         Splits(0)._ColumnProps(24)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(25)=   "Column(4)._WidthInPix=1746"
         Splits(0)._ColumnProps(26)=   "Column(4)._EditAlways=0"
         Splits(0)._ColumnProps(27)=   "Column(4).Order=5"
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
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33"
         _StyleDefs(7)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
         _StyleDefs(8)   =   "HeadingStyle:id=2,.parent=1,.namedParent=34"
         _StyleDefs(9)   =   "FooterStyle:id=3,.parent=1,.namedParent=35"
         _StyleDefs(10)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(11)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
         _StyleDefs(12)  =   "EditorStyle:id=7,.parent=1"
         _StyleDefs(13)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
         _StyleDefs(14)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
         _StyleDefs(15)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
         _StyleDefs(16)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(17)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
         _StyleDefs(18)  =   "Splits(0).Style:id=13,.parent=1"
         _StyleDefs(19)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
         _StyleDefs(20)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
         _StyleDefs(21)  =   "Splits(0).FooterStyle:id=15,.parent=3"
         _StyleDefs(22)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
         _StyleDefs(23)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
         _StyleDefs(24)  =   "Splits(0).EditorStyle:id=17,.parent=7"
         _StyleDefs(25)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
         _StyleDefs(26)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
         _StyleDefs(27)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
         _StyleDefs(28)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(29)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(30)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
         _StyleDefs(31)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
         _StyleDefs(32)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
         _StyleDefs(33)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
         _StyleDefs(34)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
         _StyleDefs(35)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
         _StyleDefs(36)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
         _StyleDefs(37)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
         _StyleDefs(38)  =   "Splits(0).Columns(2).Style:id=46,.parent=13"
         _StyleDefs(39)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
         _StyleDefs(40)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
         _StyleDefs(41)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
         _StyleDefs(42)  =   "Splits(0).Columns(3).Style:id=50,.parent=13"
         _StyleDefs(43)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
         _StyleDefs(44)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
         _StyleDefs(45)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
         _StyleDefs(46)  =   "Splits(0).Columns(4).Style:id=54,.parent=13"
         _StyleDefs(47)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14"
         _StyleDefs(48)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
         _StyleDefs(49)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
         _StyleDefs(50)  =   "Named:id=33:Normal"
         _StyleDefs(51)  =   ":id=33,.parent=0"
         _StyleDefs(52)  =   "Named:id=34:Heading"
         _StyleDefs(53)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(54)  =   ":id=34,.wraptext=-1"
         _StyleDefs(55)  =   "Named:id=35:Footing"
         _StyleDefs(56)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(57)  =   "Named:id=36:Selected"
         _StyleDefs(58)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(59)  =   "Named:id=37:Caption"
         _StyleDefs(60)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(61)  =   "Named:id=38:HighlightRow"
         _StyleDefs(62)  =   ":id=38,.parent=33,.bgcolor=&H80000014&,.fgcolor=&H80000012&"
         _StyleDefs(63)  =   "Named:id=39:EvenRow"
         _StyleDefs(64)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(65)  =   "Named:id=40:OddRow"
         _StyleDefs(66)  =   ":id=40,.parent=33"
         _StyleDefs(67)  =   "Named:id=41:RecordSelector"
         _StyleDefs(68)  =   ":id=41,.parent=34"
         _StyleDefs(69)  =   "Named:id=42:FilterBar"
         _StyleDefs(70)  =   ":id=42,.parent=33"
      End
      Begin MSDataListLib.DataCombo dbcintLogradouro 
         Height          =   315
         Left            =   1695
         TabIndex        =   0
         Top             =   495
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Text            =   ""
      End
      Begin VB.Label lblintValorDaSecao 
         AutoSize        =   -1  'True
         Caption         =   "Valor da Seção"
         Height          =   195
         Left            =   495
         TabIndex        =   20
         Top             =   1785
         Width           =   1095
      End
      Begin VB.Label lbl_MelhoramentosExistentes 
         AutoSize        =   -1  'True
         Caption         =   "Melhoramentos existentes na seção:"
         Height          =   195
         Left            =   -71130
         TabIndex        =   19
         Top             =   630
         Width           =   2580
      End
      Begin VB.Label lbl_Melhoramentos 
         AutoSize        =   -1  'True
         Caption         =   "Melhoramentos cadastrados"
         Height          =   195
         Left            =   -74490
         TabIndex        =   18
         Top             =   630
         Width           =   1995
      End
      Begin VB.Label lbl_Ano 
         AutoSize        =   -1  'True
         Caption         =   "Ano do melhoramento"
         Height          =   195
         Left            =   -74040
         TabIndex        =   16
         Top             =   4650
         Width           =   1545
      End
      Begin VB.Label lblintUtilizacao 
         AutoSize        =   -1  'True
         Caption         =   "Utilização"
         Height          =   195
         Left            =   900
         TabIndex        =   11
         Top             =   1380
         Width           =   690
      End
      Begin VB.Label lblstrInscricaoCadastral 
         AutoSize        =   -1  'True
         Caption         =   "Número da Seção"
         Height          =   195
         Left            =   300
         TabIndex        =   10
         Top             =   990
         Width           =   1290
      End
      Begin VB.Label lblintLogradouro 
         AutoSize        =   -1  'True
         Caption         =   "Logradouro"
         Height          =   195
         Left            =   780
         TabIndex        =   8
         Top             =   570
         Width           =   810
      End
      Begin VB.Label lblintLargura 
         AutoSize        =   -1  'True
         Caption         =   "Largura"
         Height          =   195
         Left            =   1050
         TabIndex        =   7
         Top             =   2145
         Width           =   540
      End
   End
   Begin VB.TextBox txtPKId 
      Height          =   270
      Left            =   2580
      TabIndex        =   5
      Top             =   240
      Visible         =   0   'False
      Width           =   645
   End
End
Attribute VB_Name = "frmCadSecoesLogradouro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    
Dim mblnPrimeiraVez     As Boolean
Dim mblnAlterando       As Boolean
Dim mcboAux             As ComboBox
Dim strSQL              As String
Dim objList             As Object
Dim adoResultado        As Object
Dim mobjGeral           As Object

Private Sub dbcintLogradouro_Click(Area As Integer)
    DropDownDataCombo dbcintLogradouro, Me, Area
    If Area = 2 And dbcintLogradouro.MatchedWithList Then
        mblnPrimeiraVez = False
        mblnAlterando = False
        'LimpaObjeto Me
        VerificaListaAutomatica gstrLogradouro, tdb_Secao, strQuerySecaoLogradouro
    End If
End Sub

Private Sub dbcintLogradouro_GotFocus()
tab_3dPasta.Tab = 0
End Sub

Private Sub dbcintLogradouro_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbcintLogradouro, Me, , KeyCode, Shift
End Sub

Private Sub dbcintLogradouro_KeyPress(KeyAscii As Integer)
   CaracterValido KeyAscii, "A", dbcintLogradouro
End Sub

Private Sub cbointUtilizacao_GotFocus()
tab_3dPasta.Tab = 0
End Sub

Private Sub cbointUtilizacao_KeyPress(KeyAscii As Integer)
CaracterValido KeyAscii, "A", cbointUtilizacao
End Sub

Private Sub cbointValorDaSecao_GotFocus()
tab_3dPasta.Tab = 0
End Sub

Private Sub cbointValorDaSecao_KeyPress(KeyAscii As Integer)
CaracterValido KeyAscii, "A", cbointValorDaSecao
End Sub

Private Sub cmd_Adicionar_Click()
    AdicionarMelhoramento
End Sub

Private Sub cmd_CadLogradouro_Click()
    ChamaFormCadastro frmCadLogradouro, dbcintLogradouro
End Sub

Private Sub cmd_Remover_Click()
    RemoverMelhoramento
End Sub

Private Sub Form_Activate()
    gintCodSeguranca = 612
    VirificaGradeListView Me
End Sub

Private Sub Form_Load()

    MontaColumnHeaders
    dbcintLogradouro.Tag = strQueryLogradouro & ";LOG.strDescricao"
    VerificaListaAutomatica gstrUtilizacaoDaTabelaDeValor, cbointUtilizacao
    VerificaListaAutomatica gstrMelhoramentoPublico, lvw_MelhoramentosCadastrados, "PKId, strNomeDoMelhoramento"
    PreencheComboValorDaSecao
    VerificaMascaraInscricao
    cbointUtilizacao.ListIndex = gintIndiceCBO(cbointUtilizacao, 2)
    
End Sub

Private Function strQueryLogradouro() As String

'******************************************************************************************
' Data: 02/05/2003
' Alteração: - Substituição do comando nativo ISNULL() do SQL Server pela função gstrISNULL
' Responsável: Everton Bianchini
'------------------------------------------------------------------------------------------
' Data: 05/05/2003
' Alteração: - Substituição do comando nativo de concatenação de strings (+) do SQL Server
'            pela variável strCONCAT.
' Responsável: Everton Bianchini
'------------------------------------------------------------------------------------------
' Data: 05/05/2003
' Alteração: - Deslocamento dos JOINS utilizados na cláusula FROM para a cláusula WHERE.
' Responsável: Everton Bianchini
'******************************************************************************************
    
    Dim strSQL As String
    
    strSQL = strSQL & " SELECT LOG.PKId, "
'    strSql = strSql & "RTRIM(ISNULL(TLO.strSigla, '')) + ' ' + LTRIM(ISNULL(TIT.strDescricao,''))  "
    strSQL = strSQL & "RTRIM(" & gstrISNULL("TLO.strSigla", "''") & ") " & strCONCAT & " ' ' " & strCONCAT & " LTRIM(" & gstrISNULL("TIT.strDescricao", "''") & ")  "
'    strSQL = strSQL & " + ' ' +  LOG.strDescricao AS strDescricao "
    strSQL = strSQL & strCONCAT & " ' ' " & strCONCAT & "  LOG.strDescricao AS strDescricao "
    strSQL = strSQL & " FROM "
    strSQL = strSQL & gstrLogradouro & " LOG "
'    strSql = strSql & "LEFT JOIN  " & gstrTituloLogradouro & " TIT "
    strSQL = strSQL & ", " & gstrTituloLogradouro & " TIT "
'    strSql = strSql & "ON LOG.intTituloLogradouro = TIT.PKId "
'    strSql = strSql & "LEFT JOIN " & gstrTipoLogradouro & " TLO "
    strSQL = strSQL & ", " & gstrTipoLogradouro & " TLO "
'    strSql = strSql & "ON LOG.intTipoLogradouro = TLO.PKId "
    
    strSQL = strSQL & " WHERE LOG.intTituloLogradouro " & strOUTJSQLServer & "= TIT.PKId " & strOUTJOracle
    strSQL = strSQL & " AND LOG.intTipoLogradouro " & strOUTJSQLServer & "= TLO.PKId " & strOUTJOracle
    
    strSQL = strSQL & " ORDER BY LOG.strDescricao "
    strQueryLogradouro = strSQL
End Function


Sub MontaColumnHeaders()
    
    With lvw_MelhoramentosCadastrados
        .ColumnHeaders.Clear
        .ColumnHeaders.Add 1, , "Descrição", 2690
    End With
    
     With lvw_Melhoramentos
        .ColumnHeaders.Clear
        .ColumnHeaders.Add 1, , "Descrição", 2400
        .ColumnHeaders.Add 2, , "Ano", 800
    End With
End Sub

Sub CarregaMelhoramentos(intCodSecao As Integer)
    lvw_Melhoramentos.ListItems.Clear
    txt_intAno = ""
    
    strSQL = ""
    strSQL = strSQL & "Select M.PKId, M.strNomeDoMelhoramento Descricao, MS.intAno "
    strSQL = strSQL & "From " & gstrMelhoramentoDaSecaoDeLogradouro & " MS,  "
    strSQL = strSQL & gstrMelhoramentoPublico & " M "
    strSQL = strSQL & "Where MS.intMelhoramento = M.PKId And MS.intSecaoDeLogradouro = " & intCodSecao
    
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
        With adoResultado
            Do While Not .EOF
                Set objList = lvw_Melhoramentos.ListItems.Add(, , !Descricao)
                objList.SubItems(1) = !intAno
                objList.Tag = !Pkid
                .MoveNext
            Loop
        End With
    End If
End Sub

Private Sub lvw_Melhoramentos_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    OrdenaColunaClicada lvw_Melhoramentos, ColumnHeader
End Sub

Private Sub lvw_Melhoramentos_GotFocus()
tab_3dPasta.Tab = 1
End Sub

Private Sub lvw_Melhoramentos_KeyPress(KeyAscii As Integer)
CaracterValido KeyAscii, "A", lvw_Melhoramentos
End Sub

Private Sub lvw_MelhoramentosCadastrados_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    OrdenaColunaClicada lvw_MelhoramentosCadastrados, ColumnHeader
End Sub

Private Sub lvw_MelhoramentosCadastrados_DblClick()
    cmd_Adicionar_Click
End Sub

Private Sub lvw_MelhoramentosCadastrados_GotFocus()
tab_3dPasta.Tab = 1
End Sub

Private Sub lvw_MelhoramentosCadastrados_KeyPress(KeyAscii As Integer)
CaracterValido KeyAscii, "A", lvw_MelhoramentosCadastrados
End Sub

Private Sub mskstrInscricaoCadastral_GotFocus()
    MarcaCampo mskstrInscricaoCadastral
    tab_3dPasta.Tab = 0
End Sub

Private Sub mskstrInscricaoCadastral_KeyPress(KeyAscii As Integer)
    On Error GoTo err_mskstrInscricaoCadastral_KeyPress
    CaracterValido KeyAscii, "A", mskstrInscricaoCadastral
    
    Select Case KeyAscii
        Case vbKeyBack
            Exit Sub
    End Select
    
    If Len(mskstrInscricaoCadastral.Mask) - Val(mskstrInscricaoCadastral.SelStart) = 1 Then
        Select Case UCase(Chr(KeyAscii))
            Case "D", "E", "X"
                KeyAscii = Asc(UCase(Chr(KeyAscii)))
            Case Else
                KeyAscii = 0
                Beep
        End Select
    Else
        Select Case KeyAscii
            Case vbKey0 To vbKey9
                Exit Sub
            Case Else
                KeyAscii = 0
                Beep
        End Select
    End If
err_mskstrInscricaoCadastral_KeyPress:
End Sub

Public Sub MantemForm(ByVal strModoOperacao As String)
    Dim intCodSecao As Integer
    If UCase(strModoOperacao) = "Salvar" And mskstrInscricaoCadastral.ClipText = "" Then
        MsgBox " O Número da Seção tem que ser digitado"
        mskstrInscricaoCadastral.SetFocus
        Exit Sub
    End If
    If mblnAlterando Then
        intCodSecao = IIf(IsNull(tdb_Secao.Columns("PKId").Value), 0, tdb_Secao.Columns("PKId").Value)
    Else
        intCodSecao = glngPegaProximaChave(gstrSecaoLogradouro, "PKId")
    End If
    
    Select Case UCase(strModoOperacao)
        Case "SALVAR"
            If blnDadosOk Then
                If ToolBarGeral(strModoOperacao, gstrSecaoLogradouro, mblnAlterando, tdb_Secao, Me, mobjGeral, strQuerySecaoLogradouro) Then
                    mblnPrimeiraVez = False
                    GravaMelhoramentos intCodSecao
                    NovaSecaoDeLogradouro
                    cbointUtilizacao.ListIndex = gintIndiceCBO(cbointUtilizacao, 2)
                End If
            End If
           
            
        Case "DELETAR"
            If ToolBarGeral(strModoOperacao, gstrSecaoLogradouro, mblnAlterando, tdb_Secao, Me, mobjGeral, strQuerySecaoLogradouro) Then
                 mblnPrimeiraVez = False
                 DeletaMelhoramentos intCodSecao
                 NovaSecaoDeLogradouro
                 cbointUtilizacao.ListIndex = gintIndiceCBO(cbointUtilizacao, 2)
            End If
        
        Case "NOVO"
            LimpaObjeto Me, mblnAlterando
            NovaSecaoDeLogradouro
            cbointUtilizacao.ListIndex = gintIndiceCBO(cbointUtilizacao, 2)
            
        'Case gstrImprimir
            'ImprimeRelatorio rptSecaoDeLogradouro, strQuerryRelatorio
        Case gstrLocalizar, gstrPreencherLista
            ToolBarGeral strModoOperacao, gstrSecaoLogradouro, mblnAlterando, tdb_Secao, Me, mobjGeral, strQuerySecaoLogradouro
            
        Case "FECHAR"
            Unload Me
    End Select
    
End Sub

Private Function strQuerySecaoLogradouro() As String

'******************************************************************************************
' Data: 02/05/2003
' Alteração: - Substituição do comando nativo ISNULL() do SQL Server pela função gstrISNULL
' Responsável: Everton Bianchini
'------------------------------------------------------------------------------------------
' Data: 05/05/2003
' Alteração: - Substituição do comando nativo de concatenação de strings (+) do SQL Server
'            pela variável strCONCAT.
' Responsável: Everton Bianchini
'------------------------------------------------------------------------------------------
' Data: 05/05/2003
' Alteração: - Retirado comando "AS" utilizado para dar apelidos às tabelas, o qual não é
'            permitido pelo Oracle, da cláusula FROM.
' Responsável: Everton Bianchini
'------------------------------------------------------------------------------------------
' Data: 05/05/2003
' Alteração: - Deslocamento dos JOINS utilizados na cláusula FROM para a cláusula WHERE.
'            Foi mantida a forma antiga para o SQL Server pois não era possível o
'            deslocamento completo devido à incompatibilidade entres os bancos.
' Responsável: Everton Bianchini
'******************************************************************************************
    
    Dim strSQL As String
    
    strSQL = ""
    strSQL = strSQL & "Select SL.PKId, SL.strInscricaoCadastral Inscricao, TV.dblValor ValorSL,  SL.intLargura Largura, "
'    strSql = strSql & "TP.strSigla, RTRIM(ISNULL(TP.strSigla, '')) + ' ' + LTRIM(ISNULL(TIT.strDescricao,''))  "
    strSQL = strSQL & "TP.strSigla, RTRIM(" & gstrISNULL("TP.strSigla", "''") & ") " & strCONCAT & " ' ' " & strCONCAT & " LTRIM(" & gstrISNULL("TIT.strDescricao", "''") & ")  "
'    strSQL = strSQL & " + ' ' +  L.strDescricao AS Logradouro "
    strSQL = strSQL & strCONCAT & " ' ' " & strCONCAT & "  L.strDescricao AS Logradouro "
'    strSQL = strSQL & "From (( " & gstrSecaoLogradouro & " AS SL "
    
    If (bytDBType = EDatabases.SQLServer) Then
    
        strSQL = strSQL & "From (( " & gstrSecaoLogradouro & " SL "
        strSQL = strSQL & "Left Join " & gstrLogradouro & " L On SL.intLogradouro = L.PKId) "
        strSQL = strSQL & "Left Join " & gstrTituloLogradouro & " TIT On L.intTituloLogradouro = TIT.PKId) "
        strSQL = strSQL & "Left Join " & gstrTipoLogradouro & " TP On L.intTipoLogradouro = TP.PKId "
    
    ElseIf (bytDBType = EDatabases.Oracle) Then
    
        strSQL = strSQL & "From " & gstrSecaoLogradouro & " SL, "
        strSQL = strSQL & gstrLogradouro & " L, "
        strSQL = strSQL & gstrTituloLogradouro & " TIT, "
        strSQL = strSQL & gstrTipoLogradouro & " TP "
    
    End If
    
    strSQL = strSQL & ", " & gstrTabelaDeValor & " TV "
    strSQL = strSQL & "WHERE TV.PKId = SL.intValorDaSecao "
    
    If (bytDBType = EDatabases.Oracle) Then
    
        strSQL = strSQL & " AND SL.intLogradouro = L.PKId " & strOUTJOracle
        strSQL = strSQL & " AND L.intTituloLogradouro = TIT.PKId " & strOUTJOracle
        strSQL = strSQL & " AND L.intTipoLogradouro = TP.PKId " & strOUTJOracle
    
    End If
    
    If dbcintLogradouro.MatchedWithList Then
        strSQL = strSQL & "AND SL.intLogradouro = " & gstrItemData(dbcintLogradouro)
    End If
    
    strQuerySecaoLogradouro = strSQL
End Function

Sub RemoverMelhoramento()
    If lvw_Melhoramentos.ListItems.Count = 0 Then Exit Sub
    If lvw_Melhoramentos.SelectedItem.Selected = False Then Exit Sub
    
    lvw_Melhoramentos.ListItems.Remove lvw_Melhoramentos.SelectedItem.Index
    lvw_Melhoramentos.Sorted = True
End Sub

Sub AdicionarMelhoramento()
    If lvw_MelhoramentosCadastrados.ListItems.Count = 0 Then Exit Sub
    If lvw_MelhoramentosCadastrados.SelectedItem.Selected = False Then Exit Sub
    
    For giContador = 1 To lvw_Melhoramentos.ListItems.Count
        If lvw_Melhoramentos.ListItems(giContador).Tag = lvw_MelhoramentosCadastrados.SelectedItem.Tag Then
            ExibeMensagem "Melhoramento já relacionado com a seção."
            Exit Sub
        End If
    Next
    If Val(txt_intAno) = 0 Then
        ExibeMensagem "O ano do melhoramento tem que ser digitado."
        txt_intAno.SetFocus
        Exit Sub
    End If
    Set objList = lvw_Melhoramentos.ListItems.Add(, , lvw_MelhoramentosCadastrados.SelectedItem.Text)
    objList.SubItems(1) = txt_intAno
    objList.Tag = lvw_MelhoramentosCadastrados.SelectedItem.Tag
End Sub

Private Sub tdb_Secao_Click()
    mblnPrimeiraVez = True
End Sub

Private Sub tdb_Secao_FilterChange()
    gblnFilraCampos tdb_Secao
End Sub

Private Sub tdb_Secao_KeyPress(KeyAscii As Integer)
    If tdb_Secao.Col = 3 Or tdb_Secao.Col = 4 Then
        CaracterValido KeyAscii, "V", tdb_Secao
    Else
        CaracterValido KeyAscii, "A", tdb_Secao
    End If
End Sub

Private Sub tdb_Secao_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If mblnPrimeiraVez Then
        If tdb_Secao.EOF Then Exit Sub
        mblnAlterando = True
       txtpkID.Text = tdb_Secao.Columns("PKId").Value
        LeDaTabelaParaObj gstrSecaoLogradouro, Me
        cbointUtilizacao.ListIndex = gintIndiceCBO(cbointUtilizacao, 2)
        CarregaMelhoramentos txtpkID.Text
    End If
End Sub

Private Sub txt_intAno_GotFocus()
    MarcaCampo txt_intAno
    tab_3dPasta.Tab = 1
End Sub

Private Sub txt_intAno_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txt_intAno
End Sub

Private Sub txtintLargura_GotFocus()
    MarcaCampo txtintLargura
    tab_3dPasta.Tab = 0
End Sub

Private Sub txtintLargura_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "V", txtintLargura
End Sub

Private Sub txtintLargura_LostFocus()
    txtintLargura = gvntConvVrDoSql(txtintLargura)
End Sub
 
'
'Private Sub txtNumero_KeyPress(KeyAscii As Integer)
'    CaracterValido KeyAscii, "A", txtNumero
'    Select Case KeyAscii
'        Case vbKeyBack
'            Exit Sub
'    End Select
'    If Len(txtNumero) = 5 Then
'        Select Case UCase(Chr(KeyAscii))
'            Case "D", "E", "X"
'                KeyAscii = Asc(UCase(Chr(KeyAscii)))
'            Case Else
'                KeyAscii = 0
'                Beep
'        End Select
'    Else
'        Select Case KeyAscii
'            Case vbKey0 To vbKey9
'                Exit Sub
'            Case Else
'                KeyAscii = 0
'                Beep
'        End Select
'    End If
'End Sub

Function blnDadosOk() As Boolean
    Select Case Right(mskstrInscricaoCadastral.ClipText, 1)
        Case "D", "E", "X"
        Case Else
            ExibeMensagem "O último caracter do número da seção tem que ser obrigatoriamente: X (toda extensão do logradouro), D (lado direito), E (lado esquerdo)"
            Exit Function
    End Select
    blnDadosOk = True
End Function

Sub GravaMelhoramentos(intCodSecao As Integer)

'******************************************************************************************
' Data: 02/05/2003
' Alteração: - Substituição do comando nativo GETDATE() do SQL Server pela variável
'            strGETDATE.
' Responsável: Everton Bianchini
'******************************************************************************************
    
    Dim intI As Integer
    DeletaMelhoramentos intCodSecao
    With lvw_Melhoramentos
        For intI = 1 To .ListItems.Count
            strSQL = ""
            strSQL = strSQL & "Insert Into " & gstrMelhoramentoDaSecaoDeLogradouro & " "
            strSQL = strSQL & "(intSecaoDeLogradouro, intMelhoramento, intAno, "
            strSQL = strSQL & "dtmDtAtualizacao, lngCodUsr "
            strSQL = strSQL & ") Values ("
            strSQL = strSQL & intCodSecao & ", "
            strSQL = strSQL & .ListItems(intI).Tag & ", "
            strSQL = strSQL & .ListItems(intI).SubItems(1) & ", "
'            strSql = strSql & "GETDATE()" & ", "
            strSQL = strSQL & strGETDATE & ", "
            strSQL = strSQL & glngCodUsr
            strSQL = strSQL & ")"
            Set gobjBanco = New clsBanco
            gobjBanco.Execute strSQL
        Next
    End With
End Sub

Sub DeletaMelhoramentos(intCodSecao As Integer)
    Dim strSQL As String

    strSQL = ""
    strSQL = strSQL & "Delete From " & gstrMelhoramentoDaSecaoDeLogradouro & " "
    strSQL = strSQL & "Where intSecaoDeLogradouro = " & intCodSecao
    
    Set gobjBanco = New clsBanco
    gobjBanco.Execute strSQL
End Sub

Sub VerificaMascaraInscricao()
    Dim adoResultado As ADODB.Recordset
    Dim strMascara   As String
    
    strMascara = ""
    
    strSQL = ""
    strSQL = strSQL & "Select * From " & gstrCampoDeInscricao & " "
    strSQL = strSQL & "Where intTipoDeInscricao = " & TYP_ACORDO
    strSQL = strSQL & "Order By intSequencia"
    
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
        With adoResultado
            Do While Not .EOF
                strMascara = strMascara & String(!intTamanho, "a") & gstrVerificaCampoNulo(!strSeparador)
                .MoveNext
            Loop
        End With
    End If
    mskstrInscricaoCadastral.Mask = strMascara
End Sub

Sub NovaSecaoDeLogradouro()
    lvw_Melhoramentos.ListItems.Clear
    txt_intAno = ""
    tab_3dPasta.Tab = 0
End Sub



'##########$$$$$$$$$$$$$$##########'

Sub PreencheComboValorDaSecao()
    Dim strSQL       As String
    Dim adoResultado As ADODB.Recordset
    
    strSQL = ""
    strSQL = strSQL & "Select PKId,dblValor ,strNomeDoValor "
    strSQL = strSQL & " From " & gstrTabelaDeValor & " "
    strSQL = strSQL & " Where intCodigoDaUtilizacao = 2"

    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
        With adoResultado
            Do While Not .EOF
                cbointValorDaSecao.AddItem gvntConvVrDoSql(!DBLVALOR) & "    - " & !strnomedovalor
                cbointValorDaSecao.ItemData(cbointValorDaSecao.NewIndex) = !Pkid
                .MoveNext
            Loop
        End With
    End If
End Sub

Function strQuerryRelatorio() As String

'******************************************************************************************
' Data: 05/05/2003
' Alteração: - Deslocamento dos JOINS utilizados na cláusula FROM para a cláusula WHERE.
'            Foi mantida a forma antiga para o SQL Server pois não era possível o
'            deslocamento completo devido à incompatibilidade entres os bancos.
' Responsável: Everton Bianchini
'******************************************************************************************

    Dim strSQL As String
    strSQL = ""
    strSQL = strSQL & "SELECT SL.PKId, SL.strInscricaoCadastral Inscricao, "
    strSQL = strSQL & "SL.intLargura Largura, RTRIM(TV.strNomeDoValor) AS NomeDoValor, "
    strSQL = strSQL & "TV.dblValor Valor, MS.intAno, TP.strSigla Sigla, "
    strSQL = strSQL & "TL.strDescricao TituloLogradouro, L.strDescricao Logradouro, "
    strSQL = strSQL & "M.strNomeDoMelhoramento Melhoramento "
        
    If (bytDBType = EDatabases.SQLServer) Then
        strSQL = strSQL & "FROM (((((tblSecaodeLogradouro SL "
        strSQL = strSQL & "LEFT JOIN " & gstrTabelaDeValor & " TV ON SL.intValorDaSecao = TV.PKId) "
        strSQL = strSQL & "LEFT JOIN " & gstrLogradouro & " L ON SL.intLogradouro = L.PKId) "
        strSQL = strSQL & "LEFT JOIN " & gstrMelhoramentoDaSecaoDeLogradouro & " MS ON SL.PKId = MS.intSecaoDeLogradouro) "
        strSQL = strSQL & "LEFT JOIN " & gstrMelhoramentoPublico & " M ON MS.intMelhoramento = M.PKId) "
        strSQL = strSQL & "LEFT JOIN " & gstrTipoLogradouro & " TP ON L.intTipoLogradouro = TP.PKId) "
        strSQL = strSQL & "LEFT JOIN " & gstrTituloLogradouro & " TL ON L.intTituloLogradouro = TL.PKId "
    
    ElseIf (bytDBType = EDatabases.Oracle) Then
        strSQL = strSQL & "FROM tblSecaodeLogradouro SL, "
        strSQL = strSQL & gstrTabelaDeValor & " TV, "
        strSQL = strSQL & gstrLogradouro & " L, "
        strSQL = strSQL & gstrMelhoramentoDaSecaoDeLogradouro & " MS, "
        strSQL = strSQL & gstrMelhoramentoPublico & " M, "
        strSQL = strSQL & gstrTipoLogradouro & " TP, "
        strSQL = strSQL & gstrTituloLogradouro & " TL, "
    
    End If
    
    If mblnAlterando = True Then
        strSQL = strSQL & " WHERE SL.PKId = " & tdb_Secao.Columns("PKId").Value
    End If

    If (bytDBType = EDatabases.Oracle) Then
        strSQL = strSQL & " AND SL.intValorDaSecao = TV.PKId " & strOUTJOracle
        strSQL = strSQL & " AND SL.intLogradouro = L.PKId " & strOUTJOracle
        strSQL = strSQL & " AND SL.PKId = MS.intSecaoDeLogradouro " & strOUTJOracle
        strSQL = strSQL & " AND MS.intMelhoramento = M.PKId " & strOUTJOracle
        strSQL = strSQL & " AND L.intTipoLogradouro = TP.PKId " & strOUTJOracle
        strSQL = strSQL & " AND L.intTituloLogradouro = TL.PKId " & strOUTJOracle
    
    End If

strQuerryRelatorio = strSQL
End Function
