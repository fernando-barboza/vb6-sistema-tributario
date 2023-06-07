VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "TODG7.OCX"
Begin VB.Form frmCadTabelaDeEditais 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Editais"
   ClientHeight    =   6720
   ClientLeft      =   1110
   ClientTop       =   2325
   ClientWidth     =   7860
   HelpContextID   =   33
   Icon            =   "CadTabelaDeEditais.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6720
   ScaleWidth      =   7860
   Begin TabDlg.SSTab tab_3DPasta 
      Height          =   6555
      Left            =   60
      TabIndex        =   9
      Top             =   60
      Width           =   7740
      _ExtentX        =   13653
      _ExtentY        =   11562
      _Version        =   393216
      Style           =   1
      TabHeight       =   529
      TabCaption(0)   =   "Editais"
      TabPicture(0)   =   "CadTabelaDeEditais.frx":1042
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblstrNomeDoEdital"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lbldblCustoDaParcela"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblPKID"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblDtmDataInicio"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblDtmDataTermino"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lbldblCustoDeTerceiros"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtstrNomeDoEdital"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtdblCustoDaParcela"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "fra_bytTipo"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtPKId"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtdtmDataDeInicio"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtdtmDataDeTermino"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtdblCustoDeTerceiros"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "tdb_TabelaEdital"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).ControlCount=   14
      TabCaption(1)   =   "Texto do Edital"
      TabPicture(1)   =   "CadTabelaDeEditais.frx":105E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Fra_Frame1"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Logradouro/Seção"
      TabPicture(2)   =   "CadTabelaDeEditais.frx":107A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fra_Secao"
      Tab(2).Control(1)=   "tdb_LogradouroSecao"
      Tab(2).ControlCount=   2
      Begin VB.Frame fra_Secao 
         Height          =   1215
         Left            =   -74790
         TabIndex        =   20
         Top             =   840
         Width           =   7335
         Begin MSDataListLib.DataCombo dbc_intSecaoDeLogradouro 
            CausesValidation=   0   'False
            Height          =   315
            Left            =   1785
            TabIndex        =   21
            Top             =   660
            Width           =   5130
            _ExtentX        =   9049
            _ExtentY        =   556
            _Version        =   393216
            IntegralHeight  =   0   'False
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dbc_intLogradouro 
            CausesValidation=   0   'False
            Height          =   315
            Left            =   1785
            TabIndex        =   22
            Top             =   300
            Width           =   5130
            _ExtentX        =   9049
            _ExtentY        =   556
            _Version        =   393216
            IntegralHeight  =   0   'False
            Text            =   ""
         End
         Begin VB.Label lbl_intLogradouro 
            AutoSize        =   -1  'True
            Caption         =   "Logradouro"
            Height          =   195
            Left            =   855
            TabIndex        =   24
            Top             =   360
            Width           =   810
         End
         Begin VB.Label lbl_intSecaoDeLogradouro 
            AutoSize        =   -1  'True
            Caption         =   "Seção de Logradouro"
            Height          =   195
            Left            =   120
            TabIndex        =   23
            Top             =   720
            Width           =   1545
         End
      End
      Begin TrueOleDBGrid70.TDBGrid tdb_TabelaEdital 
         Height          =   4275
         Left            =   120
         TabIndex        =   8
         Top             =   2130
         Width           =   7485
         _ExtentX        =   13203
         _ExtentY        =   7541
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
         Columns(1).DataField=   "PKID"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Nome do Edital"
         Columns(2).DataField=   "strNomeDoEdital"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   3
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
         Splits(0)._ColumnProps(0)=   "Columns.Count=3"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0).AllowSizing=0"
         Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
         Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(8)=   "Column(1).Width=3413"
         Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=3334"
         Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(13)=   "Column(2).Width=9208"
         Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=9128"
         Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(17)=   "Column(2).Order=3"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   0
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
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=7,.bold=0,.fontsize=825,.italic=0"
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
         _StyleDefs(42)  =   "Named:id=33:Normal"
         _StyleDefs(43)  =   ":id=33,.parent=0"
         _StyleDefs(44)  =   "Named:id=34:Heading"
         _StyleDefs(45)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(46)  =   ":id=34,.wraptext=-1"
         _StyleDefs(47)  =   "Named:id=35:Footing"
         _StyleDefs(48)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(49)  =   "Named:id=36:Selected"
         _StyleDefs(50)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(51)  =   "Named:id=37:Caption"
         _StyleDefs(52)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(53)  =   "Named:id=38:HighlightRow"
         _StyleDefs(54)  =   ":id=38,.parent=33,.bgcolor=&H80000014&,.fgcolor=&H80000012&"
         _StyleDefs(55)  =   "Named:id=39:EvenRow"
         _StyleDefs(56)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(57)  =   "Named:id=40:OddRow"
         _StyleDefs(58)  =   ":id=40,.parent=33"
         _StyleDefs(59)  =   "Named:id=41:RecordSelector"
         _StyleDefs(60)  =   ":id=41,.parent=34"
         _StyleDefs(61)  =   "Named:id=42:FilterBar"
         _StyleDefs(62)  =   ":id=42,.parent=33"
      End
      Begin VB.Frame Fra_Frame1 
         Height          =   6000
         Left            =   -74820
         TabIndex        =   17
         Top             =   390
         Width           =   7410
         Begin VB.TextBox txtstrTextoDoEdital 
            Height          =   5625
            Left            =   120
            MaxLength       =   4000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   19
            Top             =   240
            Width           =   7155
         End
      End
      Begin VB.TextBox txtdblCustoDeTerceiros 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4470
         MaxLength       =   12
         MultiLine       =   -1  'True
         TabIndex        =   4
         Top             =   1620
         Width           =   1125
      End
      Begin VB.TextBox txtdtmDataDeTermino 
         Height          =   285
         Left            =   4470
         MaxLength       =   12
         MultiLine       =   -1  'True
         TabIndex        =   2
         Top             =   1275
         Width           =   1125
      End
      Begin VB.TextBox txtdtmDataDeInicio 
         Height          =   285
         Left            =   1755
         MaxLength       =   12
         MultiLine       =   -1  'True
         TabIndex        =   1
         Top             =   1260
         Width           =   1125
      End
      Begin VB.TextBox txtPKId 
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1755
         Locked          =   -1  'True
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   585
         Width           =   1290
      End
      Begin VB.Frame fra_bytTipo 
         Caption         =   "Tipo "
         Height          =   645
         Left            =   5790
         TabIndex        =   5
         Top             =   1260
         Width           =   1770
         Begin VB.OptionButton optBytTipo 
            Caption         =   "Serviço"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   6
            Top             =   285
            Width           =   855
         End
         Begin VB.OptionButton optBytTipo 
            Caption         =   "Obra"
            Height          =   195
            Index           =   1
            Left            =   1005
            TabIndex        =   7
            Top             =   285
            Width           =   645
         End
      End
      Begin VB.TextBox txtdblCustoDaParcela 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1755
         MaxLength       =   12
         MultiLine       =   -1  'True
         TabIndex        =   3
         Top             =   1605
         Width           =   1125
      End
      Begin VB.TextBox txtstrNomeDoEdital 
         Height          =   285
         Left            =   1755
         MaxLength       =   80
         TabIndex        =   0
         Top             =   930
         Width           =   5790
      End
      Begin TrueOleDBGrid70.TDBGrid tdb_LogradouroSecao 
         Height          =   3720
         Left            =   -74805
         TabIndex        =   25
         Top             =   2415
         Width           =   7365
         _ExtentX        =   12991
         _ExtentY        =   6562
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
         Columns(1).DataField=   "intLogradouro"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Logradouro"
         Columns(2).DataField=   "strDescricao"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).DataField=   "PKIdSecao"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "Seção de Logradouro"
         Columns(4).DataField=   "strInscricaoCadastral"
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
         Splits(0)._ColumnProps(12)=   "Column(1).AllowSizing=0"
         Splits(0)._ColumnProps(13)=   "Column(1).Visible=0"
         Splits(0)._ColumnProps(14)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(15)=   "Column(2).Width=7858"
         Splits(0)._ColumnProps(16)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(17)=   "Column(2)._WidthInPix=7779"
         Splits(0)._ColumnProps(18)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(19)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(20)=   "Column(3).Width=2725"
         Splits(0)._ColumnProps(21)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(22)=   "Column(3)._WidthInPix=2646"
         Splits(0)._ColumnProps(23)=   "Column(3)._EditAlways=0"
         Splits(0)._ColumnProps(24)=   "Column(3).AllowSizing=0"
         Splits(0)._ColumnProps(25)=   "Column(3).Visible=0"
         Splits(0)._ColumnProps(26)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(27)=   "Column(4).Width=4604"
         Splits(0)._ColumnProps(28)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(29)=   "Column(4)._WidthInPix=4524"
         Splits(0)._ColumnProps(30)=   "Column(4)._EditAlways=0"
         Splits(0)._ColumnProps(31)=   "Column(4).Order=5"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   0
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
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=7,.bold=0,.fontsize=825,.italic=0"
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
         _StyleDefs(34)  =   "Splits(0).Columns(1).Style:id=50,.parent=13"
         _StyleDefs(35)  =   "Splits(0).Columns(1).HeadingStyle:id=47,.parent=14"
         _StyleDefs(36)  =   "Splits(0).Columns(1).FooterStyle:id=48,.parent=15"
         _StyleDefs(37)  =   "Splits(0).Columns(1).EditorStyle:id=49,.parent=17"
         _StyleDefs(38)  =   "Splits(0).Columns(2).Style:id=32,.parent=13"
         _StyleDefs(39)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
         _StyleDefs(40)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
         _StyleDefs(41)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
         _StyleDefs(42)  =   "Splits(0).Columns(3).Style:id=54,.parent=13"
         _StyleDefs(43)  =   "Splits(0).Columns(3).HeadingStyle:id=51,.parent=14"
         _StyleDefs(44)  =   "Splits(0).Columns(3).FooterStyle:id=52,.parent=15"
         _StyleDefs(45)  =   "Splits(0).Columns(3).EditorStyle:id=53,.parent=17"
         _StyleDefs(46)  =   "Splits(0).Columns(4).Style:id=46,.parent=13"
         _StyleDefs(47)  =   "Splits(0).Columns(4).HeadingStyle:id=43,.parent=14"
         _StyleDefs(48)  =   "Splits(0).Columns(4).FooterStyle:id=44,.parent=15"
         _StyleDefs(49)  =   "Splits(0).Columns(4).EditorStyle:id=45,.parent=17"
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Seção de Logradouro"
         Height          =   195
         Left            =   -3240
         TabIndex        =   18
         Top             =   -2580
         Width           =   1545
      End
      Begin VB.Label lbldblCustoDeTerceiros 
         AutoSize        =   -1  'True
         Caption         =   "Custo de Terceiros"
         Height          =   195
         Left            =   3015
         TabIndex        =   16
         Top             =   1665
         Width           =   1335
      End
      Begin VB.Label lblDtmDataTermino 
         AutoSize        =   -1  'True
         Caption         =   "Data de Término"
         Height          =   195
         Left            =   3165
         TabIndex        =   15
         Top             =   1320
         Width           =   1185
      End
      Begin VB.Label lblDtmDataInicio 
         AutoSize        =   -1  'True
         Caption         =   "Data de Início"
         Height          =   195
         Left            =   615
         TabIndex        =   14
         Top             =   1305
         Width           =   1020
      End
      Begin VB.Label lblPKID 
         AutoSize        =   -1  'True
         Caption         =   "Código"
         Height          =   195
         Left            =   1140
         TabIndex        =   13
         Top             =   630
         Width           =   495
      End
      Begin VB.Label lbldblCustoDaParcela 
         AutoSize        =   -1  'True
         Caption         =   "Custo da Parcela"
         Height          =   195
         Left            =   420
         TabIndex        =   12
         Top             =   1650
         Width           =   1215
      End
      Begin VB.Label lblstrNomeDoEdital 
         AutoSize        =   -1  'True
         Caption         =   "Nome"
         Height          =   195
         Left            =   1215
         TabIndex        =   11
         Top             =   990
         Width           =   420
      End
   End
End
Attribute VB_Name = "frmCadTabelaDeEditais"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim mblnAlterando   As Boolean
    Dim mobjGeral       As Object
    Dim mblnSelecionou  As Boolean
    Dim mblnClickOk     As Boolean
    Dim mblnPrimeiraVez As Boolean
    Dim mobjAux         As Object
    Dim blnOrdenacaoAsc As Boolean
    Dim bytOrdenacao    As Byte
    Dim blnPrimeiraVezLogradouroSecao As Boolean
    Dim blnAlterandoLogradouroSecao As Boolean

Private Sub dbc_intLogradouro_Click(Area As Integer)
    DropDownDataCombo dbc_intLogradouro, Me, Area
    If Area = 2 And dbc_intLogradouro.MatchedWithList Then
        dbc_intSecaoDeLogradouro.Tag = "SELECT PKId, strInscricaoCadastral FROM tblSecaoDeLogradouro WHERE intLogradouro = " & dbc_intLogradouro.BoundText & " ORDER BY strInscricaoCadastral" & ";strInscricaoCadastral"
        LeDaTabelaParaObj gstrSecaoLogradouro, dbc_intSecaoDeLogradouro, "SELECT PKId, strInscricaoCadastral FROM tblSecaoDeLogradouro WHERE intLogradouro = " & dbc_intLogradouro.BoundText & " ORDER BY strInscricaoCadastral"
    End If
End Sub

Private Sub dbc_intLogradouro_KeyDown(KeyCode As Integer, Shift As Integer)
   DropDownDataCombo dbc_intLogradouro, Me, , KeyCode, Shift
End Sub

Private Sub dbc_intSecaoDeLogradouro_Click(Area As Integer)
    DropDownDataCombo dbc_intSecaoDeLogradouro, Me, Area
End Sub

Private Sub dbc_intSecaoDeLogradouro_KeyDown(KeyCode As Integer, Shift As Integer)
   DropDownDataCombo dbc_intSecaoDeLogradouro, Me, , KeyCode, Shift
End Sub

'Private Function strQueryEdital(intSecaoDeLogradouro As Integer) As String
'Dim strSql As String
'    strSql = ""
'    strSql = strSql & "SELECT B.PKId, B.strNomeDoEdital FROM "
'    strSql = strSql & gstrSecaoLogradouro & " A, "
'    strSql = strSql & gstrTabelaDeEdital & " B "
'    strSql = strSql & " WHERE A.PKId = B.intSecaoDeLogradouro"
'    strSql = strSql & " AND A.PKId = " & intSecaoDeLogradouro
'strQueryEdital = strSql
'End Function

'Private Sub dbcintSecaoDeLogradouro_Click(Area As Integer)
'Dim strSql As String
'If Area = 2 Then
'    If dbcintSecaoDeLogradouro.MatchedWithList Then
'        strSql = strQueryEdital(dbcintSecaoDeLogradouro.BoundText)
'        Limpa_Controles Me, True, False, True, False, False
'        mblnAlterando = False
'        mblnPrimeiraVez = False
'        VerificaListaAutomatica gstrTabelaDeEdital, tdb_TabelaEdital, strSql
'    End If
'End If
'End Sub

Private Sub Form_Activate()
    gintCodSeguranca = 617
    VirificaGradeListView Me
'=============
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
'=============
End Sub

Private Sub Form_Deactivate()
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar
End Sub

Private Sub Form_Load()
    tab_3dPasta.TabEnabled(2) = False
    bytOrdenacao = 2: blnOrdenacaoAsc = True
    'LeDaTabelaParaObj gstrTabelaDeEdital, tdb_TabelaEdital, strQuery
    dbc_intLogradouro.Tag = gstrQueryLogradouro & ";L.Descricao"
    VerificaObjParaAplicar mobjAux
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar, gstrDeletar
    mblnSelecionou = False
End Sub

Private Sub optBytTipo_GotFocus(Index As Integer)
    tab_3dPasta.Tab = 0
End Sub

Private Sub optBytTipo_KeyPress(Index As Integer, KeyAscii As Integer)
    CaracterValido KeyAscii, "V", optbytTipo(Index)
End Sub

Private Sub tdb_TabelaEdital_Click()
    mblnPrimeiraVez = False
    mblnClickOk = True
    If glngQtdLinhaTDBGrid(tdb_TabelaEdital) = 1 Then
        tdb_TabelaEdital_RowColChange 0, 0
    End If
End Sub

Private Sub tdb_TabelaEdital_DblClick()
    MantemForm gstrAplicar
End Sub

Private Sub tdb_TabelaEdital_FilterChange()
    mblnPrimeiraVez = False
    gblnFilraCampos tdb_TabelaEdital
End Sub

Private Sub tdb_TabelaEdital_GotFocus()
    tab_3dPasta.Tab = 0
End Sub

Private Sub tdb_TabelaEdital_HeadClick(ByVal ColIndex As Integer)
   
   blnOrdenacaoAsc = IIf(bytOrdenacao = ColIndex, Not blnOrdenacaoAsc, True)
   
   bytOrdenacao = ColIndex: MantemForm gstrRefresh
   
End Sub

Private Sub tdb_TabelaEdital_KeyDown(KeyCode As Integer, Shift As Integer)
    mblnClickOk = False
End Sub

Private Sub tdb_TabelaEdital_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", tdb_TabelaEdital
End Sub

Private Sub tdb_TabelaEdital_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mblnClickOk = True
End Sub

Private Sub tdb_TabelaEdital_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    
    With tdb_TabelaEdital
        If (Not .EOF And Not .BOF) And mblnClickOk Then
            mblnClickOk = False
            txtPKId.Text = .Columns("PKID").Value
            
            LeDaTabelaParaObj gstrTabelaDeEdital, Me
'=============
            gCorLinhaSelecionada tdb_TabelaEdital
            
            HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrDeletar
            
            If mobjAux Is Nothing Then
                HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar
            Else

                HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrAplicar
            End If
            
            blnPrimeiraVezLogradouroSecao = False
            blnAlterandoLogradouroSecao = False
            LeDaTabelaParaObj gstrTabelaDeEdital, tdb_LogradouroSecao, strQueryLogradouroSecao
            dbc_intLogradouro.BoundText = ""
            dbc_intSecaoDeLogradouro.BoundText = ""
            Set dbc_intSecaoDeLogradouro.RowSource = Nothing
            tab_3dPasta.TabEnabled(2) = True
            
            mblnAlterando = True
            mblnSelecionou = True
        End If
    End With

End Sub

Public Sub MantemForm(ByVal strModoOperacao As String)
Dim varBookMark As Variant
Dim strSql As String
If strModoOperacao = UCase("IMPRIMIR") Then
    strSql = strQuery
    ToolBarGeral strModoOperacao, gstrTabelaDeEdital, mblnAlterando, tdb_TabelaEdital, Me, mobjAux, strSql, , rptCadEditais, strQuery
    Exit Sub
End If
    
    If UCase(strModoOperacao) = gstrPreencherLista Then
        PreencherListaDeOpcoes Me.ActiveControl
        Exit Sub
    End If
    
    If tab_3dPasta.Tab = 0 Or tab_3dPasta.Tab = 1 Then
        strSql = strQuery
        If UCase(strModoOperacao) = "SALVAR" Or UCase(strModoOperacao) = "DELETAR" Then
            mblnClickOk = False
        End If
        ToolBarGeral strModoOperacao, gstrTabelaDeEdital, mblnAlterando, tdb_TabelaEdital, Me, mobjGeral, strSql, strSql
        If Trim(txtPKId.Text) = "" Then
            tab_3dPasta.TabEnabled(2) = False
        End If
    Else
        If UCase(strModoOperacao) = "NOVO" Then
            blnPrimeiraVezLogradouroSecao = False
            blnAlterandoLogradouroSecao = False
            dbc_intLogradouro.BoundText = ""
            dbc_intSecaoDeLogradouro.BoundText = ""
            Set dbc_intSecaoDeLogradouro.RowSource = Nothing
        ElseIf UCase(strModoOperacao) = "SALVAR" Then
            SalvaAtualizaLogradouroSecao
        ElseIf UCase(strModoOperacao) = "DELETAR" Then
            DeletaLogradouroSecao
        End If
    End If
    
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrDeletar

End Sub

Private Sub SalvaAtualizaLogradouroSecao()

'******************************************************************************************
' Data: 04/04/2003
' Alteração: - Substituição do comando nativo GETDATE() do SQL Server pela variável
'        strGETDATE.
' Responsável: Everton Bianchini
'******************************************************************************************
    
    Dim strSql As String
    Dim intPosicao As Integer
    
    If Not dbc_intLogradouro.MatchedWithList Then
        ExibeMensagem "Selecione um Logradouro "
        dbc_intLogradouro.SetFocus
        Exit Sub
    ElseIf Not dbc_intSecaoDeLogradouro.MatchedWithList Then
        ExibeMensagem "Selecione uma Seção de Logradouro "
        dbc_intSecaoDeLogradouro.SetFocus
        Exit Sub
    End If
    
    strSql = ""
    intPosicao = 1
    
    If Not blnAlterandoLogradouroSecao Then
        If gblnExclusaoGravacaoOk("I") Then
           strSql = "INSERT INTO " & gstrEditalSecaoLogradouro & " (intTabelaDeEdital, "
           strSql = strSql & "intSecaoDeLogradouro, "
           strSql = strSql & "dtmDtAtualizacao,   "
           strSql = strSql & " lngCodUsr) "
           strSql = strSql & " VALUES ( " & txtPKId.Text
           strSql = strSql & ", " & dbc_intSecaoDeLogradouro.BoundText
'           strSql = strSql & ", getdate(), " & glngCodUsr & " )"
           strSql = strSql & ", " & strGETDATE & ", " & glngCodUsr & " )"
        End If
    Else
        If gblnExclusaoGravacaoOk("A") Then
            intPosicao = tdb_LogradouroSecao.Bookmark
            strSql = "UPDATE " & gstrEditalSecaoLogradouro & " SET "
            strSql = strSql & " intTabelaDeEdital = " & txtPKId.Text & ", "
            strSql = strSql & " intSecaoDeLogradouro = " & dbc_intSecaoDeLogradouro.BoundText & ", "
'            strSql = strSql & " dtmDtAtualizacao = getdate(), "
            strSql = strSql & " dtmDtAtualizacao = " & strGETDATE & ", "
            strSql = strSql & " lngCodUsr = " & glngCodUsr
            strSql = strSql & " WHERE "
            strSql = strSql & " PKId = " & tdb_LogradouroSecao.Columns("PKId").Value
        End If
    End If
    
    If strSql <> "" Then
        Set gobjBanco = New clsBanco
        If gobjBanco.Execute(strSql) Then
            blnPrimeiraVezLogradouroSecao = False
            blnAlterandoLogradouroSecao = False
            LeDaTabelaParaObj gstrTabelaDeEdital, tdb_LogradouroSecao, strQueryLogradouroSecao
            dbc_intLogradouro.BoundText = ""
            dbc_intSecaoDeLogradouro.BoundText = ""
            Set dbc_intSecaoDeLogradouro.RowSource = Nothing
            tdb_LogradouroSecao.Bookmark = intPosicao
        End If
    End If
    
End Sub

Private Sub DeletaLogradouroSecao()
    Dim strSql As String
    
    If gblnExclusaoGravacaoOk("E") Then
        
        strSql = strSql & " DELETE " & gstrEditalSecaoLogradouro & " WHERE "
        strSql = strSql & " PKId = " & tdb_LogradouroSecao.Columns("PKId").Value
        Set gobjBanco = New clsBanco
        If gobjBanco.Execute(strSql) Then
            blnPrimeiraVezLogradouroSecao = False
            blnAlterandoLogradouroSecao = False
            LeDaTabelaParaObj gstrTabelaDeEdital, tdb_LogradouroSecao, strQueryLogradouroSecao
            dbc_intLogradouro.BoundText = ""
            dbc_intSecaoDeLogradouro.BoundText = ""
            Set dbc_intSecaoDeLogradouro.RowSource = Nothing
        End If
    End If
End Sub


Private Sub tdb_LogradouroSecao_Click()
    blnPrimeiraVezLogradouroSecao = True
End Sub

Private Sub tdb_LogradouroSecao_FilterChange()
    blnPrimeiraVezLogradouroSecao = False
    gblnFilraCampos tdb_LogradouroSecao
End Sub

Private Sub tdb_LogradouroSecao_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", tdb_LogradouroSecao
End Sub

Private Sub tdb_LogradouroSecao_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    
    With tdb_LogradouroSecao
        If Not .EOF And Not .BOF Then
            
            If blnPrimeiraVezLogradouroSecao Then
                PreencherListaDeOpcoes dbc_intLogradouro, .Columns("intLogradouro").Value
                dbc_intLogradouro.BoundText = .Columns("intLogradouro").Value
                dbc_intLogradouro_Click (2)
                dbc_intSecaoDeLogradouro.BoundText = .Columns("PKIdSecao").Value
                
                gCorLinhaSelecionada tdb_LogradouroSecao
                
                HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrDeletar
                                
                blnAlterandoLogradouroSecao = True
            End If
        End If
    End With

End Sub


Private Sub txtdblCustoDeTerceiros_GotFocus()
    MarcaCampo txtdblCustoDeTerceiros
    tab_3dPasta.Tab = 0
End Sub

Private Sub txtdblCustoDeTerceiros_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "V", txtdblCustoDeTerceiros
End Sub

Private Sub txtdblCustoDaParcela_GotFocus()
    MarcaCampo txtdblCustoDaParcela
    tab_3dPasta.Tab = 0
End Sub

Private Sub txtdblCustoDaParcela_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "V", txtdblCustoDaParcela
End Sub

Private Sub txtdblCustoDaParcela_LostFocus()
    txtdblCustoDaParcela = gvntConvVrDoSql(txtdblCustoDaParcela)
End Sub

Private Sub txtdblCustoDeTerceiros_LostFocus()
    txtdblCustoDeTerceiros = gvntConvVrDoSql(txtdblCustoDeTerceiros)
End Sub

Private Sub txtdtmDataDeTermino_GotFocus()
    MarcaCampo txtdtmDataDeTermino
    tab_3dPasta.Tab = 0
End Sub

Private Sub txtdtmDataDeTermino_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "D", txtdtmDataDeTermino
End Sub

Private Sub txtdtmDataDeTermino_LostFocus()
    txtdtmDataDeTermino = gstrDataFormatada(txtdtmDataDeTermino)
End Sub
Private Sub txtDtmDataDeInicio_GotFocus()
    MarcaCampo txtDtmDataDeInicio
    tab_3dPasta.Tab = 0
End Sub

Private Sub txtDtmDataDeInicio_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "D", txtDtmDataDeInicio
End Sub

Private Sub txtDtmDataDeInicio_LostFocus()
    txtDtmDataDeInicio = gstrDataFormatada(txtDtmDataDeInicio)
End Sub

Private Sub txtPKId_GotFocus()
tab_3dPasta.Tab = 0
End Sub

Private Sub txtPKId_KeyPress(KeyAscii As Integer)
 CaracterValido KeyAscii, " ", txtPKId
End Sub

Private Sub txtstrNomeDoEdital_GotFocus()
    MarcaCampo txtstrNomeDoEdital
    tab_3dPasta.Tab = 0
End Sub

Private Sub txtstrNomeDoEdital_KeyPress(KeyAscii As Integer)
   CaracterValido KeyAscii, "A", txtstrNomeDoEdital
End Sub

Private Sub txtstrTextoDoEdital_GotFocus()
    MarcaCampo txtstrTextoDoEdital
    tab_3dPasta.Tab = 1
End Sub

Private Sub txtstrTextoDoEdital_KeyPress(KeyAscii As Integer)
If KeyAscii <> vbKeyReturn Then
    CaracterValido KeyAscii, "A", txtstrTextoDoEdital
End If
End Sub


'************************** Query *********************************

Private Function strQuery() As String
    Dim strSql  As String
    strSql = ""
    strSql = strSql & " SELECT PKId, strNomeDoEdital FROM "
    strSql = strSql & gstrTabelaDeEdital
    
    Select Case bytOrdenacao
   
      Case Is = 1
         strSql = strSql & " ORDER BY PKId" & IIf(blnOrdenacaoAsc, " ASC", " DESC")
      Case Is = 2
         strSql = strSql & " ORDER BY strNomeDoEdital" & IIf(blnOrdenacaoAsc, " ASC", " DESC")
   End Select
    
    strQuery = strSql
End Function

Private Function strQueryLogradouroSecao() As String

'******************************************************************************************
' Data: 04/04/2003
' Alteração: - Substituição do comando nativo ISNULL() do SQL Server pela variável
'            gstrISNULL.
' Responsável: Everton Bianchini
'------------------------------------------------------------------------------------------
' Data: 04/04/2003
' Alteração: - Substituição do comando nativo de concatenação de strings (+) do SQL Server
'        pela variável strCONCAT.
' Responsável: Everton Bianchini
'------------------------------------------------------------------------------------------
' Data: 08/04/2003
' Alteração: - Deslocamento dos JOINS utilizados na cláusula FROM para a cláusula WHERE.
' Responsável: Everton Bianchini
'******************************************************************************************
    
    Dim strSql  As String
    strSql = ""
    strSql = strSql & " SELECT ESL.PKId, SEC.PKId AS PKIdSecao, SEC.intLogradouro, "
'    strSql = strSql & "RTRIM(ISNULL(TLO.strSigla, '')) + ' ' + LTRIM(ISNULL(TIT.strDescricao,''))  "
    strSql = strSql & "LO.strDescricao " & strCONCAT & " ', ' " & strCONCAT & "RTRIM(" & gstrISNULL("TLO.strSigla", "''") & ") " & strCONCAT & " ' ' " & strCONCAT & " LTRIM(" & gstrISNULL("TIT.strDescricao", "''") & ")  AS strDescricao,"
'    strSql = strSql & " + ' ' +  LO.strDescricao AS strDescricao, "
    strSql = strSql & " SEC.strInscricaoCadastral FROM "
    strSql = strSql & gstrEditalSecaoLogradouro & " ESL, "
    strSql = strSql & gstrSecaoLogradouro & " SEC, "
    strSql = strSql & gstrLogradouro & " LO "
'    strSql = strSql & "LEFT JOIN  " & gstrTituloLogradouro & " TIT "
    strSql = strSql & ", " & gstrTituloLogradouro & " TIT "
'    strSql = strSql & "ON LO.intTituloLogradouro = TIT.PKId "
'    strSql = strSql & "LEFT JOIN " & gstrTipoLogradouro & " TLO "
    strSql = strSql & ", " & gstrTipoLogradouro & " TLO "
'    strSql = strSql & "ON LO.intTipoLogradouro = TLO.PKId "
    strSql = strSql & " WHERE ESL.intTabelaDeEdital = " & txtPKId.Text & " AND "
    strSql = strSql & " SEC.PKId = ESL.intSecaoDeLogradouro AND "
    strSql = strSql & " LO.PKId = SEC.intLogradouro "
    strSql = strSql & " AND LO.intTituloLogradouro " & strOUTJSQLServer & "= TIT.PKId " & strOUTJOracle
    strSql = strSql & " AND LO.intTipoLogradouro " & strOUTJSQLServer & "= TLO.PKId " & strOUTJOracle
    strSql = strSql & " AND LO.Dtmdtexclusao is null "
    strSql = strSql & " ORDER BY ESL.PKId "
    strQueryLogradouroSecao = strSql
End Function


'************************** Query *********************************
