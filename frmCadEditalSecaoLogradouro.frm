VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frmCadEditalSecaoLogradouro 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cadastro de Logradouros para o Edital"
   ClientHeight    =   6615
   ClientLeft      =   2265
   ClientTop       =   690
   ClientWidth     =   7815
   Icon            =   "frmCadEditalSecaoLogradouro.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6615
   ScaleWidth      =   7815
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab tab_3DPasta 
      Height          =   6555
      Left            =   60
      TabIndex        =   15
      Top             =   60
      Width           =   7740
      _ExtentX        =   13653
      _ExtentY        =   11562
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   529
      TabCaption(0)   =   "Editais"
      TabPicture(0)   =   "frmCadEditalSecaoLogradouro.frx":1042
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "tdb_TabelaEdital"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fra_Secao"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fra_Edital"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Texto do Edital"
      TabPicture(1)   =   "frmCadEditalSecaoLogradouro.frx":105E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Fra_Frame1"
      Tab(1).ControlCount=   1
      Begin VB.Frame fra_Edital 
         Caption         =   "Dados do Edital"
         Height          =   2235
         Left            =   180
         TabIndex        =   0
         Top             =   420
         Width           =   7335
         Begin VB.TextBox txtstrNomeDoEdital 
            Height          =   285
            Left            =   1395
            MaxLength       =   80
            TabIndex        =   3
            Top             =   1050
            Width           =   5790
         End
         Begin VB.TextBox txtdblCustoDaParcela 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1395
            MaxLength       =   12
            MultiLine       =   -1  'True
            TabIndex        =   6
            Top             =   1725
            Width           =   1125
         End
         Begin VB.Frame fra_bytTipo 
            Caption         =   "Tipo "
            Height          =   645
            Left            =   5430
            TabIndex        =   8
            Top             =   1380
            Width           =   1770
            Begin VB.OptionButton optBytTipo 
               Caption         =   "Obra"
               Height          =   195
               Index           =   1
               Left            =   1005
               TabIndex        =   10
               Top             =   285
               Width           =   645
            End
            Begin VB.OptionButton optBytTipo 
               Caption         =   "Serviço"
               Height          =   195
               Index           =   0
               Left            =   120
               TabIndex        =   9
               Top             =   285
               Width           =   855
            End
         End
         Begin VB.TextBox txtPKId 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1395
            Locked          =   -1  'True
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   705
            Width           =   1290
         End
         Begin VB.TextBox txtDtmDataDeInicio 
            Height          =   285
            Left            =   1395
            MaxLength       =   12
            MultiLine       =   -1  'True
            TabIndex        =   4
            Top             =   1380
            Width           =   1125
         End
         Begin VB.TextBox txtdtmDataDeTermino 
            Height          =   285
            Left            =   4110
            MaxLength       =   12
            MultiLine       =   -1  'True
            TabIndex        =   5
            Top             =   1395
            Width           =   1125
         End
         Begin VB.TextBox txtdblCustoDeTerceiros 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   4110
            MaxLength       =   12
            MultiLine       =   -1  'True
            TabIndex        =   7
            Top             =   1740
            Width           =   1125
         End
         Begin MSDataListLib.DataCombo dbcintTabelaDeEdital 
            CausesValidation=   0   'False
            Height          =   315
            Left            =   1395
            TabIndex        =   1
            Top             =   300
            Width           =   5790
            _ExtentX        =   10213
            _ExtentY        =   556
            _Version        =   393216
            IntegralHeight  =   0   'False
            Text            =   ""
         End
         Begin VB.Label lblintTabelaDeEdital 
            AutoSize        =   -1  'True
            Caption         =   "Edital"
            Height          =   195
            Left            =   885
            TabIndex        =   27
            Top             =   360
            Width           =   390
         End
         Begin VB.Label lblstrNomeDoEdital 
            AutoSize        =   -1  'True
            Caption         =   "Nome"
            Height          =   195
            Left            =   855
            TabIndex        =   26
            Top             =   1110
            Width           =   420
         End
         Begin VB.Label lbldblCustoDaParcela 
            AutoSize        =   -1  'True
            Caption         =   "Custo da Parcela"
            Height          =   195
            Left            =   60
            TabIndex        =   25
            Top             =   1770
            Width           =   1215
         End
         Begin VB.Label lblPKID 
            AutoSize        =   -1  'True
            Caption         =   "Código"
            Height          =   195
            Left            =   780
            TabIndex        =   24
            Top             =   750
            Width           =   495
         End
         Begin VB.Label lblDtmDataInicio 
            AutoSize        =   -1  'True
            Caption         =   "Data de Início"
            Height          =   195
            Left            =   255
            TabIndex        =   23
            Top             =   1425
            Width           =   1020
         End
         Begin VB.Label lblDtmDataTermino 
            AutoSize        =   -1  'True
            Caption         =   "Data de Término"
            Height          =   195
            Left            =   2805
            TabIndex        =   22
            Top             =   1440
            Width           =   1185
         End
         Begin VB.Label lbldblCustoDeTerceiros 
            AutoSize        =   -1  'True
            Caption         =   "Custo de Terceiros"
            Height          =   195
            Left            =   2655
            TabIndex        =   21
            Top             =   1785
            Width           =   1335
         End
      End
      Begin VB.Frame fra_Secao 
         Caption         =   "Logradouro/Seção"
         Height          =   1215
         Left            =   180
         TabIndex        =   11
         Top             =   2700
         Width           =   7335
         Begin MSDataListLib.DataCombo dbcintSecaoDeLogradouro 
            CausesValidation=   0   'False
            Height          =   315
            Left            =   1785
            TabIndex        =   13
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
            TabIndex        =   12
            Top             =   300
            Width           =   5130
            _ExtentX        =   9049
            _ExtentY        =   556
            _Version        =   393216
            IntegralHeight  =   0   'False
            Text            =   ""
         End
         Begin VB.Label lblintSecaoDeLogradouro 
            AutoSize        =   -1  'True
            Caption         =   "Seção de Logradouro"
            Height          =   195
            Left            =   120
            TabIndex        =   20
            Top             =   720
            Width           =   1545
         End
         Begin VB.Label lbl_intLogradouro 
            AutoSize        =   -1  'True
            Caption         =   "Logradouro"
            Height          =   195
            Left            =   855
            TabIndex        =   19
            Top             =   360
            Width           =   810
         End
      End
      Begin VB.Frame Fra_Frame1 
         Height          =   6000
         Left            =   -74820
         TabIndex        =   16
         Top             =   375
         Width           =   7410
         Begin VB.TextBox txtstrTextoDoEdital 
            Height          =   5625
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   17
            Top             =   240
            Width           =   7155
         End
      End
      Begin TrueOleDBGrid70.TDBGrid tdb_TabelaEdital 
         Height          =   2445
         Left            =   180
         TabIndex        =   14
         Top             =   3990
         Width           =   7365
         _ExtentX        =   12991
         _ExtentY        =   4313
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
         Splits(0)._ColumnProps(13)=   "Column(2).Width=8943"
         Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=8864"
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Seção de Logradouro"
         Height          =   195
         Left            =   -3240
         TabIndex        =   18
         Top             =   -2580
         Width           =   1545
      End
   End
End
Attribute VB_Name = "frmCadEditalSecaoLogradouro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub dbc_intLogradouro_Click(Area As Integer)
   DropDownDataCombo dbc_intLogradouro, Me, Area
End Sub

Private Sub dbc_intLogradouro_KeyDown(KeyCode As Integer, Shift As Integer)
   DropDownDataCombo dbc_intLogradouro, Me, , KeyCode, Shift
End Sub

Private Sub dbcintSecaoDeLogradouro_Click(Area As Integer)
   DropDownDataCombo dbcintSecaoDeLogradouro, Me, Area
End Sub

Private Sub dbcintSecaoDeLogradouro_KeyDown(KeyCode As Integer, Shift As Integer)
   DropDownDataCombo dbcintSecaoDeLogradouro, Me, , KeyCode, Shift
End Sub

Private Sub dbcintTabelaDeEdital_Click(Area As Integer)
   DropDownDataCombo dbcintTabelaDeEdital, Me, Area
End Sub

Private Sub dbcintTabelaDeEdital_KeyDown(KeyCode As Integer, Shift As Integer)
   DropDownDataCombo dbcintTabelaDeEdital, Me, , KeyCode, Shift
End Sub

Private Sub Form_Load()
TrocaCorObjeto txtPKId, True
TrocaCorObjeto txtstrNomeDoEdital, True
TrocaCorObjeto txtDtmDataDeInicio, True
TrocaCorObjeto txtdtmDataDeTermino, True
TrocaCorObjeto txtdblCustoDaParcela, True
TrocaCorObjeto txtdblCustoDeTerceiros, True
TrocaCorObjeto optbytTipo(0), True
TrocaCorObjeto optbytTipo(1), True

LeDaTabelaParaObj gstrTabelaDeEdital, dbcintTabelaDeEdital

End Sub

