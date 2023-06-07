VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frmCadConvenio 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Convênio / Entidade / Fundo"
   ClientHeight    =   7980
   ClientLeft      =   795
   ClientTop       =   2070
   ClientWidth     =   9660
   HelpContextID   =   247
   Icon            =   "CadConvenio.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7980
   ScaleWidth      =   9660
   Begin VB.TextBox txtPKId 
      Height          =   270
      Left            =   6630
      TabIndex        =   52
      Top             =   120
      Visible         =   0   'False
      Width           =   645
   End
   Begin TabDlg.SSTab tab_3DPasta 
      Height          =   7815
      Left            =   90
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   90
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   13785
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "Convênio"
      TabPicture(0)   =   "CadConvenio.frx":1042
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblstrDescricao"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lbldtmData"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblintObjeto"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblintOrgao"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lbldblValor"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lbldblValorContraPartida"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lblintContaContabil"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lblstrCodigo"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lblstrObservacao"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "LblstrContribuinte"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "LblstrPrevRec"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtintPrevisaoDaReceita"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "dbc_strNomeContribuinte"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "dbc_strOrgao"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "dbc_strTipLegisla"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "dbc_strTipConvenio"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "dbcintObjeto"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "tdb_Lista"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "txtdtmData"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "cmd_OrgaoFinanceiro"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "txtdblValor"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "txtdblValorContrapartida"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "cbointContaContabil"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "cbo_ContaContabil"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "cmd_PlanoConta"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "cmd_Conta"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "fra_DataAplicacao"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "fra_bytTipo"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "cmd_Objetico"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "txtstrObservacao"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "txtstrCodigo"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "txtstrDescricao"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "txt_intCodTipConvenio"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "txtintContribuinte"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "txt_intCodTipLegisla"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "txtintOrgao"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "Cmd_Contribuinte"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "Cmd_PrevReceita"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "cbo_CodOrcstrDescricao"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "cbo_CodOrcstrCodigo"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "optbytTipoConvenio(0)"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "optbytTipoConvenio(1)"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "optbytTipoConvenio(2)"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).ControlCount=   43
      Begin VB.OptionButton optbytTipoConvenio 
         Caption         =   "Fundo"
         Height          =   285
         Index           =   2
         Left            =   6330
         TabIndex        =   5
         Top             =   465
         Width           =   1575
      End
      Begin VB.OptionButton optbytTipoConvenio 
         Caption         =   "Entidade"
         Height          =   285
         Index           =   1
         Left            =   4785
         TabIndex        =   4
         Top             =   465
         Width           =   1575
      End
      Begin VB.OptionButton optbytTipoConvenio 
         Caption         =   "Convênio"
         Height          =   285
         Index           =   0
         Left            =   3240
         TabIndex        =   3
         Top             =   465
         Width           =   1575
      End
      Begin VB.ComboBox cbo_CodOrcstrCodigo 
         Height          =   315
         Left            =   1740
         OLEDragMode     =   1  'Automatic
         Sorted          =   -1  'True
         TabIndex        =   14
         Top             =   1710
         Width           =   1695
      End
      Begin VB.ComboBox cbo_CodOrcstrDescricao 
         Height          =   315
         Left            =   3480
         OLEDragMode     =   1  'Automatic
         Sorted          =   -1  'True
         TabIndex        =   15
         Top             =   1710
         Width           =   5355
      End
      Begin VB.CommandButton Cmd_PrevReceita 
         Height          =   300
         Left            =   8895
         Picture         =   "CadConvenio.frx":105E
         Style           =   1  'Graphical
         TabIndex        =   16
         TabStop         =   0   'False
         Tag             =   "193"
         ToolTipText     =   "Clique para cadastrar previsão da receita"
         Top             =   1710
         Width           =   330
      End
      Begin VB.CommandButton Cmd_Contribuinte 
         Height          =   300
         Left            =   8895
         Picture         =   "CadConvenio.frx":13E8
         Style           =   1  'Graphical
         TabIndex        =   11
         TabStop         =   0   'False
         Tag             =   "193"
         ToolTipText     =   "Clique para cadastrar contribuinte"
         Top             =   1290
         Width           =   330
      End
      Begin VB.TextBox txtintOrgao 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1740
         MaxLength       =   10
         TabIndex        =   29
         Top             =   3015
         Width           =   1215
      End
      Begin VB.TextBox txt_intCodTipLegisla 
         Alignment       =   1  'Right Justify
         BackColor       =   &H000000FF&
         Height          =   315
         Left            =   540
         MaxLength       =   10
         TabIndex        =   50
         Top             =   8400
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtintContribuinte 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1740
         MaxLength       =   10
         TabIndex        =   9
         Top             =   1290
         Width           =   1215
      End
      Begin VB.TextBox txt_intCodTipConvenio 
         Alignment       =   1  'Right Justify
         BackColor       =   &H000000FF&
         Height          =   315
         Left            =   540
         MaxLength       =   10
         TabIndex        =   48
         Top             =   7950
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtstrDescricao 
         Height          =   315
         Left            =   1740
         MaxLength       =   100
         TabIndex        =   7
         Top             =   870
         Width           =   7485
      End
      Begin VB.TextBox txtstrCodigo 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1740
         MaxLength       =   10
         TabIndex        =   2
         Top             =   450
         Width           =   1215
      End
      Begin VB.TextBox txtstrObservacao 
         Height          =   315
         Left            =   1740
         MaxLength       =   40
         TabIndex        =   37
         Top             =   3900
         Width           =   7485
      End
      Begin VB.CommandButton cmd_Objetico 
         Height          =   300
         Left            =   8895
         Picture         =   "CadConvenio.frx":1772
         Style           =   1  'Graphical
         TabIndex        =   20
         TabStop         =   0   'False
         Tag             =   "193"
         ToolTipText     =   "Clique para cadastrar objetivo"
         Top             =   2160
         Width           =   330
      End
      Begin VB.Frame fra_bytTipo 
         Caption         =   " Destinado "
         Height          =   705
         Left            =   5160
         TabIndex        =   43
         Top             =   4320
         Width           =   3945
         Begin VB.OptionButton optbytTipo 
            Caption         =   "Outros"
            Height          =   285
            Index           =   2
            Left            =   2880
            TabIndex        =   46
            Top             =   270
            Width           =   825
         End
         Begin VB.OptionButton optbytTipo 
            Caption         =   "Saúde"
            Height          =   285
            Index           =   1
            Left            =   1665
            TabIndex        =   45
            Top             =   270
            Width           =   825
         End
         Begin VB.OptionButton optbytTipo 
            Caption         =   "Educação"
            Height          =   285
            Index           =   0
            Left            =   210
            TabIndex        =   44
            Top             =   270
            Width           =   1065
         End
      End
      Begin VB.Frame fra_DataAplicacao 
         Caption         =   " Data de aplicação "
         Height          =   705
         Left            =   885
         TabIndex        =   38
         Top             =   4320
         Width           =   3885
         Begin VB.TextBox txtdtmDataAplicacaoInicial 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   810
            TabIndex        =   40
            Top             =   270
            Width           =   1005
         End
         Begin VB.TextBox txtdtmDataAplicacaoFinal 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2550
            TabIndex        =   42
            Top             =   270
            Width           =   1005
         End
         Begin VB.Label lbldtmDataAplicacaoInicial 
            AutoSize        =   -1  'True
            Caption         =   "Inicial"
            Height          =   195
            Left            =   240
            TabIndex        =   39
            Top             =   315
            Width           =   405
         End
         Begin VB.Label lbldtmDataAplicacaoFinal 
            AutoSize        =   -1  'True
            Caption         =   "Final"
            Height          =   195
            Left            =   2070
            TabIndex        =   41
            Top             =   315
            Width           =   330
         End
      End
      Begin VB.CommandButton cmd_Conta 
         Height          =   300
         Left            =   8895
         Picture         =   "CadConvenio.frx":1AFC
         Style           =   1  'Graphical
         TabIndex        =   35
         TabStop         =   0   'False
         Tag             =   "322"
         ToolTipText     =   "Clique para cadastar conta"
         Top             =   3450
         Width           =   330
      End
      Begin VB.CommandButton cmd_PlanoConta 
         Height          =   300
         Left            =   9990
         Picture         =   "CadConvenio.frx":1E86
         Style           =   1  'Graphical
         TabIndex        =   28
         TabStop         =   0   'False
         ToolTipText     =   "Clique para cadastar plano de conta"
         Top             =   3000
         Width           =   360
      End
      Begin VB.ComboBox cbo_ContaContabil 
         Height          =   315
         Left            =   3480
         OLEDragMode     =   1  'Automatic
         Sorted          =   -1  'True
         TabIndex        =   34
         ToolTipText     =   "Histórico padrão"
         Top             =   3450
         Width           =   5355
      End
      Begin VB.ComboBox cbointContaContabil 
         Height          =   315
         Left            =   1740
         OLEDragMode     =   1  'Automatic
         Sorted          =   -1  'True
         TabIndex        =   33
         ToolTipText     =   "Histórico padrão"
         Top             =   3450
         Width           =   1695
      End
      Begin VB.TextBox txtdblValorContrapartida 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   5100
         MaxLength       =   18
         TabIndex        =   24
         Top             =   2595
         Width           =   1605
      End
      Begin VB.TextBox txtdblValor 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1740
         MaxLength       =   18
         TabIndex        =   22
         Top             =   2595
         Width           =   1605
      End
      Begin VB.CommandButton cmd_OrgaoFinanceiro 
         Height          =   300
         Left            =   8895
         Picture         =   "CadConvenio.frx":2210
         Style           =   1  'Graphical
         TabIndex        =   31
         TabStop         =   0   'False
         Tag             =   "193"
         ToolTipText     =   "Clique para cadastar órgão"
         Top             =   3015
         Width           =   330
      End
      Begin VB.TextBox txtdtmData 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   8250
         MaxLength       =   10
         TabIndex        =   26
         Top             =   2595
         Width           =   975
      End
      Begin TrueOleDBGrid70.TDBGrid tdb_Lista 
         Height          =   2415
         Left            =   750
         TabIndex        =   47
         Top             =   5220
         Width           =   8475
         _ExtentX        =   14949
         _ExtentY        =   4260
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
         Columns(1).DataField=   "strCodigo"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Descrição"
         Columns(2).DataField=   "strDescricao"
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
         Splits(0)._ColumnProps(1)=   "Column(0).Width=1138"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1058"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0).AllowSizing=0"
         Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
         Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(8)=   "Column(1).Width=2328"
         Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=2249"
         Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(13)=   "Column(2).Width=12091"
         Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=12012"
         Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(17)=   "Column(2).Order=3"
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
         _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
         _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
         _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
         _StyleDefs(22)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(23)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
         _StyleDefs(24)  =   "Splits(0).Style:id=13,.parent=1"
         _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
         _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
         _StyleDefs(27)  =   "Splits(0).FooterStyle:id=15,.parent=3"
         _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
         _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
         _StyleDefs(30)  =   "Splits(0).EditorStyle:id=17,.parent=7"
         _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
         _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
         _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
         _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
         _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
         _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
         _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
         _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=46,.parent=13"
         _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=43,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=44,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=45,.parent=17"
         _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=32,.parent=13"
         _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
         _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
         _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
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
      Begin MSDataListLib.DataCombo dbcintObjeto 
         Height          =   315
         Left            =   1740
         TabIndex        =   19
         Top             =   2145
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dbc_strTipConvenio 
         Height          =   315
         Left            =   1890
         TabIndex        =   49
         Top             =   7950
         Visible         =   0   'False
         Width           =   6345
         _ExtentX        =   11192
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   255
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dbc_strTipLegisla 
         Height          =   315
         Left            =   2130
         TabIndex        =   51
         Top             =   8430
         Visible         =   0   'False
         Width           =   6345
         _ExtentX        =   11192
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   255
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dbc_strOrgao 
         Height          =   315
         Left            =   3090
         TabIndex        =   30
         Top             =   3015
         Width           =   5745
         _ExtentX        =   10134
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dbc_strNomeContribuinte 
         Height          =   315
         Left            =   3120
         TabIndex        =   10
         Top             =   1290
         Width           =   5745
         _ExtentX        =   10134
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin VB.TextBox txtintPrevisaoDaReceita 
         Height          =   345
         Left            =   1740
         TabIndex        =   13
         Top             =   1695
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fonte de Recurso"
         Height          =   195
         Left            =   -360
         TabIndex        =   17
         Top             =   2190
         Width           =   90
      End
      Begin VB.Label LblstrPrevRec 
         AutoSize        =   -1  'True
         Caption         =   "Previsão da Receita"
         Height          =   195
         Left            =   225
         TabIndex        =   12
         Top             =   1770
         Width           =   1440
      End
      Begin VB.Label LblstrContribuinte 
         AutoSize        =   -1  'True
         Caption         =   "Contribuinte"
         Height          =   195
         Left            =   825
         TabIndex        =   8
         Top             =   1350
         Width           =   840
      End
      Begin VB.Label lblstrObservacao 
         AutoSize        =   -1  'True
         Caption         =   "Observação"
         Height          =   195
         Left            =   795
         TabIndex        =   36
         Top             =   3960
         Width           =   870
      End
      Begin VB.Label lblstrCodigo 
         AutoSize        =   -1  'True
         Caption         =   "Código"
         Height          =   195
         Left            =   1170
         TabIndex        =   1
         Top             =   510
         Width           =   495
      End
      Begin VB.Label lblintContaContabil 
         AutoSize        =   -1  'True
         Caption         =   "Conta"
         Height          =   195
         Left            =   1245
         TabIndex        =   32
         Top             =   3510
         Width           =   420
      End
      Begin VB.Label lbldblValorContraPartida 
         AutoSize        =   -1  'True
         Caption         =   "Contrapartida"
         Height          =   195
         Left            =   4080
         TabIndex        =   23
         Top             =   2655
         Width           =   945
      End
      Begin VB.Label lbldblValor 
         AutoSize        =   -1  'True
         Caption         =   "Valor"
         Height          =   195
         Left            =   1305
         TabIndex        =   21
         Top             =   2655
         Width           =   360
      End
      Begin VB.Label lblintOrgao 
         AutoSize        =   -1  'True
         Caption         =   "Órgão"
         Height          =   195
         Left            =   1230
         TabIndex        =   27
         Top             =   3075
         Width           =   435
      End
      Begin VB.Label lblintObjeto 
         AutoSize        =   -1  'True
         Caption         =   "Objeto"
         Height          =   195
         Left            =   1200
         TabIndex        =   18
         Top             =   2205
         Width           =   465
      End
      Begin VB.Label lbldtmData 
         AutoSize        =   -1  'True
         Caption         =   "Data"
         Height          =   195
         Left            =   7695
         TabIndex        =   25
         Top             =   2655
         Width           =   345
      End
      Begin VB.Label lblstrDescricao 
         AutoSize        =   -1  'True
         Caption         =   "Descrição"
         Height          =   195
         Left            =   945
         TabIndex        =   6
         Top             =   930
         Width           =   720
      End
   End
End
Attribute VB_Name = "frmCadConvenio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim mblnAlterando   As Boolean
    Dim mobjAux         As Object
    Dim mblnselecionou  As Boolean
    Dim mblnClickOk     As Boolean
    Dim ValTxtStrCodigo As String
    Dim ValTxtStrDescricao As String
    Dim mblnCarregaFormConta As Boolean
    Dim mblncarrega As Boolean

Private Sub cbo_ContaContabil_GotFocus()
    If mblnCarregaFormConta = True Then
        mblnCarregaFormConta = False
        If cbo_ContaContabil.ListIndex = -1 Then cbointContaContabil.ListIndex = -1
    End If
End Sub

Private Sub cbo_ContaContabil_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub cbointContaContabil_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub cbo_CodOrcstrCodigo_Click()
    Me.cbo_CodOrcstrDescricao.ListIndex = gintIndiceCBO(Me.cbo_CodOrcstrDescricao, gstrItemData(Me.cbo_CodOrcstrCodigo))
End Sub

Private Sub cbo_CodOrcstrDescricao_Click()
    Dim tempIndice As Integer
    
    tempIndice = Me.cbo_CodOrcstrDescricao.ListIndex
    Me.cbo_CodOrcstrCodigo.ListIndex = gintIndiceCBO(Me.cbo_CodOrcstrCodigo, gstrItemData(Me.cbo_CodOrcstrDescricao))
    RetornaCodigo gstrPrevisaoDaReceita, Me.txtintPrevisaoDaReceita, "PKID", "intCodigoOrcamentario", gstrItemData(Me.cbo_CodOrcstrDescricao)
   If Me.cbo_CodOrcstrCodigo.ListIndex = -1 Then
        LePrevisaoReceitaGeral Me.cbo_CodOrcstrCodigo, Me.cbo_CodOrcstrDescricao
        Me.cbo_CodOrcstrDescricao.ListIndex = tempIndice
        Me.cbo_CodOrcstrCodigo.ListIndex = gintIndiceCBO(Me.cbo_CodOrcstrCodigo, gstrItemData(Me.cbo_CodOrcstrDescricao))
        RetornaCodigo gstrPrevisaoDaReceita, Me.txtintPrevisaoDaReceita, "PKID", "intCodigoOrcamentario", gstrItemData(Me.cbo_CodOrcstrDescricao)
   End If
    
   
   If Me.cbo_CodOrcstrDescricao.ListIndex = -1 Then Me.cbo_CodOrcstrCodigo.ListIndex = -1

End Sub

Private Sub cmd_Conta_Click()
    mblnCarregaFormConta = True
    CarregaForm frmCadPlanoConta, cbo_ContaContabil, strQueryAplicar
End Sub

Private Sub cmd_Contribuinte_Click()
    CarregaForm frmCadContribuinte, dbc_strNomeContribuinte
End Sub

Private Sub cmd_Objetico_Click()
    CarregaForm frmCadObjetivo, dbcintObjeto
End Sub

Private Sub cmd_OrgaoFinanceiro_Click()
    CarregaForm frmCadOrgao, dbc_strOrgao
End Sub

Private Function strQueryAplicar() As String
    Dim strSQL As String
    strSQL = ""
    strSQL = strSQL & "SELECT PKId, strDescricao "
    strSQL = strSQL & "FROM "
    strSQL = strSQL & gstrPlanoConta & " "
    strSQL = strSQL & "WHERE ABS(blnFinanceira) = 1 "
    strSQL = strSQL & "AND ABS(blnAnalitica) = 1"
    strQueryAplicar = strSQL
End Function

Private Sub Cmd_PrevReceita_Click()
    CarregaForm frmCadPrevisaoDaReceita, cbo_CodOrcstrDescricao, strQueryAplicar
End Sub

Private Sub dbcintObjeto_Click(Area As Integer)
   DropDownDataCombo dbcintObjeto, Me, Area
End Sub

Private Sub dbcintObjeto_KeyDown(KeyCode As Integer, Shift As Integer)
   DropDownDataCombo dbcintObjeto, Me, , KeyCode, Shift
End Sub

Private Sub dbcintObjeto_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub dbc_strNomeContribuinte_Click(Area As Integer)
    If Len(Trim(Me.dbc_strNomeContribuinte.Text)) = 0 Then
        DropDownDataCombo dbc_strNomeContribuinte, Me, Area
    Else
        Me.txtintContribuinte.Text = Me.dbc_strNomeContribuinte.BoundText
    End If
End Sub

Private Sub dbc_strOrgao_Click(Area As Integer)
    If Len(Trim(Me.dbc_strOrgao.Text)) = 0 Then
        DropDownDataCombo dbc_strOrgao, Me, Area
    Else
        Me.txtintOrgao.Text = Me.dbc_strOrgao.BoundText
    End If
End Sub

Private Sub dbc_strOrgao_KeyDown(KeyCode As Integer, Shift As Integer)
   DropDownDataCombo dbc_strOrgao, Me, , KeyCode, Shift
End Sub

Private Sub dbc_strOrgao_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub Form_Activate()
    gintCodSeguranca = 247
    VirificaGradeListView Me
    HabilitaDesabilitaBotao1 mblnselecionou, gstrMnuArquivo, gstrDeletar
    HabilitaDesabilitaBotao1 (Not mobjAux Is Nothing And mblnselecionou), gstrMnuArquivo, gstrAplicar
    HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrImprimir
End Sub

Private Sub cbointContaContabil_Click()
    cbo_ContaContabil.ListIndex = gintIndiceCBO(cbo_ContaContabil, _
                                    gstrItemData(cbointContaContabil))
End Sub

Private Sub cbo_ContaContabil_Click()
    Dim tempIndice As Integer
        
    tempIndice = cbo_ContaContabil.ListIndex
    cbointContaContabil.ListIndex = gintIndiceCBO(cbointContaContabil, _
                                    gstrItemData(cbo_ContaContabil))
                                    
   If cbointContaContabil.ListIndex = -1 Then
        LePlanoContaGeral cbointContaContabil, cbo_ContaContabil, "FN"
        cbo_ContaContabil.ListIndex = tempIndice
        cbointContaContabil.ListIndex = gintIndiceCBO(cbointContaContabil, _
                                    gstrItemData(cbo_ContaContabil))
   End If
   
   If cbo_ContaContabil.ListIndex = -1 Then cbointContaContabil.ListIndex = -1
End Sub

Private Sub Form_Deactivate()
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar
End Sub

Function strQuery() As String
    Dim strSQL As String
    strSQL = ""
    strSQL = strSQL & " SELECT PKId, strCodigo, strDescricao "
    strSQL = strSQL & " FROM " & gstrConvenio
    strSQL = strSQL & " ORDER BY strCodigo"
    strQuery = strSQL
End Function

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = vbKeyF1 Then
        Call_HtmlHelp Me.HelpContextID
    End If
End Sub

Private Sub Form_Load()
    mblnAlterando = False

    VerificaListaAutomatica "", tdb_Lista, strQuery
    VerificaObjParaAplicar mobjAux
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrImprimir
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar
    mblnselecionou = False
End Sub

Private Sub optbytTipo_KeyPress(Index As Integer, KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub tdb_Lista_Click()
    If glngQtdLinhaTDBGrid(tdb_Lista) = 1 Then
        tdb_Lista_RowColChange 0, 0
    End If
    
    
End Sub

Private Sub tdb_Lista_DblClick()
    MantemForm gstrAplicar
End Sub

Private Sub tdb_Lista_FilterChange()
    gblnFilraCampos tdb_Lista
End Sub

Private Sub tdb_Lista_KeyDown(KeyCode As Integer, Shift As Integer)
    mblnClickOk = True
End Sub

Private Sub tdb_Lista_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub tdb_Lista_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mblnClickOk = True
End Sub

Private Sub tdb_Lista_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    Dim tmp_con             As New clsBanco
    Dim tmp_rec             As New Recordset
    Dim intIndiceCombo      As Integer
        
    With tdb_Lista
        If (Not .EOF And Not .BOF) And mblnClickOk Then
            mblnClickOk = False
            LimpaObjeto Me
            LocalLimpaDBCeCBO
            txtPKId.Text = .Columns("PKID").Value
            
            If dbc_strOrgao.MatchedWithList = False Then
                LeDaTabelaParaObj gstrOrgao, dbc_strOrgao, "SELECT PKId, strDescricao FROM " & gstrOrgao & " WHERE intExercicio = " & gintExercicio & " ORDER BY strDescricao"
            End If
            
            If dbcintObjeto.MatchedWithList = False Then
                LeDaTabelaParaObj gstrObjetivo, dbcintObjeto
            End If
            
            If Me.dbc_strNomeContribuinte.MatchedWithList = False Then
                LeDaTabelaParaObj gstrContribuinte, dbc_strNomeContribuinte, "SELECT PKId, strNome FROM " & gstrContribuinte & " ORDER BY strNome"
            End If
            
            If Me.cbointContaContabil.ListCount = 0 Then
                LePlanoContaGeral cbointContaContabil, cbo_ContaContabil, "FN"
            End If
            
            If Me.cbo_CodOrcstrCodigo.ListCount = 0 Then
                LePrevisaoReceitaGeral cbo_CodOrcstrCodigo, cbo_CodOrcstrDescricao
            End If
            
            LeDaTabelaParaObj gstrConvenio, Me
            
            If Len(Trim(Me.txtintOrgao.Text)) > 0 Then
                If tmp_con.CriaADO("SELECT strDescricao FROM " & gstrOrgao & " WHERE PKid=" & Me.txtintOrgao.Text, 60, tmp_rec) = True Then
                    dbc_strOrgao.Text = tmp_rec.Fields("strDescricao")
                End If
            End If
            
            If Len(Trim(Me.txtintContribuinte.Text)) > 0 Then
                If tmp_con.CriaADO("SELECT strNome FROM " & gstrContribuinte & " WHERE PKid =" & Me.txtintContribuinte.Text, 60, tmp_rec) = True Then
                    dbc_strNomeContribuinte.Text = tmp_rec.Fields("strNome")
                End If
            End If
            
            If dbcintObjeto.BoundText <> "" Then
                If tmp_con.CriaADO("SELECT strDescricao FROM " & gstrObjetivo & " WHERE PKid=" & dbcintObjeto.BoundText, 60, tmp_rec) = True Then
                    dbcintObjeto.Text = tmp_rec!strDescricao
                End If
            End If
            
            If Me.txtintPrevisaoDaReceita.Text <> "" Then
                If tmp_con.CriaADO("SELECT intCodigoOrcamentario FROM " & gstrPrevisaoDaReceita & " WHERE PKid = " & Me.txtintPrevisaoDaReceita, 60, tmp_rec) = True Then
                    For intIndiceCombo = 0 To Me.cbo_CodOrcstrCodigo.ListCount - 1
                        If Me.cbo_CodOrcstrCodigo.ItemData(intIndiceCombo) = tmp_rec.Fields("intCodigoOrcamentario") Then
                            Me.cbo_CodOrcstrCodigo.ListIndex = intIndiceCombo
                            Me.cbo_CodOrcstrDescricao.ListIndex = intIndiceCombo
                            Exit For
                        End If
                    Next
                End If
            End If
            
            
            gCorLinhaSelecionada tdb_Lista
            
            'dbcintOrgao.BoundText = gite
            
            dbcintObjeto.BoundText = gstrItemData(dbcintObjeto, False)
            

            
            
            If mobjAux Is Nothing Then
                HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar
            Else
                HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrAplicar
            End If
            mblnselecionou = True
            mblnAlterando = True
        End If
    End With
    ValTxtStrCodigo = txtstrCodigo.Text
    ValTxtStrDescricao = txtstrDescricao.Text
End Sub

Private Function blnDadosOk() As Boolean
    'If CVDate(txtdtmDataAplicacaoFinal) < CVDate(txtdtmDataAplicacaoInicial) Then
    '    ExibeMensagem "Data Final não pode ser inferior à data inicial."
    '    txtdtmDataAplicacaoFinal.SetFocus
    '    Exit Function
    'ElseIf cbointContaContabil.ListIndex = -1 Then
    '    ExibeMensagem "A conta deve ser informada."
    '    cbointContaContabil.SetFocus
    '    Exit Function
    'Else
    
    blnDadosOk = False
    
    If txtdtmData.Text = "" Then
        ExibeMensagem "A data deve ser informada."
        txtdtmData.SetFocus
        Exit Function
    ElseIf txtdtmDataAplicacaoInicial.Text = "" Then
        ExibeMensagem "A data de aplicação inicial deve ser informada."
        txtdtmDataAplicacaoInicial.SetFocus
        Exit Function
    ElseIf txtdtmDataAplicacaoFinal.Text = "" Then
        ExibeMensagem "A data de aplicação final deve ser informada."
        txtdtmDataAplicacaoFinal.SetFocus
        Exit Function
    ElseIf cbointContaContabil.Text = "" Then
        ExibeMensagem "A conta contábil deve ser informada."
        cbointContaContabil.SetFocus
        Exit Function
    ElseIf dbcintObjeto.Text = "" Then
        ExibeMensagem "O objeto deve ser informada."
        dbcintObjeto.SetFocus
        Exit Function
    ElseIf txtstrCodigo.Text = "" Then
        ExibeMensagem "O código deve ser informada."
        txtstrCodigo.SetFocus
        Exit Function
    ElseIf txtstrDescricao.Text = "" Then
        ExibeMensagem "A descrição deve ser informada."
        txtstrDescricao.SetFocus
        Exit Function
    ElseIf mblnAlterando = False And gblnExisteCodigo(1, gstrConvenio, "strCodigo", txtstrCodigo.Text) Then
        ExibeMensagem "A código digitado já se encontra cadastrado."
        txtstrCodigo.SetFocus
        Exit Function
    ElseIf mblnAlterando = False And gblnExisteCodigo(1, gstrConvenio, "strDescricao", "'" & txtstrDescricao.Text & "'") Then
        ExibeMensagem "A descrição digitada já se encontra cadastrada."
        txtstrDescricao.SetFocus
        Exit Function
    
    'Verifica se foi alterado
    
    ElseIf txtstrCodigo <> ValTxtStrCodigo Then
        If mblnAlterando And gblnExisteCodigo(1, gstrConvenio, "strCodigo", txtstrCodigo.Text) Then
            ExibeMensagem "A código digitado já se encontra cadastrado."
            txtstrCodigo.SetFocus
            Exit Function
        End If
    
    
    ElseIf UCase(txtstrDescricao) <> UCase(ValTxtStrDescricao) Then
        If mblnAlterando And gblnExisteCodigo(1, gstrConvenio, "strDescricao", "'" & txtstrDescricao.Text & "'") Then
            ExibeMensagem "A descrição digitada já se encontra cadastrada."
            txtstrDescricao.SetFocus
            Exit Function
        End If
    
    End If
    
    blnDadosOk = True
    
End Function

Public Sub MantemForm(ByVal strModoOperacao As String)

'******************************************************************************************
' Data: 09/06/2003
' Alteração: - Substituição da chamada direta à stored procedure pela função
'            gstrStoredProcedure.
' Responsável: Everton Bianchini
'******************************************************************************************

    If strModoOperacao = gstrSalvar Then
        If blnDadosOk Then
            ToolBarGeral strModoOperacao, gstrConvenio, mblnAlterando, tdb_Lista, frmCadConvenio, mobjAux, strQuery, , rptControleConvenio, gstrStoredProcedure("sp_ControleConvenio", , True)
            LocalLimpaDBCeCBO
        End If
    ElseIf strModoOperacao = gstrPreencherLista Then
    
        If Me.ActiveControl.Name = cbointContaContabil.Name Or Me.ActiveControl.Name = cbo_ContaContabil.Name Then
            LePlanoContaGeral cbointContaContabil, cbo_ContaContabil, "FN"
            'LeDaTabelaParaObj "tblPlanoConta", cbo_ContaContabil, "SELECT PKId, strDescricao FROM tblPlanoConta WHERE ABS(blnFinanceira) = 1 AND ABS(blnAnalitica) = 1"
        End If
        
        If Me.ActiveControl.Name = cbo_CodOrcstrCodigo.Name Or Me.ActiveControl.Name = cbo_CodOrcstrDescricao.Name Then
            LePrevisaoReceitaGeral Me.cbo_CodOrcstrCodigo, Me.cbo_CodOrcstrDescricao
        End If
        
        If Me.ActiveControl.Name = dbc_strOrgao.Name Then
            LeDaTabelaParaObj gstrOrgao, dbc_strOrgao, "SELECT PKId, strDescricao FROM " & gstrOrgao & " WHERE intExercicio = " & gintExercicio & " ORDER BY strDescricao"
        End If
        
        If Me.ActiveControl.Name = dbc_strNomeContribuinte.Name Then
            LeDaTabelaParaObj gstrContribuinte, dbc_strNomeContribuinte, "SELECT PKId, strNome FROM " & gstrContribuinte & " ORDER BY strNome"
        End If
        
        If Me.ActiveControl.Name = dbcintObjeto.Name Then
            LeDaTabelaParaObj gstrObjetivo, dbcintObjeto
        End If
    
        
    ElseIf strModoOperacao = gstrAplicar Then
        LePlanoContaGeral cbointContaContabil, cbo_ContaContabil, "FN"
    Else
'            ToolBarGeral strModoOperacao, gstrConvenio, mblnAlterando, tdb_Lista, _
'                         Me, mobjAux, strQuery, , rptControleConvenio, "sp_ControleConvenio"
            ToolBarGeral strModoOperacao, gstrConvenio, mblnAlterando, tdb_Lista, _
                         Me, mobjAux, strQuery, , rptControleConvenio, gstrStoredProcedure("sp_ControleConvenio", , True)
            LocalLimpaDBCeCBO
    End If
End Sub

Private Sub txtintContribuinte_LostFocus()
    If Len(Me.txtintContribuinte.Text) > 0 Then
    
        RetornaCodigo gstrContribuinte, Me.dbc_strNomeContribuinte, "strNome", "PKID", Me.txtintContribuinte.Text
        
    Else
        Me.dbc_strNomeContribuinte.Text = ""
    End If

End Sub

Private Sub txtdtmData_GotFocus()
    MarcaCampo txtdtmData
End Sub

Private Sub txtdtmData_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "D", txtdtmData
End Sub

Private Sub txtdtmData_LostFocus()
    txtdtmData = gstrDataFormatada(txtdtmData)
End Sub

Private Sub txtdtmDataAplicacaoFinal_GotFocus()
    MarcaCampo txtdtmDataAplicacaoFinal
End Sub

Private Sub txtdtmDataAplicacaoFinal_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "D", txtdtmDataAplicacaoFinal
End Sub

Private Sub txtdtmDataAplicacaoFinal_LostFocus()
    txtdtmDataAplicacaoFinal = gstrDataFormatada(txtdtmDataAplicacaoFinal)
End Sub

Private Sub txtdtmDataAplicacaoInicial_GotFocus()
    MarcaCampo txtdtmDataAplicacaoInicial
End Sub

Private Sub txtdtmDataAplicacaoInicial_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "D", txtdtmDataAplicacaoInicial
End Sub

Private Sub txtdtmDataAplicacaoInicial_LostFocus()
    txtdtmDataAplicacaoInicial = gstrDataFormatada(txtdtmDataAplicacaoInicial)
End Sub

Private Sub txtintObjeto_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub txtintOrgao_LostFocus()
    If Len(Me.txtintOrgao.Text) > 0 Then
    
        RetornaCodigo gstrOrgao, Me.dbc_strOrgao, "strDescricao", "PKID", Me.txtintOrgao.Text, " AND intExercicio = " & gintExercicio
        
    Else
        Me.dbc_strOrgao.Text = ""
    End If
End Sub

Private Sub txtstrCodigo_GotFocus()
    gstrProximoCodigo txtstrCodigo, gstrConvenio, "strCodigo", gintCodSeguranca
    MarcaCampo txtstrCodigo
    
End Sub

Private Sub txtstrCodigo_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtstrCodigo
End Sub

Private Sub txtstrDescricao_GotFocus()
    MarcaCampo txtstrDescricao
End Sub

Private Sub txtstrDescricao_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub txtstrObservacao_GotFocus()
    MarcaCampo txtstrObservacao
End Sub

Private Sub txtstrObservacao_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrObservacao
End Sub

Private Sub txtdblValor_GotFocus()
    MarcaCampo txtdblValor
End Sub

Private Sub txtdblValor_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "V", txtdblValor
End Sub

Private Sub txtdblValor_LostFocus()
    txtdblValor = gstrConvVrDoSql(txtdblValor)
End Sub

Private Sub txtdblValorContrapartida_GotFocus()
    MarcaCampo txtdblValorContrapartida
End Sub

Private Sub txtdblValorContrapartida_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "V", txtdblValorContrapartida
End Sub

Private Sub txtdblValorContrapartida_LostFocus()
    txtdblValorContrapartida = gstrConvVrDoSql(txtdblValorContrapartida)
End Sub
Private Function RetornaCodigo(strTabela As String, objRetorno As Object, strCampoRetorno As String, _
    strCampoPesquisa As String, strValorDePesquisa As String, Optional strANDQuery As String, Optional blnSetFocus As Boolean) As String
'***************************************************************************
'Create By:             Éder Henrique                                      *
'Módulos:               Orçamentário                                       *
'Data:                  04/01/2006                                         *
'Ficha:                 orc1051                                            *
'Objetivo: Passando um parametro (strCampoPesquisa -> campo no banco) é    *
' consultado no banco e retornado (strCampoRetorno o campo passado como    *
' como parametro                                                           *
'***************************************************************************
    
    Dim strSQL          As String
    Dim adoTemp         As New ADODB.Recordset
    
    strSQL = ""
    strSQL = strSQL & "SELECT " & strCampoRetorno & " Retorno FROM " & strTabela
    strSQL = strSQL & " WHERE " & strCampoPesquisa & " = '" & strValorDePesquisa & "'"
    
    If Len(strANDQuery) > 0 Then strSQL = strSQL & strANDQuery
    
    Set gobjBanco = New clsBanco
    
    If gobjBanco.CriaADO(strSQL, 5, adoTemp) Then
        
        If Not adoTemp.EOF Then
            objRetorno.Text = gstrENulo(adoTemp.Fields("Retorno"))
        Else
            objRetorno.Text = ""
            If blnSetFocus = True Then objRetorno.SetFocus
        End If
        adoTemp.Close
    End If
End Function
Sub LocalLimpaDBCeCBO()
    Me.dbc_strNomeContribuinte.BoundText = ""
    Me.dbc_strOrgao.BoundText = ""
    Me.dbcintObjeto.BoundText = ""
    Me.cbo_CodOrcstrCodigo.ListIndex = -1
    Me.cbo_CodOrcstrDescricao.ListIndex = -1
End Sub

