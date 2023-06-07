VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "TODG7.OCX"
Begin VB.Form frmCadServicos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Serviços"
   ClientHeight    =   4410
   ClientLeft      =   2580
   ClientTop       =   2520
   ClientWidth     =   7245
   Icon            =   "frmCadServicos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4410
   ScaleWidth      =   7245
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab tab_3dPasta 
      Height          =   4335
      Left            =   60
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   60
      Width           =   7155
      _ExtentX        =   12621
      _ExtentY        =   7646
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Serviços"
      TabPicture(0)   =   "frmCadServicos.frx":1042
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblstrDescricao"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblCodigo"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblintlistaservicofederal"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "dbcintlistaservicofederal"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "tdb_Lista"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtPKId"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtstrdescricao"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtstrcodigo"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "Valores"
      TabPicture(1)   =   "frmCadServicos.frx":105E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmd_indexador"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "txt_DBLPORCENTAGEMISSVAR"
      Tab(1).Control(2)=   "txt_intExercicio"
      Tab(1).Control(3)=   "txt_DBLVALORISSFIXO"
      Tab(1).Control(4)=   "lvw_Itens"
      Tab(1).Control(5)=   "dbc_intindexadoreconomico"
      Tab(1).Control(6)=   "Label2"
      Tab(1).Control(7)=   "lblDBLVALORPROCENTAGEM"
      Tab(1).Control(8)=   "lbldblValor"
      Tab(1).Control(9)=   "Label1"
      Tab(1).ControlCount=   10
      Begin VB.CommandButton cmd_indexador 
         Height          =   315
         Left            =   -70740
         Picture         =   "frmCadServicos.frx":107A
         Style           =   1  'Graphical
         TabIndex        =   15
         TabStop         =   0   'False
         ToolTipText     =   "Ativa Cadastro Único"
         Top             =   1260
         Width           =   360
      End
      Begin VB.TextBox txtstrcodigo 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#.##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   2010
         MaxLength       =   5
         TabIndex        =   4
         Top             =   960
         Width           =   645
      End
      Begin VB.TextBox txt_DBLPORCENTAGEMISSVAR 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#.##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   -70170
         MaxLength       =   18
         TabIndex        =   12
         Top             =   930
         Width           =   1660
      End
      Begin VB.TextBox txt_intExercicio 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#.##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   -73380
         MaxLength       =   4
         TabIndex        =   8
         Top             =   540
         Width           =   495
      End
      Begin VB.TextBox txt_DBLVALORISSFIXO 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#.##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   -73380
         MaxLength       =   18
         TabIndex        =   10
         Top             =   900
         Width           =   1660
      End
      Begin VB.TextBox txtstrdescricao 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#.##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   1
         EndProperty
         Height          =   675
         Left            =   2010
         MaxLength       =   240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Top             =   1350
         Width           =   5025
      End
      Begin VB.TextBox txtPKId 
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         Left            =   3360
         Locked          =   -1  'True
         OLEDragMode     =   1  'Automatic
         OLEDropMode     =   1  'Manual
         TabIndex        =   18
         Top             =   0
         Visible         =   0   'False
         Width           =   855
      End
      Begin TrueOleDBGrid70.TDBGrid tdb_Lista 
         Height          =   2085
         Left            =   120
         TabIndex        =   17
         Top             =   2085
         Width           =   6915
         _ExtentX        =   12197
         _ExtentY        =   3678
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
         Columns(1).Caption=   "Lista de Serviço - Federal"
         Columns(1).DataField=   "strListaFederal"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Código"
         Columns(2).DataField=   "strcodigo"
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
         Splits(0)._ColumnProps(8)=   "Column(1).Width=2725"
         Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=2646"
         Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(13)=   "Column(2).Width=1217"
         Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=1138"
         Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(17)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(18)=   "Column(3).Width=7197"
         Splits(0)._ColumnProps(19)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(20)=   "Column(3)._WidthInPix=7117"
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
         _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=46,.parent=13"
         _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=43,.parent=14"
         _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=44,.parent=15"
         _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=45,.parent=17"
         _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=28,.parent=13"
         _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
         _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=50,.parent=13"
         _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=47,.parent=14"
         _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=48,.parent=15"
         _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=49,.parent=17"
         _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=32,.parent=13"
         _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=29,.parent=14"
         _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=30,.parent=15"
         _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=31,.parent=17"
         _StyleDefs(52)  =   "Named:id=33:Normal"
         _StyleDefs(53)  =   ":id=33,.parent=0"
         _StyleDefs(54)  =   "Named:id=34:Heading"
         _StyleDefs(55)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(56)  =   ":id=34,.wraptext=-1"
         _StyleDefs(57)  =   "Named:id=35:Footing"
         _StyleDefs(58)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(59)  =   "Named:id=36:Selected"
         _StyleDefs(60)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(61)  =   "Named:id=37:Caption"
         _StyleDefs(62)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(63)  =   "Named:id=38:HighlightRow"
         _StyleDefs(64)  =   ":id=38,.parent=33,.bgcolor=&H80000014&,.fgcolor=&H80000012&"
         _StyleDefs(65)  =   "Named:id=39:EvenRow"
         _StyleDefs(66)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(67)  =   "Named:id=40:OddRow"
         _StyleDefs(68)  =   ":id=40,.parent=33"
         _StyleDefs(69)  =   "Named:id=41:RecordSelector"
         _StyleDefs(70)  =   ":id=41,.parent=34"
         _StyleDefs(71)  =   "Named:id=42:FilterBar"
         _StyleDefs(72)  =   ":id=42,.parent=33"
      End
      Begin MSComctlLib.ListView lvw_Itens 
         Height          =   2235
         Left            =   -74490
         TabIndex        =   16
         Top             =   1890
         Width           =   6045
         _ExtentX        =   10663
         _ExtentY        =   3942
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Exercício"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Text            =   "Valor"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Porcentagem"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Indexador"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "intIndexador"
            Object.Width           =   0
         EndProperty
      End
      Begin MSDataListLib.DataCombo dbc_intindexadoreconomico 
         Height          =   315
         Left            =   -73380
         TabIndex        =   14
         Top             =   1260
         Width           =   2625
         _ExtentX        =   4630
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dbcintlistaservicofederal 
         Height          =   315
         Left            =   2010
         TabIndex        =   2
         Top             =   570
         Width           =   5025
         _ExtentX        =   8864
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Text            =   ""
      End
      Begin VB.Label lblintlistaservicofederal 
         AutoSize        =   -1  'True
         Caption         =   "Lista de Serviços - Federal"
         Height          =   195
         Left            =   75
         TabIndex        =   1
         Top             =   570
         Width           =   1875
      End
      Begin VB.Label lblCodigo 
         AutoSize        =   -1  'True
         Caption         =   "Código"
         Height          =   195
         Left            =   1455
         TabIndex        =   3
         Top             =   960
         Width           =   495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Indexador"
         Height          =   195
         Left            =   -74175
         TabIndex        =   13
         Top             =   1320
         Width           =   705
      End
      Begin VB.Label lblDBLVALORPROCENTAGEM 
         AutoSize        =   -1  'True
         Caption         =   "Valor Porcentagem"
         Height          =   195
         Left            =   -71550
         TabIndex        =   11
         Top             =   960
         Width           =   1350
      End
      Begin VB.Label lbldblValor 
         AutoSize        =   -1  'True
         Caption         =   "Valor"
         Height          =   195
         Left            =   -73830
         TabIndex        =   9
         Top             =   960
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Exercício"
         Height          =   195
         Left            =   -74130
         TabIndex        =   7
         Top             =   600
         Width           =   675
      End
      Begin VB.Label lblstrDescricao 
         AutoSize        =   -1  'True
         Caption         =   "Descrição"
         Height          =   195
         Left            =   1230
         TabIndex        =   5
         Top             =   1320
         Width           =   720
      End
   End
End
Attribute VB_Name = "frmCadServicos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim mblnAlterando       As Boolean
    Dim mobjAux             As Object
    Dim mblnSelecionou      As Boolean
    Dim mblnClickOk         As Boolean
    Dim bytOrdenacao        As Byte
    Dim blnOrdenacaoAsc     As Boolean
    Dim mobjLista           As Object
    Dim mblnAlterandoLista  As Boolean
    Dim intPkid             As Long
    Dim mblnAlterandoAux    As Boolean
    Dim strCodigo           As Integer
    Dim strDescricao        As String

Private Function strQuery() As String

Dim strSql  As String
   
    strSql = ""
    strSql = strSql & "Select "
    strSql = strSql & "LS.PKID, "
    strSql = strSql & "LSF.strdescricao as strListaFederal, "
    strSql = strSql & "LS.strcodigo, "
    strSql = strSql & "LS.STRDESCRICAO "
    strSql = strSql & "From " & gstrListaServico & " LS, " & gstrListaServicoFederal & " LSF "
    strSql = strSql & "Where LS.intListaServicoFederal " & strOUTJSQLServer & "= LSF.pkid" & strOUTJOracle & " "
    Select Case bytOrdenacao
        Case Is = 1
            strSql = strSql & " Order by LS.strcodigo" & IIf(blnOrdenacaoAsc, " ASC", " DESC")
        Case Is = 2
            strSql = strSql & " Order by LS.strDescricao" & IIf(blnOrdenacaoAsc, " ASC", " DESC")
    End Select
    
    strQuery = strSql
    
End Function

Private Function strQueryAplicar() As String
    Dim strSql  As String
    strSql = ""
    strSql = strSql & " SELECT PKId, strDescricao FROM "
    strSql = strSql & gstrListaServico & " ORDER BY strDescricao"
    strQueryAplicar = strSql
End Function

Private Sub cmd_indexador_Click()
    CarregaForm frmIndexadorEconomico, dbc_intindexadoreconomico
End Sub

Private Sub dbc_intindexadoreconomico_GotFocus()
    MarcaCampo dbc_intindexadoreconomico
End Sub

Private Sub dbc_intindexadoreconomico_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbc_intindexadoreconomico, Me, , , Shift
End Sub

Private Sub dbc_intindexadoreconomico_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbc_intindexadoreconomico
End Sub

Private Sub dbcintlistaservicofederal_GotFocus()
    MarcaCampo dbcintlistaservicofederal
End Sub

Private Sub dbcintlistaservicofederal_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbcintlistaservicofederal, Me, , KeyCode, Shift
End Sub

Private Sub dbcintlistaservicofederal_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbcintlistaservicofederal
End Sub

Private Sub Form_Activate()
    gintCodSeguranca = 1158
    If mblnSelecionou Then
        HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrDeletar
    Else
        HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrDeletar
    End If
    
    If tab_3dPasta.Tab = 1 Then
        HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrIncluirItem, gstrExcluirItem
    End If
    
    If mobjAux Is Nothing Then
        HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar
    Else
        HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrAplicar
    End If
End Sub

Private Sub Form_Deactivate()
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar
    HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrIncluirItem, gstrExcluirItem
End Sub

Private Sub Form_Load()

    bytOrdenacao = 1: blnOrdenacaoAsc = True
    dbc_intindexadoreconomico.Tag = strQueryIndexEconomico & ";strabreviatura"
    dbcintlistaservicofederal.Tag = strQueryListaServicoFederal & ";strDescricao"
    VerificaObjParaAplicar mobjAux
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar
    mblnSelecionou = False
End Sub


Private Sub lvw_Itens_Click()
If lvw_Itens.ListItems.Count > 0 Then
    txt_intExercicio.Text = lvw_Itens.SelectedItem.Text
    txt_DBLVALORISSFIXO = lvw_Itens.SelectedItem.SubItems(1)
    txt_DBLPORCENTAGEMISSVAR = lvw_Itens.SelectedItem.SubItems(2)
    PreencherListaDeOpcoes dbc_intindexadoreconomico, lvw_Itens.SelectedItem.SubItems(4)
    mblnAlterandoLista = True
End If
End Sub

Private Sub tab_3dPasta_Click(PreviousTab As Integer)
    If tab_3dPasta.Tab = 1 Then
        HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrIncluirItem, gstrExcluirItem
    Else
        HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrIncluirItem, gstrExcluirItem
    End If
End Sub

Private Sub tdb_Lista_Click()
    mblnClickOk = True
End Sub

Sub tdb_Lista_DblClick()
    MantemForm gstrAplicar
End Sub

Private Sub tdb_Lista_FilterChange()
    mblnClickOk = False
    gblnFilraCampos tdb_Lista
End Sub

Private Sub tdb_Lista_HeadClick(ByVal ColIndex As Integer)
    gOrdenaGrid tdb_Lista, ColIndex
End Sub

Private Sub tdb_Lista_KeyDown(KeyCode As Integer, Shift As Integer)
    mblnClickOk = True
End Sub

Private Sub tdb_Lista_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    With tdb_Lista
        If (Not .EOF And Not .BOF) And mblnClickOk Then
            mblnClickOk = False
            txtPKId.Text = .Columns("PKID").Value
            LeDaTabelaParaObj gstrListaServico, Me
            gCorLinhaSelecionada tdb_Lista
            If mobjAux Is Nothing Then
                HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar
            Else
                HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrAplicar
            End If
            strCodigo = Trim(tdb_Lista.Columns("strcodigo").Value)
            strDescricao = Trim(tdb_Lista.Columns("strDescricao").Value)
            PreencheListItens
            mblnSelecionou = True
            mblnAlterando = True
            mblnAlterandoLista = False
        End If
    End With
End Sub

Public Sub MantemForm(ByVal strModoOperacao As String)
    Dim strSql As String
    
    strSql = strQueryRelatorio
    
    If strModoOperacao = UCase("IMPRIMIR") Then
        ToolBarGeral strModoOperacao, gstrServico, mblnAlterando, tdb_Lista, Me, mobjAux, strSql, , rptservicos, strQueryRelatorio
        Exit Sub
    End If
    
    Select Case UCase(strModoOperacao)
        Case UCase(gstrNovo)
            If tab_3dPasta.Tab = 0 Then
                mblnClickOk = False
                mblnSelecionou = False
                mblnAlterando = False
                mblnAlterandoLista = False
                LimpaObjeto Me
                tab_3dPasta.Tab = 0
                gstrProximoCodigo txtstrCodigo, gstrListaServico, "strcodigo", gintCodSeguranca
                MarcaCampo txtstrCodigo
                strCodigo = 0
                strDescricao = ""
                txt_intExercicio = ""
                txt_DBLVALORISSFIXO = ""
                txt_DBLPORCENTAGEMISSVAR.Text = ""
                dbc_intindexadoreconomico.Text = ""
                lvw_Itens.ListItems.Clear
                dbcintlistaservicofederal.SetFocus
            Else
                txt_intExercicio = ""
                txt_DBLVALORISSFIXO = ""
                txt_DBLPORCENTAGEMISSVAR.Text = ""
                dbc_intindexadoreconomico.Text = ""
                mblnAlterandoLista = False
                txt_intExercicio.SetFocus
            End If
        Case UCase(gstrSalvar)
            If Not blnDadosOk Then Exit Sub
            If mblnAlterando Then
                mblnAlterandoAux = mblnAlterando
                intPkid = txtPKId
            Else
                mblnAlterandoAux = False
            End If
            Set gobjBanco = New clsBanco
            gobjBanco.ExecutaBeginTrans
            If ToolBarGeral(strModoOperacao, gstrListaServico, mblnAlterando, _
                            tdb_Lista, Me, mobjAux, strQuery, strQueryAplicar) Then gobjBanco.ExecutaCommitTrans
            strSql = StrSalvaItem
            If strSql <> "" Then
                If gobjBanco.Execute(strSql) Then
                    gobjBanco.ExecutaCommitTrans
                    tab_3dPasta.Tab = 0
                    MantemForm gstrNovo
                    LeDaTabelaParaObj "", tdb_Lista, strQuery
                Else
                    Set gobjBanco = New clsBanco
                    gobjBanco.ExecutaRollbackTrans
                End If
            Else
                tab_3dPasta.Tab = 0
                MantemForm gstrNovo
                LeDaTabelaParaObj "", tdb_Lista, strQuery
            End If
        
            Set gobjBanco = New clsBanco
            gobjBanco.ExecutaRollbackTrans
        
    Case UCase(gstrIncluirItem)
        IncluirItemNoGrid
        txt_intExercicio.SetFocus
    Case UCase(gstrExcluirItem)
        ExcluirItemNoGrid
        mblnAlterandoLista = False
    Case UCase(gstrDeletar)
        Set gobjBanco = New clsBanco
        gobjBanco.ExecutaBeginTrans
        strSql = " Delete from " & gstrListaServicoExercicio & " Where INTLISTASERVICO = " & txtPKId
        If gobjBanco.Execute(strSql) Then
            If ToolBarGeral(strModoOperacao, gstrListaServico, mblnAlterando, tdb_Lista, Me, mobjAux, strQuery, strQueryAplicar) Then
                tab_3dPasta.Tab = 0
                MantemForm gstrNovo
                LeDaTabelaParaObj "", tdb_Lista, strQuery
            Else
                Set gobjBanco = New clsBanco
                gobjBanco.ExecutaRollbackTrans
                End If
        Else
            Set gobjBanco = New clsBanco
            gobjBanco.ExecutaRollbackTrans
        End If
    Case Else
        ToolBarGeral strModoOperacao, gstrListaServico, mblnAlterando, tdb_Lista, Me, mobjAux, strQuery, strQueryAplicar
End Select

End Sub


Function strQueryRelatorio() As String
    
Dim strSql As String
   
   strSql = "select strcodigo,strdescricao from " & gstrListaServico
   
     
   Select Case bytOrdenacao
      
      Case Is = 1

      Case Is = 2

      
      Case Is = 3

         
   End Select
   
   strQueryRelatorio = strSql
   
End Function

Private Function blnDadosOk() As Boolean
    
    blnDadosOk = False
    txtstrCodigo.Text = Trim(txtstrCodigo.Text)
    If Trim(txtstrCodigo.Text) = "" Then
        ExibeMensagem "O Código é obrigatório."
        txtstrCodigo.SetFocus
        Exit Function
    ElseIf Trim(txtstrdescricao.Text) = "" Then
        ExibeMensagem "A descrição é obrigatória."
        txtstrdescricao.SetFocus
        Exit Function
    ElseIf Not mblnAlterando Then
        If gblnExisteCodigo(1, gstrListaServico, "strCodigo", "'" & txtstrCodigo.Text & "'") Then
            ExibeMensagem "A código informado já se encontra cadastrado."
            txtstrCodigo.SetFocus
            Exit Function
        End If
    End If

    blnDadosOk = True
    
End Function

Private Sub txt_DBLVALORISSFIXO_GotFocus()
    MarcaCampo txt_DBLVALORISSFIXO
End Sub

Private Sub txt_DBLVALORISSFIXO_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "V", txt_DBLVALORISSFIXO
End Sub

Private Sub txt_DBLVALORISSFIXO_LostFocus()
    txt_DBLVALORISSFIXO = gstrConvVrDoSql(txt_DBLVALORISSFIXO, 6)
End Sub


Private Sub txt_DBLPORCENTAGEMISSVAR_GotFocus()
    MarcaCampo txt_DBLPORCENTAGEMISSVAR
End Sub

Private Sub txt_DBLPORCENTAGEMISSVAR_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "V", txt_DBLPORCENTAGEMISSVAR
End Sub

Private Sub txt_DBLPORCENTAGEMISSVAR_LostFocus()
    txt_DBLPORCENTAGEMISSVAR = gstrConvVrDoSql(txt_DBLPORCENTAGEMISSVAR, 6)
End Sub

Private Sub txt_intExercicio_GotFocus()
    MarcaCampo txt_intExercicio
End Sub

Private Sub txt_intExercicio_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txt_intExercicio
End Sub
Private Function strQueryIndexEconomico() As String
    Dim strSql As String
    
    strSql = "SELECT Pkid,"
    strSql = strSql & " strabreviatura "
    strSql = strSql & " FROM "
    strSql = strSql & gstrIndexadorEconomico
    strSql = strSql & " ORDER BY strAbreviatura"
    
    strQueryIndexEconomico = strSql
End Function

Private Function IncluirItemNoGrid()
    Dim intInd          As Integer
    If blnDadosItens = False Then Exit Function
    With lvw_Itens
        If mblnAlterandoLista Then
            For intInd = 1 To .ListItems.Count
                If .SelectedItem.Index <> intInd Then
                    If Trim(txt_intExercicio) = .ListItems(intInd).Text Then
                        ExibeMensagem "Não é possível incluir itens com exercícios iguais."
                        Exit Function
                    End If
                End If
            Next
            .SelectedItem.Text = txt_intExercicio
            .SelectedItem.SubItems(1) = gstrConvVrDoSql(txt_DBLVALORISSFIXO, 6)
            .SelectedItem.SubItems(2) = gstrConvVrDoSql(txt_DBLPORCENTAGEMISSVAR, 6)
            .SelectedItem.SubItems(3) = dbc_intindexadoreconomico.Text
            .SelectedItem.SubItems(4) = dbc_intindexadoreconomico.BoundText
            mblnAlterandoLista = False
        Else
            For intInd = 1 To .ListItems.Count
                If Trim(txt_intExercicio) = .ListItems(intInd).Text Then
                    ExibeMensagem "Não é possível incluir itens com exercícios iguais."
                    Exit Function
                End If
            Next

            Set mobjLista = .ListItems.Add(, , txt_intExercicio)
            mobjLista.SubItems(1) = gstrConvVrDoSql(txt_DBLVALORISSFIXO, 6)
            mobjLista.SubItems(2) = gstrConvVrDoSql(txt_DBLPORCENTAGEMISSVAR, 6)
            mobjLista.SubItems(3) = dbc_intindexadoreconomico.Text
            mobjLista.SubItems(4) = dbc_intindexadoreconomico.BoundText
        End If
    End With
    txt_intExercicio.Text = ""
    txt_DBLVALORISSFIXO.Text = ""
    txt_DBLPORCENTAGEMISSVAR.Text = ""
    dbc_intindexadoreconomico.Text = ""
    
End Function

Private Function ExcluirItemNoGrid()
    With lvw_Itens
        If .ListItems.Count > 0 Then
            .ListItems.Remove .SelectedItem.Index
        End If
    End With
End Function

Private Function blnDadosItens() As Boolean
    blnDadosItens = False
    If Trim(Len(txt_intExercicio.Text)) <> 4 Then
        ExibeMensagem "O exercício deve ser preenchido corretamente."
        txt_intExercicio.SetFocus
        Exit Function
    ElseIf Trim(txt_DBLVALORISSFIXO) = "" Then
        ExibeMensagem "O valor deve ser preenchido corretamente."
        txt_DBLVALORISSFIXO.SetFocus
        Exit Function
    End If
    blnDadosItens = True
End Function
Private Function StrSalvaItem() As String
    Dim strSql  As String
    Dim intInd  As Integer
    
    strSql = ""
    If lvw_Itens.ListItems.Count > 0 Then
        strSql = IIf(bytDBType = Oracle, "Begin", "")
        If mblnAlterandoAux Then
            strSql = strSql & " Delete from " & gstrListaServicoExercicio & " Where INTLISTASERVICO = " & intPkid
            strSql = strSql & IIf(bytDBType = Oracle, ";", "")
        End If
    End If
    If lvw_Itens.ListItems.Count > 0 Then
        With lvw_Itens
            For intInd = 1 To .ListItems.Count
                strSql = strSql & " INSERT INTO "
                strSql = strSql & gstrListaServicoExercicio & " ("
                strSql = strSql & "INTLISTASERVICO, "
                strSql = strSql & "intexercicio, "
                strSql = strSql & "DBLVALORISSFIXO, "
                strSql = strSql & "DBLPORCENTAGEMISSVAR, "
                strSql = strSql & "intindexadoreconomico, "
                strSql = strSql & "dtmDtAtualizacao, "
                strSql = strSql & "lngCodUsr) "
                strSql = strSql & "Values("
                If mblnAlterandoAux Then
                    strSql = strSql & intPkid & ", "
                Else
                    strSql = strSql & glngPegaUltimaChave(gstrListaServico, "pkid") & ", "
                End If
                strSql = strSql & .ListItems(intInd).Text & ", "
                strSql = strSql & gstrConvVrParaSql(.ListItems(intInd).SubItems(1)) & ", "
                strSql = strSql & gstrConvVrParaSql(.ListItems(intInd).SubItems(2)) & ", "
                strSql = strSql & gstrENulo(.ListItems(intInd).SubItems(4), , True) & ", "
                strSql = strSql & gstrConvDtParaSql(gstrDataDoSistema) & ", "
                strSql = strSql & glngCodUsr & " "
                strSql = strSql & ")" & IIf(bytDBType = Oracle, ";", "")
            Next
        End With
    End If
    If lvw_Itens.ListItems.Count > 0 Then
        strSql = strSql & IIf(bytDBType = Oracle, "End;", "")
    End If
    StrSalvaItem = strSql
End Function

Private Sub PreencheListItens()
    Dim strSql As String
    Dim adoResultado As ADODB.Recordset
    
    strSql = strSql & "Select LE.*,IE.Strabreviatura,IE.Pkid As intIndex "
    strSql = strSql & "From "
    strSql = strSql & gstrListaServico & " LS, "
    
    strSql = strSql & gstrListaServicoExercicio & " LE, "
    strSql = strSql & gstrIndexadorEconomico & " IE "
    strSql = strSql & "Where "
    strSql = strSql & "LS.Pkid = LE.Intlistaservico AND "
    strSql = strSql & "IE.Pkid " & strOUTJOracle & " =" & strOUTJSQLServer & "  LE.Intindexadoreconomico AND "
    strSql = strSql & "LE.Intlistaservico = " & txtPKId
    strSql = strSql & " Order By LE.intExercicio"
    lvw_Itens.ListItems.Clear
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        With adoResultado
            If .EOF = False Then
                Do While Not .EOF
                    Set mobjLista = lvw_Itens.ListItems.Add(, , gstrENulo(!intExercicio))
                    mobjLista.SubItems(1) = gstrConvVrDoSql(gstrENulo(!DBLVALORISSFIXO), 6)
                    mobjLista.SubItems(2) = gstrConvVrDoSql(gstrENulo(!DBLPORCENTAGEMISSVAR), 6)
                    mobjLista.SubItems(3) = gstrENulo(!Strabreviatura)
                    mobjLista.SubItems(4) = gstrENulo(!intIndex)
                    .MoveNext
                Loop
            End If
        End With
    End If
End Sub

Private Sub txtstrCodigo_GotFocus()
    gstrProximoCodigo txtstrCodigo, gstrListaServico, "strcodigo", gintCodSeguranca
    MarcaCampo txtstrCodigo
End Sub

Private Sub txtstrCodigo_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrCodigo
End Sub

Private Function strQueryListaServicoFederal() As String
    Dim strSql As String
    
    strSql = "SELECT Pkid,"
    strSql = strSql & " strDescricao "
    strSql = strSql & " FROM "
    strSql = strSql & gstrListaServicoFederal
    strSql = strSql & " ORDER BY strDescricao"
    
    strQueryListaServicoFederal = strSql
End Function


