VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MsDatLst.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frmCadSubunidade 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Subunidades"
   ClientHeight    =   5655
   ClientLeft      =   2220
   ClientTop       =   2745
   ClientWidth     =   7140
   HelpContextID   =   20
   Icon            =   "CadSubunidade.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   7140
   Begin VB.TextBox txtPKId 
      Height          =   270
      Left            =   6360
      TabIndex        =   14
      Top             =   150
      Visible         =   0   'False
      Width           =   645
   End
   Begin TabDlg.SSTab tab_3dPasta 
      Height          =   5490
      Left            =   90
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   90
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   9684
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Subunidades"
      TabPicture(0)   =   "CadSubunidade.frx":1042
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblintOrgao"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblstrDescricao"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblintPeriodo"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblintUnidadeOrcamentaria"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblintGestao"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblintUnidadeGestora"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lblstrCodigo"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txt_CodigoOrcamentario"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtstrCodigo"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txt_CodigoOrgao"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "cmd_TipoDeAdm"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "cmd_Gestao"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "cmd_Periodicidade"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtstrDescricao"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "cmd_Orgao"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "tdb_Subunidade"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "dbcintUnidadeGestora"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "dbcintGestao"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "dbcintPeriodo"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "cmd_UnidadeOrcamentaria"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "dbcintOrgao"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "dbcintUnidadeOrcamentaria"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).ControlCount=   22
      Begin VB.ComboBox dbcintUnidadeOrcamentaria 
         Height          =   315
         IntegralHeight  =   0   'False
         Left            =   2640
         TabIndex        =   11
         Top             =   930
         Width           =   3825
      End
      Begin VB.ComboBox dbcintOrgao 
         Height          =   315
         IntegralHeight  =   0   'False
         Left            =   2640
         TabIndex        =   9
         Top             =   510
         Width           =   3825
      End
      Begin VB.CommandButton cmd_UnidadeOrcamentaria 
         Height          =   300
         Left            =   6510
         Picture         =   "CadSubunidade.frx":105E
         Style           =   1  'Graphical
         TabIndex        =   12
         TabStop         =   0   'False
         ToolTipText     =   "Clique aqui para cadastar a Unidade Orcamentária"
         Top             =   930
         Width           =   330
      End
      Begin MSDataListLib.DataCombo dbcintPeriodo 
         Height          =   315
         Left            =   1455
         TabIndex        =   6
         Top             =   2970
         Width           =   5025
         _ExtentX        =   8864
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dbcintGestao 
         Height          =   315
         Left            =   1455
         TabIndex        =   4
         Top             =   2550
         Width           =   5025
         _ExtentX        =   8864
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dbcintUnidadeGestora 
         Height          =   315
         Left            =   1455
         TabIndex        =   2
         Top             =   2130
         Width           =   5025
         _ExtentX        =   8864
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin TrueOleDBGrid70.TDBGrid tdb_Subunidade 
         Height          =   1875
         Left            =   150
         TabIndex        =   8
         Top             =   3450
         Width           =   6645
         _ExtentX        =   11721
         _ExtentY        =   3307
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
         Columns(1).DataField=   "strcodigo"
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
         Splits(0)._ColumnProps(13)=   "Column(2).Width=8493"
         Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=8414"
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
         AllowUpdate     =   0   'False
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
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=3,.bold=0,.fontsize=825,.italic=0"
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
         _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
         _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
         _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=46,.parent=13"
         _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
         _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
         _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
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
      Begin VB.CommandButton cmd_Orgao 
         Height          =   300
         Left            =   6510
         Picture         =   "CadSubunidade.frx":13E8
         Style           =   1  'Graphical
         TabIndex        =   10
         TabStop         =   0   'False
         ToolTipText     =   "Clique aqui para cadastar Órgão"
         Top             =   510
         Width           =   330
      End
      Begin VB.TextBox txtstrDescricao 
         Height          =   285
         Left            =   1455
         MaxLength       =   100
         TabIndex        =   1
         Top             =   1740
         Width           =   5390
      End
      Begin VB.CommandButton cmd_Periodicidade 
         Height          =   300
         Left            =   6510
         Picture         =   "CadSubunidade.frx":1772
         Style           =   1  'Graphical
         TabIndex        =   7
         TabStop         =   0   'False
         ToolTipText     =   "Clique aqui para cadastar Periodicidade"
         Top             =   2970
         Width           =   330
      End
      Begin VB.CommandButton cmd_Gestao 
         Height          =   300
         Left            =   6510
         Picture         =   "CadSubunidade.frx":1AFC
         Style           =   1  'Graphical
         TabIndex        =   5
         TabStop         =   0   'False
         ToolTipText     =   "Clique aqui para cadastar Gestão"
         Top             =   2550
         Width           =   330
      End
      Begin VB.CommandButton cmd_TipoDeAdm 
         Height          =   300
         Left            =   6510
         Picture         =   "CadSubunidade.frx":1E86
         Style           =   1  'Graphical
         TabIndex        =   3
         TabStop         =   0   'False
         ToolTipText     =   "Clique aqui para cadastar a Unidade Gestora"
         Top             =   2145
         Width           =   330
      End
      Begin VB.TextBox txt_CodigoOrgao 
         BackColor       =   &H80000000&
         CausesValidation=   0   'False
         Enabled         =   0   'False
         Height          =   315
         Left            =   1455
         MaxLength       =   10
         OLEDragMode     =   1  'Automatic
         OLEDropMode     =   1  'Manual
         TabIndex        =   16
         Top             =   510
         Width           =   1185
      End
      Begin VB.TextBox txtstrCodigo 
         Height          =   285
         Left            =   1455
         MaxLength       =   10
         TabIndex        =   0
         Top             =   1350
         Width           =   1185
      End
      Begin VB.TextBox txt_CodigoOrcamentario 
         BackColor       =   &H80000000&
         CausesValidation=   0   'False
         Enabled         =   0   'False
         Height          =   315
         Left            =   1455
         MaxLength       =   10
         OLEDragMode     =   1  'Automatic
         OLEDropMode     =   1  'Manual
         TabIndex        =   15
         Top             =   930
         Width           =   1185
      End
      Begin VB.Label lblstrCodigo 
         AutoSize        =   -1  'True
         Caption         =   "Código"
         Height          =   195
         Left            =   870
         TabIndex        =   23
         Top             =   1425
         Width           =   495
      End
      Begin VB.Label lblintUnidadeGestora 
         AutoSize        =   -1  'True
         Caption         =   "Unidade Gestora"
         Height          =   195
         Left            =   165
         TabIndex        =   22
         Top             =   2250
         Width           =   1200
      End
      Begin VB.Label lblintGestao 
         AutoSize        =   -1  'True
         Caption         =   "Gestão"
         Height          =   195
         Left            =   855
         TabIndex        =   21
         Top             =   2685
         Width           =   510
      End
      Begin VB.Label lblintUnidadeOrcamentaria 
         AutoSize        =   -1  'True
         Caption         =   "U.Orçamentária"
         Height          =   195
         Left            =   255
         TabIndex        =   20
         Top             =   1020
         Width           =   1110
      End
      Begin VB.Label lblintPeriodo 
         AutoSize        =   -1  'True
         Caption         =   "Periodicidade"
         Height          =   195
         Left            =   405
         TabIndex        =   19
         Top             =   3090
         Width           =   960
      End
      Begin VB.Label lblstrDescricao 
         AutoSize        =   -1  'True
         Caption         =   "Descrição"
         Height          =   195
         Left            =   645
         TabIndex        =   18
         Top             =   1830
         Width           =   720
      End
      Begin VB.Label lblintOrgao 
         AutoSize        =   -1  'True
         Caption         =   "Órgão"
         Height          =   195
         Left            =   930
         TabIndex        =   17
         Top             =   630
         Width           =   435
      End
   End
End
Attribute VB_Name = "frmCadSubunidade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim mblnAlterando          As Boolean
    Dim mobjGeral              As Object
    Dim mobjAux                As Object
    Dim mblnClickOk            As Boolean
    Dim intFiltroExercicio     As Integer
    Public mIntCodSeguranca    As Integer

Private Sub dbcintGestao_Click(Area As Integer)
   DropDownDataCombo dbcintGestao, Me, Area
End Sub

Private Sub dbcintGestao_KeyDown(KeyCode As Integer, Shift As Integer)
   DropDownDataCombo dbcintGestao, Me, , KeyCode, Shift
End Sub

Private Sub dbcintOrgao_Click()
   LeCodigoEspecifico txt_CodigoOrgao, gstrOrgao, dbcintOrgao
   LeDaTabelaParaObj "", dbcintUnidadeOrcamentaria, strQueryCodigoUO
End Sub

Private Sub dbcintPeriodo_Click(Area As Integer)
   DropDownDataCombo dbcintPeriodo, Me, Area
End Sub

Private Sub dbcintPeriodo_KeyDown(KeyCode As Integer, Shift As Integer)
   DropDownDataCombo dbcintPeriodo, Me, , KeyCode, Shift
End Sub

Private Sub dbcintUnidadeGestora_Click(Area As Integer)
   DropDownDataCombo dbcintUnidadeGestora, Me, Area
End Sub

Private Sub dbcintUnidadeGestora_KeyDown(KeyCode As Integer, Shift As Integer)
   DropDownDataCombo dbcintUnidadeGestora, Me, , KeyCode, Shift
End Sub

Private Sub dbcintGestao_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", dbcintGestao
End Sub

Private Sub dbcintPeriodo_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", dbcintPeriodo
End Sub

Private Sub dbcintUnidadeGestora_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", dbcintUnidadeGestora
End Sub

Private Function strQuerySub() As String
    
Dim strSQL          As String
    
    strSQL = "SELECT "
    strSQL = strSQL & "SU.PKId, "
    strSQL = strSQL & gstrCONVERT(cdt_numeric, "SU.strCodigo") & " strCodigo, "
    strSQL = strSQL & "SU.strDescricao "
    strSQL = strSQL & "FROM "
    strSQL = strSQL & "tblSubUnidade SU, "
    strSQL = strSQL & "tblOrgao OG "
    strSQL = strSQL & "WHERE "
    strSQL = strSQL & "OG.Pkid = SU.intOrgao "
    strSQL = strSQL & "AND OG.intExercicio = " & intFiltroExercicio
    
    If Trim(dbcintUnidadeOrcamentaria.Text) <> "" Then
        strSQL = strSQL & " AND intUnidadeOrcamentaria = " & gstrItemData(dbcintUnidadeOrcamentaria) & " "
    End If
    If Trim(dbcintOrgao.Text) <> "" Then
        strSQL = strSQL & " AND SU.intOrgao = " & gstrItemData(dbcintOrgao)
    End If
    
    strSQL = strSQL & " ORDER BY strCodigo "

    strQuerySub = strSQL
    
End Function

Private Function strQueryCodigoOrgao() As String

Dim strSQL  As String

    strSQL = "SELECT strCodigo FROM "
    strSQL = strSQL & gstrOrgao & " "
    strSQL = strSQL & "WHERE PKId = " & gstrItemData(dbcintOrgao)
    
    strQueryCodigoOrgao = strSQL
    
End Function

Private Function strQueryCodigoUO(Optional blnPegarCodigo As Boolean) As String
    
Dim strSQL  As String
    
    strSQL = ""
    
    If blnPegarCodigo Then
        strSQL = strSQL & "SELECT strCodigo FROM "
        strSQL = strSQL & gstrUnidadeOrcamentaria & " "
        strSQL = strSQL & "WHERE PKId = " & gstrItemData(dbcintUnidadeOrcamentaria)
    Else
        strSQL = strSQL & "SELECT PKId, strDescricao FROM "
        strSQL = strSQL & gstrUnidadeOrcamentaria & " "
        strSQL = strSQL & "WHERE intOrgao = " & gstrItemData(dbcintOrgao)
    End If
    
    strQueryCodigoUO = strSQL
    
End Function

Private Sub dbcintUnidadeOrcamentaria_Click()
    LeCodigoEspecifico txt_CodigoOrcamentario, gstrUnidadeOrcamentaria, dbcintUnidadeOrcamentaria
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
 On Error Resume Next
  If KeyCode = vbKeyF1 Then
     Call_HtmlHelp Me.HelpContextID
  End If
End Sub

Private Sub tdb_Subunidade_DblClick()
    MantemForm gstrAplicar
End Sub

Private Sub tdb_Subunidade_FilterChange()
    mblnClickOk = False
    gblnFilraCampos tdb_Subunidade
End Sub

Private Sub tdb_Subunidade_HeadClick(ByVal ColIndex As Integer)
    gOrdenaGrid tdb_Subunidade, ColIndex
End Sub

Private Sub tdb_Subunidade_KeyPress(KeyAscii As Integer)
    If tdb_Subunidade.Col = 1 Then
        CaracterValido KeyAscii, "A", tdb_Subunidade
    End If
End Sub

Private Sub tdb_Subunidade_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mblnClickOk = True
End Sub

Private Sub tdb_Subunidade_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    
    Dim rsTmp As ADODB.Recordset
    
    Set rsTmp = New ADODB.Recordset
    
    With tdb_Subunidade
        If (Not .EOF And Not .BOF) And mblnClickOk Then
            LimpaObjetos
            mblnClickOk = False
            txtPKId.Text = .Columns("PKID").Value
            LeDaTabelaParaObj gstrSubUnidade, Me
            
            PreencheCombos Trim(txtPKId)
            
            Set gobjBanco = New clsBanco
            
            If Trim(dbcintUnidadeGestora.BoundText) <> "" Then
                If gobjBanco.CriaADO("SELECT strDescricao FROM " & gstrUnidadeGestora & " WHERE PKid=" & dbcintUnidadeGestora.BoundText, 60, rsTmp) Then
                    If Not rsTmp.EOF Then
                        dbcintUnidadeGestora.Text = gstrENulo(rsTmp.Fields("strDescricao").Value)
                    End If
                End If
            End If
            
            If Trim(dbcintGestao.BoundText) <> "" Then
                If gobjBanco.CriaADO("SELECT strDescricao FROM " & gstrGestao & " WHERE PKid=" & dbcintGestao.BoundText, 60, rsTmp) Then
                    If Not rsTmp.EOF Then
                        dbcintGestao.Text = gstrENulo(rsTmp.Fields("strDescricao").Value)
                    End If
                End If
            End If
            
            If Trim(dbcintPeriodo.BoundText) <> "" Then
                If gobjBanco.CriaADO("SELECT strDescricao FROM " & gstrPeriodo & " WHERE PKid=" & dbcintPeriodo.BoundText, 60, rsTmp) Then
                    If Not rsTmp.EOF Then
                        dbcintPeriodo.Text = gstrENulo(rsTmp.Fields("strDescricao").Value)
                    End If
                End If
            End If
            
            gCorLinhaSelecionada tdb_Subunidade
            
            If mobjAux Is Nothing Then
                HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar
            Else
                HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrAplicar
            End If
            mblnAlterando = True
            
        End If
    End With
    
    If rsTmp.State = adStateOpen Then
        rsTmp.Close
    End If
    
    Set rsTmp = Nothing
End Sub

Public Sub MantemForm(ByVal strModoOperacao As String)
    ' Verifica o menu do qual o form foi chamado para atribuir o Exercício correto
    If gbytMenu = gbytMenuCadastro Then
        intFiltroExercicio = gintExercicio
    Else
        intFiltroExercicio = gintExercicio + 1
    End If
    If UCase(strModoOperacao) = UCase(gstrNovo) Then
        
        LimpaObjetos
        mblnAlterando = False
        HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrDeletar
        Exit Sub
        
    End If
    
    If UCase(strModoOperacao) = UCase(gstrPreencherLista) Then
        If Me.ActiveControl.Name = dbcintOrgao.Name Then
            LeDaTabelaParaObj "", dbcintOrgao, strQueryOrgao
            Exit Sub
        End If
    End If
    
    If UCase(strModoOperacao) = UCase(gstrSalvar) Then
        If blnDadosOk Then
            ToolBarGeral strModoOperacao, gstrSubUnidade, _
                 mblnAlterando, tdb_Subunidade, _
                 Me, mobjAux, strQuerySub, strQueryAplicar, _
                 rptSubunidade, strQueryRelatorio
       End If
    Else
         ToolBarGeral strModoOperacao, gstrSubUnidade, _
                 mblnAlterando, tdb_Subunidade, _
                 Me, mobjAux, strQuerySub, strQueryAplicar, _
                 rptSubunidade, strQueryRelatorio
    End If
End Sub

Private Sub cmd_Gestao_Click()
    CarregaForm frmCadGestao, dbcintGestao
End Sub

Private Sub cmd_Orgao_Click()
    CarregaForm frmCadOrgao, dbcintOrgao
End Sub

Private Sub cmd_Periodicidade_Click()
    CarregaForm frmCadPeriodo, dbcintPeriodo
End Sub

Private Sub cmd_TipoDeAdm_Click()
    CarregaForm frmCadUnidadeGestora, dbcintUnidadeGestora
End Sub

Private Sub cmd_UnidadeOrcamentaria_Click()
    CarregaForm frmCadUnidadeOrcamentaria, dbcintUnidadeOrcamentaria
End Sub

Private Sub Form_Activate()
    
    gintCodSeguranca = mIntCodSeguranca
    
    VirificaGradeListView Me
    
    If mblnAlterando Then
        HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrAplicar, gstrDeletar
    Else
        HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar
    End If
    If mobjAux Is Nothing Then
       HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar
    Else
       HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrAplicar
    End If
    
End Sub

Private Sub txtstrCodigo_GotFocus()
    If dbcintUnidadeOrcamentaria.ListIndex <> -1 And dbcintOrgao.ListIndex <> -1 Then
        gstrProximoCodigo txtstrCodigo, gstrSubUnidade, "strCodigo", gintCodSeguranca, "intUnidadeOrcamentaria", dbcintUnidadeOrcamentaria.ItemData(dbcintUnidadeOrcamentaria.ListIndex), , , "intOrgao", dbcintOrgao.ItemData(dbcintOrgao.ListIndex)
    End If
    
    MarcaCampo txtstrCodigo
End Sub

Private Sub txtstrCodigo_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtstrCodigo
End Sub

Private Sub txtstrCodigoOrcamentario_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub txtstrCodigoOrgao_GotFocus()
    MarcaCampo txt_CodigoOrgao
End Sub

Private Sub txtstrCodigoOrcamentario_GotFocus()
    MarcaCampo txt_CodigoOrcamentario
End Sub

Private Sub txtstrCodigoOrgao_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub txtstrDescricao_GotFocus()
    MarcaCampo txtstrDescricao
End Sub

Private Sub txtstrDescricao_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrDescricao
End Sub

Private Sub Form_Load()
    
    mblnAlterando = False
    
    ' Verifica o menu do qual o form foi chamado para atribuir o Exercício correto
    If gbytMenu = gbytMenuCadastro Then
        intFiltroExercicio = gintExercicio
    Else
        intFiltroExercicio = gintExercicio + 1
    End If
    
    LeDaTabelaParaObj gstrUnidadeGestora, dbcintUnidadeGestora
    LeDaTabelaParaObj gstrGestao, dbcintGestao
    LeDaTabelaParaObj "", dbcintOrgao, strQueryOrgao
    LeDaTabelaParaObj gstrPeriodo, dbcintPeriodo
    VerificaObjParaAplicar mobjAux
    
End Sub

Public Function strQueryRelatorio()
    
Dim strSQL  As String

    strSQL = "SELECT OG.strCodigo AS CodigoOrgao, OG.strDescricao AS Orgao, "
    strSQL = strSQL & "UO.strCodigo AS CodigoOrcamentario, "
    strSQL = strSQL & "UO.strDescricao AS UnidadeOrcamentaria, SU.strCodigo AS Codigo, "
    strSQL = strSQL & "SU.strDescricao AS Descricao, UG.strDescricao AS UnidadeGestora, "
    strSQL = strSQL & "GE.strDescricao AS Gestao, PE.strDescricao AS Periodicidade, "
    strSQL = strSQL & intFiltroExercicio & " " & " AS Exercicio "
    strSQL = strSQL & "FROM "
    strSQL = strSQL & gstrGestao & " GE, "
    strSQL = strSQL & gstrOrgao & " OG, "
    strSQL = strSQL & gstrSubUnidade & " SU, "
    strSQL = strSQL & gstrUnidadeOrcamentaria & " UO, "
    strSQL = strSQL & gstrUnidadeGestora & " UG, "
    strSQL = strSQL & gstrPeriodo & " PE "
    strSQL = strSQL & "WHERE SU.intOrgao = OG.PKId "
    strSQL = strSQL & " AND SU.intUnidadeOrcamentaria = UO.PKId "
    strSQL = strSQL & " AND SU.intUnidadeGestora " & strOUTJSQLServer & "= UG.PKId " & strOUTJOracle
    strSQL = strSQL & " AND SU.intGestao " & strOUTJSQLServer & "= GE.PKId " & strOUTJOracle
    strSQL = strSQL & " AND SU.intPeriodo " & strOUTJSQLServer & "= PE.PKId " & strOUTJOracle
    strSQL = strSQL & " AND OG.intExercicio = " & intFiltroExercicio
    strSQL = strSQL & " ORDER BY " & gstrCONVERT(CDT_INT, "OG.strCodigo") & ", " & gstrCONVERT(CDT_INT, "UO.strCodigo") & ", "
    strSQL = strSQL & "SU.strCodigo, SU.strDescricao"
    
    strQueryRelatorio = strSQL
    
End Function

Public Function strQueryOrgao()

Dim strSQL  As String

    strSQL = "SELECT PkID, strDescricao FROM "
    strSQL = strSQL & gstrOrgao & " "
    strSQL = strSQL & "WHERE intExercicio = " & intFiltroExercicio
    strSQL = strSQL & " ORDER BY strDescricao "
    
    strQueryOrgao = strSQL
    
End Function

Private Function strQueryAplicar() As String
    
    strQueryAplicar = " SELECT SU.PKId, SU.strDescricao FROM " & gstrSubUnidade & " SU, " & gstrUnidadeOrcamentaria & " UO, " & gstrOrgao & " O " & _
                    " WHERE SU.intUnidadeOrcamentaria = UO.PkId AND UO.intOrgao = O.PkId AND O.intExercicio = " & intFiltroExercicio
    If Val(Me.Tag) > 0 Then
        strQueryAplicar = strQueryAplicar & " AND UO.PKId = " & Me.Tag
    End If
    
End Function
Private Sub LimpaObjetos()
    
    LimpaObjeto Me
    dbcintUnidadeOrcamentaria.Clear
    dbcintOrgao.Clear
    
    txt_CodigoOrgao = Space$(0)
    txt_CodigoOrcamentario = Space$(0)
    
End Sub

Private Sub PreencheCombos(strPkidSubUnidade As String)

Dim strSQL As String

    strSQL = "SELECT " ' Carrega a Combo do Órgão
    strSQL = strSQL & " OG.Pkid Pkid, "
    strSQL = strSQL & " OG.strDescricao strDescricao "
    strSQL = strSQL & " FROM "
    strSQL = strSQL & gstrSubUnidade & " SU, "
    strSQL = strSQL & " tblOrgao OG "
    strSQL = strSQL & " WHERE "
    strSQL = strSQL & " SU.Pkid = " & strPkidSubUnidade
    strSQL = strSQL & " AND SU.intOrgao = OG.Pkid "

    LeDaTabelaParaObj "", dbcintOrgao, strSQL
    
    If dbcintOrgao.ListCount > 0 Then
        dbcintOrgao.ListIndex = 0
    End If
    
    strSQL = "SELECT " ' Carrega a Combo da Unidade Orçamentária
    strSQL = strSQL & " UO.Pkid Pkid, "
    strSQL = strSQL & " UO.strDescricao strDescricao "
    strSQL = strSQL & " FROM "
    strSQL = strSQL & " tblSubUnidade SU, "
    strSQL = strSQL & gstrUnidadeOrcamentaria & " UO "
    strSQL = strSQL & " WHERE "
    strSQL = strSQL & " SU.Pkid = " & strPkidSubUnidade
    strSQL = strSQL & " AND "
    strSQL = strSQL & " SU.intUnidadeOrcamentaria = UO.Pkid "
    
    LeDaTabelaParaObj "", dbcintUnidadeOrcamentaria, strSQL
    
    If dbcintUnidadeOrcamentaria.ListCount > 0 Then
        dbcintUnidadeOrcamentaria.ListIndex = 0
    End If
    
End Sub

Private Function blnDadosOk() As Boolean
    
    blnDadosOk = False
    
    If Trim(dbcintOrgao.Text) = "" Or dbcintOrgao.ListIndex < 0 Then
        ExibeMensagem "Selecione um Órgão corretamente."
        If dbcintOrgao.Enabled = True Then dbcintOrgao.SetFocus
        Exit Function
    End If
    
    If Trim(dbcintUnidadeOrcamentaria.Text) = "" Or dbcintUnidadeOrcamentaria.ListIndex < 0 Then
        ExibeMensagem "Selecione uma Unidade Orçamentária corretamente."
        If dbcintUnidadeOrcamentaria.Enabled = True Then dbcintUnidadeOrcamentaria.SetFocus
        Exit Function
    End If
    
    If Trim(txtstrCodigo) = "" Then
        ExibeMensagem "Insira um código válido."
        If txtstrCodigo.Enabled = True Then txtstrCodigo.SetFocus
        Exit Function
    End If
    
    'If mblnAlterando Then
        If blnExisteCodigoSubUnidade(Trim(txtstrCodigo), intFiltroExercicio, gstrItemData(dbcintOrgao), gstrItemData(dbcintUnidadeOrcamentaria)) Then
            ExibeMensagem "Este código já se encontra cadastrado."
            If txtstrCodigo.Enabled = True Then txtstrCodigo.SetFocus
            Exit Function
        End If
   'End If
    
    If Trim(txtstrDescricao) = "" Then
        ExibeMensagem "Insira uma descrição válida para a Sub Unidade."
        If txtstrDescricao.Enabled = True Then txtstrDescricao.SetFocus
        Exit Function
    End If
    
    
    blnDadosOk = True
    
End Function

Private Function blnExisteCodigoSubUnidade(strCodigo As String, intExercicio As Integer, intCodigoOrgao As Integer, intCodigoUnidade As Integer) As Boolean

Dim strSQL         As String
Dim adoResultado   As ADODB.Recordset

    blnExisteCodigoSubUnidade = False
    
    strSQL = "SELECT "
    strSQL = strSQL & "SU.Pkid "
    strSQL = strSQL & "FROM "
    strSQL = strSQL & gstrSubUnidade & " SU, "
    strSQL = strSQL & gstrOrgao & " OG, "
    strSQL = strSQL & gstrUnidadeOrcamentaria & " UO "
    strSQL = strSQL & "WHERE "
    strSQL = strSQL & "OG.Pkid = SU.intOrgao "
    strSQL = strSQL & "AND UO.Pkid = SU.intUnidadeOrcamentaria "
    strSQL = strSQL & "AND OG.IntExercicio = " & intExercicio & " "
    strSQL = strSQL & "AND " & gstrCONVERT(cdt_numeric, "SU.strCodigo") & " = " & strCodigo & " "
    strSQL = strSQL & "AND OG.Pkid = " & intCodigoOrgao & " "
    strSQL = strSQL & "AND UO.Pkid = " & intCodigoUnidade
    Set gobjBanco = New clsBanco
    Set adoResultado = New ADODB.Recordset
    
    If gobjBanco.CriaADO(strSQL, 60, adoResultado) Then
        If adoResultado.RecordCount > 0 Then
            blnExisteCodigoSubUnidade = True
        Else
            blnExisteCodigoSubUnidade = False
        End If
    Else
        blnExisteCodigoSubUnidade = False
    End If
End Function
