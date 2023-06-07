VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frmCadUnidadeOrcamentaria 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Unidades Orçamentárias"
   ClientHeight    =   4740
   ClientLeft      =   2595
   ClientTop       =   2040
   ClientWidth     =   6675
   HelpContextID   =   22
   Icon            =   "CadUnidadeOrcamentaria.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4740
   ScaleWidth      =   6675
   Begin VB.TextBox txtPKId 
      Height          =   270
      Left            =   5220
      TabIndex        =   9
      Top             =   120
      Visible         =   0   'False
      Width           =   645
   End
   Begin TabDlg.SSTab tab_3dPasta 
      Height          =   4500
      Left            =   120
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   120
      Width           =   6465
      _ExtentX        =   11404
      _ExtentY        =   7938
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Unidades Orçamentárias"
      TabPicture(0)   =   "CadUnidadeOrcamentaria.frx":1042
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblstrDescricao"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblstrOrdenador"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblintOrgao"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblintGestao"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblintUnidadeGestora"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblstrCodigo"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "tdb_UndOrcamentaria"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtstrDescricao"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cmd_Ordenador"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cmd_Orgao"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "cmd_Gestao"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "cmd_UnidadeGestora"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtstrCodigo"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txt_CodigoOrgao"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "dbcintOrgao"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "dbcintUnidadeGestora"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "dbcintGestao"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "dbcintOrdenador"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).ControlCount=   18
      Begin MSDataListLib.DataCombo dbcintOrdenador 
         Height          =   315
         Left            =   1440
         TabIndex        =   5
         Top             =   2280
         Width           =   4560
         _ExtentX        =   8043
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dbcintGestao 
         Height          =   315
         Left            =   1440
         TabIndex        =   4
         Top             =   1905
         Width           =   4560
         _ExtentX        =   8043
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dbcintUnidadeGestora 
         Height          =   315
         Left            =   1440
         TabIndex        =   3
         Top             =   1530
         Width           =   4560
         _ExtentX        =   8043
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dbcintOrgao 
         Height          =   315
         Left            =   2640
         TabIndex        =   0
         Top             =   420
         Width           =   3405
         _ExtentX        =   6006
         _ExtentY        =   556
         _Version        =   393216
         IntegralHeight  =   0   'False
         Text            =   ""
      End
      Begin VB.TextBox txt_CodigoOrgao 
         BackColor       =   &H80000000&
         CausesValidation=   0   'False
         Enabled         =   0   'False
         Height          =   300
         Left            =   1440
         MaxLength       =   10
         OLEDragMode     =   1  'Automatic
         OLEDropMode     =   1  'Manual
         TabIndex        =   19
         Top             =   435
         Width           =   1185
      End
      Begin VB.TextBox txtstrCodigo 
         Height          =   285
         Left            =   1440
         MaxLength       =   10
         TabIndex        =   1
         Top             =   810
         Width           =   1185
      End
      Begin VB.CommandButton cmd_UnidadeGestora 
         Height          =   300
         Left            =   6015
         Picture         =   "CadUnidadeOrcamentaria.frx":105E
         Style           =   1  'Graphical
         TabIndex        =   13
         TabStop         =   0   'False
         ToolTipText     =   "Clique aqui para cadastar unidade gestora"
         Top             =   1545
         Width           =   330
      End
      Begin VB.CommandButton cmd_Gestao 
         Height          =   300
         Left            =   6015
         Picture         =   "CadUnidadeOrcamentaria.frx":13E8
         Style           =   1  'Graphical
         TabIndex        =   12
         TabStop         =   0   'False
         ToolTipText     =   "Clique aqui para cadastar gestão"
         Top             =   1920
         Width           =   330
      End
      Begin VB.CommandButton cmd_Orgao 
         Height          =   300
         Left            =   6015
         Picture         =   "CadUnidadeOrcamentaria.frx":1772
         Style           =   1  'Graphical
         TabIndex        =   11
         TabStop         =   0   'False
         ToolTipText     =   "Clique aqui para cadastar órgão"
         Top             =   435
         Width           =   330
      End
      Begin VB.CommandButton cmd_Ordenador 
         Height          =   300
         Left            =   6015
         Picture         =   "CadUnidadeOrcamentaria.frx":1AFC
         Style           =   1  'Graphical
         TabIndex        =   10
         TabStop         =   0   'False
         ToolTipText     =   "Clique aqui para cadastar ordenador"
         Top             =   2295
         Width           =   330
      End
      Begin VB.TextBox txtstrDescricao 
         Height          =   285
         Left            =   1440
         MaxLength       =   100
         TabIndex        =   2
         Top             =   1170
         Width           =   4890
      End
      Begin TrueOleDBGrid70.TDBGrid tdb_UndOrcamentaria 
         Height          =   1725
         Left            =   120
         TabIndex        =   6
         Top             =   2670
         Width           =   6225
         _ExtentX        =   10980
         _ExtentY        =   3043
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
         Splits(0)._ColumnProps(8)=   "Column(1).Width=1376"
         Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=1296"
         Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(13)=   "Column(2).Width=9075"
         Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=8996"
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
      Begin VB.Label lblstrCodigo 
         AutoSize        =   -1  'True
         Caption         =   "Código Unidade"
         Height          =   195
         Left            =   210
         TabIndex        =   18
         Top             =   915
         Width           =   1140
      End
      Begin VB.Label lblintUnidadeGestora 
         AutoSize        =   -1  'True
         Caption         =   "Unidade Gestora"
         Height          =   195
         Left            =   150
         TabIndex        =   17
         Top             =   1650
         Width           =   1200
      End
      Begin VB.Label lblintGestao 
         AutoSize        =   -1  'True
         Caption         =   "Gestão"
         Height          =   195
         Left            =   840
         TabIndex        =   16
         Top             =   2025
         Width           =   510
      End
      Begin VB.Label lblintOrgao 
         AutoSize        =   -1  'True
         Caption         =   "Órgão"
         Height          =   195
         Left            =   915
         TabIndex        =   15
         Top             =   540
         Width           =   435
      End
      Begin VB.Label lblstrOrdenador 
         AutoSize        =   -1  'True
         Caption         =   "Ordenador"
         Height          =   195
         Left            =   600
         TabIndex        =   14
         Top             =   2400
         Width           =   750
      End
      Begin VB.Label lblstrDescricao 
         AutoSize        =   -1  'True
         Caption         =   "Descrição"
         Height          =   195
         Left            =   630
         TabIndex        =   8
         Top             =   1275
         Width           =   720
      End
   End
End
Attribute VB_Name = "frmCadUnidadeOrcamentaria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim mblnAlterando           As Boolean
    Dim mobjAux                 As Object
    Dim mblnClickOk             As Boolean
    Dim blnGridClik             As Boolean
    Public mIntCodSeguranca     As Integer
    Public intFiltroExercicio   As Integer
Private Sub dbcintGestao_Click(Area As Integer)
   DropDownDataCombo dbcintGestao, Me, Area
End Sub

Private Sub dbcintGestao_KeyDown(KeyCode As Integer, Shift As Integer)
   DropDownDataCombo dbcintGestao, Me, , KeyCode, Shift
End Sub

Private Sub dbcintGestao_KeyPress(KeyAscii As Integer)
     CaracterValido KeyAscii, "N", dbcintGestao
End Sub

Private Sub dbcintOrdenador_Click(Area As Integer)
   DropDownDataCombo dbcintOrdenador, Me, Area
End Sub

Private Sub dbcintOrdenador_KeyDown(KeyCode As Integer, Shift As Integer)
   DropDownDataCombo dbcintOrdenador, Me, , KeyCode, Shift
End Sub

Private Sub dbcintOrgao_Change()
    LeCodigoEspecifico txt_CodigoOrgao, gstrOrgao, dbcintOrgao
    If dbcintOrgao.BoundText = "" Then txt_CodigoOrgao = ""
End Sub

Private Function strQueryGrid() As String

Dim strSQL  As String

    strSQL = "SELECT UO.PkID, UO.STRDESCRICAO,"
    strSQL = strSQL & "UO.STRcodigo,UG.STRDESCRICAO intUnidadeGestora ,"
    strSQL = strSQL & "G.strDescricao intgestao,OD.strNome intOrdenador, "
    strSQL = strSQL & "O.STRDESCRICAO intorgao "
    strSQL = strSQL & "FROM " & gstrUnidadeOrcamentaria & " UO,"
    strSQL = strSQL & gstrUnidadeGestora & " UG," & gstrGestao & " G,"
    strSQL = strSQL & gstrOrgao & " O," & gstrOrdenador & " OD "
    strSQL = strSQL & "Where UO.PKId =" & txtPKId & " AND "
    strSQL = strSQL & "UG.PKID = UO.intUnidadeGestora AND G.Pkid = UO.INTGESTAO AND "
    strSQL = strSQL & "O.PKID = UO.INTORGAO AND OD.PKID = UO.Intordenador"
        
    strQueryGrid = strSQL
    
End Function

Private Function strQueryUO() As String

Dim strSQL  As String
    
    strSQL = "SELECT UO.PKId, "
    strSQL = strSQL & " UO.strCodigo strCodigo ,"
    ' Alterado pendência orc1553
    'strSQL = strSQL & gstrCONVERT(cdt_numeric, " UO.strCodigo") & "strCodigo ,"
    strSQL = strSQL & " UO.strDescricao FROM "
    strSQL = strSQL & gstrUnidadeOrcamentaria & " UO , "
    strSQL = strSQL & gstrOrgao & " O "
    strSQL = strSQL & "WHERE O.intExercicio=" & CStr(intFiltroExercicio) & " "
    strSQL = strSQL & "AND O.PKID=UO.intOrgao "
    strSQL = strSQL & IIf(Trim(dbcintOrgao.BoundText) <> "", " and UO.intOrgao = " & Trim(dbcintOrgao.BoundText), "")
    strQueryUO = strSQL
    
    strSQL = strSQL & " ORDER BY " & gstrCONVERT(cdt_numeric, "UO.strCodigo")
   
    strQueryUO = strSQL
    
End Function

Private Sub dbcintOrgao_Click(Area As Integer)
   DropDownDataCombo dbcintOrgao, Me, Area
End Sub

Private Sub dbcintOrgao_KeyDown(KeyCode As Integer, Shift As Integer)
   DropDownDataCombo dbcintOrgao, Me, , KeyCode, Shift
End Sub

Private Sub dbcintOrgao_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", dbcintOrgao
End Sub

Private Sub dbcintOrdenador_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", dbcintOrdenador
End Sub

Private Sub dbcintUnidadeGestora_Click(Area As Integer)
   DropDownDataCombo dbcintUnidadeGestora, Me, Area
End Sub

Private Sub dbcintUnidadeGestora_KeyDown(KeyCode As Integer, Shift As Integer)
   DropDownDataCombo dbcintUnidadeGestora, Me, , KeyCode, Shift
End Sub

Private Sub dbcintUnidadeGestora_KeyPress(KeyAscii As Integer)
   CaracterValido KeyAscii, "N", dbcintUnidadeGestora
End Sub

Private Sub cmd_Gestao_Click()
    CarregaForm frmCadGestao, dbcintGestao
End Sub

Private Sub cmd_Orgao_Click()
    CarregaForm frmCadOrgao, dbcintOrgao
End Sub

Private Sub cmd_Ordenador_Click()
    CarregaForm frmCadOrdenador, dbcintOrdenador
End Sub

Private Sub cmd_UnidadeGestora_Click()
    CarregaForm frmCadUnidadeGestora, dbcintUnidadeGestora
End Sub

Private Sub Form_Activate()
    
    gintCodSeguranca = mIntCodSeguranca
    
    VirificaGradeListView Me
    
    If mblnAlterando Then
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

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
 On Error Resume Next
  If KeyCode = vbKeyF1 Then
     Call_HtmlHelp Me.HelpContextID
  End If
End Sub

Private Sub tdb_UndOrcamentaria_Click()
    If glngQtdLinhaTDBGrid(tdb_UndOrcamentaria) = 1 Then
        tdb_UndOrcamentaria_RowColChange 0, 0
    End If
End Sub

Private Sub tdb_UndOrcamentaria_DblClick()
    MantemForm gstrAplicar
End Sub

Private Sub tdb_UndOrcamentaria_HeadClick(ByVal ColIndex As Integer)
    gOrdenaGrid tdb_UndOrcamentaria, ColIndex
End Sub

Private Sub tdb_UndOrcamentaria_KeyPress(KeyAscii As Integer)
    If tdb_UndOrcamentaria.Col = 1 Then
        CaracterValido KeyAscii, "N", tdb_UndOrcamentaria
    End If
End Sub

Private Sub tdb_UndOrcamentaria_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mblnClickOk = True
End Sub

Private Sub txtstrCodigo_GotFocus()
        
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' # Esta rotina naum funciona corretamente para esta tela, existe '
    ' a necessidade da criação de uma rotina específica para retornar '
    ' o próximo código.                                               '
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    If dbcintOrgao.Text <> "" Then
        gstrProximoCodigo txtstrCodigo, gstrUnidadeOrcamentaria, "strCodigo", gintCodSeguranca, "intOrgao", dbcintOrgao.BoundText
    End If
    txtstrCodigo.Text = Format(txtstrCodigo, "00")
    MarcaCampo txtstrCodigo
    
End Sub

Private Sub txtstrCodigo_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrCodigo
End Sub

Private Sub txt_CodigoOrgao_KeyPress(KeyAscii As Integer)
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
    
     'Vamos verificar qual menu que chamou o form, para definirmos o filtro
    VerificaDefineExercicio
    
    dbcintOrgao.Tag = strQueryOR & ";strDescricao"
    
    dbcintUnidadeGestora.Tag = "SELECT pkid, strDescricao FROM " & gstrUnidadeGestora & ";strdescricao"
    
    dbcintGestao.Tag = "SELECT pkid, strDescricao FROM " & gstrGestao & ";strdescricao"
    
    dbcintOrdenador.Tag = "SELECT pkid, strNome FROM " & gstrOrdenador & ";strNome"
    
    VerificaObjParaAplicar mobjAux
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar
    mblnAlterando = False
End Sub

Private Sub tdb_UndOrcamentaria_FilterChange()
    mblnClickOk = False
    gblnFilraCampos tdb_UndOrcamentaria
End Sub

Private Sub tdb_UndOrcamentaria_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    With tdb_UndOrcamentaria
        If Not .EOF And Not .BOF And mblnClickOk Then
            mblnClickOk = False
            blnGridClik = True
            txtPKId.Text = .Columns("PKID").Value
            LeDaTabelaParaObj gstrUnidadeOrcamentaria, Me
            PreencherListaDeOpcoes dbcintGestao, dbcintGestao.BoundText
            PreencherListaDeOpcoes dbcintOrdenador, dbcintOrdenador.BoundText
            PreencherListaDeOpcoes dbcintUnidadeGestora, dbcintUnidadeGestora.BoundText
            
            If mobjAux Is Nothing Then
                HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar
            Else

                HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrAplicar
            End If
            mblnAlterando = True
            blnGridClik = False
        End If
    End With
End Sub

Public Sub MantemForm(ByVal strModoOperacao As String)

    Select Case UCase(strModoOperacao)
        Case gstrNovo
            
        Case gstrSalvar
            
        Case gstrDeletar
            
        Case UCase(gstrCancelar)
            
        Case UCase(gstrPreencherLista)
    
            If Me.ActiveControl.Name = dbcintUnidadeGestora.Name Then
               LeDaTabelaParaObj gstrUnidadeGestora, dbcintUnidadeGestora
            End If
            
            If Me.ActiveControl.Name = dbcintGestao.Name Then
               LeDaTabelaParaObj gstrGestao, dbcintGestao
            End If
            
            If Me.ActiveControl.Name = dbcintOrgao.Name Then
               LeDaTabelaParaObj gstrOrgao, dbcintOrgao, strQueryOR
            End If
            
            If Me.ActiveControl.Name = dbcintOrdenador.Name Then
               LeDaTabelaParaObj gstrOrdenador, dbcintOrdenador
            End If
    
         Case gstrFechar
            Unload Me

    End Select

    If strModoOperacao = gstrSalvar Then
        If blnDadosOk Then
            ToolBarGeral strModoOperacao, gstrUnidadeOrcamentaria, mblnAlterando, _
            tdb_UndOrcamentaria, Me, mobjAux, strQueryUO, , _
            rptUnidadeOrcamentaria, strQueryRelatorio
        End If
    ElseIf strModoOperacao <> gstrSalvar And strModoOperacao <> gstrPreencherLista Then
        ToolBarGeral strModoOperacao, gstrUnidadeOrcamentaria, mblnAlterando, _
        tdb_UndOrcamentaria, Me, mobjAux, strQueryUO, strQueryAplicar, _
        rptUnidadeOrcamentaria, strQueryRelatorio
    End If

End Sub

Public Function strQueryRelatorio()
    
Dim strSQL  As String
    
    strSQL = "SELECT OG.strCodigo AS CodigoOrgao, "
    strSQL = strSQL & "OG.strDescricao AS Orgao, UO.strCodigo AS CodigoUnidade, "
    strSQL = strSQL & "UO.strDescricao AS UnidadeOrcamentaria, "
    strSQL = strSQL & "UG.strDescricao AS UnidadeGestora, GE.strDescricao AS Gestao, "
    strSQL = strSQL & "strNome AS Ordenador "
    strSQL = strSQL & "FROM "
    strSQL = strSQL & gstrUnidadeOrcamentaria & " UO, "
    strSQL = strSQL & gstrUnidadeOrcamentaria & " UG, "
    strSQL = strSQL & gstrGestao & " GE, "
    strSQL = strSQL & gstrOrgao & " OG, "
    strSQL = strSQL & gstrOrdenador & " OD "
    strSQL = strSQL & "WHERE UO.intOrgao = OG.PKId "
    strSQL = strSQL & "AND OG.intExercicio=" & intFiltroExercicio & " "
    strSQL = strSQL & "AND UO.intUnidadeGestora " & strOUTJSQLServer & "= UG.PKId " & strOUTJOracle
    strSQL = strSQL & "AND UO.intGestao " & strOUTJSQLServer & "= GE.PKId " & strOUTJOracle
    strSQL = strSQL & "AND UO.intOrdenador " & strOUTJSQLServer & "= OD.PKId " & strOUTJOracle
    strSQL = strSQL & "ORDER BY " & gstrCONVERT(CDT_INT, "OG.strCodigo") & ", "
    strSQL = strSQL & "UO.strCodigo, "
    strSQL = strSQL & "UO.strDescricao, UG.strDescricao, "
    strSQL = strSQL & "GE.strDescricao, OD.strNome"
    
    strQueryRelatorio = strSQL
    
End Function
Private Function blnDadosOk() As Boolean
    Dim strWhereComplementar    As String
    
    'Incluido orc1551 para impedir inclusão de descricoes repetidas no mesmo exercicio
    If mblnAlterando Then
        strWhereComplementar = " AND PKID <> " & Me.txtPKId.Text
    Else
        strWhereComplementar = ""
    End If
    
    blnDadosOk = False
    
    If dbcintOrgao.Text = "" Then
        ExibeMensagem "O Órgão deve ser informado."
        dbcintOrgao.SetFocus
        Exit Function
    ElseIf txtstrCodigo.Text = "" Then
        ExibeMensagem "O código deve ser informado."
        txtstrCodigo.SetFocus
        Exit Function
    ElseIf txtstrDescricao.Text = "" Then
        ExibeMensagem "A descrição deve ser informada."
        txtstrDescricao.SetFocus
        Exit Function

    End If
    
'    If mblnAlterando Then
'        If gblnExisteCodigo(2, gstrUnidadeOrcamentaria, "strDescricao", "'" & txtstrDescricao.Text & "'", "intorgao", dbcintOrgao.BoundText, , , strWhereComplementar) Then
'            ExibeMensagem "A descrição informada já se encontra cadastrada."
'            txtstrDescricao.SetFocus
'            Exit Function
'        End If
'
'    Else
        If gblnExisteCodigo(2, gstrUnidadeOrcamentaria, "strCodigo", txtstrCodigo.Text, "intorgao", dbcintOrgao.BoundText, , , strWhereComplementar) Then
            ExibeMensagem "O código informado já se encontra cadastrado."
            txtstrCodigo.SetFocus
            Exit Function
        End If
        If gblnExisteCodigo(2, gstrUnidadeOrcamentaria, "strDescricao", "'" & UCase(Trim(txtstrDescricao.Text)) & "'", "intorgao", dbcintOrgao.BoundText, , , strWhereComplementar) Then
            ExibeMensagem "A descrição informada já se encontra cadastrada."
            txtstrDescricao.SetFocus
            Exit Function
        End If

'    End If

    blnDadosOk = True
    
End Function
Private Function strQueryOR() As String
    
Dim strSQL  As String

    strSQL = strSQL & "SELECT pkid, strDescricao FROM "
    strSQL = strSQL & gstrOrgao
    strSQL = strSQL & " WHERE intExercicio = " & intFiltroExercicio
    strSQL = strSQL & " ORDER BY strDescricao"
    
    strQueryOR = strSQL
    
End Function

Private Function strQueryAplicar() As String

    strQueryAplicar = " SELECT UO.PKId, UO.strDescricao FROM " & gstrUnidadeOrcamentaria & " UO, " & gstrOrgao & " O"
    strQueryAplicar = strQueryAplicar & " WHERE O.PKId=UO.intOrgao AND O.intExercicio = " & intFiltroExercicio
    
    If Val(Me.Tag) > 0 Then
        strQueryAplicar = strQueryAplicar & " AND O.PKId = " & Me.Tag
    End If
    
End Function
Public Sub VerificaDefineExercicio()
    
    If gbytMenu = gbytMenuCadastro Then
        intFiltroExercicio = gintExercicio
    Else
        intFiltroExercicio = gintExercicio + 1
    End If
    
End Sub
