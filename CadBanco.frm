VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frmCadBanco 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bancos"
   ClientHeight    =   5400
   ClientLeft      =   2520
   ClientTop       =   2250
   ClientWidth     =   7545
   HelpContextID   =   17
   Icon            =   "CadBanco.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5400
   ScaleWidth      =   7545
   Begin TabDlg.SSTab tab_3dPasta 
      Height          =   5265
      Left            =   90
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   90
      Width           =   7425
      _ExtentX        =   13097
      _ExtentY        =   9287
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Bancos"
      TabPicture(0)   =   "CadBanco.frx":1042
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblstrSigla"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblstrDescricao"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblintBanco"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lbl_intDigitoBanco"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblstrConvenioDebito"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblintultimolote"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "tdb_Lista"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtPKId"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "fra_LogoBanco"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtintDigitoBanco"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtintSequencialDebAut"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtstrConvenioDebAut"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtintBanco"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtstrDescricao"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txtstrSigla"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).ControlCount=   15
      Begin VB.TextBox txtstrSigla 
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
         Left            =   2190
         MaxLength       =   15
         TabIndex        =   2
         Top             =   1140
         Width           =   1515
      End
      Begin VB.TextBox txtstrDescricao 
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
         Left            =   2190
         MaxLength       =   40
         TabIndex        =   3
         Top             =   1530
         Width           =   3840
      End
      Begin VB.TextBox txtintBanco 
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
         Left            =   2190
         MaxLength       =   10
         TabIndex        =   0
         Top             =   720
         Width           =   1020
      End
      Begin VB.TextBox txtstrConvenioDebAut 
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
         Left            =   2190
         MaxLength       =   20
         TabIndex        =   4
         Top             =   1920
         Width           =   1860
      End
      Begin VB.TextBox txtintSequencialDebAut 
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
         Left            =   5010
         MaxLength       =   6
         TabIndex        =   5
         Top             =   1920
         Width           =   1020
      End
      Begin VB.TextBox txtintDigitoBanco 
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
         Left            =   4035
         MaxLength       =   1
         TabIndex        =   1
         Top             =   720
         Width           =   390
      End
      Begin VB.Frame fra_LogoBanco 
         Caption         =   " Logotipo "
         Height          =   1365
         Left            =   6150
         TabIndex        =   12
         Top             =   450
         Width           =   1125
         Begin VB.TextBox txtintLogoBanco 
            Height          =   285
            Left            =   30
            TabIndex        =   13
            Top             =   600
            Visible         =   0   'False
            Width           =   1065
         End
         Begin VB.Image img_LogoBanco 
            BorderStyle     =   1  'Fixed Single
            Height          =   1110
            Left            =   0
            MouseIcon       =   "CadBanco.frx":105E
            MousePointer    =   99  'Custom
            Stretch         =   -1  'True
            Top             =   240
            Width           =   1110
         End
      End
      Begin VB.TextBox txtPKId 
         Height          =   270
         Left            =   1500
         TabIndex        =   14
         Top             =   0
         Visible         =   0   'False
         Width           =   645
      End
      Begin TrueOleDBGrid70.TDBGrid tdb_Lista 
         Height          =   2805
         Left            =   900
         TabIndex        =   6
         Top             =   2340
         Width           =   6045
         _ExtentX        =   10663
         _ExtentY        =   4948
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
         Columns(1).Caption=   "Número"
         Columns(1).DataField=   "intBanco"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Dígito"
         Columns(2).DataField=   "intDigitoBanco"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Descrição"
         Columns(3).DataField=   "strDescricao"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "Sigla"
         Columns(4).DataField=   "strSigla"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "Logotipo"
         Columns(5).DataField=   "intLogoBanco"
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
         Splits(0)._ColumnProps(8)=   "Column(1).Width=2143"
         Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=2064"
         Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=2"
         Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(14)=   "Column(2).Width=1085"
         Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=1005"
         Splits(0)._ColumnProps(17)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(18)=   "Column(2)._ColStyle=2"
         Splits(0)._ColumnProps(19)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(20)=   "Column(3).Width=5080"
         Splits(0)._ColumnProps(21)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(22)=   "Column(3)._WidthInPix=5001"
         Splits(0)._ColumnProps(23)=   "Column(3)._EditAlways=0"
         Splits(0)._ColumnProps(24)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(25)=   "Column(4).Width=1773"
         Splits(0)._ColumnProps(26)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(27)=   "Column(4)._WidthInPix=1693"
         Splits(0)._ColumnProps(28)=   "Column(4)._EditAlways=0"
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
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=188,.bold=0,.fontsize=825,.italic=0"
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
         _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=28,.parent=13,.alignment=1"
         _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
         _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=54,.parent=13,.alignment=1"
         _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=51,.parent=14"
         _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=52,.parent=15"
         _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=53,.parent=17"
         _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=32,.parent=13"
         _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=29,.parent=14"
         _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=30,.parent=15"
         _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=31,.parent=17"
         _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=50,.parent=13"
         _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=47,.parent=14"
         _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=48,.parent=15"
         _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=49,.parent=17"
         _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=58,.parent=13"
         _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=55,.parent=14"
         _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=56,.parent=15"
         _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=57,.parent=17"
         _StyleDefs(60)  =   "Named:id=33:Normal"
         _StyleDefs(61)  =   ":id=33,.parent=0"
         _StyleDefs(62)  =   "Named:id=34:Heading"
         _StyleDefs(63)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(64)  =   ":id=34,.wraptext=-1"
         _StyleDefs(65)  =   "Named:id=35:Footing"
         _StyleDefs(66)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(67)  =   "Named:id=36:Selected"
         _StyleDefs(68)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(69)  =   "Named:id=37:Caption"
         _StyleDefs(70)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(71)  =   "Named:id=38:HighlightRow"
         _StyleDefs(72)  =   ":id=38,.parent=33,.bgcolor=&H80000014&,.fgcolor=&H80000012&"
         _StyleDefs(73)  =   "Named:id=39:EvenRow"
         _StyleDefs(74)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(75)  =   "Named:id=40:OddRow"
         _StyleDefs(76)  =   ":id=40,.parent=33"
         _StyleDefs(77)  =   "Named:id=41:RecordSelector"
         _StyleDefs(78)  =   ":id=41,.parent=34"
         _StyleDefs(79)  =   "Named:id=42:FilterBar"
         _StyleDefs(80)  =   ":id=42,.parent=33"
      End
      Begin VB.Label lblintultimolote 
         AutoSize        =   -1  'True
         Caption         =   "Último lote"
         Height          =   195
         Left            =   4200
         TabIndex        =   16
         Top             =   1980
         Width           =   735
      End
      Begin VB.Label lblstrConvenioDebito 
         AutoSize        =   -1  'True
         Caption         =   "Convenio débito automático"
         Height          =   195
         Left            =   105
         TabIndex        =   15
         Top             =   1980
         Width           =   1980
      End
      Begin VB.Label lbl_intDigitoBanco 
         AutoSize        =   -1  'True
         Caption         =   "Dígito"
         Height          =   195
         Left            =   3510
         TabIndex        =   9
         Top             =   765
         Width           =   435
      End
      Begin VB.Label lblintBanco 
         AutoSize        =   -1  'True
         Caption         =   "Número"
         Height          =   195
         Left            =   1530
         TabIndex        =   8
         Top             =   765
         Width           =   555
      End
      Begin VB.Label lblstrDescricao 
         AutoSize        =   -1  'True
         Caption         =   "Nome"
         Height          =   195
         Left            =   1665
         TabIndex        =   11
         Top             =   1575
         Width           =   420
      End
      Begin VB.Label lblstrSigla 
         AutoSize        =   -1  'True
         Caption         =   "Sigla"
         Height          =   195
         Left            =   1740
         TabIndex        =   10
         Top             =   1185
         Width           =   345
      End
   End
End
Attribute VB_Name = "frmCadBanco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim mblnAlterando      As Boolean
    Dim mobjAux            As Object
    Dim mblnClickOk        As Boolean
    Dim mblnPrimeiraVez    As Boolean
    Dim bytOrdenacao       As Byte
    Dim blnOrdenacaoAsc    As Boolean
    Dim strCodigoAtual     As String
    Dim strDescricaoAtual  As String

Private Function strQuery() As String
    Dim strsql  As String
    strsql = ""
    strsql = strsql & " SELECT PKId, intBanco, strDescricao, strSigla, intDigitoBanco, intLogoBanco FROM "
    strsql = strsql & gstrBanco ' & " ORDER BY intBanco"
    
    Select Case bytOrdenacao
        Case Is = 1
            strsql = strsql & " ORDER BY intBanco" & IIf(blnOrdenacaoAsc, " ASC", " DESC")
        Case Is = 2
            strsql = strsql & " ORDER BY strDescricao" & IIf(blnOrdenacaoAsc, " ASC", " DESC")
        Case Is = 3
            strsql = strsql & " ORDER BY strSigla" & IIf(blnOrdenacaoAsc, " ASC", " DESC")
    End Select
    
    strQuery = strsql
End Function

Private Sub Form_Activate()
    gintCodSeguranca = 590
    VirificaGradeListView Me
    HabilitaDesabilitaBotao1 mblnAlterando, gstrMnuArquivo, gstrDeletar
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
    'VerificaListaAutomatica gstrBanco, tdb_Lista, strQuery
    VerificaObjParaAplicar mobjAux
    
    bytOrdenacao = 2: blnOrdenacaoAsc = True
End Sub
 
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar
End Sub

Private Sub img_LogoBanco_DblClick()
    MantemForm gstrLogotipo
End Sub

Private Sub img_LogoBanco_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        Set img_LogoBanco.Picture = Nothing
        txtintLogoBanco = ""
    End If
End Sub

Private Sub tdb_Lista_Click()
    mblnPrimeiraVez = True
    If glngQtdLinhaTDBGrid(tdb_Lista) = 1 Then
        tdb_Lista_RowColChange 0, 0
    End If
End Sub

Private Sub tdb_Lista_DblClick()
    MantemForm gstrAplicar
End Sub

Private Sub tdb_Lista_FilterChange()
    mblnPrimeiraVez = False
    gblnFilraCampos tdb_Lista
End Sub

Private Sub tdb_Lista_HeadClick(ByVal ColIndex As Integer)
    gOrdenaGrid tdb_Lista, ColIndex
End Sub

Private Sub tdb_Lista_KeyDown(KeyCode As Integer, Shift As Integer)
    mblnClickOk = True
End Sub

Private Sub tdb_Lista_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub tdb_Lista_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    mblnClickOk = True
End Sub

Private Sub tdb_Lista_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    With tdb_Lista
        If (Not .EOF And Not .BOF) And mblnClickOk Then
            mblnClickOk = False
            txtpkID.Text = .Columns("PKID").Value
            If mblnPrimeiraVez Then
                LeDaTabelaParaObj gstrBanco, Me
                gCorLinhaSelecionada tdb_Lista
                HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrDeletar
                If mobjAux Is Nothing Then
                    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar
                Else
                    HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrAplicar
                End If
                strCodigoAtual = tdb_Lista.Columns("intBanco").Value
                strDescricaoAtual = tdb_Lista.Columns("strDescricao").Value
                txtintDigitoBanco = tdb_Lista.Columns("intDigitoBanco").Value
                PreencheImagem txtpkID
             End If
                mblnAlterando = True
        End If
    End With
End Sub

Public Sub MantemForm(ByVal strModoOperacao As String)
    
    If UCase(strModoOperacao) = gstrSalvar Then
        If blnDadosOk = False Then Exit Sub
            If ToolBarGeral(strModoOperacao, gstrBanco, mblnAlterando, tdb_Lista, Me, mobjAux, strQuery, , rptBanco, strQueryRelatorio) Then
                Set img_LogoBanco.Picture = Nothing
            End If
        Exit Sub
    ElseIf strModoOperacao = gstrLogotipo Then
        frmCadImagem.CadastraFoto img_LogoBanco, txtintLogoBanco
        frmCadImagem.Caption = "Logotipo"
        Exit Sub
    ElseIf (strModoOperacao) = gstrNovo Then
        Set img_LogoBanco.Picture = Nothing
    ElseIf UCase(strModoOperacao) = gstrDeletar Then
        If ToolBarGeral(strModoOperacao, gstrBanco, mblnAlterando, tdb_Lista, Me, mobjAux, strQuery, , rptBanco, strQueryRelatorio) Then
            Set img_LogoBanco.Picture = Nothing
        End If
        Exit Sub
    End If
    
    ToolBarGeral strModoOperacao, gstrBanco, mblnAlterando, tdb_Lista, Me, mobjAux, _
                 strQuery, , rptBanco, strQueryRelatorio
End Sub

Private Sub txtintBanco_GotFocus()
    MarcaCampo txtintBanco
End Sub

Private Sub txtintBanco_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub txtintDigitoBanco_GotFocus()
    MarcaCampo txtintDigitoBanco
End Sub

Private Sub txtintDigitoBanco_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub txtintSequencialDebAut_GotFocus()
    MarcaCampo txtintSequencialDebAut
End Sub

Private Sub txtintSequencialDebAut_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtintSequencialDebAut
End Sub

Private Sub txtstrSigla_GotFocus()
    MarcaCampo txtstrSigla
End Sub

Private Sub txtstrSigla_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub txtstrDescricao_GotFocus()
    MarcaCampo txtstrdescricao
End Sub

Private Sub txtstrDescricao_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Function strQueryRelatorio() As String
    Dim strsql As String
    strsql = ""
    strsql = strsql & "SELECT intBanco, strSigla, strDescricao "
    strsql = strsql & "FROM " & gstrBanco
    If mblnAlterando = True Then
        strsql = strsql & " WHERE PKId = " & Val(txtpkID)
    End If
    strQueryRelatorio = strsql
End Function

Private Function blnDadosOk() As Boolean
    blnDadosOk = False
    If Trim(txtintBanco) = "" Then
        ExibeMensagem "O código deve  ser preenchido corretamente."
        Exit Function
    ElseIf Trim(txtstrdescricao) = "" Then
        ExibeMensagem "A descrição deve ser preenchida corretamente."
        txtstrdescricao.SetFocus
        Exit Function
    End If
    
    If Not mblnAlterando Or (mblnAlterando And UCase$(strCodigoAtual) <> UCase$(txtintBanco.Text)) Then
        If gblnExisteCodigo(1, gstrBanco, "intBanco", txtintBanco.Text) Then
            ExibeMensagem "Já existe registro com esse código."
            Exit Function
        End If
    End If
    
    If Not mblnAlterando Or (mblnAlterando And UCase$(txtstrdescricao.Text) <> UCase$(strDescricaoAtual)) Then
        If gblnExisteCodigo(1, gstrBanco, "strDescricao", "'" & Trim(txtstrdescricao) & "'") Then
            ExibeMensagem "Já existe registro com essa descrição."
            Exit Function
        End If
    End If
    
    blnDadosOk = True
    
    
End Function

Private Sub PreencheImagem(intPkid As Long)
    Dim strsql As String
    Dim adoResultado As ADODB.Recordset
    
    strsql = "Select * from " & gstrBanco & " Where pkid = " & intPkid
    
    If gobjBanco.CriaADO(strsql, 5, adoResultado) Then
        If Not adoResultado.EOF Then
            LeImagem Val(gstrENulo(adoResultado("intLogoBanco").Value)), img_LogoBanco
        End If
    End If

End Sub
