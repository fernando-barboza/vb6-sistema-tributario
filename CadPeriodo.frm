VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frmCadPeriodo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Per�odo"
   ClientHeight    =   4170
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5535
   Icon            =   "CadPeriodo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4170
   ScaleWidth      =   5535
   Begin TabDlg.SSTab tab_3dPasta 
      Height          =   3975
      Left            =   120
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   90
      Width           =   5310
      _ExtentX        =   9366
      _ExtentY        =   7011
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "Per�odo"
      TabPicture(0)   =   "CadPeriodo.frx":1042
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblstrDescricao"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblintIntervalo"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "txtPKId"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txtstrDescricao"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtintIntervalo"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "fra_Periodo"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "tdb_Periodo"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      Begin TrueOleDBGrid70.TDBGrid tdb_Periodo 
         Height          =   2505
         Left            =   120
         TabIndex        =   11
         Top             =   1350
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   4419
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "PKId"
         Columns(0).DataField=   "PKId"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Descri��o"
         Columns(1).DataField=   "strDescricao"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   2
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
         Splits(0)._ColumnProps(0)=   "Columns.Count=2"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0).Visible=0"
         Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(7)=   "Column(1).Width=8387"
         Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=8308"
         Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
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
         _StyleDefs(38)  =   "Named:id=33:Normal"
         _StyleDefs(39)  =   ":id=33,.parent=0"
         _StyleDefs(40)  =   "Named:id=34:Heading"
         _StyleDefs(41)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(42)  =   ":id=34,.wraptext=-1"
         _StyleDefs(43)  =   "Named:id=35:Footing"
         _StyleDefs(44)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(45)  =   "Named:id=36:Selected"
         _StyleDefs(46)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(47)  =   "Named:id=37:Caption"
         _StyleDefs(48)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(49)  =   "Named:id=38:HighlightRow"
         _StyleDefs(50)  =   ":id=38,.parent=33,.bgcolor=&H80000014&,.fgcolor=&H80000012&"
         _StyleDefs(51)  =   "Named:id=39:EvenRow"
         _StyleDefs(52)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(53)  =   "Named:id=40:OddRow"
         _StyleDefs(54)  =   ":id=40,.parent=33"
         _StyleDefs(55)  =   "Named:id=41:RecordSelector"
         _StyleDefs(56)  =   ":id=41,.parent=34"
         _StyleDefs(57)  =   "Named:id=42:FilterBar"
         _StyleDefs(58)  =   ":id=42,.parent=33"
      End
      Begin VB.Frame fra_Periodo 
         Caption         =   " Per�odo "
         Height          =   495
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   5055
         Begin VB.OptionButton optbytPeriodo 
            Caption         =   "Anos"
            Height          =   195
            Index           =   2
            Left            =   4110
            TabIndex        =   10
            Top             =   240
            Width           =   765
         End
         Begin VB.OptionButton optbytPeriodo 
            Caption         =   "Meses"
            Height          =   195
            Index           =   0
            Left            =   2740
            TabIndex        =   9
            Top             =   240
            Width           =   765
         End
         Begin VB.OptionButton optbytPeriodo 
            Caption         =   "Dias"
            Height          =   195
            Index           =   1
            Left            =   1490
            TabIndex        =   8
            Top             =   240
            Width           =   645
         End
         Begin VB.OptionButton optbytPeriodo 
            Caption         =   "Horas"
            Height          =   195
            Index           =   3
            Left            =   150
            TabIndex        =   7
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.TextBox txtintIntervalo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4260
         MaxLength       =   4
         MultiLine       =   -1  'True
         TabIndex        =   5
         Top             =   390
         Width           =   915
      End
      Begin VB.TextBox txtstrDescricao 
         Height          =   285
         Left            =   900
         MaxLength       =   30
         TabIndex        =   0
         Top             =   390
         Width           =   2535
      End
      Begin VB.TextBox txtPKId 
         Height          =   285
         Left            =   1830
         TabIndex        =   2
         Top             =   15
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.Label lblintIntervalo 
         AutoSize        =   -1  'True
         Caption         =   "Intervalo"
         Height          =   195
         Left            =   3600
         TabIndex        =   4
         Top             =   450
         Width           =   615
      End
      Begin VB.Label lblstrDescricao 
         AutoSize        =   -1  'True
         Caption         =   "Descri��o"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   435
         Width           =   720
      End
   End
End
Attribute VB_Name = "frmCadPeriodo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim mblnAlterando   As Boolean
    Dim mblnSelecionou  As Boolean
    Dim mblnClickOk     As Boolean
    Dim mobjAux         As Object

Private Sub Form_Activate()

    gintCodSeguranca = 499

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

Private Sub tdb_Periodo_KeyDown(KeyCode As Integer, Shift As Integer)
    mblnClickOk = True
End Sub

Private Sub tdb_Periodo_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    mblnClickOk = True
End Sub

Private Sub Form_Deactivate()
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar
End Sub

Private Sub Form_Load()
    mblnAlterando = False
    VerificaListaAutomatica gstrPeriodo, tdb_Periodo
    VerificaObjParaAplicar mobjAux
End Sub

Private Sub tdb_Periodo_Click()
    If glngQtdLinhaTDBGrid(tdb_Periodo) = 1 Then
        tdb_Periodo_RowColChange 0, 0
    End If
End Sub

Private Sub txtstrDescricao_GotFocus()
    MarcaCampo txtstrDescricao
End Sub

Private Sub txtstrDescricao_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub tdb_Periodo_DblClick()
    MantemForm gstrAplicar
End Sub

Private Sub tdb_Periodo_FilterChange()
    gblnFilraCampos tdb_Periodo
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar
    mblnSelecionou = False
End Sub

Private Sub tdb_Periodo_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    With tdb_Periodo
        If (Not .EOF And Not .BOF) And mblnClickOk Then
            mblnClickOk = False
            mblnAlterando = True
            txtPKId.Text = .Columns("PKID").Value
            LeDaTabelaParaObj gstrPeriodo, Me
            gCorLinhaSelecionada tdb_Periodo
        End If
    End With
End Sub

Public Sub MantemForm(ByVal strModoOperacao As String)
    ToolBarGeral strModoOperacao, gstrPeriodo, mblnAlterando, tdb_Periodo, Me, mobjAux
End Sub
