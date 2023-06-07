VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frmResumoBancario 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Resumo Bancário"
   ClientHeight    =   6285
   ClientLeft      =   2025
   ClientTop       =   2535
   ClientWidth     =   7320
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6285
   ScaleWidth      =   7320
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab tab_3dPasta 
      Height          =   6060
      Left            =   120
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   105
      Width           =   7080
      _ExtentX        =   12488
      _ExtentY        =   10689
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Resumo Bancário"
      TabPicture(0)   =   "frmResumoBancario.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lbldblValor"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblContaBancaria"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblLote"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblDataDoMovimento"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblAgencia"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblBanco"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "dbc_strDescricao"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "tdb_ResumoBancario"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "dbcintContaBancaria"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtdtmData"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtPKId"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtintLote"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtdblValor"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txt_strAgencia"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txt_strBanco"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "cmd_ContaCorrente"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).ControlCount=   16
      Begin VB.CommandButton cmd_ContaCorrente 
         Height          =   315
         Left            =   6585
         Picture         =   "frmResumoBancario.frx":001C
         Style           =   1  'Graphical
         TabIndex        =   16
         TabStop         =   0   'False
         Tag             =   "585"
         ToolTipText     =   "Clique para cadastar uma Conta Bancária"
         Top             =   990
         Width           =   360
      End
      Begin VB.TextBox txt_strBanco 
         Height          =   315
         Left            =   1620
         TabIndex        =   4
         Top             =   1875
         Width           =   3555
      End
      Begin VB.TextBox txt_strAgencia 
         Height          =   315
         Left            =   1620
         TabIndex        =   3
         Top             =   1440
         Width           =   1515
      End
      Begin VB.TextBox txtdblValor 
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
         Left            =   3915
         TabIndex        =   6
         Top             =   2295
         Width           =   1260
      End
      Begin VB.TextBox txtintLote 
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
         Left            =   1620
         MaxLength       =   4
         TabIndex        =   5
         Top             =   2295
         Width           =   480
      End
      Begin VB.TextBox txtPKId 
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         Left            =   2130
         Locked          =   -1  'True
         OLEDragMode     =   1  'Automatic
         OLEDropMode     =   1  'Manual
         TabIndex        =   9
         Top             =   15
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtdtmData 
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
         Left            =   1620
         MaxLength       =   10
         TabIndex        =   0
         Top             =   570
         Width           =   1125
      End
      Begin MSDataListLib.DataCombo dbcintContaBancaria 
         Height          =   315
         HelpContextID   =   1
         Left            =   1620
         TabIndex        =   1
         Top             =   990
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Text            =   ""
      End
      Begin TrueOleDBGrid70.TDBGrid tdb_ResumoBancario 
         Height          =   3210
         Left            =   120
         TabIndex        =   7
         Top             =   2715
         Width           =   6810
         _ExtentX        =   12012
         _ExtentY        =   5662
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
         Columns(1).Caption=   "Data"
         Columns(1).DataField=   "DataMovimento"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Conta Corrente"
         Columns(2).DataField=   "ContaCorrente"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Lote"
         Columns(3).DataField=   "Lote"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "Valor"
         Columns(4).DataField=   "Valor"
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
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0).AllowSizing=0"
         Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
         Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(8)=   "Column(1).Width=1799"
         Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=1720"
         Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(13)=   "Column(2).Width=5741"
         Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=5662"
         Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(17)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(18)=   "Column(3).Width=1138"
         Splits(0)._ColumnProps(19)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(20)=   "Column(3)._WidthInPix=1058"
         Splits(0)._ColumnProps(21)=   "Column(3)._EditAlways=0"
         Splits(0)._ColumnProps(22)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(23)=   "Column(4).Width=2699"
         Splits(0)._ColumnProps(24)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(25)=   "Column(4)._WidthInPix=2619"
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
         _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=32,.parent=13"
         _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
         _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
         _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
         _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=50,.parent=13"
         _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
         _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
         _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
         _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=54,.parent=13"
         _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14"
         _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
         _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
         _StyleDefs(56)  =   "Named:id=33:Normal"
         _StyleDefs(57)  =   ":id=33,.parent=0"
         _StyleDefs(58)  =   "Named:id=34:Heading"
         _StyleDefs(59)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(60)  =   ":id=34,.wraptext=-1"
         _StyleDefs(61)  =   "Named:id=35:Footing"
         _StyleDefs(62)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(63)  =   "Named:id=36:Selected"
         _StyleDefs(64)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(65)  =   "Named:id=37:Caption"
         _StyleDefs(66)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(67)  =   "Named:id=38:HighlightRow"
         _StyleDefs(68)  =   ":id=38,.parent=33,.bgcolor=&H80000014&,.fgcolor=&H80000012&"
         _StyleDefs(69)  =   "Named:id=39:EvenRow"
         _StyleDefs(70)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(71)  =   "Named:id=40:OddRow"
         _StyleDefs(72)  =   ":id=40,.parent=33"
         _StyleDefs(73)  =   "Named:id=41:RecordSelector"
         _StyleDefs(74)  =   ":id=41,.parent=34"
         _StyleDefs(75)  =   "Named:id=42:FilterBar"
         _StyleDefs(76)  =   ":id=42,.parent=33"
      End
      Begin MSDataListLib.DataCombo dbc_strDescricao 
         Height          =   315
         Left            =   2985
         TabIndex        =   2
         Top             =   990
         Width           =   3600
         _ExtentX        =   6350
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Text            =   ""
      End
      Begin VB.Label lblBanco 
         AutoSize        =   -1  'True
         Caption         =   "Banco"
         Height          =   195
         Left            =   1065
         TabIndex        =   15
         Top             =   1965
         Width           =   465
      End
      Begin VB.Label lblAgencia 
         AutoSize        =   -1  'True
         Caption         =   "Agência"
         Height          =   195
         Left            =   945
         TabIndex        =   14
         Top             =   1560
         Width           =   585
      End
      Begin VB.Label lblDataDoMovimento 
         AutoSize        =   -1  'True
         Caption         =   "Data do Movimento"
         Height          =   195
         Left            =   135
         TabIndex        =   13
         Top             =   660
         Width           =   1395
      End
      Begin VB.Label lblLote 
         AutoSize        =   -1  'True
         Caption         =   "Lote"
         Height          =   195
         Left            =   1230
         TabIndex        =   12
         Top             =   2340
         Width           =   315
      End
      Begin VB.Label lblContaBancaria 
         AutoSize        =   -1  'True
         Caption         =   "Conta Corrente"
         Height          =   195
         Left            =   465
         TabIndex        =   11
         Top             =   1095
         Width           =   1065
      End
      Begin VB.Label lbldblValor 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Valor"
         Height          =   195
         Left            =   3510
         TabIndex        =   10
         Top             =   2340
         Width           =   360
      End
   End
End
Attribute VB_Name = "frmResumoBancario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim blnAlterando            As Boolean
    Dim bytOrdenacao            As Byte
    Dim blnOrdenacaoAsc         As Boolean
    Dim blnPrimeiraVez          As Boolean
    Dim dtmDataAtual            As Date
    Dim intContaBancariaAtual   As Long
    Dim intLoteAtual            As Integer

Private Sub cmd_ContaCorrente_Click()
    CarregaForm frmCadContasBancarias, dbcintcontabancaria
End Sub

Private Sub dbc_strDescricao_Change()
    If dbc_strDescricao.MatchedWithList Then
        If dbc_strDescricao.BoundText <> dbcintcontabancaria.BoundText Then
            PreencherListaDeOpcoes dbcintcontabancaria, dbc_strDescricao.BoundText
        End If
    End If
End Sub
Private Sub dbc_strDescricao_Click(Area As Integer)
    DropDownDataCombo dbc_strDescricao, Me, Area
End Sub
Private Sub dbcintcontabancaria_Click(Area As Integer)
    DropDownDataCombo dbcintcontabancaria, Me, Area
End Sub
Private Sub dbcintcontabancaria_Change()
    If dbcintcontabancaria.MatchedWithList Then
        If dbc_strDescricao.BoundText <> dbcintcontabancaria.BoundText Then
            PreencherListaDeOpcoes dbc_strDescricao, dbcintcontabancaria.BoundText
        End If
        PreencheAgBanco (dbcintcontabancaria.BoundText)
    End If
End Sub
Private Sub dbcintContaBancaria_GotFocus()
    MarcaCampo dbcintcontabancaria
End Sub
Private Sub dbcintContaBancaria_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbcintcontabancaria, Me, , KeyCode, Shift
End Sub
Private Sub dbcintContaBancaria_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbcintcontabancaria
End Sub

Private Sub dbcintContaBancaria_LostFocus()
    LeDaTabelaParaObj "", dbcintcontabancaria, strQueryContaCorrente
    If Not dbcintcontabancaria.MatchedWithList Then
        dbc_strDescricao.BoundText = ""
        Set dbc_strDescricao.RowSource = Nothing
        txt_strAgencia.Text = ""
        txt_strBanco.Text = ""
    End If
End Sub

Private Sub Form_Activate()
    gintCodSeguranca = 1108
    HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrDeletar
End Sub
Private Sub Form_Deactivate()
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar
End Sub
Private Sub Form_Load()
    dbcintcontabancaria.Tag = strQueryContaCorrente(True) & ";intNumeroConta"
    dbc_strDescricao.Tag = strQueryContaDescricao & ";descricao"
    TrocaCorObjeto txt_strAgencia, True
    TrocaCorObjeto txt_strBanco, True
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar
End Sub
Private Sub tdb_ResumoBancario_Click()
    blnPrimeiraVez = True
End Sub
Private Sub tdb_ResumoBancario_FilterChange()
    blnPrimeiraVez = False
    gblnFilraCampos tdb_ResumoBancario
End Sub

Private Sub tdb_resumobancario_HeadClick(ByVal ColIndex As Integer)
   
    gOrdenaGrid tdb_ResumoBancario, ColIndex
    blnPrimeiraVez = False
   
End Sub

Private Sub tdb_ResumoBancario_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDown, vbKeyUp, vbKeyLeft, vbKeyRight
            blnPrimeiraVez = True
    End Select
    
End Sub

Private Sub tdb_ResumoBancario_KeyPress(KeyAscii As Integer)
    Select Case tdb_ResumoBancario.Col
        Case 1
            CaracterValido KeyAscii, "D", tdb_ResumoBancario
        
        Case 2
            CaracterValido KeyAscii, "A", tdb_ResumoBancario
        Case 3
            CaracterValido KeyAscii, "N", tdb_ResumoBancario
        Case 4
            CaracterValido KeyAscii, "V", tdb_ResumoBancario
    End Select
End Sub

Private Sub tdb_ResumoBancario_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
Dim adoGrid As ADODB.Recordset
    If Button = 2 Then
        Set adoGrid = tdb_ResumoBancario.DataSource
        ImprimeRelatorioDoGrid rptGrid, adoGrid.Source, adoGrid.ActiveConnection, adoGrid, Me.tdb_ResumoBancario, "Resumo Bancário"
    End If
End Sub

Private Sub tdb_ResumoBancario_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    With tdb_ResumoBancario
        If Not .EOF And blnPrimeiraVez Then
            txtPKId.Text = .Columns("PKID").Value
            blnAlterando = True
            LeDaTabelaParaObj gstrResumoBancario, Me
            dtmDataAtual = tdb_ResumoBancario.Columns("Data").Value
            If Val(gstrENulo(dbcintcontabancaria.BoundText)) <= 0 Then
                ExibeMensagem "A conta " & gstrENulo(tdb_ResumoBancario.Columns(2).Value) & " não foi vinculada a um plano de conta."
            Else
                intContaBancariaAtual = dbcintcontabancaria.BoundText
            End If
            intLoteAtual = tdb_ResumoBancario.Columns("Lote").Value
        End If
    End With
End Sub

Public Sub MantemForm(ByVal strModoOperacao As String)
        
    Select Case UCase(strModoOperacao)
        Case Is = UCase(gstrSalvar)
            If Not blnDadosOk Then Exit Sub
            ToolBarGeral strModoOperacao, gstrResumoBancario, blnAlterando, tdb_ResumoBancario, Me, , strQuery(gstrSalvar)
                If txtdtmData.Text = "" And dbcintcontabancaria.BoundText = "" Then
                    dbc_strDescricao.Text = ""
                    txt_strAgencia.Text = ""
                    txt_strBanco.Text = ""
                End If
            If Not blnAlterando Then
                blnPrimeiraVez = False
            End If
        Case Is = UCase(gstrPreencherLista)
            PreencherListaDeOpcoes Me.ActiveControl
        Case Is = UCase(gstrSalvar)
            If txtdtmData.Text = "" And dbcintcontabancaria.BoundText = "" Then
                dbc_strDescricao.Text = ""
                txt_strAgencia.Text = ""
                txt_strBanco.Text = ""
            End If
        Case Is = UCase(gstrNovo)
            LimpaObjeto Me
            dbc_strDescricao.BoundText = ""
            txt_strAgencia.Text = ""
            txt_strBanco.Text = ""
            blnPrimeiraVez = False
            blnAlterando = False
        Case Else
            ToolBarGeral strModoOperacao, gstrResumoBancario, blnAlterando, tdb_ResumoBancario, Me, , strQuery
            If UCase(strModoOperacao) = "DELETAR" Then
                If txtdtmData.Text = "" And dbcintcontabancaria.BoundText = "" Then
                    MantemForm gstrNovo
                End If
            End If
    End Select
                 
End Sub
Private Function blnDadosOk()
    blnDadosOk = False
    
    If Not gblnDataValida(txtdtmData) Then
        ExibeMensagem "A Data informada não é válida."
        txtdtmData.SetFocus
        Exit Function
    End If
    
    If Not dbcintcontabancaria.MatchedWithList Then
        ExibeMensagem "Selecione uma Conta Corrente válida."
        dbcintcontabancaria.SetFocus
        Exit Function
    End If
    
    If txtintLote.Text = "" Then
        ExibeMensagem "O Lote deve ser preenchido."
        txtintLote.SetFocus
        Exit Function
    End If
    
    If CCur(txtdblValor.Text) = 0 Then
        ExibeMensagem "O Campo Valor deve ser informado."
        txtdblValor.SetFocus
        Exit Function
    End If
    
    
    If Not blnAlterando Or (blnAlterando And CDate(dtmDataAtual) <> CDate(txtdtmData.Text) And _
            Val(intContaBancariaAtual) <> Val(dbcintcontabancaria.BoundText) And _
            Val(intLoteAtual) <> Val(txtintLote.Text)) Then
    
        If gblnExisteCodigo(2, gstrResumoBancario, "dtmData", CDate(txtdtmData.Text), "intContaBancaria", dbcintcontabancaria.BoundText, "intLote", Val(txtintLote.Text)) Then
            ExibeMensagem "Já existe um registro com a mesma Data, Conta Bancária e Lote."
            txtdtmData.SetFocus
            Exit Function
        End If
    End If

    blnDadosOk = True
    
End Function

Private Sub txtdblValor_GotFocus()
    MarcaCampo txtdblValor
End Sub

Private Sub txtdblValor_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "V", txtdblValor
End Sub

Private Sub txtdblValor_LostFocus()
    txtdblValor = gstrConvVrDoSql(txtdblValor, 2)
End Sub

Private Sub txtdtmData_GotFocus()
    If txtdtmData = "" Then txtdtmData = gstrDataDoSistema
    MarcaCampo txtdtmData
End Sub

Private Sub txtdtmData_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "D", txtdtmData
End Sub

Private Sub txtdtmData_LostFocus()
    txtdtmData = gstrDataFormatada(txtdtmData)
End Sub

Private Sub txtintLote_GotFocus()
    MarcaCampo txtintLote
End Sub

Private Sub txtintLote_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtintLote
End Sub

Private Sub txtintLote_LostFocus()
    txtintLote = Format(txtintLote, "0000")
End Sub

Private Function strQuery(Optional strModoOperacao As String) As String
Dim strSql As String

    strSql = "SELECT RB.Pkid,"
    strSql = strSql & " RB.dtmData DataMovimento,"
    strSql = strSql & " RB.intLote Lote,"
    strSql = strSql & " RB.dblValor Valor, "
    strSql = strSql & gstrCONVERT(CDT_VARCHAR, "CB.strConta") & strCONCAT & "'-'" & strCONCAT & " CB.strDigitoVerificador ContaCorrente"
    strSql = strSql & " FROM "
    strSql = strSql & gstrResumoBancario & " RB, "
    strSql = strSql & gstrContaBancaria & " CB"
    strSql = strSql & " WHERE"
    strSql = strSql & " RB.intContaBancaria " & strOUTJSQLServer & "= CB.Pkid " & strOUTJOracle
       
    If UCase(strModoOperacao) = UCase(gstrSalvar) Then
        If Not blnAlterando Then
            strSql = strSql & " AND RB.Pkid = " & glngPegaUltimaChave(gstrResumoBancario, "Pkid") + 1
        Else
            strSql = strSql & " AND RB.Pkid = " & txtPKId.Text
        End If
    End If
    
    Select Case bytOrdenacao
        Case Is = 1
            strSql = strSql & " ORDER BY RB.dtmData " & IIf(blnOrdenacaoAsc, " ASC", " DESC")
        Case Is = 2
            'strSql = strSql & " ORDER BY CB.intNumeroConta, " & gstrCONVERT(cdt_numeric, "CB.strDigitoVerificador") & IIf(blnOrdenacaoAsc, " ASC", " DESC")
            strSql = strSql & " ORDER BY CB.intNumeroConta " & IIf(blnOrdenacaoAsc, " ASC", " DESC")
        Case Is = 3
            strSql = strSql & " ORDER BY RB.intLote" & IIf(blnOrdenacaoAsc, " ASC", " DESC")
        Case Is = 4
            strSql = strSql & " ORDER BY RB.dblValor" & IIf(blnOrdenacaoAsc, " ASC", " DESC")
    End Select
    
    strQuery = strSql

End Function

Private Function strQueryContaCorrente(Optional blnF5 As Boolean) As String
    Dim strSql As String

    strSql = "SELECT CB.Pkid, "
    strSql = strSql & "intNumeroConta ContaCorrente"
    strSql = strSql & " FROM " & gstrContaBancaria & " CB, "
    strSql = strSql & gstrPlanoConta & " PC"
    strSql = strSql & " Where"
    strSql = strSql & " CB.Pkid = PC.Intcontabancaria"
    If blnF5 = False Then
        strSql = strSql & " AND CB.intNumeroConta = " & Val(dbcintcontabancaria.Text)
    End If
    strSql = strSql & " ORDER BY intNumeroConta, strDigitoVerificador"
    
    strQueryContaCorrente = strSql

End Function

Private Sub PreencheAgBanco(lngPkidContaBancaria As Long)
Dim adoResultado    As ADODB.Recordset
Dim strSql          As String

    strSql = "SELECT BA.strDescricao Banco,"
    strSql = strSql & " AG.strDescricao Agencia"
    strSql = strSql & " FROM "
    strSql = strSql & gstrContaBancaria & " CB, "
    strSql = strSql & gstrBanco & " BA, "
    strSql = strSql & gstrAgencia & " AG"
    strSql = strSql & " WHERE"
    strSql = strSql & " BA.Pkid " & strOUTJOracle & "=" & strOUTJSQLServer & " CB.intBanco AND"
    strSql = strSql & " AG.Pkid " & strOUTJOracle & "=" & strOUTJSQLServer & " CB.intAgencia AND"
    strSql = strSql & " CB.Pkid = " & lngPkidContaBancaria

    Set gobjBanco = New clsBanco
    
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        If Not adoResultado.EOF Then
            txt_strAgencia.Text = gstrENulo(adoResultado!Agencia)
            txt_strBanco.Text = gstrENulo(adoResultado!Banco)
        Else
            txt_strAgencia.Text = ""
            txt_strBanco.Text = ""
        End If
    End If
    
End Sub

Private Function strQueryContaDescricao() As String
    Dim strSql As String

    strSql = strSql & "SELECT "
    strSql = strSql & "CB.Pkid, "
    strSql = strSql & "CB.strdescricao " & strCONCAT & "'('" & strCONCAT & " CB.strConta " & strCONCAT & "'-'" & strCONCAT & " CB.strdigitoverificador" & strCONCAT & "')' Descricao"
    strSql = strSql & " FROM " & gstrContaBancaria & " CB, "
    strSql = strSql & gstrPlanoConta & " PC"
    strSql = strSql & " Where"
    strSql = strSql & " CB.Pkid = PC.Intcontabancaria"
    strSql = strSql & " ORDER BY CB.strdescricao"
    
    strQueryContaDescricao = strSql

End Function

