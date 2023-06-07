VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "TODG7.OCX"
Begin VB.Form frmCadOrgao 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Órgãos"
   ClientHeight    =   4830
   ClientLeft      =   855
   ClientTop       =   2235
   ClientWidth     =   8520
   ClipControls    =   0   'False
   HelpContextID   =   24
   Icon            =   "CadOrgao.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form4"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4830
   ScaleWidth      =   8520
   Begin VB.TextBox txtPKId 
      Height          =   285
      Left            =   870
      TabIndex        =   18
      Top             =   120
      Visible         =   0   'False
      Width           =   525
   End
   Begin TabDlg.SSTab tab_3dPasta 
      Height          =   4650
      Left            =   120
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   90
      Width           =   8325
      _ExtentX        =   14684
      _ExtentY        =   8202
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Órgãos"
      TabPicture(0)   =   "CadOrgao.frx":1042
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lbllintTipoDeAdm"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblintPoder"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblstrDirigente"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblintUnidadeFinanceira"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblstrCNPJ"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblstrDescricao"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lblstrCodigo"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lbldblLimite"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "tdb_Orgao"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtstrDirigente"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "cmd_TipoDeAdm"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "cmd_Poder"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "cmd_UG"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtstrCNPJ"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txtstrDescricao"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txtstrCodigo"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "dbcintTipoDeAdm"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "dbcintPoder"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "dbcintUnidadeFinanceira"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "txtintExercicio"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "txtdblLimite"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).ControlCount=   21
      Begin VB.TextBox txtdblLimite 
         Height          =   285
         Left            =   7350
         MaxLength       =   25
         OLEDragMode     =   1  'Automatic
         TabIndex        =   3
         Top             =   390
         Width           =   810
      End
      Begin VB.TextBox txtintExercicio 
         Enabled         =   0   'False
         Height          =   285
         Left            =   5400
         TabIndex        =   2
         Top             =   390
         Visible         =   0   'False
         Width           =   735
      End
      Begin MSDataListLib.DataCombo dbcintUnidadeFinanceira 
         Height          =   315
         Left            =   1845
         TabIndex        =   9
         Top             =   1410
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dbcintPoder 
         Height          =   315
         Left            =   5460
         TabIndex        =   7
         Top             =   1050
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dbcintTipoDeAdm 
         Height          =   315
         Left            =   1845
         TabIndex        =   5
         Top             =   1050
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
      Begin VB.TextBox txtstrCodigo 
         Height          =   285
         Left            =   1845
         MaxLength       =   8
         TabIndex        =   0
         Top             =   390
         Width           =   1050
      End
      Begin VB.TextBox txtstrDescricao 
         Height          =   285
         Left            =   1845
         MaxLength       =   100
         TabIndex        =   4
         Top             =   720
         Width           =   6315
      End
      Begin VB.TextBox txtstrCNPJ 
         Height          =   285
         Left            =   3720
         MaxLength       =   18
         TabIndex        =   1
         Top             =   390
         Width           =   1650
      End
      Begin VB.CommandButton cmd_UG 
         Height          =   300
         Left            =   4275
         Picture         =   "CadOrgao.frx":105E
         Style           =   1  'Graphical
         TabIndex        =   10
         TabStop         =   0   'False
         ToolTipText     =   "Cliqui aqui para cadastrar unidade financeira"
         Top             =   1425
         Width           =   360
      End
      Begin VB.CommandButton cmd_Poder 
         Height          =   300
         Left            =   7845
         Picture         =   "CadOrgao.frx":13E8
         Style           =   1  'Graphical
         TabIndex        =   8
         TabStop         =   0   'False
         ToolTipText     =   "Cliqui aqui para cadastrar poder"
         Top             =   1065
         Width           =   360
      End
      Begin VB.CommandButton cmd_TipoDeAdm 
         Height          =   300
         Left            =   4275
         Picture         =   "CadOrgao.frx":1772
         Style           =   1  'Graphical
         TabIndex        =   6
         TabStop         =   0   'False
         ToolTipText     =   "Cliqui aqui para cadastrar tipo de administração"
         Top             =   1065
         Width           =   360
      End
      Begin VB.TextBox txtstrDirigente 
         Height          =   285
         Left            =   5460
         MaxLength       =   27
         TabIndex        =   11
         Top             =   1440
         Width           =   2745
      End
      Begin TrueOleDBGrid70.TDBGrid tdb_Orgao 
         Height          =   2715
         Left            =   120
         TabIndex        =   12
         Top             =   1830
         Width           =   8085
         _ExtentX        =   14261
         _ExtentY        =   4789
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "PKID"
         Columns(0).DataField=   "PKId"
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
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0).AllowSizing=0"
         Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
         Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(8)=   "Column(1).Width=1614"
         Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=1535"
         Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(13)=   "Column(2).Width=12118"
         Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=12039"
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
         _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=33,.bgcolor=&H80000014&"
         _StyleDefs(20)  =   ":id=8,.fgcolor=&H80000012&"
         _StyleDefs(21)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
         _StyleDefs(22)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
         _StyleDefs(23)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(24)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
         _StyleDefs(25)  =   "Splits(0).Style:id=13,.parent=1"
         _StyleDefs(26)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
         _StyleDefs(27)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
         _StyleDefs(28)  =   "Splits(0).FooterStyle:id=15,.parent=3"
         _StyleDefs(29)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
         _StyleDefs(30)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
         _StyleDefs(31)  =   "Splits(0).EditorStyle:id=17,.parent=7"
         _StyleDefs(32)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
         _StyleDefs(33)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
         _StyleDefs(34)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
         _StyleDefs(35)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(36)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(37)  =   "Splits(0).Columns(0).Style:id=46,.parent=13"
         _StyleDefs(38)  =   "Splits(0).Columns(0).HeadingStyle:id=43,.parent=14"
         _StyleDefs(39)  =   "Splits(0).Columns(0).FooterStyle:id=44,.parent=15"
         _StyleDefs(40)  =   "Splits(0).Columns(0).EditorStyle:id=45,.parent=17"
         _StyleDefs(41)  =   "Splits(0).Columns(1).Style:id=28,.parent=13"
         _StyleDefs(42)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14"
         _StyleDefs(43)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
         _StyleDefs(44)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
         _StyleDefs(45)  =   "Splits(0).Columns(2).Style:id=32,.parent=13"
         _StyleDefs(46)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
         _StyleDefs(47)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
         _StyleDefs(48)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
         _StyleDefs(49)  =   "Named:id=33:Normal"
         _StyleDefs(50)  =   ":id=33,.parent=0"
         _StyleDefs(51)  =   "Named:id=34:Heading"
         _StyleDefs(52)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(53)  =   ":id=34,.wraptext=-1"
         _StyleDefs(54)  =   "Named:id=35:Footing"
         _StyleDefs(55)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(56)  =   "Named:id=36:Selected"
         _StyleDefs(57)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(58)  =   "Named:id=37:Caption"
         _StyleDefs(59)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(60)  =   "Named:id=38:HighlightRow"
         _StyleDefs(61)  =   ":id=38,.parent=33,.bgcolor=&H80000014&,.fgcolor=&H80000012&"
         _StyleDefs(62)  =   "Named:id=39:EvenRow"
         _StyleDefs(63)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(64)  =   "Named:id=40:OddRow"
         _StyleDefs(65)  =   ":id=40,.parent=33"
         _StyleDefs(66)  =   "Named:id=41:RecordSelector"
         _StyleDefs(67)  =   ":id=41,.parent=34"
         _StyleDefs(68)  =   "Named:id=42:FilterBar"
         _StyleDefs(69)  =   ":id=42,.parent=33"
      End
      Begin VB.Label lbldblLimite 
         AutoSize        =   -1  'True
         Caption         =   "Limite (%)"
         Height          =   195
         Left            =   6600
         TabIndex        =   22
         Top             =   420
         Width           =   660
      End
      Begin VB.Label lblstrCodigo 
         AutoSize        =   -1  'True
         Caption         =   "Código"
         Height          =   195
         Left            =   1290
         TabIndex        =   21
         Top             =   420
         Width           =   495
      End
      Begin VB.Label lblstrDescricao 
         AutoSize        =   -1  'True
         Caption         =   "Descrição"
         Height          =   195
         Left            =   1065
         TabIndex        =   20
         Top             =   765
         Width           =   720
      End
      Begin VB.Label lblstrCNPJ 
         AutoSize        =   -1  'True
         Caption         =   "CNPJ"
         Height          =   195
         Left            =   3240
         TabIndex        =   19
         Top             =   420
         Width           =   405
      End
      Begin VB.Label lblintUnidadeFinanceira 
         AutoSize        =   -1  'True
         Caption         =   "Unidade Financeira"
         Height          =   195
         Left            =   405
         TabIndex        =   17
         Top             =   1455
         Width           =   1380
      End
      Begin VB.Label lblstrDirigente 
         AutoSize        =   -1  'True
         Caption         =   "Dirigente"
         Height          =   195
         Left            =   4770
         TabIndex        =   16
         Top             =   1455
         Width           =   630
      End
      Begin VB.Label lblintPoder 
         AutoSize        =   -1  'True
         Caption         =   "Poder"
         Height          =   195
         Left            =   4980
         TabIndex        =   15
         Top             =   1110
         Width           =   420
      End
      Begin VB.Label lbllintTipoDeAdm 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Administração"
         Height          =   195
         Left            =   210
         TabIndex        =   14
         Top             =   1110
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frmCadOrgao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mblnAlterando         As Boolean
Dim mobjAux               As Object
Dim mblnselecionou        As Boolean
Dim mblnClickOk           As Boolean
Dim intFiltroExercicio    As Integer
Public mIntCodSeguranca   As Integer
Public strcodigoantigo       As String
Dim strCodigoAtual  As String
Dim strDescricaoAtual As String
   
Private Function blnDadosOk() As Boolean
Dim strWhereComplementar    As String
    blnDadosOk = False
    
    'Incluido ORC1550 para impedir inclusão de descricoes repetidas no mesmo exercicio
    If mblnAlterando Then
        strWhereComplementar = " AND PKID <> " & Me.txtPKId.Text
    Else
        strWhereComplementar = ""
    End If
    
    
    If Trim(txtstrCodigo.Text) = Space$(0) Then
        ExibeMensagem "O código tem que ser informado."
        txtstrCodigo.SetFocus
        Exit Function
    End If
    
    If Trim(txtdblLimite.Text) = "" Then
        ExibeMensagem "O limite tem que ser informado."
        txtdblLimite.SetFocus
        Exit Function
    End If
    
    If txtstrDescricao.Text = Space$(0) Then
        ExibeMensagem "A descrição tem que ser informada."
        txtstrDescricao.SetFocus
        Exit Function
    End If
    
    If dbcintPoder.BoundText = Space$(0) Then
        ExibeMensagem "O poder tem que ser informado."
        dbcintPoder.SetFocus
        Exit Function
    End If
        
    If Not mblnAlterando Or (mblnAlterando And UCase$(txtstrCodigo.Text) <> UCase$(strCodigoAtual)) Then
         If gblnExisteCodigo(1, gstrOrgao, "strCodigo", "'" & txtstrCodigo.Text & "'", , , , , " AND intExercicio = " & IIf(gbytMenu = gbytMenuProposta, gintExercicio + 1, gintExercicio)) Then
            ExibeMensagem "O código digitado já se encontra cadastrado!"
            txtstrCodigo.SetFocus
            Exit Function
        End If
    End If
    
    If Not mblnAlterando Or (mblnAlterando And UCase$(txtstrDescricao.Text) <> UCase$(strDescricaoAtual)) Then
        If gblnExisteCodigo(2, gstrOrgao, "strDescricao", "'" & txtstrDescricao.Text & "'", "intExercicio", Str(intFiltroExercicio), , , strWhereComplementar) Then
             ExibeMensagem "A descrição digitada já se encontra cadastrada!"
             txtstrDescricao.SetFocus
             Exit Function
        End If
    End If
    

'
'    If (txtstrDescricao.Text <> txtstrDescricao.Tag) Then
'            If gblnExisteCodigo(1, gstrOrgao, "strDescricao", "'" & txtstrDescricao & "'", , , , , " AND intExercicio = " & IIf(gbytMenu = gbytMenuProposta, gintExercicio + 1, gintExercicio)) Then
'                ExibeMensagem "A descrição digitada já se encontra cadastrada!"
'                txtstrDescricao.SetFocus
'                Exit Function
'            End If
'    End If
    
    blnDadosOk = True
    
End Function

Private Sub dbcintPoder_Click(Area As Integer)
   DropDownDataCombo dbcintPoder, Me, Area
End Sub

Private Sub dbcintPoder_KeyDown(KeyCode As Integer, Shift As Integer)
   DropDownDataCombo dbcintPoder, Me, , KeyCode, Shift
End Sub

Private Sub dbcintPoder_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", dbcintPoder
End Sub

Private Sub dbcintTipoDeAdm_Click(Area As Integer)
   DropDownDataCombo dbcintTipoDeAdm, Me, Area
End Sub

Private Sub dbcintTipoDeAdm_KeyDown(KeyCode As Integer, Shift As Integer)
   DropDownDataCombo dbcintTipoDeAdm, Me, , KeyCode, Shift
End Sub

Private Sub dbcintTipoDeAdm_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", dbcintUnidadeFinanceira
End Sub

Private Sub dbcintUnidadeFinanceira_Click(Area As Integer)
   DropDownDataCombo dbcintUnidadeFinanceira, Me, Area
End Sub

Private Sub dbcintUnidadeFinanceira_KeyDown(KeyCode As Integer, Shift As Integer)
   DropDownDataCombo dbcintUnidadeFinanceira, Me, , KeyCode, Shift
End Sub

Private Sub dbcintUnidadeFinanceira_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", dbcintUnidadeFinanceira
End Sub

Private Sub cmd_Poder_Click()
    CarregaForm frmCadPoder, dbcintPoder
End Sub

Private Sub cmd_TipoDeAdm_Click()
    CarregaForm frmCadTipoDeAdministracao, dbcintTipoDeAdm
End Sub

Private Sub cmd_UG_Click()
    CarregaForm frmCadUnidadeFinanceira, dbcintUnidadeFinanceira
End Sub

Private Sub Form_Activate()
    gintCodSeguranca = mIntCodSeguranca
    VirificaGradeListView Me
'=============
    If mblnselecionou Then
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

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = vbKeyF1 Then
        Call_HtmlHelp Me.HelpContextID
    End If
End Sub

Private Sub Form_Load()
    
    'Vamos verificar qual menu que chamou o form, para definirmos o filtro
     intFiltroExercicio = IIf(gbytMenu = gbytMenuProposta, gintExercicio + 1, gintExercicio)
    'If gbytMenu = gbytMenuCadastro Then
        'intFiltroExercicio = gintExercicio
    'Else
       ' intFiltroExercicio = gintExercicio + 1
    'End If
    
    txtintExercicio = intFiltroExercicio
    
    mblnAlterando = False
    LeDaTabelaParaObj gstrTipoAdministracao, dbcintTipoDeAdm
    LeDaTabelaParaObj gstrPoder, dbcintPoder
    LeDaTabelaParaObj gstrUnidadeFinanceira, dbcintUnidadeFinanceira
    VerificaListaAutomatica gstrOrgao, tdb_Orgao, strQueryOrgao
    VerificaObjParaAplicar mobjAux
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar, gstrDeletar
    mblnselecionou = False
End Sub

Private Sub tdb_Orgao_Click()
    If glngQtdLinhaTDBGrid(tdb_Orgao) = 1 Then
        tdb_Orgao_RowColChange 0, 0
    End If
End Sub

Private Sub tdb_Orgao_DblClick()
    MantemForm gstrAplicar
End Sub

Private Sub tdb_Orgao_FilterChange()
    mblnClickOk = False
    gblnFilraCampos tdb_Orgao
End Sub

Private Sub tdb_Orgao_HeadClick(ByVal ColIndex As Integer)
    gOrdenaGrid tdb_Orgao, ColIndex
End Sub

Private Sub tdb_Orgao_KeyPress(KeyAscii As Integer)
 If tdb_Orgao.Col = 1 Then
     CaracterValido KeyAscii, "N", tdb_Orgao
 End If
End Sub

Private Sub tdb_Orgao_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mblnClickOk = True
End Sub

Private Sub tdb_Orgao_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    With tdb_Orgao
        If (Not .EOF And Not .BOF) And mblnClickOk Then
            mblnClickOk = False
            txtPKId.Text = .Columns("PKID").Value
            LeDaTabelaParaObj gstrOrgao, Me
            
            PreencheControles
            strcodigoantigo = .Columns("Código").Value
            txtstrDescricao.Tag = txtstrDescricao.Text
            
            HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrDeletar
            If mobjAux Is Nothing Then
                HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar
            Else
                HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrAplicar
            End If
            TrocaCorObjeto txtdblLimite, True
            txtdblLimite.BackColor = vbWindowBackground
            txtdblLimite.Enabled = True
            mblnselecionou = True
            mblnAlterando = True
        End If
        strCodigoAtual = txtstrCodigo
        strDescricaoAtual = txtstrDescricao
    End With
End Sub


Private Sub PreencheControles()

Dim strSql          As String
Dim adoResultado    As ADODB.Recordset
        

   strSql = "SELECT intTipoDeAdm, intPoder, intUnidadeFinanceira "
   strSql = strSql & " FROM " & gstrOrgao & " O"
   strSql = strSql & " WHERE O.PKID = " & txtPKId.Text
 
   Set gobjBanco = New clsBanco
   
   If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
       With adoResultado
           While Not .EOF
               dbcintTipoDeAdm.BoundText = IIf(IsNull(!intTipoDeAdm), "", !intTipoDeAdm)
               dbcintUnidadeFinanceira.BoundText = IIf(IsNull(!intUnidadeFinanceira), "", !intUnidadeFinanceira)
               dbcintPoder.BoundText = IIf(IsNull(!intPoder), "", !intPoder)
               .MoveNext
           Wend
       End With
   End If


End Sub

Public Sub MantemForm(ByVal strModoOperacao As String)

    If UCase(strModoOperacao) = UCase(gstrSalvar) Then
            
        If blnDadosOk Then
           ToolBarGeral strModoOperacao, gstrOrgao, _
                         mblnAlterando, tdb_Orgao, Me, _
                         mobjAux, strQueryOrgao, , rptOrgao, strQueryRelatorio
        End If
        
    Else
       
       ToolBarGeral strModoOperacao, gstrOrgao, mblnAlterando, tdb_Orgao, Me, mobjAux, strQueryOrgao, strQueryAplicar, rptOrgao, strQueryRelatorio
       
        
        If UCase(strModoOperacao) = UCase(gstrNovo) Or UCase(strModoOperacao) = UCase(gstrLimpar) Or UCase(strModoOperacao) = UCase(gstrDeletar) Then
            strcodigoantigo = ""
            txtintExercicio = intFiltroExercicio
            txtstrDescricao.Tag = Space$(0)
            mblnAlterando = False
        End If
        
    End If
    
End Sub

Private Function strQueryOrgao() As String

Dim strSql  As String
    
    strSql = " SELECT PKId, strCodigo, strDescricao FROM " & gstrOrgao & _
             " WHERE intExercicio = " & intFiltroExercicio & _
             " ORDER BY " & gstrCONVERT(cdt_numeric, "strCodigo")
    
    strQueryOrgao = strSql
    
End Function

Public Function strQueryRelatorio()
   
Dim strSql  As String

    strSql = "SELECT OG.strCodigo, OG.strDescricao AS Orgao, OG.strCNPJ, "
    strSql = strSql & "TA.strDescricao TipoAdm, PO.strDescricao AS Poder, "
    strSql = strSql & "UN.strDescricao AS UnidadeFinanceira, OG.strDirigente "
    strSql = strSql & "FROM "
    strSql = strSql & gstrOrgao & " OG, "
    strSql = strSql & gstrTipoAdministracao & " TA, "
    strSql = strSql & gstrPoder & " PO, "
    strSql = strSql & gstrUnidadeFinanceira & " UN "
    strSql = strSql & "WHERE OG.intTipoDeAdm " & strOUTJSQLServer & "= TA.PKId " & strOUTJOracle
    strSql = strSql & "AND OG.intPoder " & strOUTJSQLServer & "= PO.PKId " & strOUTJOracle
    strSql = strSql & "AND OG.intUnidadeFinanceira " & strOUTJSQLServer & "= UN.PKId " & strOUTJOracle
    strSql = strSql & "AND OG.intExercicio = " & intFiltroExercicio
    strSql = strSql & " ORDER BY " & gstrCONVERT(cdt_numeric, "OG.strCodigo")
         
    strQueryRelatorio = strSql
    
End Function

Private Sub txtdblLimite_GotFocus()
   MarcaCampo txtdblLimite
End Sub

Private Sub txtdblLimite_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "V", txtdblLimite
End Sub

Private Sub txtdblLimite_LostFocus()
    txtdblLimite = gstrConvVrDoSql(txtdblLimite)
End Sub

Private Sub txtstrCNPJ_GotFocus()
    txtstrCNPJ = gstrValorSemMascara(txtstrCNPJ)
    MarcaCampo txtstrCNPJ
End Sub

Private Sub txtstrCNPJ_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtstrCNPJ
End Sub

Private Sub txtstrCNPJ_LostFocus()
    txtstrCNPJ = gstrCGCCPFFormatado(txtstrCNPJ)
End Sub

Private Sub txtstrCodigo_Change()
    If txtstrCodigo.Text = "" Then
    mblnAlterando = False
    End If
End Sub

Private Sub txtstrCodigo_GotFocus()

    gstrProximoCodigo txtstrCodigo, gstrOrgao, "strCodigo", gintCodSeguranca, , , , , , , "intExercicio", CStr(intFiltroExercicio)
    txtstrCodigo = gstrValorSemMascara(txtstrCodigo)

    MarcaCampo txtstrCodigo
    
End Sub

Private Sub txtstrCodigo_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtstrCodigo
End Sub

Private Sub txtstrDescricao_GotFocus()
    MarcaCampo txtstrDescricao
End Sub

Private Sub txtstrDescricao_KeyPress(KeyAscii As Integer)
   CaracterValido KeyAscii, "A", txtstrDescricao
End Sub

Private Sub txtstrDirigente_GotFocus()
    MarcaCampo txtstrDirigente
End Sub

Private Sub txtstrDirigente_KeyPress(KeyAscii As Integer)
  CaracterValido KeyAscii, "A", txtstrDirigente
End Sub

Private Function strQueryAplicar() As String

    strQueryAplicar = " SELECT PKId, strDescricao FROM " & gstrOrgao & " WHERE intExercicio = " & intFiltroExercicio

End Function


