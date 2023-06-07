VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "Tabctl32.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frmCadResponsavelPatrimonio 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cadastro de Responsáveis"
   ClientHeight    =   6030
   ClientLeft      =   3030
   ClientTop       =   3105
   ClientWidth     =   7680
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6030
   ScaleWidth      =   7680
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtPKId 
      Height          =   270
      Left            =   5460
      TabIndex        =   11
      Top             =   90
      Visible         =   0   'False
      Width           =   645
   End
   Begin TabDlg.SSTab tab_3dPasta 
      Height          =   5715
      Left            =   120
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   180
      Width           =   7440
      _ExtentX        =   13123
      _ExtentY        =   10081
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Responsáveis"
      TabPicture(0)   =   "frmCadResponsavelPatrimonio.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lbl_strMatricula"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lbl_intLocais"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lbl_dtmDtDataCadastro"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lbl_strOservacao"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblstrNome"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblstrCargo"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "tdb_Lista"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtstrMatricula"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtstrObservacao"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "dbcintLocal"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtstrResponsavel"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtstrCargo"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtdtmDtDemissao"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txtdtmDtCadastro"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).ControlCount=   15
      Begin VB.TextBox txtdtmDtCadastro 
         Height          =   285
         Left            =   4425
         MaxLength       =   10
         TabIndex        =   1
         Top             =   420
         Width           =   990
      End
      Begin VB.TextBox txtdtmDtDemissao 
         Height          =   285
         Left            =   6330
         MaxLength       =   10
         TabIndex        =   2
         Top             =   420
         Width           =   990
      End
      Begin VB.TextBox txtstrCargo 
         Height          =   285
         Left            =   2085
         MaxLength       =   50
         TabIndex        =   5
         Top             =   1560
         Width           =   5130
      End
      Begin VB.TextBox txtstrResponsavel 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Left            =   2085
         MaxLength       =   50
         TabIndex        =   3
         Top             =   795
         Width           =   5130
      End
      Begin MSDataListLib.DataCombo dbcintLocal 
         Height          =   315
         Left            =   2085
         TabIndex        =   4
         Top             =   1170
         Width           =   5145
         _ExtentX        =   9075
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin VB.TextBox txtstrObservacao 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#.##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   1
         EndProperty
         Height          =   795
         Left            =   2085
         MaxLength       =   500
         MultiLine       =   -1  'True
         TabIndex        =   6
         Top             =   1935
         Width           =   5130
      End
      Begin VB.TextBox txtstrMatricula 
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
         Left            =   2085
         MaxLength       =   15
         TabIndex        =   0
         Top             =   420
         Width           =   1155
      End
      Begin TrueOleDBGrid70.TDBGrid tdb_Lista 
         Height          =   2595
         Left            =   150
         TabIndex        =   7
         Top             =   2955
         Width           =   7125
         _ExtentX        =   12568
         _ExtentY        =   4577
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
         Columns(1).Caption=   "Matrícula"
         Columns(1).DataField=   "strMatricula"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Nome do Responsável"
         Columns(2).DataField=   "strResponsavel"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Unidade Centro de Custo"
         Columns(3).DataField=   "UnidadeCC"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "Cargo"
         Columns(4).DataField=   "strCargo"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "Data Cadastro"
         Columns(5).DataField=   "dtmDtCadastro"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   6
         Splits(0)._UserFlags=   0
         Splits(0).MarqueeStyle=   3
         Splits(0).RecordSelectors=   0   'False
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
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
         Splits(0)._ColumnProps(8)=   "Column(1).Width=2170"
         Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=2090"
         Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(13)=   "Column(2).Width=4657"
         Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=4577"
         Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(17)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(18)=   "Column(3).Width=5450"
         Splits(0)._ColumnProps(19)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(20)=   "Column(3)._WidthInPix=5371"
         Splits(0)._ColumnProps(21)=   "Column(3)._EditAlways=0"
         Splits(0)._ColumnProps(22)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(23)=   "Column(4).Width=2725"
         Splits(0)._ColumnProps(24)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(25)=   "Column(4)._WidthInPix=2646"
         Splits(0)._ColumnProps(26)=   "Column(4)._EditAlways=0"
         Splits(0)._ColumnProps(27)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(28)=   "Column(5).Width=2196"
         Splits(0)._ColumnProps(29)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(30)=   "Column(5)._WidthInPix=2117"
         Splits(0)._ColumnProps(31)=   "Column(5)._EditAlways=0"
         Splits(0)._ColumnProps(32)=   "Column(5).Order=6"
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
         _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=28,.parent=13"
         _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
         _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=54,.parent=13"
         _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=51,.parent=14"
         _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=52,.parent=15"
         _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=53,.parent=17"
         _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=32,.parent=13"
         _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=29,.parent=14"
         _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=30,.parent=15"
         _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=31,.parent=17"
         _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=58,.parent=13"
         _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=55,.parent=14"
         _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=56,.parent=15"
         _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=57,.parent=17"
         _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=50,.parent=13"
         _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=47,.parent=14"
         _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=48,.parent=15"
         _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=49,.parent=17"
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Demissão"
         Height          =   195
         Left            =   5535
         TabIndex        =   16
         Top             =   465
         Width           =   690
      End
      Begin VB.Label lblstrCargo 
         Caption         =   "Cargo do Responsável"
         Height          =   195
         Left            =   375
         TabIndex        =   15
         Top             =   1620
         Width           =   1620
      End
      Begin VB.Label lblstrNome 
         AutoSize        =   -1  'True
         Caption         =   "Nome do Responsável"
         Height          =   195
         Left            =   375
         TabIndex        =   14
         Top             =   840
         Width           =   1620
      End
      Begin VB.Label lbl_strOservacao 
         AutoSize        =   -1  'True
         Caption         =   "Observação"
         Height          =   195
         Left            =   1125
         TabIndex        =   13
         Top             =   1995
         Width           =   870
      End
      Begin VB.Label lbl_dtmDtDataCadastro 
         AutoSize        =   -1  'True
         Caption         =   "Data Cadastro"
         Height          =   195
         Left            =   3345
         TabIndex        =   12
         Top             =   465
         Width           =   1020
      End
      Begin VB.Label lbl_intLocais 
         AutoSize        =   -1  'True
         Caption         =   "Unidade Centro de Custo"
         Height          =   195
         Left            =   210
         TabIndex        =   10
         Top             =   1230
         Width           =   1785
      End
      Begin VB.Label lbl_strMatricula 
         AutoSize        =   -1  'True
         Caption         =   "Matrícula/Código"
         Height          =   195
         Left            =   750
         TabIndex        =   9
         Top             =   465
         Width           =   1245
      End
   End
End
Attribute VB_Name = "frmCadResponsavelPatrimonio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim mblnAlterando       As Boolean
    Dim mobjAux             As Object
    Dim mblnClickOk         As Boolean
    Dim mblnPrimeiraVez     As Boolean
    
    Dim bytOrdenacao        As Byte
    Dim blnOrdenacaoAsc     As Boolean
        
    Dim strCodigo           As String
    Dim strCodigoAtual      As String
    
    Public strWhereQuery    As String 'Recebe valor para filtro no grid caso exista grupo

Private Function strQuery() As String
    Dim strSql  As String
    strSql = ""
    strSql = strSql & "SELECT A.PKId, A.strResponsavel, A.strMatricula, A.dtmDtCadastro, A.strObservacao, A.strCargo, B.strDescricao AS UnidadeCC" & _
                    " FROM " & gstrResponsavelPatrimonio & " A, " & gstrLocais & " B" & _
                    " WHERE A.intLocal=B.PKId " 'ORDER BY strMatricula"
    
    strSql = strSql & strWhereQuery
    
    Select Case bytOrdenacao
        Case Is = 1
            strSql = strSql & " ORDER BY strMatricula" & IIf(blnOrdenacaoAsc, " ASC", " DESC")
        Case Is = 2
            strSql = strSql & " ORDER BY strResponsavel" & IIf(blnOrdenacaoAsc, " ASC", " DESC")
        Case Is = 3
            strSql = strSql & " ORDER BY UnidadeCC" & IIf(blnOrdenacaoAsc, " ASC", " DESC")
        Case Is = 4
            strSql = strSql & " ORDER BY strCargo" & IIf(blnOrdenacaoAsc, " ASC", " DESC")
        Case Is = 5
            strSql = strSql & " ORDER BY dtmDtCadastro" & IIf(blnOrdenacaoAsc, " ASC", " DESC")
    End Select

    strQuery = strSql
End Function

Private Sub dbcintLocal_Click(Area As Integer)
    DropDownDataCombo dbcintLocal, Me, Area
End Sub

Private Sub dbcintLocal_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbcintLocal, Me, , KeyCode, Shift
End Sub

Private Sub Form_Activate()
    gintCodSeguranca = 16
    If UCase(MDIMenu.Tag) = "OUVIDORIA" Then
        Me.Caption = "Cadastro de Funcionários"
        tab_3dPasta.TabCaption(0) = "Funcionários"
        lblstrNome.Caption = "Nome do Funcionário"
        lblstrCargo.Caption = "Cargo do Funcionário"
        lblstrCargo.Left = 500
        lblstrNome.Left = 500
        tdb_Lista.Columns(2).Caption = "Nome do Funcionário"
    End If
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
   
    bytOrdenacao = 2: blnOrdenacaoAsc = True
    
    'VerificaListaAutomatica gstrBanco, tdb_Lista, strQuery
    VerificaObjParaAplicar mobjAux
    dbcintLocal.Tag = strQueryUnidadeCC & ";strDescricao"
    
    strWhereQuery = Space$(0)
End Sub
 
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar
    mblnPrimeiraVez = False
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
   
   blnOrdenacaoAsc = IIf(bytOrdenacao = ColIndex, Not blnOrdenacaoAsc, True)
   
   bytOrdenacao = ColIndex: MantemForm gstrRefresh

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
    With tdb_Lista
        If (Not .EOF And Not .BOF) And mblnClickOk Then
            If mblnPrimeiraVez = True Then
                mblnClickOk = False
                txtPKId.Text = .Columns("PKID").Value
                LeDaTabelaParaObj gstrResponsavelPatrimonio, Me
                gCorLinhaSelecionada tdb_Lista
                HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrDeletar
                If mobjAux Is Nothing Then
                    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar
                Else
                    HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrAplicar
                End If
                mblnAlterando = True
                strCodigoAtual = txtstrMatricula.Text
            End If
        End If
    End With
End Sub

Public Sub MantemForm(ByVal strModoOperacao As String)
    If UCase(strModoOperacao) = gstrSalvar Then
        If Not blnDadosOk Then Exit Sub
    End If
    
    If UCase(strModoOperacao) = gstrLocalizar Then
        mblnClickOk = True
        mblnPrimeiraVez = True
    End If
    
    ToolBarGeral strModoOperacao, gstrResponsavelPatrimonio, mblnAlterando, tdb_Lista, Me, mobjAux, _
                strQuery, strQuery, rptCadResponsavelPatrimonio, strQuery
End Sub

Private Sub txtdtmDemissao_Change()

End Sub

Private Sub txtdtmDtDemissao_GotFocus()
    MarcaCampo txtdtmDtDemissao
End Sub

Private Sub txtdtmDtCadastro_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "D", txtdtmDtCadastro
End Sub

Private Sub txtdtmDtCadastro_LostFocus()
    txtdtmDtCadastro = gstrDataFormatada(txtdtmDtCadastro, False)
End Sub

Private Sub txtdtmDtDemissao_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "D", txtdtmDtDemissao
End Sub

Private Sub txtdtmDtDemissao_LostFocus()
    txtdtmDtDemissao = gstrDataFormatada(txtdtmDtDemissao, False)
End Sub

Private Sub txtstrCargo_GotFocus()
    MarcaCampo txtstrCargo
End Sub

Private Sub txtstrMatricula_GotFocus()
    gstrProximoCodigo txtstrMatricula, gstrResponsavelPatrimonio, "strMatricula", gintCodSeguranca
    MarcaCampo txtstrMatricula
End Sub

Private Sub txtdtmDtCadastro_GotFocus()
    MarcaCampo txtdtmDtCadastro
End Sub

Private Sub txtstrObservacao_GotFocus()
    MarcaCampo txtstrObservacao
End Sub

'Function strQueryRelatorio() As String
'    Dim strSQL As String
'    strSQL = ""
'    strSQL = strSQL & "SELECT intBanco, strSigla, strDescricao "
'    strSQL = strSQL & "FROM " & gstrBanco
'    If mblnAlterando = True Then
'        strSQL = strSQL & " WHERE PKId = " & Val(txtPKId)
'    End If
'    strQueryRelatorio = strSQL
'End Function
'
 Private Function strQueryUnidadeCC() As String
    Dim strSql As String
    
    strSql = ""
    strSql = "SELECT PKId, strDescricao" & _
            " FROM " & gstrLocais & _
            " ORDER BY strDescricao"
            
    strQueryUnidadeCC = strSql
 End Function

Private Function blnDadosOk() As Boolean
Dim strSql          As String
Dim adoResultado    As ADODB.Recordset

blnDadosOk = False
    
If Not mblnAlterando Or (mblnAlterando And UCase$(strCodigoAtual) <> UCase$(txtstrMatricula.Text)) Then

ProximoCodigo:

        If gblnExisteCodigo(1, gstrResponsavelPatrimonio, "strMatricula", txtstrMatricula.Text) Then
            strCodigo = (gstrProximoCodigo(txtstrMatricula, gstrResponsavelPatrimonio, "strMatricula", gintCodSeguranca, , , , True))
            If MsgBox("O código informado já se encontra cadastrado. Deseja usar o código " & strCodigo & "?", vbYesNo + vbQuestion) = vbNo Then
                txtstrMatricula.SetFocus
                Exit Function
            Else
                txtstrMatricula.Text = strCodigo
                GoTo ProximoCodigo
            End If
        End If
    End If
    
     strCodigo = (gstrProximoCodigo(txtstrMatricula, gstrResponsavelPatrimonio, "strMatricula", gintCodSeguranca, , , , True))
    If Trim(txtstrMatricula.Text) = 0 Then
        If MsgBox("O campo " & lbl_strMatricula.Caption & " não pode ser nulo!. Deseja usar o código " & strCodigo & "?", vbYesNo + vbQuestion) = vbNo Then
            txtstrMatricula.SetFocus
            Exit Function
        Else
            txtstrMatricula.Text = strCodigo
            GoTo ProximoCodigo
        End If
   
    End If
       
    If Trim(txtstrMatricula) = "" Then
        ExibeMensagem "Preencha o campo Matrícula corretamente."
        txtstrMatricula.SetFocus
        Exit Function
    ElseIf Trim(txtdtmDtCadastro) = "" Then
        ExibeMensagem "Informe a data do cadastro."
        txtdtmDtCadastro.SetFocus
        Exit Function
    ElseIf Trim(txtstrResponsavel.Text) = "" Then
        ExibeMensagem "Preencha o campo Nome do Responsável corretamente."
        txtstrResponsavel.SetFocus
        Exit Function
    ElseIf dbcintLocal.MatchedWithList = False Then
        ExibeMensagem "Selecione uma Unidade Centro de Custo."
        dbcintLocal.SetFocus
        Exit Function
    End If
    
    If Len(txtdtmDtDemissao.Text) > 0 Then
        If DateDiff("D", txtdtmDtCadastro.Text, txtdtmDtDemissao.Text) <= 0 Then
            ExibeMensagem "A data de demissão não pode ser inferior do que a data de cadastro."
            txtdtmDtDemissao.SetFocus
            Exit Function
        End If
    End If
    
    strSql = "SELECT strDescricao FROM " & gstrLocais & " WHERE PKId=" & dbcintLocal.BoundText & " AND NOT dtmCancelamento IS NULL"
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 10, adoResultado) Then
        If Not adoResultado.EOF Then
            ExibeMensagem "A Unidade Centro de Custo """ & adoResultado!strDescricao & """ está cancelada, selecione outra Unidade Centro de Custo."
            If dbcintLocal.Enabled = True Then dbcintLocal.SetFocus
            Exit Function
        End If
    End If

        
    blnDadosOk = True
    
End Function


Private Sub txtstrResponsavel_GotFocus()
    MarcaCampo txtstrResponsavel
End Sub
