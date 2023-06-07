VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "TODG7.OCX"
Begin VB.Form frmParametrosDividaAtiva 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Parametros para Divida Ativa"
   ClientHeight    =   4905
   ClientLeft      =   2175
   ClientTop       =   3105
   ClientWidth     =   7515
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4905
   ScaleWidth      =   7515
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtPKID 
      Height          =   285
      Left            =   2880
      TabIndex        =   17
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
   Begin TabDlg.SSTab tab_3dPasta 
      Height          =   4560
      Left            =   120
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   7260
      _ExtentX        =   12806
      _ExtentY        =   8043
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Parâmetros para Dívida Ativa"
      TabPicture(0)   =   "frmParametrosDividaAtiva.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fra_CompReceita"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin VB.Frame fra_CompReceita 
         Caption         =   "Ultimos Parâmetros Utilizados"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3660
         Left            =   240
         TabIndex        =   1
         Top             =   720
         Width           =   6795
         Begin VB.TextBox txtintFolhaPorLivro 
            Height          =   285
            Left            =   1680
            MaxLength       =   4
            TabIndex        =   12
            Top             =   1080
            Width           =   735
         End
         Begin VB.TextBox txtintCertidaoPorFolha 
            Height          =   285
            Left            =   4680
            MaxLength       =   2
            TabIndex        =   14
            Top             =   1080
            Width           =   735
         End
         Begin VB.TextBox txtintqtdcertidaoultfolha 
            Height          =   285
            Left            =   1680
            MaxLength       =   1
            TabIndex        =   16
            Top             =   1440
            Width           =   735
         End
         Begin VB.TextBox txtintCertidao 
            Height          =   285
            Left            =   1680
            TabIndex        =   6
            Top             =   720
            Width           =   975
         End
         Begin VB.TextBox txtintLivro 
            Height          =   285
            Left            =   3240
            MaxLength       =   5
            TabIndex        =   8
            Top             =   720
            Width           =   735
         End
         Begin VB.TextBox txtintFolha 
            Height          =   285
            Left            =   4680
            MaxLength       =   3
            TabIndex        =   10
            Top             =   720
            Width           =   735
         End
         Begin VB.CommandButton cmd_Composicao 
            Height          =   300
            Left            =   5475
            Picture         =   "frmParametrosDividaAtiva.frx":001C
            Style           =   1  'Graphical
            TabIndex        =   4
            TabStop         =   0   'False
            Tag             =   "590"
            ToolTipText     =   "Ativa Cadastro de Composição da Receita"
            Top             =   360
            Width           =   360
         End
         Begin MSDataListLib.DataCombo dbcintComposicaoDaReceita 
            Height          =   315
            Left            =   1680
            TabIndex        =   3
            Top             =   360
            Width           =   3735
            _ExtentX        =   6588
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin TrueOleDBGrid70.TDBGrid tdb_Parametros 
            Height          =   1455
            Left            =   120
            TabIndex        =   18
            Top             =   2040
            Width           =   6525
            _ExtentX        =   11509
            _ExtentY        =   2566
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
            Columns(1).Caption=   "Composição da Receita"
            Columns(1).DataField=   "strComposicaoDaReceita"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Certidão"
            Columns(2).DataField=   "intCertidao"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Livro"
            Columns(3).DataField=   "intLivro"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "Folha"
            Columns(4).DataField=   "intFolha"
            Columns(4).ConvertEmptyCell=   1
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "intFolhaPorlivro"
            Columns(5).DataField=   "intCertidaoporfolha"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(6)._VlistStyle=   0
            Columns(6)._MaxComboItems=   5
            Columns(6).DataField=   ""
            Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   7
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
            Splits(0)._ColumnProps(0)=   "Columns.Count=7"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0).AllowSizing=0"
            Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
            Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(8)=   "Column(1).Width=6535"
            Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=6456"
            Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(13)=   "Column(2).Width=1879"
            Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=1799"
            Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(17)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(18)=   "Column(3).Width=1349"
            Splits(0)._ColumnProps(19)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(20)=   "Column(3)._WidthInPix=1270"
            Splits(0)._ColumnProps(21)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(22)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(23)=   "Column(4).Width=1191"
            Splits(0)._ColumnProps(24)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(25)=   "Column(4)._WidthInPix=1111"
            Splits(0)._ColumnProps(26)=   "Column(4)._EditAlways=0"
            Splits(0)._ColumnProps(27)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(28)=   "Column(5).Width=2725"
            Splits(0)._ColumnProps(29)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(30)=   "Column(5)._WidthInPix=2646"
            Splits(0)._ColumnProps(31)=   "Column(5)._EditAlways=0"
            Splits(0)._ColumnProps(32)=   "Column(5).Order=6"
            Splits(0)._ColumnProps(33)=   "Column(6).Width=2725"
            Splits(0)._ColumnProps(34)=   "Column(6).DividerColor=0"
            Splits(0)._ColumnProps(35)=   "Column(6)._WidthInPix=2646"
            Splits(0)._ColumnProps(36)=   "Column(6)._EditAlways=0"
            Splits(0)._ColumnProps(37)=   "Column(6).Order=7"
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
            _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=38,.bold=0,.fontsize=825,.italic=0"
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
            _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=38,.bgcolor=&H8000000D&"
            _StyleDefs(18)  =   ":id=6,.fgcolor=&H8000000E&"
            _StyleDefs(19)  =   "EditorStyle:id=7,.parent=1"
            _StyleDefs(20)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
            _StyleDefs(21)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
            _StyleDefs(22)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
            _StyleDefs(23)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
            _StyleDefs(24)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
            _StyleDefs(25)  =   "Splits(0).Style:id=13,.parent=1"
            _StyleDefs(26)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
            _StyleDefs(27)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
            _StyleDefs(28)  =   "Splits(0).FooterStyle:id=15,.parent=3"
            _StyleDefs(29)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
            _StyleDefs(30)  =   "Splits(0).SelectedStyle:id=18,.parent=6,.namedParent=38,.bgcolor=&H8000000D&"
            _StyleDefs(31)  =   ":id=18,.fgcolor=&H8000000E&"
            _StyleDefs(32)  =   "Splits(0).EditorStyle:id=17,.parent=7"
            _StyleDefs(33)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8,.namedParent=38,.bgcolor=&H8000000D&"
            _StyleDefs(34)  =   ":id=19,.fgcolor=&H8000000E&"
            _StyleDefs(35)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
            _StyleDefs(36)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
            _StyleDefs(37)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
            _StyleDefs(38)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
            _StyleDefs(39)  =   "Splits(0).Columns(0).Style:id=46,.parent=13"
            _StyleDefs(40)  =   "Splits(0).Columns(0).HeadingStyle:id=43,.parent=14"
            _StyleDefs(41)  =   "Splits(0).Columns(0).FooterStyle:id=44,.parent=15"
            _StyleDefs(42)  =   "Splits(0).Columns(0).EditorStyle:id=45,.parent=17"
            _StyleDefs(43)  =   "Splits(0).Columns(1).Style:id=28,.parent=13"
            _StyleDefs(44)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14"
            _StyleDefs(45)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
            _StyleDefs(46)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
            _StyleDefs(47)  =   "Splits(0).Columns(2).Style:id=32,.parent=13"
            _StyleDefs(48)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
            _StyleDefs(49)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
            _StyleDefs(50)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
            _StyleDefs(51)  =   "Splits(0).Columns(3).Style:id=50,.parent=13"
            _StyleDefs(52)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
            _StyleDefs(53)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
            _StyleDefs(54)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
            _StyleDefs(55)  =   "Splits(0).Columns(4).Style:id=54,.parent=13"
            _StyleDefs(56)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14"
            _StyleDefs(57)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
            _StyleDefs(58)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
            _StyleDefs(59)  =   "Splits(0).Columns(5).Style:id=58,.parent=13"
            _StyleDefs(60)  =   "Splits(0).Columns(5).HeadingStyle:id=55,.parent=14"
            _StyleDefs(61)  =   "Splits(0).Columns(5).FooterStyle:id=56,.parent=15"
            _StyleDefs(62)  =   "Splits(0).Columns(5).EditorStyle:id=57,.parent=17"
            _StyleDefs(63)  =   "Splits(0).Columns(6).Style:id=62,.parent=13"
            _StyleDefs(64)  =   "Splits(0).Columns(6).HeadingStyle:id=59,.parent=14"
            _StyleDefs(65)  =   "Splits(0).Columns(6).FooterStyle:id=60,.parent=15"
            _StyleDefs(66)  =   "Splits(0).Columns(6).EditorStyle:id=61,.parent=17"
            _StyleDefs(67)  =   "Named:id=33:Normal"
            _StyleDefs(68)  =   ":id=33,.parent=0"
            _StyleDefs(69)  =   "Named:id=34:Heading"
            _StyleDefs(70)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(71)  =   ":id=34,.wraptext=-1"
            _StyleDefs(72)  =   "Named:id=35:Footing"
            _StyleDefs(73)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(74)  =   "Named:id=36:Selected"
            _StyleDefs(75)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(76)  =   "Named:id=37:Caption"
            _StyleDefs(77)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(78)  =   "Named:id=38:HighlightRow"
            _StyleDefs(79)  =   ":id=38,.parent=33,.bgcolor=&H80000014&,.fgcolor=&H80000012&"
            _StyleDefs(80)  =   "Named:id=39:EvenRow"
            _StyleDefs(81)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(82)  =   "Named:id=40:OddRow"
            _StyleDefs(83)  =   ":id=40,.parent=33"
            _StyleDefs(84)  =   "Named:id=41:RecordSelector"
            _StyleDefs(85)  =   ":id=41,.parent=34"
            _StyleDefs(86)  =   "Named:id=42:FilterBar"
            _StyleDefs(87)  =   ":id=42,.parent=33"
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Comp. da Receita:"
            Height          =   195
            Left            =   270
            TabIndex        =   2
            Top             =   420
            Width           =   1320
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Folhas p/ Livro:"
            Height          =   195
            Left            =   480
            TabIndex        =   11
            Top             =   1125
            Width           =   1110
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Certidões por Folha:"
            Height          =   195
            Left            =   3240
            TabIndex        =   13
            Top             =   1125
            Width           =   1410
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Qtd Certidão - Ult. Fl:"
            Height          =   195
            Left            =   120
            TabIndex        =   15
            Top             =   1485
            Width           =   1470
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Certidão:"
            Height          =   195
            Left            =   960
            TabIndex        =   5
            Top             =   765
            Width           =   630
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Livro:"
            Height          =   195
            Left            =   2760
            TabIndex        =   7
            Top             =   765
            Width           =   390
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Folha:"
            Height          =   195
            Left            =   4200
            TabIndex        =   9
            Top             =   765
            Width           =   435
         End
      End
   End
End
Attribute VB_Name = "frmParametrosDividaAtiva"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private blnPrimeiraVez As Boolean
Private blnAlterando    As Boolean

Private Function strQueryComposicaoReceita()
    
    Dim strSQL As String
    
    strSQL = ""
    strSQL = strSQL & "SELECT PKId, Ltrim(Rtrim(intCodigo))  " & strCONCAT & "' - '" & strCONCAT & " Ltrim(Rtrim(strDescricao))  strComposicaoDaReceita "
    strSQL = strSQL & "FROM " & gstrComposicaoDaReceita & " "
    strSQL = strSQL & "WHERE bytDividaAtiva = 1 AND "
    strSQL = strSQL & "intUtilizacao <> 3 "
    strSQL = strSQL & "ORDER BY intCodigo"
    
    strQueryComposicaoReceita = strSQL
    
End Function

Private Sub cmd_Composicao_Click()
    ChamaFormCadastro frmCadComposicaoDaReceita, dbcintComposicaoDaReceita, strQueryComposicaoReceita
End Sub

Private Sub dbcintComposicaoDaReceita_GotFocus()
    MarcaCampo dbcintComposicaoDaReceita
End Sub

Private Sub Form_Load()
    dbcintComposicaoDaReceita.Tag = strQueryComposicaoReceita & ";strDescricao"
    TrocaCorObjeto txtintqtdcertidaoultfolha, True
    blnPrimeiraVez = False
    blnAlterando = False
End Sub

Public Sub MantemForm(ByVal strModoOperacao As String)
    
    Select Case UCase(strModoOperacao)
        Case UCase(gstrNovo)
            LimpaControles
        Case UCase(gstrPreencherLista)
            PreencherListaDeOpcoes Me.ActiveControl
        Case UCase(gstrSalvar)
            If blnDadosOK Then
                If gblnExclusaoGravacaoOk(IIf(blnAlterando, "A", "I")) Then
                    GravaParametros
                    LimpaControles
                    MantemForm gstrLocalizar
                End If
            End If
        Case UCase(gstrDeletar)
            If Not tdb_Parametros.EOF Then
                If gblnExclusaoGravacaoOk("E") Then
                    ExcluiParametros
                    LimpaControles
                    MantemForm gstrLocalizar
                End If
            End If
        Case UCase(gstrLocalizar)
            ToolBarGeral gstrLocalizar, gstrParametroDividaAtiva, False, tdb_Parametros, Me, , strQueryLocalizar
    End Select
    
End Sub

Private Function blnDadosOK() As Boolean

    blnDadosOK = False
   
    If Trim(txtintCertidao) = "" And Trim(txtintFolha) = "" And Trim(txtintLivro) = "" And Trim(txtintFolhaPorLivro) = "" And Trim(txtintCertidaoPorFolha) = "" Then
        MsgBox "Preencha os campos antes de salvar!", vbCritical + vbOKOnly
        Exit Function
    End If
           
    If Trim(txtintFolha.Text) <> "" And Trim(txtintLivro.Text) <> "" And Trim(txtintFolhaPorLivro.Text) = "" Then
        MsgBox "O campo Folhas por livro não pode estar em branco.", vbCritical + vbOKOnly
        txtintFolhaPorLivro.SetFocus
        Exit Function
    End If
    
    If Trim(txtintFolha.Text) <> "" And Trim(txtintLivro) <> "" And Trim(txtintCertidaoPorFolha.Text) = "" Then
        MsgBox "O campo Certidão por Folha não pode estar em branco.", vbCritical + vbOKOnly
        txtintCertidaoPorFolha.SetFocus
        Exit Function
    End If
    
    If dbcintComposicaoDaReceita.BoundText = "" Then
        If blnExisteParametroGenerico Then
            MsgBox "Já existe um Parametro Genérico gravado!", vbCritical + vbOKOnly
            Exit Function
        End If
    End If
    
    blnDadosOK = True
    
End Function

' Verifica se já existe um parametro genérico na tblParametroDividaAtiva

Private Function blnExisteParametroGenerico() As Boolean

    blnExisteParametroGenerico = False

    Dim strSQL As String
    Dim adoRec As New ADODB.Recordset
    
    strSQL = " SELECT PKID" & _
             " FROM " & gstrParametroDividaAtiva & _
             " WHERE intComposicaoDaReceita IS NULL"
             
    Set gobjBanco = New clsBanco
    
    If gobjBanco.CriaADO(strSQL, 10, adoRec) Then
        If adoRec.EOF Then
            Exit Function
        End If
    End If

    blnExisteParametroGenerico = True
            
End Function

Private Function strQueryLocalizar() As String

    Dim strSQL As String

    strSQL = "SELECT PDA.PKID, cr.strdescricao strComposicaoDaReceita," & _
                   " PDA.intComposicaoDaReceita, " & _
                   " PDA.Intcertidao," & _
                   " PDA.intlivro," & _
                   " PDA.intfolha, " & _
                   " PDA.intFolhaPorLivro, " & _
                   " PDA.intcertidaoporfolha, " & _
                   " PDA.intqtdcertidaoultfolha " & _
             " FROM " & gstrComposicaoDaReceita & " CR, " & _
                       gstrParametroDividaAtiva & " PDA " & _
             " WHERE PDA.intComposicaoDaReceita " & strOUTJSQLServer & "= CR.Pkid " & strOUTJOracle
             
    If blnPrimeiraVez Then
        strSQL = strSQL & " AND PDA.pkid = " & tdb_Parametros.Columns("pkid").Value
    End If
                
    strQueryLocalizar = strSQL

End Function

Private Sub tdb_Parametros_Click()
    blnPrimeiraVez = True
    blnAlterando = True
End Sub

Private Sub tdb_Parametros_HeadClick(ByVal ColIndex As Integer)
    gOrdenaGrid tdb_Parametros, ColIndex
End Sub

Private Sub tdb_Parametros_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If Not tdb_Parametros.EOF Then
        If blnPrimeiraVez Then
            PreencheFormulario strQueryLocalizar
            HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrDeletar
        End If
    End If
End Sub

Private Sub LimpaControles()

    Set dbcintComposicaoDaReceita.RowSource = Nothing
    txtPKID.Text = ""
    dbcintComposicaoDaReceita.Text = ""
    txtintCertidao.Text = ""
    txtintLivro.Text = ""
    txtintFolha.Text = ""
    txtintFolhaPorLivro.Text = ""
    txtintCertidaoPorFolha.Text = ""
    txtintqtdcertidaoultfolha.Text = ""
    dbcintComposicaoDaReceita.SetFocus
    blnPrimeiraVez = False
    blnAlterando = False

End Sub

Private Sub GravaParametros()

    Dim strSQL As String
    
    If Not blnAlterando Then
        strSQL = "INSERT INTO " & gstrParametroDividaAtiva & "(" & _
                                 "intComposicaoDaReceita, " & _
                                 "Intcertidao, " & _
                                 "intlivro, " & _
                                 "intfolha, " & _
                                 "intFolhaPorLivro, " & _
                                 "intcertidaoporfolha, " & _
                                 "intqtdcertidaoultfolha," & _
                                 "bitmanutencao," & _
                                 "dtmdtAtualizacao, " & _
                                 "lngCodUsr) " & _
                 "VALUES (" & _
                                 IIf(dbcintComposicaoDaReceita.BoundText = "", "NULL", dbcintComposicaoDaReceita.BoundText) & ", " & _
                                 IIf(Trim(txtintCertidao) = "", "NULL", Trim(txtintCertidao)) & ", " & _
                                 IIf(Trim(txtintLivro) = "", "NULL", Trim(txtintLivro)) & ", " & _
                                 IIf(Trim(txtintFolha) = "", "NULL", Trim(txtintFolha)) & ", " & _
                                 IIf(Trim(txtintFolhaPorLivro) = "", "NULL", Trim(txtintFolhaPorLivro)) & ", " & _
                                 IIf(Trim(txtintCertidaoPorFolha) = "", "NULL", Trim(txtintCertidaoPorFolha)) & ", " & _
                                 "0, " & _
                                 "0, " & _
                                 strGETDATE & ", " & _
                                 glngCodUsr & _
                                 ")"
    Else
        strSQL = "UPDATE " & gstrParametroDividaAtiva & _
                 " SET intcomposicaodareceita = " & IIf(dbcintComposicaoDaReceita.BoundText = "", "NULL", dbcintComposicaoDaReceita.BoundText) & ", " & _
                     "Intcertidao  = " & IIf(Trim(txtintCertidao) = "", "NULL", Trim(txtintCertidao)) & ", " & _
                     "intlivro = " & IIf(Trim(txtintLivro) = "", "NULL", Trim(txtintLivro)) & ", " & _
                     "intfolha = " & IIf(Trim(txtintFolha) = "", "NULL", Trim(txtintFolha)) & ", " & _
                     "intFolhaporlivro = " & IIf(Trim(txtintFolhaPorLivro) = "", "NULL", Trim(txtintFolhaPorLivro)) & ", " & _
                     "intcertidaoporfolha = " & IIf(Trim(txtintCertidaoPorFolha) = "", "NULL", Trim(txtintCertidaoPorFolha)) & ", " & _
                     "intqtdcertidaoultfolha = 0, " & _
                     "bitManutencao = 0, " & _
                     "dtmdtAtualizacao = " & strGETDATE & ", " & _
                     "lngCodUsr = " & glngCodUsr & _
                 " WHERE Pkid = " & tdb_Parametros.Columns("pkid").Value
    End If
    
    Set gobjBanco = New clsBanco
    gobjBanco.Execute strSQL

End Sub

Private Sub ExcluiParametros()

    Dim strSQL As String

    strSQL = "DELETE FROM " & gstrParametroDividaAtiva & _
             " WHERE Pkid = " & tdb_Parametros.Columns("PKID").Value
    
    Set gobjBanco = New clsBanco
    
    gobjBanco.Execute strSQL
    
    Set gobjBanco = Nothing
                                 
End Sub

Private Sub txtintCertidao_GotFocus()
    MarcaCampo txtintCertidao
End Sub

Private Sub txtintCertidaoPorFolha_GotFocus()
MarcaCampo txtintCertidaoPorFolha
End Sub

Private Sub txtintFolha_GotFocus()
    MarcaCampo txtintFolha
End Sub

Private Sub txtintFolhaPorLivro_GotFocus()
    MarcaCampo txtintFolhaPorLivro
End Sub

Private Sub txtintLivro_GotFocus()
    MarcaCampo txtintLivro
End Sub
' Função para preencher os campos do formulário manualmente
' ** Não foi possível utilizar a função genérica LeDaTabelaParaObjeto **
Private Sub PreencheFormulario(strQuery As String)

    Dim adoRec As New ADODB.Recordset

    Set gobjBanco = New clsBanco
    
    If gobjBanco.CriaADO(strQuery, 10, adoRec) Then
        With adoRec
            If Not .EOF Then
                
                dbcintComposicaoDaReceita.Text = ""
                txtintCertidao.Text = ""
                txtintLivro.Text = ""
                txtintFolha.Text = ""
                txtintFolhaPorLivro.Text = ""
                txtintCertidaoPorFolha.Text = ""
                txtintqtdcertidaoultfolha.Text = ""
                
                If IsNull(!intComposicaoDaReceita) = False Then PreencherListaDeOpcoes dbcintComposicaoDaReceita, !intComposicaoDaReceita
                txtintCertidao = IIf(IsNull(!intCertidao), "", !intCertidao)
                txtintLivro = IIf(IsNull(!intLivro), "", !intLivro)
                txtintFolha = IIf(IsNull(!intFolha), "", !intFolha)
                txtintFolhaPorLivro = IIf(IsNull(!intFolhaPorLivro), "", !intFolhaPorLivro)
                txtintCertidaoPorFolha = IIf(IsNull(!intCertidaoPorFolha), "", !intCertidaoPorFolha)
                txtintqtdcertidaoultfolha = IIf(IsNull(!intQtdCertidaoUltFolha), "", !intQtdCertidaoUltFolha)
                
            End If
        End With
    End If

End Sub
