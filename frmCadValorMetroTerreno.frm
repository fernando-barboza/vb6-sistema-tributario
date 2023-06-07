VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "TODG7.OCX"
Begin VB.Form frmCadValorMetroTerreno 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Valor Metro Terreno"
   ClientHeight    =   3735
   ClientLeft      =   2670
   ClientTop       =   3690
   ClientWidth     =   7725
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   7725
   Begin VB.TextBox txtPKId 
      Height          =   270
      Left            =   1875
      TabIndex        =   9
      Top             =   90
      Visible         =   0   'False
      Width           =   645
   End
   Begin TabDlg.SSTab tab_3dPasta 
      Height          =   3570
      Left            =   90
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   75
      Width           =   7590
      _ExtentX        =   13388
      _ExtentY        =   6297
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Valor Metro Terreno"
      TabPicture(0)   =   "frmCadValorMetroTerreno.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lbldblValor"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblintCodigo"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblintExercicio"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblMoedas"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "dbcintMoeda"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "tdb_ValorMetroTerreno"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtdblValor"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtintCodigo"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtintExercicio"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cmd_Moedas"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).ControlCount=   10
      Begin VB.CommandButton cmd_Moedas 
         Height          =   315
         Left            =   7050
         Picture         =   "frmCadValorMetroTerreno.frx":001C
         Style           =   1  'Graphical
         TabIndex        =   8
         TabStop         =   0   'False
         Tag             =   "583"
         ToolTipText     =   "Ativa Cadastro de Moedas"
         Top             =   450
         Width           =   360
      End
      Begin VB.TextBox txtintExercicio 
         Height          =   285
         Left            =   825
         MaxLength       =   4
         TabIndex        =   2
         Top             =   480
         Width           =   675
      End
      Begin VB.TextBox txtintCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2130
         MaxLength       =   5
         TabIndex        =   4
         Top             =   480
         Width           =   1005
      End
      Begin VB.TextBox txtdblValor 
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
         Left            =   3660
         MaxLength       =   50
         TabIndex        =   6
         Top             =   480
         Width           =   1530
      End
      Begin TrueOleDBGrid70.TDBGrid tdb_ValorMetroTerreno 
         Height          =   2415
         Left            =   90
         TabIndex        =   11
         Top             =   945
         Width           =   7380
         _ExtentX        =   13018
         _ExtentY        =   4260
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "PKID"
         Columns(0).DataField=   "PKID"
         Columns(0).NumberFormat=   "FormatText Event"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Exercício"
         Columns(1).DataField=   "intExercicio"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Código"
         Columns(2).DataField=   "intCodigo"
         Columns(2).NumberFormat=   "FormatText Event"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Valor"
         Columns(3).DataField=   "dblValor"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "Moeda"
         Columns(4).DataField=   "STRABREVIATURA"
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
         Splits(0)._ColumnProps(1)=   "Column(0).Width=1058"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=979"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0).AllowSizing=0"
         Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
         Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(8)=   "Column(1).Width=1773"
         Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=1693"
         Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(13)=   "Column(2).Width=1984"
         Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=1905"
         Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(17)=   "Column(2)._ColStyle=2"
         Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(19)=   "Column(3).Width=2725"
         Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=2646"
         Splits(0)._ColumnProps(22)=   "Column(3)._EditAlways=0"
         Splits(0)._ColumnProps(23)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(24)=   "Column(4).Width=2540"
         Splits(0)._ColumnProps(25)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(26)=   "Column(4)._WidthInPix=2461"
         Splits(0)._ColumnProps(27)=   "Column(4)._EditAlways=0"
         Splits(0)._ColumnProps(28)=   "Column(4).Order=5"
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
         RowDividerStyle =   3
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
         _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38,.bgcolor=&H80000014&"
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
         _StyleDefs(37)  =   "Splits(0).Columns(0).Style:id=50,.parent=13"
         _StyleDefs(38)  =   "Splits(0).Columns(0).HeadingStyle:id=47,.parent=14"
         _StyleDefs(39)  =   "Splits(0).Columns(0).FooterStyle:id=48,.parent=15"
         _StyleDefs(40)  =   "Splits(0).Columns(0).EditorStyle:id=49,.parent=17"
         _StyleDefs(41)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
         _StyleDefs(42)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
         _StyleDefs(43)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
         _StyleDefs(44)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
         _StyleDefs(45)  =   "Splits(0).Columns(2).Style:id=46,.parent=13,.alignment=1"
         _StyleDefs(46)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
         _StyleDefs(47)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
         _StyleDefs(48)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
         _StyleDefs(49)  =   "Splits(0).Columns(3).Style:id=28,.parent=13"
         _StyleDefs(50)  =   "Splits(0).Columns(3).HeadingStyle:id=25,.parent=14"
         _StyleDefs(51)  =   "Splits(0).Columns(3).FooterStyle:id=26,.parent=15"
         _StyleDefs(52)  =   "Splits(0).Columns(3).EditorStyle:id=27,.parent=17"
         _StyleDefs(53)  =   "Splits(0).Columns(4).Style:id=54,.parent=13"
         _StyleDefs(54)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14"
         _StyleDefs(55)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
         _StyleDefs(56)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
         _StyleDefs(57)  =   "Named:id=33:Normal"
         _StyleDefs(58)  =   ":id=33,.parent=0"
         _StyleDefs(59)  =   "Named:id=34:Heading"
         _StyleDefs(60)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(61)  =   ":id=34,.wraptext=-1"
         _StyleDefs(62)  =   "Named:id=35:Footing"
         _StyleDefs(63)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(64)  =   "Named:id=36:Selected"
         _StyleDefs(65)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(66)  =   "Named:id=37:Caption"
         _StyleDefs(67)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(68)  =   "Named:id=38:HighlightRow"
         _StyleDefs(69)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(70)  =   "Named:id=39:EvenRow"
         _StyleDefs(71)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(72)  =   "Named:id=40:OddRow"
         _StyleDefs(73)  =   ":id=40,.parent=33"
         _StyleDefs(74)  =   "Named:id=41:RecordSelector"
         _StyleDefs(75)  =   ":id=41,.parent=34"
         _StyleDefs(76)  =   "Named:id=42:FilterBar"
         _StyleDefs(77)  =   ":id=42,.parent=33"
      End
      Begin MSDataListLib.DataCombo dbcintMoeda 
         Height          =   315
         Left            =   5880
         TabIndex        =   7
         Top             =   450
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Text            =   ""
      End
      Begin VB.Label lblMoedas 
         AutoSize        =   -1  'True
         Caption         =   "Moedas"
         Height          =   195
         Left            =   5250
         TabIndex        =   10
         Top             =   510
         Width           =   570
      End
      Begin VB.Label lblintExercicio 
         AutoSize        =   -1  'True
         Caption         =   "Exercício"
         Height          =   195
         Left            =   90
         TabIndex        =   1
         Top             =   525
         Width           =   705
      End
      Begin VB.Label lblintCodigo 
         AutoSize        =   -1  'True
         Caption         =   "Código"
         Height          =   195
         Left            =   1545
         TabIndex        =   3
         Top             =   525
         Width           =   495
      End
      Begin VB.Label lbldblValor 
         AutoSize        =   -1  'True
         Caption         =   "Valor"
         Height          =   195
         Left            =   3210
         TabIndex        =   5
         Top             =   525
         Width           =   360
      End
   End
End
Attribute VB_Name = "frmCadValorMetroTerreno"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim mblnAlterando   As Boolean
    Dim mobjAux         As Object
    Dim mblnSelecionou  As Boolean
    Dim mblnPrimeiraVez As Boolean
    Dim bytOrdenacao    As Byte
    Dim blnOrdenacaoAsc As Boolean
    Dim strCodigoAtual  As String
    
Private Function strQuery() As String
    Dim strSQL  As String
    
    strSQL = ""
    strSQL = strSQL & " SELECT MT.PKId, MT.intCodigo, MT.intExercicio, MT.dblValor, M.STRABREVIATURA FROM "
    strSQL = strSQL & gstrValorMetroTerreno & " MT, "
    strSQL = strSQL & gstrMoedas & " M "
    strSQL = strSQL & "Where MT.intmoeda = M.pkid"
    
    Select Case bytOrdenacao
      Case Is = 1
         strSQL = strSQL & " ORDER BY intExercicio" & IIf(blnOrdenacaoAsc, " ASC", " DESC") & ", intCodigo"
      Case Is = 2
         strSQL = strSQL & " ORDER BY intExercicio" & IIf(blnOrdenacaoAsc, " ASC", " DESC") & ", intCodigo"
      Case Is = 3
         strSQL = strSQL & " ORDER BY intExercicio" & IIf(blnOrdenacaoAsc, " ASC", " DESC") & ", intCodigo"
    End Select
    
    strQuery = strSQL
End Function


Private Sub cmd_Moedas_Click()
    CarregaForm frmCadMoedas, dbcintMoeda
End Sub

Private Sub dbcintMoeda_Click(Area As Integer)
    DropDownDataCombo dbcintMoeda, Me, Area
End Sub

Private Sub dbcintMoeda_GotFocus()
    MarcaCampo dbcintMoeda
End Sub

Private Sub dbcintMoeda_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbcintMoeda, Me, , KeyCode, Shift
End Sub

Private Sub dbcintMoeda_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub Form_Activate()
    gintCodSeguranca = 1025
    VirificaGradeListView Me
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
End Sub

Private Sub Form_Deactivate()
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar
End Sub

Private Sub Form_Load()
    bytOrdenacao = 1: blnOrdenacaoAsc = False
    mblnAlterando = False
    VerificaListaAutomatica gstrValorMetroTerreno, tdb_ValorMetroTerreno, strQuery
    VerificaObjParaAplicar mobjAux
    dbcintMoeda.Tag = CarregaDataComboMoedas & ";strAbreviatura"
    
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar
    mblnSelecionou = False
    mblnPrimeiraVez = False
End Sub

Private Sub tdb_ValorMetroTerreno_Click()
    mblnPrimeiraVez = True
    If glngQtdLinhaTDBGrid(tdb_ValorMetroTerreno) = 1 Then
        tdb_ValorMetroTerreno_RowColChange 0, 0
    End If
End Sub

Private Sub tdb_ValorMetroTerreno_DblClick()
    MantemForm gstrAplicar
End Sub

Private Sub tdb_ValorMetroTerreno_FilterChange()
    mblnPrimeiraVez = False
    gblnFilraCampos tdb_ValorMetroTerreno
End Sub

Private Sub tdb_ValorMetroTerreno_FormatText(ByVal ColIndex As Integer, Value As Variant, Bookmark As Variant)
    If ColIndex = 3 Then
        Value = gstrConvVrDoSql(Value, 5)
    End If
End Sub

Private Sub tdb_ValorMetroTerreno_HeadClick(ByVal ColIndex As Integer)
    gOrdenaGrid tdb_ValorMetroTerreno, ColIndex
End Sub

Private Sub tdb_ValorMetroTerreno_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
  Dim strCarregaMoedas As String
    
    With tdb_ValorMetroTerreno
        If Not .EOF And Not .BOF Then
            mblnPrimeiraVez = True
            If mblnPrimeiraVez Then
                txtPKId.Text = .Columns("PKID").Value
                LeDaTabelaParaObj gstrValorMetroTerreno, Me
                gCorLinhaSelecionada tdb_ValorMetroTerreno
               
                HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrDeletar

                If mobjAux Is Nothing Then
                    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar
                Else
                    HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrAplicar
                End If
                mblnSelecionou = True
                mblnAlterando = True
                
                strCodigoAtual = txtintCodigo.Text
                
            End If
        End If
    End With

End Sub

Public Sub MantemForm(ByVal strModoOperacao As String)
Dim varBookmark As Variant
Dim strSQL As String

If Not tdb_ValorMetroTerreno.EOF Then
    varBookmark = tdb_ValorMetroTerreno.Bookmark
Else
    mblnAlterando = False
End If

If UCase(strModoOperacao) = gstrNovo Then
    LimpaCampos
    txtintCodigo.SetFocus
End If

If UCase(strModoOperacao) = gstrSalvar Then
    If Not blnDadosOk Then Exit Sub
End If

If UCase(strModoOperacao) = gstrSalvar Or UCase(strModoOperacao) = gstrDeletar Then
    mblnPrimeiraVez = False
End If
'ElseIf UCase(strModoOperacao) <> gstrAplicar Then
'    strSQL = strQuery
'End If

If UCase(strModoOperacao) = UCase(gstrPreencherLista) Then
    PreencherListaDeOpcoes Me.ActiveControl
End If

strSQL = strQuery

ToolBarGeral strModoOperacao, gstrValorMetroTerreno, mblnAlterando, tdb_ValorMetroTerreno, Me, mobjAux, strSQL, "PKId, dblValor", rptValorMetroTerreno, strQuery

HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrDeletar

End Sub

Private Function CarregaDataComboMoedas() As String
    Dim strSQL          As String
    
    strSQL = ""
    strSQL = strSQL & "Select Pkid, strabreviatura "
    strSQL = strSQL & "From " & gstrMoedas
    
    CarregaDataComboMoedas = strSQL
    
End Function

Private Sub txtdblValor_GotFocus()
    MarcaCampo txtdblValor
End Sub

Private Sub txtdblValor_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "V", txtdblValor
End Sub

Private Sub txtdblValor_LostFocus()
    txtdblValor = gstrConvVrDoSql(txtdblValor, 5)
End Sub

Private Sub txtintCodigo_GotFocus()
    MarcaCampo txtintCodigo
    If txtintExercicio.Text = Space$(0) Then txtintExercicio.Text = Year(Date)
    gstrProximoCodigo txtintCodigo, gstrValorMetroTerreno, "intCodigo", gintCodSeguranca, "intExercicio", txtintExercicio.Text
End Sub

Private Sub txtintCodigo_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtintCodigo
End Sub

Private Function blnDadosOk() As Boolean
Dim strCodigo As String

    If Val(Trim(txtintCodigo)) = 0 Then
        ExibeMensagem "O código tem que ser digitado."
        txtintCodigo.SetFocus
        Exit Function
    ElseIf Trim(txtintExercicio) = "" Then
        ExibeMensagem "O exercício tem que ser informado."
        txtintExercicio.SetFocus
        Exit Function
    ElseIf Trim(txtdblValor) = "" Then
        ExibeMensagem "O valor tem que ser informado."
        txtdblValor.SetFocus
        Exit Function
    ElseIf dbcintMoeda.BoundText = "" Then
        ExibeMensagem "A moeda tem que ser informada."
        dbcintMoeda.SetFocus
        Exit Function
    End If
    
    
    If Not mblnAlterando Or (mblnAlterando And UCase$(strCodigoAtual) <> UCase$(txtintCodigo.Text)) Then
        If gblnExisteCodigo(2, gstrValorMetroTerreno, "intCodigo", "'" & txtintCodigo.Text & "'", "intExercicio", txtintExercicio.Text) Then
            strCodigo = (gstrProximoCodigo(txtintCodigo, gstrValorMetroTerreno, "intCodigo", gintCodSeguranca, , , , True))
            ExibeMensagem "O código informado já se encontra cadastrado."
            txtintCodigo.SetFocus
            Exit Function
        End If
    End If
    
    
    blnDadosOk = True
    
End Function

Private Sub txtintExercicio_GotFocus()
    MarcaCampo txtintExercicio
End Sub

Private Sub txtintExercicio_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtintExercicio
End Sub

Private Sub LimpaCampos()
    
    txtintExercicio.Text = Year(Date)
    txtintCodigo.Text = Space$(0)
    txtdblValor.Text = Space$(0)
    dbcintMoeda.Text = Space$(0)
    
End Sub
