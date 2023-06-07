VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "TODG7.OCX"
Begin VB.Form frmCadCidade 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Municípios"
   ClientHeight    =   4815
   ClientLeft      =   4260
   ClientTop       =   4545
   ClientWidth     =   6765
   HelpContextID   =   17
   Icon            =   "CadCidade.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4815
   ScaleWidth      =   6765
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtPKId 
      Height          =   270
      Left            =   5820
      TabIndex        =   12
      Top             =   60
      Visible         =   0   'False
      Width           =   645
   End
   Begin TabDlg.SSTab tab_3dPasta 
      Height          =   4620
      Left            =   120
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   105
      Width           =   6555
      _ExtentX        =   11562
      _ExtentY        =   8149
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Municípios"
      TabPicture(0)   =   "CadCidade.frx":1042
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblstrDescricao"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblintUF"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblintCepInicial"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblintCepFinal"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblintCodigo"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtstrDescricao"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtintCepInicial"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtintCepFinal"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtintCodigo"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "dbcintUf"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "tdb_cidade"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).ControlCount=   11
      Begin TrueOleDBGrid70.TDBGrid tdb_cidade 
         Height          =   3075
         Left            =   120
         TabIndex        =   6
         Top             =   1410
         Width           =   6285
         _ExtentX        =   11086
         _ExtentY        =   5424
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
         Columns(1).DataField=   "intCodigo"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Descrição"
         Columns(2).DataField=   "strDescricao"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "UF"
         Columns(3).DataField=   "strSigla"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "CEP Inicial"
         Columns(4).DataField=   "intCEPInicial"
         Columns(4).NumberFormat=   "FormatText Event"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "CEP Final"
         Columns(5).DataField=   "intCEPFinal"
         Columns(5).NumberFormat=   "FormatText Event"
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
         Splits(0)._ColumnProps(8)=   "Column(1).Width=1296"
         Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=1217"
         Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=2"
         Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(14)=   "Column(2).Width=3360"
         Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=3281"
         Splits(0)._ColumnProps(17)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(19)=   "Column(3).Width=926"
         Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=847"
         Splits(0)._ColumnProps(22)=   "Column(3)._EditAlways=0"
         Splits(0)._ColumnProps(23)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(24)=   "Column(4).Width=2064"
         Splits(0)._ColumnProps(25)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(26)=   "Column(4)._WidthInPix=1984"
         Splits(0)._ColumnProps(27)=   "Column(4)._EditAlways=0"
         Splits(0)._ColumnProps(28)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(29)=   "Column(5).Width=2910"
         Splits(0)._ColumnProps(30)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(31)=   "Column(5)._WidthInPix=2831"
         Splits(0)._ColumnProps(32)=   "Column(5)._EditAlways=0"
         Splits(0)._ColumnProps(33)=   "Column(5).Order=6"
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
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=76,.bold=0,.fontsize=825,.italic=0"
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
         _StyleDefs(30)  =   "Splits(0).Columns(0).Style:id=58,.parent=13"
         _StyleDefs(31)  =   "Splits(0).Columns(0).HeadingStyle:id=55,.parent=14"
         _StyleDefs(32)  =   "Splits(0).Columns(0).FooterStyle:id=56,.parent=15"
         _StyleDefs(33)  =   "Splits(0).Columns(0).EditorStyle:id=57,.parent=17"
         _StyleDefs(34)  =   "Splits(0).Columns(1).Style:id=54,.parent=13,.alignment=1"
         _StyleDefs(35)  =   "Splits(0).Columns(1).HeadingStyle:id=51,.parent=14"
         _StyleDefs(36)  =   "Splits(0).Columns(1).FooterStyle:id=52,.parent=15"
         _StyleDefs(37)  =   "Splits(0).Columns(1).EditorStyle:id=53,.parent=17"
         _StyleDefs(38)  =   "Splits(0).Columns(2).Style:id=50,.parent=13"
         _StyleDefs(39)  =   "Splits(0).Columns(2).HeadingStyle:id=47,.parent=14"
         _StyleDefs(40)  =   "Splits(0).Columns(2).FooterStyle:id=48,.parent=15"
         _StyleDefs(41)  =   "Splits(0).Columns(2).EditorStyle:id=49,.parent=17"
         _StyleDefs(42)  =   "Splits(0).Columns(3).Style:id=46,.parent=13"
         _StyleDefs(43)  =   "Splits(0).Columns(3).HeadingStyle:id=43,.parent=14"
         _StyleDefs(44)  =   "Splits(0).Columns(3).FooterStyle:id=44,.parent=15"
         _StyleDefs(45)  =   "Splits(0).Columns(3).EditorStyle:id=45,.parent=17"
         _StyleDefs(46)  =   "Splits(0).Columns(4).Style:id=28,.parent=13"
         _StyleDefs(47)  =   "Splits(0).Columns(4).HeadingStyle:id=25,.parent=14"
         _StyleDefs(48)  =   "Splits(0).Columns(4).FooterStyle:id=26,.parent=15"
         _StyleDefs(49)  =   "Splits(0).Columns(4).EditorStyle:id=27,.parent=17"
         _StyleDefs(50)  =   "Splits(0).Columns(5).Style:id=32,.parent=13"
         _StyleDefs(51)  =   "Splits(0).Columns(5).HeadingStyle:id=29,.parent=14"
         _StyleDefs(52)  =   "Splits(0).Columns(5).FooterStyle:id=30,.parent=15"
         _StyleDefs(53)  =   "Splits(0).Columns(5).EditorStyle:id=31,.parent=17"
         _StyleDefs(54)  =   "Named:id=33:Normal"
         _StyleDefs(55)  =   ":id=33,.parent=0"
         _StyleDefs(56)  =   "Named:id=34:Heading"
         _StyleDefs(57)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(58)  =   ":id=34,.wraptext=-1"
         _StyleDefs(59)  =   "Named:id=35:Footing"
         _StyleDefs(60)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(61)  =   "Named:id=36:Selected"
         _StyleDefs(62)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(63)  =   "Named:id=37:Caption"
         _StyleDefs(64)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(65)  =   "Named:id=38:HighlightRow"
         _StyleDefs(66)  =   ":id=38,.parent=33,.bgcolor=&H80000014&,.fgcolor=&H80000012&"
         _StyleDefs(67)  =   "Named:id=39:EvenRow"
         _StyleDefs(68)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(69)  =   "Named:id=40:OddRow"
         _StyleDefs(70)  =   ":id=40,.parent=33"
         _StyleDefs(71)  =   "Named:id=41:RecordSelector"
         _StyleDefs(72)  =   ":id=41,.parent=34"
         _StyleDefs(73)  =   "Named:id=42:FilterBar"
         _StyleDefs(74)  =   ":id=42,.parent=33"
      End
      Begin MSDataListLib.DataCombo dbcintUf 
         Height          =   315
         Left            =   5760
         TabIndex        =   5
         Top             =   1050
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
      Begin VB.TextBox txtintCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   960
         MaxLength       =   8
         OLEDragMode     =   1  'Automatic
         TabIndex        =   1
         Top             =   375
         Width           =   1005
      End
      Begin VB.TextBox txtintCepFinal 
         Height          =   285
         Left            =   3450
         MaxLength       =   9
         TabIndex        =   4
         Top             =   1050
         Width           =   975
      End
      Begin VB.TextBox txtintCepInicial 
         Height          =   285
         Left            =   960
         MaxLength       =   9
         TabIndex        =   3
         Top             =   1050
         Width           =   975
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
         Left            =   960
         MaxLength       =   60
         TabIndex        =   2
         Top             =   720
         Width           =   5430
      End
      Begin VB.Label lblintCodigo 
         AutoSize        =   -1  'True
         Caption         =   "Código"
         Height          =   195
         Left            =   375
         TabIndex        =   11
         Top             =   420
         Width           =   495
      End
      Begin VB.Label lblintCepFinal 
         AutoSize        =   -1  'True
         Caption         =   "CEP final"
         Height          =   195
         Left            =   2700
         TabIndex        =   10
         Top             =   1110
         Width           =   645
      End
      Begin VB.Label lblintCepInicial 
         AutoSize        =   -1  'True
         Caption         =   "CEP inicial"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   1110
         Width           =   750
      End
      Begin VB.Label lblintUF 
         AutoSize        =   -1  'True
         Caption         =   "UF"
         Height          =   195
         Left            =   5490
         TabIndex        =   8
         Top             =   1110
         Width           =   210
      End
      Begin VB.Label lblstrDescricao 
         AutoSize        =   -1  'True
         Caption         =   "Descrição"
         Height          =   195
         Left            =   150
         TabIndex        =   7
         Top             =   765
         Width           =   720
      End
   End
End
Attribute VB_Name = "frmCadCidade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    
Dim mblnAlterando       As Boolean
Dim mobjAux             As Object
Dim mblnPrimeiraVez     As Boolean

Dim strCodigoAtual       As String
Dim strDescricaoAtual    As String
Dim strCodigo            As String
 
Private Function strQuery() As String
Dim strSQL  As String
   
   strSQL = ""
   
   strSQL = strSQL & "SELECT C.PKId, C.intCodigo, C.strDescricao, C.intCepInicial, "
   strSQL = strSQL & "C.intCepFinal, U.strSigla "
   strSQL = strSQL & "FROM " & gstrCidade & " C, " & gstrUF & " U "
   strSQL = strSQL & "WHERE C.intUF = U.PKId "
   strSQL = strSQL & " ORDER BY C.intCodigo"
   
   strQuery = strSQL
   
End Function

Private Sub dbcintUF_Click(Area As Integer)
   DropDownDataCombo dbcintUf, Me, Area
End Sub

Private Sub dbcintUF_KeyDown(KeyCode As Integer, Shift As Integer)
   DropDownDataCombo dbcintUf, Me, , KeyCode, Shift
End Sub

Private Sub dbcintUF_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub Form_Activate()
    gintCodSeguranca = 53
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
   
   mblnAlterando = False
   
   dbcintUf.Tag = gstrQueryUF
   LeDaTabelaParaObj gstrUF, dbcintUf, gstrQueryUF
  
   VerificaObjParaAplicar mobjAux
   
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar
    mblnPrimeiraVez = False
End Sub

Private Sub tdb_cidade_Click()
    mblnPrimeiraVez = True
     If glngQtdLinhaTDBGrid(tdb_cidade) = 1 Then
        tdb_Cidade_RowColChange 0, 0
    End If
End Sub

Private Sub tdb_cidade_DblClick()
    MantemForm gstrAplicar
End Sub

Private Sub tdb_Cidade_FilterChange()
    mblnPrimeiraVez = False
    gblnFilraCampos tdb_cidade
End Sub

Private Sub tdb_Cidade_HeadClick(ByVal ColIndex As Integer)
   gOrdenaGrid tdb_cidade, ColIndex
End Sub

Private Sub tdb_cidade_FormatText(ByVal ColIndex As Integer, Value As Variant, Bookmark As Variant)
    Value = gstrCEPFormatado(Value)
End Sub

Private Sub tdb_cidade_KeyPress(KeyAscii As Integer)
    Select Case tdb_cidade.Col
        Case 1, 4, 5
            CaracterValido KeyAscii, "N", tdb_cidade
        Case Else
            CaracterValido KeyAscii, "A", tdb_cidade
    End Select
End Sub

Private Sub tdb_Cidade_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    With tdb_cidade
        If Not .EOF And Not .BOF Then
            If mblnPrimeiraVez Then
                txtPKId.Text = .Columns("PKID").Value
                LeDaTabelaParaObj gstrCidade, Me
                gCorLinhaSelecionada tdb_cidade
                HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrDeletar
                If mobjAux Is Nothing Then
                    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar
                Else
                    HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrAplicar
                End If
                mblnAlterando = True
                
                strCodigoAtual = txtintCodigo
                strDescricaoAtual = txtstrDescricao

            End If
        End If
    End With
End Sub

Public Sub MantemForm(ByVal strModoOperacao As String)
Dim varBookMark As Variant
Dim strSQL As String

    strSQL = strQuery

    If UCase(strModoOperacao) = "SALVAR" Or UCase(strModoOperacao) = "DELETAR" Then
        mblnPrimeiraVez = False
    End If

    If UCase(strModoOperacao) = gstrSalvar Then
        If Not blnDadosOk Then Exit Sub
    End If

    If ToolBarGeral(strModoOperacao, gstrCidade, mblnAlterando, tdb_cidade, Me, mobjAux, strSQL, "PKId, strDescricao", rptMunicipio, strQueryRelatorio) Then
        If UCase(strModoOperacao) = gstrDeletar And Not tdb_cidade.EOF And Not tdb_cidade.BOF Then
            tdb_cidade.MoveFirst
        ElseIf UCase(strModoOperacao) = gstrNovo Or UCase(strModoOperacao) = gstrSalvar Then
            If txtintCodigo.Enabled Then
                txtintCodigo.SetFocus
            End If
        End If
    End If

    If Not gobjGeral Is Nothing Then HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrDeletar

End Sub

Private Sub txtintCepFinal_GotFocus()
    MarcaCampo txtintCepFinal
End Sub

Private Sub txtintCepInicial_GotFocus()
    MarcaCampo txtintCepInicial
End Sub

Private Sub txtintCepInicial_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "E", txtintCepInicial
End Sub

Private Sub txtintCepInicial_LostFocus()
    txtintCepInicial = gstrCEPFormatado(txtintCepInicial)
End Sub

Private Sub txtintCepFinal_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "E", txtintCepFinal
End Sub

Private Sub txtintCepFinal_LostFocus()
    txtintCepFinal = gstrCEPFormatado(txtintCepFinal)
End Sub

Private Sub txtintCodigo_GotFocus()
    MarcaCampo txtintCodigo
    gstrProximoCodigo txtintCodigo, gstrCidade, "intCodigo", gintCodSeguranca
End Sub

Private Sub txtintCodigo_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtintCodigo
End Sub

Private Sub txtstrDescricao_GotFocus()
    MarcaCampo txtstrDescricao
End Sub

Private Sub txtstrDescricao_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrDescricao
End Sub

Private Function blnDadosOk() As Boolean
    
    If Val(Trim(txtintCodigo)) = 0 Then
        ExibeMensagem "O código do município tem que ser digitado."
        txtintCodigo.SetFocus
        Exit Function
    End If
    
    If Len(Trim(txtstrDescricao)) = 0 Then
        ExibeMensagem "A descrição do município tem que ser digitada."
        txtstrDescricao.SetFocus
        Exit Function
    End If
    
    'If Trim(txtintCepInicial) = "" Then
    '    ExibeMensagem "O cep inicial do município tem que ser informado."
    '    txtintCepInicial.SetFocus
    '    Exit Function
    'ElseIf Trim(txtintCepFinal) = "" Then
    '    ExibeMensagem "O cep final do município tem que ser informado."
    '    txtintCepFinal.SetFocus
    '    Exit Function
    'ElseIf Val(gstrValorSemMascara(txtintCepInicial)) > Val(gstrValorSemMascara(txtintCepFinal)) Then
    '    ExibeMensagem "O cep inicial não pode ser maior que o cep final."
    '    txtintCepInicial.SetFocus
    '    Exit Function
    'End If
    
    If dbcintUf.MatchedWithList = False Then
        ExibeMensagem "A sigla UF não foi selecionada."
        dbcintUf.SetFocus
        Exit Function
    End If
    
    If mblnAlterando Then
        If txtintCepInicial.DataChanged = True Or txtintCepFinal.DataChanged = True Then
            ExibeMensagem "Atenção: A alteração da faixa de cep do município irá interferir na validação dos ceps dos logradouros."
        End If
    End If
    
    If Not mblnAlterando Or (mblnAlterando And UCase$(strCodigoAtual) <> UCase$(txtintCodigo.Text)) Then

ProximoCodigo:

        If gblnExisteCodigo(1, gstrCidade, "intCodigo", "'" & txtintCodigo.Text & "'") Then
            strCodigo = (gstrProximoCodigo(txtintCodigo, gstrCidade, "intCodigo", gintCodSeguranca, , , , True, , , , , 1))
            If MsgBox("O código informado já se encontra cadastrado. Deseja usar o código " & strCodigo & "?", vbYesNo + vbQuestion) = vbNo Then
                txtintCodigo.SetFocus
                Exit Function
            Else
                txtintCodigo.Text = strCodigo
                GoTo ProximoCodigo
            End If
        End If
    End If
    
    If Not mblnAlterando Or (mblnAlterando And UCase$(txtstrDescricao.Text) <> UCase$(strDescricaoAtual)) Then
            
        If gblnExisteCodigo(1, gstrCidade, "strDescricao", "'" & txtstrDescricao.Text & "'") Then
            ExibeMensagem "A descrição informada já se encontra cadastrada."
            txtstrDescricao.SetFocus
            Exit Function
        End If
    End If

    blnDadosOk = True
    
End Function

Private Function strQueryRelatorio() As String
Dim strSQL  As String
   
   strSQL = ""
   
   strSQL = strSQL & "SELECT C.PKId, C.intCodigo, C.strDescricao, C.intCepInicial, "
   strSQL = strSQL & "C.intCepFinal, U.strSigla "
   strSQL = strSQL & "FROM " & gstrCidade & " C, " & gstrUF & " U "
   
   If mblnAlterando = True And Val(txtPKId) <> 0 Then
      strSQL = strSQL & " WHERE C.intUF = U.PKId AND C.PKId = " & Val(txtPKId)
   Else
      strSQL = strSQL & " WHERE C.intUF = U.PKId"
   End If
   
   strSQL = strSQL & " ORDER BY C.intCodigo"
    
   strQueryRelatorio = strSQL
   
End Function
