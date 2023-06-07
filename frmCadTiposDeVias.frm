VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "TODG7.OCX"
Begin VB.Form frmCadTiposDeVias 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tipos de Vias"
   ClientHeight    =   4695
   ClientLeft      =   3585
   ClientTop       =   2325
   ClientWidth     =   5700
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4695
   ScaleWidth      =   5700
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab tab_3dPasta 
      Height          =   4530
      Left            =   90
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   75
      Width           =   5520
      _ExtentX        =   9737
      _ExtentY        =   7990
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Tipos de Vias"
      TabPicture(0)   =   "frmCadTiposDeVias.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lbl_Codigo"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblstrDescricao"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblstrSigla"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "tdb_TiposDeVias"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtstrDescricao"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtstrSigla"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtPkid"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtintCodigo"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      Begin VB.TextBox txtintCodigo 
         Height          =   300
         Left            =   1020
         MaxLength       =   10
         TabIndex        =   0
         Top             =   450
         Width           =   870
      End
      Begin VB.TextBox txtPkid 
         Height          =   285
         Left            =   1455
         TabIndex        =   8
         Top             =   30
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.TextBox txtstrSigla 
         Height          =   285
         Left            =   1020
         MaxLength       =   3
         TabIndex        =   2
         Top             =   1170
         Width           =   1140
      End
      Begin VB.TextBox txtstrDescricao 
         Height          =   285
         Left            =   1020
         MaxLength       =   40
         TabIndex        =   1
         Top             =   810
         Width           =   4365
      End
      Begin TrueOleDBGrid70.TDBGrid tdb_TiposDeVias 
         Height          =   2865
         Left            =   120
         TabIndex        =   3
         Top             =   1530
         Width           =   5265
         _ExtentX        =   9287
         _ExtentY        =   5054
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Pkid"
         Columns(0).DataField=   "Pkid"
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
         Columns(3).Caption=   "Sigla"
         Columns(3).DataField=   "strSigla"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   4
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
         Splits(0)._ColumnProps(0)=   "Columns.Count=4"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=1455"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1376"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0).Visible=0"
         Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(7)=   "Column(1).Width=1482"
         Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=1402"
         Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(12)=   "Column(2).Width=5662"
         Splits(0)._ColumnProps(13)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(14)=   "Column(2)._WidthInPix=5583"
         Splits(0)._ColumnProps(15)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(16)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(17)=   "Column(3).Width=1640"
         Splits(0)._ColumnProps(18)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(19)=   "Column(3)._WidthInPix=1561"
         Splits(0)._ColumnProps(20)=   "Column(3)._EditAlways=0"
         Splits(0)._ColumnProps(21)=   "Column(3).Order=4"
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
         TabAction       =   1
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
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=121,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=38,.bgcolor=&H8000000E&,.bold=0,.fontsize=825"
         _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=0"
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
         _StyleDefs(20)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38,.bgcolor=&H8000000D&"
         _StyleDefs(21)  =   ":id=8,.fgcolor=&H8000000E&"
         _StyleDefs(22)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
         _StyleDefs(23)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
         _StyleDefs(24)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(25)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
         _StyleDefs(26)  =   "Splits(0).Style:id=13,.parent=1"
         _StyleDefs(27)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
         _StyleDefs(28)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
         _StyleDefs(29)  =   "Splits(0).FooterStyle:id=15,.parent=3"
         _StyleDefs(30)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
         _StyleDefs(31)  =   "Splits(0).SelectedStyle:id=18,.parent=6,.namedParent=38,.bgcolor=&H8000000D&"
         _StyleDefs(32)  =   ":id=18,.fgcolor=&H8000000E&"
         _StyleDefs(33)  =   "Splits(0).EditorStyle:id=17,.parent=7"
         _StyleDefs(34)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
         _StyleDefs(35)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
         _StyleDefs(36)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
         _StyleDefs(37)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(38)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(39)  =   "Splits(0).Columns(0).Style:id=50,.parent=13"
         _StyleDefs(40)  =   "Splits(0).Columns(0).HeadingStyle:id=47,.parent=14"
         _StyleDefs(41)  =   "Splits(0).Columns(0).FooterStyle:id=48,.parent=15"
         _StyleDefs(42)  =   "Splits(0).Columns(0).EditorStyle:id=49,.parent=17"
         _StyleDefs(43)  =   "Splits(0).Columns(1).Style:id=28,.parent=13"
         _StyleDefs(44)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14"
         _StyleDefs(45)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
         _StyleDefs(46)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
         _StyleDefs(47)  =   "Splits(0).Columns(2).Style:id=32,.parent=13"
         _StyleDefs(48)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
         _StyleDefs(49)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
         _StyleDefs(50)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
         _StyleDefs(51)  =   "Splits(0).Columns(3).Style:id=46,.parent=13"
         _StyleDefs(52)  =   "Splits(0).Columns(3).HeadingStyle:id=43,.parent=14"
         _StyleDefs(53)  =   "Splits(0).Columns(3).FooterStyle:id=44,.parent=15"
         _StyleDefs(54)  =   "Splits(0).Columns(3).EditorStyle:id=45,.parent=17"
         _StyleDefs(55)  =   "Named:id=33:Normal"
         _StyleDefs(56)  =   ":id=33,.parent=0"
         _StyleDefs(57)  =   "Named:id=34:Heading"
         _StyleDefs(58)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(59)  =   ":id=34,.wraptext=-1"
         _StyleDefs(60)  =   "Named:id=35:Footing"
         _StyleDefs(61)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(62)  =   "Named:id=36:Selected"
         _StyleDefs(63)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(64)  =   "Named:id=37:Caption"
         _StyleDefs(65)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(66)  =   "Named:id=38:HighlightRow"
         _StyleDefs(67)  =   ":id=38,.parent=33,.bgcolor=&H80000014&,.fgcolor=&H80000012&"
         _StyleDefs(68)  =   "Named:id=39:EvenRow"
         _StyleDefs(69)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(70)  =   "Named:id=40:OddRow"
         _StyleDefs(71)  =   ":id=40,.parent=33"
         _StyleDefs(72)  =   "Named:id=41:RecordSelector"
         _StyleDefs(73)  =   ":id=41,.parent=34"
         _StyleDefs(74)  =   "Named:id=42:FilterBar"
         _StyleDefs(75)  =   ":id=42,.parent=33"
      End
      Begin VB.Label lblstrSigla 
         AutoSize        =   -1  'True
         Caption         =   "Sigla"
         Height          =   195
         Left            =   615
         TabIndex        =   7
         Top             =   1215
         Width           =   345
      End
      Begin VB.Label lblstrDescricao 
         AutoSize        =   -1  'True
         Caption         =   "Descrição"
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   855
         Width           =   720
      End
      Begin VB.Label lbl_Codigo 
         AutoSize        =   -1  'True
         Caption         =   "Código"
         Height          =   195
         Left            =   465
         TabIndex        =   5
         Top             =   495
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmCadTiposDeVias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim blnPrimeiraVez  As Boolean
    Dim mblnPrimeiraVez As Boolean
    Dim mobjAux         As Object
    Dim blnAlterando    As Boolean
    Dim mblnClickOk     As Boolean
    Dim strCodigoAtual     As String
    Dim strDescricaoAtual  As String
    Dim strSiglaAtual      As String

Private Sub Form_Activate()
    
    gintCodSeguranca = 1058
    HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrDeletar
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
    VerificaObjParaAplicar mobjAux
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar
    blnPrimeiraVez = False
    blnAlterando = False
End Sub

Private Sub tdb_TiposDeVias_Click()
    blnPrimeiraVez = True
    If glngQtdLinhaTDBGrid(tdb_TiposDeVias) = 1 Then
        tdb_TiposDeVias_RowColChange 0, 0
    End If
End Sub

Private Function strQueryTiposDeVias() As String

Dim strSql As String

    strSql = "SELECT Pkid Pkid, "
    strSql = strSql & "intCodigo intCodigo, "
    strSql = strSql & "strDescricao strDescricao, "
    strSql = strSql & "strSigla strSigla "
    strSql = strSql & "FROM "
    strSql = strSql & gstrTiposDeVias & " "
'strSql = strSql & " ORDER BY "
'strSql = strSql & "intcodigo"

    strSql = strSql & "ORDER BY intCodigo "

    strQueryTiposDeVias = strSql

End Function

Public Sub MantemForm(ByVal strModoOperacao As String)
Dim strSql As String

strSql = strQueryTiposDeVias
If strModoOperacao = UCase(gstrImprimir) Then
    ToolBarGeral strModoOperacao, gstrFormulaDeCalculo, blnAlterando, tdb_TiposDeVias, Me, mobjAux, strSql, , rpttiposdevias, strQueryTiposDeVias
    Exit Sub
End If


Select Case UCase(strModoOperacao)
    Case Is = UCase(gstrLocalizar)
        ToolBarGeral strModoOperacao, gstrTiposDeVias, False, tdb_TiposDeVias, Me, mobjAux, strQueryTiposDeVias
    Case Is = UCase(gstrNovo)
        TrocaCorObjeto txtintCodigo, False
        LimpaObjeto Me
        blnPrimeiraVez = False
        blnAlterando = False
    Case Is = UCase(gstrFechar)
        Unload Me
    Case Is = UCase(gstrSalvar)
        If blnDadosOk Then
            ToolBarGeral strModoOperacao, gstrTiposDeVias, blnAlterando, tdb_TiposDeVias, _
                     Me, mobjAux, strQueryTiposDeVias
        End If
    Case Else
        ToolBarGeral strModoOperacao, gstrTiposDeVias, blnAlterando, tdb_TiposDeVias, _
                    Me, mobjAux, strQueryTiposDeVias, strQueryAplicar
        blnPrimeiraVez = False
        blnAlterando = False
        TrocaCorObjeto txtintCodigo, False
        
End Select
End Sub


Private Sub tdb_TiposDeVias_DblClick()
    MantemForm gstrAplicar
End Sub

Private Sub tdb_TiposDeVias_FilterChange()
    mblnPrimeiraVez = False
    gblnFilraCampos tdb_TiposDeVias
End Sub

Private Sub tdb_TiposDeVias_HeadClick(ByVal ColIndex As Integer)
   gOrdenaGrid tdb_TiposDeVias, ColIndex
End Sub

Private Sub tdb_TiposDeVias_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
mblnClickOk = True
End Sub

Private Sub tdb_TiposDeVias_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
With tdb_TiposDeVias
    If (Not .EOF And Not .BOF) And mblnClickOk Then
        mblnClickOk = False
        blnAlterando = True
        txtPkid = tdb_TiposDeVias.Columns("Pkid")
        If blnPrimeiraVez Then
                TrocaCorObjeto txtintCodigo, True
                LeDaTabelaParaObj gstrTiposDeVias, Me
        End If
        strCodigoAtual = tdb_TiposDeVias.Columns("intcodigo").Value
        strDescricaoAtual = tdb_TiposDeVias.Columns("strDescricao").Value
        strSiglaAtual = tdb_TiposDeVias.Columns("strSigla").Value
    End If
End With
End Sub

Private Sub txtintCodigo_GotFocus()
    gstrProximoCodigo txtintCodigo, gstrTiposDeVias, "intCodigo", gintCodSeguranca
    MarcaCampo txtintCodigo
End Sub

Private Sub txtintCodigo_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtintCodigo
End Sub

Private Function blnDadosOk() As Boolean

    blnDadosOk = False
    
    If Trim(txtintCodigo.text) = "" Then
        ExibeMensagem "O Código deve ser preenchido."
        If txtintCodigo.Enabled = True Then
           txtintCodigo.SetFocus
        End If
        Exit Function
    ElseIf Trim(txtstrDescricao.text) = "" Then
        ExibeMensagem "A Descrição deve ser preenchida."
        txtstrDescricao.SetFocus
        Exit Function
    ElseIf Trim(txtstrSigla.text) = "" Then
        ExibeMensagem "A Sigla deve ser preenchida."
        txtstrSigla.SetFocus
        Exit Function
    End If
        
    If Not blnAlterando Or (blnAlterando And UCase$(strCodigoAtual) <> UCase$(txtintCodigo.text)) Then
        If gblnExisteCodigo(1, gstrTiposDeVias, "intCodigo", txtintCodigo.text) Then
            ExibeMensagem "Já existe registro com esse código."
            txtintCodigo.SetFocus
            Exit Function
        End If
    End If
    If Not blnAlterando Or (blnAlterando And UCase$(txtstrDescricao.text) <> UCase$(strDescricaoAtual)) Then
        If gblnExisteCodigo(1, gstrTiposDeVias, "strDescricao", "'" & txtstrDescricao & "'") Then
            ExibeMensagem "Já existe registro com essa descrição."
            txtstrDescricao.SetFocus
            Exit Function
        End If
    End If
    If Not blnAlterando Or (blnAlterando And UCase$(txtstrSigla.text) <> UCase$(strSiglaAtual)) Then
        If gblnExisteCodigo(1, gstrTiposDeVias, "strSigla", "'" & txtstrSigla & "'") Then
            ExibeMensagem "Já existe registro com essa sigla."
            txtstrSigla.SetFocus
            Exit Function
        End If
    End If

blnDadosOk = True

End Function

Private Function strQueryAplicar() As String
Dim strSql As String
strSql = "SELECT PKId, strSigla "
strSql = strSql & " FROM " & gstrTiposDeVias
strSql = strSql & " ORDER BY strSigla"
strQueryAplicar = strSql

End Function
