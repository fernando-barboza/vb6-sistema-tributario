VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "TODG7.OCX"
Begin VB.Form frmCadCaracteristicasGerais 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Características Gerais"
   ClientHeight    =   5820
   ClientLeft      =   5910
   ClientTop       =   4095
   ClientWidth     =   6660
   Icon            =   "CadCaracteristicasGerais.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5820
   ScaleWidth      =   6660
   Begin VB.TextBox txtPKId 
      Height          =   270
      Left            =   1560
      TabIndex        =   12
      Top             =   135
      Visible         =   0   'False
      Width           =   645
   End
   Begin TabDlg.SSTab tab_3DPasta 
      Height          =   5505
      Left            =   150
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   150
      Width           =   6315
      _ExtentX        =   11139
      _ExtentY        =   9710
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   529
      TabCaption(0)   =   "Características"
      TabPicture(0)   =   "CadCaracteristicasGerais.frx":1042
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblintUtilizacaoDaCaracteristica"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblintCodigoDaCaracteristica"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblstrNomeDaCaracteristica"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblintCategoriaConstrucao"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "dbcintCategoriaConstrucao"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "tdb_CaracteristicaGeral"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtintCodigoDaCaracteristica"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtstrNomeDaCaracteristica"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "dbcintUtilizacaoDaCaracteristica"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "chkBytFator"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "chkBytCaracteristica"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).ControlCount=   11
      Begin VB.CheckBox chkBytCaracteristica 
         Caption         =   "Característica de Boletim"
         Height          =   195
         Left            =   3585
         TabIndex        =   10
         Top             =   1920
         Width           =   2160
      End
      Begin VB.CheckBox chkBytFator 
         Caption         =   "Fator de Correção"
         Height          =   195
         Left            =   1455
         TabIndex        =   9
         Top             =   1920
         Width           =   1635
      End
      Begin MSDataListLib.DataCombo dbcintUtilizacaoDaCaracteristica 
         Height          =   315
         Left            =   1275
         TabIndex        =   2
         Top             =   450
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   556
         _Version        =   393216
         IntegralHeight  =   0   'False
         Text            =   ""
      End
      Begin VB.TextBox txtstrNomeDaCaracteristica 
         Height          =   285
         Left            =   1275
         MaxLength       =   50
         TabIndex        =   8
         Top             =   1515
         Width           =   4815
      End
      Begin VB.TextBox txtintCodigoDaCaracteristica 
         Height          =   285
         Left            =   1275
         MaxLength       =   15
         TabIndex        =   6
         Top             =   1170
         Width           =   1035
      End
      Begin TrueOleDBGrid70.TDBGrid tdb_CaracteristicaGeral 
         Height          =   3060
         Left            =   255
         TabIndex        =   11
         Top             =   2295
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   5398
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
         Columns(1).DataField=   "Codigo"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Descrição"
         Columns(2).DataField=   "strNomeDaCaracteristica"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   3
         Splits(0)._UserFlags=   0
         Splits(0).MarqueeStyle=   3
         Splits(0).RecordSelectors=   0   'False
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).DividerColor=   12632256
         Splits(0).FilterBar=   -1  'True
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=3"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=1746"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1667"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0).Visible=0"
         Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(7)=   "Column(1).Width=1455"
         Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=1376"
         Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(12)=   "Column(2).Width=8017"
         Splits(0)._ColumnProps(13)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(14)=   "Column(2)._WidthInPix=7938"
         Splits(0)._ColumnProps(15)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(16)=   "Column(2).Order=3"
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
         _StyleDefs(31)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
         _StyleDefs(32)  =   "Splits(0).EditorStyle:id=17,.parent=7"
         _StyleDefs(33)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
         _StyleDefs(34)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
         _StyleDefs(35)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
         _StyleDefs(36)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(37)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(38)  =   "Splits(0).Columns(0).Style:id=50,.parent=13"
         _StyleDefs(39)  =   "Splits(0).Columns(0).HeadingStyle:id=47,.parent=14"
         _StyleDefs(40)  =   "Splits(0).Columns(0).FooterStyle:id=48,.parent=15"
         _StyleDefs(41)  =   "Splits(0).Columns(0).EditorStyle:id=49,.parent=17"
         _StyleDefs(42)  =   "Splits(0).Columns(1).Style:id=46,.parent=13"
         _StyleDefs(43)  =   "Splits(0).Columns(1).HeadingStyle:id=43,.parent=14"
         _StyleDefs(44)  =   "Splits(0).Columns(1).FooterStyle:id=44,.parent=15"
         _StyleDefs(45)  =   "Splits(0).Columns(1).EditorStyle:id=45,.parent=17"
         _StyleDefs(46)  =   "Splits(0).Columns(2).Style:id=58,.parent=13"
         _StyleDefs(47)  =   "Splits(0).Columns(2).HeadingStyle:id=55,.parent=14"
         _StyleDefs(48)  =   "Splits(0).Columns(2).FooterStyle:id=56,.parent=15"
         _StyleDefs(49)  =   "Splits(0).Columns(2).EditorStyle:id=57,.parent=17"
         _StyleDefs(50)  =   "Named:id=33:Normal"
         _StyleDefs(51)  =   ":id=33,.parent=0"
         _StyleDefs(52)  =   "Named:id=34:Heading"
         _StyleDefs(53)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(54)  =   ":id=34,.wraptext=-1"
         _StyleDefs(55)  =   "Named:id=35:Footing"
         _StyleDefs(56)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(57)  =   "Named:id=36:Selected"
         _StyleDefs(58)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(59)  =   "Named:id=37:Caption"
         _StyleDefs(60)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(61)  =   "Named:id=38:HighlightRow"
         _StyleDefs(62)  =   ":id=38,.parent=33,.bgcolor=&H80000014&,.fgcolor=&H80000012&"
         _StyleDefs(63)  =   "Named:id=39:EvenRow"
         _StyleDefs(64)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(65)  =   "Named:id=40:OddRow"
         _StyleDefs(66)  =   ":id=40,.parent=33"
         _StyleDefs(67)  =   "Named:id=41:RecordSelector"
         _StyleDefs(68)  =   ":id=41,.parent=34"
         _StyleDefs(69)  =   "Named:id=42:FilterBar"
         _StyleDefs(70)  =   ":id=42,.parent=33"
      End
      Begin MSDataListLib.DataCombo dbcintCategoriaConstrucao 
         Height          =   315
         Left            =   1275
         TabIndex        =   4
         Top             =   810
         Width           =   4830
         _ExtentX        =   8520
         _ExtentY        =   556
         _Version        =   393216
         IntegralHeight  =   0   'False
         Text            =   ""
      End
      Begin VB.Label lblintCategoriaConstrucao 
         AutoSize        =   -1  'True
         Caption         =   "Categoria"
         Height          =   195
         Left            =   525
         TabIndex        =   3
         Top             =   915
         Width           =   675
      End
      Begin VB.Label lblstrNomeDaCaracteristica 
         AutoSize        =   -1  'True
         Caption         =   "Nome"
         Height          =   195
         Left            =   750
         TabIndex        =   7
         Top             =   1590
         Width           =   420
      End
      Begin VB.Label lblintCodigoDaCaracteristica 
         AutoSize        =   -1  'True
         Caption         =   "Código"
         Height          =   195
         Left            =   690
         TabIndex        =   5
         Top             =   1260
         Width           =   495
      End
      Begin VB.Label lblintUtilizacaoDaCaracteristica 
         AutoSize        =   -1  'True
         Caption         =   "Utilização"
         Height          =   195
         Left            =   525
         TabIndex        =   1
         Top             =   525
         Width           =   690
      End
   End
End
Attribute VB_Name = "frmCadCaracteristicasGerais"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim blnAlterando        As Boolean
    Dim mobjAux             As Object
    Dim blnSelecionou       As Boolean
    Dim bytOrdenacao        As Byte
    Dim blnOrdenacaoAsc     As Boolean
    Dim blnPrimeiraVez      As Boolean
    Dim strCodigoAtual      As String
    Dim strNomeCaracAtual   As String
    
    
    
Private Sub dbcintUtilizacaoDaCaracteristica_Click(Area As Integer)
    If Area = 0 Then
        DropDownDataCombo dbcintUtilizacaoDaCaracteristica, Me, Area
    ElseIf Area = 2 Then
        dbcintCategoriaConstrucao.Tag = strQueryCategoriaConstrucao & ";strDescricao"
        LeDaTabelaParaObj gstrCategoriaConstrucao, dbcintCategoriaConstrucao, strQueryCategoriaConstrucao
        Set tdb_CaracteristicaGeral.DataSource = Nothing
        txtintCodigoDaCaracteristica.Text = ""
        txtstrNomeDaCaracteristica.Text = ""
        chkBytCaracteristica.Value = 0
        chkBytFator.Value = 0
    End If
       
End Sub

Private Sub dbcintUtilizacaoDaCaracteristica_GotFocus()
    MarcaCampo dbcintUtilizacaoDaCaracteristica
End Sub

Private Sub dbcintUtilizacaoDaCaracteristica_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbcintUtilizacaoDaCaracteristica, Me, , KeyCode, Shift
End Sub

Private Sub dbcintUtilizacaoDaCaracteristica_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbcintUtilizacaoDaCaracteristica
End Sub

Private Sub dbcintCategoriaConstrucao_Click(Area As Integer)
    If Area = 0 Then
        DropDownDataCombo dbcintCategoriaConstrucao, Me, Area
    ElseIf Area = 2 Then
        LeDaTabelaParaObj gstrCaracteristicaGeral, tdb_CaracteristicaGeral, strQueryGrid
        txtintCodigoDaCaracteristica.Text = ""
        txtstrNomeDaCaracteristica.Text = ""
        chkBytCaracteristica.Value = 0
        chkBytFator.Value = 0
    End If
       
End Sub

Private Sub dbcintCategoriaConstrucao_GotFocus()
    MarcaCampo dbcintCategoriaConstrucao
End Sub

Private Sub dbcintCategoriaConstrucao_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbcintCategoriaConstrucao, Me, , KeyCode, Shift
End Sub

Private Sub dbcintCategoriaConstrucao_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbcintCategoriaConstrucao
End Sub

Private Sub Form_Activate()
    
    gintCodSeguranca = 1060
    If mobjAux Is Nothing Then
        HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar
    Else
        HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrAplicar
        If VerificaFormAtivo Then PreencherListaDeOpcoes dbcintUtilizacaoDaCaracteristica, frmCadDetalheDaCaracteristica.dbcintUtilizacaoDaCaracteristica.BoundText
        If VerificaFormAtivo Then dbcintUtilizacaoDaCaracteristica_Click 2
        If VerificaFormAtivo Then TrocaCorObjeto dbcintUtilizacaoDaCaracteristica, True
        If InStr(UCase(Mid(mobjAux.Name, 4, Len(mobjAux.Name) - 3)), "CARACTERISTICA") > 0 Then
            chkBytFator.Enabled = False
            chkBytFator.Value = 0
            chkBytCaracteristica.Value = 1
            chkBytCaracteristica.Enabled = False
            chkBytCaracteristica.CausesValidation = False
        End If
        If InStr(UCase(Mid(mobjAux.Name, 4, Len(mobjAux.Name) - 3)), "HORIZONTAL") > 0 Then 'FATOR
            chkBytCaracteristica.Enabled = False
            chkBytFator.Value = 1
            chkBytCaracteristica.Value = 0
            chkBytFator.Enabled = False
            chkBytFator.CausesValidation = False
        End If
    End If
    If VerificaFormAtivo Then
        If Not frmCadDetalheDaCaracteristica.dbcintUtilizacaoDaCaracteristica.MatchedWithList Then
            ExibeMensagem "Selecione a utilização."
            Unload frmCadCaracteristicasGerais
            frmCadDetalheDaCaracteristica.dbcintUtilizacaoDaCaracteristica.SetFocus
            Exit Sub
        Else
            If Not frmCadDetalheDaCaracteristica.dbcintCategoriaConstrucao.MatchedWithList Then
                ExibeMensagem "Selecione a Categoria"
                Unload frmCadCaracteristicasGerais
                frmCadDetalheDaCaracteristica.dbcintCategoriaConstrucao.SetFocus
                Exit Sub
            End If
        End If
        dbcintCategoriaConstrucao.Text = frmCadDetalheDaCaracteristica.dbcintCategoriaConstrucao.Text
        TrocaCorObjeto dbcintCategoriaConstrucao, True, True
        txtintCodigoDaCaracteristica.Text = Empty
        txtintCodigoDaCaracteristica.SetFocus
    End If
End Sub

Private Sub Form_Deactivate()
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar
End Sub

Private Sub Form_Load()
    VerificaObjParaAplicar mobjAux
    dbcintUtilizacaoDaCaracteristica.Tag = strQueryUtilizacao & ";strNomeDaUtilizacao"
    LeDaTabelaParaObj gstrUtilizacaoDaTabelaDeValor, dbcintUtilizacaoDaCaracteristica, strQueryUtilizacao
    dbcintCategoriaConstrucao.Text = frmCadDetalheDaCaracteristica.dbcintCategoriaConstrucao
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar, gstrDeletar
    blnPrimeiraVez = False
End Sub

Private Sub tdb_CaracteristicaGeral_Click()
    blnPrimeiraVez = True
End Sub

Private Sub tdb_CaracteristicaGeral_DblClick()
    MantemForm gstrAplicar
End Sub

Private Sub tdb_CaracteristicaGeral_FilterChange()
    gblnFilraCampos tdb_CaracteristicaGeral
End Sub
Public Sub MantemForm(ByVal strModoOperacao As String)

Select Case UCase(strModoOperacao)
    Case Is = UCase(gstrPreencherLista)
        PreencherListaDeOpcoes Me.ActiveControl
    Case Else
        If UCase(strModoOperacao) = UCase(gstrSalvar) Then
            If Not blnDadosOk Then Exit Sub
        End If
        
        ToolBarGeral strModoOperacao, gstrCaracteristicaGeral, blnAlterando, tdb_CaracteristicaGeral, _
                Me, mobjAux, strQueryGrid, strQueryAplicar
End Select

If UCase(strModoOperacao) = UCase(gstrNovo) Or UCase(strModoOperacao) = UCase(gstrSalvar) Then
    ProximoCodigo txtintCodigoDaCaracteristica, gstrCaracteristicaGeral, "intCodigoDaCaracteristica", gintCodSeguranca, "intUtilizacaoDaCaracteristica", dbcintUtilizacaoDaCaracteristica.BoundText
    txtintCodigoDaCaracteristica.SetFocus
    MantemForm gstrRefresh
End If

End Sub

Private Function strQueryUtilizacao() As String
Dim strSql As String
    strSql = "SELECT Pkid, strNomeDaUtilizacao "
    strSql = strSql & "FROM "
    strSql = strSql & gstrUtilizacaoDaTabelaDeValor
    strSql = strSql & " ORDER BY strNomeDaUtilizacao"

strQueryUtilizacao = strSql

End Function

Private Function strQueryGrid() As String
Dim strSql As String
strSql = "SELECT Pkid,"
strSql = strSql & " intCodigoDaCaracteristica Codigo,"
strSql = strSql & " strNomeDaCaracteristica"
strSql = strSql & " FROM "
strSql = strSql & gstrCaracteristicaGeral
strSql = strSql & " WHERE intUtilizacaoDaCaracteristica = '" & dbcintUtilizacaoDaCaracteristica.BoundText & "'"

If dbcintCategoriaConstrucao.MatchedWithList Then
    strSql = strSql & " AND intCategoriaConstrucao = " & dbcintCategoriaConstrucao.BoundText
End If

If InStr(UCase(Mid(mobjAux.Name, 4, Len(mobjAux.Name) - 3)), "CARACTERISTICA") > 0 Then
    strSql = strSql & " AND bytCaracteristica = 1"
ElseIf InStr(UCase(Mid(mobjAux.Name, 4, Len(mobjAux.Name) - 3)), "HORIZONTAL") > 0 Then 'FATOR
    strSql = strSql & " AND bytFator =1"
End If

Select Case bytOrdenacao
    Case Is = 1
        strSql = strSql & " ORDER BY intCodigoDaCaracteristica " & IIf(blnOrdenacaoAsc, "ASC", "DESC")
    Case Is = 2
        strSql = strSql & " ORDER BY strNomeDaCaracteristica " & IIf(blnOrdenacaoAsc, "ASC", "DESC")
    Case Is = 3
        strSql = strSql & " ORDER BY strNomeDaCaracteristica " & IIf(blnOrdenacaoAsc, "ASC", "DESC")
End Select
strQueryGrid = strSql

End Function

Private Function strQueryCategoriaConstrucao() As String
Dim strSql As String

strSql = "SELECT Pkid,"
strSql = strSql & " strDescricao"
strSql = strSql & " FROM "
strSql = strSql & gstrCategoriaConstrucao
strSql = strSql & " WHERE intUtilizacaoTabelaValor = '" & dbcintUtilizacaoDaCaracteristica.BoundText & "'"

strQueryCategoriaConstrucao = strSql

End Function

Private Sub tdb_CaracteristicaGeral_HeadClick(ByVal ColIndex As Integer)
   
   blnOrdenacaoAsc = IIf(bytOrdenacao = ColIndex, Not blnOrdenacaoAsc, True)
   
   bytOrdenacao = ColIndex: MantemForm gstrRefresh

End Sub

Private Function strQueryAplicar() As String

Dim strSql As String
strSql = "SELECT PKId, strNomeDaCaracteristica "
strSql = strSql & " FROM " & gstrCaracteristicaGeral
strSql = strSql & " WHERE intUtilizacaoDaCaracteristica = '" & dbcintUtilizacaoDaCaracteristica.BoundText & "'"
strSql = strSql & " AND bytCaracteristica = 1"
strSql = strSql & " ORDER BY strNomeDaCaracteristica"

strQueryAplicar = strSql

End Function

Private Sub tdb_CaracteristicaGeral_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
If blnPrimeiraVez Then
    blnAlterando = True
    txtPKId = tdb_CaracteristicaGeral.Columns("Pkid")
    LeDaTabelaParaObj gstrCaracteristicaGeral, Me
    blnPrimeiraVez = False
    
    strCodigoAtual = tdb_CaracteristicaGeral.Columns("Código")
    strNomeCaracAtual = tdb_CaracteristicaGeral.Columns("Descrição")
    
End If
End Sub

Public Function ProximoCodigo(txtDestino As TextBox, _
                              strTabela As String, _
                              strCampo As String, _
                              intCodigo As Integer, _
                              Optional strGrupo As String, _
                              Optional strValorGrupo As String, _
                              Optional intMascaraEspecifica As Integer, _
                              Optional Retorno As Boolean, _
                              Optional strSubGrupo As String, _
                              Optional strValorSubGrupo As String, _
                              Optional strParametroEspecifico As String, _
                              Optional strValorParametroEspecifico As String) As String
                              
Dim strSql          As String
Dim AdoResultado    As ADODB.Recordset

Dim strSistema      As String


    
    If Not Retorno Then If txtDestino.Text <> "" Then Exit Function
    Set gobjBanco = New clsBanco
    
    
    strSql = "SELECT bitAutoNumeracao FROM " & gstrItens
    strSql = strSql & " WHERE intCodigo = " & intCodigo & " AND "
    strSql = strSql & "UPPER(" & strSUBSTRING & "(strCodItem,1,1)) = '" & strSistema & "' "
            
    strSql = "SELECT " & gstrTOPnSQLServer(1) & " (" & gstrREPLICATE(strCampo, "0", 10) & ") as ProximoCodigo,10 - " & strLEN & "(" & strCampo & ") As TotalZeros FROM " & strTabela
    strSql = strSql & " WHERE " & strGrupo & " = " & dbcintUtilizacaoDaCaracteristica.BoundText
    strSql = strSql & " AND intCategoriaConstrucao = " & dbcintCategoriaConstrucao.BoundText
    If InStr(UCase(Mid(mobjAux.Name, 4, Len(mobjAux.Name) - 3)), "CARACTERISTICA") > 0 Then
        strSql = strSql & " AND bytCaracteristica = 1"
    ElseIf InStr(UCase(Mid(mobjAux.Name, 4, Len(mobjAux.Name) - 3)), "HORIZONTAL") > 0 Then 'FATOR
        strSql = strSql & " AND bytFator =1"
    End If

    
    If strParametroEspecifico <> "" Then
        strSql = strSql & " AND " & strParametroEspecifico & "=" & IIf(UCase$(Left(strParametroEspecifico, 3)) = "STR", "'" & strValorParametroEspecifico & "'", Val(strValorParametroEspecifico))
    End If
    strSql = strSql & " GROUP BY " & strCampo & " ORDER BY ProximoCodigo DESC "
    strSql = gstrTOPnOracle(strSql, 1)
    
    If gobjBanco.CriaADO(strSql, 5, AdoResultado) Then
        If Not AdoResultado.EOF Then
            If Not IsNull(AdoResultado("ProximoCodigo")) Then
                If Retorno Then
                    If IsNumeric(Right(Trim(AdoResultado("ProximoCodigo")), Len(Trim(AdoResultado("ProximoCodigo"))) - AdoResultado("TotalZeros"))) Then
                        ProximoCodigo = Right(Trim(AdoResultado("ProximoCodigo")), Len(Trim(AdoResultado("ProximoCodigo"))) - AdoResultado("TotalZeros")) + 1
                    Else
                        If IsNumeric(Right(Right(Trim(AdoResultado("ProximoCodigo")), Len(Trim(AdoResultado("ProximoCodigo"))) - AdoResultado("TotalZeros")), 1)) Then
                            ProximoCodigo = Left(Right(Trim(AdoResultado("ProximoCodigo")), Len(Trim(AdoResultado("ProximoCodigo"))) - AdoResultado("TotalZeros")), Len(Right(Trim(AdoResultado("ProximoCodigo")), Len(Trim(AdoResultado("ProximoCodigo"))) - AdoResultado("TotalZeros"))) - 1) & (Right(Right(Trim(AdoResultado("ProximoCodigo")), Len(Trim(AdoResultado("ProximoCodigo"))) - AdoResultado("TotalZeros")), 1) + 1)
                        Else
                            ProximoCodigo = Left(Right(Trim(AdoResultado("ProximoCodigo")), Len(Trim(AdoResultado("ProximoCodigo"))) - AdoResultado("TotalZeros")), Len(Right(Trim(AdoResultado("ProximoCodigo")), Len(Trim(AdoResultado("ProximoCodigo"))) - AdoResultado("TotalZeros"))) - 1) & Chr(Asc(Right(Right(Trim(AdoResultado("ProximoCodigo")), Len(Trim(AdoResultado("ProximoCodigo"))) - AdoResultado("TotalZeros")), 1)) + 1)
                        End If
                    End If
                Else
                    If IsNumeric(Right(Trim(AdoResultado("ProximoCodigo")), Len(Trim(AdoResultado("ProximoCodigo"))) - AdoResultado("TotalZeros"))) Then
                        txtDestino.Text = Right(Trim(AdoResultado("ProximoCodigo")), Len(Trim(AdoResultado("ProximoCodigo"))) - AdoResultado("TotalZeros")) + 1
                    Else
                        If IsNumeric(Right(Right(Trim(AdoResultado("ProximoCodigo")), Len(Trim(AdoResultado("ProximoCodigo"))) - AdoResultado("TotalZeros")), 1)) Then
                            txtDestino.Text = Left(Right(Trim(AdoResultado("ProximoCodigo")), Len(Trim(AdoResultado("ProximoCodigo"))) - AdoResultado("TotalZeros")), Len(Right(Trim(AdoResultado("ProximoCodigo")), Len(Trim(AdoResultado("ProximoCodigo"))) - AdoResultado("TotalZeros"))) - 1) & (Right(Right(Trim(AdoResultado("ProximoCodigo")), Len(Trim(AdoResultado("ProximoCodigo"))) - AdoResultado("TotalZeros")), 1) + 1)
                        Else
                            txtDestino.Text = Left(Right(Trim(AdoResultado("ProximoCodigo")), Len(Trim(AdoResultado("ProximoCodigo"))) - AdoResultado("TotalZeros")), Len(Right(Trim(AdoResultado("ProximoCodigo")), Len(Trim(AdoResultado("ProximoCodigo"))) - AdoResultado("TotalZeros"))) - 1) & Chr(Asc(Right(Right(Trim(AdoResultado("ProximoCodigo")), Len(Trim(AdoResultado("ProximoCodigo"))) - AdoResultado("TotalZeros")), 1)) + 1)
                        End If
                    End If
                End If
            Else
                If Retorno Then
                    ProximoCodigo = "1"
                Else
                    txtDestino.Text = "1"
                End If
            End If
        Else
            If strValorGrupo <> "" And strValorGrupo <> "0" Then
                If Retorno Then
                    ProximoCodigo = "1"
                Else
                    txtDestino.Text = "1"
                End If
            End If
        End If
        AdoResultado.Close
        Set AdoResultado = Nothing
    End If
    
Set gobjBanco = Nothing
    
End Function


Private Sub txtintCodigoDaCaracteristica_GotFocus()
    If dbcintCategoriaConstrucao.MatchedWithList Then
        ProximoCodigo txtintCodigoDaCaracteristica, gstrCaracteristicaGeral, "intCodigoDaCaracteristica", gintCodSeguranca, "intUtilizacaoDaCaracteristica", dbcintUtilizacaoDaCaracteristica.BoundText
        MarcaCampo txtintCodigoDaCaracteristica
    Else
        ExibeMensagem "Selecione a categoria"
        dbcintCategoriaConstrucao.SetFocus
    End If
           
End Sub

Private Sub txtintCodigoDaCaracteristica_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtintCodigoDaCaracteristica
End Sub

Private Sub txtstrNomeDaCaracteristica_GotFocus()
    MarcaCampo txtstrNomeDaCaracteristica
End Sub


Private Sub txtstrNomeDaCaracteristica_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtintCodigoDaCaracteristica
End Sub

Private Function blnDadosOk() As Boolean

blnDadosOk = False

If txtintCodigoDaCaracteristica.Text = "" Then
    ExibeMensagem "Informe um Código válido."
    Exit Function
End If

If txtstrNomeDaCaracteristica.Text = "" Then
    ExibeMensagem "Informe um Nome da Característica válido."
    Exit Function
End If

If Not dbcintCategoriaConstrucao.MatchedWithList Then
    ExibeMensagem "Selecione uma Categoria de Construção."
    Exit Function
End If

'    If Not blnAlterando Or (blnAlterando And UCase$(strCodigoAtual) <> UCase$(txtintCodigoDaCaracteristica.Text)) Then

'ProximoCodigo:
        Dim strCodigo As String

  '      If gblnExisteCodigo(2, gstrCaracteristicaGeral, "intCodigoDaCaracteristica", "'" & txtintCodigoDaCaracteristica.Text & "'", "intUtilizacaoDaCaracteristica", "'" & dbcintUtilizacaoDaCaracteristica.BoundText & "'") Then
   '         txtintCodigoDaCaracteristica.Text = strCodigo
    '        GoTo ProximoCodigo
     '   End If
    'End If
'
    'If Not blnAlterando Or (blnAlterando And UCase$(txtstrNomeDaCaracteristica.Text) <> UCase$(strNomeCaracAtual)) Then
     '
     '   If gblnExisteCodigo(2, gstrCaracteristicaGeral, "strNomeDaCaracteristica", "'" & txtstrNomeDaCaracteristica.Text & "'", "intUtilizacaoDaCaracteristica", "'" & dbcintUtilizacaoDaCaracteristica.BoundText & "'") Then
     '       ExibeMensagem "O Nome da Característica já se encontra cadastrado."
     '       txtstrNomeDaCaracteristica.SetFocus
     '       Exit Function
     '   End If
    'End If


blnDadosOk = True

End Function
