VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "TODG7.OCX"
Begin VB.Form frmCadTabelaDeValores 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Valores"
   ClientHeight    =   5715
   ClientLeft      =   2640
   ClientTop       =   2550
   ClientWidth     =   7200
   HelpContextID   =   28
   Icon            =   "CadTabelaDeValores.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5715
   ScaleWidth      =   7200
   Begin VB.TextBox txtPKId 
      Height          =   270
      Left            =   5250
      TabIndex        =   10
      Top             =   90
      Visible         =   0   'False
      Width           =   645
   End
   Begin TabDlg.SSTab tab_3DPasta 
      Height          =   5535
      Left            =   120
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   90
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   9763
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   529
      TabCaption(0)   =   "Valores"
      TabPicture(0)   =   "CadTabelaDeValores.frx":1042
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lbldblValor"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblstrNomeDoValor"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lbl_CodigoDaUtilizacao"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "fra_bytTipo"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtdblValor"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtstrNomeDoValor"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "tdb_TabelaDeValores"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "dbcintCodigoDaUtilizacao"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      Begin MSDataListLib.DataCombo dbcintCodigoDaUtilizacao 
         Height          =   315
         Left            =   1020
         TabIndex        =   0
         Top             =   450
         Width           =   4605
         _ExtentX        =   8123
         _ExtentY        =   556
         _Version        =   393216
         IntegralHeight  =   0   'False
         Text            =   ""
      End
      Begin TrueOleDBGrid70.TDBGrid tdb_TabelaDeValores 
         Height          =   3165
         Left            =   120
         TabIndex        =   8
         Top             =   2220
         Width           =   6705
         _ExtentX        =   11827
         _ExtentY        =   5583
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
         Columns(1).Caption=   "Utilizacao"
         Columns(1).DataField=   "strUtilizacao"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Descrição"
         Columns(2).DataField=   "strNomeDoValor"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Valor"
         Columns(3).DataField=   "dblValor"
         Columns(3).NumberFormat=   "FormatText Event"
         Columns(3).EditMask=   "0,0000#.####"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   4
         Splits(0)._UserFlags=   0
         Splits(0).MarqueeStyle=   3
         Splits(0).RecordSelectors=   0   'False
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).DividerColor=   12632256
         Splits(0).FilterBar=   -1  'True
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=4"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0).AllowSizing=0"
         Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
         Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(8)=   "Column(1).Width=4657"
         Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=4577"
         Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(13)=   "Column(2).Width=4815"
         Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=4736"
         Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(17)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(18)=   "Column(3).Width=2170"
         Splits(0)._ColumnProps(19)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(20)=   "Column(3)._WidthInPix=2090"
         Splits(0)._ColumnProps(21)=   "Column(3)._EditAlways=0"
         Splits(0)._ColumnProps(22)=   "Column(3)._ColStyle=2"
         Splits(0)._ColumnProps(23)=   "Column(3).Order=4"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   0
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
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=32,.bold=0,.fontsize=825,.italic=0"
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
         _StyleDefs(34)  =   "Splits(0).Columns(1).Style:id=50,.parent=13"
         _StyleDefs(35)  =   "Splits(0).Columns(1).HeadingStyle:id=47,.parent=14"
         _StyleDefs(36)  =   "Splits(0).Columns(1).FooterStyle:id=48,.parent=15"
         _StyleDefs(37)  =   "Splits(0).Columns(1).EditorStyle:id=49,.parent=17"
         _StyleDefs(38)  =   "Splits(0).Columns(2).Style:id=32,.parent=13"
         _StyleDefs(39)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
         _StyleDefs(40)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
         _StyleDefs(41)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
         _StyleDefs(42)  =   "Splits(0).Columns(3).Style:id=46,.parent=13,.alignment=1"
         _StyleDefs(43)  =   "Splits(0).Columns(3).HeadingStyle:id=43,.parent=14"
         _StyleDefs(44)  =   "Splits(0).Columns(3).FooterStyle:id=44,.parent=15"
         _StyleDefs(45)  =   "Splits(0).Columns(3).EditorStyle:id=45,.parent=17"
         _StyleDefs(46)  =   "Named:id=33:Normal"
         _StyleDefs(47)  =   ":id=33,.parent=0"
         _StyleDefs(48)  =   "Named:id=34:Heading"
         _StyleDefs(49)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(50)  =   ":id=34,.wraptext=-1"
         _StyleDefs(51)  =   "Named:id=35:Footing"
         _StyleDefs(52)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(53)  =   "Named:id=36:Selected"
         _StyleDefs(54)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(55)  =   "Named:id=37:Caption"
         _StyleDefs(56)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(57)  =   "Named:id=38:HighlightRow"
         _StyleDefs(58)  =   ":id=38,.parent=33,.bgcolor=&H80000014&,.fgcolor=&H80000012&"
         _StyleDefs(59)  =   "Named:id=39:EvenRow"
         _StyleDefs(60)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(61)  =   "Named:id=40:OddRow"
         _StyleDefs(62)  =   ":id=40,.parent=33"
         _StyleDefs(63)  =   "Named:id=41:RecordSelector"
         _StyleDefs(64)  =   ":id=41,.parent=34"
         _StyleDefs(65)  =   "Named:id=42:FilterBar"
         _StyleDefs(66)  =   ":id=42,.parent=33"
      End
      Begin VB.TextBox txtstrNomeDoValor 
         Height          =   285
         Left            =   1020
         MaxLength       =   30
         TabIndex        =   1
         Top             =   840
         Width           =   4605
      End
      Begin VB.TextBox txtdblValor 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1020
         MaxLength       =   12
         MultiLine       =   -1  'True
         TabIndex        =   6
         Top             =   1200
         Width           =   1515
      End
      Begin VB.Frame fra_bytTipo 
         Caption         =   "Tipo "
         Height          =   585
         Left            =   1020
         TabIndex        =   9
         Top             =   1530
         Width           =   4605
         Begin VB.OptionButton optbytTipoDoValor 
            Caption         =   "Fator"
            Height          =   195
            Index           =   3
            Left            =   3780
            TabIndex        =   5
            Top             =   270
            Width           =   705
         End
         Begin VB.OptionButton optbytTipoDoValor 
            Caption         =   "Moeda"
            Height          =   195
            Index           =   2
            Left            =   2730
            TabIndex        =   4
            Top             =   270
            Width           =   945
         End
         Begin VB.OptionButton optbytTipoDoValor 
            Caption         =   "Quantidade"
            Height          =   195
            Index           =   1
            Left            =   1440
            TabIndex        =   3
            Top             =   270
            Width           =   1245
         End
         Begin VB.OptionButton optbytTipoDoValor 
            Caption         =   "Percentual"
            Height          =   195
            Index           =   0
            Left            =   180
            TabIndex        =   2
            Top             =   270
            Width           =   1095
         End
      End
      Begin VB.Label lbl_CodigoDaUtilizacao 
         AutoSize        =   -1  'True
         Caption         =   "Utilização"
         Height          =   195
         Left            =   240
         TabIndex        =   13
         Top             =   540
         Width           =   690
      End
      Begin VB.Label lblstrNomeDoValor 
         AutoSize        =   -1  'True
         Caption         =   "Descrição"
         Height          =   195
         Left            =   210
         TabIndex        =   12
         Top             =   885
         Width           =   720
      End
      Begin VB.Label lbldblValor 
         AutoSize        =   -1  'True
         Caption         =   "Valor"
         Height          =   195
         Left            =   585
         TabIndex        =   11
         Top             =   1290
         Width           =   360
      End
   End
End
Attribute VB_Name = "frmCadTabelaDeValores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim mblnAlterando           As Boolean
    Dim mobjAux                 As Object
    Dim mlngUltimo              As Long
    Dim mblnGuardaUltimo        As Boolean
    Dim mblnSelecionou          As Boolean
    Dim mblnPrimeiraVez         As Boolean
    Dim strUtilizacaoAtual      As String
    Dim strDescricaoAtual       As String
    
Private Sub dbcintCodigoDaUtilizacao_Click(Area As Integer)
    DropDownDataCombo dbcintCodigoDaUtilizacao, Me, Area
  '  If Area = 2 And dbcintCodigoDaUtilizacao.MatchedWithList Then
  '      mlngUltimo = dbcintCodigoDaUtilizacao.BoundText
  '      LeDaTabelaParaObj gstrTabelaDeValor, tdb_TabelaDeValores, strQueryTabelaDeValor
  '  End If
End Sub

Private Sub dbcintCodigoDaUtilizacao_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbcintCodigoDaUtilizacao, Me, , KeyCode, Shift
End Sub

Private Sub dbcintCodigoDaUtilizacao_KeyPress(KeyAscii As Integer)
CaracterValido KeyAscii, "A", dbcintCodigoDaUtilizacao
End Sub

Private Sub Form_Activate()
    gintCodSeguranca = 593
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
'============
End Sub

Private Sub Form_Deactivate()
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Dim cmlbtnBotao As MSComctlLib.Button
    If cmlbtnBotao Is Nothing = False Then
    End If
End Sub

Private Sub Form_Load()
    dbcintCodigoDaUtilizacao.Tag = strQueryDataComboUtilizacao & ";strNomeDaUtilizacao"
    LeDaTabelaParaObj "", dbcintCodigoDaUtilizacao, strQueryDataComboUtilizacao
    VerificaObjParaAplicar mobjAux
    optbytTipoDoValor(0).Value = True
End Sub

Private Function strQueryDataComboUtilizacao()
    Dim strSql As String
    strSql = ""
    strSql = strSql & "SELECT PKId, strNomeDaUtilizacao "
    strSql = strSql & "FROM " & gstrUtilizacaoDaTabelaDeValor & " "
    strSql = strSql & "ORDER BY strNomeDaUtilizacao"
    strQueryDataComboUtilizacao = strSql
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar, gstrDeletar
    mblnSelecionou = False
    mblnPrimeiraVez = False
End Sub

Private Sub optbytTipoDoValor_KeyPress(Index As Integer, KeyAscii As Integer)
CaracterValido KeyAscii, "A", optbytTipoDoValor(Index)
End Sub


Private Sub tdb_TabelaDeValores_Click()
    mblnPrimeiraVez = True
   If glngQtdLinhaTDBGrid(tdb_TabelaDeValores) = 1 Then
        tdb_TabelaDeValores_RowColChange 0, 0
    End If
End Sub

Private Sub tdb_TabelaDeValores_FilterChange()
    mblnPrimeiraVez = False
    gblnFilraCampos tdb_TabelaDeValores
End Sub

Private Sub tdb_TabelaDeValores_FormatText(ByVal ColIndex As Integer, Value As Variant, Bookmark As Variant)
If ColIndex = 2 Then
    Value = gstrConvVrDoSql(Value, 4)
End If
End Sub

Private Sub tdb_TabelaDeValores_HeadClick(ByVal ColIndex As Integer)
gOrdenaGrid tdb_TabelaDeValores, ColIndex
End Sub

Private Sub tdb_TabelaDeValores_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    
    With tdb_TabelaDeValores
        If Not .EOF And Not .BOF Then
            If mblnPrimeiraVez Then
                txtPKId.Text = .Columns("PKID").Value
                mblnAlterando = True
                LeDaTabelaParaObj gstrTabelaDeValor, Me

                gCorLinhaSelecionada tdb_TabelaDeValores
                
                HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrDeletar
                
                If mobjAux Is Nothing Then
                    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar
                Else

                    HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrAplicar
                End If
                strUtilizacaoAtual = dbcintCodigoDaUtilizacao.BoundText
                strDescricaoAtual = tdb_TabelaDeValores.Columns("strNomeDoValor").Value
                mblnSelecionou = True
            End If
        End If
    End With
End Sub

Public Sub MantemForm(ByVal strModoOperacao As String)
Dim varBookMark As Variant
Dim strSql As String

strSql = strQuery

If strModoOperacao = UCase(gstrImprimir) Then
    ToolBarGeral strModoOperacao, gstrTabelaDeValor, mblnAlterando, tdb_TabelaDeValores, Me, mobjAux, strSql, , rptTabelaDeValores, strQuery
    Exit Sub
End If

If UCase(strModoOperacao) = "SALVAR" Or UCase(strModoOperacao) = "DELETAR" Then
    mblnPrimeiraVez = False
End If

If UCase(strModoOperacao) = "SALVAR" Then
    If blnDadosOk = False Then Exit Sub
    ToolBarGeral strModoOperacao, gstrTabelaDeValor, mblnAlterando, tdb_TabelaDeValores, Me, mobjAux, strQueryTabelaDeValor, strSql
    Exit Sub
End If

If UCase(strModoOperacao) = gstrNovo Then
    mblnSelecionou = False
    mblnPrimeiraVez = False
End If

If UCase(strModoOperacao) = UCase(gstrLocalizar) Then
    LeDaTabelaParaObj gstrTabelaDeValor, tdb_TabelaDeValores, strQueryTabelaDevalores
    Exit Sub
End If

ToolBarGeral strModoOperacao, gstrTabelaDeValor, mblnAlterando, tdb_TabelaDeValores, Me, mobjAux, strQueryTabelaDeValor, strSql

HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrDeletar



End Sub
Private Sub txtstrNomeDoValor_GotFocus()
    MarcaCampo txtstrNomeDoValor
End Sub

Private Sub txtstrNomeDoValor_KeyPress(KeyAscii As Integer)
CaracterValido KeyAscii, "A", txtstrNomeDoValor
End Sub

Private Sub txtdblValor_GotFocus()
    MarcaCampo txtdblValor
End Sub

Private Sub txtdblValor_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "V", txtdblValor
End Sub

Private Sub txtdblValor_LostFocus()
    txtdblValor = gvntConvVrDoSql(txtdblValor)
End Sub

Private Function strQuery() As String
    Dim strSql As String
    strSql = ""
    strSql = strSql & "SELECT TB.PKId, TB.strNomeDoValor, TB.dblValor "
    strSql = strSql & "FROM " & gstrTabelaDeValor & " TB "
    strQuery = strSql
End Function

Private Function strQueryTabelaDeValor() As String
    Dim strSql As String
    Dim intFor As Integer
    
    strSql = ""
    strSql = strSql & "Select PKId, PKId Codigo, strNomeDoValor, dblValor "
    strSql = strSql & "From " & gstrTabelaDeValor & " "
'    If dbcintCodigoDaUtilizacao.MatchedWithList Then
'        strSql = strSql & "Where intCodigoDaUtilizacao = " & dbcintCodigoDaUtilizacao.BoundText
'    End If
    For intFor = 0 To 3
        If optbytTipoDoValor(intFor).Value Then
            strSql = strSql & " Where bytTipoDoValor = " & intFor
            Exit For
        End If
    Next
    strQueryTabelaDeValor = strSql
End Function

Private Function blnDadosOk() As Boolean
    blnDadosOk = False
    If dbcintCodigoDaUtilizacao.MatchedWithList = False Then
        ExibeMensagem "O utilização código deve ser preenchido corretamente."
        dbcintCodigoDaUtilizacao.SetFocus
        Exit Function
    ElseIf Trim(txtstrNomeDoValor) = "" Then
        ExibeMensagem "O campo descrição deve ser preenchido corretamente."
        txtstrNomeDoValor.SetFocus
        Exit Function
    ElseIf Trim(txtdblValor) = "" Then
        ExibeMensagem "O campo valor deve ser preenchido corretamente."
        txtdblValor.SetFocus
        Exit Function
    End If
    
    If Not mblnAlterando Or (mblnAlterando And UCase$(Trim(txtstrNomeDoValor.Text)) <> UCase$(Trim(strDescricaoAtual))) Or (mblnAlterando And (strUtilizacaoAtual <> dbcintCodigoDaUtilizacao.BoundText)) Then
        If gblnExisteCodigo(2, gstrTabelaDeValor, "Intcodigodautilizacao", dbcintCodigoDaUtilizacao.BoundText, "strNomeDoValor", "'" & Trim(txtstrNomeDoValor) & "'") Then
            ExibeMensagem "A descrição já existe para esta Utilização."
            Exit Function
        End If
    End If
    
    blnDadosOk = True
End Function


Private Function strQueryTabelaDevalores() As String
Dim strSql As String
Dim intFor As Integer

    strSql = "SELECT UTV.Strnomedautilizacao strUtilizacao, TVL.Strnomedovalor, TVL.Dblvalor "
    strSql = strSql & "FROM " & gstrUtilizacaoDaTabelaDeValor & " UTV, " & gstrTabelaDeValor & " TVL "
    strSql = strSql & "WHERE TVL.Intcodigodautilizacao = UTV.Pkid"
    
    For intFor = 0 To 3
        If optbytTipoDoValor(intFor).Value Then
            strSql = strSql & " AND bytTipoDoValor = " & intFor
            Exit For
        End If
    Next
    
    If dbcintCodigoDaUtilizacao.MatchedWithList Then
        strSql = strSql & " AND TVL.intCodigoDaUtilizacao = " & dbcintCodigoDaUtilizacao.BoundText
    End If
    
    strQueryTabelaDevalores = strSql

End Function
