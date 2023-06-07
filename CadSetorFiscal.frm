VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "TODG7.OCX"
Begin VB.Form frmCadSetorFiscal 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Setores Fiscais"
   ClientHeight    =   5115
   ClientLeft      =   3300
   ClientTop       =   3060
   ClientWidth     =   6465
   HelpContextID   =   22
   Icon            =   "CadSetorFiscal.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5115
   ScaleWidth      =   6465
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtPKId 
      Enabled         =   0   'False
      Height          =   270
      Left            =   3570
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   120
      Visible         =   0   'False
      Width           =   645
   End
   Begin TabDlg.SSTab tab_3dPasta 
      Height          =   4965
      Left            =   90
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   90
      Width           =   6285
      _ExtentX        =   11086
      _ExtentY        =   8758
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   529
      TabCaption(0)   =   "Setores Fiscais"
      TabPicture(0)   =   "CadSetorFiscal.frx":1042
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblintCodigo"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblstrDescricao"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "tdb_SetorFiscal"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txtintCodigo"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtstrDescricao"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      Begin VB.TextBox txtstrDescricao 
         Height          =   285
         Left            =   1020
         MaxLength       =   100
         TabIndex        =   1
         Top             =   780
         Width           =   5085
      End
      Begin VB.TextBox txtintCodigo 
         Height          =   285
         Left            =   1020
         MaxLength       =   8
         TabIndex        =   0
         Top             =   450
         Width           =   1305
      End
      Begin TrueOleDBGrid70.TDBGrid tdb_SetorFiscal 
         Height          =   3525
         Left            =   180
         TabIndex        =   4
         Top             =   1260
         Width           =   5985
         _ExtentX        =   10557
         _ExtentY        =   6218
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
         Columns(1).Caption=   "Código"
         Columns(1).DataField=   "intCodigo"
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
         Splits(0)._ColumnProps(8)=   "Column(1).Width=1667"
         Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=1588"
         Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=2"
         Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(14)=   "Column(2).Width=8334"
         Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=8255"
         Splits(0)._ColumnProps(17)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
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
         _StyleDefs(34)  =   "Splits(0).Columns(1).Style:id=46,.parent=13,.alignment=1"
         _StyleDefs(35)  =   "Splits(0).Columns(1).HeadingStyle:id=43,.parent=14"
         _StyleDefs(36)  =   "Splits(0).Columns(1).FooterStyle:id=44,.parent=15"
         _StyleDefs(37)  =   "Splits(0).Columns(1).EditorStyle:id=45,.parent=17"
         _StyleDefs(38)  =   "Splits(0).Columns(2).Style:id=32,.parent=13"
         _StyleDefs(39)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
         _StyleDefs(40)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
         _StyleDefs(41)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
         _StyleDefs(42)  =   "Named:id=33:Normal"
         _StyleDefs(43)  =   ":id=33,.parent=0"
         _StyleDefs(44)  =   "Named:id=34:Heading"
         _StyleDefs(45)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(46)  =   ":id=34,.wraptext=-1"
         _StyleDefs(47)  =   "Named:id=35:Footing"
         _StyleDefs(48)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(49)  =   "Named:id=36:Selected"
         _StyleDefs(50)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(51)  =   "Named:id=37:Caption"
         _StyleDefs(52)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(53)  =   "Named:id=38:HighlightRow"
         _StyleDefs(54)  =   ":id=38,.parent=33,.bgcolor=&H80000014&,.fgcolor=&H80000012&"
         _StyleDefs(55)  =   "Named:id=39:EvenRow"
         _StyleDefs(56)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(57)  =   "Named:id=40:OddRow"
         _StyleDefs(58)  =   ":id=40,.parent=33"
         _StyleDefs(59)  =   "Named:id=41:RecordSelector"
         _StyleDefs(60)  =   ":id=41,.parent=34"
         _StyleDefs(61)  =   "Named:id=42:FilterBar"
         _StyleDefs(62)  =   ":id=42,.parent=33"
      End
      Begin VB.Label lblstrDescricao 
         AutoSize        =   -1  'True
         Caption         =   "Descrição"
         Height          =   195
         Left            =   180
         TabIndex        =   6
         Top             =   870
         Width           =   720
      End
      Begin VB.Label lblintCodigo 
         AutoSize        =   -1  'True
         Caption         =   "Código"
         Height          =   195
         Left            =   405
         TabIndex        =   5
         Top             =   510
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmCadSetorFiscal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim mblnAlterando      As Boolean
    Dim mobjAux            As Object
    Dim mblnSelecionou     As Boolean
    Dim mblnPrimeiraVez    As Boolean
    Dim bytOrdenacao       As Byte
    Dim blnOrdenacaoAsc    As Boolean
    Dim strCodigoAtual     As String
    Dim strDescricaoAtual  As String

Private Sub Form_Activate()
    gintCodSeguranca = 586
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

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
  If KeyCode = vbKeyF1 Then
     Call_HtmlHelp Me.HelpContextID
  End If
End Sub

Private Sub Form_Load()
    
    bytOrdenacao = 2: blnOrdenacaoAsc = True
    
    mblnAlterando = False
    VerificaListaAutomatica gstrSetorFiscal, tdb_SetorFiscal, strQuery
    VerificaObjParaAplicar mobjAux
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar, gstrDeletar
    mblnSelecionou = False
    mblnPrimeiraVez = False
End Sub

Private Sub tdb_SetorFiscal_Click()
    mblnPrimeiraVez = True
    If glngQtdLinhaTDBGrid(tdb_SetorFiscal) = 1 Then
        tdb_SetorFiscal_RowColChange 0, 0
    End If
End Sub

Private Sub tdb_SetorFiscal_DblClick()
    MantemForm gstrAplicar
End Sub

Private Sub tdb_SetorFiscal_FilterChange()
    mblnPrimeiraVez = False
    gblnFilraCampos tdb_SetorFiscal
End Sub

Private Function strQuery() As String
    
Dim strSql  As String

    strSql = ""
    strSql = strSql & " SELECT PKId, intCodigo, strDescricao FROM "
    strSql = strSql & gstrSetorFiscal
    
    Select Case bytOrdenacao
    
        Case 0
            strSql = strSql & " ORDER BY pkid " & IIf(blnOrdenacaoAsc, " ASC", " DESC")
        Case 1
            strSql = strSql & " ORDER BY intCodigo " & IIf(blnOrdenacaoAsc, " ASC", " DESC")
        Case 2
            strSql = strSql & " ORDER BY strDescricao " & IIf(blnOrdenacaoAsc, " ASC", " DESC")
    
    End Select
    
    strQuery = strSql
    
End Function

Private Sub tdb_SetorFiscal_HeadClick(ByVal ColIndex As Integer)
    gOrdenaGrid tdb_SetorFiscal, ColIndex
End Sub

Private Sub tdb_SetorFiscal_KeyPress(KeyAscii As Integer)
    Select Case tdb_SetorFiscal.Col
        Case 1
            CaracterValido KeyAscii, "N", tdb_SetorFiscal
        Case Else
            CaracterValido KeyAscii, "A", tdb_SetorFiscal
    End Select
End Sub

Private Sub tdb_SetorFiscal_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    With tdb_SetorFiscal
        If Not .EOF And Not .BOF Then
            If mblnPrimeiraVez Then
                mblnAlterando = True
                txtPKID.Text = .Columns("PKID").Value
                LeDaTabelaParaObj gstrSetorFiscal, Me
                gCorLinhaSelecionada tdb_SetorFiscal
                HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrDeletar
                If mobjAux Is Nothing Then
                    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar
                Else
                    HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrAplicar
                End If
                strCodigoAtual = tdb_SetorFiscal.Columns("intcodigo").Value
                strDescricaoAtual = tdb_SetorFiscal.Columns("strDescricao").Value
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
        ToolBarGeral strModoOperacao, gstrSetorFiscal, mblnAlterando, tdb_SetorFiscal, Me, mobjAux, strSql, , rptsetorfiscal, strQuery
        Exit Sub
    End If
    
    If UCase(strModoOperacao) = "SALVAR" Then
        If blnDadosOk = False Then Exit Sub
        mblnPrimeiraVez = False
        ToolBarGeral strModoOperacao, gstrSetorFiscal, mblnAlterando, tdb_SetorFiscal, Me, mobjAux, strSql
        Exit Sub
    ElseIf UCase(strModoOperacao) = "DELETAR" Then
        mblnPrimeiraVez = False
        ToolBarGeral strModoOperacao, gstrSetorFiscal, mblnAlterando, tdb_SetorFiscal, Me, mobjAux, strSql
        Exit Sub
    End If
    
    ToolBarGeral strModoOperacao, gstrSetorFiscal, mblnAlterando, tdb_SetorFiscal, Me, mobjAux, strSql
    
    If Not gobjGeral Is Nothing Then HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrDeletar
    
End Sub

Private Sub txtintCodigo_GotFocus()
    MarcaCampo txtintCodigo
    gstrProximoCodigo txtintCodigo, gstrSetorFiscal, "intCodigo", gintCodSeguranca
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
    blnDadosOk = False
    If Trim(txtintCodigo.Text) = "" Then
        ExibeMensagem "O Código deve ser informado."
        txtintCodigo.SetFocus
        Exit Function
    End If
    
    If Trim(txtstrDescricao.Text) = "" Then
        ExibeMensagem "A Descrição deve ser informada."
        txtstrDescricao.SetFocus
        Exit Function
    End If
    
    If Not mblnAlterando Or (mblnAlterando And UCase$(strCodigoAtual) <> UCase$(txtintCodigo.Text)) Then
        If gblnExisteCodigo(1, gstrSetorFiscal, "intCodigo", txtintCodigo.Text) Then
            ExibeMensagem "O Código informado já existe."
            txtintCodigo.SetFocus
            Exit Function
        End If
    End If
    If Not mblnAlterando Or (mblnAlterando And UCase$(txtstrDescricao.Text) <> UCase$(strDescricaoAtual)) Then
        If gblnExisteCodigo(1, gstrSetorFiscal, "strDescricao", "'" & txtstrDescricao & "'") Then
            ExibeMensagem "A Descrição informada já existe."
            txtstrDescricao.SetFocus
            Exit Function
        End If
    End If
    
    blnDadosOk = True
End Function
