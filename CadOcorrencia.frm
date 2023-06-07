VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "TODG7.OCX"
Begin VB.Form frmCadOcorrencia 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ocorrências"
   ClientHeight    =   5595
   ClientLeft      =   3390
   ClientTop       =   2775
   ClientWidth     =   6000
   HelpContextID   =   40
   Icon            =   "CadOcorrencia.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5595
   ScaleWidth      =   6000
   Begin TabDlg.SSTab tab_3DPasta 
      Height          =   5535
      Left            =   60
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   30
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   9763
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   529
      TabCaption(0)   =   "Ocorrências"
      TabPicture(0)   =   "CadOcorrencia.frx":1042
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblintUtilizacaoDaOcorrencia"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblstrDescricao"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblintCodigo"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txtstrDescricao"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtintCodigo"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "tdb_Ocorrencia"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "chkbytRemido"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "dbcintUtilizacaoDaOcorrencia"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtPKId"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).ControlCount=   9
      Begin VB.TextBox txtPKId 
         Height          =   270
         Left            =   1770
         TabIndex        =   9
         Top             =   30
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.ComboBox dbcintUtilizacaoDaOcorrencia 
         Height          =   315
         Left            =   1050
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   450
         Width           =   4695
      End
      Begin VB.CheckBox chkbytRemido 
         Caption         =   "Remido"
         Height          =   225
         Left            =   4890
         TabIndex        =   2
         Top             =   930
         Width           =   885
      End
      Begin TrueOleDBGrid70.TDBGrid tdb_Ocorrencia 
         Height          =   3675
         Left            =   150
         TabIndex        =   8
         Top             =   1680
         Width           =   5595
         _ExtentX        =   9869
         _ExtentY        =   6482
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
         Columns(3).Caption=   "Remido"
         Columns(3).DataField=   "bytRemido"
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
         Splits(0)._ColumnProps(13)=   "Column(2).Width=6615"
         Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=6535"
         Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(17)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(18)=   "Column(3).Width=1164"
         Splits(0)._ColumnProps(19)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(20)=   "Column(3)._WidthInPix=1085"
         Splits(0)._ColumnProps(21)=   "Column(3)._EditAlways=0"
         Splits(0)._ColumnProps(22)=   "Column(3).Order=4"
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
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=101,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
         _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
         _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34"
         _StyleDefs(11)  =   "FooterStyle:id=3,.parent=1,.namedParent=35"
         _StyleDefs(12)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(13)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
         _StyleDefs(14)  =   "EditorStyle:id=7,.parent=1"
         _StyleDefs(15)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
         _StyleDefs(16)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
         _StyleDefs(17)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
         _StyleDefs(18)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(19)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
         _StyleDefs(20)  =   "Splits(0).Style:id=13,.parent=1"
         _StyleDefs(21)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
         _StyleDefs(22)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
         _StyleDefs(23)  =   "Splits(0).FooterStyle:id=15,.parent=3"
         _StyleDefs(24)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
         _StyleDefs(25)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
         _StyleDefs(26)  =   "Splits(0).EditorStyle:id=17,.parent=7"
         _StyleDefs(27)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
         _StyleDefs(28)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
         _StyleDefs(29)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
         _StyleDefs(30)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(31)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(32)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
         _StyleDefs(33)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
         _StyleDefs(34)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
         _StyleDefs(35)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
         _StyleDefs(36)  =   "Splits(0).Columns(1).Style:id=46,.parent=13"
         _StyleDefs(37)  =   "Splits(0).Columns(1).HeadingStyle:id=43,.parent=14"
         _StyleDefs(38)  =   "Splits(0).Columns(1).FooterStyle:id=44,.parent=15"
         _StyleDefs(39)  =   "Splits(0).Columns(1).EditorStyle:id=45,.parent=17"
         _StyleDefs(40)  =   "Splits(0).Columns(2).Style:id=32,.parent=13"
         _StyleDefs(41)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
         _StyleDefs(44)  =   "Splits(0).Columns(3).Style:id=50,.parent=13"
         _StyleDefs(45)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
         _StyleDefs(46)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
         _StyleDefs(47)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
         _StyleDefs(48)  =   "Named:id=33:Normal"
         _StyleDefs(49)  =   ":id=33,.parent=0"
         _StyleDefs(50)  =   "Named:id=34:Heading"
         _StyleDefs(51)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(52)  =   ":id=34,.wraptext=-1"
         _StyleDefs(53)  =   "Named:id=35:Footing"
         _StyleDefs(54)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(55)  =   "Named:id=36:Selected"
         _StyleDefs(56)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(57)  =   "Named:id=37:Caption"
         _StyleDefs(58)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(59)  =   "Named:id=38:HighlightRow"
         _StyleDefs(60)  =   ":id=38,.parent=33,.bgcolor=&H80000014&,.fgcolor=&H80000012&"
         _StyleDefs(61)  =   "Named:id=39:EvenRow"
         _StyleDefs(62)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(63)  =   "Named:id=40:OddRow"
         _StyleDefs(64)  =   ":id=40,.parent=33"
         _StyleDefs(65)  =   "Named:id=41:RecordSelector"
         _StyleDefs(66)  =   ":id=41,.parent=34"
         _StyleDefs(67)  =   "Named:id=42:FilterBar"
         _StyleDefs(68)  =   ":id=42,.parent=33"
      End
      Begin VB.TextBox txtintCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1050
         MaxLength       =   8
         MultiLine       =   -1  'True
         TabIndex        =   1
         Top             =   870
         Width           =   1125
      End
      Begin VB.TextBox txtstrDescricao 
         Height          =   285
         Left            =   1050
         MaxLength       =   100
         TabIndex        =   3
         Top             =   1230
         Width           =   4695
      End
      Begin VB.Label lblintCodigo 
         AutoSize        =   -1  'True
         Caption         =   "Código"
         Height          =   195
         Left            =   465
         TabIndex        =   7
         Top             =   960
         Width           =   495
      End
      Begin VB.Label lblstrDescricao 
         AutoSize        =   -1  'True
         Caption         =   "Descrição"
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   1320
         Width           =   720
      End
      Begin VB.Label lblintUtilizacaoDaOcorrencia 
         AutoSize        =   -1  'True
         Caption         =   "Utilização"
         Height          =   195
         Left            =   270
         TabIndex        =   5
         Top             =   585
         Width           =   690
      End
   End
End
Attribute VB_Name = "frmCadOcorrencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim mblnAlterando      As Boolean
    Dim mobjAux            As Object
    Dim mlngUltimo         As Long
    Dim mblnGuardaUltimo   As Boolean
    Dim mblnSelecionou     As Boolean
    Dim mblnPrimeiraVez    As Boolean
    Dim adoResultado       As ADODB.Recordset
    Dim strUtilizacaoAtual As String
    Dim strCodigoAtual     As String
    Dim strDescricaoAtual  As String
    Dim bytOrdenacao       As Byte
    Dim blnOrdenacaoAsc    As Boolean

    
Function blnQuerryDuplicataCodigo(strCodigo As String) As Boolean
Dim strSQL As String
blnQuerryDuplicataCodigo = False
    strSQL = ""
    strSQL = strSQL & "SELECT count(*) as Contador "
    strSQL = strSQL & "FROM " & gstrOcorrencia & " WHERE intCodigo= '" & strCodigo & "' "
    strSQL = strSQL & "AND intUtilizacaodaOcorrencia = " & dbcintUtilizacaoDaOcorrencia.ListIndex & " "
    If mblnAlterando Then
        strSQL = strSQL & "AND PKID <> " & txtPKId
    End If
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
        If Not adoResultado.EOF Then
            If adoResultado!Contador <= 0 Then
                blnQuerryDuplicataCodigo = False
                Exit Function
            Else
                ExibeMensagem "O código desta Ocorrência já existe para esta Utilização"
                blnQuerryDuplicataCodigo = True
            End If
        End If
    End If
End Function

Private Sub chkbytRemido_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", chkbytRemido
End Sub

Private Sub dbcintUtilizacaoDaOcorrencia_Click()
'Exit Sub
        'If mblnGuardaUltimo = False Then
        '    mlngUltimo = dbcintUtilizacaoDaOcorrencia.ListIndex
        '    Limpa_Controles Me, True, True, False, False, False
        'End If
        'mblnPrimeiraVez = False
        'LeDaTabelaParaObj gstrOcorrencia, tdb_Ocorrencia, strQueryTabelaOcorrencia
End Sub

Private Sub dbcintUtilizacaoDaOcorrencia_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbcintUtilizacaoDaOcorrencia
End Sub

Private Sub Form_Activate()
    gintCodSeguranca = 588
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
    mblnAlterando = False
    
    dbcintUtilizacaoDaOcorrencia.AddItem "Cálculos"
    dbcintUtilizacaoDaOcorrencia.ItemData(dbcintUtilizacaoDaOcorrencia.NewIndex) = "1"
    dbcintUtilizacaoDaOcorrencia.AddItem "Baixa"
    dbcintUtilizacaoDaOcorrencia.ItemData(dbcintUtilizacaoDaOcorrencia.NewIndex) = "2"
    dbcintUtilizacaoDaOcorrencia.AddItem "Entrega"
    dbcintUtilizacaoDaOcorrencia.ItemData(dbcintUtilizacaoDaOcorrencia.NewIndex) = "3"
    dbcintUtilizacaoDaOcorrencia.AddItem "Dívida Ativa"
    dbcintUtilizacaoDaOcorrencia.ItemData(dbcintUtilizacaoDaOcorrencia.NewIndex) = "4"
    dbcintUtilizacaoDaOcorrencia.AddItem "Econômicas"
    dbcintUtilizacaoDaOcorrencia.ItemData(dbcintUtilizacaoDaOcorrencia.NewIndex) = "5"
    dbcintUtilizacaoDaOcorrencia.AddItem "Imobiliárias"
    dbcintUtilizacaoDaOcorrencia.ItemData(dbcintUtilizacaoDaOcorrencia.NewIndex) = "6"
    dbcintUtilizacaoDaOcorrencia.AddItem "Suspensão de Exigências"
    dbcintUtilizacaoDaOcorrencia.ItemData(dbcintUtilizacaoDaOcorrencia.NewIndex) = "7"
    dbcintUtilizacaoDaOcorrencia.AddItem "Protocolos Diversos"
    dbcintUtilizacaoDaOcorrencia.ItemData(dbcintUtilizacaoDaOcorrencia.NewIndex) = "8"
    dbcintUtilizacaoDaOcorrencia.AddItem "Ações Fiscais"
    dbcintUtilizacaoDaOcorrencia.ItemData(dbcintUtilizacaoDaOcorrencia.NewIndex) = "9"
    dbcintUtilizacaoDaOcorrencia.AddItem "Ações Judiciais"
    dbcintUtilizacaoDaOcorrencia.ItemData(dbcintUtilizacaoDaOcorrencia.NewIndex) = "10"
    
    VerificaObjParaAplicar mobjAux
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar, gstrDeletar
    mblnSelecionou = False
    mblnPrimeiraVez = False
End Sub

Private Sub tdb_Ocorrencia_Click()
    mblnPrimeiraVez = True
    If glngQtdLinhaTDBGrid(tdb_Ocorrencia) = 1 Then
        tdb_Ocorrencia_RowColChange 0, 0
    End If
End Sub

Private Sub tdb_Ocorrencia_DblClick()
    MantemForm gstrAplicar
End Sub

Private Sub tdb_Ocorrencia_HeadClick(ByVal ColIndex As Integer)
    gOrdenaGrid tdb_Ocorrencia, ColIndex
    'blnOrdenacaoAsc = IIf(bytOrdenacao = ColIndex, Not blnOrdenacaoAsc, True)
    'bytOrdenacao = ColIndex: MantemForm gstrRefresh
End Sub

Private Sub tdb_Ocorrencia_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", tdb_Ocorrencia
End Sub

Private Sub txtstrDescricao_GotFocus()
    MarcaCampo txtstrDescricao
End Sub

Private Sub txtstrDescricao_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrDescricao
End Sub

Private Sub tdb_Ocorrencia_FilterChange()
    mblnPrimeiraVez = False
    gblnFilraCampos tdb_Ocorrencia
End Sub

Private Sub tdb_Ocorrencia_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    With tdb_Ocorrencia
        If Not .EOF And Not .BOF Then
            If mblnPrimeiraVez Then
                mblnAlterando = True
                txtPKId.Text = .Columns("PKID").Value
                LeDaTabelaParaObj gstrOcorrencia, Me
                gCorLinhaSelecionada tdb_Ocorrencia
                HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrDeletar
                If mobjAux Is Nothing Then
                    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar
                Else
                    HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrAplicar
                End If
                strUtilizacaoAtual = gstrItemData(dbcintUtilizacaoDaOcorrencia)
                strCodigoAtual = tdb_Ocorrencia.Columns("intcodigo").Value
                strDescricaoAtual = tdb_Ocorrencia.Columns("strDescricao").Value
                mblnSelecionou = True
            End If
        End If
    End With
End Sub

Public Sub MantemForm(ByVal strModoOperacao As String)
Dim varBookMark As Variant
Dim strSQL As String
    
    mblnGuardaUltimo = True
    
'    If UCase(strModoOperacao) = "SALVAR" Then
'        If blnQuerryDuplicataCodigo(txtintCodigo) Then Exit Sub
'    End If
    
    strSQL = strQuery
    
    If UCase(strModoOperacao) = "SALVAR" Or UCase(strModoOperacao) = "DELETAR" Then
        mblnPrimeiraVez = False
    End If
    
    If UCase(strModoOperacao) = "DELETAR" Then
       ToolBarGeral strModoOperacao, gstrOcorrencia, mblnAlterando, tdb_Ocorrencia, Me, mobjAux, strQueryTabelaOcorrencia, strQueryAplicar, rptOcorrencias, strQuerryRelatorio
       MantemForm gstrRefresh
       Exit Sub
    End If
    
    If UCase(strModoOperacao) = "SALVAR" Then
        If blnDadosOk = False Then Exit Sub
        mblnPrimeiraVez = False
        ToolBarGeral strModoOperacao, gstrOcorrencia, mblnAlterando, tdb_Ocorrencia, Me, mobjAux, strSQL, strQueryAplicar, rptOcorrencias, strQuerryRelatorio
        MantemForm gstrRefresh
        Exit Sub
    End If
    
    ToolBarGeral strModoOperacao, gstrOcorrencia, mblnAlterando, tdb_Ocorrencia, Me, mobjAux, strQueryTabelaOcorrencia, strQueryAplicar, rptOcorrencias, strQuerryRelatorio
    
    If UCase(strModoOperacao) <> gstrAplicar Then
       HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrDeletar
    End If
        
    'dbcintUtilizacaoDaOcorrencia.ListIndex = mlngUltimo
    mblnGuardaUltimo = False

End Sub

Private Function strQuery() As String
Dim strSQL As String

    strSQL = ""
    strSQL = strSQL & "SELECT PKId, intCodigo, strDescricao, " & gstrCASEWHEN("bytRemido", "0, 'Não', 1, 'Sim'") & " as bytRemido "
    strSQL = strSQL & "FROM " & gstrOcorrencia & " "
    strSQL = strSQL & "ORDER BY strDescricao "
    strQuery = strSQL
End Function

Private Function strQueryTabelaOcorrencia() As String
    Dim strSQL As String
    
'txtstrDescricao = ""
'txtintCodigo = ""
'chkbytRemido.Value = 0
    strSQL = ""
'    strSql = strSql & "Select PKId, intCodigo, strDescricao, CASE bytRemido WHEN 0 THEN 'Não' WHEN 1 THEN 'Sim' END as bytRemido "
    strSQL = strSQL & "Select PKId, intCodigo, strDescricao, " & gstrCASEWHEN("bytRemido", "0, 'Não', 1, 'Sim'") & " as bytRemido "
    strSQL = strSQL & "From " & gstrOcorrencia & " "
    If Trim(dbcintUtilizacaoDaOcorrencia.Text) <> "" Then
       strSQL = strSQL & "WHERE intUtilizacaoDaOcorrencia = " & gstrItemData(dbcintUtilizacaoDaOcorrencia) & " "
    End If
    strSQL = strSQL & "ORDER BY strDescricao "
strQueryTabelaOcorrencia = strSQL
End Function

Private Sub txtintCodigo_GotFocus()
    MarcaCampo txtintCodigo
End Sub

Private Sub txtintCodigo_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtintCodigo
End Sub

Function strQuerryRelatorio() As String
    
Dim strSQL As String
    strSQL = ""
    strSQL = strSQL & " SELECT "
    
    strSQL = strSQL & gstrCASEWHEN("intUtilizacaoDaOcorrencia", _
        "1, 'Cálculos', " & _
        "2, 'Baixa', " & _
        "3, 'Entrega', " & _
        "4, 'Dívida Ativa', " & _
        "5, 'Econômicas', " & _
        "6, 'Imobiliárias', " & _
        "7, 'Suspensão de Exigências', " & _
        "8, 'Protocolos Diversos', " & _
        "9, 'Ações Fiscais', " & _
        "10, 'Ações Judiciais'") & " AS Utilizacao, "
    strSQL = strSQL & " intUtilizacaoDaOcorrencia as CODUTILIZACAO, intCodigo AS Codigo, strDescricao AS Ocorrencia "
    
    strSQL = strSQL & " FROM " & gstrOcorrencia

    strSQL = strSQL & " ORDER BY intUtilizacaoDaOcorrencia, Utilizacao "
strQuerryRelatorio = strSQL
End Function


Private Function strQueryAplicar() As String
    Dim strSQL  As String
    strSQL = ""
    strSQL = strSQL & " SELECT PKId, strDescricao FROM "
    strSQL = strSQL & gstrOcorrencia & " ORDER BY strDescricao"
    strQueryAplicar = strSQL
End Function

Private Function blnDadosOk() As Boolean
    blnDadosOk = False
    If gstrItemData(dbcintUtilizacaoDaOcorrencia) = 0 Then
        ExibeMensagem "O utilização deve ser preenchida corretamente."
        Exit Function
    ElseIf Trim(txtintCodigo) = "" Then
        ExibeMensagem "O código tem que ser digitado."
        Exit Function
    ElseIf Trim(txtstrDescricao) = "" Then
        ExibeMensagem "A descrição tem que ser digitada."
        Exit Function
    End If
    
    If Not mblnAlterando Or (mblnAlterando And UCase$(strCodigoAtual) <> UCase$(txtintCodigo.Text)) Or (mblnAlterando And strUtilizacaoAtual <> gstrItemData(dbcintUtilizacaoDaOcorrencia)) Then
        If gblnExisteCodigo(2, gstrOcorrencia, "intutilizacaodaocorrencia", gstrItemData(dbcintUtilizacaoDaOcorrencia), "intCodigo", txtintCodigo.Text) Then
            ExibeMensagem "O código desta Ocorrência já existe para esta Utilização"
            Exit Function
        End If
    End If
    
    If Not mblnAlterando Or (mblnAlterando And UCase$(txtstrDescricao.Text) <> UCase$(strDescricaoAtual)) Or (mblnAlterando And (strUtilizacaoAtual <> gstrItemData(dbcintUtilizacaoDaOcorrencia))) Then
        If gblnExisteCodigo(2, gstrOcorrencia, "intutilizacaodaocorrencia", gstrItemData(dbcintUtilizacaoDaOcorrencia), "strDescricao", "'" & txtstrDescricao & "'") Then
            ExibeMensagem "A descrição desta Ocorrência já existe para esta Utilização"
            Exit Function
        End If
    End If
    
    blnDadosOk = True
End Function



