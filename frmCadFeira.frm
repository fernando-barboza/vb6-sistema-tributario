VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "Tabctl32.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frmCadFeira 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Feiras"
   ClientHeight    =   4245
   ClientLeft      =   2805
   ClientTop       =   3390
   ClientWidth     =   6285
   Icon            =   "frmCadFeira.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   6285
   Begin TabDlg.SSTab tab_3dPasta 
      Height          =   4155
      Left            =   60
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   60
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   7329
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Feira"
      TabPicture(0)   =   "frmCadFeira.frx":1042
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblstrDescricao"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "tdb_Lista"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txtstrdescricao"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtPKId"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cbointDiaDaSemana"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      Begin VB.ComboBox cbointDiaDaSemana 
         Height          =   315
         Left            =   1560
         TabIndex        =   6
         Top             =   600
         Width           =   1695
      End
      Begin VB.TextBox txtPKId 
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         OLEDragMode     =   1  'Automatic
         OLEDropMode     =   1  'Manual
         TabIndex        =   2
         Top             =   0
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtstrdescricao 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#.##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   1
         EndProperty
         Height          =   315
         Left            =   1560
         MaxLength       =   30
         TabIndex        =   1
         Top             =   960
         Width           =   4305
      End
      Begin TrueOleDBGrid70.TDBGrid tdb_Lista 
         Height          =   2475
         Left            =   120
         TabIndex        =   3
         Top             =   1455
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   4366
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
         Columns(1).Caption=   "Descrição"
         Columns(1).DataField=   "strDescricao"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   16
         Columns(2)._MaxComboItems=   5
         Columns(2).ValueItems(0)._DefaultItem=   0
         Columns(2).ValueItems(0).Value=   "1"
         Columns(2).ValueItems(0).Value.vt=   8
         Columns(2).ValueItems(0).DisplayValue=   "Domingo"
         Columns(2).ValueItems(0).DisplayValue.vt=   8
         Columns(2).ValueItems(0)._PropDict=   "_DefaultItem,517,2"
         Columns(2).ValueItems(1)._DefaultItem=   0
         Columns(2).ValueItems(1).Value=   "2"
         Columns(2).ValueItems(1).Value.vt=   8
         Columns(2).ValueItems(1).DisplayValue=   "Segunda"
         Columns(2).ValueItems(1).DisplayValue.vt=   8
         Columns(2).ValueItems(1)._PropDict=   "_DefaultItem,517,2"
         Columns(2).ValueItems(2)._DefaultItem=   0
         Columns(2).ValueItems(2).Value=   "3"
         Columns(2).ValueItems(2).Value.vt=   8
         Columns(2).ValueItems(2).DisplayValue=   "Terça"
         Columns(2).ValueItems(2).DisplayValue.vt=   8
         Columns(2).ValueItems(2)._PropDict=   "_DefaultItem,517,2"
         Columns(2).ValueItems(3)._DefaultItem=   0
         Columns(2).ValueItems(3).Value=   "4"
         Columns(2).ValueItems(3).Value.vt=   8
         Columns(2).ValueItems(3).DisplayValue=   "Quarta"
         Columns(2).ValueItems(3).DisplayValue.vt=   8
         Columns(2).ValueItems(3)._PropDict=   "_DefaultItem,517,2"
         Columns(2).ValueItems(4)._DefaultItem=   0
         Columns(2).ValueItems(4).Value=   "5"
         Columns(2).ValueItems(4).Value.vt=   8
         Columns(2).ValueItems(4).DisplayValue=   "Quinta"
         Columns(2).ValueItems(4).DisplayValue.vt=   8
         Columns(2).ValueItems(4)._PropDict=   "_DefaultItem,517,2"
         Columns(2).ValueItems(5)._DefaultItem=   0
         Columns(2).ValueItems(5).Value=   "6"
         Columns(2).ValueItems(5).Value.vt=   8
         Columns(2).ValueItems(5).DisplayValue=   "Sexta"
         Columns(2).ValueItems(5).DisplayValue.vt=   8
         Columns(2).ValueItems(5)._PropDict=   "_DefaultItem,517,2"
         Columns(2).ValueItems(6)._DefaultItem=   0
         Columns(2).ValueItems(6).Value=   "7"
         Columns(2).ValueItems(6).Value.vt=   8
         Columns(2).ValueItems(6).DisplayValue=   "Sabado"
         Columns(2).ValueItems(6).DisplayValue.vt=   8
         Columns(2).ValueItems(6)._PropDict=   "_DefaultItem,517,2"
         Columns(2).ValueItems.Count=   7
         Columns(2).Caption=   "Dia da semana"
         Columns(2).DataField=   "INTDIADASEMANA"
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
         Splits(0)._ColumnProps(8)=   "Column(1).Width=6906"
         Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=6826"
         Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(13)=   "Column(2).Width=2725"
         Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=2646"
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
         _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
         _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
         _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=28,.parent=13"
         _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=25,.parent=14"
         _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=26,.parent=15"
         _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=27,.parent=17"
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
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Dia Da Semana:"
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   660
         Width           =   1170
      End
      Begin VB.Label lblstrDescricao 
         AutoSize        =   -1  'True
         Caption         =   "Descrição:"
         Height          =   195
         Left            =   600
         TabIndex        =   4
         Top             =   1050
         Width           =   765
      End
   End
End
Attribute VB_Name = "frmCadFeira"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim mblnAlterando     As Boolean
    Dim mobjAux           As Object
    Dim mblnSelecionou    As Boolean
    Dim mblnClickOk       As Boolean
    Dim mblnPrimeiraVez   As Boolean
    Dim bytOrdenacao      As Byte
    Dim blnOrdenacaoAsc   As Boolean
    Dim strDescricaoAtual As String

Private Function strQuery() As String
    
    Dim strSQL  As String
    
    strSQL = ""
    
    strSQL = strSQL & " SELECT Pkid,strDescricao, intdiadasemana FROM "
    
    strSQL = strSQL & gstrFeira
    
    Select Case bytOrdenacao
    
    Case Is = 1
        strSQL = strSQL & " ORDER BY strDescricao" & IIf(blnOrdenacaoAsc, " ASC", " DESC")
    End Select
    
    strQuery = strSQL
    
End Function

Private Function strQueryAplicar() As String
    Dim strSQL  As String
    
    strSQL = ""
    strSQL = strSQL & " SELECT PKId, strDescricao, intdiadasemana FROM "
    strSQL = strSQL & gstrFeira & " ORDER BY strDescricao, intDiaDaSemana"
    
    strQueryAplicar = strSQL
    
End Function

Private Sub Form_Activate()
    gintCodSeguranca = 1159
    VerificaListaAutomatica gstrFeira, tdb_Lista, ""
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

    bytOrdenacao = 3: blnOrdenacaoAsc = True
    VerificaObjParaAplicar mobjAux
    
    PreencheDiasDaSemana

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar
    mblnSelecionou = False
End Sub

Private Sub tdb_Lista_Click()
    mblnClickOk = True
    mblnPrimeiraVez = True
End Sub

Sub tdb_Lista_DblClick()
    MantemForm gstrAplicar
End Sub

Private Sub tdb_Lista_FilterChange()
    mblnPrimeiraVez = False
    mblnClickOk = False
    gblnFilraCampos tdb_Lista
End Sub

Private Sub tdb_Lista_HeadClick(ByVal ColIndex As Integer)
   
'   blnOrdenacaoAsc = IIf(bytOrdenacao = ColIndex, Not blnOrdenacaoAsc, True)
'   bytOrdenacao = ColIndex: MantemForm gstrRefresh
    gOrdenaGrid tdb_Lista, ColIndex
   mblnClickOk = True
   
End Sub

Private Sub tdb_Lista_KeyPress(KeyAscii As Integer)
    Select Case tdb_Lista.Col
        Case 1
            CaracterValido KeyAscii, "A", tdb_Lista
    End Select
End Sub

Private Sub tdb_Lista_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    With tdb_Lista
        If (Not .EOF And Not .BOF) And mblnClickOk Then
            mblnClickOk = False
            txtPKId.Text = .Columns("PKID").Value
            If mblnPrimeiraVez Then
                LeDaTabelaParaObj gstrFeira, Me
                gCorLinhaSelecionada tdb_Lista
                If mobjAux Is Nothing Then
                    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar
                Else
                    HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrAplicar
                End If
                strDescricaoAtual = tdb_Lista.Columns("strDescricao").Value
                cbointDiaDaSemana.ListIndex = Val(tdb_Lista.Columns("intDiaDaSemana").Value) - 1
                mblnSelecionou = True
                mblnAlterando = True
            End If
        End If
    End With
End Sub

Public Sub MantemForm(ByVal strModoOperacao As String)
Dim strSQL As String
If strModoOperacao = UCase("IMPRIMIR") Then
    strSQL = strQueryRelatorio
    ToolBarGeral strModoOperacao, gstrFeira, mblnAlterando, tdb_Lista, Me, mobjAux, strSQL, , rptFeira, strQueryRelatorio
    Exit Sub
End If
    Select Case UCase(strModoOperacao)
        Case gstrSalvar
            If Not blnDadosOK Then Exit Sub
            SalvarFeira
            MantemForm gstrRefresh
            MantemForm gstrNovo
        Case gstrPreencherLista
            If Me.ActiveControl.Name = "cbointDiaDaSemana" Then
                PreencheDiasDaSemana
            End If
        Case Else
            ToolBarGeral strModoOperacao, gstrFeira, mblnAlterando, tdb_Lista, _
                 Me, mobjAux, strQuery, strQueryAplicar, rptBairro, strQueryRelatorio
    End Select
                 
End Sub

Private Sub txtstrDescricao_GotFocus()
    MarcaCampo txtstrdescricao
End Sub

Private Sub txtstrDescricao_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrdescricao
End Sub

Private Sub txtPKId_GotFocus()
    MarcaCampo txtPKId
End Sub

Private Sub txtPKId_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Function strQueryRelatorio() As String
    Dim strSQL As String
    
    strSQL = ""
    
    strSQL = "SELECT strdescricao FROM " & gstrFeira
    
    strQueryRelatorio = strSQL
   
End Function

Private Function blnDadosOK()
    
    blnDadosOK = False
    If Not mblnAlterando Or (mblnAlterando And UCase$(txtstrdescricao.Text) <> UCase$(strDescricaoAtual)) Then
        If gblnExisteCodigo(1, gstrFeira, "strDescricao", "'" & txtstrdescricao.Text & "'") Then
            ExibeMensagem "A descrição informada já se encontra cadastrada."
            txtstrdescricao.SetFocus
            Exit Function
        End If
    End If
    
    If Trim(txtstrdescricao.Text) = "" Then
        ExibeMensagem "O campo Descrição é obrigatório"
        txtstrdescricao.SetFocus
        Exit Function
    End If
    
    If cbointDiaDaSemana.ListIndex = -1 Then
        ExibeMensagem "O campo Dia da Semana é obrigatório"
        cbointDiaDaSemana.SetFocus
        Exit Function
    End If
        
    blnDadosOK = True
    
End Function

Private Sub PreencheDiasDaSemana()
    cbointDiaDaSemana.Clear
    cbointDiaDaSemana.AddItem "Domingo"
    cbointDiaDaSemana.ItemData(cbointDiaDaSemana.NewIndex) = DOMINGO
    cbointDiaDaSemana.AddItem "Segunda-Feira"
    cbointDiaDaSemana.ItemData(cbointDiaDaSemana.NewIndex) = SEGUNDA
    cbointDiaDaSemana.AddItem "Terça-Feira"
    cbointDiaDaSemana.ItemData(cbointDiaDaSemana.NewIndex) = TERCA
    cbointDiaDaSemana.AddItem "Quarta-Feira"
    cbointDiaDaSemana.ItemData(cbointDiaDaSemana.NewIndex) = QUARTA
    cbointDiaDaSemana.AddItem "Quinta-Feira"
    cbointDiaDaSemana.ItemData(cbointDiaDaSemana.NewIndex) = QUINTA
    cbointDiaDaSemana.AddItem "Sexta-Feira"
    cbointDiaDaSemana.ItemData(cbointDiaDaSemana.NewIndex) = SEXTA
    cbointDiaDaSemana.AddItem "Sabado"
    cbointDiaDaSemana.ItemData(cbointDiaDaSemana.NewIndex) = SABADO
End Sub

Private Sub SalvarFeira()
    
    Dim strSQL As String
    Dim adoRec As New ADODB.Recordset
    
    If Not mblnAlterando Then
        strSQL = "INSERT INTO " & gstrFeira & "(" & _
                         "strDescricao, " & _
                         "dtmdtAtualizacao, " & _
                         "lngCodUsr, " & _
                         "Intdiadasemana)" & _
                 "VALUES ('" & txtstrdescricao & "'," & strGETDATE & "," & glngCodUsr & ", " & cbointDiaDaSemana.ItemData(cbointDiaDaSemana.ListIndex) & ")"
                 
        Set gobjBanco = New clsBanco
        
        gobjBanco.Execute (strSQL)
    Else
        
        If gblnExclusaoGravacaoOk("", "Tem certeza que deseja alterar") Then
        
            strSQL = "UPDATE " & gstrFeira & " SET " & _
                                "strDescricao = '" & txtstrdescricao.Text & "', " & _
                                "dtmdtAtualizacao = " & strGETDATE & ", " & _
                                "lngcodusr = " & glngCodUsr & ", " & _
                                "intdiadasemana = " & cbointDiaDaSemana.ItemData(cbointDiaDaSemana.ListIndex) & _
                                " WHERE Pkid = " & txtPKId.Text
        End If
                            
        Set gobjBanco = New clsBanco
        gobjBanco.Execute (strSQL)
                            
    End If
End Sub
