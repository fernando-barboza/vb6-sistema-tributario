VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "TODG7.OCX"
Begin VB.Form frmCadTipoLogradouro 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tipos de Logradouro"
   ClientHeight    =   4260
   ClientLeft      =   4935
   ClientTop       =   4800
   ClientWidth     =   5940
   HelpContextID   =   24
   Icon            =   "CadTipoLogradouro.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4260
   ScaleWidth      =   5940
   Begin TabDlg.SSTab tab_3dPasta 
      Height          =   4080
      Left            =   90
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   90
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   7197
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Tipos de Logradouro"
      TabPicture(0)   =   "CadTipoLogradouro.frx":1042
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblstrDescricao"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblstrSigla"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblPKID"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txtstrDescricao"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtstrSigla"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtPKId"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "tdb_TipoLogradouro"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      Begin TrueOleDBGrid70.TDBGrid tdb_TipoLogradouro 
         Height          =   2775
         Left            =   120
         TabIndex        =   2
         Top             =   1170
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   4895
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Código"
         Columns(0).DataField=   "PKID"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Descrição"
         Columns(1).DataField=   "strdescricao"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Sigla"
         Columns(2).DataField=   "strsigla"
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
         Splits(0)._ColumnProps(1)=   "Column(0).Width=1455"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1376"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=2"
         Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(7)=   "Column(1).Width=5953"
         Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=5874"
         Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(12)=   "Column(2).Width=1826"
         Splits(0)._ColumnProps(13)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(14)=   "Column(2)._WidthInPix=1746"
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
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=236,.bold=0,.fontsize=825,.italic=0"
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
         _StyleDefs(30)  =   "Splits(0).Columns(0).Style:id=28,.parent=13,.alignment=1"
         _StyleDefs(31)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
         _StyleDefs(32)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
         _StyleDefs(33)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
         _StyleDefs(34)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
         _StyleDefs(35)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
         _StyleDefs(36)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
         _StyleDefs(37)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
         _StyleDefs(38)  =   "Splits(0).Columns(2).Style:id=46,.parent=13"
         _StyleDefs(39)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
         _StyleDefs(40)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
         _StyleDefs(41)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
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
      Begin VB.TextBox txtPKId 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1005
         Locked          =   -1  'True
         OLEDragMode     =   1  'Automatic
         OLEDropMode     =   1  'Manual
         TabIndex        =   4
         Top             =   450
         Width           =   885
      End
      Begin VB.TextBox txtstrSigla 
         Height          =   285
         Left            =   4935
         MaxLength       =   3
         TabIndex        =   1
         Top             =   810
         Width           =   660
      End
      Begin VB.TextBox txtstrDescricao 
         Height          =   285
         Left            =   1005
         MaxLength       =   15
         TabIndex        =   0
         Top             =   810
         Width           =   3405
      End
      Begin VB.Label lblPKID 
         AutoSize        =   -1  'True
         Caption         =   "Código"
         Height          =   195
         Left            =   420
         TabIndex        =   7
         Top             =   495
         Width           =   495
      End
      Begin VB.Label lblstrSigla 
         AutoSize        =   -1  'True
         Caption         =   "Sigla"
         Height          =   195
         Left            =   4500
         TabIndex        =   6
         Top             =   855
         Width           =   345
      End
      Begin VB.Label lblstrDescricao 
         AutoSize        =   -1  'True
         Caption         =   "Descrição"
         Height          =   195
         Left            =   195
         TabIndex        =   5
         Top             =   855
         Width           =   720
      End
   End
End
Attribute VB_Name = "frmCadTipoLogradouro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mblnAlterando   As Boolean
Dim mobjAux         As Object
    
Dim mblnSelecionou  As Boolean
Dim mblnPrimeiraVez As Boolean
    
Private Sub Form_Activate()
    gintCodSeguranca = 582
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
    
    VerificaListaAutomatica "", tdb_TipoLogradouro, gstrQueryTipoLogradouro(True)
    VerificaObjParaAplicar mobjAux
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar, gstrDeletar
    mblnSelecionou = False
    mblnPrimeiraVez = False
End Sub

Private Sub tdb_TipoLogradouro_Click()
    mblnPrimeiraVez = True
    If glngQtdLinhaTDBGrid(tdb_TipoLogradouro) = 1 Then
        tdb_TipoLogradouro_RowColChange 0, 0
    End If
End Sub

Private Sub tdb_TipoLogradouro_DblClick()
    MantemForm gstrAplicar
End Sub

Private Sub tdb_TipoLogradouro_FilterChange()
    mblnPrimeiraVez = False
    gblnFilraCampos tdb_TipoLogradouro
End Sub

Private Sub tdb_TipoLogradouro_HeadClick(ByVal ColIndex As Integer)
    gOrdenaGrid tdb_TipoLogradouro, ColIndex
End Sub

Private Sub tdb_TipoLogradouro_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    With tdb_TipoLogradouro
        If Not .EOF And Not .BOF Then
            If mblnPrimeiraVez Then
                txtPKId.Text = .Columns("PKID").Value
                mblnAlterando = True
                LeDaTabelaParaObj gstrTipoLogradouro, Me
                gCorLinhaSelecionada tdb_TipoLogradouro
                
                HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrDeletar
                
                If mobjAux Is Nothing Then
                    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar
                Else
                    HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrAplicar
                End If
                mblnSelecionou = True
            End If
        End If
    End With
End Sub

Private Function strQueryAplicar() As String
Dim strSql As String
    
    strSql = ""
    strSql = strSql & "SELECT PKId, strSigla "
    strSql = strSql & " FROM " & gstrTipoLogradouro
    strSql = strSql & " ORDER BY strSigla"
    
    strQueryAplicar = strSql
    
End Function

Private Function strQuery() As String
Dim strSql1 As String
    
    strSql1 = ""
    strSql1 = strSql1 & "SELECT * "
    strSql1 = strSql1 & "FROM " & gstrTipoLogradouro & " "
    strSql1 = strSql1 & "ORDER BY PKID"
    
    strQuery = strSql1
    
End Function

Public Sub MantemForm(ByVal strModoOperacao As String)
Dim varBookMark As Variant
Dim strSql As String
Dim strSql1 As String

    strSql = strQueryTipoLogradouro
    strSql1 = strQuery
    
    If strModoOperacao = UCase(gstrImprimir) Then
        ToolBarGeral strModoOperacao, gstrTipoLogradouro, mblnAlterando, tdb_TipoLogradouro, Me, mobjAux, strSql1, , rptTipoDeLogradouro, strQuery
        Exit Sub
    End If

    If UCase(strModoOperacao) = "SALVAR" Then
        If Not blnDadosOk Then Exit Sub
        mblnPrimeiraVez = False
    End If
    
    If UCase(strModoOperacao) = "DELETAR" Then
        mblnPrimeiraVez = False
    End If
    
    ToolBarGeral strModoOperacao, gstrTipoLogradouro, mblnAlterando, tdb_TipoLogradouro, Me, mobjAux, strSql, strQueryAplicar, rptTipoDeLogradouro, strQueryRelatorio
    If Not gobjGeral Is Nothing Then HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrDeletar
    
End Sub

Private Sub txtstrDescricao_GotFocus()
    MarcaCampo txtstrDescricao
End Sub

Private Sub txtstrDescricao_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrDescricao
End Sub

Private Sub txtPKId_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub txtstrSigla_GotFocus()
    MarcaCampo txtstrSigla
End Sub

Private Sub txtstrSigla_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrSigla
End Sub

Function strQueryRelatorio() As String
Dim strSql As String
    
    strSql = ""
    strSql = strSql & "SELECT * "
    strSql = strSql & " FROM " & gstrTipoLogradouro
    If mblnAlterando = True Then
        strSql = strSql & " WHERE PKId = " & tdb_TipoLogradouro.Columns("PKId").Value
    End If
    strSql = strSql & " ORDER BY strDescricao "

    strQueryRelatorio = strSql
    
End Function

Private Function blnDadosOk() As Boolean
Dim strSql As String
Dim adoRec As ADODB.Recordset
    
    blnDadosOk = False
    
    If Trim(txtstrDescricao) = "" Then
        ExibeMensagem "A descrição do logradouro deve ser informada corretamente."
        txtstrDescricao.SetFocus
        Exit Function
    ElseIf Trim(txtstrSigla) = "" Then
        ExibeMensagem "A sigla do logradouro deve ser informada corretamente."
        txtstrSigla.SetFocus
        Exit Function
    ElseIf mblnAlterando = False Then
'        strSql = "SELECT *, 'Campo'= CASE UPPER(strDescricao) WHEN '" & UCase(txtstrDescricao) & "' THEN 'descrição' ELSE 'sigla' END"
        strSql = "SELECT " & gstrTipoLogradouro & ".*, (" & gstrCASEWHEN("UPPER(strDescricao)", "'" & UCase(txtstrDescricao) & "','descrição'", "'sigla'") & ") Campo "
        strSql = strSql & " FROM " & gstrTipoLogradouro & " WHERE "
        strSql = strSql & " UPPER(strDescricao)='" & UCase(txtstrDescricao) & "'"
        strSql = strSql & " OR UPPER(strSigla)='" & UCase(txtstrSigla) & "'"
        
        Set gobjBanco = New clsBanco
        If gobjBanco.CriaADO(strSql, 10, adoRec) Then
            If Not (adoRec.EOF Or adoRec.BOF) Then
                ExibeMensagem "Esta " & adoRec("Campo") & " já esta cadastrada."
                Set adoRec = Nothing
                Exit Function
            End If
        End If
    End If
    
    blnDadosOk = True
    
End Function

Private Function strQueryTipoLogradouro() As String
Dim strSql  As String

    strSql = strSql & "SELECT PKId, strDescricao, strSigla "
    strSql = strSql & "FROM " & gstrTipoLogradouro
    strSql = strSql & " ORDER BY strDescricao "
    
    strQueryTipoLogradouro = strSql
    
End Function

