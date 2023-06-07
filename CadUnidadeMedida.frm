VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "TODG7.OCX"
Begin VB.Form frmCadUnidadeMedida 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Unidades de Medida"
   ClientHeight    =   5280
   ClientLeft      =   1365
   ClientTop       =   3270
   ClientWidth     =   6375
   HelpContextID   =   148
   Icon            =   "CadUnidadeMedida.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5280
   ScaleWidth      =   6375
   Begin TabDlg.SSTab tab_3dPasta 
      Height          =   5145
      Left            =   60
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   60
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   9075
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Unidades de Medida"
      TabPicture(0)   =   "CadUnidadeMedida.frx":1042
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblstrDescricao"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblstrAbreviatura"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblPKID"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lbl_dtmInativo"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtstrDescricao"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtstrAbreviatura"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtPKId"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "tdb_UnidadeMedida"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtdtmInativo"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).ControlCount=   9
      Begin VB.TextBox txtdtmInativo 
         Height          =   285
         Left            =   4950
         TabIndex        =   2
         Top             =   840
         Width           =   1125
      End
      Begin TrueOleDBGrid70.TDBGrid tdb_UnidadeMedida 
         Height          =   3405
         Left            =   135
         TabIndex        =   4
         Top             =   1560
         Width           =   5985
         _ExtentX        =   10557
         _ExtentY        =   6006
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
         Columns(2).DataField=   "strAbreviatura"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Inativo"
         Columns(3).DataField=   "dtmInativo"
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
         Splits(0)._ColumnProps(1)=   "Column(0).Width=1588"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1508"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(1).Width=5318"
         Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=5239"
         Splits(0)._ColumnProps(9)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(11)=   "Column(2).Width=1376"
         Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=1296"
         Splits(0)._ColumnProps(14)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(16)=   "Column(3).Width=1720"
         Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=1640"
         Splits(0)._ColumnProps(19)=   "Column(3)._EditAlways=0"
         Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
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
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=192,.bold=0,.fontsize=825,.italic=0"
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
         _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
         _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
         _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
         _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
         _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
         _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
         _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=46,.parent=13"
         _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
         _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
         _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
         _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=50,.parent=13"
         _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
         _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
         _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
         _StyleDefs(52)  =   "Named:id=33:Normal"
         _StyleDefs(53)  =   ":id=33,.parent=0"
         _StyleDefs(54)  =   "Named:id=34:Heading"
         _StyleDefs(55)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(56)  =   ":id=34,.wraptext=-1"
         _StyleDefs(57)  =   "Named:id=35:Footing"
         _StyleDefs(58)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(59)  =   "Named:id=36:Selected"
         _StyleDefs(60)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(61)  =   "Named:id=37:Caption"
         _StyleDefs(62)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(63)  =   "Named:id=38:HighlightRow"
         _StyleDefs(64)  =   ":id=38,.parent=33,.bgcolor=&H80000014&,.fgcolor=&H80000012&"
         _StyleDefs(65)  =   "Named:id=39:EvenRow"
         _StyleDefs(66)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(67)  =   "Named:id=40:OddRow"
         _StyleDefs(68)  =   ":id=40,.parent=33"
         _StyleDefs(69)  =   "Named:id=41:RecordSelector"
         _StyleDefs(70)  =   ":id=41,.parent=34"
         _StyleDefs(71)  =   "Named:id=42:FilterBar"
         _StyleDefs(72)  =   ":id=42,.parent=33"
      End
      Begin VB.TextBox txtPKId 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1065
         OLEDragMode     =   1  'Automatic
         OLEDropMode     =   1  'Manual
         TabIndex        =   0
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox txtstrAbreviatura 
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
         Left            =   1065
         MaxLength       =   10
         TabIndex        =   1
         Top             =   840
         Width           =   1335
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
         Left            =   1065
         MaxLength       =   30
         TabIndex        =   3
         Top             =   1200
         Width           =   2865
      End
      Begin VB.Label lbl_dtmInativo 
         AutoSize        =   -1  'True
         Caption         =   "Inativo"
         Height          =   195
         Left            =   4380
         TabIndex        =   9
         Top             =   930
         Width           =   480
      End
      Begin VB.Label lblPKID 
         AutoSize        =   -1  'True
         Caption         =   "Código"
         Height          =   195
         Left            =   465
         TabIndex        =   8
         Top             =   570
         Width           =   495
      End
      Begin VB.Label lblstrAbreviatura 
         AutoSize        =   -1  'True
         Caption         =   "Sigla"
         Height          =   195
         Left            =   615
         TabIndex        =   7
         Top             =   930
         Width           =   345
      End
      Begin VB.Label lblstrDescricao 
         AutoSize        =   -1  'True
         Caption         =   "Descrição"
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   1290
         Width           =   720
      End
   End
End
Attribute VB_Name = "frmCadUnidadeMedida"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim mblnAlterando   As Boolean
    Dim mobjAux         As Object
    Dim mblnSelecionou  As Boolean
    Dim mblnPrimeiraVez As Boolean
    Dim blnAlterando    As Boolean
    
    Dim strDescricaoAtual   As String
    Dim strSiglaAtual       As String
    
    Dim bytOrdenacao            As Byte
    Dim blnOrdenacaoAsc         As Boolean
    Dim blnOrdenaGrid           As Boolean

Private Sub Form_Activate()
    gintCodSeguranca = 8
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
    'LeDaTabelaParaObj gstrUnidadeMedida, tdb_UnidadeMedida, "PKId, strDescricao, strAbreviatura"
    VerificaObjParaAplicar mobjAux
    Me.Icon = MDIMenu.Icon
    
    bytOrdenacao = 2: blnOrdenacaoAsc = True
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar, gstrDeletar
    mblnSelecionou = False
    mblnPrimeiraVez = False
End Sub

Private Function strQuery() As String
    Dim strSql  As String
    strSql = ""
    strSql = strSql & " SELECT PKId, strDescricao, strAbreviatura, dtmInativo FROM "
    strSql = strSql & gstrUnidadeMedida & " ORDER BY PKId "
    strQuery = strSql
End Function

Private Sub tdb_UnidadeMedida_Click()
    mblnPrimeiraVez = True
End Sub

Private Sub tdb_UnidadeMedida_DblClick()
    MantemForm gstrAplicar
End Sub

Private Sub tdb_UnidadeMedida_FilterChange()
    mblnPrimeiraVez = False
    gblnFilraCampos tdb_UnidadeMedida
End Sub

Private Sub tdb_UnidadeMedida_HeadClick(ByVal ColIndex As Integer)
    gOrdenaGrid tdb_UnidadeMedida, ColIndex
End Sub

Private Sub tdb_UnidadeMedida_KeyPress(KeyAscii As Integer)
    Select Case tdb_UnidadeMedida.Col
        Case 0
            CaracterValido KeyAscii, "N", tdb_UnidadeMedida
        Case Else
            CaracterValido KeyAscii, "A", tdb_UnidadeMedida
    End Select
End Sub

Private Sub tdb_UnidadeMedida_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    With tdb_UnidadeMedida
        If Not .EOF And Not .BOF Then
            If mblnPrimeiraVez Then
                txtPKID.Text = .Columns("PKID").Value
                LeDaTabelaParaObj gstrUnidadeMedida, Me
                
                gCorLinhaSelecionada tdb_UnidadeMedida
                
                HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrDeletar
                If mobjAux Is Nothing Then
                    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar
                Else
                    HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrAplicar
                End If
                mblnSelecionou = True
                'mblnPrimeiraVez = False
                mblnAlterando = True
                
                strDescricaoAtual = txtstrDescricao.Text
                strSiglaAtual = txtstrAbreviatura.Text
                
                HabilitaDesabilitaObjParaAlteracao (Trim(txtdtmInativo) = "")
                
            End If
        End If
    End With
End Sub

Public Sub MantemForm(ByVal strModoOperacao As String)

If strModoOperacao = UCase(gstrSalvar) Then
    If blnDadosOk Then
        blnAlterando = mblnAlterando
        ToolBarGeral strModoOperacao, gstrUnidadeMedida, mblnAlterando, _
            tdb_UnidadeMedida, Me, mobjAux, strQuery, strQuery, _
                rptUnidadeDeMedida, strQueryRelatorio
    End If
Else
    ToolBarGeral strModoOperacao, gstrUnidadeMedida, mblnAlterando, _
        tdb_UnidadeMedida, Me, mobjAux, strQuery, strQuery, _
            rptUnidadeDeMedida, strQueryRelatorio
End If

HabilitaDesabilitaObjParaAlteracao IIf(Trim(txtdtmInativo) = "", True, False)

End Sub

Private Sub txtdtmInativo_GotFocus()
    MarcaCampo txtdtmInativo
End Sub

Private Sub txtdtmInativo_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "D", txtdtmInativo
End Sub

Private Sub txtdtmInativo_LostFocus()
    txtdtmInativo = gstrDataFormatada(txtdtmInativo)
End Sub

Private Sub txtstrAbreviatura_GotFocus()
    MarcaCampo txtstrAbreviatura
End Sub

Private Sub txtstrAbreviatura_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrAbreviatura
End Sub

Private Sub txtstrDescricao_GotFocus()
    MarcaCampo txtstrDescricao
End Sub

Private Sub txtstrDescricao_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrDescricao
End Sub

Function strQueryRelatorio() As String
    Dim strSql As String
    strSql = ""
    strSql = strSql & " SELECT * "
    strSql = strSql & " FROM " & gstrUnidadeMedida
    If mblnSelecionou = True Then
        If txtPKID.Text <> "" Then
            strSql = strSql & " WHERE PKId = " & Val(txtPKID.Text)
        End If
    End If
    strSql = strSql & " ORDER BY strDescricao "
    strQueryRelatorio = strSql
End Function
Private Function blnDadosOk() As Boolean
    
    If Not mblnAlterando Or (mblnAlterando And UCase(txtstrAbreviatura.Text) <> UCase(strSiglaAtual)) Then
        If gblnExisteCodigo(1, gstrUnidadeMedida, "strAbreviatura", "'" & txtstrAbreviatura.Text & "'") Then
            ExibeMensagem "A sigla informada já se encontra cadastrada."
            txtstrAbreviatura.SetFocus
            Exit Function
        End If
    End If


    If Not mblnAlterando Or (mblnAlterando And UCase(txtstrDescricao.Text) <> UCase(strDescricaoAtual)) Then
            
        If gblnExisteCodigo(1, gstrUnidadeMedida, "strDescricao", "'" & txtstrDescricao.Text & "'") Then
            ExibeMensagem "A Descrição informada já se encontra cadastrada."
            txtstrDescricao.SetFocus
            Exit Function
        End If
    End If
    
    If txtstrAbreviatura = "" Then
        ExibeMensagem "A sigla deve ser informada."
        txtstrAbreviatura.SetFocus
        Exit Function
    End If
    
    If txtstrDescricao = "" Then
        ExibeMensagem "A descrição deve ser informada."
        txtstrDescricao.SetFocus
        Exit Function
    End If
    
    blnDadosOk = True


End Function

Private Sub HabilitaDesabilitaObjParaAlteracao(ByVal blnHabilita As Boolean)
    TrocaCorObjeto txtstrAbreviatura, Not blnHabilita
    TrocaCorObjeto txtstrDescricao, Not blnHabilita
End Sub
Private Function PreencheGrid()

Dim strSql As String

strSql = ""
strSql = strSql & "SELECT * FROM " & gstrUnidadeMedida & " "

    Select Case bytOrdenacao
        
        Case 0
            If blnOrdenacaoAsc Then
                strSql = strSql & "ORDER BY pkid ASC"
            Else
                strSql = strSql & "ORDER BY pkid DESC"
            End If
        Case 1
            If blnOrdenacaoAsc Then
                strSql = strSql & "ORDER BY strDescricao ASC"
            Else
                strSql = strSql & "ORDER BY strDescricao DESC"
            End If
        Case 2
            If blnOrdenacaoAsc Then
                strSql = strSql & "ORDER BY strAbreviatura ASC"
            Else
                strSql = strSql & "ORDER BY strAbreviatura DESC"
            End If
        Case 3
            If blnOrdenacaoAsc Then
                strSql = strSql & "ORDER BY dtmInativo ASC"
            Else
                strSql = strSql & "ORDER BY dtmInativo DESC"
            End If
    End Select
        
LeDaTabelaParaObj "", tdb_UnidadeMedida, strSql

End Function
