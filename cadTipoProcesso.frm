VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "Tabctl32.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frmCadTipoProcesso 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tipos de Processo"
   ClientHeight    =   5340
   ClientLeft      =   1425
   ClientTop       =   1875
   ClientWidth     =   9390
   HelpContextID   =   14
   Icon            =   "cadTipoProcesso.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5340
   ScaleWidth      =   9390
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab Tab_3dpasta 
      Height          =   5145
      Left            =   120
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   120
      Width           =   9165
      _ExtentX        =   16166
      _ExtentY        =   9075
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Tipos de Processo"
      TabPicture(0)   =   "cadTipoProcesso.frx":1042
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblintPrazoArquivamento"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblintPrazoTramitacao"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblstrdescricao"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblintcodigo"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblstrTipoArquivamento"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lbl_Prazo(0)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lbl_Prazo(1)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "tdb_TipoProcesso"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtintPrazoArquivamento"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtstrtipoArquivamento"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtintprazotramitacao"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtstrdescricao"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtstrcodigo"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).ControlCount=   13
      Begin VB.TextBox txtstrcodigo 
         Height          =   285
         Left            =   1635
         MaxLength       =   5
         TabIndex        =   0
         Top             =   420
         Width           =   645
      End
      Begin VB.TextBox txtstrdescricao 
         Height          =   285
         Left            =   1635
         MaxLength       =   100
         TabIndex        =   1
         Top             =   750
         Width           =   7425
      End
      Begin VB.TextBox txtintprazotramitacao 
         Height          =   285
         Left            =   1635
         MaxLength       =   4
         TabIndex        =   3
         Top             =   1410
         Width           =   1395
      End
      Begin VB.TextBox txtstrtipoArquivamento 
         Height          =   285
         Left            =   1635
         MaxLength       =   60
         TabIndex        =   2
         Top             =   1080
         Width           =   7425
      End
      Begin VB.TextBox txtintPrazoArquivamento 
         Height          =   285
         Left            =   5415
         MaxLength       =   4
         TabIndex        =   4
         Top             =   1410
         Width           =   1095
      End
      Begin TrueOleDBGrid70.TDBGrid tdb_TipoProcesso 
         Height          =   3195
         Left            =   105
         TabIndex        =   5
         Top             =   1800
         Width           =   8925
         _ExtentX        =   15743
         _ExtentY        =   5636
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
         Columns(1).DataField=   "strCodigo"
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
         Splits(0)._ColumnProps(8)=   "Column(1).Width=1905"
         Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=1826"
         Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(13)=   "Column(2).Width=13335"
         Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=13256"
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
         _StyleDefs(37)  =   "Splits(0).Columns(0).Style:id=46,.parent=13"
         _StyleDefs(38)  =   "Splits(0).Columns(0).HeadingStyle:id=43,.parent=14"
         _StyleDefs(39)  =   "Splits(0).Columns(0).FooterStyle:id=44,.parent=15"
         _StyleDefs(40)  =   "Splits(0).Columns(0).EditorStyle:id=45,.parent=17"
         _StyleDefs(41)  =   "Splits(0).Columns(1).Style:id=28,.parent=13"
         _StyleDefs(42)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14"
         _StyleDefs(43)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
         _StyleDefs(44)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
         _StyleDefs(45)  =   "Splits(0).Columns(2).Style:id=32,.parent=13"
         _StyleDefs(46)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
         _StyleDefs(47)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
         _StyleDefs(48)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
         _StyleDefs(49)  =   "Named:id=33:Normal"
         _StyleDefs(50)  =   ":id=33,.parent=0"
         _StyleDefs(51)  =   "Named:id=34:Heading"
         _StyleDefs(52)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(53)  =   ":id=34,.wraptext=-1"
         _StyleDefs(54)  =   "Named:id=35:Footing"
         _StyleDefs(55)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(56)  =   "Named:id=36:Selected"
         _StyleDefs(57)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(58)  =   "Named:id=37:Caption"
         _StyleDefs(59)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(60)  =   "Named:id=38:HighlightRow"
         _StyleDefs(61)  =   ":id=38,.parent=33,.bgcolor=&H80000013&,.fgcolor=&H80000012&"
         _StyleDefs(62)  =   "Named:id=39:EvenRow"
         _StyleDefs(63)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(64)  =   "Named:id=40:OddRow"
         _StyleDefs(65)  =   ":id=40,.parent=33"
         _StyleDefs(66)  =   "Named:id=41:RecordSelector"
         _StyleDefs(67)  =   ":id=41,.parent=34"
         _StyleDefs(68)  =   "Named:id=42:FilterBar"
         _StyleDefs(69)  =   ":id=42,.parent=33"
      End
      Begin VB.Label lbl_Prazo 
         AutoSize        =   -1  'True
         Caption         =   "dias"
         Height          =   195
         Index           =   1
         Left            =   3120
         TabIndex        =   14
         Top             =   1455
         Width           =   285
      End
      Begin VB.Label lbl_Prazo 
         AutoSize        =   -1  'True
         Caption         =   "dias"
         Height          =   195
         Index           =   0
         Left            =   6570
         TabIndex        =   13
         Top             =   1455
         Width           =   285
      End
      Begin VB.Label lblstrTipoArquivamento 
         AutoSize        =   -1  'True
         Caption         =   "Arquivamento"
         Height          =   195
         Left            =   600
         TabIndex        =   12
         Top             =   1110
         Width           =   975
      End
      Begin VB.Label lblintcodigo 
         AutoSize        =   -1  'True
         Caption         =   "Código"
         Height          =   195
         Left            =   1080
         TabIndex        =   10
         Top             =   465
         Width           =   495
      End
      Begin VB.Label lblstrdescricao 
         AutoSize        =   -1  'True
         Caption         =   "Descrição"
         Height          =   195
         Left            =   855
         TabIndex        =   9
         Top             =   795
         Width           =   720
      End
      Begin VB.Label lblintPrazoTramitacao 
         AutoSize        =   -1  'True
         Caption         =   "Prazo de Tramitação"
         Height          =   195
         Left            =   105
         TabIndex        =   8
         Top             =   1455
         Width           =   1470
      End
      Begin VB.Label lblintPrazoArquivamento 
         AutoSize        =   -1  'True
         Caption         =   "Prazo de Arquivamento"
         Height          =   195
         Left            =   3690
         TabIndex        =   7
         Top             =   1455
         Width           =   1650
      End
   End
   Begin VB.TextBox txtPKId 
      Enabled         =   0   'False
      Height          =   600
      Left            =   360
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   330
      Visible         =   0   'False
      Width           =   825
   End
End
Attribute VB_Name = "frmCadTipoProcesso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim mblnAlterando        As Boolean
    Dim mobjAux              As Object
    Dim mblnSelecionou       As Boolean
    Dim mblnPrimeiraVez      As Boolean
    
    Dim strCodigoAtual       As String
    Dim strDescricaoAtual    As String
    Dim strCodigo            As String

Private Sub Form_Activate()
    gintCodSeguranca = 436
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
    mblnAlterando = False
    VerificaListaAutomatica gstrTipoProcesso, tdb_TipoProcesso, strQuery
    VerificaObjParaAplicar mobjAux
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar
    mblnSelecionou = False
    mblnPrimeiraVez = False
End Sub

Private Sub tab_3dPasta_KeyPress(KeyAscii As Integer)
CaracterValido KeyAscii, , Tab_3dpasta
End Sub

Private Sub tdb_TipoProcesso_Click()
    mblnPrimeiraVez = True
    With tdb_TipoProcesso
        If Not .EOF And Not .BOF Then
            If .Bookmark = 1 Then
                tdb_TipoProcesso_RowColChange 0, 0
            End If
        End If
    End With
End Sub

Sub tdb_TipoProcesso_DblClick()
    MantemForm gstrAplicar
End Sub

Private Sub tdb_TipoProcesso_FilterChange()
    gblnFilraCampos tdb_TipoProcesso
End Sub

Private Sub tdb_TipoProcesso_KeyPress(KeyAscii As Integer)
CaracterValido KeyAscii, "A", tdb_TipoProcesso
End Sub

Private Sub tdb_TipoProcesso_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    With tdb_TipoProcesso
        If Not .EOF And Not .BOF Then
            txtPKId.Text = .Columns("PKID").Value

            If mblnPrimeiraVez Then
                LeDaTabelaParaObj gstrTipoProcesso, Me

                gCorLinhaSelecionada tdb_TipoProcesso

                HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrDeletar

                If mobjAux Is Nothing Then
                    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar
                Else
                    HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrAplicar
                End If
                mblnSelecionou = True
                mblnAlterando = True
                
                strCodigoAtual = txtstrcodigo.Text
                strDescricaoAtual = txtstrdescricao.Text
                
            End If

        End If
    End With
End Sub

Public Sub MantemForm(ByVal strModoOperacao As String)
    Dim varBookMark As Variant
    Dim strSql As String
       
    strSql = strQuery
    
    If UCase(strModoOperacao) = UCase(gstrSalvar) Or UCase(strModoOperacao) = UCase(gstrDeletar) Then
        mblnPrimeiraVez = False
    End If
    
    If UCase(strModoOperacao) = UCase(gstrSalvar) Then
        If Not blnDadosOk Then Exit Sub
    End If
    
    ToolBarGeral strModoOperacao, gstrTipoProcesso, mblnAlterando, tdb_TipoProcesso, Me, mobjAux, strSql, , rptTipoDeProcesso, strQueryRelatorio
    
    If Not gobjGeral Is Nothing Then HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrDeletar
    
End Sub

Private Function strQuery() As String
Dim strSql As String
    strSql = ""
    strSql = strSql & "SELECT * "
    strSql = strSql & "FROM " & gstrTipoProcesso & " "
    strSql = strSql & "ORDER BY PKID"
strQuery = strSql
End Function

Private Sub txtstrcodigo_GotFocus()
gstrProximoCodigo txtstrcodigo, gstrTipoProcesso, "strCodigo", gintCodSeguranca
End Sub

Private Sub txtstrcodigo_KeyPress(KeyAscii As Integer)
CaracterValido KeyAscii, "A", txtstrcodigo
End Sub

Private Sub txtintPrazoArquivamento_KeyPress(KeyAscii As Integer)
CaracterValido KeyAscii, "N", txtintPrazoArquivamento
End Sub

Private Sub txtstrDescricao_KeyPress(KeyAscii As Integer)
CaracterValido KeyAscii, "A", txtstrdescricao
End Sub

Private Sub txtintprazotramitacao_KeyPress(KeyAscii As Integer)
CaracterValido KeyAscii, "N", txtintprazotramitacao
End Sub

Private Sub txtstrtipoArquivamento_KeyPress(KeyAscii As Integer)
CaracterValido KeyAscii, "A", txtstrtipoArquivamento
End Sub

Private Function strQueryRelatorio() As String
Dim strSql As String
    strSql = ""
    strSql = strSql & " SELECT "
    strSql = strSql & " strCodigo, strDescricao "
    strSql = strSql & " FROM "
    strSql = strSql & gstrTipoProcesso
    If mblnSelecionou = True Then
        If txtPKId.Text <> "" Then
            strSql = strSql & " WHERE "
            strSql = strSql & " PKId = " & Val(txtPKId.Text)
        End If
    End If
    strSql = strSql & " ORDER BY "
    strSql = strSql & " strDescricao "
strQueryRelatorio = strSql
End Function

Private Function blnDadosOk() As Boolean

    If (Trim(txtstrcodigo)) = "" Then
        ExibeMensagem "O código do tipo de processo tem que ser digitado."
        Exit Function
    End If
    
    If Trim(txtstrdescricao) = "" Then
        ExibeMensagem "O descrição do tipo de processo tem que ser informada."
        txtstrdescricao.SetFocus
        Exit Function
    ElseIf Trim(txtstrtipoArquivamento) = "" Then
        ExibeMensagem "O arquivamento tem que ser informado."
        txtstrtipoArquivamento.SetFocus
        Exit Function
    ElseIf Trim(txtintprazotramitacao) = "" Then
        ExibeMensagem "O prazo de tramitação tem que ser informado."
        txtintprazotramitacao.SetFocus
        Exit Function
    ElseIf Trim(txtintPrazoArquivamento) = "" Then
        ExibeMensagem "O prazo de arquivamento tem que ser informado."
        txtintPrazoArquivamento.SetFocus
        Exit Function
    End If
    
    If Not mblnAlterando Or (mblnAlterando And UCase$(strCodigoAtual) <> UCase$(txtstrcodigo.Text)) Then

ProximoCodigo:

        If gblnExisteCodigo(1, gstrTipoProcesso, "strCodigo", "'" & txtstrcodigo.Text & "'") Then
            strCodigo = (gstrProximoCodigo(txtstrcodigo, gstrTipoProcesso, "strCodigo", gintCodSeguranca, , , , True))
            If MsgBox("O código informado já se encontra cadastrado. Deseja usar o código " & strCodigo & "?", vbYesNo + vbQuestion) = vbNo Then
                txtstrcodigo.SetFocus
                Exit Function
            Else
                txtstrcodigo.Text = strCodigo
                GoTo ProximoCodigo
            End If
        End If
    End If
    
    If Not mblnAlterando Or (mblnAlterando And UCase$(txtstrdescricao.Text) <> UCase$(strDescricaoAtual)) Then
            
        If gblnExisteCodigo(1, gstrTipoProcesso, "strDescricao", "'" & txtstrdescricao.Text & "'") Then
            ExibeMensagem "A descrição informada já se encontra cadastrada."
            txtstrdescricao.SetFocus
            Exit Function
        End If
    End If

    blnDadosOk = True
    
End Function

