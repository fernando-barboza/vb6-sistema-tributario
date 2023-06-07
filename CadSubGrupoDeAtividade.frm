VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "TODG7.OCX"
Begin VB.Form frmCadSubGrupoDeAtividade 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SubGrupo de Atividades"
   ClientHeight    =   5160
   ClientLeft      =   1575
   ClientTop       =   2550
   ClientWidth     =   6555
   Icon            =   "CadSubGrupoDeAtividade.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5160
   ScaleWidth      =   6555
   Begin TabDlg.SSTab tab_3dPasta 
      Height          =   4965
      Left            =   120
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   90
      Width           =   6345
      _ExtentX        =   11192
      _ExtentY        =   8758
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "SubGrupo"
      TabPicture(0)   =   "CadSubGrupoDeAtividade.frx":1042
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblintCodigo"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblintCodigodoGrupo"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblstrNomeDoSubGrupo"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "tdb_SubGupoDeAtividade"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "dbcintCodigoDoGrupo"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtintCodigo"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmd_Grupo"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtstrNomeDoSubGrupo"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      Begin VB.TextBox txtstrNomeDoSubGrupo 
         Height          =   285
         Left            =   1080
         MaxLength       =   30
         TabIndex        =   2
         Top             =   1215
         Width           =   5085
      End
      Begin VB.CommandButton cmd_Grupo 
         Height          =   315
         Left            =   5820
         Picture         =   "CadSubGrupoDeAtividade.frx":105E
         Style           =   1  'Graphical
         TabIndex        =   6
         TabStop         =   0   'False
         ToolTipText     =   "Ativa Cadastro de Grupos de Atividade"
         Top             =   540
         Width           =   360
      End
      Begin VB.TextBox txtintCodigo 
         Height          =   285
         Left            =   1080
         MaxLength       =   15
         TabIndex        =   1
         Top             =   900
         Width           =   1305
      End
      Begin MSDataListLib.DataCombo dbcintCodigoDoGrupo 
         Height          =   315
         Left            =   1080
         TabIndex        =   0
         Top             =   540
         Width           =   4755
         _ExtentX        =   8387
         _ExtentY        =   556
         _Version        =   393216
         IntegralHeight  =   0   'False
         Text            =   ""
      End
      Begin TrueOleDBGrid70.TDBGrid tdb_SubGupoDeAtividade 
         Height          =   3255
         Left            =   180
         TabIndex        =   3
         Top             =   1590
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   5741
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
         Columns(2).DataField=   "strNomeDoSubGrupo"
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
         Splits(0)._ColumnProps(5)=   "Column(0).Visible=0"
         Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(7)=   "Column(1).Width=1799"
         Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=1720"
         Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(12)=   "Column(2).Width=8229"
         Splits(0)._ColumnProps(13)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(14)=   "Column(2)._WidthInPix=8149"
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
         _StyleDefs(44)  =   "Named:id=33:Normal"
         _StyleDefs(45)  =   ":id=33,.parent=0"
         _StyleDefs(46)  =   "Named:id=34:Heading"
         _StyleDefs(47)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(48)  =   ":id=34,.wraptext=-1"
         _StyleDefs(49)  =   "Named:id=35:Footing"
         _StyleDefs(50)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(51)  =   "Named:id=36:Selected"
         _StyleDefs(52)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(53)  =   "Named:id=37:Caption"
         _StyleDefs(54)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(55)  =   "Named:id=38:HighlightRow"
         _StyleDefs(56)  =   ":id=38,.parent=33,.bgcolor=&H80000014&,.fgcolor=&H80000012&"
         _StyleDefs(57)  =   "Named:id=39:EvenRow"
         _StyleDefs(58)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(59)  =   "Named:id=40:OddRow"
         _StyleDefs(60)  =   ":id=40,.parent=33"
         _StyleDefs(61)  =   "Named:id=41:RecordSelector"
         _StyleDefs(62)  =   ":id=41,.parent=34"
         _StyleDefs(63)  =   "Named:id=42:FilterBar"
         _StyleDefs(64)  =   ":id=42,.parent=33"
      End
      Begin VB.Label lblstrNomeDoSubGrupo 
         AutoSize        =   -1  'True
         Caption         =   "SubGrupo"
         Height          =   195
         Left            =   240
         TabIndex        =   9
         Top             =   1260
         Width           =   720
      End
      Begin VB.Label lblintCodigodoGrupo 
         AutoSize        =   -1  'True
         Caption         =   "Grupo"
         Height          =   195
         Left            =   525
         TabIndex        =   8
         Top             =   600
         Width           =   435
      End
      Begin VB.Label lblintCodigo 
         AutoSize        =   -1  'True
         Caption         =   "Código"
         Height          =   195
         Left            =   450
         TabIndex        =   7
         Top             =   945
         Width           =   495
      End
   End
   Begin VB.TextBox txtPKId 
      Enabled         =   0   'False
      Height          =   270
      Left            =   1470
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "frmCadSubGrupoDeAtividade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim mblnAlterando      As Boolean
    Dim mobjAux            As Object
    Dim mlngUltimo         As Long
    Dim mblnGuardaUltimo   As Boolean
    Dim strDuplicataCodigo As String
    Dim mblnSelecionou     As Boolean
    Dim mblnPrimeiraVez    As Boolean

Private Sub cmd_Grupo_Click()
    ChamaFormCadastro frmCadGrupoDeAtividade, dbcintCodigoDoGrupo
    frmCadGrupoDeAtividade.dbcintUtilizacaoDaTabelaDeValor.Text = frmCadAtividadeEconomica.dbcintUtilizacao.Text
    TrocaCorObjeto frmCadGrupoDeAtividade.dbcintUtilizacaoDaTabelaDeValor, True, True
    TrocaCorObjeto dbcintCodigoDoGrupo, False, False
End Sub

Private Sub dbcintCodigoDoGrupo_Click(Area As Integer)
    DropDownDataCombo dbcintCodigoDoGrupo, Me, Area
    If Area = 2 And dbcintCodigoDoGrupo.MatchedWithList Then
        If mblnGuardaUltimo = False Then
            mlngUltimo = dbcintCodigoDoGrupo.BoundText
        End If
        VerificaListaAutomatica gstrSubGrupoDeAtividade, tdb_SubGupoDeAtividade, strQuerySubGrupo
        MantemForm gstrNovo
        mblnPrimeiraVez = False
    End If
End Sub

Function blnTemDuplicata() As Boolean
Dim strSQl             As String
Dim adoResultado       As ADODB.Recordset
blnTemDuplicata = False
strDuplicataCodigo = ""
    If Val(txtintCodigo) = 0 Then Exit Function
    strSQl = ""
    strSQl = strSQl & "SELECT * FROM "
    strSQl = strSQl & gstrSubGrupoDeAtividade
    strSQl = strSQl & " WHERE intCodigo = " & txtintCodigo
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSQl, 5, adoResultado) Then
        With adoResultado
            If Not .EOF Then
                strDuplicataCodigo = strSQl
                blnTemDuplicata = True
                txtPKId.Text = !Pkid
                Exit Function
            End If
        End With
    End If
End Function

Private Sub dbcintCodigoDoGrupo_KeyDown(KeyCode As Integer, Shift As Integer)
   DropDownDataCombo dbcintCodigoDoGrupo, Me, , KeyCode, Shift
End Sub

Private Sub dbcintCodigoDoGrupo_KeyPress(KeyAscii As Integer)
  CaracterValido KeyAscii, "A", dbcintCodigoDoGrupo
End Sub

Private Sub tdb_SubGupoDeAtividade_KeyPress(KeyAscii As Integer)
CaracterValido KeyAscii, "A", tdb_SubGupoDeAtividade
End Sub

Private Sub txtintCodigo_GotFocus()
    MarcaCampo txtintCodigo
End Sub

Private Sub txtintCodigo_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtintCodigo
End Sub

Private Sub Form_Activate()
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
    dbcintCodigoDoGrupo.Tag = strQueryGrupo & ";strNomeDoGrupo"
    PreencherListaDeOpcoes dbcintCodigoDoGrupo
    VerificaObjParaAplicar mobjAux
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar, gstrDeletar
    mblnSelecionou = False
    mblnPrimeiraVez = False
End Sub

Private Function strQueryGrupo() As String
Dim strSQl As String
        strSQl = ""
    strSQl = strSQl & " SELECT GA.PKId, GA.Strnomedogrupo FROM "
    strSQl = strSQl & gstrSubGrupoDeAtividade & " SGA, "
    strSQl = strSQl & gstrGrupoDeAtividade & " GA"
    If dbcintCodigoDoGrupo.MatchedWithList Then
        strSQl = strSQl & " WHERE GA.pkid = " & dbcintCodigoDoGrupo.BoundText
    End If
    strSQl = strSQl & " ORDER BY strNomeDoSubGrupo"
    strQueryGrupo = strSQl

End Function

Private Function strQuery() As String
    Dim strSQl  As String
    strSQl = ""
    strSQl = strSQl & " SELECT PKId, intCodigo, strNomeDoSubGrupo FROM "
    strSQl = strSQl & gstrSubGrupoDeAtividade & " "
    If dbcintCodigoDoGrupo.MatchedWithList Then
        strSQl = strSQl & "WHERE intCodigoDoGrupo = " & dbcintCodigoDoGrupo.BoundText & " "
    End If
    strSQl = strSQl & "ORDER BY strNomeDoSubGrupo"
    strQuery = strSQl
End Function

Private Function strQueryAplicar() As String
    Dim strSQl  As String
    strSQl = ""
    strSQl = strSQl & "PKId, strNomeDoSubGrupo"
    strQueryAplicar = strSQl
End Function

Private Sub tdb_SubGupoDeAtividade_Click()
    mblnPrimeiraVez = True
End Sub

Private Sub tdb_SubGupoDeAtividade_FilterChange()
    gblnFilraCampos tdb_SubGupoDeAtividade
End Sub

Private Sub tdb_SubGupoDeAtividade_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    With tdb_SubGupoDeAtividade
        If Not .EOF And Not .BOF Then
            If mblnPrimeiraVez Then
                mblnAlterando = True
                txtPKId.Text = .Columns("PKID").Value
                LeDaTabelaParaObj gstrSubGrupoDeAtividade, Me
                gCorLinhaSelecionada tdb_SubGupoDeAtividade
                
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

Public Sub MantemForm(ByVal strModoOperacao As String)
    Dim varBookMark As Variant
    Dim strSQl As String
    
    mblnGuardaUltimo = True
    
    If Not tdb_SubGupoDeAtividade.EOF Then
        varBookMark = tdb_SubGupoDeAtividade.Bookmark
    Else
        mblnAlterando = False
    End If
    
    strSQl = strQuery
    
    If UCase(strModoOperacao) = gstrPreencherLista Then
        PreencherListaDeOpcoes Me.ActiveControl
        Exit Sub
    End If
    
    If UCase(strModoOperacao) = "SALVAR" And txtintCodigo = "" Then
        ExibeMensagem "O código tem que ser digitado."
        txtintCodigo.SetFocus
        Exit Sub
    End If
    
    If UCase(strModoOperacao) = "SALVAR" Or UCase(strModoOperacao) = "DELETAR" Then
        mblnPrimeiraVez = False
    End If
    
    ToolBarGeral strModoOperacao, gstrSubGrupoDeAtividade, mblnAlterando, tdb_SubGupoDeAtividade, Me, mobjAux, strSQl, strQueryAplicar
    
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrDeletar
End Sub

Private Sub txtstrNomeDoSubGrupo_GotFocus()
    MarcaCampo txtstrNomeDoSubGrupo
End Sub

Private Sub txtstrNomeDoSubGrupo_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrNomeDoSubGrupo
End Sub

Private Function strQuerry() As String
    Dim strSQl As String
    strSQl = ""
    strSQl = strSQl & "SELECT SB.PKId, SB.strNomeDoSubGrupo "
    strSQl = strSQl & "FROM " & gstrSubGrupoDeAtividade & " SB"
    strQuerry = strSQl
End Function

Private Function strQuerySubGrupo() As String
    Dim strSQl As String
    strSQl = ""
    strSQl = strSQl & "SELECT PKId, intCodigo, strNomeDoSubGrupo "
    strSQl = strSQl & "FROM " & gstrSubGrupoDeAtividade & " "
    strSQl = strSQl & "WHERE intCodigoDoGrupo = " & dbcintCodigoDoGrupo.BoundText
    strQuerySubGrupo = strSQl
End Function
    
