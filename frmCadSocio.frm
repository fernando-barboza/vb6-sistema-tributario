VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "TODG7.OCX"
Begin VB.Form frmCadSocio 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Sócios"
   ClientHeight    =   5265
   ClientLeft      =   2610
   ClientTop       =   2475
   ClientWidth     =   6435
   HelpContextID   =   10
   Icon            =   "frmCadSocio.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5265
   ScaleWidth      =   6435
   Begin TabDlg.SSTab tab_3dPasta 
      Height          =   5100
      Left            =   90
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   90
      Width           =   6240
      _ExtentX        =   11007
      _ExtentY        =   8996
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Sócio"
      TabPicture(0)   =   "frmCadSocio.frx":1042
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblPKID"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblintContribuinte"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmd_Contribuinte"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txtPKId"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "tdb_Socio"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "dbcintContribuinte"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      Begin MSDataListLib.DataCombo dbcintContribuinte 
         Height          =   315
         Left            =   840
         TabIndex        =   6
         Top             =   900
         Width           =   4875
         _ExtentX        =   8599
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin TrueOleDBGrid70.TDBGrid tdb_Socio 
         Height          =   3615
         Left            =   150
         TabIndex        =   5
         Top             =   1320
         Width           =   5925
         _ExtentX        =   10451
         _ExtentY        =   6376
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
         Columns(1).Caption=   "Nome"
         Columns(1).DataField=   "strNome"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   2
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
         Splits(0)._ColumnProps(0)=   "Columns.Count=2"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2064"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1984"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=2"
         Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(7)=   "Column(1).Width=7858"
         Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=7779"
         Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
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
         _StyleDefs(30)  =   "Splits(0).Columns(0).Style:id=28,.parent=13,.alignment=1"
         _StyleDefs(31)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
         _StyleDefs(32)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
         _StyleDefs(33)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
         _StyleDefs(34)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
         _StyleDefs(35)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
         _StyleDefs(36)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
         _StyleDefs(37)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
         _StyleDefs(38)  =   "Named:id=33:Normal"
         _StyleDefs(39)  =   ":id=33,.parent=0"
         _StyleDefs(40)  =   "Named:id=34:Heading"
         _StyleDefs(41)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(42)  =   ":id=34,.wraptext=-1"
         _StyleDefs(43)  =   "Named:id=35:Footing"
         _StyleDefs(44)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(45)  =   "Named:id=36:Selected"
         _StyleDefs(46)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(47)  =   "Named:id=37:Caption"
         _StyleDefs(48)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(49)  =   "Named:id=38:HighlightRow"
         _StyleDefs(50)  =   ":id=38,.parent=33,.bgcolor=&H80000014&,.fgcolor=&H80000012&"
         _StyleDefs(51)  =   "Named:id=39:EvenRow"
         _StyleDefs(52)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(53)  =   "Named:id=40:OddRow"
         _StyleDefs(54)  =   ":id=40,.parent=33"
         _StyleDefs(55)  =   "Named:id=41:RecordSelector"
         _StyleDefs(56)  =   ":id=41,.parent=34"
         _StyleDefs(57)  =   "Named:id=42:FilterBar"
         _StyleDefs(58)  =   ":id=42,.parent=33"
      End
      Begin VB.TextBox txtPKId 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         Left            =   840
         TabIndex        =   4
         Top             =   510
         Width           =   960
      End
      Begin VB.CommandButton cmd_Contribuinte 
         Height          =   315
         Left            =   5730
         Picture         =   "frmCadSocio.frx":105E
         Style           =   1  'Graphical
         TabIndex        =   2
         TabStop         =   0   'False
         Tag             =   "15"
         ToolTipText     =   "Ativa Cadastro de Contribuintes"
         Top             =   900
         Width           =   360
      End
      Begin VB.Label lblintContribuinte 
         AutoSize        =   -1  'True
         Caption         =   "Nome"
         Height          =   195
         Left            =   345
         TabIndex        =   3
         Top             =   1020
         Width           =   420
      End
      Begin VB.Label lblPKID 
         AutoSize        =   -1  'True
         Caption         =   "Código"
         Height          =   195
         Left            =   270
         TabIndex        =   1
         Top             =   600
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmCadSocio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mblnAlterando   As Boolean
Dim mobjAux         As Object
Dim mblnSelecionou  As Boolean
Dim mblnPrimeiraVez As Boolean
Dim mblnClickOk     As Boolean
Dim bytOrdenacao    As Byte
Dim blnOrdenacaoAsc As Boolean

Private Sub cmd_Contribuinte_Click()
    ChamaFormCadastro frmCadContribuinte, dbcintContribuinte
End Sub

Private Sub dbcintContribuinte_Click(Area As Integer)
    DropDownDataCombo dbcintContribuinte, Me, Area
End Sub

Private Sub dbcintContribuinte_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbcintContribuinte, Me, , KeyCode, Shift
End Sub

Private Sub dbcintContribuinte_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbcintContribuinte
End Sub

Private Sub Form_Activate()
    gintCodSeguranca = 627
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
    dbcintContribuinte.Tag = strQueryDataComboContribuinte & ";strNome"
    bytOrdenacao = 1: blnOrdenacaoAsc = True
    VerificaObjParaAplicar mobjAux
End Sub

Private Function strQueryDataComboContribuinte()
Dim strSql As String
    strSql = ""
    strSql = strSql & "SELECT PKId, strNome "
    strSql = strSql & "FROM " & gstrContribuinte & " "
    strSql = strSql & "ORDER BY strNome"
    strQueryDataComboContribuinte = strSql
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar, gstrDeletar
    mblnSelecionou = False
    mblnPrimeiraVez = False
End Sub

Private Function strQuerySocio() As String

'******************************************************************************************
' Data: 05/05/2003
' Alteração: - Retirado comando "AS" utilizado para dar apelidos às tabelas, o qual não é
'            permitido pelo Oracle, da cláusula FROM.
' Responsável: Everton Bianchini
'******************************************************************************************
Dim strSql As String
    
    strSql = ""
    strSql = strSql & "Select SO.PKId, CO.strNome "
    strSql = strSql & "From " & gstrSocio & " SO, " & gstrContribuinte & " CO "
    strSql = strSql & "Where SO.intContribuinte = CO.PKId "
    
    Select Case bytOrdenacao
      Case Is = 0
         strSql = strSql & " ORDER BY SO.PKId" & IIf(blnOrdenacaoAsc, " ASC", " DESC")
      Case Is = 1
         strSql = strSql & " ORDER BY CO.strNome" & IIf(blnOrdenacaoAsc, " ASC", " DESC")
   End Select
    
    strQuerySocio = strSql
    
End Function

Private Sub tdb_Socio_Click()
    mblnPrimeiraVez = True
    If glngQtdLinhaTDBGrid(tdb_Socio) = 1 Then
        tdb_Socio_RowColChange 0, 0
    End If
End Sub

Private Sub tdb_Socio_DblClick()
    MantemForm gstrAplicar
End Sub

Private Sub tdb_Socio_FilterChange()
    mblnPrimeiraVez = False
    gblnFilraCampos tdb_Socio
End Sub

Private Sub tdb_Socio_HeadClick(ByVal ColIndex As Integer)
    gOrdenaGrid tdb_Socio, ColIndex
End Sub

Private Sub tdb_Socio_KeyDown(KeyCode As Integer, Shift As Integer)
    mblnClickOk = True
End Sub

Private Sub tdb_Socio_KeyPress(KeyAscii As Integer)
    Select Case tdb_Socio.Col
        Case 0 'PKId
            CaracterValido KeyAscii, "N", tdb_Socio
        Case Else
            CaracterValido KeyAscii, "A", tdb_Socio
    End Select
End Sub

Private Sub tdb_Socio_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mblnClickOk = True
End Sub

Private Sub tdb_Socio_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    
    With tdb_Socio
        If (Not .EOF And Not .BOF) And mblnClickOk Then
            If mblnPrimeiraVez Then
                mblnClickOk = False
                mblnAlterando = True
                txtPKId.Text = .Columns("PKID").Value
                LeDaTabelaParaObj gstrSocio, Me
                gCorLinhaSelecionada tdb_Socio
                    
                gCorLinhaSelecionada tdb_Socio
                
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

Private Sub txtPKId_GotFocus()
    MarcaCampo txtPKId
End Sub

Function strQuerryRelatorio() As String
    Dim strSql As String
    strSql = ""
    strSql = strSql & "SELECT SC.PKId Codigo, CO.strNome Socio FROM "
    strSql = strSql & gstrSocio & " SC,"
    strSql = strSql & gstrContribuinte & " CO "
    strSql = strSql & " WHERE SC.intContribuinte = CO.PKId"
    strSql = strSql & " ORDER BY strNome "
    strQuerryRelatorio = strSql
End Function

Public Sub MantemForm(ByVal strModoOperacao As String)
    Dim varBookMark As Variant
    Dim strSql As String
    
    If UCase(strModoOperacao) = "SALVAR" Or UCase(strModoOperacao) = "DELETAR" Then
        mblnPrimeiraVez = False
    End If
    Select Case UCase(strModoOperacao)
        Case gstrNovo
            LimpaObjeto Me, mblnAlterando
            'txtPKId = glngPegaProximaChave(gstrSocio, "PKId")
        Case gstrSalvar
            If blnDadosOk Then
                ToolBarGeral strModoOperacao, gstrSocio, mblnAlterando, tdb_Socio, Me, mobjAux, strQuerySocio, strQuerySocio, rptSocios, strQuerryRelatorio
            End If
        Case Else
            ToolBarGeral strModoOperacao, gstrSocio, mblnAlterando, tdb_Socio, Me, mobjAux, strQuerySocio, strQuerySocio, rptSocios, strQuerryRelatorio
    End Select
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrDeletar

End Sub

Private Sub txtPKId_KeyPress(KeyAscii As Integer)
 CaracterValido KeyAscii, "", txtPKId
End Sub

Private Function blnDadosOk() As Boolean
    If Not dbcintContribuinte.MatchedWithList Then
        ExibeMensagem "O nome digitado não se encontra cadastrado na tabela de contribuintes!"
        dbcintContribuinte.SetFocus
        Exit Function
    End If
    
    blnDadosOk = True
    
End Function
