VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "TODG7.OCX"
Begin VB.Form frmCadDocumentoEmitido 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Documentos Emitidos"
   ClientHeight    =   5220
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7605
   HelpContextID   =   44
   Icon            =   "CadDocumentoEmitido.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5220
   ScaleWidth      =   7605
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab tab_3dPasta 
      Height          =   5115
      Left            =   120
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   30
      Width           =   7395
      _ExtentX        =   13044
      _ExtentY        =   9022
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Documentos Emitidos"
      TabPicture(0)   =   "CadDocumentoEmitido.frx":1042
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblPKID"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblstrDescricao"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "tdb_DocumentoEmitido"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txtPKId"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtstrDescricao"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "chkbytVerificaDivida"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "fra_FrameSaco"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      Begin VB.Frame fra_FrameSaco 
         Caption         =   " Texto "
         Height          =   1935
         Left            =   120
         TabIndex        =   7
         Top             =   1020
         Width           =   7155
         Begin VB.TextBox txtstrTexto 
            Height          =   1725
            Left            =   0
            MaxLength       =   4790
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   8
            Top             =   210
            Width           =   7155
         End
      End
      Begin VB.CheckBox chkbytVerificaDivida 
         Caption         =   "Verificar Dívidas"
         Height          =   210
         Left            =   1980
         TabIndex        =   6
         Top             =   420
         Width           =   1470
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
         Left            =   855
         MaxLength       =   30
         TabIndex        =   3
         Top             =   720
         Width           =   6405
      End
      Begin VB.TextBox txtPKId 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         Left            =   855
         OLEDragMode     =   1  'Automatic
         OLEDropMode     =   1  'Manual
         TabIndex        =   2
         Top             =   390
         Width           =   855
      End
      Begin TrueOleDBGrid70.TDBGrid tdb_DocumentoEmitido 
         Height          =   1965
         Left            =   120
         TabIndex        =   1
         Top             =   3030
         Width           =   7185
         _ExtentX        =   12674
         _ExtentY        =   3466
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
         Splits(0)._ColumnProps(1)=   "Column(0).Width=1746"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1667"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=2"
         Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(7)=   "Column(1).Width=10398"
         Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=10319"
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
         _StyleDefs(24)  =   "Splits(0).Style:id=43,.parent=1"
         _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=52,.parent=4"
         _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=44,.parent=2"
         _StyleDefs(27)  =   "Splits(0).FooterStyle:id=45,.parent=3"
         _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=46,.parent=5"
         _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=48,.parent=6"
         _StyleDefs(30)  =   "Splits(0).EditorStyle:id=47,.parent=7"
         _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=49,.parent=8"
         _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=50,.parent=9"
         _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=51,.parent=10"
         _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=53,.parent=11"
         _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=54,.parent=12"
         _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=58,.parent=43,.alignment=1"
         _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=55,.parent=44"
         _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=56,.parent=45"
         _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=57,.parent=47"
         _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=62,.parent=43"
         _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=59,.parent=44"
         _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=60,.parent=45"
         _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=61,.parent=47"
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
      Begin VB.Label lblstrDescricao 
         AutoSize        =   -1  'True
         Caption         =   "Decriçao"
         Height          =   195
         Left            =   135
         TabIndex        =   5
         Top             =   750
         Width           =   645
      End
      Begin VB.Label lblPKID 
         AutoSize        =   -1  'True
         Caption         =   "Código"
         Height          =   195
         Left            =   285
         TabIndex        =   4
         Top             =   420
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmCadDocumentoEmitido"
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
    gintCodSeguranca = 758
    VirificaGradeListView Me
    HabilitaDesabilitaBotao1 True, gstrMnuArquivo
    HabilitaDesabilitaBotao1 mblnSelecionou, gstrMnuArquivo, gstrDeletar
    HabilitaDesabilitaBotao1 ((mobjAux Is Nothing) And mblnSelecionou), gstrMnuArquivo, gstrAplicar
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
    LeDaTabelaParaObj gstrDocumentoEmitido, tdb_DocumentoEmitido, strQuery
    VerificaObjParaAplicar mobjAux
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar, gstrDeletar
    mblnSelecionou = False
    mblnPrimeiraVez = False
End Sub

Private Function strQuery() As String
    Dim strSQL  As String
    strSQL = ""
    strSQL = strSQL & " SELECT PKId, strDescricao,strtexto FROM "
    strSQL = strSQL & gstrDocumentoEmitido & " ORDER BY strDescricao"
    strQuery = strSQL
End Function

Private Sub tdb_DocumentoEmitido_Click()
mblnPrimeiraVez = True
    With tdb_DocumentoEmitido
        If Not .BOF And Not .EOF Then
            If .Bookmark = 1 Then
                tdb_DocumentoEmitido_RowColChange 0, 0
            End If
       End If
    End With
End Sub

Private Sub tdb_DocumentoEmitido_FilterChange()
    gblnFilraCampos tdb_DocumentoEmitido
End Sub

Private Sub tdb_DocumentoEmitido_HeadClick(ByVal ColIndex As Integer)
    gOrdenaGrid tdb_DocumentoEmitido, ColIndex
End Sub

Private Sub tdb_DocumentoEmitido_KeyPress(KeyAscii As Integer)
    Select Case tdb_DocumentoEmitido.Col
        Case 0
            CaracterValido KeyAscii, "N", tdb_DocumentoEmitido
        Case Else
            CaracterValido KeyAscii, "A", tdb_DocumentoEmitido
    End Select
End Sub

Private Sub tdb_DocumentoEmitido_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    With tdb_DocumentoEmitido
        If Not .EOF And Not .BOF Then
            If mblnPrimeiraVez Then
                txtPKID.Text = .Columns("PKID").Value
                mblnAlterando = True
                LeDaTabelaParaObj gstrDocumentoEmitido, Me
                gCorLinhaSelecionada tdb_DocumentoEmitido
                
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
    Dim strSQL As String
    
    strSQL = strQuery
    
    If strModoOperacao = UCase("IMPRIMIR") Then
        ToolBarGeral strModoOperacao, gstrDocumentoEmitido, mblnAlterando, tdb_DocumentoEmitido, Me, mobjAux, strSQL, , rptdocumentosemitidos, strQuery
        Exit Sub
    End If
    If UCase(strModoOperacao) = UCase(gstrSalvar) And txtstrTexto.Text = "" Then
        ExibeMensagem "O campo texto não pode ser nulo."
        Exit Sub
    End If
    If UCase(strModoOperacao) = UCase(gstrSalvar) And txtstrDescricao.Text = "" Then
        ExibeMensagem "O campo descrição não pode ser nulo."
        Exit Sub
    End If
    
    ToolBarGeral strModoOperacao, gstrDocumentoEmitido, mblnAlterando, tdb_DocumentoEmitido, Me, mobjAux, strSQL, strSQL
    
    If Not gobjGeral Is Nothing Then HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrDeletar
    tab_3dPasta.Tab = 0
End Sub

Private Sub txtstrDescricao_GotFocus()
    MarcaCampo txtstrDescricao
End Sub

Private Sub txtstrDescricao_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrDescricao
End Sub

Private Sub txtstrTexto_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrTexto
End Sub

