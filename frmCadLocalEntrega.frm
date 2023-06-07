VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "TODG7.OCX"
Begin VB.Form frmCadLocalEntrega 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Locais de Entrega"
   ClientHeight    =   4755
   ClientLeft      =   2520
   ClientTop       =   2490
   ClientWidth     =   6015
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4755
   ScaleWidth      =   6015
   Begin TabDlg.SSTab tab_3dPasta 
      Height          =   4590
      Left            =   105
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   90
      Width           =   5835
      _ExtentX        =   10292
      _ExtentY        =   8096
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Locais de Entrega"
      TabPicture(0)   =   "frmCadLocalEntrega.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblstrDescricao"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblintCep"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblstrComplemento"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblintLogradouro"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblintBairro"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblintNumero"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "dbcintLogradouro"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "dbcintBairro"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "tdb_Lista"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtPKId"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtstrDescricao"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "cmd_Bairro"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "cmd_Logradouro"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtintCep"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txtstrComplemento"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txtintNumero"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).ControlCount=   16
      Begin VB.TextBox txtintNumero 
         Height          =   285
         Left            =   1080
         MaxLength       =   8
         TabIndex        =   2
         Top             =   1380
         Width           =   855
      End
      Begin VB.TextBox txtstrComplemento 
         Height          =   285
         Left            =   2685
         MaxLength       =   20
         TabIndex        =   3
         Top             =   1380
         Width           =   1350
      End
      Begin VB.TextBox txtintCep 
         Height          =   285
         Left            =   4590
         MaxLength       =   9
         TabIndex        =   4
         Top             =   1380
         Width           =   1080
      End
      Begin VB.CommandButton cmd_Logradouro 
         Height          =   300
         Left            =   5280
         Picture         =   "frmCadLocalEntrega.frx":001C
         Style           =   1  'Graphical
         TabIndex        =   11
         TabStop         =   0   'False
         Tag             =   "584"
         ToolTipText     =   "Ativa Cadastro de Logradouro"
         Top             =   960
         Width           =   330
      End
      Begin VB.CommandButton cmd_Bairro 
         Height          =   300
         Left            =   3645
         Picture         =   "frmCadLocalEntrega.frx":0302
         Style           =   1  'Graphical
         TabIndex        =   10
         TabStop         =   0   'False
         Tag             =   "581"
         ToolTipText     =   "Ativa Cadastro de Bairro"
         Top             =   1785
         Width           =   330
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
         MaxLength       =   50
         TabIndex        =   0
         Top             =   570
         Width           =   4545
      End
      Begin VB.TextBox txtPKId 
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         OLEDragMode     =   1  'Automatic
         OLEDropMode     =   1  'Manual
         TabIndex        =   8
         Top             =   -30
         Visible         =   0   'False
         Width           =   855
      End
      Begin TrueOleDBGrid70.TDBGrid tdb_Lista 
         Height          =   1965
         Left            =   150
         TabIndex        =   6
         Top             =   2430
         Width           =   5595
         _ExtentX        =   9869
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
         Columns(1).DataField=   "strDescricao"
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
         Splits(0)._ColumnProps(1)=   "Column(0).Width=1455"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1376"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0).AllowSizing=0"
         Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(7)=   "Column(1).Width=7726"
         Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=7646"
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
      Begin MSDataListLib.DataCombo dbcintBairro 
         Height          =   315
         Left            =   1065
         TabIndex        =   5
         Top             =   1785
         Width           =   2595
         _ExtentX        =   4577
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dbcintLogradouro 
         Height          =   315
         Left            =   1065
         TabIndex        =   1
         Top             =   960
         Width           =   4230
         _ExtentX        =   7461
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin VB.Label lblintNumero 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "N°"
         Height          =   195
         Left            =   810
         TabIndex        =   16
         Top             =   1425
         Width           =   180
      End
      Begin VB.Label lblintBairro 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Bairro"
         Height          =   195
         Left            =   555
         TabIndex        =   15
         Top             =   1905
         Width           =   405
      End
      Begin VB.Label lblintLogradouro 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Logradouro"
         Height          =   195
         Left            =   165
         TabIndex        =   14
         Top             =   1080
         Width           =   810
      End
      Begin VB.Label lblstrComplemento 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Compl."
         Height          =   195
         Left            =   2130
         TabIndex        =   13
         Top             =   1425
         Width           =   480
      End
      Begin VB.Label lblintCep 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Cep"
         Height          =   195
         Left            =   4215
         TabIndex        =   12
         Top             =   1425
         Width           =   285
      End
      Begin VB.Label lblstrDescricao 
         AutoSize        =   -1  'True
         Caption         =   "Descrição"
         Height          =   195
         Left            =   255
         TabIndex        =   9
         Top             =   615
         Width           =   720
      End
   End
End
Attribute VB_Name = "frmCadLocalEntrega"
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

    Dim strDescricaoAtual As String

    Dim bytOrdenacao      As Byte
    Dim blnOrdenacaoAsc   As Boolean


Private Sub CarregaListaBairro()
Dim strSQL As String
Dim adoResultado As adodb.Recordset

strSQL = ""
strSQL = "SELECT intBairro FROM " & gstrLogradouro & " WHERE PKId='" & dbcintLogradouro.BoundText & "'"

Set gobjBanco = New clsBanco
gobjBanco.CriaADO strSQL, 5, adoResultado

With adoResultado
    If Not (.BOF And .EOF) Then
        PreencherListaDeOpcoes dbcintBairro, !intBairro
    End If
End With

Set adoResultado = Nothing

End Sub



Private Function strQuery() As String

Dim strSQL  As String
   
    strSQL = ""
    
    strSQL = strSQL & " SELECT PKId, strDescricao FROM "
    
    strSQL = strSQL & gstrLocalEntrega
   
    Select Case bytOrdenacao
   
        Case Is = 1
            strSQL = strSQL & " ORDER BY strDescricao" & IIf(blnOrdenacaoAsc, " ASC", " DESC")
        Case Is = 0
            strSQL = strSQL & " ORDER BY PKId" & IIf(blnOrdenacaoAsc, " ASC", " DESC")
            
    End Select
    
    strQuery = strSQL
    
End Function

Private Function strQueryAplicar() As String
    Dim strSQL  As String
    strSQL = ""
    strSQL = strSQL & " SELECT PKId, strDescricao FROM "
    strSQL = strSQL & gstrLocalEntrega & " ORDER BY strDescricao"
    strQueryAplicar = strSQL
End Function

Private Sub cmd_Logradouro_Click()
    ChamaFormCadastro frmCadLogradouro, dbcintLogradouro
End Sub

Private Sub Form_Activate()
    gintCodSeguranca = 1024
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

   bytOrdenacao = 2: blnOrdenacaoAsc = True
   
    'VerificaListaAutomatica "", tdb_Lista, strQuery
    VerificaObjParaAplicar mobjAux
    
    dbcintBairro.Tag = gstrQueryDataComboBairro & ";strDescricao"
    dbcintLogradouro.Tag = gstrQueryLogradouro & ";L.strDescricao"
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar
    mblnSelecionou = False
    mblnPrimeiraVez = False
End Sub

Private Sub tdb_Lista_Click()
    If glngQtdLinhaTDBGrid(tdb_Lista) = 1 Then
        tdb_Lista_RowColChange 0, 0
    End If
    mblnPrimeiraVez = True
End Sub

Sub tdb_Lista_DblClick()
    MantemForm gstrAplicar
End Sub

Private Sub tdb_Lista_FilterChange()
    mblnPrimeiraVez = False
    gblnFilraCampos tdb_Lista
End Sub

Private Sub tdb_Lista_HeadClick(ByVal ColIndex As Integer)
   
   blnOrdenacaoAsc = IIf(bytOrdenacao = ColIndex, Not blnOrdenacaoAsc, True)
   
   bytOrdenacao = ColIndex: MantemForm gstrRefresh
   
End Sub

Private Sub tdb_Lista_KeyDown(KeyCode As Integer, Shift As Integer)
    mblnClickOk = True
End Sub

Private Sub tdb_Lista_KeyPress(KeyAscii As Integer)
    Select Case tdb_Lista.Col
        Case 1
'            CaracterValido KeyAscii, "N", tdb_Lista
        Case Else
            CaracterValido KeyAscii, "A", tdb_Lista
    End Select
End Sub

Private Sub tdb_Lista_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mblnClickOk = True
End Sub

Private Sub tdb_Lista_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    With tdb_Lista
        If (Not .EOF And Not .BOF) And mblnClickOk Then
            If mblnPrimeiraVez Then
                mblnClickOk = False
                txtPKId.Text = .Columns("PKID").Value
                LeDaTabelaParaObj gstrLocalEntrega, Me
                gCorLinhaSelecionada tdb_Lista
                If mobjAux Is Nothing Then
                    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar
                Else
                    HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrAplicar
                End If
                mblnSelecionou = True
                mblnAlterando = True
                
                strDescricaoAtual = txtstrDescricao
            End If
        End If
    End With
End Sub

Public Sub MantemForm(ByVal strModoOperacao As String)
        
    If UCase(strModoOperacao) = gstrSalvar Then
        If Not blnDadosOK Then Exit Sub
    End If
    
    ToolBarGeral strModoOperacao, gstrLocalEntrega, mblnAlterando, tdb_Lista, _
                 Me, mobjAux, strQuery, strQueryAplicar, rptCadLocalEntrega, strQueryRelatorio
                 
End Sub

Private Sub txtstrDescricao_GotFocus()
    MarcaCampo txtstrDescricao
End Sub

Private Sub txtstrDescricao_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrDescricao
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
    strSQL = strSQL & "SELECT LE.PKId, LE.strDescricao, LO.strDescricao AS strLogradouro, "
    strSQL = strSQL & "BA.strDescricao AS strBairro, LE.intNumero, LE.intCEP "
    strSQL = strSQL & "FROM "
    strSQL = strSQL & gstrLocalEntrega & " LE, "
    strSQL = strSQL & gstrLogradouro & " LO, "
    strSQL = strSQL & gstrBairro & " BA "
    strSQL = strSQL & "WHERE LE.intLogradouro " & strOUTJSQLServer & "= LO.PKId " & strOUTJOracle & " AND LE.intBairro = BA.PKId"
      
    Select Case bytOrdenacao
      
        Case Is = 1
            strSQL = strSQL & " ORDER BY PKId" & IIf(blnOrdenacaoAsc, " ASC", " DESC")
         
        Case Is = 2
            strSQL = strSQL & " ORDER BY LE.strDescricao" & IIf(blnOrdenacaoAsc, " ASC", " DESC")
      
    End Select
   
    strQueryRelatorio = strSQL
   
End Function

Private Function blnDadosOK()
    
    blnDadosOK = False
    If Trim(txtstrDescricao.Text) = "" Then
        ExibeMensagem "Preencha corretamente o campo descrição!"
        txtstrDescricao.SetFocus
        Exit Function
    ElseIf dbcintLogradouro.Text = "" Or Not dbcintLogradouro.MatchedWithList Then
        ExibeMensagem "Preencha corretamente o campo logradouro!"
        dbcintLogradouro.SetFocus
        Exit Function
    ElseIf dbcintBairro.Text = "" Or Not dbcintBairro.MatchedWithList Then
        ExibeMensagem "Preencha corretamente o campo bairro!"
        dbcintBairro.SetFocus
        Exit Function
    End If
   
    If Not mblnAlterando Or (mblnAlterando And UCase$(txtstrDescricao.Text) <> UCase$(strDescricaoAtual)) Then
        If gblnExisteCodigo(1, gstrLocalEntrega, "strDescricao", "'" & Trim(txtstrDescricao.Text) & "'") Then
            ExibeMensagem "A descrição informada já se encontra cadastrada."
            txtstrDescricao.SetFocus
            Exit Function
        End If
    End If

    blnDadosOK = True
    
End Function

Private Sub dbcintLogradouro_Click(Area As Integer)
   If Area = 0 Then DropDownDataCombo dbcintLogradouro, Me, Area
   CarregaListaBairro
End Sub

Private Sub dbcintLogradouro_KeyDown(KeyCode As Integer, Shift As Integer)
   DropDownDataCombo dbcintLogradouro, Me, , KeyCode, Shift
End Sub

Private Sub dbcintLogradouro_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub
Private Sub txtintNumero_GotFocus()
    MarcaCampo txtintNumero
End Sub

Private Sub txtintNumero_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtintNumero
End Sub

Private Sub txtstrComplemento_GotFocus()
    MarcaCampo txtstrComplemento
End Sub

Private Sub txtstrComplemento_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrComplemento
End Sub

Private Sub dbcintBairro_Click(Area As Integer)
   If Area = 0 Then DropDownDataCombo dbcintBairro, Me, Area
End Sub

Private Sub dbcintBairro_KeyDown(KeyCode As Integer, Shift As Integer)
   DropDownDataCombo dbcintBairro, Me, , KeyCode, Shift
End Sub

Private Sub dbcintBairro_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub txtintCEP_GotFocus()
    MarcaCampo txtintCep
End Sub

Private Sub txtintCEP_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "E", txtintCep
End Sub

Private Sub txtintCEP_LostFocus()
    txtintCep = gstrCEPFormatado(txtintCep)
End Sub

Private Sub cmd_Bairro_Click()
    ChamaFormCadastro frmCadBairro, dbcintBairro
End Sub
