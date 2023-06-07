VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frmCadModalidadeAplicacao 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Modalidades de Aplica��o"
   ClientHeight    =   4980
   ClientLeft      =   1995
   ClientTop       =   1740
   ClientWidth     =   6600
   Icon            =   "CadModalidadeAplicacao.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4980
   ScaleWidth      =   6600
   Begin TabDlg.SSTab tab_3DPasta 
      Height          =   4785
      Left            =   90
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   6435
      _ExtentX        =   11351
      _ExtentY        =   8440
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "Modalidades de Aplica��o"
      TabPicture(0)   =   "CadModalidadeAplicacao.frx":1042
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblstrDescricao"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblstrCodigo"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "tdb_Lista"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txtstrDescricao"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtstrCodigo"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "chkbytPermiteAplicacao"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "fra_Convenio"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      Begin VB.Frame fra_Convenio 
         Caption         =   "Conv�nio / Entidade / Fundo"
         Height          =   675
         Left            =   420
         TabIndex        =   3
         Top             =   870
         Width           =   5925
         Begin VB.CommandButton cmd_Convenio 
            Height          =   315
            Left            =   5490
            Picture         =   "CadModalidadeAplicacao.frx":105E
            Style           =   1  'Graphical
            TabIndex        =   5
            TabStop         =   0   'False
            Tag             =   "193"
            ToolTipText     =   "Clique aqui para cadastar conv�nio"
            Top             =   270
            Width           =   360
         End
         Begin MSDataListLib.DataCombo dbcintConvenio 
            Height          =   315
            Left            =   555
            TabIndex        =   4
            Top             =   270
            Width           =   4905
            _ExtentX        =   8652
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
      End
      Begin VB.CheckBox chkbytPermiteAplicacao 
         Caption         =   "Permite Aplica��o"
         Height          =   225
         Left            =   990
         TabIndex        =   6
         Top             =   1620
         Value           =   1  'Checked
         Width           =   1605
      End
      Begin VB.TextBox txtstrCodigo 
         Height          =   285
         Left            =   990
         MaxLength       =   3
         TabIndex        =   2
         Top             =   480
         Width           =   405
      End
      Begin VB.TextBox txtstrDescricao 
         Height          =   285
         Left            =   990
         MaxLength       =   100
         TabIndex        =   8
         Top             =   1950
         Width           =   5295
      End
      Begin TrueOleDBGrid70.TDBGrid tdb_Lista 
         Height          =   2085
         Left            =   120
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   2550
         Width           =   6165
         _ExtentX        =   10874
         _ExtentY        =   3678
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
         Columns(1).Caption=   "C�digo"
         Columns(1).DataField=   "strCodigo"
         Columns(1).NumberFormat=   "FormatText Event"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Conv�nio"
         Columns(2).DataField=   "strConvenio"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Permite Aplica��o"
         Columns(3).DataField=   "bytPermiteAplicacao"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "Descri��o"
         Columns(4).DataField=   "strDescricao"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   5
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
         Splits(0)._ColumnProps(0)=   "Columns.Count=5"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0).AllowSizing=0"
         Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
         Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(8)=   "Column(1).Width=1058"
         Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=979"
         Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(13)=   "Column(2).Width=2963"
         Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=2884"
         Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(17)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(18)=   "Column(3).Width=2381"
         Splits(0)._ColumnProps(19)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(20)=   "Column(3)._WidthInPix=2302"
         Splits(0)._ColumnProps(21)=   "Column(3)._EditAlways=0"
         Splits(0)._ColumnProps(22)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(23)=   "Column(4).Width=3916"
         Splits(0)._ColumnProps(24)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(25)=   "Column(4)._WidthInPix=3836"
         Splits(0)._ColumnProps(26)=   "Column(4)._EditAlways=0"
         Splits(0)._ColumnProps(27)=   "Column(4).Order=5"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   3
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         AllowUpdate     =   0   'False
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
         _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=28,.parent=13"
         _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
         _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=50,.parent=13"
         _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=47,.parent=14"
         _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=48,.parent=15"
         _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=49,.parent=17"
         _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=54,.parent=13"
         _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=51,.parent=14"
         _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=52,.parent=15"
         _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=53,.parent=17"
         _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=32,.parent=13"
         _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=29,.parent=14"
         _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=30,.parent=15"
         _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=31,.parent=17"
         _StyleDefs(56)  =   "Named:id=33:Normal"
         _StyleDefs(57)  =   ":id=33,.parent=0"
         _StyleDefs(58)  =   "Named:id=34:Heading"
         _StyleDefs(59)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(60)  =   ":id=34,.wraptext=-1"
         _StyleDefs(61)  =   "Named:id=35:Footing"
         _StyleDefs(62)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(63)  =   "Named:id=36:Selected"
         _StyleDefs(64)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(65)  =   "Named:id=37:Caption"
         _StyleDefs(66)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(67)  =   "Named:id=38:HighlightRow"
         _StyleDefs(68)  =   ":id=38,.parent=33,.bgcolor=&H80000014&,.fgcolor=&H80000012&"
         _StyleDefs(69)  =   "Named:id=39:EvenRow"
         _StyleDefs(70)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(71)  =   "Named:id=40:OddRow"
         _StyleDefs(72)  =   ":id=40,.parent=33"
         _StyleDefs(73)  =   "Named:id=41:RecordSelector"
         _StyleDefs(74)  =   ":id=41,.parent=34"
         _StyleDefs(75)  =   "Named:id=42:FilterBar"
         _StyleDefs(76)  =   ":id=42,.parent=33"
      End
      Begin VB.Label lblstrCodigo 
         AutoSize        =   -1  'True
         Caption         =   "C�digo"
         Height          =   195
         Left            =   405
         TabIndex        =   1
         Top             =   525
         Width           =   495
      End
      Begin VB.Label lblstrDescricao 
         AutoSize        =   -1  'True
         Caption         =   "Descri��o"
         Height          =   195
         Left            =   180
         TabIndex        =   7
         Top             =   1995
         Width           =   720
      End
   End
   Begin VB.TextBox txtPKId 
      Height          =   270
      Left            =   1170
      TabIndex        =   10
      Top             =   390
      Visible         =   0   'False
      Width           =   645
   End
End
Attribute VB_Name = "frmCadModalidadeAplicacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    
Private mobjAux             As Object
Private mblnClickOk         As Boolean
Private mblnAlterando       As Boolean
Private strCodigoAtual      As String
Private strDescricaoAtual   As String
     
     
Private Function strQueryConvenio() As String
    strQueryConvenio = "Select Pkid, strDescricao From " & gstrConvenio & " Order By strDescricao"
End Function


Private Sub cmd_Convenio_Click()
    CarregaForm frmCadConvenio, dbcintConvenio
End Sub

Private Sub Form_Activate()
    
    gintCodSeguranca = 210
    
    VirificaGradeListView Me
    
    HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrImprimir
    
    If mblnAlterando Then
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

Private Function strQuery() As String

Dim strSQL  As String

    strSQL = "SELECT tm.PKId, " & gstrCONVERT(CDT_INT, "tm.strCodigo") & " strCodigo, tm.strDescricao, " & gstrCASEWHEN("tm.bytPermiteAplicacao", "0,'N�o',1,'Sim'") & " bytPermiteAplicacao, tc.strDescricao strConvenio FROM "
    strSQL = strSQL & gstrModalidade & " tm, " & gstrConvenio & " tc "
    strSQL = strSQL & " Where tc.pkid " & strOUTJOracle & "=" & strOUTJSQLServer & " tm.intConvenio "
    strSQL = strSQL & " ORDER BY " & gstrCONVERT(CDT_INT, "tm.strCodigo")
   
    strQuery = strSQL
    
End Function

Private Sub Form_Load()
    
    mblnAlterando = False
    
    dbcintConvenio.Tag = strQueryConvenio & ";strDescricao"
    
    VerificaListaAutomatica gstrModalidade, tdb_Lista, strQuery
    VerificaObjParaAplicar mobjAux
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrImprimir
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar
End Sub

Private Sub tdb_Lista_Click()
    
    mblnClickOk = True
    
    If glngQtdLinhaTDBGrid(tdb_Lista) = 1 Then
        tdb_Lista_RowColChange 0, 0
    End If
    
End Sub

Private Sub tdb_Lista_DblClick()
    MantemForm gstrAplicar
End Sub

Private Sub tdb_Lista_FilterChange()
    gblnFilraCampos tdb_Lista
End Sub

Private Sub tdb_Lista_HeadClick(ByVal ColIndex As Integer)
   gOrdenaGrid tdb_Lista, ColIndex
End Sub

Private Sub tdb_Lista_KeyPress(KeyAscii As Integer)
    If tdb_Lista.Col = 1 Then
        CaracterValido KeyAscii, "N", tdb_Lista
    End If
End Sub

Private Sub tdb_Lista_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mblnClickOk = True
End Sub

Private Sub tdb_Lista_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    
    With tdb_Lista
        If (Not .EOF And Not .BOF) And mblnClickOk Then
            mblnClickOk = False
            txtPKId.Text = .Columns("PKID").Value
            LeDaTabelaParaObj gstrModalidade, Me
            gCorLinhaSelecionada tdb_Lista
            If mobjAux Is Nothing Then
                HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar
            Else
                HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrAplicar
            End If
            strCodigoAtual = txtstrCodigo.Text
            strDescricaoAtual = txtstrDescricao.Text
            mblnAlterando = True
        End If
    End With

End Sub
Public Sub MantemForm(ByVal strModoOperacao As String)
Dim strCod As String

    If strModoOperacao = gstrSalvar Then
        If blnDadosOk Then
            ToolBarGeral strModoOperacao, gstrModalidade, mblnAlterando, _
                         tdb_Lista, Me, mobjAux, strQuery, , _
                         rptCadModalidadeDeAplicacao, strQueryRelatorio
        End If
    Else
        
        ToolBarGeral strModoOperacao, gstrModalidade, mblnAlterando, _
                    tdb_Lista, Me, mobjAux, strQuery, , _
                    rptCadModalidadeDeAplicacao, strQueryRelatorio
        
        If strModoOperacao = gstrNovo Then chkbytPermiteAplicacao.Value = vbChecked
        
    End If
End Sub
Private Sub txtstrDescricao_GotFocus()
    MarcaCampo txtstrDescricao
End Sub

Private Sub txtstrDescricao_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrDescricao
End Sub

Private Sub txtstrCodigo_GotFocus()
    gstrProximoCodigo txtstrCodigo, gstrModalidade, "strCodigo", gintCodSeguranca
    MarcaCampo txtstrCodigo
End Sub

Private Sub txtstrCodigo_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtstrCodigo
End Sub

Function strQueryRelatorio() As String

Dim strSQL As String

'    strSQL = " SELECT * "
'    strSQL = strSQL & " FROM " & gstrModalidade
'    strSQL = strSQL & " ORDER BY strDescricao "

    strSQL = "SELECT tm.PKId, " & gstrCONVERT(CDT_INT, "tm.strCodigo") & " strCodigo, tm.strDescricao, " & gstrCASEWHEN("tm.bytPermiteAplicacao", "0,'N�o',1,'Sim'") & " bytPermiteAplicacao, tc.strDescricao strConvenio FROM "
    strSQL = strSQL & gstrModalidade & " tm, " & gstrConvenio & " tc "
    strSQL = strSQL & " Where tc.pkid " & strOUTJOracle & "=" & strOUTJSQLServer & " tm.intConvenio "
    strSQL = strSQL & " ORDER BY " & gstrCONVERT(CDT_INT, "tm.strCodigo")
    
    strQueryRelatorio = strSQL
    
End Function

Private Function blnDadosOk() As Boolean
    
    Dim strCodigo         As String
    
    blnDadosOk = False
    
    If Trim(txtstrCodigo.Text) = "" Then
        ExibeMensagem "O c�digo deve ser informado."
        txtstrCodigo.SetFocus
        Exit Function
    End If
    If Trim(txtstrDescricao.Text) = "" Then
        ExibeMensagem "A descri��o deve ser informada."
        txtstrDescricao.SetFocus
        Exit Function
    End If
    
'    If mblnAlterando Then
'        If gblnExisteCodigo(1, gstrModalidade, "strDescricao", "'" & txtstrDescricao.Text & "'") Then
'            ExibeMensagem "A descri��o informada j� se encontra cadastrada."
'            txtstrDescricao.SetFocus
'            Exit Function
'        End If
'
'    Else
'        If gblnExisteCodigo(1, gstrModalidade, "strCodigo", Format(Replace(txtstrCodigo.Text, ".", ""), "00000")) Then
'            ExibeMensagem "O c�digo informado j� se encontra cadastrado."
'            txtstrCodigo.SetFocus
'            Exit Function
'        End If
'        If gblnExisteCodigo(1, gstrModalidade, "strDescricao", "'" & txtstrDescricao.Text & "'") Then
'            ExibeMensagem "A descri��o informada j� se encontra cadastrada."
'            txtstrDescricao.SetFocus
'            Exit Function
'        End If
'
'    End If
    
    If Not mblnAlterando Or (mblnAlterando And UCase$(strCodigoAtual) <> UCase$(txtstrCodigo.Text)) Then

ProximoCodigo:

        If gblnExisteCodigo(1, gstrModalidade, "strCodigo", "'" & txtstrCodigo.Text & "'") Then
            strCodigo = (gstrProximoCodigo(txtstrCodigo, gstrModalidade, "strCodigo", gintCodSeguranca, , , , True, , , , , 1))
            If MsgBox("O c�digo informado j� se encontra cadastrado. Deseja usar o c�digo " & strCodigo & "?", vbYesNo + vbQuestion) = vbNo Then
                txtstrCodigo.SetFocus
                Exit Function
            Else
                txtstrCodigo.Text = strCodigo
                GoTo ProximoCodigo
            End If
        End If
    End If
    
    If Not mblnAlterando Or (mblnAlterando And UCase$(txtstrDescricao.Text) <> UCase$(strDescricaoAtual)) Then
            
        If gblnExisteCodigo(1, gstrModalidade, "strDescricao", "'" & txtstrDescricao.Text & "'") Then
            ExibeMensagem "A descri��o informada j� se encontra cadastrada."
            txtstrDescricao.SetFocus
            Exit Function
        End If
    End If
    
    blnDadosOk = True
    
End Function

Private Sub dbcintConvenio_Click(Area As Integer)
    DropDownDataCombo dbcintConvenio, Me, Area
    If Not mblnAlterando Then
        If dbcintConvenio.MatchedWithList Then
            txtstrDescricao.Text = dbcintConvenio.Text
        Else
            txtstrDescricao.Text = Space$(0)
        End If
    End If
End Sub

Private Sub dbcintConvenio_GotFocus()
    MarcaCampo dbcintConvenio
End Sub

Private Sub dbcintConvenio_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbcintConvenio, Me, , KeyCode, Shift
End Sub

Private Sub dbcintConvenio_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbcintConvenio
End Sub
