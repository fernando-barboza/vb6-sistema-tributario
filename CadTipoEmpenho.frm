VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frmCadTipoEmpenho 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tipo de Empenho"
   ClientHeight    =   4140
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7095
   Icon            =   "CadTipoEmpenho.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4140
   ScaleWidth      =   7095
   Begin VB.TextBox txtPKId 
      Height          =   270
      Left            =   2580
      TabIndex        =   7
      Top             =   150
      Visible         =   0   'False
      Width           =   645
   End
   Begin TabDlg.SSTab tab_3dPasta 
      Height          =   3915
      Left            =   135
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   120
      Width           =   6840
      _ExtentX        =   12065
      _ExtentY        =   6906
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Tipo de Empenho"
      TabPicture(0)   =   "CadTipoEmpenho.frx":1042
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblstrDescricao"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblstrCodigo"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "txtstrDescricao"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txtstrCodigo"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "tdb_tipoEmpenho"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "chkbytAdiantamento"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      Begin VB.CheckBox chkbytAdiantamento 
         Caption         =   "Adiantamento"
         Height          =   195
         Left            =   2490
         TabIndex        =   1
         Top             =   495
         Width           =   1275
      End
      Begin TrueOleDBGrid70.TDBGrid tdb_tipoEmpenho 
         Height          =   2475
         Left            =   120
         TabIndex        =   3
         Top             =   1320
         Width           =   6585
         _ExtentX        =   11615
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
         Columns(1).Caption=   "Código"
         Columns(1).DataField=   "strCodigo"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Descrição"
         Columns(2).DataField=   "strDescricao"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   3
         Splits(0)._UserFlags=   1
         Splits(0).MarqueeStyle=   5
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
         Splits(0)._ColumnProps(7)=   "Column(1).Width=2434"
         Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=2355"
         Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(12)=   "Column(2).Width=8652"
         Splits(0)._ColumnProps(13)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(14)=   "Column(2)._WidthInPix=8573"
         Splits(0)._ColumnProps(15)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(16)=   "Column(2).Order=3"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   0
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
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&H80000005&"
         _StyleDefs(7)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
         _StyleDefs(8)   =   "HeadingStyle:id=2,.parent=1,.namedParent=34"
         _StyleDefs(9)   =   "FooterStyle:id=3,.parent=1,.namedParent=35"
         _StyleDefs(10)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(11)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
         _StyleDefs(12)  =   "EditorStyle:id=7,.parent=1,.bgcolor=&H80000005&"
         _StyleDefs(13)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38,.bgcolor=&H80000002&"
         _StyleDefs(14)  =   ":id=8,.fgcolor=&H80000009&"
         _StyleDefs(15)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
         _StyleDefs(16)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
         _StyleDefs(17)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(18)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
         _StyleDefs(19)  =   "Splits(0).Style:id=13,.parent=1"
         _StyleDefs(20)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
         _StyleDefs(21)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
         _StyleDefs(22)  =   "Splits(0).FooterStyle:id=15,.parent=3"
         _StyleDefs(23)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
         _StyleDefs(24)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
         _StyleDefs(25)  =   "Splits(0).EditorStyle:id=17,.parent=7"
         _StyleDefs(26)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
         _StyleDefs(27)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
         _StyleDefs(28)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
         _StyleDefs(29)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(30)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(31)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
         _StyleDefs(32)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
         _StyleDefs(33)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
         _StyleDefs(34)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
         _StyleDefs(35)  =   "Splits(0).Columns(1).Style:id=46,.parent=13"
         _StyleDefs(36)  =   "Splits(0).Columns(1).HeadingStyle:id=43,.parent=14"
         _StyleDefs(37)  =   "Splits(0).Columns(1).FooterStyle:id=44,.parent=15"
         _StyleDefs(38)  =   "Splits(0).Columns(1).EditorStyle:id=45,.parent=17"
         _StyleDefs(39)  =   "Splits(0).Columns(2).Style:id=32,.parent=13"
         _StyleDefs(40)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
         _StyleDefs(41)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
         _StyleDefs(42)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
         _StyleDefs(43)  =   "Named:id=33:Normal"
         _StyleDefs(44)  =   ":id=33,.parent=0"
         _StyleDefs(45)  =   "Named:id=34:Heading"
         _StyleDefs(46)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(47)  =   ":id=34,.wraptext=-1"
         _StyleDefs(48)  =   "Named:id=35:Footing"
         _StyleDefs(49)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(50)  =   "Named:id=36:Selected"
         _StyleDefs(51)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(52)  =   "Named:id=37:Caption"
         _StyleDefs(53)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(54)  =   "Named:id=38:HighlightRow"
         _StyleDefs(55)  =   ":id=38,.parent=33,.bgcolor=&H80000014&,.fgcolor=&H80000012&"
         _StyleDefs(56)  =   "Named:id=39:EvenRow"
         _StyleDefs(57)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(58)  =   "Named:id=40:OddRow"
         _StyleDefs(59)  =   ":id=40,.parent=33"
         _StyleDefs(60)  =   "Named:id=41:RecordSelector"
         _StyleDefs(61)  =   ":id=41,.parent=34"
         _StyleDefs(62)  =   "Named:id=42:FilterBar"
         _StyleDefs(63)  =   ":id=42,.parent=33"
      End
      Begin VB.TextBox txtstrCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   990
         MaxLength       =   8
         MultiLine       =   -1  'True
         OLEDragMode     =   1  'Automatic
         TabIndex        =   0
         Top             =   480
         Width           =   1215
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
         Left            =   990
         MaxLength       =   100
         TabIndex        =   2
         Top             =   840
         Width           =   5715
      End
      Begin VB.Label lblstrCodigo 
         AutoSize        =   -1  'True
         Caption         =   "Código"
         Height          =   195
         Left            =   420
         TabIndex        =   6
         Top             =   585
         Width           =   495
      End
      Begin VB.Label lblstrDescricao 
         AutoSize        =   -1  'True
         Caption         =   "Descrição"
         Height          =   195
         Left            =   180
         TabIndex        =   5
         Top             =   915
         Width           =   720
      End
   End
End
Attribute VB_Name = "frmCadTipoEmpenho"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim mblnAlterando   As Boolean
    Dim mobjAux         As Object
    Dim mblnSelecionou As Boolean
    Dim mblnPrimeiraVez As Boolean

Private Sub chkbytAdiantamento_KeyPress(KeyAscii As Integer)
   CaracterValido KeyAscii
End Sub

Private Sub Form_Activate()
    VirificaGradeListView Me
    
'=============
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
'=============
End Sub

Private Sub Form_Deactivate()
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar
End Sub

Private Sub Form_Load()
    mblnAlterando = False
    VerificaListaAutomatica gstrTipoEmpenho, tdb_tipoEmpenho, strQuery
    VerificaObjParaAplicar mobjAux
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar
    mblnSelecionou = False
    mblnPrimeiraVez = False
End Sub

Private Function strQuery() As String
    Dim strSQL  As String
    strSQL = ""
    strSQL = strSQL & "SELECT PKId, strCodigo, strDescricao FROM "
    strSQL = strSQL & gstrTipoEmpenho
    strSQL = strSQL & " ORDER BY " & gstrCONVERT(CDT_INT, "strCodigo")
    strQuery = strSQL
End Function
Private Function strQueryAplicar() As String
    Dim strSQL  As String
    strSQL = ""
    strSQL = strSQL & "SELECT PKId, strDescricao, strCodigo FROM "
    strSQL = strSQL & gstrTipoEmpenho
    strSQL = strSQL & " ORDER BY " & gstrCONVERT(CDT_INT, "strCodigo")
    strQueryAplicar = strSQL
End Function

Private Function LeTabelaTipoEmpenho() As String
    Dim strSQL       As String
    Dim adoResultado As ADODB.Recordset
    
    strSQL = ""
    strSQL = strSQL & "SELECT PKId, strCodigo, strDescricao FROM "
    strSQL = strSQL & gstrTipoEmpenho
    strSQL = strSQL & " ORDER BY " & gstrCONVERT(CDT_INT, "strCodigo")
        
    Set gobjBanco = New clsBanco
    
    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
       Set tdb_tipoEmpenho.DataSource = adoResultado
       tdb_tipoEmpenho.Refresh
    End If
        
End Function

Private Sub tdb_tipoEmpenho_Click()
    mblnPrimeiraVez = True
   If glngQtdLinhaTDBGrid(tdb_tipoEmpenho) = 1 Then
        tdb_TipoEmpenho_RowColChange 0, 0
   End If
End Sub

Private Sub tdb_tipoEmpenho_DblClick()
   
   MantemForm gstrAplicar
   
End Sub

Private Sub tdb_TipoEmpenho_FilterChange()
    gblnFilraCampos tdb_tipoEmpenho
End Sub

Private Sub tdb_tipoEmpenho_KeyPress(KeyAscii As Integer)
     CaracterValido KeyAscii, "A", tdb_tipoEmpenho
End Sub

Private Sub tdb_TipoEmpenho_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    
    With tdb_tipoEmpenho
        If Not .EOF And Not .BOF Then
            txtPKId.Text = .Columns("PKID").Value
            If mblnPrimeiraVez Then
                LeDaTabelaParaObj gstrTipoEmpenho, Me
                gCorLinhaSelecionada tdb_tipoEmpenho
                HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrDeletar
                If mobjAux Is Nothing Then
                    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar
                Else
                    HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrAplicar
                End If
                mblnSelecionou = True
                mblnAlterando = True
            End If
        End If
    End With
End Sub

Public Sub MantemForm(ByVal strModoOperacao As String)
    
    If UCase(strModoOperacao) = UCase(gstrSalvar) Then
       If gblnExisteCodigo(1, gstrTipoEmpenho, "strCodigo", txtstrCodigo) Then
         'mstrCodigo = (gstrProximoCodigo(txtstrCodigo, gstrTipoCredito, "strCodigo", gintCodSeguranca, , , , True))
         ExibeMensagem "O número de tipo de recurso informado já se encontra cadastrado."
             If txtstrCodigo.Enabled = True Then
                txtstrCodigo.SetFocus
             End If
             Exit Sub
      End If
'      If VerificaAdiantamento Then
'         ExibeMensagem "Não é possível cadastrar mais de um tipo de empenho de adiantamento. "
'         Exit Sub
'      End If
    End If
    ToolBarGeral strModoOperacao, gstrTipoEmpenho, mblnAlterando, tdb_tipoEmpenho, Me, mobjAux, strQuery, strQueryAplicar
    
    Select Case UCase(strModoOperacao)
       Case UCase(gstrSalvar)
          LeTabelaTipoEmpenho
       Case UCase(gstrDeletar)
          LeTabelaTipoEmpenho
    End Select
    
End Sub

Private Sub txtstrDescricao_GotFocus()
    MarcaCampo txtstrDescricao
End Sub

Private Sub txtstrDescricao_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrDescricao
End Sub

Private Sub txtstrCodigo_GotFocus()
    Dim strSQL       As String
    Dim adoResultado As ADODB.Recordset
    
    strSQL = "SELECT (" & gstrISNULL("MAX(" & gstrCONVERT(CDT_INT, "strCodigo") & ")", "0") & " + 1) AS strCodigo   FROM " & gstrTipoEmpenho
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
       txtstrCodigo = adoResultado!strCodigo
       MarcaCampo txtstrCodigo
    End If
    
    adoResultado.Close: Set adoResultado = Nothing
    
End Sub

Private Sub txtstrcodigo_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtstrCodigo
End Sub
Private Function VerificaAdiantamento() As Boolean

   Dim strSQL       As String
   Dim adoResultado As ADODB.Recordset
   
   strSQL = "SELECT bytAdiantamento FROM " & gstrTipoEmpenho
   strSQL = strSQL & " WHERE bytAdiantamento = 1"
   
   If chkbytAdiantamento.Value = vbChecked Then
      Set gobjBanco = New clsBanco
      
      If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
         If Not adoResultado.EOF Then
            VerificaAdiantamento = True
         End If
      End If
   
   End If
End Function
