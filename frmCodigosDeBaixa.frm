VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "TODG7.OCX"
Begin VB.Form frmCodigosDeBaixa 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Códigos de Baixa"
   ClientHeight    =   5220
   ClientLeft      =   1680
   ClientTop       =   1935
   ClientWidth     =   7110
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5220
   ScaleWidth      =   7110
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab tab_3dPasta 
      Height          =   5025
      Left            =   90
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   75
      Width           =   6885
      _ExtentX        =   12144
      _ExtentY        =   8864
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Códigos de Baixa"
      TabPicture(0)   =   "frmCodigosDeBaixa.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblDescricao"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblAbreviatura"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblstrhistorico"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "tdb_CodigosDeBaixa"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "fra_Tipo"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtPKId"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtstrDescricao"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtstrAbreviatura"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtstrhistorico"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).ControlCount=   9
      Begin VB.TextBox txtstrhistorico 
         Height          =   705
         Left            =   990
         MaxLength       =   150
         MultiLine       =   -1  'True
         TabIndex        =   2
         Top             =   1200
         Width           =   5790
      End
      Begin VB.TextBox txtstrAbreviatura 
         Height          =   315
         Left            =   990
         MaxLength       =   10
         TabIndex        =   1
         Top             =   825
         Width           =   1380
      End
      Begin VB.TextBox txtstrDescricao 
         Height          =   315
         Left            =   990
         MaxLength       =   50
         TabIndex        =   0
         Top             =   465
         Width           =   4950
      End
      Begin VB.TextBox txtPKId 
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         Left            =   2130
         Locked          =   -1  'True
         OLEDragMode     =   1  'Automatic
         OLEDropMode     =   1  'Manual
         TabIndex        =   12
         Top             =   15
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Frame fra_Tipo 
         Caption         =   "Tipos"
         Height          =   645
         Left            =   180
         TabIndex        =   11
         Top             =   1935
         Width           =   6630
         Begin VB.OptionButton optbytTipo 
            Caption         =   "Acordo"
            CausesValidation=   0   'False
            Height          =   225
            Index           =   5
            Left            =   5730
            TabIndex        =   8
            Top             =   300
            Width           =   825
         End
         Begin VB.OptionButton optbytTipo 
            Caption         =   "Atraso"
            CausesValidation=   0   'False
            Height          =   225
            Index           =   4
            Left            =   4860
            TabIndex        =   7
            Top             =   300
            Width           =   1290
         End
         Begin VB.OptionButton optbytTipo 
            Caption         =   "Normal"
            Height          =   225
            Index           =   0
            Left            =   150
            TabIndex        =   3
            Top             =   300
            Value           =   -1  'True
            Width           =   795
         End
         Begin VB.OptionButton optbytTipo 
            Caption         =   "Única"
            CausesValidation=   0   'False
            Height          =   225
            Index           =   1
            Left            =   1095
            TabIndex        =   4
            Top             =   300
            Width           =   750
         End
         Begin VB.OptionButton optbytTipo 
            Caption         =   "Cancelamento"
            CausesValidation=   0   'False
            Height          =   225
            Index           =   2
            Left            =   2025
            TabIndex        =   5
            Top             =   300
            Width           =   1350
         End
         Begin VB.OptionButton optbytTipo 
            Caption         =   "Dívida Ativa"
            CausesValidation=   0   'False
            Height          =   225
            Index           =   3
            Left            =   3510
            TabIndex        =   6
            Top             =   300
            Width           =   1290
         End
      End
      Begin TrueOleDBGrid70.TDBGrid tdb_CodigosDeBaixa 
         Height          =   2055
         Left            =   195
         TabIndex        =   9
         Top             =   2775
         Width           =   6600
         _ExtentX        =   11642
         _ExtentY        =   3625
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
         Columns(1).DataField=   "Descricao"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Abreviatura"
         Columns(2).DataField=   "Abreviatura"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Tipo"
         Columns(3).DataField=   "Tipo"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "Histórico"
         Columns(4).DataField=   "strhistorico"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   5
         Splits(0)._UserFlags=   0
         Splits(0).MarqueeStyle=   3
         Splits(0).RecordSelectors=   0   'False
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
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
         Splits(0)._ColumnProps(7)=   "Column(0).AllowFocus=0"
         Splits(0)._ColumnProps(8)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(9)=   "Column(1).Width=6456"
         Splits(0)._ColumnProps(10)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(11)=   "Column(1)._WidthInPix=6376"
         Splits(0)._ColumnProps(12)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(14)=   "Column(2).Width=1984"
         Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=1905"
         Splits(0)._ColumnProps(17)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(19)=   "Column(3).Width=2619"
         Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=2540"
         Splits(0)._ColumnProps(22)=   "Column(3)._EditAlways=0"
         Splits(0)._ColumnProps(23)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(24)=   "Column(4).Width=8758"
         Splits(0)._ColumnProps(25)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(26)=   "Column(4)._WidthInPix=8678"
         Splits(0)._ColumnProps(27)=   "Column(4)._EditAlways=0"
         Splits(0)._ColumnProps(28)=   "Column(4).Order=5"
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
         TabAction       =   1
         WrapCellPointer =   -1  'True
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
         _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=28,.parent=13"
         _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
         _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=32,.parent=13"
         _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
         _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
         _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
         _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=50,.parent=13"
         _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
         _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
         _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
         _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=54,.parent=13"
         _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14"
         _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
         _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
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
      Begin VB.Label lblstrhistorico 
         AutoSize        =   -1  'True
         Caption         =   "Histórico"
         Height          =   195
         Left            =   330
         TabIndex        =   15
         Top             =   1155
         Width           =   615
      End
      Begin VB.Label lblAbreviatura 
         AutoSize        =   -1  'True
         Caption         =   "Abreviatura"
         Height          =   195
         Left            =   135
         TabIndex        =   14
         Top             =   810
         Width           =   810
      End
      Begin VB.Label lblDescricao 
         AutoSize        =   -1  'True
         Caption         =   "Descrição"
         Height          =   195
         Left            =   225
         TabIndex        =   13
         Top             =   435
         Width           =   720
      End
   End
End
Attribute VB_Name = "frmCodigosDeBaixa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim blnAlterando            As Boolean
    Dim bytOrdenacao            As Byte
    Dim blnOrdenacaoAsc         As Boolean
    Dim blnPrimeiraVez          As Boolean
    Dim strDescricaoAtual       As String
    Dim strAbreviaturaAtual     As String

Private Sub Form_Activate()
    gintCodSeguranca = 1120
    HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrDeletar
End Sub

Private Sub Form_Deactivate()
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar
End Sub

Private Sub Label1_Click()

End Sub

Private Sub tdb_CodigosDeBaixa_Click()
    blnPrimeiraVez = True
End Sub

Private Sub tdb_CodigosDeBaixa_FilterChange()
    blnPrimeiraVez = False
    gblnFilraCampos tdb_CodigosDeBaixa
End Sub

Private Sub tdb_CodigosDeBaixa_HeadClick(ByVal ColIndex As Integer)
    gOrdenaGrid tdb_CodigosDeBaixa, ColIndex
End Sub

Private Sub tdb_CodigosDeBaixa_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDown, vbKeyUp, vbKeyLeft, vbKeyRight
            blnPrimeiraVez = True
    End Select
    
End Sub

Private Sub tdb_CodigosDeBaixa_KeyPress(KeyAscii As Integer)
    Select Case tdb_CodigosDeBaixa.Col
        Case 1
            CaracterValido KeyAscii, "A", tdb_CodigosDeBaixa
        Case 2
            CaracterValido KeyAscii, "A", tdb_CodigosDeBaixa
        Case 3
            CaracterValido KeyAscii, "A", tdb_CodigosDeBaixa
    End Select
End Sub

Private Sub tdb_CodigosDeBaixa_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    With tdb_CodigosDeBaixa
        If Not .EOF And blnPrimeiraVez Then
            txtPKID.Text = .Columns("PKID").Value
            blnAlterando = True
            LeDaTabelaParaObj gstrCodigoDeBaixa, Me
            strDescricaoAtual = tdb_CodigosDeBaixa.Columns("Descrição").Value
            strAbreviaturaAtual = tdb_CodigosDeBaixa.Columns("Abreviatura").Value
        End If
    End With
End Sub

Public Sub MantemForm(ByVal strModoOperacao As String)
Dim strsql As String
    Select Case UCase(strModoOperacao)
        Case Is = UCase(gstrSalvar)
            If Not blnDadosOk Then Exit Sub
            ToolBarGeral strModoOperacao, gstrCodigoDeBaixa, blnAlterando, tdb_CodigosDeBaixa, Me, , strQuery(gstrSalvar), , , , True
            blnPrimeiraVez = False
        Case Is = UCase(gstrNovo)
            LimpaObjeto Me
            blnPrimeiraVez = False
            blnAlterando = False
            txtstrDescricao.SetFocus
        Case Is = UCase(gstrImprimir)
            ImprimeRelatorio rptcodigodebaixa, strQueryRelatorio
        Case Else
            ToolBarGeral strModoOperacao, gstrCodigoDeBaixa, blnAlterando, tdb_CodigosDeBaixa, Me, , strQuery
    End Select
                 
End Sub

Private Function blnDadosOk()
    blnDadosOk = False
    
    If txtstrDescricao.Text = "" Then
        ExibeMensagem "É necessário preencher o campo Descrição."
        If txtstrDescricao.Enabled Then txtstrDescricao.SetFocus
        Exit Function
    End If
    
    If txtstrabreviatura.Text = "" Then
        ExibeMensagem "É necessário preencher a Abreviatura."
        If txtstrabreviatura.Enabled Then txtstrabreviatura.SetFocus
        Exit Function
    End If
    
    If Not blnAlterando Or (blnAlterando And RTrim(LTrim(strDescricaoAtual)) <> LTrim(RTrim(txtstrDescricao.Text))) Then
        If gblnExisteCodigo(1, gstrCodigoDeBaixa, "strDescricao", "'" & txtstrDescricao.Text & "'") Then
            ExibeMensagem "Já existe um registro com a mesma Descrição informada."
            If txtstrDescricao.Enabled Then txtstrDescricao.SetFocus
            Exit Function
        End If
    End If
    If Not blnAlterando Or (blnAlterando And LTrim(RTrim(strAbreviaturaAtual)) <> LTrim(RTrim(txtstrabreviatura.Text))) Then
        If gblnExisteCodigo(1, gstrCodigoDeBaixa, "strAbreviatura", "'" & txtstrabreviatura.Text & "'") Then
            ExibeMensagem "Já existe um registro com a mesma abreviatura informado."
            If txtstrabreviatura.Enabled Then txtstrabreviatura.SetFocus
            Exit Function
        End If
    End If
        
    blnDadosOk = True
    
End Function

Private Function strQueryRelatorio() As String
Dim strsql As String
strsql = ""
strsql = strsql & " Select CB.strDescricao strDescricao,"
    strsql = strsql & " CB.strAbreviatura strAbreviatura, "
    strsql = strsql & gstrCASEWHEN("CB.BYTTipo", " 0, 'Normal', 1, 'Única', 2, 'Cancelamento', 3, 'Dívida Ativa', 4, 'Atraso', 5, 'Acordo'") & " strTipo, "
    strsql = strsql & " CB.strhistorico strhistorico"
    strsql = strsql & " FROM "
    strsql = strsql & gstrCodigoDeBaixa & " CB"
strQueryRelatorio = strsql
End Function

Private Function strQuery(Optional strModoOperacao As String) As String
Dim strsql As String

    strsql = "SELECT CB.Pkid,"
    strsql = strsql & " CB.strDescricao Descricao,"
    strsql = strsql & " CB.strAbreviatura Abreviatura, "
    strsql = strsql & " CB.strhistorico, "
    strsql = strsql & gstrCASEWHEN("CB.BYTTipo", " 0, 'Normal', 1, 'Única', 2, 'Cancelamento', 3, 'Dívida Ativa', 4, 'Atraso', 5, 'Acordo'") & " Tipo"
    strsql = strsql & " FROM "
    strsql = strsql & gstrCodigoDeBaixa & " CB"

    If UCase(strModoOperacao) = UCase(gstrSalvar) Then
        If Not blnAlterando Then
            If glngPegaUltimaChave(gstrCodigoDeBaixa, "Pkid") + 1 = 1 Then
                MantemForm gstrLocalizar
            Else
                strsql = strsql & " WHERE CB.Pkid = " & glngPegaUltimaChave(gstrCodigoDeBaixa, "Pkid") + 1
            End If
        Else
            strsql = strsql & " WHERE CB.Pkid = " & Val(txtPKID.Text)
        End If
    End If
    
    Select Case bytOrdenacao
        Case Is = 1
            strsql = strsql & " ORDER BY CB.strDescricao " & IIf(blnOrdenacaoAsc, " ASC", " DESC")
        Case Is = 2
            strsql = strsql & " ORDER BY CB.strAbreviatura " & IIf(blnOrdenacaoAsc, " ASC", " DESC")
        Case Is = 3
            strsql = strsql & " ORDER BY Tipo " & IIf(blnOrdenacaoAsc, " ASC", " DESC")
    End Select
    
    strQuery = strsql

End Function
Private Sub txtstrAbreviatura_GotFocus()
    MarcaCampo txtstrabreviatura
End Sub

Private Sub txtstrAbreviatura_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrabreviatura
End Sub

Private Sub txtstrDescricao_GotFocus()
    MarcaCampo txtstrDescricao
End Sub

Private Sub txtstrDescricao_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrDescricao
End Sub


