VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "TODG7.OCX"
Begin VB.Form frmCadIndexadorEconomico 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Indexadores Econômicos"
   ClientHeight    =   5430
   ClientLeft      =   3540
   ClientTop       =   2610
   ClientWidth     =   6285
   HelpContextID   =   26
   Icon            =   "CadIndexadorEconomico.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5430
   ScaleWidth      =   6285
   Begin VB.TextBox txtPKId 
      Height          =   270
      Left            =   3120
      TabIndex        =   9
      Top             =   150
      Visible         =   0   'False
      Width           =   645
   End
   Begin TabDlg.SSTab tab_3DPasta 
      Height          =   5235
      Left            =   135
      TabIndex        =   6
      Top             =   90
      Width           =   5985
      _ExtentX        =   10557
      _ExtentY        =   9234
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Indexador"
      TabPicture(0)   =   "CadIndexadorEconomico.frx":1042
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblstrSiglaIndexador"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblstrDescricaoIndexador"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "txtstrSiglaIndexador"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txtstrDescricaoIndexador"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "tdb_IndexadorEconomico"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "fra_bytCaracteristica"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Índices econômicos"
      TabPicture(1)   =   "CadIndexadorEconomico.frx":105E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lbl_dblValor"
      Tab(1).Control(1)=   "lbl_dtmData"
      Tab(1).Control(2)=   "lbl_intDivisao"
      Tab(1).Control(3)=   "img_Aux"
      Tab(1).Control(4)=   "lvw_Indices"
      Tab(1).Control(5)=   "ssp_ToolbarCeps"
      Tab(1).Control(6)=   "txt_dblValor"
      Tab(1).Control(7)=   "txt_dtmData"
      Tab(1).Control(8)=   "txt_intDivisao"
      Tab(1).ControlCount=   9
      Begin VB.Frame fra_bytCaracteristica 
         Caption         =   "Característica"
         Height          =   540
         Left            =   300
         TabIndex        =   16
         Top             =   1260
         Width           =   3120
         Begin VB.OptionButton optbytCaracteristica 
            Caption         =   "Valor"
            CausesValidation=   0   'False
            Height          =   195
            Index           =   0
            Left            =   165
            TabIndex        =   19
            Top             =   270
            Width           =   675
         End
         Begin VB.OptionButton optbytCaracteristica 
            Caption         =   "Fator"
            CausesValidation=   0   'False
            Height          =   195
            Index           =   1
            Left            =   1080
            TabIndex        =   18
            Top             =   270
            Width           =   660
         End
         Begin VB.OptionButton optbytCaracteristica 
            Caption         =   "Percentual"
            CausesValidation=   0   'False
            Height          =   195
            Index           =   2
            Left            =   1875
            TabIndex        =   17
            Top             =   270
            Width           =   1080
         End
      End
      Begin TrueOleDBGrid70.TDBGrid tdb_IndexadorEconomico 
         Height          =   3075
         Left            =   150
         TabIndex        =   7
         Top             =   2010
         Width           =   5685
         _ExtentX        =   10028
         _ExtentY        =   5424
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
         Columns(1).DataField=   "strDescricaoIndexador"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Sigla"
         Columns(2).DataField=   "strSiglaIndexador"
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
         Splits(0)._ColumnProps(1)=   "Column(0).Width=1667"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1588"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(1).Width=5927"
         Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=5847"
         Splits(0)._ColumnProps(9)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(11)=   "Column(2).Width=1879"
         Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=1799"
         Splits(0)._ColumnProps(14)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
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
         _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=28,.parent=13"
         _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
         _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=32,.parent=13"
         _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
         _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
         _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
         _StyleDefs(48)  =   "Named:id=33:Normal"
         _StyleDefs(49)  =   ":id=33,.parent=0"
         _StyleDefs(50)  =   "Named:id=34:Heading"
         _StyleDefs(51)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(52)  =   ":id=34,.wraptext=-1"
         _StyleDefs(53)  =   "Named:id=35:Footing"
         _StyleDefs(54)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(55)  =   "Named:id=36:Selected"
         _StyleDefs(56)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(57)  =   "Named:id=37:Caption"
         _StyleDefs(58)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(59)  =   "Named:id=38:HighlightRow"
         _StyleDefs(60)  =   ":id=38,.parent=33,.bgcolor=&H80000014&,.fgcolor=&H80000012&"
         _StyleDefs(61)  =   "Named:id=39:EvenRow"
         _StyleDefs(62)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(63)  =   "Named:id=40:OddRow"
         _StyleDefs(64)  =   ":id=40,.parent=33"
         _StyleDefs(65)  =   "Named:id=41:RecordSelector"
         _StyleDefs(66)  =   ":id=41,.parent=34"
         _StyleDefs(67)  =   "Named:id=42:FilterBar"
         _StyleDefs(68)  =   ":id=42,.parent=33"
      End
      Begin VB.TextBox txt_intDivisao 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -72075
         MaxLength       =   9
         TabIndex        =   4
         Top             =   870
         Width           =   1155
      End
      Begin VB.TextBox txt_dtmData 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   -74370
         TabIndex        =   2
         Top             =   465
         Width           =   1005
      End
      Begin VB.TextBox txt_dblValor 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -74370
         MaxLength       =   12
         TabIndex        =   3
         Top             =   885
         Width           =   1425
      End
      Begin VB.TextBox txtstrDescricaoIndexador 
         Height          =   285
         Left            =   1020
         MaxLength       =   100
         TabIndex        =   0
         Top             =   480
         Width           =   4815
      End
      Begin VB.TextBox txtstrSiglaIndexador 
         Height          =   285
         Left            =   1035
         MaxLength       =   10
         TabIndex        =   1
         Top             =   825
         Width           =   1245
      End
      Begin Threed.SSPanel ssp_ToolbarCeps 
         Height          =   390
         Left            =   -70350
         TabIndex        =   5
         Top             =   765
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   688
         _StockProps     =   15
         BackColor       =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin MSComctlLib.Toolbar tlb_Indices 
            Height          =   330
            Left            =   30
            TabIndex        =   14
            Top             =   30
            Width           =   1050
            _ExtentX        =   1852
            _ExtentY        =   582
            ButtonWidth     =   609
            ButtonHeight    =   582
            Style           =   1
            ImageList       =   "img_Aux"
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   3
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Novo"
                  Object.ToolTipText     =   "Novo"
                  ImageIndex      =   1
               EndProperty
               BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Salvar"
                  Object.ToolTipText     =   "Adicionar / Alterar"
                  ImageIndex      =   2
               EndProperty
               BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Deletar"
                  Object.ToolTipText     =   "Remover"
                  ImageIndex      =   3
               EndProperty
            EndProperty
         End
      End
      Begin MSComctlLib.ListView lvw_Indices 
         Height          =   3885
         Left            =   -74880
         TabIndex        =   8
         Top             =   1245
         Width           =   5730
         _ExtentX        =   10107
         _ExtentY        =   6853
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin MSComctlLib.ImageList img_Aux 
         Left            =   -71070
         Top             =   390
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   3
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CadIndexadorEconomico.frx":107A
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CadIndexadorEconomico.frx":11DA
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CadIndexadorEconomico.frx":1336
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Label lbl_intDivisao 
         AutoSize        =   -1  'True
         Caption         =   "Divisão"
         Height          =   195
         Left            =   -72690
         TabIndex        =   15
         Top             =   960
         Width           =   525
      End
      Begin VB.Label lbl_dtmData 
         AutoSize        =   -1  'True
         Caption         =   "Data"
         Height          =   195
         Left            =   -74790
         TabIndex        =   13
         Top             =   570
         Width           =   345
      End
      Begin VB.Label lbl_dblValor 
         AutoSize        =   -1  'True
         Caption         =   "Valor"
         Height          =   195
         Left            =   -74805
         TabIndex        =   12
         Top             =   960
         Width           =   360
      End
      Begin VB.Label lblstrDescricaoIndexador 
         AutoSize        =   -1  'True
         Caption         =   "Descrição"
         Height          =   195
         Left            =   210
         TabIndex        =   11
         Top             =   570
         Width           =   720
      End
      Begin VB.Label lblstrSiglaIndexador 
         AutoSize        =   -1  'True
         Caption         =   "Sigla"
         Height          =   195
         Left            =   585
         TabIndex        =   10
         Top             =   900
         Width           =   345
      End
   End
End
Attribute VB_Name = "frmCadIndexadorEconomico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim mblnAlterando       As Boolean
    Dim mblnAlterandoIndice As Boolean
    Dim mobjAux             As Object
    Dim objList             As Object
    Dim mblnSelecionou      As Boolean
    Dim mblnPrimeiraVez     As Boolean
    
' TIMTIM - 11/02/2003 - Pendência nº 5
   Dim bytOrdenacao         As Byte
   Dim blnOrdenacaoAsc      As Boolean
   
   'Cláudio
   Dim strDescriAtual       As String
   Dim strSiglaAtual        As String
   
Private Sub Form_Activate()
    gintCodSeguranca = 397
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

   ' TIMTIM - 11/02/2003 - Pendência nº 5
   bytOrdenacao = 2: blnOrdenacaoAsc = True
   
    MontaColumnHeaders
    mblnAlterando = False
    'VerificaListaAutomatica gstrIndexadorEconomico, tdb_IndexadorEconomico, strQuery
    VerificaObjParaAplicar mobjAux
End Sub

Private Function strQuery() As String
Dim strSQL As String
   
   strSQL = ""
   
   strSQL = strSQL & "SELECT PKId, strDescricaoIndexador, strSiglaIndexador from " & gstrIndexadorEconomico
    
 ' TIMTIM - 11/02/2003 - Pendência nº 5
 
   Select Case bytOrdenacao
      
      Case Is = 0
         strSQL = strSQL & " ORDER BY PKId" & IIf(blnOrdenacaoAsc, " ASC", " DESC")
         
      Case Is = 1
         strSQL = strSQL & " ORDER BY strDescricaoIndexador" & IIf(blnOrdenacaoAsc, " ASC", " DESC")
      
      Case Is = 2
         strSQL = strSQL & " ORDER BY strSiglaIndexador" & IIf(blnOrdenacaoAsc, " ASC", " DESC")
         
   End Select
   
   strQuery = strSQL
   
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar, gstrDeletar
    mblnSelecionou = False
    mblnPrimeiraVez = False
End Sub

Private Sub lvw_Indices_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    OrdenaColunaClicada lvw_Indices, ColumnHeader
End Sub

Private Sub lvw_Indices_GotFocus()
tab_3DPasta.Tab = 1
End Sub

Private Sub lvw_Indices_ItemClick(ByVal Item As MSComctlLib.ListItem)
    With lvw_Indices
        txt_dtmData = gstrDataFormatada(.SelectedItem.Text)
        txt_dblValor = gvntConvVrDoSql(.SelectedItem.SubItems(1), 5)
        txt_intDivisao = .SelectedItem.SubItems(2)
    End With
    txt_dtmData.Locked = True
    txt_dtmData.BackColor = &HC0C0C0
    mblnAlterandoIndice = True
End Sub

Private Sub lvw_Indices_KeyPress(KeyAscii As Integer)
CaracterValido KeyAscii, "V", lvw_Indices
End Sub

Private Sub tab_3dPasta_KeyPress(KeyAscii As Integer)
CaracterValido KeyAscii, "", tab_3DPasta
End Sub

Private Sub tdb_IndexadorEconomico_Click()
    mblnPrimeiraVez = True
    If glngQtdLinhaTDBGrid(tdb_IndexadorEconomico) = 1 Then
        tdb_IndexadorEconomico_RowColChange 0, 0
    End If
End Sub

Private Sub tdb_IndexadorEconomico_DblClick()
    MantemForm gstrAplicar
End Sub

Private Sub tdb_IndexadorEconomico_FilterChange()
    mblnPrimeiraVez = False
    gblnFilraCampos tdb_IndexadorEconomico
End Sub

' TIMTIM - 11/02/2003 - Pendência nº 5
Private Sub tdb_IndexadorEconomico_HeadClick(ByVal ColIndex As Integer)
   
   blnOrdenacaoAsc = IIf(bytOrdenacao = ColIndex, Not blnOrdenacaoAsc, True)
   
   bytOrdenacao = ColIndex ': MantemForm gstrRefresh
   
   ToolBarGeral gstrRefresh, gstrIndexadorEconomico, mblnAlterando, tdb_IndexadorEconomico, Me, mobjAux, strQuery
   
End Sub

Private Sub tdb_IndexadorEconomico_GotFocus()
tab_3DPasta.Tab = 0
End Sub

Private Sub tdb_IndexadorEconomico_KeyPress(KeyAscii As Integer)
    Select Case tdb_IndexadorEconomico.Col
        Case 0
            CaracterValido KeyAscii, "N", tdb_IndexadorEconomico
        Case Else
            CaracterValido KeyAscii, "A", tdb_IndexadorEconomico
    End Select
End Sub

Private Sub tdb_IndexadorEconomico_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    With tdb_IndexadorEconomico
        If Not .EOF And Not .BOF Then
            If mblnPrimeiraVez Then
                mblnAlterando = True
                txtPKId.Text = .Columns("PKID").Value
                gCorLinhaSelecionada tdb_IndexadorEconomico
                LeDaTabelaParaObj gstrIndexadorEconomico, Me
                If txtPKId <> "" Then
                    Novo_Indexador
                    CarregaListViewIndices txtPKId
                End If
                gCorLinhaSelecionada tdb_IndexadorEconomico
                HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrDeletar
                
                If mobjAux Is Nothing Then
                    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar
                Else
                    HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrAplicar
                End If
                mblnSelecionou = True
                strDescriAtual = tdb_IndexadorEconomico.Columns("Descrição")
                strSiglaAtual = tdb_IndexadorEconomico.Columns("Sigla")
            End If
        End If
    End With
End Sub

Public Sub MantemForm(ByVal strModoOperacao As String)
    Dim varBookMark     As Variant
    Dim intIndexador    As Integer
    Dim blnAlterandoAux As Boolean
    
    If UCase(strModoOperacao) = UCase(gstrSalvar) Or UCase(strModoOperacao) = UCase(gstrDeletar) Then
        mblnPrimeiraVez = False
    End If
    
    blnAlterandoAux = mblnAlterando
    
    If mblnAlterando Then
        intIndexador = Val(txtPKId)
    End If
    
    Select Case UCase(strModoOperacao)
        Case UCase(gstrNovo)
            LimpaObjeto Me, mblnAlterando
            Novo_Indexador
            
        Case UCase(gstrSalvar)
            If blnDadosOk Then
                If ToolBarGeral(strModoOperacao, gstrIndexadorEconomico, mblnAlterando, tdb_IndexadorEconomico, Me, mobjAux, strQuery) Then
                    If blnAlterandoAux Then
                        Grava_Indices intIndexador
                    Else
                        intIndexador = glngPegaUltimaChave(gstrIndexadorEconomico, "PKId")
                        Grava_Indices intIndexador
                    End If
                    Novo_Indexador
                End If
            End If
            MantemForm gstrRefresh
            
        Case UCase(gstrDeletar)
            Deleta_Indices intIndexador
            If ToolBarGeral(strModoOperacao, gstrIndexadorEconomico, mblnAlterando, tdb_IndexadorEconomico, Me, mobjAux, "PKId, PKId, strDescricaoIndexador, strSiglaIndexador") Then
                Novo_Indexador
            End If
            MantemForm gstrRefresh
        Case Else
            ToolBarGeral strModoOperacao, gstrIndexadorEconomico, mblnAlterando, tdb_IndexadorEconomico, Me, mobjAux, strQuery
    End Select
    
    If strModoOperacao = UCase(gstrImprimir) Then
        ImprimeRelatorio rptCadIndexadorEconomico, strQuery, "Indexador Econômico"
    End If
    
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrDeletar

   
End Sub

Private Sub tlb_Indices_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error GoTo Err_tlb_Indices_ButtonClick
    Select Case UCase(Button.Key)
        Case gstrSalvar
            If Not gblnDataValida(txt_dtmData) Then
                ExibeMensagem "Data inválida."
                txt_dtmData.SetFocus
                Exit Sub
            ElseIf Trim(txt_dblValor) = "" Or Trim(txt_dblValor) = "," Then
                ExibeMensagem "Valor inválido."
                txt_dblValor.SetFocus
                Exit Sub
            End If
            If Trim(txt_intDivisao) = "," Or Trim(txt_intDivisao) = "0" Or Trim(txt_intDivisao) = "" Then
                txt_intDivisao = ""
            ElseIf CDbl(gvntConvVrDoSql(txt_intDivisao)) <= 0 Then
                txt_intDivisao = ""
            Else
                txt_intDivisao = gvntConvVrDoSql(txt_intDivisao)
            End If
            If mblnAlterandoIndice = False Then
                For giContador = 1 To lvw_Indices.ListItems.Count
                    If gstrDataFormatada(txt_dtmData) = gstrDataFormatada(lvw_Indices.ListItems(giContador).Text) Then
                        ExibeMensagem "Índice já cadastrado para a data informada."
                        txt_dtmData.SetFocus
                        Exit Sub
                    End If
                Next
                Set objList = lvw_Indices.ListItems.Add(, , gstrDataFormatada(txt_dtmData))
                objList.SubItems(1) = gvntConvVrDoSql(txt_dblValor, 5)
                objList.SubItems(2) = Trim(txt_intDivisao)
                
                Call gblnEncontroItemNoListView(lvw_Indices, gstrDataFormatada(txt_dtmData), lvwText)
            Else
                lvw_Indices.SelectedItem.Text = gstrDataFormatada(txt_dtmData)
                lvw_Indices.SelectedItem.SubItems(1) = gvntConvVrDoSql(txt_dblValor, 5)
                lvw_Indices.SelectedItem.SubItems(2) = Trim(txt_intDivisao)
            End If
    
        Case gstrNovo
            
        Case gstrDeletar
            If lvw_Indices.ListItems.Count = 0 Then Exit Sub
            If lvw_Indices.SelectedItem.Selected = False Then Exit Sub
            lvw_Indices.ListItems.Remove lvw_Indices.SelectedItem.Index
            
    End Select
    txt_dtmData = ""
    txt_dblValor = ""
    txt_intDivisao = ""
    txt_dtmData.Locked = False
    txt_dtmData.BackColor = &H80000005
    txt_dtmData.SetFocus
    If lvw_Indices.ListItems.Count <> 0 Then
        lvw_Indices.SelectedItem.Selected = False
    End If
    mblnAlterandoIndice = False
Err_tlb_Indices_ButtonClick:
End Sub

Private Sub txt_dtmData_GotFocus()
    MarcaCampo txt_dtmData
    tab_3DPasta.Tab = 1
End Sub

Private Sub txt_dtmData_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "D", txt_dtmData
End Sub

Private Sub txt_dtmData_LostFocus()
    txt_dtmData = gstrDataFormatada(txt_dtmData)
End Sub

Private Sub txt_dblValor_GotFocus()
    MarcaCampo txt_dblValor
    tab_3DPasta.Tab = 1
End Sub

Private Sub txt_dblValor_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "V", txt_dblValor
End Sub

Private Sub txt_dblValor_LostFocus()
    txt_dblValor = gstrConvVrDoSql(txt_dblValor, 5)
End Sub

Private Sub txt_intDivisao_LostFocus()
    txt_intDivisao = gstrConvVrDoSql(txt_intDivisao)
End Sub

Private Sub txtstrDescricaoIndexador_GotFocus()
    MarcaCampo txtstrDescricaoIndexador
    tab_3DPasta.Tab = 0
End Sub

Sub MontaColumnHeaders()
    With lvw_Indices
        .ColumnHeaders.Clear
        .ColumnHeaders.Add 1, , "Data", 1200
        .ColumnHeaders.Add 2, , "Valor", 2500
        .ColumnHeaders.Add 3, , "Divisão", 2550.047
        
    End With
End Sub

Function strQueryRelatorio() As String
Dim strSQL As String
   
   strSQL = ""
   
   strSQL = strSQL & "select * from " & gstrIndexadorEconomico
   
   If mblnAlterando = True Then
      strSQL = strSQL & " WHERE PKId = " & Val(txtPKId)
   End If
    
 ' TIMTIM - 11/02/2003 - Pendência nº 5
 ' strSql = strSql & " ORDER BY strDescricaoIndexador"
 
   Select Case bytOrdenacao
      
      Case Is = 0
         strSQL = strSQL & " ORDER BY PKId" & IIf(blnOrdenacaoAsc, " ASC", " DESC")
         
      Case Is = 1
         strSQL = strSQL & " ORDER BY strDescricaoIndexador" & IIf(blnOrdenacaoAsc, " ASC", " DESC")
      
      Case Is = 2
         strSQL = strSQL & " ORDER BY strSiglaIndexador" & IIf(blnOrdenacaoAsc, " ASC", " DESC")
         
   End Select
   
   strQueryRelatorio = strSQL
   
End Function

Private Sub CarregaListViewIndices(intIndexador As Integer)
    Dim adoResultado As ADODB.Recordset
    Dim strSQL       As String
    
    strSQL = ""
    strSQL = strSQL & "Select IE.dtmData Data, IE.dblValor Valor, IE.intDivisao Divisao"
    strSQL = strSQL & " From " & gstrIndiceEconomico & " IE "
    strSQL = strSQL & "Where IE.intIndexador = " & intIndexador & " "
    strSQL = strSQL & "Order By Data Desc"

    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
        With adoResultado
            lvw_Indices.ListItems.Clear
            Do While .EOF = False
                Set objList = lvw_Indices.ListItems.Add(, , gstrVerificaCampoNulo(!Data))
                objList.SubItems(1) = gvntConvVrDoSql(!Valor, 4)
                objList.SubItems(2) = gvntConvVrDoSql(!Divisao, 2)
                .MoveNext
            Loop
        End With
        Set gobjBanco = Nothing
        adoResultado.Close
        Set adoResultado = Nothing
    End If
End Sub

Sub Grava_Indices(intIndexador As Integer)

'******************************************************************************************
' Data: 14/04/2003
' Alteração: - Substituição do comando nativo GETDATE() do SQL Server pela variável
'            strGETDATE.
' Responsável: Everton Bianchini
'******************************************************************************************
    
    Dim strSQL     As String
    Dim strDivisao As String
    
    On Error GoTo err_Grava_Indices
    
    Deleta_Indices intIndexador
    
    With lvw_Indices
        For giContador = 1 To .ListItems.Count
            strDivisao = gstrConvVrParaSql(Trim(.ListItems(giContador).SubItems(2)))
            strSQL = ""
            strSQL = strSQL & "Insert Into " & gstrIndiceEconomico & " "
            strSQL = strSQL & "(intIndexador, dtmData, dblValor, intDivisao,"
            strSQL = strSQL & "dtmDtAtualizacao, lngCodUsr "
            strSQL = strSQL & ") Values ("
            strSQL = strSQL & intIndexador & ", "
            strSQL = strSQL & gstrConvDtParaSql(.ListItems(giContador).Text) & ", "
            strSQL = strSQL & gstrConvVrParaSql(.ListItems(giContador).SubItems(1)) & ", "
            strSQL = strSQL & strDivisao & ", "
'            strSql = strSql & "GETDATE()" & ", "
            strSQL = strSQL & strGETDATE & ", "
            strSQL = strSQL & glngCodUsr
            strSQL = strSQL & ")"
            Set gobjBanco = New clsBanco
            gobjBanco.Execute strSQL
        Next
    End With
err_Grava_Indices:
End Sub

Sub Deleta_Indices(intIndexador As Integer)
    Dim strSQL As String

    strSQL = ""
    strSQL = strSQL & "Delete From " & gstrIndiceEconomico & " "
    strSQL = strSQL & "Where intIndexador = " & intIndexador
    
    Set gobjBanco = New clsBanco
    gobjBanco.Execute strSQL
End Sub

Private Function blnDadosOk() As Boolean
blnDadosOk = False

If Trim(txtstrDescricaoIndexador) = "" Then
    ExibeMensagem "A descrição do indexador tem que ser digitada."
    txtstrDescricaoIndexador.SetFocus
    Exit Function
ElseIf Trim(txtstrSiglaIndexador) = "" Then
    ExibeMensagem "A sigla do indexador tem que ser digitada."
    txtstrSiglaIndexador.SetFocus
    Exit Function
End If
    
If txt_dtmData.Text <> "" Then
    If gblnDataValida(txt_dtmData.Text) = False Then
        ExibeMensagem "Data inválida."
        txt_dtmData.SetFocus
        Exit Function
    End If
End If
    
If Not mblnAlterando Or (mblnAlterando And UCase$(txtstrDescricaoIndexador.Text) <> UCase$(strDescriAtual)) Then
    If gblnExisteCodigo(1, gstrIndexadorEconomico, "STRDESCRICAOINDEXADOR", txtstrDescricaoIndexador) Then
        ExibeMensagem "A descrição informada já se encontra cadastrada."
        txtstrDescricaoIndexador.SetFocus
        Exit Function
    End If
End If

If Not mblnAlterando Or (mblnAlterando And UCase$(txtstrSiglaIndexador.Text) <> UCase$(strSiglaAtual)) Then
    If gblnExisteCodigo(1, gstrIndexadorEconomico, "STRSIGLAINDEXADOR", txtstrSiglaIndexador) Then
        ExibeMensagem "A sigla informada já se encontra cadastrada."
        txtstrSiglaIndexador.SetFocus
        Exit Function
    End If
End If
    
blnDadosOk = True
End Function

Sub Novo_Indexador()
    lvw_Indices.ListItems.Clear
    mblnAlterandoIndice = False
    txt_dtmData = ""
    txt_dtmData.Locked = False
    txt_dtmData.BackColor = &H80000005
    txt_dblValor = ""
    txt_intDivisao = ""
    tab_3DPasta.Tab = 0
    
End Sub

Private Sub txtstrDescricaoIndexador_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrDescricaoIndexador
End Sub

Private Sub txt_intDivisao_GotFocus()
    MarcaCampo txt_intDivisao
    tab_3DPasta.Tab = 1
End Sub

Private Sub txt_intDivisao_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "V", txt_intDivisao
End Sub

Private Sub txtstrSiglaIndexador_GotFocus()
    MarcaCampo txtstrSiglaIndexador
    tab_3DPasta.Tab = 0
End Sub

Private Sub txtstrSiglaIndexador_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrSiglaIndexador
End Sub

