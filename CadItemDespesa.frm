VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frmCadItemDespesa 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Itens de Despesa"
   ClientHeight    =   3555
   ClientLeft      =   495
   ClientTop       =   3270
   ClientWidth     =   8460
   HelpContextID   =   44
   Icon            =   "CadItemDespesa.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3555
   ScaleWidth      =   8460
   Begin VB.TextBox txtPKId 
      Height          =   270
      Left            =   6870
      TabIndex        =   13
      Top             =   30
      Visible         =   0   'False
      Width           =   645
   End
   Begin TabDlg.SSTab tab_3dPasta 
      Height          =   3375
      Left            =   90
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   60
      Width           =   8280
      _ExtentX        =   14605
      _ExtentY        =   5953
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Itens de Despesa"
      TabPicture(0)   =   "CadItemDespesa.frx":1042
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblstrCodigo"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblstrDescricao"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "txtstrCodigo"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txtstrDescricao"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "fra_Lancamento"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "tdb_Lista"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      Begin TrueOleDBGrid70.TDBGrid tdb_Lista 
         Height          =   2085
         Left            =   120
         TabIndex        =   9
         Top             =   1050
         Width           =   8055
         _ExtentX        =   14208
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
         Columns(1).Caption=   "Código"
         Columns(1).DataField=   "strCodigo"
         Columns(1).NumberFormat=   "FormatText Event"
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
         Splits(0)._ColumnProps(1)=   "Column(0).Width=1138"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1058"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0).AllowSizing=0"
         Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
         Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(8)=   "Column(1).Width=2831"
         Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=2752"
         Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(13)=   "Column(2).Width=10821"
         Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=10742"
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
         AllowUpdate     =   0   'False
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
         _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
         _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
         _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
         _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
         _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=46,.parent=13"
         _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=43,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=44,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=45,.parent=17"
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
      Begin VB.Frame fra_Lancamento 
         Caption         =   " Lançamento "
         Height          =   1605
         Left            =   120
         TabIndex        =   14
         Top             =   1050
         Visible         =   0   'False
         Width           =   8055
         Begin VB.Frame fra_bytTipo 
            Caption         =   " Tipo "
            Height          =   525
            Left            =   120
            TabIndex        =   19
            Top             =   990
            Width           =   4155
            Begin VB.OptionButton optbytTipo 
               Caption         =   "Diversos"
               Height          =   195
               Index           =   0
               Left            =   90
               TabIndex        =   6
               Top             =   240
               Value           =   -1  'True
               Width           =   975
            End
            Begin VB.OptionButton optbytTipo 
               Caption         =   "Mutação"
               CausesValidation=   0   'False
               Height          =   195
               Index           =   2
               Left            =   3090
               TabIndex        =   8
               Top             =   240
               Width           =   1005
            End
            Begin VB.OptionButton optbytTipo 
               Caption         =   "Independente"
               Height          =   195
               Index           =   1
               Left            =   1410
               TabIndex        =   7
               Top             =   240
               Width           =   1335
            End
         End
         Begin VB.CommandButton cmd_ContaCredito 
            Height          =   315
            Left            =   7590
            Picture         =   "CadItemDespesa.frx":105E
            Style           =   1  'Graphical
            TabIndex        =   16
            TabStop         =   0   'False
            Tag             =   "322"
            ToolTipText     =   "Clique aqui para cadastrar conta"
            Top             =   240
            Width           =   360
         End
         Begin VB.CommandButton cmd_ContaDebito 
            Height          =   315
            Left            =   7590
            Picture         =   "CadItemDespesa.frx":13E8
            Style           =   1  'Graphical
            TabIndex        =   15
            TabStop         =   0   'False
            Tag             =   "322"
            ToolTipText     =   "Clique aqui para cadastrar conta"
            Top             =   630
            Width           =   360
         End
         Begin VB.ComboBox cbo_ContaContabilCredito 
            Height          =   315
            Left            =   630
            OLEDragMode     =   1  'Automatic
            Sorted          =   -1  'True
            TabIndex        =   2
            ToolTipText     =   "Crédito"
            Top             =   240
            Width           =   1665
         End
         Begin VB.ComboBox cbointContaCredito 
            Height          =   315
            Left            =   2280
            OLEDragMode     =   1  'Automatic
            Sorted          =   -1  'True
            TabIndex        =   3
            ToolTipText     =   "Histórico padrão"
            Top             =   240
            Width           =   5325
         End
         Begin VB.ComboBox cbo_ContaContabilDebito 
            Height          =   315
            Left            =   630
            OLEDragMode     =   1  'Automatic
            Sorted          =   -1  'True
            TabIndex        =   4
            ToolTipText     =   "Débito"
            Top             =   630
            Width           =   1665
         End
         Begin VB.ComboBox cbointContaDebito 
            Height          =   315
            Left            =   2280
            OLEDragMode     =   1  'Automatic
            Sorted          =   -1  'True
            TabIndex        =   5
            ToolTipText     =   "Histórico padrão"
            Top             =   630
            Width           =   5325
         End
         Begin VB.Label lblintContaCredito 
            AutoSize        =   -1  'True
            Caption         =   "Crédito"
            Height          =   195
            Left            =   60
            TabIndex        =   18
            Top             =   285
            Width           =   495
         End
         Begin VB.Label lblintContaDebito 
            AutoSize        =   -1  'True
            Caption         =   "Débito"
            Height          =   195
            Left            =   90
            TabIndex        =   17
            Top             =   660
            Width           =   465
         End
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
         Left            =   960
         MaxLength       =   50
         TabIndex        =   1
         Top             =   720
         Width           =   7200
      End
      Begin VB.TextBox txtstrCodigo 
         Alignment       =   1  'Right Justify
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
         Left            =   960
         MaxLength       =   18
         MultiLine       =   -1  'True
         TabIndex        =   0
         Top             =   375
         Width           =   1680
      End
      Begin VB.Label lblstrDescricao 
         AutoSize        =   -1  'True
         Caption         =   "Descrição"
         Height          =   195
         Left            =   150
         TabIndex        =   12
         Top             =   765
         Width           =   720
      End
      Begin VB.Label lblstrCodigo 
         AutoSize        =   -1  'True
         Caption         =   "Código"
         Height          =   195
         Left            =   375
         TabIndex        =   11
         Top             =   420
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmCadItemDespesa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim mblnAlterando         As Boolean
    Dim mobjAux               As Object
    Dim mblnClickOk           As Boolean
    Dim ValTxtStrCodigo       As String
    Dim ValTxtStrDescricao    As String
    Dim mblnCarregaFormConta  As Boolean


Private Sub cbointContaCredito_Click()
    
Dim tempIndice As Integer

    tempIndice = cbointContaCredito.ListIndex
    cbo_ContaContabilCredito.ListIndex = gintIndiceCBO(cbo_ContaContabilCredito, _
                                  gstrItemData(cbointContaCredito))
                                  
   If cbo_ContaContabilCredito.ListIndex = -1 Then
        LePlanoContaGeral1 cbo_ContaContabilCredito, cbointContaCredito, cbo_ContaContabilDebito, cbointContaDebito
        cbointContaCredito.ListIndex = tempIndice
        cbo_ContaContabilCredito.ListIndex = gintIndiceCBO(cbo_ContaContabilCredito, _
                                  gstrItemData(cbointContaCredito))
   End If

End Sub

Private Sub cbo_ContaContabilCredito_Click()
    cbointContaCredito.ListIndex = gintIndiceCBO(cbointContaCredito, _
                              gstrItemData(cbo_ContaContabilCredito))
End Sub

Private Sub cbointContaCredito_GotFocus()
    If mblnCarregaFormConta = True Then
        mblnCarregaFormConta = False
        If cbointContaCredito.ListIndex = -1 Then cbo_ContaContabilCredito.ListIndex = -1
    End If
End Sub

Private Sub cbointContaDebito_Click()

Dim tempIndice As Integer

    tempIndice = cbointContaDebito.ListIndex
    cbo_ContaContabilDebito.ListIndex = gintIndiceCBO(cbo_ContaContabilDebito, _
                                  gstrItemData(cbointContaDebito))
                                  
   If cbo_ContaContabilDebito.ListIndex = -1 Then
        LePlanoContaGeral1 cbo_ContaContabilCredito, cbointContaCredito, cbo_ContaContabilDebito, cbointContaDebito
        cbointContaDebito.ListIndex = tempIndice
    cbo_ContaContabilDebito.ListIndex = gintIndiceCBO(cbo_ContaContabilDebito, _
                                  gstrItemData(cbointContaDebito))
   End If
   
End Sub

Private Sub cbo_ContaContabilDebito_Click()
    cbointContaDebito.ListIndex = gintIndiceCBO(cbointContaDebito, _
                              gstrItemData(cbo_ContaContabilDebito))
End Sub

Private Sub cbointContaCredito_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", cbointContaCredito
End Sub

Private Sub cbointContaDebito_GotFocus()
    If mblnCarregaFormConta = True Then
        mblnCarregaFormConta = False
        If cbointContaDebito.ListIndex = -1 Then cbo_ContaContabilDebito.ListIndex = -1
    End If
End Sub

Private Sub cmd_ContaCredito_Click()
    mblnCarregaFormConta = True
    CarregaForm frmCadPlanoConta, cbointContaCredito, strQueryPlanoConta
End Sub

Private Sub cmd_ContaDebito_Click()
    mblnCarregaFormConta = True
    CarregaForm frmCadPlanoConta, cbointContaDebito, strQueryPlanoConta
End Sub

Private Sub cbointContaDebito_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", cbointContaDebito
End Sub

Private Sub Form_Activate()
    
    gintCodSeguranca = 244
    
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
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar
End Sub

Private Function strQuery() As String

Dim strSQL  As String

    strSQL = "SELECT PKId, strCodigo, strDescricao FROM "
    strSQL = strSQL & gstrItemDespesa & " ORDER BY strCodigo"
    
    strQuery = strSQL
    
End Function

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = vbKeyF1 Then
        Call_HtmlHelp Me.HelpContextID
    End If
End Sub

Private Sub Form_Load()
    VerificaListaAutomatica gstrItemDespesa, tdb_Lista, strQuery
    VerificaObjParaAplicar mobjAux
    mblnAlterando = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrImprimir
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar
End Sub

Private Sub optbytTipo_KeyPress(Index As Integer, KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub tdb_Lista_Click()
    If glngQtdLinhaTDBGrid(tdb_Lista) = 1 Then
        tdb_Lista_RowColChange 0, 0
    End If
End Sub

Private Sub tdb_Lista_DblClick()
    MantemForm gstrAplicar
End Sub

Private Sub tdb_Lista_FilterChange()
    mblnClickOk = False
    gblnFilraCampos tdb_Lista
End Sub

Private Sub tdb_lista_FormatText(ByVal ColIndex As Integer, Value As Variant, Bookmark As Variant)
    If ColIndex = 1 Then
        Value = gvntFormatacaoEspecifica(Value, 4)
    End If
End Sub

Private Sub tdb_Lista_HeadClick(ByVal ColIndex As Integer)
    gOrdenaGrid tdb_Lista, ColIndex
End Sub

Private Sub tdb_Lista_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", tdb_Lista
End Sub

Private Sub LimpaCombo()
    cbo_ContaContabilCredito = ""
    cbo_ContaContabilDebito = ""
End Sub

Private Sub tdb_Lista_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mblnClickOk = True
End Sub

Private Sub tdb_Lista_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    
    With tdb_Lista
        If (Not .EOF And Not .BOF) And mblnClickOk Then
            mblnClickOk = False
            txtPKId.Text = .Columns("PKID").Value
            LeDaTabelaParaObj gstrItemDespesa, Me
            txtstrCodigo = gvntFormatacaoEspecifica(txtstrCodigo, 4)
            gCorLinhaSelecionada tdb_Lista
            If mobjAux Is Nothing Then
                HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar
            Else
                HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrAplicar
            End If
            mblnAlterando = True
        End If
    End With
    
    ValTxtStrCodigo = txtstrCodigo.Text
    ValTxtStrDescricao = txtstrDescricao.Text

End Sub

Private Sub GravaItemDespesa()

Dim strSQL As String
Dim bytTipo As Integer
    
    If optbytTipo(0).Value = True Then bytTipo = 0
    If optbytTipo(1).Value = True Then bytTipo = 1
    If optbytTipo(2).Value = True Then bytTipo = 2
    
    If blnDadosOk = True Then
    If mblnAlterando Then
        If gblnExclusaoGravacaoOk("A") Then
            strSQL = ""
            strSQL = strSQL & "Update " & gstrItemDespesa & " SET "
            strSQL = strSQL & "strCodigo ='" & ValorParaGravar & "', "
            strSQL = strSQL & "strDescricao='" & txtstrDescricao & "', "
            strSQL = strSQL & "bytTipo=" & bytTipo & ", "
            strSQL = strSQL & "dtmDtAtualizacao=" & gstrConvDtParaSql(gstrDataDoSistema) & ", "
            strSQL = strSQL & "lngCodUsr=" & glngCodUsr
            strSQL = strSQL & " WHERE PKID=" & txtPKId
        Else
            Exit Sub
        End If
    Else
        If gblnExclusaoGravacaoOk("I") Then
            strSQL = ""
            strSQL = strSQL & "INSERT INTO " & gstrItemDespesa & " ("
            strSQL = strSQL & "strCodigo , strDescricao, bytTipo, dtmDtAtualizacao, lngCodUsr) "
            strSQL = strSQL & " VALUES ('" & ValorParaGravar & "', '"
            strSQL = strSQL & txtstrDescricao & "', " & bytTipo & ", "
            strSQL = strSQL & gstrConvDtParaSql(gstrDataDoSistema) & ", "
            strSQL = strSQL & glngCodUsr & ")"
        Else
            Exit Sub
        End If
    End If
        
        Set gobjBanco = New clsBanco
        gobjBanco.Execute (strSQL)
        Set gobjBanco = Nothing
        MantemForm gstrNovo
        MantemForm gstrLocalizar
        
    End If
End Sub

Public Sub MantemForm(ByVal strModoOperacao As String)
    
    If strModoOperacao = gstrSalvar Then
        GravaItemDespesa
        Exit Sub
    End If
    
    If strModoOperacao = gstrPreencherLista Then
        LePlanoContaGeral1 cbo_ContaContabilCredito, cbointContaCredito, cbo_ContaContabilDebito, cbointContaDebito
    End If
    
    If ToolBarGeral(strModoOperacao, gstrItemDespesa, mblnAlterando, tdb_Lista, Me, mobjAux, strQuery, "PKId, strDescricao", rptItemDeDespesa, strQuerryRelatorio) Then
        LimpaCombo
    End If
    
End Sub

Function strQueryPlanoConta() As String

Dim strSQL          As String
    
    strSQL = "SELECT PKId, strDescricao "
    strSQL = strSQL & "FROM "
    strSQL = strSQL & gstrPlanoConta & " "
    strSQL = strSQL & "WHERE ABS(blnAnalitica) = 1"
    
    strQueryPlanoConta = strSQL
    
End Function

Private Sub cbo_ContaContabilCredito_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", cbo_ContaContabilCredito
End Sub

Private Sub cbo_ContaContabilDebito_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", cbo_ContaContabilDebito
End Sub

Private Sub txtstrCodigo_LostFocus()
    txtstrCodigo = gvntFormatacaoEspecifica(txtstrCodigo, 4)
End Sub

Private Sub txtstrDescricao_GotFocus()
    MarcaCampo txtstrDescricao
End Sub

Private Sub txtstrDescricao_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrDescricao
End Sub

Private Sub txtstrCodigo_GotFocus()
    txtstrCodigo = gstrValorSemMascara(txtstrCodigo)
    MarcaCampo txtstrCodigo
End Sub

Private Sub txtstrCodigo_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrCodigo
End Sub

Private Function strQuerryRelatorio() As String

Dim strSQL As String
        
    strSQL = " SELECT "
    strSQL = strSQL & " strCodigo, strDescricao, "
    strSQL = strSQL & gstrCASEWHEN("bytTipo", "1, 'Independente', 2, 'Mutação'") & " AS TIPO "
    strSQL = strSQL & " FROM "
    strSQL = strSQL & gstrItemDespesa
    strSQL = strSQL & " ORDER BY strDescricao "

    strQuerryRelatorio = strSQL

End Function

Private Function ValorParaGravar() As String
    
Dim strMascara     As String
Dim vntValor       As String
Dim intInd         As Integer
Dim intLenMascara  As Integer
    
    strMascara = gstrMascaraElementoDespesa
    
    vntValor = gstrValorSemMascara(txtstrCodigo)
    
    If Trim(strMascara) <> "" Then
        For intInd = 1 To Len(Trim(strMascara))
            If Mid(strMascara, intInd, 1) = "0" Then
                intLenMascara = intLenMascara + 1
            End If
        Next
        If Len(gstrENulo(vntValor)) < 15 Then
            intLenMascara = 15 - Len(gstrENulo(vntValor))
        ElseIf Len(gstrENulo(vntValor)) > 15 Then
            intLenMascara = Len(gstrENulo(vntValor)) - 15
        Else
            intLenMascara = 0
        End If
        ValorParaGravar = Trim(vntValor) & String(intLenMascara, "0")
    End If
    
End Function

Private Function blnDadosOk() As Boolean

    If txtstrCodigo.Text = "" Then
        ExibeMensagem "O campo Código devers er preenchido"
        txtstrCodigo.SetFocus
        Exit Function
    ElseIf txtstrDescricao.Text = "" Then
        ExibeMensagem "O campo Descrição dever ser preenchido!"
        txtstrDescricao.SetFocus
        Exit Function
    ElseIf mblnAlterando = False And gblnExisteCodigo(1, gstrItemDespesa, "strCodigo", ValorParaGravar) Then
        ExibeMensagem "A código digitado já se encontra cadastrado."
        txtstrCodigo.SetFocus
        Exit Function
    ElseIf mblnAlterando = False And gblnExisteCodigo(1, gstrItemDespesa, "strDescricao", "'" & txtstrDescricao.Text & "'") Then
        ExibeMensagem "A descrição digitada já se encontra cadastrada."
        txtstrDescricao.SetFocus
        Exit Function
    
    'Verifica se foi alterado
    
    ElseIf txtstrCodigo <> ValTxtStrCodigo Then
        If mblnAlterando And gblnExisteCodigo(1, gstrItemDespesa, "strCodigo", ValorParaGravar) Then
            ExibeMensagem "A código digitado já se encontra cadastrado."
            txtstrCodigo.SetFocus
            Exit Function
        End If
    
    ElseIf UCase(txtstrDescricao) <> UCase(ValTxtStrDescricao) Then
        If mblnAlterando And gblnExisteCodigo(1, gstrItemDespesa, "strDescricao", "'" & txtstrDescricao.Text & "'") Then
            ExibeMensagem "A descrição digitada já se encontra cadastrada."
            txtstrDescricao.SetFocus
            Exit Function
        End If
    
    End If

    blnDadosOk = True

End Function
