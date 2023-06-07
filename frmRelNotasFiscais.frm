VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MsDatLst.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frmRelNotasFiscais 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Notas Fiscais - Emissão"
   ClientHeight    =   4920
   ClientLeft      =   1860
   ClientTop       =   2475
   ClientWidth     =   8445
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   8445
   Begin TabDlg.SSTab tab_3dPasta 
      Height          =   4875
      Left            =   60
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   30
      Width           =   8385
      _ExtentX        =   14790
      _ExtentY        =   8599
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Notas Fiscais"
      TabPicture(0)   =   "frmRelNotasFiscais.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fra_Inscricao"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin VB.Frame fra_Inscricao 
         Caption         =   "Inscrição"
         Height          =   4395
         Left            =   120
         TabIndex        =   1
         Top             =   390
         Width           =   8115
         Begin VB.TextBox txtCancelamento 
            Height          =   315
            Left            =   7260
            TabIndex        =   30
            Top             =   990
            Visible         =   0   'False
            Width           =   750
         End
         Begin VB.TextBox txtFantasia 
            Height          =   315
            Left            =   1215
            TabIndex        =   13
            Top             =   1350
            Width           =   6810
         End
         Begin VB.TextBox txtTelefone 
            Height          =   315
            Left            =   4410
            TabIndex        =   11
            Top             =   990
            Width           =   2340
         End
         Begin VB.TextBox txtCNPJ 
            Height          =   315
            Left            =   1215
            TabIndex        =   9
            Top             =   990
            Width           =   2340
         End
         Begin VB.TextBox txtEndereco 
            Height          =   315
            Left            =   1215
            TabIndex        =   7
            Top             =   630
            Width           =   6810
         End
         Begin VB.TextBox txtSerie 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   6570
            MaxLength       =   5
            TabIndex        =   29
            Top             =   3990
            Width           =   720
         End
         Begin VB.TextBox txtControleFinal 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   4890
            TabIndex        =   23
            Top             =   3660
            Width           =   1140
         End
         Begin VB.TextBox txtControleInicial 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2280
            TabIndex        =   21
            Top             =   3660
            Width           =   1140
         End
         Begin VB.TextBox txtQuantidade 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   4890
            MaxLength       =   5
            TabIndex        =   19
            Top             =   3330
            Width           =   1140
         End
         Begin VB.TextBox txtDTBase 
            Height          =   285
            Left            =   2280
            TabIndex        =   17
            Top             =   3330
            Width           =   1140
         End
         Begin VB.TextBox txtNotaFinal 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   4890
            TabIndex        =   27
            Top             =   3990
            Width           =   1140
         End
         Begin VB.TextBox txtNotaInicial 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2280
            MaxLength       =   10
            TabIndex        =   25
            Top             =   3990
            Width           =   1140
         End
         Begin VB.TextBox txtStrRazao 
            Height          =   315
            Left            =   3870
            TabIndex        =   5
            Top             =   270
            Width           =   4140
         End
         Begin VB.Frame fra_Lancamentos 
            Caption         =   "Atividades "
            Height          =   1485
            Left            =   90
            TabIndex        =   14
            Top             =   1770
            Width           =   7860
            Begin TrueOleDBGrid70.TDBGrid tdb_Atividade 
               Height          =   1140
               Left            =   90
               TabIndex        =   15
               Top             =   240
               Width           =   7665
               _ExtentX        =   13520
               _ExtentY        =   2011
               _LayoutType     =   4
               _RowHeight      =   -2147483647
               _WasPersistedAsPixels=   0
               Columns(0)._VlistStyle=   0
               Columns(0)._MaxComboItems=   5
               Columns(0).Caption=   "Atividade"
               Columns(0).DataField=   "strDescricao"
               Columns(0).NumberFormat=   "Short Date"
               Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(1)._VlistStyle=   512
               Columns(1)._MaxComboItems=   5
               Columns(1).ValueItems(0)._DefaultItem=   0
               Columns(1).ValueItems(0).Value=   "0"
               Columns(1).ValueItems(0).Value.vt=   8
               Columns(1).ValueItems(0).DisplayValue=   "Secundária"
               Columns(1).ValueItems(0).DisplayValue.vt=   8
               Columns(1).ValueItems(0)._PropDict=   "_DefaultItem,517,2"
               Columns(1).ValueItems(1)._DefaultItem=   0
               Columns(1).ValueItems(1).Value=   "1"
               Columns(1).ValueItems(1).Value.vt=   8
               Columns(1).ValueItems(1).DisplayValue=   "Primária"
               Columns(1).ValueItems(1).DisplayValue.vt=   8
               Columns(1).ValueItems(1)._PropDict=   "_DefaultItem,517,2"
               Columns(1).ValueItems.Count=   2
               Columns(1).Caption=   "Principal / Secundária"
               Columns(1).DataField=   "Status"
               Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns.Count   =   2
               Splits(0)._UserFlags=   0
               Splits(0).MarqueeStyle=   3
               Splits(0).RecordSelectors=   0   'False
               Splits(0).RecordSelectorWidth=   503
               Splits(0)._SavedRecordSelectors=   0   'False
               Splits(0).ScrollBars=   2
               Splits(0).DividerColor=   12632256
               Splits(0).SpringMode=   0   'False
               Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
               Splits(0)._ColumnProps(0)=   "Columns.Count=2"
               Splits(0)._ColumnProps(1)=   "Column(0).Width=9684"
               Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
               Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=9604"
               Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
               Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
               Splits(0)._ColumnProps(6)=   "Column(1).Width=3016"
               Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
               Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2937"
               Splits(0)._ColumnProps(9)=   "Column(1)._EditAlways=0"
               Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
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
               _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=38,.bold=0,.fontsize=825,.italic=0"
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
               _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=38,.bgcolor=&H8000000D&"
               _StyleDefs(18)  =   ":id=6,.fgcolor=&H8000000E&"
               _StyleDefs(19)  =   "EditorStyle:id=7,.parent=1"
               _StyleDefs(20)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38,.bgcolor=&H8000000D&"
               _StyleDefs(21)  =   ":id=8,.fgcolor=&H8000000E&"
               _StyleDefs(22)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
               _StyleDefs(23)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
               _StyleDefs(24)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
               _StyleDefs(25)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
               _StyleDefs(26)  =   "Splits(0).Style:id=13,.parent=1"
               _StyleDefs(27)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
               _StyleDefs(28)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
               _StyleDefs(29)  =   "Splits(0).FooterStyle:id=15,.parent=3"
               _StyleDefs(30)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
               _StyleDefs(31)  =   "Splits(0).SelectedStyle:id=18,.parent=6,.namedParent=38,.bgcolor=&H8000000D&"
               _StyleDefs(32)  =   ":id=18,.fgcolor=&H8000000E&"
               _StyleDefs(33)  =   "Splits(0).EditorStyle:id=17,.parent=7"
               _StyleDefs(34)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8,.namedParent=38,.bgcolor=&H8000000D&"
               _StyleDefs(35)  =   ":id=19,.fgcolor=&H8000000E&"
               _StyleDefs(36)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
               _StyleDefs(37)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
               _StyleDefs(38)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
               _StyleDefs(39)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
               _StyleDefs(40)  =   "Splits(0).Columns(0).Style:id=32,.parent=13"
               _StyleDefs(41)  =   "Splits(0).Columns(0).HeadingStyle:id=29,.parent=14"
               _StyleDefs(42)  =   "Splits(0).Columns(0).FooterStyle:id=30,.parent=15"
               _StyleDefs(43)  =   "Splits(0).Columns(0).EditorStyle:id=31,.parent=17"
               _StyleDefs(44)  =   "Splits(0).Columns(1).Style:id=50,.parent=13"
               _StyleDefs(45)  =   "Splits(0).Columns(1).HeadingStyle:id=47,.parent=14"
               _StyleDefs(46)  =   "Splits(0).Columns(1).FooterStyle:id=48,.parent=15"
               _StyleDefs(47)  =   "Splits(0).Columns(1).EditorStyle:id=49,.parent=17"
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
         End
         Begin MSDataListLib.DataCombo dbc_strInscricao 
            Height          =   315
            Left            =   1215
            TabIndex        =   3
            Top             =   270
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin VB.Label lblFantasia 
            AutoSize        =   -1  'True
            Caption         =   "Nome fantasia"
            Height          =   195
            Left            =   90
            TabIndex        =   12
            Top             =   1410
            Width           =   1020
         End
         Begin VB.Label lblTelefone 
            AutoSize        =   -1  'True
            Caption         =   "Telefone"
            Height          =   195
            Left            =   3660
            TabIndex        =   10
            Top             =   1050
            Width           =   630
         End
         Begin VB.Label lblCNPJ 
            AutoSize        =   -1  'True
            Caption         =   "CNPJ/CPF"
            Height          =   195
            Left            =   90
            TabIndex        =   8
            Top             =   1050
            Width           =   780
         End
         Begin VB.Label lblEndereco 
            AutoSize        =   -1  'True
            Caption         =   "Endereço"
            Height          =   195
            Left            =   90
            TabIndex        =   6
            Top             =   690
            Width           =   690
         End
         Begin VB.Label lblSerie 
            AutoSize        =   -1  'True
            Caption         =   "Série"
            Height          =   195
            Left            =   6150
            TabIndex        =   28
            Top             =   4080
            Width           =   360
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Controle Final"
            Height          =   195
            Left            =   3870
            TabIndex        =   22
            Top             =   3750
            Width           =   960
         End
         Begin VB.Label lblControle1 
            AutoSize        =   -1  'True
            Caption         =   "Controle Inicial"
            Height          =   195
            Left            =   1170
            TabIndex        =   20
            Top             =   3750
            Width           =   1035
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Quantidade"
            Height          =   195
            Left            =   4005
            TabIndex        =   18
            Top             =   3420
            Width           =   825
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Data Base"
            Height          =   195
            Left            =   1455
            TabIndex        =   16
            Top             =   3420
            Width           =   750
         End
         Begin VB.Label lblNotaFinal 
            AutoSize        =   -1  'True
            Caption         =   "Nota Fiscal Final"
            Height          =   195
            Left            =   3660
            TabIndex        =   26
            Top             =   4080
            Width           =   1170
         End
         Begin VB.Label lblNotaInicial 
            AutoSize        =   -1  'True
            Caption         =   "Nota Fiscal Inicial"
            Height          =   195
            Left            =   960
            TabIndex        =   24
            Top             =   4080
            Width           =   1245
         End
         Begin VB.Label lblRazao 
            AutoSize        =   -1  'True
            Caption         =   "Razão Social"
            Height          =   195
            Left            =   2820
            TabIndex        =   4
            Top             =   330
            Width           =   975
         End
         Begin VB.Label lblInscricao 
            AutoSize        =   -1  'True
            Caption         =   "Inscrição"
            Height          =   195
            Left            =   90
            TabIndex        =   2
            Top             =   330
            Width           =   645
         End
      End
   End
End
Attribute VB_Name = "frmRelNotasFiscais"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mobjAux             As Object
Dim vetNotas()          As String

Private Function ProximoNumeroNotaFiscal() As Long
    Dim strSQL       As String
    Dim adoResultado As ADODB.Recordset
    
    strSQL = "select max(" & gstrCONVERT(cdt_numeric, "strnotafiscalnr") & ") + 1 ProximoNumero "
    strSQL = strSQL & "From "
    strSQL = strSQL & "tblnotafiscaliss NF "
    
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSQL, 10, adoResultado) Then
        If Not adoResultado.EOF Then
            ProximoNumeroNotaFiscal = IIf(IsNull(adoResultado!ProximoNumero), "1", adoResultado!ProximoNumero)
        Else
            ProximoNumeroNotaFiscal = ""
        End If
    End If
    
End Function

Public Sub MantemForm(ByVal strModoOperacao As String)
    
    Select Case UCase(strModoOperacao)
        
    Case UCase(gstrImprimir)
        
        If blnDadosOK Then
            If PreencheVetorNota Then
                ImprimeRelatorioPorArray rptNotaFiscalIss, vetNotas, "Notas Fiscais"
            Else
                ExibeMensagem "Ocorreu erro na geração das notas fiscais."
            End If
        End If
        
    Case UCase(gstrNovo)
        Limpa_Controles Me, True, True, True, True, True
        txtControleInicial = strNotaFiscal
        Set tdb_Atividade.DataSource = Nothing
        dbc_strInscricao.SetFocus
    Case UCase(gstrPreencherLista)
        PreencherListaDeOpcoes Me.ActiveControl
    End Select
    
End Sub

Private Sub dbc_strInscricao_Change()
    If dbc_strInscricao.MatchedWithList Then
        '        txtStrRazao = strRazãoSocial(Val(dbc_strInscricao.BoundText))
        PreencheGrdAtividades Val(dbc_strInscricao.BoundText)
        '    Else
        '        txtStrRazao = ""
    End If
    
    strRazãoSocial
    
End Sub

Private Sub dbc_strInscricao_GotFocus()
    MarcaCampo dbc_strInscricao
End Sub

Private Sub dbc_strInscricao_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbc_strInscricao, Me, , KeyCode, Shift
End Sub

Private Sub dbc_strInscricao_KeyPress(KeyAscii As Integer)
    'TRI0748
    'CaracterValido KeyAscii, "N", dbc_strInscricao
    gstrLimitaCampoValor dbc_strInscricao, KeyAscii, 9, 0
End Sub

Private Sub dbc_strInscricao_LostFocus()
    'TRI0748
    If blnVerificaEncerrado Then Exit Sub
End Sub

Private Sub Form_Activate()
    gintCodSeguranca = 1376
    If mobjAux Is Nothing Then
        HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar
    Else
        HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrAplicar
    End If
    'TRI0748
    dbc_strInscricao.SetFocus
    dbc_strInscricao.HelpContextID = 1
    MantemForm gstrPreencherLista
    dbc_strInscricao.HelpContextID = 0
    If blnVerificaEncerrado Then Exit Sub
End Sub

Private Sub Form_Deactivate()
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar
End Sub

Private Sub Form_GotFocus()

End Sub

Private Sub Form_Initialize()

End Sub

Private Sub Form_Load()
    dbc_strInscricao.Tag = strQueryInscricao & ";Strinscricaocadastral"
    TrocaCorObjeto txtStrRazao, True
    TrocaCorObjeto txtEndereco, True
    TrocaCorObjeto txtCNPJ, True
    TrocaCorObjeto txtTelefone, True
    TrocaCorObjeto txtFantasia, True
    TrocaCorObjeto txtControleInicial, True
    TrocaCorObjeto txtControleFinal, True
    TrocaCorObjeto txtNotaFinal, True
    txtControleInicial = strNotaFiscal
End Sub

Private Sub Form_Unload(Cancel As Integer)
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar
End Sub

Private Function strQueryInscricao() As String
    Dim strSQL As String
    
    strSQL = ""
    strSQL = strSQL & "SELECT "
    strSQL = strSQL & "Pkid, " & gstrRIGHT("Strinscricaocadastral", gintRetornaTamanhoMascara(TYP_ECONOMICA)) & " Strinscricaocadastral "
    strSQL = strSQL & "FROM "
    strSQL = strSQL & gstrEconomico
    'TRI0748
    'strSQL = strSQL & " Where dtmdataencerramento Is null "
    strSQL = strSQL & " ORDER BY Strinscricaocadastral "
    
    strQueryInscricao = strSQL
    
End Function

Private Sub strRazãoSocial()
    Dim strSQL As String
    Dim adoResultado As ADODB.Recordset
    
    If dbc_strInscricao.MatchedWithList Then
        
        '        strSQL = ""
        '        strSQL = strSQL & "SELECT "
        '        strSQL = strSQL & " Ltrim(Rtrim(CO.strnome)) strnome "
        '        strSQL = strSQL & "FROM " & gstrEconomico & " EC, "
        '        strSQL = strSQL & gstrContribuinte & " CO "
        '        strSQL = strSQL & " Where EC.intContribuinte = CO.Pkid And EC.Pkid = " & dbc_strInscricao.BoundText
        
        strSQL = _
        "select " & _
        "co.strnome strnome, " & _
        "co.strcnpjcpf, " & _
        "co.strnomefantasia, " & _
        "ec.dtmdataencerramento, " & _
        "tl.strdescricao strtitulo, " & _
        "lo.strdescricao strlogradouro, " & _
        "ba.strdescricao strbairro, " & _
        "ec.intnumero, " & _
        "mu.strdescricao strmunicipio, " & _
        "uf.strsigla struf, " & _
        "ec.intcep, "
        
        If bytDBType = SQLServer Then
            strSQL = strSQL & "(select top 1 fc.strconteudo from " & gstrTipoDeComunicacao & " tc where tc.pkid = fc.inttipodecomunicacao and tc.inttipo = 6) strtelefone "
        Else
            strSQL = strSQL & "(SELECT fc.strconteudo FROM (select tc.pkid from tblTipoDeComunicacao tc where tc.inttipo = 6) tc Where tc.Pkid = fc.inttipodecomunicacao and ROWNUM <= 1) strtelefone "
        End If
        
        strSQL = strSQL & _
        "from " & _
        gstrEconomico & " ec, " & _
        gstrContribuinte & " co, " & _
        gstrLogradouro & " lo, " & _
        gstrTituloLogradouro & " tl, " & _
        gstrBairro & " ba, " & _
        gstrCidade & " mu, " & _
        gstrUF & " uf, " & _
        gstrFormaDeComunicacao & " fc "
        strSQL = strSQL & _
        "Where " & _
        "ec.intcontribuinte = co.pkid and " & _
        "ec.intlogradouro = lo.pkid and " & _
        "lo.inttitulologradouro " & strOUTJSQLServer & "= tl.pkid " & strOUTJOracle & " and " & _
        "ec.intbairro = ba.pkid and " & _
        "ba.intmunicipio = mu.pkid and " & _
        "mu.intuf = uf.pkid and " & _
        "co.pkid " & strOUTJSQLServer & "= fc.intcontribuinte " & strOUTJOracle & " and " & _
        "ec.pkid = " & dbc_strInscricao.BoundText
        
        Set gobjBanco = New clsBanco
        If gobjBanco.CriaADO(strSQL, 10, adoResultado) Then
            If Not adoResultado.EOF Then
                txtStrRazao = gstrENulo(adoResultado!STRNOME)
                txtCancelamento = gstrENulo(adoResultado!DTMDATAENCERRAMENTO)
                txtEndereco = gstrENulo(adoResultado!strTitulo) & " " & gstrENulo(adoResultado!strLogradouro) & ", " & gstrENulo(adoResultado!INTNUMERO) & " - " & gstrENulo(adoResultado!STRBAIRRO) & " - " & gstrENulo(adoResultado!STRMUNICIPIO) & " - " & gstrENulo(adoResultado!STRUF) & " - " & gstrCEPFormatado(gstrENulo(adoResultado!INTCEP))
                txtCNPJ = gstrENulo(adoResultado!STRCNPJCPF)
                txtTelefone = gstrENulo(adoResultado!strTelefone)
                txtFantasia = gstrENulo(adoResultado!strNomeFantasia)
            Else
                txtStrRazao = ""
                txtCancelamento = ""
                txtEndereco = ""
                txtCNPJ = ""
                txtTelefone = ""
                txtFantasia = ""
            End If
        End If
        
    End If
    
End Sub

Private Sub PreencheGrdAtividades(lngPkid As Long)
    Dim strSQL          As String
    Dim adoResultado    As ADODB.Recordset
    
    strSQL = "SELECT "
    strSQL = strSQL & "Case When AE.BLNPRINCIPAL = 1 Then 'Principal' Else 'Secundária' end Status, "
    strSQL = strSQL & gstrCONVERT(CDT_NVARCHAR, "AEC.intCodigo") & strCONCAT & " ' - ' " & strCONCAT & " Ltrim(Rtrim(AEC.strDescricao)) " & strCONCAT & " ' / ' " & strCONCAT & " Ltrim(Rtrim(SA.STRNOMEDOSUBGRUPO)) AS strDescricao "
    strSQL = strSQL & "FROM "
    strSQL = strSQL & gstrAtividadeDaEmpresa & " AE, "
    strSQL = strSQL & gstrSubGrupoDeAtividade & " SA, "
    strSQL = strSQL & gstrAtividadeEC & " AEC "
    strSQL = strSQL & "Where "
    strSQL = strSQL & "AE.Intatividade = AEC.Pkid AND "
    strSQL = strSQL & "Sa.Pkid = AEC.intSubGrupo AND "
    strSQL = strSQL & "AE.Inteconomico = " & lngPkid
    strSQL = strSQL & " ORDER BY "
    strSQL = strSQL & "AEC.intCodigo "
    
    LeDaTabelaParaObj "", tdb_Atividade, strSQL
    
End Sub

Private Sub txtDTBase_GotFocus()
    MarcaCampo txtDTBase
    txtDTBase = gstrDataDoSistema
End Sub

Private Sub txtDTBase_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "D", txtDTBase
End Sub

Private Sub txtDTBase_LostFocus()
    txtDTBase = gstrDataFormatada(txtDTBase)
End Sub

Private Sub txtNotaInicial_Change()
    If Val(txtQuantidade) = 1 Then
        txtNotaFinal = Val(txtNotaInicial)
    ElseIf Val(txtQuantidade) > 1 Then
        txtNotaFinal = Val(txtNotaInicial) + Val(txtQuantidade) - 1
    End If
End Sub

Private Sub txtNotaInicial_GotFocus()
    MarcaCampo txtNotaInicial
    txtNotaInicial = ProximoNumeroNotaFiscal
End Sub

Private Sub txtNotaInicial_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtNotaInicial
End Sub

Private Sub txtQuantidade_Change()
    If Val(txtQuantidade) = 1 Then
        txtControleFinal = Val(txtControleInicial)
        txtNotaFinal = Val(txtNotaInicial)
    ElseIf Val(txtQuantidade) > 1 Then
        txtControleFinal = Val(txtControleInicial) + Val(txtQuantidade) - 1
        txtNotaFinal = Val(txtNotaInicial) + Val(txtQuantidade) - 1
    End If
End Sub

Private Sub txtQuantidade_GotFocus()
    MarcaCampo txtQuantidade
End Sub

Private Sub txtQuantidade_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtQuantidade
End Sub

Private Function strNotaFiscal() As String
    Dim strSQL As String
    Dim adoResultado As ADODB.Recordset
    
    strSQL = ""
    strSQL = strSQL & "SELECT "
    strSQL = strSQL & " intnotafiscalcontrole "
    strSQL = strSQL & "FROM " & gstrParametrosTributario
    
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSQL, 10, adoResultado) Then
        If Not adoResultado.EOF Then
            strNotaFiscal = IIf(Val(gstrENulo(adoResultado!intnotafiscalcontrole)) <= 0, 1, Val(gstrENulo(adoResultado!intnotafiscalcontrole)))
        Else
            strNotaFiscal = "1"
        End If
    End If
End Function

Private Function blnDadosOK() As Boolean
    Dim adoResultado    As ADODB.Recordset
    Dim strSQL          As String
    
    blnDadosOK = False
    
    If Not dbc_strInscricao.MatchedWithList Then
        ExibeMensagem "É necessário preenchimento correto da inscrição."
        dbc_strInscricao.SetFocus
        Exit Function
    ElseIf Trim(txtDTBase.Text) = "" Then
        ExibeMensagem "É necessário preenchimento correto da Data Base."
        txtDTBase.SetFocus
        Exit Function
    ElseIf Val(txtQuantidade.Text) <= 0 Then
        ExibeMensagem "É necessário preenchimento correto da quantidade."
        txtQuantidade.SetFocus
        Exit Function
    ElseIf Val(txtNotaInicial.Text) <= 0 Then
        ExibeMensagem "É necessário preenchimento correto da nota Fiscal Inicial."
        txtNotaInicial.SetFocus
        Exit Function
    ElseIf Val(txtSerie) <= 0 Then
        ExibeMensagem "É necessário preenchimento correto da série."
        txtSerie.SetFocus
        Exit Function
    ElseIf Len(Trim(txtCancelamento)) > 0 Then
        ExibeMensagem "O contribuinte está com data de encerramento."
        dbc_strInscricao.SetFocus
        Exit Function
    End If
    
    strSQL = "select * "
    strSQL = strSQL & "From "
    strSQL = strSQL & "tblnotafiscaliss NF "
    strSQL = strSQL & "Where "
    strSQL = strSQL & "NF.Strnotafiscalnr Between " & txtNotaInicial & " And " & txtNotaInicial & "  AND "
    strSQL = strSQL & "NF.Strnotafiscalserie = " & txtSerie & " AND "
    strSQL = strSQL & "NF.Inteconomico = " & Val(dbc_strInscricao.BoundText)
    
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSQL, 10, adoResultado) Then
        If Not adoResultado.EOF Then
            ExibeMensagem "Já se encontra cadastrado no sistema numeração de notas fiscais solicitadas."
            Exit Function
        End If
    End If
    
    blnDadosOK = True
    
End Function

Private Function PreencheVetorNota() As Boolean
    
    Dim strSQL                  As String
    Dim adoResultado            As ADODB.Recordset
    Dim intNotaFiscalLimitedias As Integer
    Dim intFor                  As Integer
    
    PreencheVetorNota = False
    
    On Error GoTo ErroNaRotina
    intFor = 1
    
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO("select intnotafiscallimitedias from " & gstrParametrosTributario, 10, adoResultado) Then
        If Not adoResultado.EOF Then
            intNotaFiscalLimitedias = Val(gstrENulo(adoResultado!intNotaFiscalLimitedias))
        Else
            intNotaFiscalLimitedias = 0
        End If
    End If
    
    strSQL = ""
    strSQL = strSQL & "SELECT "
    strSQL = strSQL & "EC.PKID Pkid, "
    strSQL = strSQL & gstrRIGHT("EC.strInscricaoCadastral", gintRetornaTamanhoMascara(TYP_ECONOMICA)) & " strInscricaoCadastral, "
    strSQL = strSQL & "CO.Strnome RazaoSocial, "
    strSQL = strSQL & "Co.Strinscricaoestadual, "
    strSQL = strSQL & "CO.bytNaturezaJuridica, "
    strSQL = strSQL & "CO.Strcnpjcpf, CO.strnomefantasia, "
    strSQL = strSQL & gstrISNULL("TL.Strsigla", "''") & strCONCAT & "' '" & strCONCAT & gstrISNULL("LG.Strdescricao", "''") & strCONCAT & "' '" & _
    strCONCAT & gstrCONVERT(CDT_NVARCHAR, "EC.Intnumero") & strCONCAT & "' - '" & strCONCAT & gstrISNULL("BA.strDescricao", "''") & " strLogradouro, "
    strSQL = strSQL & gstrISNULL("EC.intCep", "''") & " intCep, "
    
    If bytDBType = SQLServer Then
        strSQL = strSQL & "(select top 1 fc.strconteudo from " & gstrTipoDeComunicacao & " tc where tc.pkid = fc.inttipodecomunicacao and tc.inttipo = 6) strtelefone "
    Else
        strSQL = strSQL & "(SELECT fc.strconteudo FROM (select tc.pkid from tblTipoDeComunicacao tc where tc.inttipo = 6) tc Where tc.Pkid = fc.inttipodecomunicacao and ROWNUM <= 1) strtelefone "
    End If
    
    strSQL = strSQL & " FROM "
    strSQL = strSQL & gstrEconomico & " EC, "
    strSQL = strSQL & gstrContribuinte & " CO, "
    strSQL = strSQL & gstrTipoLogradouro & " TL, "
    strSQL = strSQL & gstrLogradouro & " LG, "
    strSQL = strSQL & gstrBairro & " BA, "
    strSQL = strSQL & gstrFormaDeComunicacao & " fc "
    strSQL = strSQL & " WHERE "
    strSQL = strSQL & "EC.pkid = " & dbc_strInscricao.BoundText & " And "
    strSQL = strSQL & "EC.Intcontribuinte" & strOUTJSQLServer & "= CO.pkid" & strOUTJOracle & " and "
    strSQL = strSQL & "EC.Intlogradouro = LG.Pkid" & strOUTJOracle & " and "
    strSQL = strSQL & "LG.INTTIPOLOGRADOURO" & strOUTJSQLServer & "= TL.Pkid" & strOUTJOracle & " and "
    strSQL = strSQL & "EC.Intbairro" & strOUTJSQLServer & "= BA.Pkid" & strOUTJOracle & " and "
    strSQL = strSQL & "EC.Intcontribuinte " & strOUTJSQLServer & "= fc.intcontribuinte " & strOUTJOracle
    
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSQL, 10, adoResultado) Then
        If Not adoResultado.EOF Then
            Set gobjBanco = New clsBanco
            gobjBanco.ExecutaBeginTrans
            Do While (intFor) <= txtQuantidade
                
                ReDim Preserve vetNotas(13, intFor - 1)
                
                vetNotas(0, intFor - 1) = gstrENulo(adoResultado("RazaoSocial").Value)
                vetNotas(1, intFor - 1) = gstrENulo(adoResultado("Strinscricaoestadual").Value)
                vetNotas(2, intFor - 1) = gstrENulo(adoResultado("bytNaturezaJuridica").Value)
                vetNotas(3, intFor - 1) = gstrENulo(adoResultado("Strcnpjcpf").Value)
                vetNotas(4, intFor - 1) = gstrENulo(adoResultado("strLogradouro").Value)
                vetNotas(5, intFor - 1) = gstrENulo(adoResultado("intCep").Value)
                vetNotas(6, intFor - 1) = gstrDataFormatada(txtDTBase)
                vetNotas(7, intFor - 1) = Val(txtControleInicial) + (intFor - 1) 'Controle
                vetNotas(8, intFor - 1) = Val(txtNotaInicial) + (intFor - 1)    'Nota Fiscal
                vetNotas(9, intFor - 1) = Val(txtSerie)                     'Serie
                vetNotas(10, intFor - 1) = gstrDataFormatada(DateAdd("d", intNotaFiscalLimitedias, txtDTBase)) 'Data Limite
                vetNotas(11, intFor - 1) = gstrENulo(adoResultado("strInscricaoCadastral").Value)
                vetNotas(12, intFor - 1) = gstrENulo(adoResultado("strnomefantasia"))
                vetNotas(13, intFor - 1) = gstrENulo(adoResultado("strtelefone"))
                
                strSQL = "insert into tblnotafiscaliss (inteconomico, dtmdtbase, intcontrolenr, strnotafiscalnr, strnotafiscalserie, dtmdtnotafiscalbaixa, dblnotafiscalvalor, intnotafiscalocorrencia, dtmdtatualizacao, lngcodusr, dtmdtlimite) "
                strSQL = strSQL & "values( "
                strSQL = strSQL & dbc_strInscricao.BoundText & ", "              'v_inteconomico
                strSQL = strSQL & gstrConvDtParaSql(txtDTBase) & ", "            'v_dtmdtbase
                strSQL = strSQL & Val(txtControleInicial) + (intFor - 1) & ", "   'v_intcontrolenr
                strSQL = strSQL & Val(txtNotaInicial) + (intFor - 1) & ", "       'v_strnotafiscalnr
                strSQL = strSQL & Val(txtSerie) & ", "                           'v_strnotafiscalserie
                strSQL = strSQL & "Null" & ", "                                  'v_dtmdtnotafiscalbaixa
                strSQL = strSQL & "Null" & ", "                                  'v_dblnotafiscalvalor
                strSQL = strSQL & "Null" & ", "                                  'v_intnotafiscalocorrencia
                strSQL = strSQL & gstrConvDtParaSql(gstrDataDoSistema) & ", "    'v_dtmdtatualizacao
                strSQL = strSQL & glngCodUsr & ", "                              'v_lngcodusr
                strSQL = strSQL & gstrConvDtParaSql(vetNotas(10, intFor - 1)) & ") " 'v_dtmdtlimite
                
                If Not gobjBanco.Execute(strSQL) Then
                    GoTo ErroNaRotina
                End If
                intFor = intFor + 1
            Loop
        End If
    End If
    
    strSQL = "Update " & gstrParametrosTributario & " set intnotafiscalcontrole = " & Val(txtControleInicial) + (intFor - 1)
    
    If Not gobjBanco.Execute(strSQL) Then
        GoTo ErroNaRotina
    End If
    
    gobjBanco.ExecutaCommitTrans
    
    PreencheVetorNota = True
    Exit Function
    
ErroNaRotina:
    gobjBanco.ExecutaRollbackTrans
    
End Function

Private Sub txtSerie_GotFocus()
    MarcaCampo txtSerie
End Sub

Private Sub txtSerie_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtSerie
End Sub

Private Function blnVerificaEncerrado() As Boolean
    
    'TRI0748
    If Len(dbc_strInscricao.BoundText) <> 0 And Trim(txtCancelamento.Text) <> "" Then
        ExibeMensagem "Esta inscrição foi encerrada em : " & txtCancelamento.Text
        HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrImprimir
        'dbc_strInscricao.SetFocus
        blnVerificaEncerrado = True
        Exit Function
    Else
        HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrImprimir
        blnVerificaEncerrado = False
    End If
    
End Function
