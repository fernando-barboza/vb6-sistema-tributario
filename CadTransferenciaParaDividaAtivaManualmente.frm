VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "TODG7.OCX"
Begin VB.Form frmCadTransferenciaParaDividaAtivaManualmente 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Débitos em Transferência para Dívida Ativa - Incluídos Manualmente"
   ClientHeight    =   3345
   ClientLeft      =   135
   ClientTop       =   1425
   ClientWidth     =   8490
   Icon            =   "CadTransferenciaParaDividaAtivaManualmente.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3345
   ScaleWidth      =   8490
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab tab_3dPasta 
      Height          =   3225
      Left            =   90
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   60
      Width           =   8355
      _ExtentX        =   14737
      _ExtentY        =   5689
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Transferência para Dívida Ativa"
      TabPicture(0)   =   "CadTransferenciaParaDividaAtivaManualmente.frx":1042
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lbl_ContribuinteFinal"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lbl_ContribuinteInicial"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lbl_ExercicioInicial"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lbl_strComposicaoReceita"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lbl_ExercicioFinal"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "dbc_ContribuinteInicial"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "dbc_strComposicaoReceita"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "dbc_ContribuinteFinal"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txt_ExercicioInicial"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txt_ExercicioFinal"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txt_PKId"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "chk_Selecionar"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).ControlCount=   12
      TabCaption(1)   =   "Composição da Receita"
      TabPicture(1)   =   "CadTransferenciaParaDividaAtivaManualmente.frx":105E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "tdb_Atividades"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.CheckBox chk_Selecionar 
         Caption         =   "Selecionar todos os Contribuintes"
         Height          =   255
         Left            =   2220
         TabIndex        =   13
         Top             =   1440
         Width           =   2835
      End
      Begin VB.TextBox txt_PKId 
         Height          =   285
         Left            =   7680
         TabIndex        =   12
         Top             =   0
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.TextBox txt_ExercicioFinal 
         Height          =   285
         Left            =   7470
         MaxLength       =   4
         TabIndex        =   3
         Top             =   1830
         Width           =   525
      End
      Begin VB.TextBox txt_ExercicioInicial 
         Height          =   285
         Left            =   2220
         MaxLength       =   4
         TabIndex        =   2
         Top             =   1845
         Width           =   525
      End
      Begin MSDataListLib.DataCombo dbc_ContribuinteFinal 
         Height          =   315
         Left            =   2220
         TabIndex        =   1
         Top             =   1020
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   556
         _Version        =   393216
         IntegralHeight  =   0   'False
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dbc_strComposicaoReceita 
         Height          =   315
         Left            =   2220
         TabIndex        =   4
         Top             =   2250
         Width           =   5805
         _ExtentX        =   10239
         _ExtentY        =   556
         _Version        =   393216
         IntegralHeight  =   0   'False
         Style           =   2
         Text            =   ""
      End
      Begin TrueOleDBGrid70.TDBGrid tdb_Atividades 
         Height          =   2475
         Left            =   -74850
         TabIndex        =   6
         Top             =   570
         Width           =   8025
         _ExtentX        =   14155
         _ExtentY        =   4366
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "PKId"
         Columns(0).DataField=   "PKId"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   68
         Columns(1)._MaxComboItems=   20
         Columns(1).ValueItems(0)._DefaultItem=   0
         Columns(1).ValueItems(0).Value=   ""
         Columns(1).ValueItems(0).Value.vt=   8
         Columns(1).ValueItems(0).DisplayValue=   ""
         Columns(1).ValueItems(0).DisplayValue.vt=   8
         Columns(1).ValueItems(0)._PropDict=   "_DefaultItem,517,2"
         Columns(1).ValueItems.Count=   1
         Columns(1).DataField=   ""
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Descrição"
         Columns(2).DataField=   "strDescricao"
         Columns(2).DropDown=   "tdd_Atividades"
         Columns(2).DropDown.vt=   8
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   3
         Splits(0)._UserFlags=   1
         Splits(0).MarqueeStyle=   5
         Splits(0).RecordSelectors=   0   'False
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).ScrollBars=   2
         Splits(0).DividerColor=   12632256
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=3"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0).AllowSizing=0"
         Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
         Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(8)=   "Column(1).Width=450"
         Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=370"
         Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(12)=   "Column(1).AllowSizing=0"
         Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(14)=   "Column(2).Width=13123"
         Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=13044"
         Splits(0)._ColumnProps(17)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(19)=   "Column(2).AutoDropDown=1"
         Splits(0)._ColumnProps(20)=   "Column(2).DropDownList=1"
         Splits(0)._ColumnProps(21)=   "Column(2).AutoCompletion=1"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   0
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         DataMode        =   4
         DefColWidth     =   0
         HeadLines       =   1
         FootLines       =   1
         MultipleLines   =   0
         CellTipsWidth   =   0
         DeadAreaBackColor=   12632256
         RowDividerColor =   12632256
         RowSubDividerColor=   12632256
         DirectionAfterEnter=   0
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
         _StyleDefs(30)  =   "Splits(0).Columns(0).Style:id=46,.parent=13"
         _StyleDefs(31)  =   "Splits(0).Columns(0).HeadingStyle:id=43,.parent=14"
         _StyleDefs(32)  =   "Splits(0).Columns(0).FooterStyle:id=44,.parent=15"
         _StyleDefs(33)  =   "Splits(0).Columns(0).EditorStyle:id=45,.parent=17"
         _StyleDefs(34)  =   "Splits(0).Columns(1).Style:id=28,.parent=13"
         _StyleDefs(35)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14"
         _StyleDefs(36)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
         _StyleDefs(37)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
         _StyleDefs(38)  =   "Splits(0).Columns(2).Style:id=50,.parent=13"
         _StyleDefs(39)  =   "Splits(0).Columns(2).HeadingStyle:id=47,.parent=14"
         _StyleDefs(40)  =   "Splits(0).Columns(2).FooterStyle:id=48,.parent=15"
         _StyleDefs(41)  =   "Splits(0).Columns(2).EditorStyle:id=49,.parent=17"
         _StyleDefs(42)  =   "Named:id=33:Normal"
         _StyleDefs(43)  =   ":id=33,.parent=0"
         _StyleDefs(44)  =   "Named:id=34:Heading"
         _StyleDefs(45)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(46)  =   ":id=34,.wraptext=-1"
         _StyleDefs(47)  =   "Named:id=35:Footing"
         _StyleDefs(48)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(49)  =   "Named:id=36:Selected"
         _StyleDefs(50)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(51)  =   "Named:id=37:Caption"
         _StyleDefs(52)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(53)  =   "Named:id=38:HighlightRow"
         _StyleDefs(54)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(55)  =   "Named:id=39:EvenRow"
         _StyleDefs(56)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(57)  =   "Named:id=40:OddRow"
         _StyleDefs(58)  =   ":id=40,.parent=33"
         _StyleDefs(59)  =   "Named:id=41:RecordSelector"
         _StyleDefs(60)  =   ":id=41,.parent=34"
         _StyleDefs(61)  =   "Named:id=42:FilterBar"
         _StyleDefs(62)  =   ":id=42,.parent=33"
      End
      Begin MSDataListLib.DataCombo dbc_ContribuinteInicial 
         Height          =   315
         Left            =   2220
         TabIndex        =   0
         Top             =   570
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   556
         _Version        =   393216
         IntegralHeight  =   0   'False
         Text            =   ""
      End
      Begin VB.Label lbl_ExercicioFinal 
         AutoSize        =   -1  'True
         Caption         =   "Exercício Final"
         Height          =   195
         Left            =   6270
         TabIndex        =   11
         Top             =   1875
         Width           =   1050
      End
      Begin VB.Label lbl_strComposicaoReceita 
         AutoSize        =   -1  'True
         Caption         =   "Composição da Receita"
         Height          =   195
         Left            =   360
         TabIndex        =   10
         Top             =   2295
         Width           =   1695
      End
      Begin VB.Label lbl_ExercicioInicial 
         AutoSize        =   -1  'True
         Caption         =   "Exercício Inicial"
         Height          =   195
         Left            =   930
         TabIndex        =   9
         Top             =   1890
         Width           =   1125
      End
      Begin VB.Label lbl_ContribuinteInicial 
         AutoSize        =   -1  'True
         Caption         =   "Contribuinte Inicial"
         Height          =   195
         Left            =   765
         TabIndex        =   8
         Top             =   600
         Width           =   1290
      End
      Begin VB.Label lbl_ContribuinteFinal 
         AutoSize        =   -1  'True
         Caption         =   "Contribuinte Final"
         Height          =   195
         Left            =   840
         TabIndex        =   7
         Top             =   1080
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmCadTransferenciaParaDividaAtivaManualmente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mblnAlterando                   As Boolean
Dim mobjAux                         As Object
Dim mblnSelecionou                  As Boolean
Dim mblnPrimeiraVez                 As Boolean
Dim xarReceita                      As XArrayDB
Dim adoRecContribuinte              As ADODB.Recordset
    
Private Sub chk_Selecionar_Click()
    If chk_Selecionar.Value = 1 Then
        dbc_ContribuinteInicial.BoundText = ""
        dbc_ContribuinteFinal.BoundText = ""
        dbc_ContribuinteInicial.Enabled = False
        TrocaCorObjeto dbc_ContribuinteInicial, True
        dbc_ContribuinteFinal.Enabled = False
        TrocaCorObjeto dbc_ContribuinteFinal, True
        txt_ExercicioInicial.SetFocus
    Else
        dbc_ContribuinteInicial.Enabled = True
        TrocaCorObjeto dbc_ContribuinteInicial, False
        dbc_ContribuinteFinal.Enabled = True
        TrocaCorObjeto dbc_ContribuinteFinal, False
        dbc_ContribuinteInicial.SetFocus
    End If

End Sub

Private Sub dbc_ContribuinteFinal_Click(Area As Integer)
    DropDownDataCombo dbc_ContribuinteFinal, Me, Area
End Sub

Private Sub dbc_ContribuinteFinal_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbc_ContribuinteFinal, Me, , KeyCode, Shift
End Sub

Private Sub dbc_ContribuinteInicial_Click(Area As Integer)
    DropDownDataCombo dbc_ContribuinteInicial, Me, Area
End Sub

Private Sub dbc_ContribuinteInicial_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbc_ContribuinteInicial, Me, , KeyCode, Shift
End Sub

Private Sub dbc_strComposicaoReceita_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbc_strComposicaoReceita, Me, , KeyCode, Shift
End Sub

Private Sub Form_Activate()
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrNovo
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrSalvar
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrDeletar
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar
End Sub

Private Sub Form_Load()
    mblnAlterando = False
    
    dbc_ContribuinteInicial.Tag = strQueryContribuinte & ";strNOme"
    dbc_ContribuinteFinal.Tag = dbc_ContribuinteInicial.Tag

    LeDaTabelaParaObj gstrComposicaoDaReceita, dbc_strComposicaoReceita, QueryComposicao
    tab_3dPasta.TabEnabled(1) = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrCalcularReajuste
    mblnSelecionou = False
    mblnPrimeiraVez = False
End Sub

Public Sub MantemForm(ByVal strModoOperacao As String)
Dim strSQL As String

    If UCase(strModoOperacao) = UCase(gstrFechar) Then
        Unload Me
    End If
    
    If strModoOperacao = gstrCalcularReajuste Then
        EfetuaCalculodeReceitasDiversas
    End If
    
    If strModoOperacao = gstrPreencherLista Then
        PreencherListaDeOpcoes Me.ActiveControl
    End If
End Sub

Private Sub tab_3dPasta_Click(PreviousTab As Integer)
    Select Case tab_3dPasta.Tab
        Case 0
            HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrCalcularReajuste
        Case 1
            'HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrCalcularReajuste
            'If blnDadosOK Then
                HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrCalcularReajuste
           ' End If
    End Select
    
End Sub

Private Sub dbc_strComposicaoReceita_Click(Area As Integer)
    DropDownDataCombo dbc_strComposicaoReceita, Me, Area
    If Area = 2 And dbc_strComposicaoReceita.MatchedWithList Then
        MontaAtividade dbc_strComposicaoReceita.BoundText
        tab_3dPasta.TabEnabled(1) = True
    End If
End Sub

Private Function QueryComposicao() As String
    Dim strSQL As String
    strSQL = ""
    strSQL = strSQL & "SELECT PKId, strDescricao FROM " & gstrComposicaoDaReceita
'    strSql = strSql & " WHERE intUtilizacao = 1 "
    strSQL = strSQL & " ORDER BY strDescricao "
    QueryComposicao = strSQL
End Function

Private Function strQueryContribuinte() As String

'******************************************************************************************
' Data: 02/05/2003
' Alteração: - Substituição do comando nativo CONVERT do SQL Server pela função gstrCONVERT
' Responsável: Everton Bianchini
'------------------------------------------------------------------------------------------
' Data: 05/05/2003
' Alteração: - Substituição do comando nativo de concatenação de strings (+) do SQL Server
'            pela variável strCONCAT.
' Responsável: Everton Bianchini
'******************************************************************************************
    
Dim strSQL As String

    strSQL = ""
'    strSql = strSql & " SELECT CO.PKID, CONVERT(NVARCHAR,DA.intContribuinte) + ' - ' + CO.strNome AS Nome "
    strSQL = strSQL & " SELECT CO.PKID, " & gstrCONVERT(CDT_NVARCHAR, "DA.intContribuinte") & strCONCAT & " ' - ' " & strCONCAT & " CO.strNome AS Nome "
    strSQL = strSQL & " FROM " & gstrDividaAtiva & " DA, "
    strSQL = strSQL & gstrContribuinte & " CO "
    strSQL = strSQL & " WHERE CO.PKID = DA.intContribuinte "
    strSQL = strSQL & " ORDER BY intContribuinte "
    
strQueryContribuinte = strSQL
End Function

Private Sub MontaAtividade(intComposicaoReceita As Integer)
Dim strSQL As String
Dim adoRec As ADODB.Recordset
Dim varAux As String

On Error GoTo Err_Handle

Set xarReceita = New XArrayDB
xarReceita.Clear

xarReceita.ReDim 0, 0, 0, 2

strSQL = ""
strSQL = strSQL & " SELECT A.PKId, A.strDescricao FROM "
strSQL = strSQL & gstrReceita & " A,"
strSQL = strSQL & gstrValorCompoRec & " B"
strSQL = strSQL & " WHERE A.PKId = B.intReceita "
strSQL = strSQL & " AND B.intComposicaoDaReceita = " & intComposicaoReceita

Set gobjBanco = New clsBanco

If gobjBanco.CriaADO(strSQL, 10, adoRec) Then
    With adoRec
        If Not .EOF Then
            xarReceita.ReDim 0, .RecordCount - 1, 0, 2
            Do While Not .EOF
                varAux = !Pkid
                xarReceita(.AbsolutePosition - 1, 0) = varAux
                
                varAux = False
                xarReceita(.AbsolutePosition - 1, 1) = varAux
            
                varAux = !strDescricao
                xarReceita(.AbsolutePosition - 1, 2) = varAux
                
                .MoveNext
            Loop
        End If
    End With
End If

Set tdb_Atividades.Array = xarReceita
tdb_Atividades.Rebind
tdb_Atividades.Refresh

Exit Sub
Err_Handle:

End Sub

Private Sub tdb_Atividades_AfterColUpdate(ByVal ColIndex As Integer)
    tdb_Atividades.Update
End Sub

Private Function strPKId() As String
    Dim strSQL As String
    Dim i As Integer
    strSQL = ""
    For i = 0 To xarReceita.Count(1) - 1
        If xarReceita(i, 2) = -1 Then
            If strSQL <> "" Then
               strSQL = strSQL & ","
            End If
        strSQL = strSQL & xarReceita(i, 0)
        End If
    Next

    strPKId = strSQL
End Function

Private Sub EfetuaCalculodeReceitasDiversas()

'******************************************************************************************
' Data: 02/05/2003
' Alteração: - Substituição do comando nativo GETDATE() do SQL Server pela variável
'            strGETDATE.
' Responsável: Everton Bianchini
'------------------------------------------------------------------------------------------
' Data: 09/05/2003
' Alteração: - Incluídos comandos de inicialização e finalização de bloco PL/SQL permitindo
'            , assim, a execução de múltiplos comandos SQL de uma única vez.
' Responsável: Everton Bianchini
'******************************************************************************************
    
Dim strSQL                  As String
Dim lngSequencia            As Long

Set gobjBanco = New clsBanco

If blnDadosOk Then

    Contribuintes
    If adoRecContribuinte.RecordCount = 0 Then
        ExibeMensagem "Todos os lançamentos neste período já foram efetuados!"
        tab_3dPasta.Tab = 0
        Exit Sub
    End If
    
    adoRecContribuinte.MoveFirst
    strSQL = ""

    gobjBanco.ExecutaBeginTrans
    Screen.MousePointer = vbHourglass

    Do While Not adoRecContribuinte.EOF

        'Pesquisa a sequência da composição da receita
        lngSequencia = lngRetornaSequencia(adoRecContribuinte!intContribuinte, adoRecContribuinte!intExercicio)

        If (bytDBType = EDatabases.Oracle) Then
            strSQL = strSQL & "DECLARE "
            strSQL = strSQL & "TYPE tp_csr IS REF CURSOR; "
            strSQL = strSQL & "csr tp_csr; "
            strSQL = strSQL & "V_DESCONTO NUMBER := 0; "
            strSQL = strSQL & "BEGIN "
        End If

        strSQL = strSQL & " INSERT INTO " & gstrLancamentoCalculo
        strSQL = strSQL & " (intExercicio, intContribuinte, intComposicaoReceita, intMensagem, strInscricaoCadastral, "
        strSQL = strSQL & " dtmLancamento, dtmVencimento, intNumeroDeParcelas, intIntervaloEntreParcelas, "
        strSQL = strSQL & " bitUtilizacaoDebito, intOcorrencia, bytOrigem, "
        strSQL = strSQL & " strSequencia, dtmDtAtualizacao, lngCodUsr ) VALUES ( "
        strSQL = strSQL & adoRecContribuinte!intExercicio
        strSQL = strSQL & ", " & adoRecContribuinte!intContribuinte
        strSQL = strSQL & ", " & adoRecContribuinte!intComposicaoReceita
        strSQL = strSQL & ", NULL" 'Mensagem - pode conter null
        strSQL = strSQL & ", '" & adoRecContribuinte!strInscricaoCadastral 'Inscrição cadastral (Para receitas diversas - código do contribuinte)
        strSQL = strSQL & "', " & gstrConvDtParaSql(adoRecContribuinte!dtmInscricao)
        strSQL = strSQL & ", " & gstrConvDtParaSql(adoRecContribuinte!dtmVencimento)
        strSQL = strSQL & ", 1" 'número de parcelas
        strSQL = strSQL & ", 0" 'intervalo entre parcelas
        strSQL = strSQL & ", 2" 'Utilização do débito
        strSQL = strSQL & ", NULL" 'Ocorrência
        strSQL = strSQL & ", 4" 'Origem
        strSQL = strSQL & ", " & CStr(lngSequencia)
'        strSql = strSql & ", GETDATE()"
        strSQL = strSQL & ", " & strGETDATE
        strSQL = strSQL & ", " & glngCodUsr
        strSQL = strSQL & " )"

        strSQL = strSQL & IIf((bytDBType = EDatabases.Oracle), "; ", "")

        strSQL = strSQL & " INSERT INTO " & gstrParcelaReceita
        strSQL = strSQL & " (intLancamentoCalculo, intComposicaoDaReceita, intNumeroParcela, dtmDataVencimento, "
        strSQL = strSQL & " bytDividaAjuizada, bytSimulado, bytPrescrita, bytCancelada, bytAtiva, bytSuspensaoDeExigencia, "
        strSQL = strSQL & " dblValorParcela, dtmDtAtualizacao, lngCodUsr) "
        strSQL = strSQL & " (SELECT MAX(PKId) "
        strSQL = strSQL & ", " & adoRecContribuinte!intComposicaoReceita
        strSQL = strSQL & ", 1"
        strSQL = strSQL & ", " & gstrConvDtParaSql(adoRecContribuinte!dtmVencimento)
        strSQL = strSQL & ", 0" 'Dívida Ajuizada
        strSQL = strSQL & ", 0" 'Simulado
        strSQL = strSQL & ", 0" 'Prescrita
        strSQL = strSQL & ", 0" 'Cancelada
        strSQL = strSQL & ", 1" 'Divida Ativa
        strSQL = strSQL & ", 0" 'Suspensão de Exigência
        strSQL = strSQL & ", " & gstrConvVrParaSql(adoRecContribuinte!dblValorOriginal)
'        strSql = strSql & ", GETDATE()"
        strSQL = strSQL & ", " & strGETDATE
        strSQL = strSQL & ", " & glngCodUsr
        strSQL = strSQL & " FROM " & gstrLancamentoCalculo & ")"
        
        strSQL = strSQL & IIf((bytDBType = EDatabases.Oracle), "; ", " ")
        
        'Desmarca o flag de débito incluido manualmente do cadastro de dívida ativa
        strSQL = strSQL & " UPDATE " & gstrDetalheDividaAtiva & " SET "
        strSQL = strSQL & " bytDebitoGeradoManualmente = 0 "
        strSQL = strSQL & " WHERE PKId = " & adoRecContribuinte!Pkid
        
        strSQL = strSQL & IIf((bytDBType = EDatabases.Oracle), "; ", "")
        
        'Grava as Parcelas Taxas
'        strSQL = strSQL & " EXECUTE sp_EfetuaCalculo '" & strPKId & "', " & dbc_strComposicaoReceita.BoundText & ",23,1,"
        strSQL = strSQL & IIf((bytDBType = EDatabases.Oracle), "pe_EfetuaCalculo.", " EXECUTE ")
        strSQL = strSQL & "sp_EfetuaCalculo" & IIf((bytDBType = EDatabases.Oracle), "(", " ") & _
            "'" & strPKId & "', " & dbc_strComposicaoReceita.BoundText & ",23,1,"
        
'        strSql = strSql & gstrConvDtParaSql(adoRecContribuinte!dtmVencimento) & ",0,0,0, " & glngCodUsr
        strSQL = strSQL & gstrConvDtParaSql(adoRecContribuinte!dtmVencimento) & ",0,0," & _
            IIf((bytDBType = EDatabases.Oracle), " V_DESCONTO", " 0") & ", " & glngCodUsr
        
        strSQL = strSQL & IIf((bytDBType = EDatabases.Oracle), ", csr); ", "")
        
        adoRecContribuinte.MoveNext

    Loop

    strSQL = strSQL & IIf((bytDBType = EDatabases.Oracle), "END; ", "")

        Set gobjBanco = New clsBanco
        If gobjBanco.Execute(strSQL, False) Then
            gobjBanco.ExecutaCommitTrans
            Screen.MousePointer = vbNormal
            ExibeMensagem "Tranferência efetuada com sucesso!"
            tab_3dPasta.Tab = 0
        Else
            Screen.MousePointer = vbNormal
            gobjBanco.ExecutaRollbackTrans
        End If

End If
End Sub

Private Sub Contribuintes()
Dim strSQL          As String

strSQL = ""
strSQL = strSQL & " SELECT DA.intcontribuinte, DDA.* "
strSQL = strSQL & " FROM " & gstrDividaAtiva & " DA, "
strSQL = strSQL & gstrDetalheDividaAtiva & " DDA "
strSQL = strSQL & " WHERE DA.PKId = DDA.intDividaAtiva "
If chk_Selecionar.Value <> 1 Then
    strSQL = strSQL & " AND DA.intContribuinte BETWEEN " & dbc_ContribuinteInicial.BoundText & " AND " & dbc_ContribuinteFinal.BoundText
End If
strSQL = strSQL & " AND DDA.intExercicio BETWEEN " & Val(txt_ExercicioInicial.Text) & " AND " & Val(txt_ExercicioFinal.Text)
strSQL = strSQL & " AND DDA.bytDebitoGeradoManualmente = 1 "

Set gobjBanco = New clsBanco
gobjBanco.CriaADO strSQL, 5, adoRecContribuinte

End Sub

Private Function lngRetornaSequencia(intContribuinte As Long, intExercicio As Integer) As Long

'******************************************************************************************
' Data: 02/05/2003
' Alteração: - Substituição do comando nativo ISNULL() do SQL Server pela função gstrISNULL
' Responsável: Everton Bianchini
'******************************************************************************************
    
Dim strMax As String
Dim adoRec As ADODB.Recordset

'Pesquisa a sequência da composição da receita
strMax = ""
'strMax = strMax & " SELECT ISNULL(MAX(strSequencia),0) + 1 AS Maximo FROM " & gstrLancamentoCalculo
strMax = strMax & " SELECT " & gstrISNULL("MAX(strSequencia)", "0") & " + 1 AS Maximo FROM " & gstrLancamentoCalculo
strMax = strMax & " WHERE intComposicaoReceita = " & dbc_strComposicaoReceita.BoundText
strMax = strMax & " AND intContribuinte = " & intContribuinte
strMax = strMax & " AND intExercicio = " & intExercicio
Set gobjBanco = New clsBanco
If gobjBanco.CriaADO(strMax, 10, adoRec) Then
    lngRetornaSequencia = adoRec!Maximo
End If

End Function

Private Function blnDadosOk() As Boolean
    Dim i As Integer
    blnDadosOk = False
    If chk_Selecionar.Value <> 1 Then
        If dbc_ContribuinteInicial.BoundText = "" Then
            ExibeMensagem "O campo " & lbl_ContribuinteInicial.Caption & " não pode ser em branco."
            dbc_ContribuinteInicial.SetFocus
            Exit Function
        End If
        If dbc_ContribuinteFinal.BoundText = "" Then
            ExibeMensagem "O campo " & lbl_ContribuinteFinal.Caption & " não pode ser em branco."
            dbc_ContribuinteFinal.SetFocus
            Exit Function
        End If
        If Val(dbc_ContribuinteFinal.BoundText) < Val(dbc_ContribuinteInicial.BoundText) Then
            ExibeMensagem "O Código do Contribuinte Final não pode ser inferior ao do Contribuinte Inicial."
            dbc_ContribuinteFinal.SetFocus
            Exit Function
        End If
    End If
    If txt_ExercicioInicial.Text = "" Then
        ExibeMensagem "O campo " & lbl_ExercicioInicial.Caption & " não pode ser em branco."
        txt_ExercicioInicial.SetFocus
        Exit Function
    End If
    If txt_ExercicioFinal.Text = "" Then
        ExibeMensagem "O campo " & lbl_ExercicioInicial.Caption & " não pode ser em branco."
        txt_ExercicioFinal.SetFocus
        Exit Function
    End If
    If dbc_strComposicaoReceita.BoundText = "" Then
        ExibeMensagem "O campo " & lbl_strComposicaoReceita.Caption & " não pode ser em branco."
        dbc_strComposicaoReceita.SetFocus
        Exit Function
    End If
    For i = 0 To xarReceita.Count(1) - 1
        If xarReceita(i, 2) = -1 Then
            blnDadosOk = True
            Exit Function
        End If
    Next
    ExibeMensagem "Selecione uma receita para efetuar o cálculo!"
End Function

'###################### CARACTER VÁLIDO E MARCA CAMPO ##############################

Private Sub dbc_ContribuinteInicial_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "", dbc_ContribuinteInicial
End Sub

Private Sub dbc_ContribuinteFinal_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "", dbc_ContribuinteFinal
End Sub

Private Sub txt_ExercicioInicial_GotFocus()
    MarcaCampo txt_ExercicioInicial
End Sub

Private Sub txt_ExercicioInicial_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txt_ExercicioInicial
End Sub

Private Sub txt_ExercicioFinal_GotFocus()
    MarcaCampo txt_ExercicioFinal
End Sub

Private Sub txt_ExercicioFinal_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txt_ExercicioFinal
End Sub

''' L
