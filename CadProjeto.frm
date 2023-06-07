VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MsDatLst.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frmCadProjetoAtividade 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Projetos e Atividades"
   ClientHeight    =   4800
   ClientLeft      =   2145
   ClientTop       =   3870
   ClientWidth     =   7065
   HelpContextID   =   17
   Icon            =   "CadProjeto.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   7065
   Begin VB.TextBox txtPKId 
      Height          =   270
      Left            =   4680
      TabIndex        =   7
      Top             =   120
      Visible         =   0   'False
      Width           =   645
   End
   Begin TabDlg.SSTab tab_3dPasta 
      Height          =   4605
      Left            =   120
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   90
      Width           =   6840
      _ExtentX        =   12065
      _ExtentY        =   8123
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Projeto e Atividade "
      TabPicture(0)   =   "CadProjeto.frx":1042
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblstrNome"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblstrCodigo"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblstrObjetivo"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "tdb_Lista"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtstrDescricao"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtstrCodigo"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "chkblnProjeto"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cbointObjetivo"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cmd_Objetico"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtintExercicio"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).ControlCount=   10
      Begin VB.TextBox txtintExercicio 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2250
         TabIndex        =   11
         Top             =   390
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.CommandButton cmd_Objetico 
         Height          =   300
         Left            =   6390
         Picture         =   "CadProjeto.frx":105E
         Style           =   1  'Graphical
         TabIndex        =   10
         TabStop         =   0   'False
         Tag             =   "204"
         ToolTipText     =   "Clique para consultar objetivo"
         Top             =   1080
         Width           =   330
      End
      Begin MSDataListLib.DataCombo cbointObjetivo 
         Height          =   315
         Left            =   990
         TabIndex        =   3
         Top             =   1080
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin VB.CheckBox chkblnProjeto 
         Caption         =   "Projeto"
         Height          =   285
         Left            =   3030
         TabIndex        =   1
         Top             =   420
         Visible         =   0   'False
         Width           =   3105
      End
      Begin VB.TextBox txtstrCodigo 
         Height          =   285
         Left            =   990
         MaxLength       =   10
         TabIndex        =   0
         Top             =   375
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
         Top             =   720
         Width           =   5715
      End
      Begin TrueOleDBGrid70.TDBGrid tdb_Lista 
         Height          =   2985
         Left            =   120
         TabIndex        =   4
         Top             =   1470
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   5265
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
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0).AllowSizing=0"
         Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
         Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(8)=   "Column(1).Width=1773"
         Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=1693"
         Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(13)=   "Column(2).Width=9340"
         Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=9260"
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
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=48,.bold=0,.fontsize=825,.italic=0"
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
         _StyleDefs(38)  =   "Splits(0).Columns(2).Style:id=32,.parent=13"
         _StyleDefs(39)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
         _StyleDefs(40)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
         _StyleDefs(41)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
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
         _StyleDefs(54)  =   ":id=38,.parent=33,.bgcolor=&H80000014&,.fgcolor=&H80000012&"
         _StyleDefs(55)  =   "Named:id=39:EvenRow"
         _StyleDefs(56)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(57)  =   "Named:id=40:OddRow"
         _StyleDefs(58)  =   ":id=40,.parent=33"
         _StyleDefs(59)  =   "Named:id=41:RecordSelector"
         _StyleDefs(60)  =   ":id=41,.parent=34"
         _StyleDefs(61)  =   "Named:id=42:FilterBar"
         _StyleDefs(62)  =   ":id=42,.parent=33"
      End
      Begin VB.Label lblstrObjetivo 
         AutoSize        =   -1  'True
         Caption         =   "Objetivo"
         Height          =   195
         Left            =   330
         TabIndex        =   9
         Top             =   1170
         Width           =   585
      End
      Begin VB.Label lblstrCodigo 
         AutoSize        =   -1  'True
         Caption         =   "Código"
         Height          =   195
         Left            =   420
         TabIndex        =   8
         Top             =   420
         Width           =   495
      End
      Begin VB.Label lblstrNome 
         AutoSize        =   -1  'True
         Caption         =   "Descrição"
         Height          =   195
         Left            =   195
         TabIndex        =   6
         Top             =   795
         Width           =   720
      End
   End
End
Attribute VB_Name = "frmCadProjetoAtividade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    
    Dim mblnAlterando              As Boolean
    Dim mobjAux                    As Object
    Dim mblnselecionou             As Boolean
    Dim mblnClickOk                As Boolean
    Dim intFiltroExercicio         As Integer
    Public blnProgramaDeTrabalho   As Boolean ' Quando True esta sendo incluido pelo form Programa de Trabalho
    Public strValorFonteRecurso    As String ' Valor que vem do programa de trabalho, para ser inserido junto com os dados gerados.
    Public mIntCodSeguranca        As Integer
    
Private Sub cbointObjetivo_click(Area As Integer)
    DropDownDataCombo cbointObjetivo, Me, Area
End Sub

Private Sub cbointObjetivo_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo cbointObjetivo, Me, , KeyCode, Shift
End Sub

Private Sub cbointObjetivo_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub cmd_Objetico_Click()
    CarregaForm frmCadObjetivo, cbointObjetivo
End Sub

Private Sub Form_Activate()
    
    gintCodSeguranca = mIntCodSeguranca
    
    LimpaCampos
    
    VirificaGradeListView Me
    
    If mblnselecionou Then
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
    
    strSQL = "SELECT PKId, strCodigo, strDescricao FROM " & gstrProjeto & " WHERE intExercicio = " & intFiltroExercicio
    
    strSQL = strSQL & " ORDER BY " & gstrCONVERT(cdt_numeric, "strCodigo")
    
    strQuery = strSQL
    
End Function

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = vbKeyF1 Then
        Call_HtmlHelp Me.HelpContextID
    End If
End Sub

Private Sub Form_Load()
    
    mblnAlterando = False
    
    'Vamos verificar qual menu que chamou o form, para definirmos o filtro
    If gbytMenu = gbytMenuCadastro Then
        intFiltroExercicio = gintExercicio
    Else
        intFiltroExercicio = gintExercicio + 1
    End If
    
    txtintExercicio = intFiltroExercicio
        
    VerificaListaAutomatica gstrProjeto, tdb_Lista, strQuery
    LeDaTabelaParaObj gstrObjetivo, cbointObjetivo
    VerificaObjParaAplicar mobjAux
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    blnProgramaDeTrabalho = False
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

Private Sub tdb_Lista_HeadClick(ByVal ColIndex As Integer)
   gOrdenaGrid tdb_Lista, ColIndex
End Sub

Private Sub tdb_Lista_KeyDown(KeyCode As Integer, Shift As Integer)
    mblnClickOk = True
End Sub

Private Sub tdb_Lista_KeyPress(KeyAscii As Integer)
    If tdb_Lista.Col = 1 Then
        CaracterValido KeyAscii, "A", tdb_Lista
    End If
End Sub

Private Sub tdb_Lista_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mblnClickOk = True
End Sub

Private Sub tdb_Lista_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    Dim rsTmp As ADODB.Recordset
    With tdb_Lista
        If (Not .EOF And Not .BOF) And mblnClickOk Then
            mblnClickOk = False
            txtPKId.Text = .Columns("PKID").Value
            LeDaTabelaParaObj gstrProjeto, Me
            
            If cbointObjetivo.BoundText <> "" Then
                If gobjBanco.CriaADO("SELECT strDescricao FROM " & gstrObjetivo & " WHERE PKid=" & cbointObjetivo.BoundText, 60, rsTmp) = True Then
                    cbointObjetivo.Text = rsTmp.Fields("strDescricao").Value
                End If
            End If
            gCorLinhaSelecionada tdb_Lista
            HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrDeletar
            If mobjAux Is Nothing Then
                HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar
            Else
                HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrAplicar
            End If
            mblnAlterando = True
        End If
    End With
    If Not rsTmp Is Nothing Then
        If rsTmp.State = adStateOpen Then
            rsTmp.Close
        End If
        Set rsTmp = Nothing
    End If
End Sub

Public Sub MantemForm(ByVal strModoOperacao As String)
    'Vamos verificar qual menu que chamou o form, para definirmos o filtro
    If gbytMenu = gbytMenuCadastro Then
        intFiltroExercicio = gintExercicio
    Else
        intFiltroExercicio = gintExercicio + 1
    End If
    Select Case strModoOperacao
    
    Case Is = gstrNovo
        txtstrCodigo.Enabled = True
        
        LimpaCampos
        Exit Sub
    Case Is = gstrSalvar
        
        If Not blnDadosOk Then Exit Sub
        
'        If Not mblnAlterando Then
'            If gblnExisteCodigo(2, gstrProjeto, "strCodigo", "'" & txtstrCodigo & "'", "intExercicio", Str(intFiltroExercicio)) Then  'Or gblnExisteCodigo(1, gstrProjeto, "strDescricao", "'" & txtstrDescricao & "'", "intExercicio", Str(intFiltroExercicio)) Then
'                MsgBox "Este Código já se encontra cadastrado.", vbOKOnly, "Mensagem ao Usuário"
'                Exit Sub
'            End If
'        End If
        
    End Select

    ToolBarGeral strModoOperacao, gstrProjeto, mblnAlterando, tdb_Lista, Me, mobjAux, strQuery, strQueryAplicar, rptProjetoAtividade, strQueryRelatorio
    
    'Vamos criar um projeto de atividade para cada elemento de despesa distinto
    If strModoOperacao = gstrSalvar And Not mblnAlterando And blnProgramaDeTrabalho Then
        
        CriaProjetoAtividadePorElementoDespesa
        
        frmCadProgramaDeTrabalho.AtualizaListaPosProjetoAtividade

        Unload Me
        
        Exit Sub
        
    End If
    
    If strModoOperacao = gstrNovo Or strModoOperacao = gstrDeletar Then txtintExercicio = intFiltroExercicio

End Sub

Private Sub txtstrDescricao_GotFocus()
    MarcaCampo txtstrDescricao
End Sub

Private Sub txtstrDescricao_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrDescricao
End Sub

Private Sub txtstrCodigo_GotFocus()
    MarcaCampo txtstrCodigo
End Sub

Private Sub txtstrCodigo_KeyPress(KeyAscii As Integer)
    If txtstrCodigo.SelStart = 0 Then
        'If KeyAscii = 48 Then
        '    KeyAscii = 0
        If InStr(UCase("50,52,54,56"), KeyAscii) Then
            chkblnProjeto = 0
        Else
            chkblnProjeto = 1
        End If
    End If
    CaracterValido KeyAscii, "N", txtstrCodigo
End Sub

Private Function blnDadosOk() As Boolean
Dim strWhereComplementar    As String
    
    'Incluido orc1556 para impedir inclusão de descricoes repetidas no mesmo exercicio
    If mblnAlterando Then
        strWhereComplementar = " AND PKID <> " & Me.txtPKId.Text
    Else
        strWhereComplementar = ""
    End If
    blnDadosOk = False

    If Trim(txtstrCodigo) = "" Then
        ExibeMensagem "O código tem que ser digitado."
        txtstrCodigo.SetFocus
        Exit Function
    End If
    
    If Trim(txtstrDescricao) = "" Then
        ExibeMensagem "A descrição tem que ser digitada."
        txtstrDescricao.SetFocus
        Exit Function
    End If
    
    
    If gblnExisteCodigo(1, gstrProjeto, "strCodigo", "'" & gvntConvFormatoEspecificoParaSQL(txtstrCodigo) & "'", , , , , " AND intExercicio = " & intFiltroExercicio & strWhereComplementar) Then
        ExibeMensagem "O código digitado já se encontra cadastrado!"
        txtstrCodigo.SetFocus
        Exit Function
    End If


    If gblnExisteCodigo(1, gstrProjeto, "strDescricao", "'" & gvntConvFormatoEspecificoParaSQL(txtstrDescricao) & "'", , , , , " AND intExercicio = " & intFiltroExercicio & strWhereComplementar) Then
        ExibeMensagem "A Descrição digitado já se encontra cadastrado!"
        txtstrCodigo.SetFocus
        Exit Function
    End If

    blnDadosOk = True

End Function

Private Sub LimpaCampos()
    
    txtstrCodigo = Space$(0)
    
    gstrProximoCodigo txtstrCodigo, gstrProjeto, "strCodigo", gintCodSeguranca, , , , , , , "intExercicio", CStr(intFiltroExercicio)
    
    txtstrDescricao = Space$(0)
    cbointObjetivo.BoundText = Space$(0)
    chkblnProjeto.Value = vbUnchecked
    txtPKId = Space$(0)
    
    mblnAlterando = False
    
End Sub

Function strQueryRelatorio() As String

Dim strSQL As String

    strSQL = "SELECT PJ.strCodigo, PJ.strDescricao, OJ.strDescricao AS strObjetivo, "
    strSQL = strSQL & intFiltroExercicio & " AS Exercicio "
    strSQL = strSQL & "FROM " & gstrProjeto & " PJ, " & gstrObjetivo & " OJ "
    
    strSQL = strSQL & "WHERE PJ.intObjetivo " & strOUTJSQLServer & "= OJ.PKId" & strOUTJOracle
    strSQL = strSQL & " AND PJ.intExercicio = " & intFiltroExercicio
    strSQL = strSQL & " ORDER BY " & gstrCONVERT(CDT_INT, "PJ.strCodigo")
    
    strQueryRelatorio = strSQL
    
End Function

Private Sub CriaProjetoAtividadePorElementoDespesa()

Dim adoTemp As ADODB.Recordset
Dim strSQL  As String
Dim lngPkid As Long

On Error GoTo TrataErroLocal
    
    lngPkid = glngPegaUltimaChave(gstrProjeto, "PkId")
    
    strSQL = "INSERT INTO " & gstrProgramaDeTrabalho & "(intOrgao, intUnidadeOrcamentaria, intSubUnidade, " & _
             "intTipoCredito, intFuncao, intSubFuncao, intPrograma, intSubPrograma, intProjetoAtividade, intElementoDespesa, dblValor, intExercicio, bytSituacao, strCodigo, intFonteRecurso) " & _
             "SELECT DISTINCT intOrgao, intUnidadeOrcamentaria, intSubUnidade, " & _
             "intTipoCredito, intFuncao, intSubFuncao, intPrograma, intSubPrograma, " & _
             lngPkid & " intProjetoAtividade, intElementoDespesa, 0, " & _
             intFiltroExercicio & " intExercicio, 0 bytSituacao, '" & Replace(frmCadProgramaDeTrabalho.txtstrCodigo, "." & frmCadProgramaDeTrabalho.txt_intProjetoAtividade & ".", "." & Format(glngPegaUltimaChave(gstrProjeto, "strCodigo", "PkId", lngPkid), "0000") & ".") & "' strCodigo " & _
             ", " & strValorFonteRecurso & " FonteRecurso " & _
             "FROM " & gstrProgramaDeTrabalho & _
             " WHERE intOrgao = " & frmCadProgramaDeTrabalho.dbcintOrgao.ItemData(frmCadProgramaDeTrabalho.dbcintOrgao.ListIndex) & _
             " AND intUnidadeOrcamentaria = " & frmCadProgramaDeTrabalho.dbcintUnidadeOrcamentaria.ItemData(frmCadProgramaDeTrabalho.dbcintUnidadeOrcamentaria.ListIndex) & _
             " AND intSubUnidade = " & frmCadProgramaDeTrabalho.dbcintSubunidade.ItemData(frmCadProgramaDeTrabalho.dbcintSubunidade.ListIndex) & _
             " AND intFuncao = " & frmCadProgramaDeTrabalho.dbcintFuncao.ItemData(frmCadProgramaDeTrabalho.dbcintFuncao.ListIndex) & _
             " AND intSubFuncao = " & frmCadProgramaDeTrabalho.dbcintSubFuncao.ItemData(frmCadProgramaDeTrabalho.dbcintSubFuncao.ListIndex) & _
             " AND intPrograma = " & frmCadProgramaDeTrabalho.dbcintPrograma.ItemData(frmCadProgramaDeTrabalho.dbcintPrograma.ListIndex) & _
             " AND intProjetoAtividade = " & frmCadProgramaDeTrabalho.dbcintProjetoAtividade.ItemData(frmCadProgramaDeTrabalho.dbcintProjetoAtividade.ListIndex) & _
             " AND intExercicio = " & intFiltroExercicio
                        
    Set gobjBanco = New clsBanco
    
    gobjBanco.ExecutaBeginTrans
    gobjBanco.Execute strSQL
    gobjBanco.ExecutaCommitTrans
    
    Set gobjBanco = Nothing

    Exit Sub
    
TrataErroLocal:
    gobjBanco.ExecutaRollbackTrans
    
End Sub

Private Function strQueryAplicar() As String

    strQueryAplicar = "SELECT PKId, strDescricao FROM " & gstrProjeto & " WHERE intExercicio=" & intFiltroExercicio

End Function
