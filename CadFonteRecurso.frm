VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frmCadFonteRecurso 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fontes de Recurso "
   ClientHeight    =   4680
   ClientLeft      =   945
   ClientTop       =   2025
   ClientWidth     =   6870
   HelpContextID   =   13
   Icon            =   "CadFonteRecurso.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   6870
   Begin VB.TextBox txtPKId 
      Height          =   270
      Left            =   5910
      TabIndex        =   8
      Top             =   150
      Visible         =   0   'False
      Width           =   645
   End
   Begin TabDlg.SSTab tab_3DPasta 
      Height          =   4515
      Left            =   90
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   90
      Width           =   6705
      _ExtentX        =   11827
      _ExtentY        =   7964
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "Fontes de Recurso "
      TabPicture(0)   =   "CadFonteRecurso.frx":1042
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblstrDescricao"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblstrCodigo"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblintGrupo"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txtstrDescricao"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "tdb_FonteRecurso"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtstrCodigo"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cbointGrupo"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmd_Grupo"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtintExercicio"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "chkbytRecursoProprio"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).ControlCount=   10
      Begin VB.CheckBox chkbytRecursoProprio 
         Caption         =   "Recurso Pr�prio"
         Height          =   195
         Left            =   2295
         TabIndex        =   2
         Top             =   930
         Width           =   1455
      End
      Begin VB.TextBox txtintExercicio 
         Height          =   285
         Left            =   4035
         TabIndex        =   11
         Top             =   15
         Visible         =   0   'False
         Width           =   840
      End
      Begin VB.CommandButton cmd_Grupo 
         Height          =   315
         Left            =   6240
         Picture         =   "CadFonteRecurso.frx":105E
         Style           =   1  'Graphical
         TabIndex        =   5
         TabStop         =   0   'False
         ToolTipText     =   "Ativa Cadastro de Grupos"
         Top             =   465
         Width           =   330
      End
      Begin VB.ComboBox cbointGrupo 
         CausesValidation=   0   'False
         Height          =   315
         ItemData        =   "CadFonteRecurso.frx":13E8
         Left            =   990
         List            =   "CadFonteRecurso.frx":13EA
         OLEDragMode     =   1  'Automatic
         TabIndex        =   0
         Top             =   465
         Width           =   5235
      End
      Begin VB.TextBox txtstrCodigo 
         Height          =   285
         Left            =   990
         MaxLength       =   2
         OLEDragMode     =   1  'Automatic
         TabIndex        =   1
         Top             =   870
         Width           =   1065
      End
      Begin TrueOleDBGrid70.TDBGrid tdb_FonteRecurso 
         Height          =   2715
         Left            =   120
         TabIndex        =   4
         Top             =   1650
         Width           =   6435
         _ExtentX        =   11351
         _ExtentY        =   4789
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "C�digo"
         Columns(0).DataField=   "PKID"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "C�digo"
         Columns(1).DataField=   "strCodigo"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Descri��o"
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
         Splits(0)._ColumnProps(13)=   "Column(2).Width=9049"
         Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=8969"
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
         _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bgcolor=&H8000000A&"
         _StyleDefs(11)  =   ":id=2,.fgcolor=&H80000008&,.bold=0,.fontsize=825,.italic=0,.underline=0"
         _StyleDefs(12)  =   ":id=2,.strikethrough=0,.charset=0"
         _StyleDefs(13)  =   ":id=2,.fontname=MS Sans Serif"
         _StyleDefs(14)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(15)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(16)  =   ":id=3,.fontname=MS Sans Serif"
         _StyleDefs(17)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(18)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
         _StyleDefs(19)  =   "EditorStyle:id=7,.parent=1"
         _StyleDefs(20)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38,.bgcolor=&H80000002&"
         _StyleDefs(21)  =   ":id=8,.fgcolor=&H80000009&"
         _StyleDefs(22)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
         _StyleDefs(23)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
         _StyleDefs(24)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(25)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
         _StyleDefs(26)  =   "Splits(0).Style:id=13,.parent=1"
         _StyleDefs(27)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
         _StyleDefs(28)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
         _StyleDefs(29)  =   "Splits(0).FooterStyle:id=15,.parent=3"
         _StyleDefs(30)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
         _StyleDefs(31)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
         _StyleDefs(32)  =   "Splits(0).EditorStyle:id=17,.parent=7"
         _StyleDefs(33)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
         _StyleDefs(34)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
         _StyleDefs(35)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
         _StyleDefs(36)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(37)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(38)  =   "Splits(0).Columns(0).Style:id=46,.parent=13"
         _StyleDefs(39)  =   "Splits(0).Columns(0).HeadingStyle:id=43,.parent=14"
         _StyleDefs(40)  =   "Splits(0).Columns(0).FooterStyle:id=44,.parent=15"
         _StyleDefs(41)  =   "Splits(0).Columns(0).EditorStyle:id=45,.parent=17"
         _StyleDefs(42)  =   "Splits(0).Columns(1).Style:id=28,.parent=13"
         _StyleDefs(43)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14"
         _StyleDefs(44)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
         _StyleDefs(45)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
         _StyleDefs(46)  =   "Splits(0).Columns(2).Style:id=32,.parent=13"
         _StyleDefs(47)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
         _StyleDefs(48)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
         _StyleDefs(49)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
         _StyleDefs(50)  =   "Named:id=33:Normal"
         _StyleDefs(51)  =   ":id=33,.parent=0"
         _StyleDefs(52)  =   "Named:id=34:Heading"
         _StyleDefs(53)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(54)  =   ":id=34,.wraptext=-1"
         _StyleDefs(55)  =   "Named:id=35:Footing"
         _StyleDefs(56)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(57)  =   "Named:id=36:Selected"
         _StyleDefs(58)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(59)  =   "Named:id=37:Caption"
         _StyleDefs(60)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(61)  =   "Named:id=38:HighlightRow"
         _StyleDefs(62)  =   ":id=38,.parent=33,.bgcolor=&H80000014&,.fgcolor=&H80000012&"
         _StyleDefs(63)  =   "Named:id=39:EvenRow"
         _StyleDefs(64)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(65)  =   "Named:id=40:OddRow"
         _StyleDefs(66)  =   ":id=40,.parent=33"
         _StyleDefs(67)  =   "Named:id=41:RecordSelector"
         _StyleDefs(68)  =   ":id=41,.parent=34"
         _StyleDefs(69)  =   "Named:id=42:FilterBar"
         _StyleDefs(70)  =   ":id=42,.parent=33"
      End
      Begin VB.TextBox txtstrDescricao 
         Height          =   285
         Left            =   990
         MaxLength       =   100
         TabIndex        =   3
         Top             =   1260
         Width           =   5565
      End
      Begin VB.Label lblintGrupo 
         AutoSize        =   -1  'True
         Caption         =   "Grupo"
         Height          =   195
         Left            =   435
         TabIndex        =   10
         Top             =   600
         Width           =   435
      End
      Begin VB.Label lblstrCodigo 
         AutoSize        =   -1  'True
         Caption         =   "C�digo"
         Height          =   195
         Left            =   375
         TabIndex        =   9
         Top             =   960
         Width           =   495
      End
      Begin VB.Label lblstrDescricao 
         AutoSize        =   -1  'True
         Caption         =   "Descri��o"
         Height          =   195
         Left            =   150
         TabIndex        =   7
         Top             =   1350
         Width           =   720
      End
   End
End
Attribute VB_Name = "frmCadFonteRecurso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim mblnAlterando     As Boolean
    Dim mblnPrimeiraVez   As Boolean
    Dim mobjAux           As Object
    Dim mblnClickOk       As Boolean
    
    Public mIntCodSeguranca  As Integer

Private Sub cbointGrupo_Click()
    
    If gbytMenu = gbytMenuCadastro Then
        txtintExercicio = gintExercicio
    Else
        txtintExercicio = gintExercicio + 1
    End If
    
    LeDaTabelaParaObj "", tdb_FonteRecurso, strQuery
    
    txtstrCodigo_GotFocus
    
End Sub

Private Sub cbointGrupo_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub cmd_Grupo_Click()
    CarregaForm frmCadGrupoDeFonteRecurso, cbointGrupo
End Sub

Private Sub Form_Activate()
    
    gintCodSeguranca = mIntCodSeguranca
    
    VirificaGradeListView Me
    
    HabilitaDesabilitaBotao1 mblnAlterando, gstrMnuArquivo, gstrDeletar
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
    
    strSQL = "SELECT PKId, strCodigo, strDescricao, bytRecursoProprio "
    strSQL = strSQL & "FROM " & gstrFonteRecurso & " "
    strSQL = strSQL & "WHERE intExercicio=" & txtintExercicio
    
    If cbointGrupo.ListIndex > 0 Then
        strSQL = strSQL & " AND intGrupo = " & gstrItemData(cbointGrupo)
    End If

    strSQL = strSQL & " ORDER BY " & gstrCONVERT(cdt_numeric, "strCodigo")
   
    strQuery = strSQL
    
End Function
 
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
  If KeyCode = vbKeyF1 Then
     Call_HtmlHelp Me.HelpContextID
  End If
End Sub

Private Sub tdb_FonteRecurso_Click()
    
    mblnClickOk = True
    
    If glngQtdLinhaTDBGrid(tdb_FonteRecurso) = 1 Then
        tdb_fonterecurso_RowColChange 0, 0
    End If
    
End Sub

Private Sub tdb_FonteRecurso_DblClick()
    MantemForm gstrAplicar
End Sub

Private Sub tdb_FonteRecurso_KeyDown(KeyCode As Integer, Shift As Integer)
    'mblnClickOk = True
End Sub

Private Sub tdb_FonteRecurso_KeyPress(KeyAscii As Integer)
  CaracterValido KeyAscii, "A", tdb_FonteRecurso
End Sub

Private Sub tdb_FonteRecurso_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    mblnClickOk = True
End Sub

Private Sub txtstrCodigo_GotFocus()

    gstrProximoCodigo txtstrCodigo, gstrFonteRecurso, "strCodigo", gintCodSeguranca, "intGrupo", gstrItemData(cbointGrupo), , , , , "intExercicio", CStr(txtintExercicio)
    txtstrCodigo = Format(gstrValorSemMascara(txtstrCodigo), "00")
    
    'MarcaCampo txtstrCodigo
    
End Sub
Private Sub txtstrCodigo_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtstrCodigo
End Sub

Private Sub txtstrCodigo_LostFocus()
    txtstrCodigo.Text = Format(txtstrCodigo.Text, "00")
End Sub

Private Sub txtstrDescricao_GotFocus()
    MarcaCampo txtstrDescricao
End Sub

Private Sub txtstrDescricao_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrDescricao
End Sub

Private Sub Form_Load()
    mblnAlterando = False
    LeDaTabelaParaObj gstrGrupoDeFonteRecurso, cbointGrupo
    VerificaObjParaAplicar mobjAux
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar
End Sub

Private Sub tdb_fonterecurso_FilterChange()
    mblnClickOk = False
    gblnFilraCampos tdb_FonteRecurso
End Sub

Private Sub tdb_fonterecurso_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If gbytMenu = gbytMenuCadastro Then
        txtintExercicio = gintExercicio
    Else
        txtintExercicio = gintExercicio + 1
    End If
    With tdb_FonteRecurso
        If (Not .EOF And Not .BOF) And mblnClickOk Then
            mblnClickOk = False
            With tdb_FonteRecurso
                txtPKId.Text = .Columns("PKID").Value
                LeDaTabelaParaObj gstrFonteRecurso, Me
                
                txtstrDescricao.Tag = txtstrDescricao.Text

                HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrDeletar
                If mobjAux Is Nothing Then
                    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar
                Else
                    HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrAplicar
                End If
                mblnAlterando = True
            End With
        End If
    End With
 TrocaCorObjeto txtstrCodigo, False
    
End Sub

Public Sub MantemForm(ByVal strModoOperacao As String)
    
    If gbytMenu = gbytMenuCadastro Then
        txtintExercicio = gintExercicio
    Else
        txtintExercicio = gintExercicio + 1
    End If
    
    If strModoOperacao = gstrSalvar Then
        If DadosOk Then
            ToolBarGeral strModoOperacao, gstrFonteRecurso, _
                         mblnAlterando, tdb_FonteRecurso, Me, _
                         mobjAux, strQuery, , _
                         rptFonteRecurso, strQueryRelatorio
        End If
    Else
        ToolBarGeral strModoOperacao, gstrFonteRecurso, _
                     mblnAlterando, tdb_FonteRecurso, Me, _
                     mobjAux, strQuery, strQueryAplicar, _
                     rptFonteRecurso, strQueryRelatorio
        
        If UCase(strModoOperacao) = UCase(gstrNovo) Or UCase(strModoOperacao) = UCase(gstrLimpar) Or UCase(strModoOperacao) = UCase(gstrDeletar) Then
            txtstrDescricao.Tag = Space$(0)
        End If
    End If
End Sub

Function strQueryRelatorio() As String
Dim strSQL As String
    
    strSQL = ""
    strSQL = strSQL & " SELECT "
    strSQL = strSQL & " FR.strCodigo, FR.strDescricao, "
    strSQL = strSQL & " GF.strCodigo AS CodGRUPO, GF.strDescricao AS GRUPO "
    
    strSQL = strSQL & "FROM "
    strSQL = strSQL & gstrFonteRecurso & " FR, "
    strSQL = strSQL & gstrGrupoDeFonteRecurso & " GF "
    
    strSQL = strSQL & " WHERE "
    strSQL = strSQL & " FR.intGrupo = GF.PKId AND intExercicio =" & txtintExercicio
    
   strSQL = strSQL & " ORDER BY "
   strSQL = strSQL & gstrCONVERT(CDT_INT, "GF.strCodigo") & ", GF.strDescricao, " & gstrCONVERT(CDT_INT, "FR.strCodigo") & " , FR.strDescricao "
       
    strQueryRelatorio = strSQL
    
End Function

Private Function DadosOk() As Boolean
    
    If cbointGrupo.ListIndex = -1 Then
        ExibeMensagem "O campo ""Grupo"" � obrigat�rio!"
        cbointGrupo.SetFocus
        DadosOk = False
        Exit Function
    End If
    
    If Len(Trim(txtstrCodigo.Text)) = 0 Then
        ExibeMensagem "O campo ""C�digo"" � obrigat�rio!"
        txtstrCodigo.SetFocus
        DadosOk = False
        Exit Function
    End If
    
    If Len(Trim(txtstrDescricao.Text)) = 0 Then
        ExibeMensagem "O campo ""Descri��o"" � obrigat�rio!"
        txtstrDescricao.SetFocus
        DadosOk = False
        Exit Function
    End If
    
    
        Select Case Len(txtstrCodigo)
            Case 1
                If gblnExisteCodigo(1, gstrFonteRecurso, "strCodigo", "'" & "0" & (txtstrCodigo) & "'", "intGrupo", gstrItemData(cbointGrupo), , , " AND intExercicio = " & txtintExercicio) Then
                    ExibeMensagem "O c�digo digitado j� se encontra cadastrado!"
                    txtstrCodigo.SetFocus
                    DadosOk = False
                    Exit Function
                End If
            Case Else
                If gblnExisteCodigo(1, gstrFonteRecurso, "strCodigo", "'" & gvntConvFormatoEspecificoParaSQL(txtstrCodigo) & "'", "intGrupo", gstrItemData(cbointGrupo), , , " AND intExercicio = " & txtintExercicio) Then
                    ExibeMensagem "O c�digo digitado j� se encontra cadastrado!"
                    txtstrCodigo.SetFocus
                    DadosOk = False
                    Exit Function
                End If
      End Select
         
    
    
    If (txtstrDescricao.Text <> txtstrDescricao.Tag) Then

            If gblnExisteCodigo(1, gstrFonteRecurso, "strDescricao", "'" & txtstrDescricao & "'", , , , , " AND intExercicio = " & txtintExercicio) Then
                ExibeMensagem "A descri��o digitada j� se encontra cadastrada!"
                txtstrDescricao.SetFocus
                DadosOk = False
                Exit Function
            End If

    End If
    
    DadosOk = True
    
End Function

Private Sub tdb_FonteRecurso_HeadClick(ByVal ColIndex As Integer)
   
   gOrdenaGrid tdb_FonteRecurso, ColIndex
   
End Sub

Private Function strQueryAplicar() As String

    strQueryAplicar = "SELECT PKId, strDescricao FROM " & gstrFonteRecurso & " WHERE intExercicio=" & txtintExercicio

End Function
