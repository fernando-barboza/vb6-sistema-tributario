VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "TODG7.OCX"
Begin VB.Form frmCadFaixaDeValores 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Faixa de Valores"
   ClientHeight    =   5385
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6315
   HelpContextID   =   27
   Icon            =   "CadFaixaDeValores.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   359
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   421
   Begin VB.TextBox txtPKId 
      Height          =   270
      Left            =   2520
      TabIndex        =   3
      Top             =   60
      Visible         =   0   'False
      Width           =   645
   End
   Begin TabDlg.SSTab tab_3DPasta 
      Height          =   5145
      Left            =   120
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   6105
      _ExtentX        =   10769
      _ExtentY        =   9075
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   529
      TabCaption(0)   =   "Faixa de Valor"
      TabPicture(0)   =   "CadFaixaDeValores.frx":1042
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblstrNomeDaFaixa"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "txtstrNomeDaFaixa"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "tdb_FaixaDeValores"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      Begin TrueOleDBGrid70.TDBGrid tdb_FaixaDeValores 
         Height          =   4185
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Width           =   5835
         _ExtentX        =   10292
         _ExtentY        =   7382
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
         Columns(1).Caption=   "Descri��o"
         Columns(1).DataField=   "strNomeDaFaixa"
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
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0).Visible=0"
         Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(7)=   "Column(1).Width=9737"
         Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=9657"
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
         _StyleDefs(30)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
         _StyleDefs(31)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
         _StyleDefs(32)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
         _StyleDefs(33)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
         _StyleDefs(34)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
         _StyleDefs(35)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
         _StyleDefs(36)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
         _StyleDefs(37)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
         _StyleDefs(38)  =   "Named:id=33:Normal"
         _StyleDefs(39)  =   ":id=33,.parent=0"
         _StyleDefs(40)  =   "Named:id=34:Heading"
         _StyleDefs(41)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(42)  =   ":id=34,.wraptext=-1"
         _StyleDefs(43)  =   "Named:id=35:Footing"
         _StyleDefs(44)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(45)  =   "Named:id=36:Selected"
         _StyleDefs(46)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(47)  =   "Named:id=37:Caption"
         _StyleDefs(48)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(49)  =   "Named:id=38:HighlightRow"
         _StyleDefs(50)  =   ":id=38,.parent=33,.bgcolor=&H80000014&,.fgcolor=&H80000012&"
         _StyleDefs(51)  =   "Named:id=39:EvenRow"
         _StyleDefs(52)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(53)  =   "Named:id=40:OddRow"
         _StyleDefs(54)  =   ":id=40,.parent=33"
         _StyleDefs(55)  =   "Named:id=41:RecordSelector"
         _StyleDefs(56)  =   ":id=41,.parent=34"
         _StyleDefs(57)  =   "Named:id=42:FilterBar"
         _StyleDefs(58)  =   ":id=42,.parent=33"
      End
      Begin VB.TextBox txtstrNomeDaFaixa 
         Height          =   285
         Left            =   930
         MaxLength       =   30
         TabIndex        =   1
         Top             =   480
         Width           =   5025
      End
      Begin VB.Label lblstrNomeDaFaixa 
         AutoSize        =   -1  'True
         Caption         =   "Descri��o"
         Height          =   195
         Left            =   150
         TabIndex        =   4
         Top             =   585
         Width           =   720
      End
   End
End
Attribute VB_Name = "frmCadFaixaDeValores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim mblnAlterando   As Boolean
    Dim mobjAux         As Object
    Dim mblnSelecionou As Boolean
    Dim mblnPrimeiraVez As Boolean

Private Sub Form_Activate()
VirificaGradeListView Me
'=============
    gintCodSeguranca = 594

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

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
  If KeyCode = vbKeyF1 Then
     Call_HtmlHelp Me.HelpContextID
  End If
End Sub

Private Sub Form_Load()
    LeDaTabelaParaObj gstrFaixaDeValor, tdb_FaixaDeValores, strQuery
    VerificaObjParaAplicar mobjAux
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar, gstrDeletar
    mblnSelecionou = False
    mblnPrimeiraVez = False
End Sub

Private Sub tdb_FaixaDeValores_Click()
mblnPrimeiraVez = True
    With tdb_FaixaDeValores
        If Not .BOF And Not .EOF Then
            If .Bookmark = 1 Then
                tdb_FaixaDeValores_RowColChange 0, 0
            End If
        End If
    End With
End Sub

Private Sub tdb_FaixaDeValores_DblClick()
    MantemForm gstrAplicar
End Sub

Private Sub tdb_FaixaDeValores_FilterChange()
    gblnFilraCampos tdb_FaixaDeValores
End Sub

Private Sub tdb_FaixaDeValores_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", tdb_FaixaDeValores
End Sub

Private Sub tdb_FaixaDeValores_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    
    With tdb_FaixaDeValores
        If Not .EOF And Not .BOF Then
            txtPKId.Text = .Columns("PKID").Value
            If mblnPrimeiraVez Then
                mblnAlterando = True
                LeDaTabelaParaObj gstrFaixaDeValor, Me
'=============
                gCorLinhaSelecionada tdb_FaixaDeValores
                
                HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrDeletar
                
                If mobjAux Is Nothing Then
                    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar
                Else

                    HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrAplicar
                End If
'=============
                mblnSelecionou = True
            End If
        End If
    End With

End Sub

Public Sub MantemForm(ByVal strModoOperacao As String)
Dim varBookMark As Variant
Dim strSql As String

If Not tdb_FaixaDeValores.EOF Then
    varBookMark = tdb_FaixaDeValores.Bookmark
Else
    mblnAlterando = False
End If

strSql = strQuery
If UCase(strModoOperacao) = "SALVAR" Or UCase(strModoOperacao) = "DELETAR" Then
    If Trim(txtstrNomeDaFaixa.Text) = "" Then
        ExibeMensagem "Digite uma descri��o v�lida antes de salvar."
        txtstrNomeDaFaixa.SetFocus
        Exit Sub
    End If
    
    mblnPrimeiraVez = False
End If
ToolBarGeral strModoOperacao, gstrFaixaDeValor, mblnAlterando, tdb_FaixaDeValores, Me, mobjAux, strSql, strSql, rptFaixaDeValores, strQueryRelatorio
HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrDeletar
If UCase(strModoOperacao) <> "FECHAR" Then
    If Not IsEmpty(varBookMark) Then
        If UCase(strModoOperacao) = "DELETAR" Then
            tdb_FaixaDeValores.MoveFirst
        Else
            tdb_FaixaDeValores.Bookmark = varBookMark
        End If
        If mblnAlterando Then
            tdb_FaixaDeValores_RowColChange 0, 0
        End If
    End If
End If
End Sub

Private Sub txtstrNomeDaFaixa_GotFocus()
    MarcaCampo txtstrNomeDaFaixa
End Sub

Private Sub txtstrNomeDaFaixa_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrNomeDaFaixa
End Sub

Function strQueryRelatorio() As String
Dim strSql As String
    strSql = ""
    strSql = strSql & "SELECT * FROM " & gstrFaixaDeValor
    If mblnAlterando = True Then
        strSql = strSql & " WHERE PKId = " & Val(txtPKId)
    End If
    strSql = strSql & " ORDER BY strNomeDaFaixa "
strQueryRelatorio = strSql
End Function

Private Function strQuery() As String
Dim strSql  As String
    strSql = ""
    strSql = strSql & " SELECT "
    strSql = strSql & " PKId, strNomeDaFaixa "
    strSql = strSql & " FROM "
    strSql = strSql & gstrFaixaDeValor
    strSql = strSql & " ORDER BY strNomeDaFaixa "
strQuery = strSql
End Function






