VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "TODG7.OCX"
Begin VB.Form frmCadTiposDeTestada 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tipos de Testada"
   ClientHeight    =   5130
   ClientLeft      =   1125
   ClientTop       =   4485
   ClientWidth     =   6420
   HelpContextID   =   38
   Icon            =   "CadTiposDeTestada.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5130
   ScaleWidth      =   6420
   Begin VB.TextBox txtPKId 
      Height          =   270
      Left            =   2550
      TabIndex        =   5
      Top             =   60
      Visible         =   0   'False
      Width           =   645
   End
   Begin TabDlg.SSTab tab_3dPasta 
      Height          =   4905
      Left            =   120
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   120
      Width           =   6195
      _ExtentX        =   10927
      _ExtentY        =   8652
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   529
      TabCaption(0)   =   "Tipos de Testada"
      TabPicture(0)   =   "CadTiposDeTestada.frx":1042
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblintCodigoDaTestada"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblstrNomeDaTestada"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "chkbytPassivaDeCM"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txtintCodigoDaTestada"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtstrNomeDaTestada"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "tdb_TiposDeTestada"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "chkBytPrincipal"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      Begin VB.CheckBox chkBytPrincipal 
         Caption         =   "Principal"
         Height          =   195
         Left            =   3270
         TabIndex        =   8
         Top             =   1110
         Width           =   915
      End
      Begin TrueOleDBGrid70.TDBGrid tdb_TiposDeTestada 
         Height          =   3375
         Left            =   120
         TabIndex        =   4
         Top             =   1380
         Width           =   5925
         _ExtentX        =   10451
         _ExtentY        =   5953
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "PKId"
         Columns(0).DataField=   "PKId"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Código"
         Columns(1).DataField=   "intCodigoDaTestada"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Nome"
         Columns(2).DataField=   "strNomeDaTestada"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "bytPrincipal"
         Columns(3).DataField=   "bytPrincipal"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   4
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
         Splits(0)._ColumnProps(0)=   "Columns.Count=4"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0).AllowSizing=0"
         Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
         Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(8)=   "Column(1).Width=2170"
         Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=2090"
         Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(13)=   "Column(2).Width=7752"
         Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=7673"
         Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(17)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(18)=   "Column(3).Width=2725"
         Splits(0)._ColumnProps(19)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(20)=   "Column(3)._WidthInPix=2646"
         Splits(0)._ColumnProps(21)=   "Column(3)._EditAlways=0"
         Splits(0)._ColumnProps(22)=   "Column(3).Order=4"
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
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=96,.bold=0,.fontsize=825,.italic=0"
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
         _StyleDefs(38)  =   "Splits(0).Columns(2).Style:id=46,.parent=13"
         _StyleDefs(39)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
         _StyleDefs(40)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
         _StyleDefs(41)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
         _StyleDefs(42)  =   "Splits(0).Columns(3).Style:id=50,.parent=13"
         _StyleDefs(43)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
         _StyleDefs(44)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
         _StyleDefs(45)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
         _StyleDefs(46)  =   "Named:id=33:Normal"
         _StyleDefs(47)  =   ":id=33,.parent=0"
         _StyleDefs(48)  =   "Named:id=34:Heading"
         _StyleDefs(49)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(50)  =   ":id=34,.wraptext=-1"
         _StyleDefs(51)  =   "Named:id=35:Footing"
         _StyleDefs(52)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(53)  =   "Named:id=36:Selected"
         _StyleDefs(54)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(55)  =   "Named:id=37:Caption"
         _StyleDefs(56)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(57)  =   "Named:id=38:HighlightRow"
         _StyleDefs(58)  =   ":id=38,.parent=33,.bgcolor=&H80000014&,.fgcolor=&H80000012&"
         _StyleDefs(59)  =   "Named:id=39:EvenRow"
         _StyleDefs(60)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(61)  =   "Named:id=40:OddRow"
         _StyleDefs(62)  =   ":id=40,.parent=33"
         _StyleDefs(63)  =   "Named:id=41:RecordSelector"
         _StyleDefs(64)  =   ":id=41,.parent=34"
         _StyleDefs(65)  =   "Named:id=42:FilterBar"
         _StyleDefs(66)  =   ":id=42,.parent=33"
      End
      Begin VB.TextBox txtstrNomeDaTestada 
         Height          =   285
         Left            =   1020
         MaxLength       =   30
         TabIndex        =   1
         Top             =   750
         Width           =   5025
      End
      Begin VB.TextBox txtintCodigoDaTestada 
         Height          =   285
         Left            =   1020
         MaxLength       =   9
         TabIndex        =   0
         Top             =   390
         Width           =   1035
      End
      Begin VB.CheckBox chkbytPassivaDeCM 
         Caption         =   "Contribuição de Melhoria"
         Height          =   195
         Left            =   1020
         TabIndex        =   3
         Top             =   1110
         Width           =   2115
      End
      Begin VB.Label lblstrNomeDaTestada 
         AutoSize        =   -1  'True
         Caption         =   "Descrição"
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   825
         Width           =   720
      End
      Begin VB.Label lblintCodigoDaTestada 
         AutoSize        =   -1  'True
         Caption         =   "Código"
         Height          =   195
         Left            =   465
         TabIndex        =   6
         Top             =   465
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmCadTiposDeTestada"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim mblnAlterando       As Boolean
    Dim mobjAux             As Object
    Dim mblnSelecionou      As Boolean
    Dim mblnPrimeiraVez     As Boolean
    Dim chkPrincipalAtual   As Byte
    Dim blnOrdenacaoAsc     As Boolean
    Dim bytOrdenacao        As Byte

Private Sub chkbytPassivaDeCM_KeyPress(KeyAscii As Integer)
 CaracterValido KeyAscii, "A", chkbytPassivaDeCM
End Sub

Private Sub Form_Activate()
    gintCodSeguranca = 610
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

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
  If KeyCode = vbKeyF1 Then
     Call_HtmlHelp Me.HelpContextID
  End If
End Sub

Private Sub Form_Load()
    bytOrdenacao = 2: blnOrdenacaoAsc = True
    VerificaListaAutomatica gstrTipoDeTestada, tdb_TiposDeTestada, "PKId, intCodigoDaTestada, strNomeDaTestada"
    VerificaObjParaAplicar mobjAux
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar, gstrDeletar
mblnSelecionou = False
mblnPrimeiraVez = False
End Sub

Private Function strQuery() As String
    Dim strsql  As String
    strsql = ""
    strsql = strsql & " SELECT PKId, intCodigoDaTestada, strNomeDaTestada, bytPrincipal FROM "
    strsql = strsql & gstrTipoDeTestada
    
    Select Case bytOrdenacao
      Case Is = 1
         strsql = strsql & " ORDER BY intCodigoDaTestada" & IIf(blnOrdenacaoAsc, " ASC", " DESC")
      Case Is = 2
         strsql = strsql & " ORDER BY strNomeDaTestada" & IIf(blnOrdenacaoAsc, " ASC", " DESC")
    End Select
    
    strQuery = strsql
End Function


Private Sub tdb_TiposDeTestada_Click()
    mblnPrimeiraVez = True
    If glngQtdLinhaTDBGrid(tdb_TiposDeTestada) = 1 Then
        tdb_TiposdeTestada_RowColChange 0, 0
    End If
End Sub

Private Sub tdb_TiposdeTestada_FilterChange()
    mblnPrimeiraVez = False
    gblnFilraCampos tdb_TiposDeTestada
End Sub

Private Sub tdb_TiposDeTestada_HeadClick(ByVal ColIndex As Integer)
    gOrdenaGrid tdb_TiposDeTestada, ColIndex
End Sub

Private Sub tdb_TiposDeTestada_KeyPress(KeyAscii As Integer)
CaracterValido KeyAscii, "A", tdb_TiposDeTestada
End Sub

Private Sub tdb_TiposdeTestada_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    
    With tdb_TiposDeTestada
        If Not .EOF And Not .BOF Then
            txtPKID.Text = .Columns("PKID").Value
            If mblnPrimeiraVez Then
                mblnAlterando = True
                LeDaTabelaParaObj gstrTipoDeTestada, Me
'=============
                gCorLinhaSelecionada tdb_TiposDeTestada
                
                HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrDeletar
                
                If mobjAux Is Nothing Then
                    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar
                Else

                    HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrAplicar
                End If
'=============
                mblnSelecionou = True
                chkPrincipalAtual = chkBytPrincipal.Value
            End If
        End If
    End With

End Sub

Public Sub MantemForm(ByVal strModoOperacao As String)
Dim varBookMark As Variant
Dim strsql As String

  strsql = strQuery
  If strModoOperacao = UCase("IMPRIMIR") Then
      ToolBarGeral strModoOperacao, gstrTipoDeTestada, mblnAlterando, tdb_TiposDeTestada, Me, mobjAux, strsql, , rptTipoDeTestada, strQuerryRelatorio
      Exit Sub
  End If
  If UCase(strModoOperacao) = "SALVAR" Then
     mblnPrimeiraVez = False
     If blnDadosOk = False Then
        Exit Sub
     End If
     If Not mblnAlterando And chkBytPrincipal.Value Or (mblnAlterando And chkPrincipalAtual = 0 And chkBytPrincipal.Value <> chkPrincipalAtual) Then
        If gblnExisteCodigo(1, gstrTipoDeTestada, "bytPrincipal", "'1'") Then
           ExibeMensagem "Já existe uma Testada Principal."
           Exit Sub
        End If
     End If
  End If
   
   
  If UCase(strModoOperacao) = "DELETAR" Then
     mblnPrimeiraVez = False
  End If
  
  If UCase(strModoOperacao) = "LOCALIZAR" Then
     If Trim(txtintCodigoDaTestada.Text) = "" And Trim(txtstrNomeDaTestada.Text) = "" Then
        txtPKID.Text = ""
     End If
  End If
  
  
  ToolBarGeral strModoOperacao, gstrTipoDeTestada, mblnAlterando, tdb_TiposDeTestada, Me, mobjAux, strsql, strsql
  
  If (UCase(strModoOperacao) = "SALVAR" Or UCase(strModoOperacao) = "DELETAR") And _
      gblnCancelarInclusao = False Then
      LeDaTabelaParaObj gstrTipoDeTestada, tdb_TiposDeTestada, strsql
      HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrDeletar
  End If
  
  If UCase(strModoOperacao) = "NOVO" Then
     HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrDeletar
  End If
  
End Sub

Private Sub txtintCodigoDaTestada_Change()
  If Trim(txtintCodigoDaTestada.Text) = "" And _
     Trim(txtstrNomeDaTestada.Text) = "" Then
     txtPKID.Text = ""
  End If
End Sub

Private Sub txtintCodigoDaTestada_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtintCodigoDaTestada
End Sub

Private Sub txtstrNomeDaTestada_Change()
  If Trim(txtintCodigoDaTestada.Text) = "" And _
     Trim(txtstrNomeDaTestada.Text) = "" Then
     txtPKID.Text = ""
  End If
End Sub

Private Sub txtstrNomeDaTestada_GotFocus()
    MarcaCampo txtstrNomeDaTestada
End Sub

Private Sub txtintCodigoDaTestada_GotFocus()
    MarcaCampo txtintCodigoDaTestada
End Sub

Private Sub txtstrNomeDaTestada_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrNomeDaTestada
End Sub

Function strQuerryRelatorio() As String
Dim strsql As String
strsql = ""
strsql = strsql & " SELECT * FROM " & gstrTipoDeTestada
If mblnAlterando = True Then
    strsql = strsql & " WHERE PKId = " & Val(txtPKID)
End If
strsql = strsql & " ORDER BY strNomeDaTestada "
strQuerryRelatorio = strsql
End Function

Private Function blnDadosOk() As Boolean

   If Trim(txtintCodigoDaTestada.Text) = "" Then
      ExibeMensagem "O campo Código deve ser preenchido."
      txtintCodigoDaTestada.SetFocus
      Exit Function
   End If
   
   If Trim(txtstrNomeDaTestada.Text) = "" Then
      ExibeMensagem "O campo Descrição deve ser preenchido."
      txtstrNomeDaTestada.SetFocus
      Exit Function
   End If
   
   If mblnAlterando = True And (Val(txtintCodigoDaTestada.Text) = tdb_TiposDeTestada.Columns("Código") And _
      Trim(txtstrNomeDaTestada.Text) = tdb_TiposDeTestada.Columns("Nome")) Then
      blnDadosOk = True
      Exit Function
   End If
   
   If mblnAlterando = True Then
      If Val(txtintCodigoDaTestada.Text) <> tdb_TiposDeTestada.Columns("Código") Then
         If gblnExisteCodigo(1, gstrTipoDeTestada, "intCodigoDaTestada", Val(txtintCodigoDaTestada.Text)) Then
            ExibeMensagem "Este Código já está cadastrado."
            txtintCodigoDaTestada.SetFocus
            Exit Function
         End If
      End If
      If Trim(txtstrNomeDaTestada.Text) <> tdb_TiposDeTestada.Columns("Nome") Then
         If gblnExisteCodigo(1, gstrTipoDeTestada, "strNomeDaTestada", Trim(txtstrNomeDaTestada.Text)) Then
            ExibeMensagem "Esta Descrição já está cadastrada."
            txtstrNomeDaTestada.SetFocus
            Exit Function
         End If
      End If

   Else
      If gblnExisteCodigo(1, gstrTipoDeTestada, "intCodigoDaTestada", Val(txtintCodigoDaTestada.Text)) Then
         ExibeMensagem "Este Código já está cadastrado."
         txtintCodigoDaTestada.SetFocus
         Exit Function
      End If
      If gblnExisteCodigo(1, gstrTipoDeTestada, "strNomeDaTestada", Trim(txtstrNomeDaTestada.Text)) Then
         ExibeMensagem "Esta Descrição já está cadastrada."
         txtstrNomeDaTestada.SetFocus
         Exit Function
      End If
   End If
   
   blnDadosOk = True
End Function

