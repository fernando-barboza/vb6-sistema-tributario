VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "TODG7.OCX"
Begin VB.Form frmCadGrupoDeAtividade 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Grupos de Atividades"
   ClientHeight    =   5205
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6435
   Icon            =   "CadGrupoDeAtividade.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5205
   ScaleWidth      =   6435
   Begin TabDlg.SSTab tab_3dPasta 
      Height          =   5085
      Left            =   90
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   90
      Width           =   6285
      _ExtentX        =   11086
      _ExtentY        =   8969
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Grupo"
      TabPicture(0)   =   "CadGrupoDeAtividade.frx":1042
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblintUtilizacaoDaTabelaDeValor"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblintCodigo"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblstrNomeDoGrupo"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "dbcintUtilizacaoDaTabelaDeValor"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "tdb_GrupoAtividade"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtintCodigo"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtstrNomeDoGrupo"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      Begin VB.TextBox txtstrNomeDoGrupo 
         Height          =   285
         Left            =   1050
         MaxLength       =   30
         TabIndex        =   4
         Top             =   1260
         Width           =   5055
      End
      Begin VB.TextBox txtintCodigo 
         Height          =   285
         Left            =   1050
         MaxLength       =   8
         TabIndex        =   2
         Top             =   900
         Width           =   1305
      End
      Begin TrueOleDBGrid70.TDBGrid tdb_GrupoAtividade 
         Height          =   3075
         Left            =   180
         TabIndex        =   3
         Top             =   1710
         Width           =   5925
         _ExtentX        =   10451
         _ExtentY        =   5424
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
         Columns(1).Caption=   "C�digo"
         Columns(1).DataField=   "intCodigo"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Descri��o"
         Columns(2).DataField=   "strNomeDoGrupo"
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
         Splits(0)._ColumnProps(1)=   "Column(0).Width=1376"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1296"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0).AllowSizing=0"
         Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
         Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(8)=   "Column(1).Width=1667"
         Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=1588"
         Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(13)=   "Column(2).Width=8176"
         Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=8096"
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
         _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34"
         _StyleDefs(11)  =   "FooterStyle:id=3,.parent=1,.namedParent=35"
         _StyleDefs(12)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(13)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
         _StyleDefs(14)  =   "EditorStyle:id=7,.parent=1"
         _StyleDefs(15)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
         _StyleDefs(16)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
         _StyleDefs(17)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
         _StyleDefs(18)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(19)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
         _StyleDefs(20)  =   "Splits(0).Style:id=13,.parent=1"
         _StyleDefs(21)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
         _StyleDefs(22)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
         _StyleDefs(23)  =   "Splits(0).FooterStyle:id=15,.parent=3"
         _StyleDefs(24)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
         _StyleDefs(25)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
         _StyleDefs(26)  =   "Splits(0).EditorStyle:id=17,.parent=7"
         _StyleDefs(27)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
         _StyleDefs(28)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
         _StyleDefs(29)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
         _StyleDefs(30)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(31)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(32)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
         _StyleDefs(33)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
         _StyleDefs(34)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
         _StyleDefs(35)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
         _StyleDefs(36)  =   "Splits(0).Columns(1).Style:id=46,.parent=13"
         _StyleDefs(37)  =   "Splits(0).Columns(1).HeadingStyle:id=43,.parent=14"
         _StyleDefs(38)  =   "Splits(0).Columns(1).FooterStyle:id=44,.parent=15"
         _StyleDefs(39)  =   "Splits(0).Columns(1).EditorStyle:id=45,.parent=17"
         _StyleDefs(40)  =   "Splits(0).Columns(2).Style:id=32,.parent=13"
         _StyleDefs(41)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
         _StyleDefs(44)  =   "Named:id=33:Normal"
         _StyleDefs(45)  =   ":id=33,.parent=0"
         _StyleDefs(46)  =   "Named:id=34:Heading"
         _StyleDefs(47)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(48)  =   ":id=34,.wraptext=-1"
         _StyleDefs(49)  =   "Named:id=35:Footing"
         _StyleDefs(50)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(51)  =   "Named:id=36:Selected"
         _StyleDefs(52)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(53)  =   "Named:id=37:Caption"
         _StyleDefs(54)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(55)  =   "Named:id=38:HighlightRow"
         _StyleDefs(56)  =   ":id=38,.parent=33,.bgcolor=&H80000014&,.fgcolor=&H80000012&"
         _StyleDefs(57)  =   "Named:id=39:EvenRow"
         _StyleDefs(58)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(59)  =   "Named:id=40:OddRow"
         _StyleDefs(60)  =   ":id=40,.parent=33"
         _StyleDefs(61)  =   "Named:id=41:RecordSelector"
         _StyleDefs(62)  =   ":id=41,.parent=34"
         _StyleDefs(63)  =   "Named:id=42:FilterBar"
         _StyleDefs(64)  =   ":id=42,.parent=33"
      End
      Begin MSDataListLib.DataCombo dbcintUtilizacaoDaTabelaDeValor 
         Height          =   315
         Left            =   1050
         TabIndex        =   5
         Top             =   510
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   556
         _Version        =   393216
         IntegralHeight  =   0   'False
         Style           =   2
         Text            =   ""
      End
      Begin VB.Label lblstrNomeDoGrupo 
         AutoSize        =   -1  'True
         Caption         =   "Descri��o"
         Height          =   195
         Left            =   210
         TabIndex        =   8
         Top             =   1305
         Width           =   720
      End
      Begin VB.Label lblintCodigo 
         AutoSize        =   -1  'True
         Caption         =   "C�digo"
         Height          =   195
         Left            =   435
         TabIndex        =   7
         Top             =   945
         Width           =   495
      End
      Begin VB.Label lblintUtilizacaoDaTabelaDeValor 
         AutoSize        =   -1  'True
         Caption         =   "Utiliza��o"
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   570
         Width           =   690
      End
   End
   Begin VB.TextBox txtPKId 
      Enabled         =   0   'False
      Height          =   270
      Left            =   1440
      TabIndex        =   0
      TabStop         =   0   'False
      Text            =   "txtPKId"
      Top             =   -60
      Visible         =   0   'False
      Width           =   645
   End
End
Attribute VB_Name = "frmCadGrupoDeAtividade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim mblnAlterando      As Boolean
    Dim mobjAux            As Object
    Dim mblnSelecionou     As Boolean
    Dim mblnPrimeiraVez    As Boolean
    Dim strDuplicataCodigo As String

Private Sub dbcintUtilizacaoDaTabelaDeValor_Click(Area As Integer)
    DropDownDataCombo dbcintUtilizacaoDaTabelaDeValor, Me, Area
    If Area = 2 And dbcintUtilizacaoDaTabelaDeValor.MatchedWithList Then
        VerificaListaAutomatica gstrGrupoDeAtividade, tdb_GrupoAtividade, strQuery
        mblnPrimeiraVez = False
        MantemForm gstrNovo
    End If
End Sub

Private Sub dbcintUtilizacaoDaTabelaDeValor_KeyDown(KeyCode As Integer, Shift As Integer)
   DropDownDataCombo dbcintUtilizacaoDaTabelaDeValor, Me, , KeyCode, Shift
End Sub

Private Sub Form_Activate()
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
    mblnAlterando = False
    dbcintUtilizacaoDaTabelaDeValor.Tag = strQuery & ";strdescricao"
    LeDaTabelaParaObj gstrUtilizacaoDaTabelaDeValor, dbcintUtilizacaoDaTabelaDeValor, "PKId , strNomeDaUtilizacao ", "WHERE bitIdentificador = 0"
    VerificaObjParaAplicar mobjAux
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar, gstrDeletar
    mblnSelecionou = False
    mblnPrimeiraVez = False
End Sub

Private Sub tdb_GrupoAtividade_Click()
    mblnPrimeiraVez = True
End Sub

Private Sub tdb_GrupoAtividade_FilterChange()
    gblnFilraCampos tdb_GrupoAtividade
End Sub

Private Function strQuery() As String

'******************************************************************************************
' Data: 05/05/2003
' Altera��o: - Retirado comando "AS" utilizado para dar apelidos �s tabelas, o qual n�o �
'            permitido pelo Oracle, da cl�usula FROM.
' Respons�vel: Everton Bianchini
'******************************************************************************************

    Dim strSql  As String
    strSql = ""
    strSql = strSql & " SELECT A.PKId, A.intCodigo, A.strNomeDoGrupo FROM "
'    strSQL = strSQL & gstrGrupoDeAtividade & " AS A, "
    strSql = strSql & gstrGrupoDeAtividade & " A, "
'    strSQL = strSQL & gstrUtilizacaoDaTabelaDeValor & " AS B "
    strSql = strSql & gstrUtilizacaoDaTabelaDeValor & " B "
    strSql = strSql & " WHERE B.PKId = A.intUtilizacaoDaTabelaDeValor "
    If dbcintUtilizacaoDaTabelaDeValor.MatchedWithList Then
        strSql = strSql & " AND B.PKId = " & dbcintUtilizacaoDaTabelaDeValor.BoundText
    End If
    strSql = strSql & "ORDER BY A.intCodigo "
    strQuery = strSql
End Function

Private Function strQueryAplicar() As String
    Dim strSql  As String
    strSql = ""
    strSql = strSql & " SELECT PKId, strNomeDoGrupo FROM "
    strSql = strSql & gstrGrupoDeAtividade & " ORDER BY PKID"
    strQueryAplicar = strSql
End Function

Private Sub tdb_GrupoAtividade_KeyPress(KeyAscii As Integer)
    Select Case tdb_GrupoAtividade.Col
        Case 0
            CaracterValido KeyAscii, "N", tdb_GrupoAtividade
        Case Else
            CaracterValido KeyAscii, "A", tdb_GrupoAtividade
    End Select
End Sub

Private Sub tdb_GrupoAtividade_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    With tdb_GrupoAtividade
        If Not .EOF And Not .BOF Then
            If mblnPrimeiraVez Then
                mblnAlterando = True
                txtPKId.Text = .Columns("PKID").Value
                LeDaTabelaParaObj gstrGrupoDeAtividade, Me
                gCorLinhaSelecionada tdb_GrupoAtividade
                
                HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrDeletar
                
                If mobjAux Is Nothing Then
                    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar
                Else
                    HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrAplicar
                End If
                mblnSelecionou = True
            End If
        End If
    End With
End Sub

Public Sub MantemForm(ByVal strModoOperacao As String)
    Dim varBookMark As Variant
    Dim strSql As String
    
    
    strSql = strQuery
    
    If ToolBarGeral(strModoOperacao, gstrGrupoDeAtividade, mblnAlterando, tdb_GrupoAtividade, Me, mobjAux, strSql, strQueryAplicar) Then
        If UCase(strModoOperacao) = gstrSalvar Or UCase(strModoOperacao) = gstrDeletar Then
            mblnPrimeiraVez = False
        End If
    End If
    
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrDeletar
    
    If UCase(strModoOperacao) = gstrSalvar Or UCase(strModoOperacao) = gstrDeletar Then
        mblnPrimeiraVez = False
    ElseIf UCase(strModoOperacao) = gstrNovo Then
        If txtintCodigo.Enabled Then
            txtintCodigo.SetFocus
        End If
    End If
End Sub

Function blnTemDuplicata() As Boolean
Dim strSql             As String
Dim adoResultado       As adodb.Recordset
blnTemDuplicata = False
strDuplicataCodigo = ""
    If Val(txtintCodigo) = 0 Then Exit Function
    strSql = ""
    strSql = strSql & "SELECT * FROM "
    strSql = strSql & gstrGrupoDeAtividade
    strSql = strSql & " WHERE intCodigo = " & txtintCodigo
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        With adoResultado
            If Not .EOF Then
                strDuplicataCodigo = strSql
                blnTemDuplicata = True
                txtPKId.Text = !Pkid
                Exit Function
            End If
        End With
    End If
End Function

Private Sub txtintCodigo_GotFocus()
    MarcaCampo txtintCodigo
End Sub

Private Sub txtintCodigo_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtintCodigo
End Sub

Private Sub txtstrNomeDoGrupo_GotFocus()
    MarcaCampo txtstrNomeDoGrupo
End Sub

Private Sub txtstrNomeDoGrupo_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrNomeDoGrupo
End Sub

