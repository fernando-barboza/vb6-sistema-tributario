VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "TODG7.OCX"
Begin VB.Form frmCadDocumentos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tipos de Documento"
   ClientHeight    =   4020
   ClientLeft      =   2250
   ClientTop       =   1560
   ClientWidth     =   5865
   HelpContextID   =   153
   Icon            =   "frmCadDocumentos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4020
   ScaleWidth      =   5865
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtPKId 
      Height          =   285
      Left            =   5010
      TabIndex        =   6
      Top             =   0
      Visible         =   0   'False
      Width           =   645
   End
   Begin TabDlg.SSTab tab_3dPasta 
      Height          =   3810
      Left            =   120
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   120
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   6720
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Tipos de Documento"
      TabPicture(0)   =   "frmCadDocumentos.frx":1042
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblstrDescricao"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblintCodigo"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "tdb_Documentos"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txtstrDescricao"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtintCodigo"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      Begin VB.TextBox txtintCodigo 
         Height          =   285
         Left            =   1140
         MaxLength       =   5
         OLEDragMode     =   1  'Automatic
         TabIndex        =   0
         Top             =   480
         Width           =   1005
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
         Left            =   1140
         MaxLength       =   45
         TabIndex        =   1
         Top             =   840
         Width           =   4380
      End
      Begin TrueOleDBGrid70.TDBGrid tdb_Documentos 
         Height          =   2475
         Left            =   120
         TabIndex        =   5
         Top             =   1200
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   4366
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "PKID"
         Columns(0).DataField=   "PKID"
         Columns(0).NumberFormat=   "FormatText Event"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Código"
         Columns(1).DataField=   "intCodigo"
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
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0).AllowSizing=0"
         Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
         Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(8)=   "Column(1).Width=1561"
         Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=1482"
         Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(13)=   "Column(2).Width=7541"
         Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=7461"
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
         RowDividerStyle =   3
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
         _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38,.bgcolor=&H80000014&"
         _StyleDefs(20)  =   ":id=8,.fgcolor=&H80000012&"
         _StyleDefs(21)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
         _StyleDefs(22)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
         _StyleDefs(23)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(24)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
         _StyleDefs(25)  =   "Splits(0).Style:id=13,.parent=1"
         _StyleDefs(26)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
         _StyleDefs(27)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
         _StyleDefs(28)  =   "Splits(0).FooterStyle:id=15,.parent=3"
         _StyleDefs(29)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
         _StyleDefs(30)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
         _StyleDefs(31)  =   "Splits(0).EditorStyle:id=17,.parent=7"
         _StyleDefs(32)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
         _StyleDefs(33)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
         _StyleDefs(34)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
         _StyleDefs(35)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(36)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(37)  =   "Splits(0).Columns(0).Style:id=50,.parent=13"
         _StyleDefs(38)  =   "Splits(0).Columns(0).HeadingStyle:id=47,.parent=14"
         _StyleDefs(39)  =   "Splits(0).Columns(0).FooterStyle:id=48,.parent=15"
         _StyleDefs(40)  =   "Splits(0).Columns(0).EditorStyle:id=49,.parent=17"
         _StyleDefs(41)  =   "Splits(0).Columns(1).Style:id=46,.parent=13"
         _StyleDefs(42)  =   "Splits(0).Columns(1).HeadingStyle:id=43,.parent=14"
         _StyleDefs(43)  =   "Splits(0).Columns(1).FooterStyle:id=44,.parent=15"
         _StyleDefs(44)  =   "Splits(0).Columns(1).EditorStyle:id=45,.parent=17"
         _StyleDefs(45)  =   "Splits(0).Columns(2).Style:id=28,.parent=13"
         _StyleDefs(46)  =   "Splits(0).Columns(2).HeadingStyle:id=25,.parent=14"
         _StyleDefs(47)  =   "Splits(0).Columns(2).FooterStyle:id=26,.parent=15"
         _StyleDefs(48)  =   "Splits(0).Columns(2).EditorStyle:id=27,.parent=17"
         _StyleDefs(49)  =   "Named:id=33:Normal"
         _StyleDefs(50)  =   ":id=33,.parent=0"
         _StyleDefs(51)  =   "Named:id=34:Heading"
         _StyleDefs(52)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(53)  =   ":id=34,.wraptext=-1"
         _StyleDefs(54)  =   "Named:id=35:Footing"
         _StyleDefs(55)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(56)  =   "Named:id=36:Selected"
         _StyleDefs(57)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(58)  =   "Named:id=37:Caption"
         _StyleDefs(59)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(60)  =   "Named:id=38:HighlightRow"
         _StyleDefs(61)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(62)  =   "Named:id=39:EvenRow"
         _StyleDefs(63)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(64)  =   "Named:id=40:OddRow"
         _StyleDefs(65)  =   ":id=40,.parent=33"
         _StyleDefs(66)  =   "Named:id=41:RecordSelector"
         _StyleDefs(67)  =   ":id=41,.parent=34"
         _StyleDefs(68)  =   "Named:id=42:FilterBar"
         _StyleDefs(69)  =   ":id=42,.parent=33"
      End
      Begin VB.Label lblintCodigo 
         AutoSize        =   -1  'True
         Caption         =   "Código"
         Height          =   195
         Left            =   585
         TabIndex        =   4
         Top             =   585
         Width           =   495
      End
      Begin VB.Label lblstrDescricao 
         AutoSize        =   -1  'True
         Caption         =   "Descrição"
         Height          =   195
         Left            =   360
         TabIndex        =   3
         Top             =   915
         Width           =   720
      End
   End
End
Attribute VB_Name = "frmCadDocumentos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mblnAlterando    As Boolean
Dim mobjAux          As Object
Dim mblnSelecionou   As Boolean
Dim mblnPrimeiraVez  As Boolean
Dim mstrQueryAplicar As String
Dim blnAlterando     As Boolean
Dim bytOrdenacao     As Byte
Dim blnOrdenacaoAsc  As Boolean

Dim strCodigo        As String
Dim strCodigoAtual   As String
Dim strDescriAtual   As String

Private Function strQuery() As String
    Dim strSql  As String
    strSql = ""
    strSql = strSql & " SELECT PKId, intCodigo, strDescricao FROM "
    strSql = strSql & gstrDocumentos
    
    Select Case bytOrdenacao
      Case Is = 1
         strSql = strSql & " ORDER BY intCodigo" & IIf(blnOrdenacaoAsc, " ASC", " DESC")
         
      Case Is = 2
         strSql = strSql & " ORDER BY strDescricao" & IIf(blnOrdenacaoAsc, " ASC", " DESC")
    End Select
    
    strQuery = strSql

End Function

Private Sub Form_Activate()
    gintCodSeguranca = 6
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
    bytOrdenacao = 2: blnOrdenacaoAsc = True
    'VerificaListaAutomatica gstrFormaPagamento, tdb_Documentos, strQuery
    VerificaObjParaAplicar mobjAux, mstrQueryAplicar
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar
    mblnSelecionou = False
    mblnPrimeiraVez = False
    mblnAlterando = False
End Sub

Private Sub tdb_Documentos_Click()
    mblnPrimeiraVez = True
    If glngQtdLinhaTDBGrid(tdb_Documentos) = 1 Then
        tdb_Documentos_RowColChange 0, 0
    End If
End Sub

Private Sub tdb_Documentos_DblClick()
    MantemForm gstrAplicar
End Sub

Private Sub tdb_Documentos_FilterChange()
    mblnPrimeiraVez = False
    gblnFilraCampos tdb_Documentos
End Sub

Private Sub tdb_Documentos_GotFocus()
    tab_3dPasta.Tab = 0
End Sub

Private Sub tdb_Documentos_HeadClick(ByVal ColIndex As Integer)
    gOrdenaGrid tdb_Documentos, ColIndex
End Sub

Private Sub tdb_Documentos_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", tdb_Documentos
End Sub

Private Sub tdb_Documentos_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    
    With tdb_Documentos
        If Not .EOF And Not .BOF Then
            If mblnPrimeiraVez Then
                txtPKID.Text = .Columns("PKID").Value
                LeDaTabelaParaObj gstrDocumentos, Me
                gCorLinhaSelecionada tdb_Documentos

                HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrDeletar

                If mobjAux Is Nothing Then
                    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar
                Else
                    HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrAplicar
                End If
                mblnSelecionou = True
                mblnAlterando = True
                mblnPrimeiraVez = False
                
                strCodigoAtual = txtintCodigo.Text
                strDescriAtual = txtstrDescricao.Text
 
            End If
        End If
    End With

End Sub

Public Sub MantemForm(ByVal strModoOperacao As String)

If strModoOperacao = UCase(gstrSalvar) Then
    If blnDadosOk Then
        blnAlterando = mblnAlterando
        ToolBarGeral strModoOperacao, gstrDocumentos, _
                     mblnAlterando, tdb_Documentos, Me, mobjAux, strQuery, mstrQueryAplicar, rptDocumentos, strQueryRelatorio
        HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrDeletar
    End If
Else
    ToolBarGeral strModoOperacao, gstrDocumentos, _
                 mblnAlterando, tdb_Documentos, Me, mobjAux, strQuery, mstrQueryAplicar, rptDocumentos, strQueryRelatorio
    If Not gobjGeral Is Nothing Then HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrDeletar
End If

End Sub

Private Sub txtintCodigo_GotFocus()
    MarcaCampo txtintCodigo
    tab_3dPasta.Tab = 0
    gstrProximoCodigo txtintCodigo, gstrDocumentos, "intCodigo", gintCodSeguranca
End Sub

Private Sub txtintCodigo_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtintCodigo
End Sub

Private Sub txtintCodigo_KeyUp(KeyCode As Integer, Shift As Integer)
'    PesquisaListView KeyCode, txtintCodigo, tdb_Documentos, mblnAlterando
End Sub

Private Sub txtstrDescricao_GotFocus()
    MarcaCampo txtstrDescricao
    tab_3dPasta.Tab = 0
End Sub

Private Sub txtstrDescricao_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrDescricao
End Sub

Private Sub txtstrDescricao_KeyUp(KeyCode As Integer, Shift As Integer)
'    PesquisaListView KeyCode, txtstrDescricao, tdb_Documentos, mblnAlterando, lvwSubItem
End Sub

Private Function strQueryRelatorio() As String
Dim strSql As String
    
    strSql = ""
    strSql = strSql & " SELECT "
    strSql = strSql & " intCodigo, strDescricao "
    strSql = strSql & " FROM "
    strSql = strSql & gstrDocumentos
    If mblnSelecionou = True Then
        If txtPKID.Text <> "" Then
            strSql = strSql & " WHERE PKId = " & Val(txtPKID.Text)
        End If
    End If
    strSql = strSql & " ORDER BY "
    strSql = strSql & " strDescricao"
    
strQueryRelatorio = strSql
End Function
Private Function blnDadosOk() As Boolean
On Error GoTo err_blnDadosOK
blnDadosOk = False

    If Trim(txtintCodigo) = "" Then
        ExibeMensagem "O Código deve ser informado."
        txtintCodigo.SetFocus
        Exit Function
    End If
    
    If Trim(txtstrDescricao) = "" Then
        ExibeMensagem "A descrição deve ser informada."
        txtstrDescricao.SetFocus
        Exit Function
    End If
    
    If Not mblnAlterando Or (mblnAlterando And UCase$(strCodigoAtual) <> UCase$(txtintCodigo.Text)) Then

ProximoCodigo:

        If gblnExisteCodigo(1, gstrDocumentos, "intCodigo", "'" & txtintCodigo.Text & "'") Then
            strCodigo = (gstrProximoCodigo(txtintCodigo, gstrDocumentos, "intCodigo", gintCodSeguranca, , , , True))
            If MsgBox("O código informado já se encontra cadastrado. Deseja usar o código " & strCodigo & "?", vbYesNo + vbQuestion) = vbNo Then
                txtintCodigo.SetFocus
                Exit Function
            Else
                txtintCodigo.Text = strCodigo
                GoTo ProximoCodigo
            End If
        End If
    End If
    
    If Not mblnAlterando Or (mblnAlterando And UCase$(txtstrDescricao.Text) <> UCase$(strDescriAtual)) Then
            
        If gblnExisteCodigo(1, gstrDocumentos, "strDescricao", "'" & txtstrDescricao.Text & "'") Then
            ExibeMensagem "A descrição informada já se encontra cadastrada."
            txtstrDescricao.SetFocus
            Exit Function
        End If
    End If
    
    blnDadosOk = True
    
err_blnDadosOK:

End Function


