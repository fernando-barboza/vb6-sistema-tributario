VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frmCadPrograma 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Programas do Governo"
   ClientHeight    =   3525
   ClientLeft      =   1920
   ClientTop       =   2850
   ClientWidth     =   6390
   HelpContextID   =   16
   Icon            =   "CadPrograma.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3525
   ScaleWidth      =   6390
   Begin VB.TextBox txtPKId 
      Height          =   315
      Left            =   5790
      TabIndex        =   6
      Top             =   0
      Visible         =   0   'False
      Width           =   510
   End
   Begin TabDlg.SSTab tab_3DPasta 
      Height          =   3345
      Left            =   120
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   90
      Width           =   6165
      _ExtentX        =   10874
      _ExtentY        =   5900
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "Programas do Governo"
      TabPicture(0)   =   "CadPrograma.frx":1042
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblstrDescricao"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblstrCodigo"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "tdb_Lista"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txtstrDescricao"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtstrCodigo"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtintExercicio"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      Begin VB.TextBox txtintExercicio 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2250
         TabIndex        =   7
         Top             =   450
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.TextBox txtstrCodigo 
         Height          =   285
         Left            =   930
         MaxLength       =   4
         OLEDragMode     =   1  'Automatic
         TabIndex        =   0
         Top             =   450
         Width           =   1245
      End
      Begin VB.TextBox txtstrDescricao 
         Height          =   285
         Left            =   930
         MaxLength       =   100
         TabIndex        =   1
         Top             =   810
         Width           =   5115
      End
      Begin TrueOleDBGrid70.TDBGrid tdb_Lista 
         Height          =   2025
         Left            =   90
         TabIndex        =   2
         Top             =   1200
         Width           =   5955
         _ExtentX        =   10504
         _ExtentY        =   3572
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
         Splits(0)._ColumnProps(8)=   "Column(1).Width=1879"
         Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=1799"
         Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(13)=   "Column(2).Width=8096"
         Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=8017"
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
      Begin VB.Label lblstrCodigo 
         AutoSize        =   -1  'True
         Caption         =   "C�digo"
         Height          =   195
         Left            =   375
         TabIndex        =   5
         Top             =   555
         Width           =   495
      End
      Begin VB.Label lblstrDescricao 
         AutoSize        =   -1  'True
         Caption         =   "Descri��o"
         Height          =   195
         Left            =   150
         TabIndex        =   4
         Top             =   915
         Width           =   720
      End
   End
End
Attribute VB_Name = "frmCadPrograma"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim mblnAlterando       As Boolean
    Dim mblnClickOk         As Boolean
    Dim mobjAux             As Object
    Dim mblnselecionou      As Boolean
    'Dim intFiltroExercicio  As Integer
    Public mIntCodSeguranca As Integer

Private Function strQuery() As String
    
Dim strSQL  As String

    strSQL = "SELECT PKId, strCodigo, strDescricao FROM " & gstrPrograma & " WHERE intExercicio = " & txtintExercicio
    strSQL = strSQL & " ORDER BY " & gstrCONVERT(cdt_numeric, "strCodigo")
    
    strQuery = strSQL
    
End Function

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

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
  If KeyCode = vbKeyF1 Then
     Call_HtmlHelp Me.HelpContextID
  End If
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

Private Sub tdb_Lista_KeyPress(KeyAscii As Integer)
    
    If tdb_Lista.Col = 1 Then
        CaracterValido KeyAscii, "N", tdb_Lista
    End If
    
End Sub

Private Sub tdb_Lista_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    mblnClickOk = True
End Sub

Private Sub tdb_Lista_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    
    With tdb_Lista
        If (Not .EOF And Not .BOF) And mblnClickOk Then
            mblnClickOk = False
            txtPKId.Text = .Columns("PKID").Value
            LeDaTabelaParaObj gstrPrograma, Me
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
    
TrocaCorDeFundoObjeto False, txtstrCodigo
    
End Sub

Public Sub MantemForm(ByVal strModoOperacao As String)
    'Vamos verificar qual menu que chamou o form, para definirmos o filtro
    If gbytMenu = gbytMenuCadastro Then
        txtintExercicio = gintExercicio
    Else
        txtintExercicio = gintExercicio + 1
    End If
    Select Case strModoOperacao
    
    Case Is = gstrNovo
        
        TrocaCorObjeto txtstrCodigo, False, False
        
        LimpaCampos
                
        Exit Sub
    
    Case Is = gstrSalvar
        
        If Not blnDadosOk Then Exit Sub
        
        If Not mblnAlterando Then
            'If gblnExisteCodigo(1, gstrPrograma, "strDescricao", "'" & txtstrDescricao & "'") Or gblnExisteCodigo(2, gstrPrograma, "strCodigo", "'" & txtstrCodigo & "'", "intExercicio", txtintExercicio) Then
            If ConsultaCodigo(txtstrCodigo) > 0 Then
                MsgBox "Este programa j� se encontra cadastrado.", vbOKOnly, "Mensagem ao Usu�rio"
                Exit Sub
            End If
        End If
        
    End Select
    
    ToolBarGeral strModoOperacao, gstrPrograma, mblnAlterando, tdb_Lista, _
                 Me, mobjAux, strQuery, strQueryAplicar, rptPrograma, strQueryRelatorio
                 
    'If strModoOperacao = gstrNovo Or strModoOperacao = gstrDeletar Then txtintExercicio = intFiltroExercicio
    
End Sub

Private Sub txtstrCodigo_GotFocus()
    MarcaCampo txtstrCodigo
End Sub

Private Sub txtstrCodigo_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N"
End Sub

Private Sub txtstrCodigo_LostFocus()
    txtstrCodigo.Text = Format(txtstrCodigo, "0000")
End Sub

Private Sub txtstrDescricao_GotFocus()
    MarcaCampo txtstrDescricao
End Sub

Private Sub txtstrDescricao_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub Form_Load()
       
    'Vamos verificar qual menu que chamou o form, para definirmos o filtro
    If gbytMenu = gbytMenuCadastro Then
        txtintExercicio = gintExercicio
    Else
        txtintExercicio = gintExercicio + 1
    End If
    
    mblnAlterando = False
    LeDaTabelaParaObj gstrPrograma, tdb_Lista, strQuery
    VerificaObjParaAplicar mobjAux
    
End Sub

Private Sub LimpaCampos()
   
   txtstrCodigo = Space$(0)
   
'   gstrProximoCodigo txtstrCodigo, gstrPrograma, "strCodigo", gintCodSeguranca, , , , , , , "intExercicio", CStr(intFiltroExercicio)
    gstrProximoCodigo txtstrCodigo, gstrPrograma, "strCodigo", 186, , , , , , , "intExercicio", CStr(txtintExercicio)
   txtstrCodigo.Text = Format(txtstrCodigo, "0000")
   txtstrDescricao.Tag = Space$(0)
   txtstrDescricao = Space$(0)
   txtPKId = Space$(0)

   mblnAlterando = False
   txtstrCodigo.SetFocus
   
End Sub

Function strQueryRelatorio() As String

Dim strSQL As String
    
    strSQL = "SELECT strCodigo, strDescricao "
    strSQL = strSQL & "FROM " & gstrPrograma
    strSQL = strSQL & " WHERE intExercicio = " & txtintExercicio
    strSQL = strSQL & " ORDER BY " & gstrCONVERT(CDT_INT, "strCodigo")
    
    strQueryRelatorio = strSQL
    
End Function

Private Function blnDadosOk() As Boolean
On Error GoTo err_blnDadosOK

    blnDadosOk = False

    If Trim(txtstrCodigo) = "" Then
        ExibeMensagem "O c�digo tem que ser digitado."
        txtstrCodigo.SetFocus
        Exit Function
    End If
    
    If Trim(txtstrDescricao) = "" Then
        ExibeMensagem "A descri��o tem que ser digitada."
        txtstrDescricao.SetFocus
        Exit Function
    End If
    
    blnDadosOk = True
    
err_blnDadosOK:

End Function

Private Function strQueryAplicar() As String
    strQueryAplicar = "SELECT PKId, strDescricao FROM " & gstrPrograma & " WHERE intExercicio = " & txtintExercicio
End Function


Private Function ConsultaCodigo(strCodigo As String) As String
 Dim strSQL As String
 Dim adoResultado As New ADODB.Recordset
    
    strSQL = " SELECT  PKId, "
    strSQL = strSQL & "strCodigo, "
    strSQL = strSQL & "strDescricao "
    strSQL = strSQL & " From "
    strSQL = strSQL & gstrPrograma
    strSQL = strSQL & " Where "
    strSQL = strSQL & "intExercicio=" & txtintExercicio
    strSQL = strSQL & " and  " & gstrCONVERT(cdt_numeric, "StrCodigo") & "=" & strCodigo
    strSQL = strSQL & " ORDER BY" & gstrCONVERT(cdt_numeric, "strCodigo")

    Set gobjBanco = New clsBanco
    
    
    If gobjBanco.CriaADO(strSQL, 0, adoResultado) Then
        If adoResultado.RecordCount > 0 Then
            ConsultaCodigo = adoResultado.RecordCount
        Else
            ConsultaCodigo = 0
        End If
    End If

End Function

