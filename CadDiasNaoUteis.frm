VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "TODG7.OCX"
Begin VB.Form frmCadDiasNaoUteis 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Dias não Úteis"
   ClientHeight    =   5355
   ClientLeft      =   3090
   ClientTop       =   3270
   ClientWidth     =   6090
   HelpContextID   =   43
   Icon            =   "CadDiasNaoUteis.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5355
   ScaleWidth      =   6090
   Begin VB.TextBox txtPKId 
      Height          =   270
      Left            =   2640
      TabIndex        =   12
      Top             =   60
      Visible         =   0   'False
      Width           =   645
   End
   Begin TabDlg.SSTab tab_3dPasta 
      Height          =   5130
      Left            =   120
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   5880
      _ExtentX        =   10372
      _ExtentY        =   9049
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Dias Não Úteis"
      TabPicture(0)   =   "CadDiasNaoUteis.frx":1042
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblstrDescricao"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblintExercicio"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lbldtmData"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "tdb_Dias"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtstrDescricao"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "fra_Tipo"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "dtpdtmData"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cbointExercicio"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cmd_Exercicio"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).ControlCount=   9
      Begin VB.CommandButton cmd_Exercicio 
         Height          =   315
         Left            =   1950
         Picture         =   "CadDiasNaoUteis.frx":10A0
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Incluir novo exercício"
         Top             =   420
         Width           =   315
      End
      Begin VB.ComboBox cbointExercicio 
         Height          =   315
         ItemData        =   "CadDiasNaoUteis.frx":11EA
         Left            =   1050
         List            =   "CadDiasNaoUteis.frx":11EC
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   420
         Width           =   885
      End
      Begin MSComCtl2.DTPicker dtpdtmData 
         Height          =   285
         Left            =   1050
         TabIndex        =   7
         Top             =   1170
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         Format          =   24969217
         CurrentDate     =   37987
      End
      Begin VB.Frame fra_Tipo 
         Caption         =   " Tipo "
         Height          =   525
         Left            =   1035
         TabIndex        =   8
         Top             =   1560
         Width           =   3525
         Begin VB.OptionButton optbytTipo 
            Caption         =   "Domingo"
            Height          =   195
            Index           =   2
            Left            =   2430
            TabIndex        =   11
            Top             =   270
            Width           =   1005
         End
         Begin VB.OptionButton optbytTipo 
            Caption         =   "Sábado"
            Height          =   195
            Index           =   1
            Left            =   1290
            TabIndex        =   10
            Top             =   270
            Width           =   975
         End
         Begin VB.OptionButton optbytTipo 
            Caption         =   "Feriado"
            Height          =   195
            Index           =   0
            Left            =   180
            TabIndex        =   9
            Top             =   270
            Width           =   945
         End
      End
      Begin VB.TextBox txtstrDescricao 
         Height          =   285
         Left            =   1050
         MaxLength       =   38
         TabIndex        =   5
         Top             =   810
         Width           =   4695
      End
      Begin TrueOleDBGrid70.TDBGrid tdb_Dias 
         Height          =   2805
         Left            =   150
         TabIndex        =   13
         Top             =   2160
         Width           =   5565
         _ExtentX        =   9816
         _ExtentY        =   4948
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).DataField=   "PKID"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Descrição"
         Columns(1).DataField=   "strDescricao"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Data"
         Columns(2).DataField=   "dtmData"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   16
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Tipo"
         Columns(3).DataField=   "bytTipo"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   4
         Splits(0)._UserFlags=   0
         Splits(0).MarqueeStyle=   3
         Splits(0).RecordSelectors=   0   'False
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).DividerColor=   12632256
         Splits(0).FilterBar=   -1  'True
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=4"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=1138"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1058"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0).Visible=0"
         Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(7)=   "Column(1).Width=4868"
         Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=4789"
         Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(12)=   "Column(2).Width=1984"
         Splits(0)._ColumnProps(13)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(14)=   "Column(2)._WidthInPix=1905"
         Splits(0)._ColumnProps(15)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(16)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(17)=   "Column(3).Width=2275"
         Splits(0)._ColumnProps(18)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(19)=   "Column(3)._WidthInPix=2196"
         Splits(0)._ColumnProps(20)=   "Column(3)._EditAlways=0"
         Splits(0)._ColumnProps(21)=   "Column(3).Order=4"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   0
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
      Begin VB.Label lbldtmData 
         AutoSize        =   -1  'True
         Caption         =   "Data"
         Height          =   195
         Left            =   600
         TabIndex        =   6
         Top             =   1260
         Width           =   345
      End
      Begin VB.Label lblintExercicio 
         AutoSize        =   -1  'True
         Caption         =   "Exercício"
         Height          =   195
         Left            =   270
         TabIndex        =   1
         Top             =   540
         Width           =   675
      End
      Begin VB.Label lblstrDescricao 
         AutoSize        =   -1  'True
         Caption         =   "Descrição"
         Height          =   195
         Left            =   225
         TabIndex        =   4
         Top             =   900
         Width           =   720
      End
   End
End
Attribute VB_Name = "frmCadDiasNaoUteis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mblnAlterando       As Boolean
Dim mobjAux             As Object
Dim mblnSelecionou      As Boolean
Dim mblnPrimeiraVez     As Boolean
Dim bytOrdenacao        As Byte
Dim blnOrdenacaoAsc     As Boolean
    
Private Sub cbointExercicio_Click()
    mblnAlterando = False
End Sub

Private Sub cmd_Exercicio_Click()
    Dim vntExercicio As Variant
    vntExercicio = InputBox("Digite o exercício:", "Dias não úteis", "")
    If Val(vntExercicio) = 0 Then
        ExibeMensagem "Exercício inválido."
        Exit Sub
    Else
        If Val(vntExercicio) >= 1930 And Val(vntExercicio) <= 2099 Then
            If gblnExisteValorNaTabela(gstrDiasNaoUteis, "intExercicio", Val(vntExercicio)) Then
               ExibeMensagem "Exercício já existente."
               Exit Sub
            End If
            Preenche_SDF Val(vntExercicio)
            LeDaTabelaParaObj gstrDiasNaoUteis, cbointExercicio, strQueryExercicio
            cbointExercicio.ListIndex = gintIndiceCBO(cbointExercicio, Val(vntExercicio))
        Else
            ExibeMensagem "Exercício inválido."
            Exit Sub
        End If
    End If
End Sub

Private Sub Form_Activate()
    gintCodSeguranca = 599
    VirificaGradeListView Me
    If mblnSelecionou Then
        HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrAplicar, gstrDeletar
    Else
        HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar
    End If
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
    bytOrdenacao = 1: blnOrdenacaoAsc = True
    LeDaTabelaParaObj gstrDiasNaoUteis, cbointExercicio, strQueryExercicio
'    VerificaPermissoes Me, Me.tlb_BarraFermta, Me.Tag
    VerificaObjParaAplicar mobjAux
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar, gstrDeletar
    mblnSelecionou = False
    mblnPrimeiraVez = False
End Sub

Private Sub tdb_Dias_Click()
    
    mblnPrimeiraVez = True
    
    With tdb_Dias
        If Not .BOF And Not .EOF Then
            If .Bookmark = 1 Then
                tdb_Dias_RowColChange 0, 0
            End If
        End If
    End With

End Sub

Private Sub tdb_Dias_FilterChange()
    mblnPrimeiraVez = False
    gblnFilraCampos tdb_Dias
End Sub

Private Sub tdb_Dias_HeadClick(ByVal ColIndex As Integer)
    gOrdenaGrid tdb_Dias, ColIndex
End Sub

Private Sub tdb_Dias_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    
    With tdb_Dias
        If Not .EOF And Not .BOF Then
            txtPKID.Text = .Columns("PKID").Value
            If mblnPrimeiraVez Then
                mblnAlterando = True
                LeDaTabelaParaObj gstrDiasNaoUteis, Me
'=============
                gCorLinhaSelecionada tdb_Dias
                
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

    Select Case UCase(strModoOperacao)
        
        Case UCase(gstrNovo)
            LimpaObjeto Me
            Exit Sub
            
        Case UCase(gstrSalvar)
            If Not blnDadosOk Then Exit Sub
        
        Case UCase(gstrLocalizar)
            LeDaTabelaParaObj gstrDiasNaoUteis, tdb_Dias, strQueryDias(True)
            Exit Sub
            
        Case UCase(gstrFechar)
            Unload Me
            Exit Sub
            
    End Select
    
    If UCase(strModoOperacao) = "SALVAR" Or UCase(strModoOperacao) = "DELETAR" Then
        mblnPrimeiraVez = False
    End If
    
    ToolBarGeral strModoOperacao, gstrDiasNaoUteis, mblnAlterando, tdb_Dias, Me, mobjAux, strQueryDias, , rptDiasNaoUteis, strQuerryRelatorio
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrDeletar
    
End Sub

Private Function blnDadosOk() As Boolean
    
    blnDadosOk = False
    
    If cbointExercicio.ListIndex = -1 Then
        ExibeMensagem "O exercício tem que ser selecionado."
        cbointExercicio.SetFocus
        Exit Function
    ElseIf Trim(txtstrDescricao.Text) = "" Then
        ExibeMensagem "A descrição tem que ser informada."
        txtstrDescricao.SetFocus
        Exit Function
    ElseIf optbytTipo(0).Value = False And optbytTipo(1).Value = False And optbytTipo(2).Value = False Then
        ExibeMensagem "Deve ser selecionado algum tipo."
        optbytTipo(0).SetFocus
        Exit Function
    End If
    
    blnDadosOk = True
    
End Function

Private Sub txtstrDescricao_GotFocus()
    MarcaCampo txtstrDescricao
End Sub

Private Sub txtstrDescricao_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrDescricao
End Sub

Sub Preenche_SDF(intExercicio As Integer)
    Dim vntData      As Date
    Dim strDescricao As String
    Dim STRTIPO      As String
    
    For vntData = CDate("01/01/" & intExercicio) To CDate("31/12/" & intExercicio)
        Select Case Weekday(vntData)
            Case vbSaturday
                STRTIPO = "1"
                strDescricao = "Sábado"
            Case vbSunday
                STRTIPO = "2"
                strDescricao = "Domingo"
            Case Else
                GoTo Proximo
        End Select
        strDescricao = ""
        AdicionaDiasNaoUteis intExercicio, strDescricao, vntData, STRTIPO
Proximo:
    Next
End Sub

Private Function strQueryExercicio() As String

'******************************************************************************************
' Data: 06/05/2003
' Alteração: - Adaptação da cláusula ORDER BY, de forma que a cláusula utiliza-se os
'            índices das colunas e não os nomes.
' Responsável: Everton Bianchini
'******************************************************************************************

    Dim strsql As String
    strsql = ""
    strsql = strsql & "Select Distinct intExercicio , intExercicio "
    strsql = strsql & "From " & gstrDiasNaoUteis & " "
'    strSQL = strSQL & "Order By intExercicio Desc"
    strsql = strsql & "Order By 1 Desc"
    strQueryExercicio = strsql
End Function

Private Function strQueryDias(Optional blnFiltrar As Boolean = False) As String
Dim strsql As String
    
    strsql = ""
    strsql = strsql & "Select PKId, strDescricao, dtmData, " & gstrCASEWHEN("bytTipo", "0,'Feriado',1,'Sábado',2,'Domingo'") & " bytTipo "
    strsql = strsql & "From " & gstrDiasNaoUteis & " "
    
    If blnFiltrar Then
    
        If cbointExercicio.ListIndex > -1 Then
            strsql = strsql & "Where intExercicio = " & gstrItemData(cbointExercicio)
        End If
    
        If Trim(txtstrDescricao.Text) <> "" Then
            strsql = strsql & IIf(InStr(1, strsql, "Where") > 0, " And ", "Where") & " strDescricao Like '" & txtstrDescricao.Text & "%'"
        End If
    
    End If
    
    Select Case bytOrdenacao
        Case Is = 1
            strsql = strsql & " ORDER BY strDescricao" & IIf(blnOrdenacaoAsc, " ASC", " DESC")
        Case Is = 2
            strsql = strsql & " ORDER BY dtmData" & IIf(blnOrdenacaoAsc, " ASC", " DESC")
        Case Is = 3
            strsql = strsql & " ORDER BY bytTipo" & IIf(blnOrdenacaoAsc, " ASC", " DESC")
    End Select
    
    strQueryDias = strsql
    
End Function

Sub AdicionaDiasNaoUteis(intAuxExercicio As Integer, strAuxDescricao As String, DTMDATA As Date, strAuxTipo As String)

'******************************************************************************************
' Data: 02/05/2003
' Alteração: - Substituição do comando nativo GETDATE() do SQL Server pela variável
'            strGETDATE.
' Responsável: Everton Bianchini
'******************************************************************************************
    
    Dim strsql As String

    strsql = ""
    strsql = strsql & "Insert Into " & gstrDiasNaoUteis & " ("
    strsql = strsql & "intExercicio, strDescricao, dtmData, bytTipo, dtmDtAtualizacao, "
    strsql = strsql & "lngCodUsr"
    strsql = strsql & ") Values ("
    strsql = strsql & intAuxExercicio & ", '"
    strsql = strsql & strAuxDescricao & "', "
    strsql = strsql & gstrConvDtParaSql(DTMDATA) & ", "
    strsql = strsql & strAuxTipo & ", "
'    strSql = strSql & "GETDATE(), "
    strsql = strsql & strGETDATE & ", "
    strsql = strsql & glngCodUsr & ")"
    
    Set gobjBanco = New clsBanco
    gobjBanco.Execute strsql
End Sub

Function strQuerryRelatorio() As String
Dim strsql As String
    strsql = ""
    strsql = strsql & "SELECT * "
    strsql = strsql & " FROM " & gstrDiasNaoUteis
    If mblnAlterando = True Then
        strsql = strsql & " WHERE PKId = " & tdb_Dias.Columns("PKId").Value  '& " and intExercicio = " & gstrItemData(cbointExercicio)
    End If
    If mblnAlterando = False And cbointExercicio.Text <> "" Then
        strsql = strsql & " WHERE intExercicio = " & gstrItemData(cbointExercicio)
    End If
strQuerryRelatorio = strsql
End Function


