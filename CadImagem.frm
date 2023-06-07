VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCadImagem 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Imagens (Brasão/Logotipo)"
   ClientHeight    =   3720
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   5880
   HelpContextID   =   111
   Icon            =   "CadImagem.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   5880
   Begin VB.TextBox txtPKId 
      Height          =   285
      Left            =   7380
      TabIndex        =   2
      Top             =   30
      Visible         =   0   'False
      Width           =   765
   End
   Begin TabDlg.SSTab tabFotos 
      Height          =   3525
      Left            =   105
      TabIndex        =   0
      Top             =   90
      Width           =   5670
      _ExtentX        =   10001
      _ExtentY        =   6218
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   420
      TabCaption(0)   =   "Imagem"
      TabPicture(0)   =   "CadImagem.frx":1042
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "imgFotinho"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lvw_Lista"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      Begin MSComctlLib.ListView lvw_Lista 
         Height          =   2985
         Left            =   90
         TabIndex        =   1
         Top             =   420
         Width           =   3825
         _ExtentX        =   6747
         _ExtentY        =   5265
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         Icons           =   "img_Arquivo"
         SmallIcons      =   "img_Arquivo"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Arquivo"
            Object.Width           =   5997
         EndProperty
      End
      Begin VB.Image imgFotinho 
         BorderStyle     =   1  'Fixed Single
         Height          =   1860
         Left            =   3990
         MouseIcon       =   "CadImagem.frx":105E
         MousePointer    =   99  'Custom
         Stretch         =   -1  'True
         Top             =   450
         Width           =   1470
      End
   End
End
Attribute VB_Name = "frmCadImagem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim mimgAux          As Image
    Dim mtxtCodFoto      As TextBox
    Dim mblnAlterando    As Boolean
    Dim mobjAux          As Object

Private Sub Form_Activate()
    VirificaGradeListView Me
    HabilitaDesabilitaBotao1 mblnAlterando, gstrMnuArquivo, gstrDeletar
    HabilitaDesabilitaBotao1 ((Not mobjAux Is Nothing) And mblnAlterando), gstrMnuArquivo, gstrAplicar, gstrSalvar
End Sub

Private Sub Form_Deactivate()
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrSalvar, gstrAplicar, gstrDeletar, gstrBrasao, gstrLogotipo
End Sub

Public Sub Aplicar()
    If lvw_Lista.ListItems.Count > 0 Then
        mtxtCodFoto.Text = Val(lvw_Lista.SelectedItem.Tag)
        mimgAux.Picture = imgFotinho.Picture
        Unload Me
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar
End Sub

Public Sub CadastraFoto(imgAux As Image, txtCodFoto As TextBox)
    Set mtxtCodFoto = txtCodFoto
    Set mimgAux = imgAux
    Me.Show
    Me.SetFocus
End Sub

Sub NovaFoto()

'******************************************************************************************
' Data: 11/03/2003
' Alteração: - Alteração da string de conexão.
' Responsável: Everton Bianchini
'******************************************************************************************
    
    Const intChunkSize  As Integer = 16384
    Dim adoResultado    As ADODB.Recordset
    Dim adoConexao      As ADODB.Connection
    Dim strFiltro       As String
    Dim lngInd          As Long
    Dim intNumArquivo   As Integer
    Dim strCaminhoFoto  As String
    Dim strNomeFoto     As String
    Dim strConexao      As String
    Dim intAux          As Integer
    Dim strSql          As String
    Dim intFragmento    As Integer
    Dim intChunks       As Integer
    Dim bytChunk()      As Byte
    
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrSalvar
    strFiltro = ""
    strFiltro = strFiltro & "Bitmap Files (*.bmp)|*.bmp|GIF Files (*.gif)|*.gif|"
    strFiltro = strFiltro & "JPG Files (*.jpg)|*.jpg|All Files (*.*)|*.*"

    strCaminhoFoto = gstrNomeArquivoParaAbrir(strNomeFoto, True, strFiltro, , , _
                                              "Selecione o arquivo", "*.txt", Me.Hwnd, _
                                              cdlOFNExplorer Or cdlOFNHideReadOnly Or _
                                              cdlOFNLongNames)
    If Trim(strCaminhoFoto) <> "" Then
        If gblnExclusaoGravacaoOk("I", "do " + CStr(frmCadImagem.Caption)) Then
            intNumArquivo = FreeFile
            Open strCaminhoFoto For Binary Access Read As intNumArquivo
            'Verifica o tamanho da imagem
            If LOF(intNumArquivo) = 0 Then
                Close intNumArquivo
                Exit Sub
            End If
            intChunks = LOF(intNumArquivo) \ intChunkSize
            intFragmento = LOF(intNumArquivo) Mod intChunkSize
            'Abre a conexão
'            strConexao = "driver={SQL Server};server=" & gstrServidor & ";" & _
'                         "uid=" & gstrLoginUser & ";pwd=" & gstrPwdUser & ";database=" & gstrDatabase
            If bytDBType = EDatabases.SQLServer Then
                strConexao = "driver={SQL Server};server=" & gstrServidor & ";" & _
                             "uid=" & gstrUsername & ";pwd=" & gstrPassword & ";database=" & gstrDatabase
            ElseIf bytDBType = EDatabases.Oracle Then
            strConexao = "Provider=MSDAORA.1;Password=" & gstrPassword & ";User ID=" & gstrUsername & _
                        ";Data Source=" & gstrServidor & ";Persist Security Info=True"
            
                'strConexao = "Provider=MSDASQL.1;Password=" & gstrPassword & ";User ID=" & gstrUsername & _
                             '";Data Source=" & gstrServidor & ";Persist Security Info=True"
            End If
            Set adoConexao = New ADODB.Connection
            adoConexao.CursorLocation = adUseClient
            adoConexao.Open strConexao
            ' Open the table.
            Set adoResultado = New ADODB.Recordset
            adoResultado.CursorLocation = adUseClient
            adoResultado.CursorType = adOpenKeyset
            adoResultado.LockType = adLockOptimistic
            adoResultado.Open gstrImagem, adoConexao, , , adCmdTable
            adoResultado.AddNew
            adoResultado!imgImagem.AppendChunk ""
            adoResultado!dtmDtAtualizacao = Format(Date, "yyyy/mm/dd")
            adoResultado!lngCodUsr = glngCodUsr
            adoResultado!strDescricao = strNomeFoto
            ReDim bytChunk(intFragmento)
            Get intNumArquivo, , bytChunk()
            adoResultado!imgImagem.AppendChunk bytChunk()
            ReDim bytChunk(intChunkSize)
            For lngInd = 1 To intChunks
                Get intNumArquivo, , bytChunk()
                adoResultado!imgImagem.AppendChunk bytChunk()
            Next lngInd
            Close intNumArquivo
            adoResultado.Update
            adoResultado.Close
            CarregaFotos
        End If
    End If
End Sub

Public Sub MantemForm(ByVal strModoOperacao As String)
    Select Case UCase(strModoOperacao)
    Case "NOVO"
        LimpaObjeto Me, mblnAlterando
        NovaFoto
    'Case "SALVAR"
        'If ToolBarGeral(strModoOperacao, gstrAgencia, _
           'mblnAlterando, lvw_Lista, Me) Then
        'End If
    Case "DELETAR"
        DeletaImagem
    Case "APLICAR"
        Aplicar
    Case "FECHAR"
        Unload Me
    End Select
End Sub

Private Sub Form_Load()
    CarregaFotos
    If mtxtCodFoto Is Nothing = False Then
        With lvw_Lista
            Call gblnEncontroItemNoListView(lvw_Lista, CStr(mtxtCodFoto))
            LeImagem Val(mtxtCodFoto), imgFotinho
        End With
    End If
End Sub

Sub CarregaFotos()
    Dim adoResultado    As ADODB.Recordset
    Dim objList         As Object
    Dim strSql          As String
    lvw_Lista.ListItems.Clear
    strSql = ""
    strSql = strSql & "SELECT PKId, strDescricao FROM "
    strSql = strSql & gstrImagem & " "
    strSql = strSql & "ORDER BY strDescricao"
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        With adoResultado
            Do While Not .EOF
                Set objList = lvw_Lista.ListItems.Add(, , Trim(!strDescricao))
                objList.Tag = !Pkid
                .MoveNext
            Loop
        End With
    End If
End Sub

Private Sub imgFotinho_DblClick()
    Aplicar
End Sub

Private Sub imgFotinho_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    FinalizaDragDrop imgFotinho, State
End Sub

Private Sub lvw_Lista_DblClick()
    Aplicar
End Sub

Private Sub lvw_Lista_ItemClick(ByVal Item As MSComctlLib.ListItem)
    With lvw_Lista
        txtPKId = lvw_Lista.SelectedItem.Tag
        LeImagem Val(txtPKId), imgFotinho
          HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrDeletar, gstrAplicar
           HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrSalvar
    End With
End Sub

Private Function DeletaImagem()
    Dim strSql As String
    If gblnExclusaoGravacaoOk("E", "Deseja excluir o " + CStr(frmCadImagem.Caption) + " ?", True) Then
        strSql = ""
        strSql = strSql & "DELETE FROM " & gstrImagem & " "
        strSql = strSql & "WHERE PKId = " & lvw_Lista.SelectedItem.Tag
        Set gobjBanco = New clsBanco
        gobjBanco.Execute strSql
    End If
    CarregaFotos
    LeImagem Val(txtPKId), imgFotinho
End Function


