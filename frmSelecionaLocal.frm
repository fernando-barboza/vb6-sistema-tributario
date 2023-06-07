VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmSelecionaLocal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Selecione Unidade Centro de Custo"
   ClientHeight    =   1275
   ClientLeft      =   3240
   ClientTop       =   5355
   ClientWidth     =   6210
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1275
   ScaleWidth      =   6210
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   405
      Left            =   4800
      TabIndex        =   2
      Top             =   690
      Width           =   1290
   End
   Begin MSDataListLib.DataCombo dbcintCustoDestino 
      Height          =   315
      Left            =   750
      TabIndex        =   0
      Top             =   195
      Width           =   5355
      _ExtentX        =   9446
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
   End
   Begin VB.Label lbl_intCustoDestino 
      AutoSize        =   -1  'True
      Caption         =   "Destino"
      Height          =   195
      Left            =   150
      TabIndex        =   1
      Top             =   240
      Width           =   540
   End
End
Attribute VB_Name = "frmSelecionaLocal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public intProtocolo     As Long
Public intExercicio     As Integer
Public strCentroCusto   As String
Public strDataProcesso  As String

Private Sub cmdOK_Click()
    Dim strSql As String
    Dim strTop As String
    Dim adoResultado As ADODB.Recordset
    Dim ProtocolizacaoVolume As String
    
    If Not blnDadosOk Then Exit Sub
    
    
    
    strSql = "INSERT INTO tblTramiteProtocolo "
    strSql = strSql & "(intProtocolizacaoVolume,strCodigoDestino, strcodigoorigem, intTempoEstimado, intCustoDestino, intCodDespacho, dtmDataSolicitacao, "
    strSql = strSql & "lngCodUsr, dtmDtAtualizacao, strDescricaoDespacho, intFolhas) "
    
    If bytDBType = SQLServer Then
        strTop = "SELECT " & gstrTOPnSQLServer(1) & " PKId FROM " & gstrProtocolizacaoVolume & " WHERE intProtocolizacaoProcesso = " & glngPegaUltimaChave(gstrProtocolizacaoProcesso, "PKID", "strCodigo", intProtocolo, "intExercicio", intExercicio) & " ORDER BY intVolume DESC"
        Set gobjBanco = New clsBanco
            If gobjBanco.CriaADO(strTop, 10, adoResultado) Then
                With adoResultado
                    If .EOF = False Then
                       ProtocolizacaoVolume = Trim(!Pkid)
                    End If
                End With
            End If
    Else
        strTop = "SELECT  PKId FROM " & gstrProtocolizacaoVolume & " WHERE intProtocolizacaoProcesso = " & glngPegaUltimaChave(gstrProtocolizacaoProcesso, "PKID", "strCodigo", intProtocolo, "intExercicio", intExercicio) & " ORDER BY intVolume DESC"
        strTop = gstrTOPnOracle(strTop, 1)
        Set gobjBanco = New clsBanco
                If gobjBanco.CriaADO(strTop, 10, adoResultado) Then
                    With adoResultado
                        If .EOF = False Then
                           ProtocolizacaoVolume = Trim(!Pkid)
                           
                        End If
                    End With
                End If
    
    End If
    
    strSql = strSql & "VALUES (" & ProtocolizacaoVolume & ",'','','', " & dbcintCustoDestino.BoundText & ",1, " & strDataProcesso & ", " & glngCodUsr & ", " & strGETDATE & ", 'PRIMEIRA TRAMITAÇÃO',1)"
    
    Set gobjBanco = New clsBanco
    
    gobjBanco.Execute strSql
    
    Unload Me
    
    Set gobjBanco = Nothing

End Sub

Private Sub dbcintCustoDestino_Click(Area As Integer)
    DropDownDataCombo dbcintCustoDestino, Me, Area
End Sub

Private Sub dbcintCustoDestino_GotFocus()
    MarcaCampo dbcintCustoDestino
End Sub

Private Sub dbcintCustoDestino_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbcintCustoDestino, Me, , KeyCode, Shift
End Sub

Private Sub dbcintCustoDestino_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbcintCustoDestino
End Sub

Private Sub Form_Load()
    dbcintCustoDestino.Tag = strQueryLocais & ";strDescricao"
    LeDaTabelaParaObj gstrLocais, dbcintCustoDestino, strQueryLocais
    
End Sub

Private Function strQueryLocais() As String
Dim strSql As String

    strSql = ""
    strSql = strSql & " SELECT A.PkId, A.strDescricao"
    strSql = strSql & " FROM"
    strSql = strSql & " " & gstrLocais & " A"
    strSql = strSql & " WHERE A.bitInformatizada = 1 AND dtmCancelamento IS NULL "
    strSql = strSql & " ORDER BY A.strDescricao"
    strQueryLocais = strSql
    
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If UnloadMode <> vbFormCode Then Cancel = 1
End Sub

Private Function blnDadosOk() As Boolean

    If Not dbcintCustoDestino.MatchedWithList Then
        ExibeMensagem "Preencha corretamente o campo Centro de Custo Destino!", vbInformation
        dbcintCustoDestino.SetFocus
        Exit Function
    End If
    
    blnDadosOk = True
End Function

Public Sub MantemForm(ByVal strModoOperacao As String)
    If strModoOperacao = gstrPreencherLista Then
        PreencherListaDeOpcoes Me.ActiveControl
    End If

End Sub

