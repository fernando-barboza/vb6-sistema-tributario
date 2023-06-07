VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmAcordoCarne2ViaAtualizada 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Acordo - Carnê 2ª Via Atualizada"
   ClientHeight    =   1920
   ClientLeft      =   4065
   ClientTop       =   5205
   ClientWidth     =   4590
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1920
   ScaleWidth      =   4590
   Begin VB.TextBox txtExercicio 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   3855
      MaxLength       =   4
      TabIndex        =   6
      Top             =   90
      Width           =   630
   End
   Begin VB.Frame fra_Mensagem1 
      Caption         =   "Faixa de Acordo"
      Height          =   1380
      Left            =   90
      TabIndex        =   0
      Top             =   450
      Width           =   4425
      Begin VB.CheckBox chkTodas 
         Caption         =   "Selecionar todos os Acordos"
         Height          =   255
         Left            =   1305
         TabIndex        =   1
         Top             =   1035
         Width           =   2835
      End
      Begin MSDataListLib.DataCombo dbcstrAcordoInicial 
         Height          =   315
         Left            =   1305
         TabIndex        =   2
         Top             =   270
         Width           =   3000
         _ExtentX        =   5292
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dbcstrAcordoFinal 
         Height          =   315
         Left            =   1305
         TabIndex        =   3
         Top             =   675
         Width           =   3000
         _ExtentX        =   5292
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin VB.Label lbl_Label 
         AutoSize        =   -1  'True
         Caption         =   "Acordo Inicial:"
         Height          =   195
         Left            =   180
         TabIndex        =   5
         Top             =   360
         Width           =   1005
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Acordo Final:"
         Height          =   195
         Left            =   270
         TabIndex        =   4
         Top             =   765
         Width           =   930
      End
   End
   Begin VB.Label lblExercicioVencto 
      AutoSize        =   -1  'True
      Caption         =   "Exercício do Vencimento das Parcelas:"
      Height          =   195
      Left            =   960
      TabIndex        =   7
      Top             =   135
      Width           =   2790
   End
End
Attribute VB_Name = "frmAcordoCarne2ViaAtualizada"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim mobjAux          As Object
    Dim mblnSelecionou   As Boolean

Private Sub chkTodas_Click()
  If chkTodas.Value = 1 Then
     dbcstrAcordoInicial.Enabled = False
     dbcstrAcordoFinal.Enabled = False
  Else
     dbcstrAcordoInicial.Enabled = True
     dbcstrAcordoFinal.Enabled = True
  End If
End Sub

Private Sub dbcstrAcordoFinal_Click(Area As Integer)
  DropDownDataCombo dbcstrAcordoFinal, Me, Area
End Sub

Private Sub dbcstrAcordoFinal_GotFocus()
  If Trim(dbcstrAcordoFinal.Text) = "" Then
     dbcstrAcordoFinal.Text = Trim(dbcstrAcordoInicial.Text)
     dbcstrAcordoFinal_Click 0
     dbcstrAcordoFinal.Text = Trim(dbcstrAcordoInicial.Text)
  End If
  MarcaCampo dbcstrAcordoFinal
End Sub

Private Sub dbcstrAcordoInicial_Click(Area As Integer)
  DropDownDataCombo dbcstrAcordoInicial, Me, Area
End Sub

Private Sub dbcstrAcordoInicial_GotFocus()
  If Trim(dbcstrAcordoInicial.Text) = "" Then
     dbcstrAcordoInicial.Text = Trim(dbcstrAcordoFinal.Text)
     dbcstrAcordoInicial_Click 0
     dbcstrAcordoInicial.Text = Trim(dbcstrAcordoFinal.Text)
  End If
  MarcaCampo dbcstrAcordoInicial
End Sub

Private Sub Form_Load()
    HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrSalvar, gstrDeletar
    txtExercicio.Text = Year(gstrDataDoSistema)
End Sub

Private Sub Form_Activate()

    gintCodSeguranca = 1408
    
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

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar
    mblnSelecionou = False
End Sub

Private Sub txtExercicio_Change()
  dbcstrAcordoFinal.Text = ""
  dbcstrAcordoInicial.Text = ""
  Set dbcstrAcordoFinal.RowSource = Nothing
  Set dbcstrAcordoInicial.RowSource = Nothing
End Sub

Private Sub txtExercicio_GotFocus()
  MarcaCampo txtExercicio
End Sub

Private Sub txtExercicio_KeyPress(KeyAscii As Integer)
  CaracterValido KeyAscii, "N", txtExercicio
End Sub

Public Sub MantemForm(ByVal strModoOperacao As String)
Dim strInscricaoInicial As String
Dim strInscricaoFinal As String

  Select Case UCase(strModoOperacao)
  
    Case UCase(gstrPreencherLista)
      If Left(Me.ActiveControl.Name, 3) = "dbc" Then
        If Len(txtExercicio.Text) = 4 Then
            LeDaTabelaParaObj "", Me.ActiveControl, strQueryComboAcordo(Trim(Me.ActiveControl.Text))
        Else
            ExibeMensagem "É necessário preencher o campo de exercício."
            txtExercicio.SetFocus
        End If
      End If
      
    Case Is = UCase(gstrImprimir)
      If blnDadosOk = True Then
          
          If blnEmissaoDe2ViaAtualizada Then
              
              If blnExisteIndexadorNoExercicio Then
              
                  If dbcstrAcordoInicial.MatchedWithList Then
                      strInscricaoInicial = Replace(dbcstrAcordoInicial.Text, "/", "")
                  End If
                  If dbcstrAcordoFinal.MatchedWithList Then
                      strInscricaoFinal = Replace(dbcstrAcordoFinal.Text, "/", "")
                  End If
                
                  rptCapaCarneAcordo.blnParcelasAtualizadas = True
                  rptCapaCarneAcordo.intExercicioAtualizadas = txtExercicio.Text
                
                  ImprimeRelatorio rptCapaCarneAcordo, gstrQueryCarneAcordoAtualizadas(strInscricaoInicial, strInscricaoFinal, txtExercicio.Text, IIf(chkTodas.Value = 1, True, False)), , 1000
              Else
                  ExibeMensagem "Não existe Indexador para o Exercício " & txtExercicio.Text & "."
              End If
          Else
              ExibeMensagem "Operação não permitida. Consulte em Parâmetros a opção 2ª Via Atualizada."
          End If
          
      End If
      
    Case UCase(gstrNovo)
      LimpaObjetos
    
    Case UCase(gstrFechar)
      Unload Me
      
  End Select
    
End Sub

Private Function blnDadosOk() As Boolean
Dim intAux              As Integer
Dim strInscricaoInicial As String
Dim strInscricaoFinal   As String

  blnDadosOk = False
  
  If Len(txtExercicio.Text) <> 4 Then
     ExibeMensagem "O exercício deve ser preenchido corretamente."
     txtExercicio.SetFocus
     Exit Function
  End If

  If dbcstrAcordoFinal.MatchedWithList = False And dbcstrAcordoInicial.MatchedWithList = False And chkTodas.Value = 0 Then
     ExibeMensagem "O Acordo deve ser informado."
     dbcstrAcordoInicial.SetFocus
     Exit Function
  End If
  
  If dbcstrAcordoInicial.MatchedWithList = True And dbcstrAcordoFinal.MatchedWithList = True Then
    
     strInscricaoInicial = Replace(dbcstrAcordoInicial.Text, "/", "")
     strInscricaoFinal = Replace(dbcstrAcordoFinal.Text, "/", "")

     If Val(Right(String(gintLenInscricao - Len(Trim(strInscricaoFinal)), "0") & Trim(strInscricaoFinal), 4) & Left(String(gintLenInscricao - Len(Trim(strInscricaoFinal)), "0") & Trim(strInscricaoFinal), 16)) < Val(Right(String(gintLenInscricao - Len(Trim(strInscricaoInicial)), "0") & Trim(strInscricaoInicial), 4) & Left(String(gintLenInscricao - Len(Trim(strInscricaoInicial)), "0") & Trim(strInscricaoInicial), 16)) Then
        ExibeMensagem "O Acordo Inicial não pode ser maior que o Acordo Final."
        dbcstrAcordoFinal.SetFocus
        Exit Function
     End If
  End If
  
  blnDadosOk = True
  
End Function

Private Function strQueryComboAcordo(Inscricao As String) As String
Dim strSql As String

  strSql = ""
  strSql = strSql & "SELECT "
  strSql = strSql & "LA.pkID, "
  strSql = strSql & strSUBSTRING & "(LA.strInscricao, " & gintLenInscricao - gintRetornaTamanhoMascara(TYP_ACORDO) + 1 & ", " & gintRetornaTamanhoMascara(TYP_ACORDO) - 4 & ")" & strCONCAT & " '/' " & strCONCAT & strSUBSTRING & "(LA.strInscricao, 17, 4) strInscricao "
  strSql = strSql & "FROM "
  strSql = strSql & gstrLancamentoAlfa & " LA, " & gstrLancamentoValor & " LV "
  strSql = strSql & "WHERE "
 
  If Inscricao <> "" Then
    If Len(Trim(txtExercicio.Text)) = 4 Then
        strSql = strSql & " LA.strInscricao Like '" & String(gintLenInscricao - Len(Val(Inscricao) & txtExercicio.Text), "0") & Val(Inscricao) & "%' AND "
    End If
  End If
  
  strSql = strSql & " LV.intLancamentoAlfa = LA.Pkid AND "
  strSql = strSql & gstrDATEPART(strYEAR, "LV.dtmDtVencimento") & " = " & Trim(txtExercicio.Text) & " AND "
  strSql = strSql & " LA.dtmDtCancelamento Is Null AND "
  strSql = strSql & " LA.intUtilizacao = " & TYP_ACORDO
  
  strSql = strSql & " GROUP BY LA.Pkid, LA.strInscricao, LA.intExercicio "
  strSql = strSql & " ORDER BY LA.intExercicio, strInscricao "

  strQueryComboAcordo = strSql
End Function

Private Sub LimpaObjetos()
  chkTodas.Value = 0
  dbcstrAcordoFinal.Enabled = True
  dbcstrAcordoFinal.Text = ""
  dbcstrAcordoInicial.Enabled = True
  dbcstrAcordoInicial.Text = ""
  txtExercicio.Text = Year(gstrDataDoSistema)
  txtExercicio.SetFocus
End Sub

Private Function blnEmissaoDe2ViaAtualizada() As Boolean
Dim strSql As String
Dim adoResultado As ADODB.Recordset
    
    blnEmissaoDe2ViaAtualizada = False
    
    strSql = "SELECT PT.blnEmite2ViaAcordoAtualizado "
    strSql = strSql & "FROM " & gstrParametrosTributario & " PT "
  
    Set gobjBanco = New clsBanco
    Set adoResultado = New ADODB.Recordset
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
       If Not adoResultado.EOF Then
           blnEmissaoDe2ViaAtualizada = Abs(adoResultado("blnEmite2ViaAcordoAtualizado").Value) = 1
       End If
    End If
    Set adoResultado = Nothing
  
End Function

Private Function blnExisteIndexadorNoExercicio() As Boolean
Dim strSql       As String
Dim adoResultado As New ADODB.Recordset

    strSql = "SELECT  FA.DBLVALOR "
    strSql = strSql & " FROM " & gstrParametroAtualizacao & " PA, "
    strSql = strSql & gstrFormaAtualizacaoValor & " FA, "
    strSql = strSql & gstrComposicaoDaReceita & " CR "
    strSql = strSql & " WHERE FA.INTINDEXADORECONOMICO = PA.intIndexadorEconomico AND "
    strSql = strSql & " FA.DTMDATA = " & gstrConvDtParaSql("01/01/" & txtExercicio.Text) & " AND "
    strSql = strSql & " PA.INTCOMPOSICAORECEITA = CR.Pkid AND "
    strSql = strSql & " PA.intExercicio = " & txtExercicio.Text & " AND "
    strSql = strSql & " CR.intUtilizacao = " & TYP_ACORDO
    
    Set gobjBanco = New clsBanco
    Set adoResultado = New ADODB.Recordset
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
       blnExisteIndexadorNoExercicio = Not adoResultado.EOF
    End If
    
    Set adoResultado = Nothing
            
End Function
