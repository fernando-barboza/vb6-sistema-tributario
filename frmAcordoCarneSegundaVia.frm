VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Begin VB.Form frmAcordoCarneSegundaVia 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Acordo - Carnê 2ª Via"
   ClientHeight    =   3990
   ClientLeft      =   4980
   ClientTop       =   2700
   ClientWidth     =   4620
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3990
   ScaleWidth      =   4620
   Begin VB.Frame Frame1 
      Caption         =   "Parcelas"
      Height          =   2055
      Left            =   90
      TabIndex        =   8
      Top             =   1845
      Width           =   4425
      Begin MSComctlLib.ListView lvwParcelas 
         Height          =   1635
         Left            =   135
         TabIndex        =   4
         Top             =   270
         Width           =   4155
         _ExtentX        =   7329
         _ExtentY        =   2884
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Parcelas"
            Object.Width           =   5644
         EndProperty
      End
   End
   Begin VB.Frame fra_Mensagem1 
      Caption         =   "Faixa de Acordo"
      Height          =   1380
      Left            =   90
      TabIndex        =   5
      Top             =   450
      Width           =   4425
      Begin VB.CheckBox chkTodas 
         Caption         =   "Selecionar todos os Acordos"
         Height          =   255
         Left            =   1305
         TabIndex        =   3
         Top             =   1035
         Width           =   2835
      End
      Begin MSDataListLib.DataCombo dbcstrAcordoInicial 
         Height          =   315
         Left            =   1305
         TabIndex        =   1
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
         TabIndex        =   2
         Top             =   675
         Width           =   3000
         _ExtentX        =   5292
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Acordo Final:"
         Height          =   195
         Left            =   270
         TabIndex        =   7
         Top             =   765
         Width           =   930
      End
      Begin VB.Label lbl_Label 
         AutoSize        =   -1  'True
         Caption         =   "Acordo Inicial:"
         Height          =   195
         Left            =   180
         TabIndex        =   6
         Top             =   360
         Width           =   1005
      End
   End
   Begin VB.TextBox txtExercicio 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1005
      MaxLength       =   4
      TabIndex        =   0
      Top             =   90
      Width           =   1200
   End
   Begin VB.Label lblCapaDeLote 
      AutoSize        =   -1  'True
      Caption         =   "Exercício:"
      Height          =   195
      Left            =   240
      TabIndex        =   9
      Top             =   135
      Width           =   765
   End
End
Attribute VB_Name = "frmAcordoCarneSegundaVia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub MantemForm(ByVal strModoOperacao As String)
Dim strInscricaoInicial As String
Dim strInscricaoFinal As String

  Select Case UCase(strModoOperacao)
  
    Case UCase(gstrPreencherLista) 'COMBO
      'PreencherListaDeOpcoes Me.ActiveControl
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
         
         rptCapaCarneAcordo.strParcelasSelecionadas = strParcelas
         rptCapaCarneAcordo.blnParcelasAtualizadas = False
         rptCapaCarneAcordo.intExercicioAtualizadas = 0

         If Len(txtExercicio.Text) = 4 Then
            If dbcstrAcordoInicial.MatchedWithList Then
               strInscricaoInicial = dbcstrAcordoInicial.Text & txtExercicio.Text
            End If
            If dbcstrAcordoFinal.MatchedWithList Then
               strInscricaoFinal = dbcstrAcordoFinal.Text & txtExercicio.Text
            End If
            ImprimeRelatorio rptCapaCarneAcordo, gstrQueryCarneAcordo(strInscricaoInicial, strInscricaoFinal, strParcelas, IIf(chkTodas.Value = 1, True, False))
         Else
            If dbcstrAcordoInicial.MatchedWithList Then
               strInscricaoInicial = dbcstrAcordoInicial.Text
            End If
            If dbcstrAcordoFinal.MatchedWithList Then
               strInscricaoFinal = dbcstrAcordoFinal.Text
            End If
            If dbcstrAcordoInicial.MatchedWithList And dbcstrAcordoFinal.MatchedWithList Then
               If Exercicios(strInscricaoInicial, strInscricaoFinal) = False Then 'Busca os exercícios das inscrições informadas
                  ExibeMensagem "Erro na consulta dos exercícios das inscrições informadas. Não foi possível gerar o carnê."
                  Exit Sub
               End If
               ImprimeRelatorio rptCapaCarneAcordo, gstrQueryCarneAcordo(strInscricaoInicial, strInscricaoFinal, strParcelas, IIf(chkTodas.Value = 1, True, False))
            Else
               ImprimeRelatorio rptCapaCarneAcordo, gstrQueryCarneAcordo(strInscricaoInicial, strInscricaoFinal, strParcelas, IIf(chkTodas.Value = 1, True, False), False)
            End If
         End If
         
      End If
      
    Case UCase(gstrLocalizar) 'PREENCHE PARCELAS
      If dbcstrAcordoFinal.MatchedWithList = False And dbcstrAcordoInicial.MatchedWithList = False And chkTodas.Value = 0 Then
         ExibeMensagem "O Acordo deve ser informado."
         dbcstrAcordoInicial.SetFocus
         Exit Sub
      End If
      If dbcstrAcordoInicial.MatchedWithList = True And dbcstrAcordoFinal.MatchedWithList = True Then
         If Int(dbcstrAcordoFinal.Text) < Int(dbcstrAcordoInicial.Text) Then
            ExibeMensagem "O Acordo Inicial não pode ser maior que o Acordo Final."
            dbcstrAcordoFinal.SetFocus
            Exit Sub
         End If
      End If
      If Len(txtExercicio.Text) > 0 And Len(txtExercicio.Text) < 4 Then
         ExibeMensagem "O exercício deve ser preenchido corretamente."
         txtExercicio.SetFocus
         Exit Sub
      End If
      lvwParcelas.Checkboxes = True
      If PreencheParcelas Then
         If lvwParcelas.ListItems.Count <= 0 Then
            lvwParcelas.ListItems.Add , , "Pressione F5 para preencher as parcelas."
            lvwParcelas.Checkboxes = False
         End If
      Else
         Exit Sub
      End If
    Case UCase(gstrNovo)
      LimpaObjetos
    Case UCase(gstrFechar)
      Unload Me
  End Select
    
End Sub

Private Function blnDadosOk() As Boolean
Dim intAux As Integer
  blnDadosOk = False
  
  If Len(txtExercicio.Text) > 0 And Len(txtExercicio.Text) < 4 Then
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
     If Int(dbcstrAcordoFinal.Text) < Int(dbcstrAcordoInicial.Text) Then
        ExibeMensagem "O Acordo Inicial não pode ser maior que o Acordo Final."
        dbcstrAcordoFinal.SetFocus
        Exit Function
     End If
  End If
  
  If lvwParcelas.ListItems.Count = 0 Or lvwParcelas.ListItems.Item(1) = "Pressione F5 para preencher as parcelas." Then
     ExibeMensagem "A(s) parcela(s) devem ser selecionadas."
     lvwParcelas.SetFocus
     Exit Function
  End If
  
  For intAux = 1 To lvwParcelas.ListItems.Count
    If lvwParcelas.ListItems.Item(intAux).Checked = True Then
       Exit For
    Else
       If intAux = lvwParcelas.ListItems.Count Then
          ExibeMensagem "A(s) parcela(s) devem ser selecionadas."
          lvwParcelas.SetFocus
          Exit Function
       End If
    End If
  Next

  blnDadosOk = True
End Function

Private Function Exercicios(ByRef strInscricaoInicial As String, ByRef strInscricaoFinal As String) As Boolean
Dim strsql As String
Dim adoResultado As ADODB.Recordset
  
  If strInscricaoInicial <> "" Then
     strsql = "SELECT LA.intExercicio "
     strsql = strsql & "FROM " & gstrLancamentoAlfa & " LA, "
     strsql = strsql & gstrAcordo & " AC "
     strsql = strsql & "WHERE LA.strInscricao LIKE '" & String(gintLenInscricao - Len(strInscricaoInicial) - 4, "0") & strInscricaoInicial & "%' AND "
     strsql = strsql & "AC.intLancamentoAlfa = LA.pkID "
     strsql = strsql & "ORDER BY intExercicio "
  
     Set gobjBanco = New clsBanco
     Set adoResultado = New ADODB.Recordset
     If gobjBanco.CriaADO(strsql, 5, adoResultado) Then
        If Not adoResultado.EOF Then
           adoResultado.MoveFirst
           strInscricaoInicial = strInscricaoInicial & adoResultado("intExercicio")
        End If
     Else
        Exit Function
     End If
     Set adoResultado = Nothing
  End If
  
  If strInscricaoFinal <> "" Then
     strsql = "SELECT LA.intExercicio "
     strsql = strsql & "FROM " & gstrLancamentoAlfa & " LA, "
     strsql = strsql & gstrAcordo & " AC "
     strsql = strsql & "WHERE LA.strInscricao LIKE '" & String(gintLenInscricao - Len(strInscricaoFinal) - 4, "0") & strInscricaoFinal & "%' AND "
     strsql = strsql & "AC.intLancamentoAlfa = LA.pkID "
     strsql = strsql & "ORDER BY intExercicio "
    
     Set gobjBanco = New clsBanco
     Set adoResultado = New ADODB.Recordset
     If gobjBanco.CriaADO(strsql, 5, adoResultado) Then
        If Not adoResultado.EOF Then
           adoResultado.MoveLast
           strInscricaoFinal = strInscricaoFinal & adoResultado("intExercicio")
        End If
     Else
        Exit Function
     End If
     Set adoResultado = Nothing
  End If
  
  Exercicios = True
End Function

Private Sub chkTodas_Click()
  If chkTodas.Value = 1 Then
     dbcstrAcordoInicial.Enabled = False
     dbcstrAcordoFinal.Enabled = False
  Else
     dbcstrAcordoInicial.Enabled = True
     dbcstrAcordoFinal.Enabled = True
  End If
End Sub

Private Sub dbcstrAcordoFinal_Change()
  lvwParcelas.ListItems.Clear
  lvwParcelas.Checkboxes = False
  lvwParcelas.ListItems.Add , , "Pressione F5 para preencher as parcelas."
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

Private Sub dbcstrAcordoInicial_Change()
  lvwParcelas.ListItems.Clear
  lvwParcelas.Checkboxes = False
  lvwParcelas.ListItems.Add , , "Pressione F5 para preencher as parcelas."
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

Private Sub LimpaObjetos()
  chkTodas.Value = 0
  dbcstrAcordoFinal.Enabled = True
  dbcstrAcordoFinal.Text = ""
  dbcstrAcordoInicial.Enabled = True
  dbcstrAcordoInicial.Text = ""
  lvwParcelas.ListItems.Clear
  lvwParcelas.ListItems.Add , , "Pressione F5 para preencher as parcelas."
  lvwParcelas.Checkboxes = False
  txtExercicio.Text = Year(gstrDataDoSistema)
  txtExercicio.SetFocus
End Sub

Private Function strQueryComboAcordo(Inscricao As String) As String
Dim strsql As String

  strsql = ""
  strsql = strsql & "SELECT "
  strsql = strsql & "LA.pkID, "
  strsql = strsql & strSUBSTRING & "(LA.strInscricao, " & gintLenInscricao - gintRetornaTamanhoMascara(TYP_ACORDO) + 1 & ", " & gintRetornaTamanhoMascara(TYP_ACORDO) - 4 & ") strInscricao "
  strsql = strsql & "FROM "
  strsql = strsql & gstrLancamentoAlfa & " LA, "
  strsql = strsql & gstrAcordo & " AC "
  strsql = strsql & "WHERE "
  strsql = strsql & "LA.pkID = AC.intLancamentoAlfa "
  
  If Len(Trim(txtExercicio.Text)) = 4 Then
     strsql = strsql & "AND LA.intExercicio = " & txtExercicio.Text & " "
  End If
  
  If Inscricao <> "" Then
    If Len(Trim(txtExercicio.Text)) = 4 Then
        strsql = strsql & " AND LA.strInscricao = '" & String(gintLenInscricao - Len(Val(Inscricao) & txtExercicio.Text), "0") & Val(Inscricao) & txtExercicio.Text & "' "
    End If
  End If
  strsql = strsql & "ORDER BY strInscricao"

  strQueryComboAcordo = strsql
End Function

Private Sub Form_Load()
  txtExercicio.Text = Year(gstrDataDoSistema)
  lvwParcelas.ListItems.Add , , "Pressione F5 para preencher as parcelas."
  lvwParcelas.Checkboxes = False
End Sub

Private Sub lvwParcelas_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
Dim intCount As Integer
  For intCount = 1 To lvwParcelas.ListItems.Count
    If lvwParcelas.ListItems.Item(intCount).Checked = True Then
       lvwParcelas.ListItems.Item(intCount).Checked = False
    Else
       lvwParcelas.ListItems.Item(intCount).Checked = True
    End If
  Next
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

Private Function strParcelas() As String
Dim strParc As String
Dim intCont As Integer
  
  strParc = ""
  For intCont = 1 To lvwParcelas.ListItems.Count
      If lvwParcelas.ListItems.Item(intCont).Checked = True Then
          strParc = strParc & lvwParcelas.ListItems.Item(intCont).Text & ","
      End If
  
  Next
  strParc = Left(strParc, Len(strParc) - 1)
  strParcelas = strParc
End Function

Private Function PreencheParcelas() As Boolean
Dim strsql As String
Dim strInsc As String
Dim adoResultado As ADODB.Recordset
Dim blnExercicio As Boolean
Dim strInscricaoInicial As String
Dim strInscricaoFinal As String
  
  lvwParcelas.ListItems.Clear
  
  blnExercicio = IIf(Len(txtExercicio.Text) = 4, True, False)
  PreencheParcelas = False
  
  If Len(txtExercicio.Text) = 4 Then
     If dbcstrAcordoInicial.MatchedWithList Then
        strInscricaoInicial = dbcstrAcordoInicial.Text & txtExercicio.Text
     End If
     If dbcstrAcordoFinal.MatchedWithList Then
        strInscricaoFinal = dbcstrAcordoFinal.Text & txtExercicio.Text
     End If
  Else
     If dbcstrAcordoInicial.MatchedWithList Then
        strInscricaoInicial = dbcstrAcordoInicial.Text
     End If
     If dbcstrAcordoFinal.MatchedWithList Then
        strInscricaoFinal = dbcstrAcordoFinal.Text
     End If
     
     If dbcstrAcordoInicial.MatchedWithList And dbcstrAcordoFinal.MatchedWithList Then
        If Exercicios(strInscricaoInicial, strInscricaoFinal) = False Then 'Busca os exercícios das inscrições informadas
           ExibeMensagem "Erro na consulta dos exercícios das inscrições informadas. Não foi possível preencher as parcelas."
           Exit Function
        End If
        blnExercicio = True
     Else
        blnExercicio = False
     End If
  End If
  
  
  strInsc = ""
  If chkTodas.Value = 0 Then
     If strInscricaoInicial <> "" And strInscricaoFinal <> "" Then
           strInsc = "LA.strInscricao BETWEEN '" & String(gintLenInscricao - Len(strInscricaoInicial), "0") & strInscricaoInicial & "' AND '"
           strInsc = strInsc & String(gintLenInscricao - Len(strInscricaoFinal), "0") & strInscricaoFinal & "' "
     Else
        If strInscricaoInicial <> "" Then
           If blnExercicio Then
              strInsc = "LA.strInscricao = '" & String(gintLenInscricao - Len(strInscricaoInicial), "0") & strInscricaoInicial & "' "
           Else
              strInsc = "LA.strInscricao LIKE '" & String(gintLenInscricao - Len(strInscricaoInicial) - 4, "0") & strInscricaoInicial & "%' "
           End If
        Else
           If blnExercicio Then
              strInsc = "LA.strInscricao = '" & String(gintLenInscricao - Len(strInscricaoFinal), "0") & strInscricaoFinal & "' "
           Else
              strInsc = "LA.strInscricao LIKE '" & String(gintLenInscricao - Len(strInscricaoFinal) - 4, "0") & strInscricaoFinal & "%' "
           End If
        End If
     End If
  End If
  
  strsql = ""
  strsql = strsql & "SELECT DISTINCT "
  strsql = strsql & "LV.intParcela "
  strsql = strsql & "FROM "
  strsql = strsql & gstrLancamentoAlfa & " LA, "
  strsql = strsql & gstrLancamentoValor & " LV, "
  strsql = strsql & gstrAcordo & " AC "
  strsql = strsql & "WHERE "
  strsql = strsql & "LA.pkID = AC.intLancamentoAlfa AND "
  strsql = strsql & "LV.intLancamentoalfa = AC.intLancamentoAlfa AND "
  strsql = strsql & "LA.intUtilizacao = 4 "
  
  
  If strInsc <> "" Then
     strsql = strsql & " AND " & strInsc & " "
  End If
  
  strsql = strsql & "ORDER BY LV.intParcela "

  Set gobjBanco = New clsBanco
  Set adoResultado = New ADODB.Recordset
  If gobjBanco.CriaADO(strsql, 5, adoResultado) Then
     Do While Not adoResultado.EOF
        lvwParcelas.ListItems.Add , , adoResultado(0)
        adoResultado.MoveNext
     Loop
  End If
  Set adoResultado = Nothing

End Function
