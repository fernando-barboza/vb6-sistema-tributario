VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Begin VB.Form frmISSCarneSegundaVia 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ISS - Carnê 2ª Via"
   ClientHeight    =   5490
   ClientLeft      =   3960
   ClientTop       =   2790
   ClientWidth     =   5550
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5490
   ScaleWidth      =   5550
   Begin VB.Frame Frame1 
      Caption         =   "Parcelas"
      Height          =   2055
      Left            =   135
      TabIndex        =   16
      Top             =   3375
      Width           =   5280
      Begin MSComctlLib.ListView lvwParcelas 
         Height          =   1635
         Left            =   450
         TabIndex        =   8
         Top             =   270
         Width           =   4335
         _ExtentX        =   7646
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
      Caption         =   "Opções de Consulta"
      Height          =   2295
      Left            =   135
      TabIndex        =   11
      Top             =   990
      Width           =   5280
      Begin VB.TextBox txtNumeroAviso 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1755
         MaxLength       =   10
         TabIndex        =   6
         Top             =   1590
         Width           =   1200
      End
      Begin VB.CheckBox chkTodasInscricoes 
         Caption         =   "Selecionar todas as inscrições"
         Height          =   255
         Left            =   1755
         TabIndex        =   7
         Top             =   1935
         Width           =   2835
      End
      Begin VB.Frame Frame2 
         Height          =   555
         Left            =   495
         TabIndex        =   12
         Top             =   180
         Width           =   4305
         Begin VB.OptionButton optOpcao 
            Caption         =   "Faixa de Inscrição"
            CausesValidation=   0   'False
            Height          =   225
            Index           =   0
            Left            =   495
            TabIndex        =   2
            Top             =   225
            Value           =   -1  'True
            Width           =   1680
         End
         Begin VB.OptionButton optOpcao 
            Caption         =   "Emissao"
            CausesValidation=   0   'False
            Height          =   225
            Index           =   1
            Left            =   2745
            TabIndex        =   3
            Top             =   225
            Width           =   1005
         End
      End
      Begin MSDataListLib.DataCombo dbcstrInscricaoInicial 
         Height          =   315
         Left            =   1755
         TabIndex        =   4
         Top             =   810
         Width           =   3045
         _ExtentX        =   5371
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dbcstrInscricaoFinal 
         Height          =   315
         Left            =   1755
         TabIndex        =   5
         Top             =   1215
         Width           =   3045
         _ExtentX        =   5371
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dbcstrEmissao 
         Height          =   315
         Left            =   1755
         TabIndex        =   9
         Top             =   810
         Visible         =   0   'False
         Width           =   3045
         _ExtentX        =   5371
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin VB.CheckBox chkTodasEmissoes 
         Caption         =   "Selecionar todas as emissões"
         Height          =   255
         Left            =   1770
         TabIndex        =   10
         Top             =   1170
         Visible         =   0   'False
         Width           =   2430
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "N° do Aviso"
         Height          =   195
         Left            =   750
         TabIndex        =   19
         Top             =   1635
         Width           =   840
      End
      Begin VB.Label lblFinal 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Inscrição Final:"
         Height          =   195
         Left            =   555
         TabIndex        =   15
         Top             =   1305
         Width           =   1065
      End
      Begin VB.Label lblInicial 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Inscrição Inicial:"
         Height          =   195
         Left            =   555
         TabIndex        =   14
         Top             =   870
         Width           =   1140
      End
      Begin VB.Label lblEmissao 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Emissão:"
         Height          =   195
         Left            =   1065
         TabIndex        =   13
         Top             =   870
         Visible         =   0   'False
         Width           =   630
      End
   End
   Begin VB.TextBox txtExercicio 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1125
      MaxLength       =   4
      TabIndex        =   1
      Top             =   495
      Width           =   1200
   End
   Begin MSDataListLib.DataCombo dbcstrComposicao 
      Height          =   315
      Left            =   1125
      TabIndex        =   0
      Top             =   90
      Width           =   4305
      _ExtentX        =   7594
      _ExtentY        =   556
      _Version        =   393216
      IntegralHeight  =   0   'False
      Text            =   ""
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Composição:"
      Height          =   195
      Left            =   180
      TabIndex        =   18
      Top             =   135
      Width           =   915
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      Caption         =   "Exercício:"
      Height          =   195
      Left            =   180
      TabIndex        =   17
      Top             =   540
      Width           =   720
   End
End
Attribute VB_Name = "frmISSCarneSegundaVia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub MantemForm(ByVal strModoOperacao As String)

  Select Case UCase(strModoOperacao)
  
    Case Is = UCase(gstrPreencherLista)
      If Me.ActiveControl.Name = "dbcstrInscricaoInicial" Or Me.ActiveControl.Name = "dbcstrInscricaoFinal" Then
         If blnCompExerOK = False Then
            dbcstrInscricaoFinal.Text = ""
            dbcstrInscricaoInicial.Text = ""
            Set dbcstrInscricaoFinal.RowSource = Nothing
            Set dbcstrInscricaoInicial.RowSource = Nothing
            Exit Sub
         Else
            dbcstrInscricaoFinal.Tag = strQueryInscricao & ";strInscricao"
            dbcstrInscricaoInicial.Tag = strQueryInscricao & ";strInscricao"
         End If
      End If
      PreencherListaDeOpcoes Me.ActiveControl
      
    Case Is = UCase(gstrImprimir)
      If blnDadosOk = True Then
         rptCapaCarneISS.strParcelasSelecionadas = strParcelas
         rptCapaCarneISS.strEmpresaFebraban = strFebraban
         
         If Trim(txtNumeroAviso) <> "" Then
            rptCapaCarneISS.strNumeroAviso = txtNumeroAviso
         Else
            rptCapaCarneISS.strNumeroAviso = ""
         End If
         
         If strFebraban = "" Then
            ExibeMensagem "Não foi cadastrado o nº Febraban no módulo de Segurança."
            Exit Sub
         End If
         ImprimeRelatorio rptCapaCarneISS, strQueryRelatorio
      End If
      
    Case Is = UCase(gstrLocalizar)
      If blnCompExerOK = False Then Exit Sub
      
      lvwParcelas.Checkboxes = True
      PreencheParcelas
      If lvwParcelas.ListItems.Count <= 0 Then
         lvwParcelas.ListItems.Add , , "Pressione F5 para preencher as parcelas."
         lvwParcelas.Checkboxes = False
      End If
      
    Case UCase(gstrNovo)
      HabilitaFaixa 0
      LimpaObjetos
    
    Case UCase(gstrFechar)
      Unload Me
      
  End Select
    
End Sub

Private Function blnCompExerOK() As Boolean
  If Trim(txtExercicio.Text) = "" Then
     ExibeMensagem "O exercício deve ser informado."
     txtExercicio.SetFocus
     Exit Function
  End If
  If Len(txtExercicio.Text) < 4 Then
     ExibeMensagem "O exercício deve ser preenchido corretamente."
     txtExercicio.SetFocus
     Exit Function
  End If
  If dbcstrComposicao.MatchedWithList = False Then
     ExibeMensagem "A composição deve ser informada."
     dbcstrComposicao.SetFocus
     Exit Function
  End If
  blnCompExerOK = True
End Function

Private Function blnDadosOk() As Boolean
Dim intAux As Integer
  blnDadosOk = False
  If Trim(txtExercicio.Text) = "" Then
     ExibeMensagem "O exercício deve ser informado."
     txtExercicio.SetFocus
     Exit Function
  End If
  If Len(txtExercicio.Text) < 4 Then
     ExibeMensagem "Preencha o exercício corretamente."
     txtExercicio.SetFocus
     Exit Function
  End If
  If dbcstrComposicao.MatchedWithList = False Then
     ExibeMensagem "A composição deve ser informada."
     dbcstrComposicao.SetFocus
     Exit Function
  End If
  
  If optOpcao(0).Value = True Then
     If dbcstrInscricaoFinal.MatchedWithList = False And dbcstrInscricaoInicial.MatchedWithList = False And chkTodasInscricoes.Value = 0 Then
        ExibeMensagem "A inscrição deve ser informada."
        dbcstrInscricaoInicial.SetFocus
        Exit Function
     End If
     If dbcstrInscricaoInicial.MatchedWithList = True And dbcstrInscricaoFinal.MatchedWithList = True Then
        If Int(dbcstrInscricaoFinal.Text) < Int(dbcstrInscricaoInicial.Text) Then
           ExibeMensagem "A inscrição inicial não pode ser maior que a inscrição final."
           dbcstrInscricaoFinal.SetFocus
           Exit Function
        End If
     End If
  Else
     If dbcstrEmissao.MatchedWithList = False And chkTodasEmissoes.Value = True Then
        ExibeMensagem "A Emissão deve ser informada."
        dbcstrEmissao.SetFocus
        Exit Function
     End If
  End If
     
  If lvwParcelas.ListItems.Count = 0 Then
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

Private Sub chkTodasInscricoes_Click()
  If chkTodasInscricoes.Value = 1 Then
     dbcstrInscricaoInicial.Enabled = False
     dbcstrInscricaoFinal.Enabled = False
  Else
     dbcstrInscricaoInicial.Enabled = True
     dbcstrInscricaoFinal.Enabled = True
  End If
End Sub

Private Sub chkTodasEmissoes_Click()
  If chkTodasEmissoes.Value = 1 Then
     dbcstrEmissao.Enabled = False
  Else
     dbcstrEmissao.Enabled = True
  End If
End Sub

Private Sub LimpaObjetos()
  optOpcao(0).Value = True
  chkTodasEmissoes.Value = 0
  chkTodasInscricoes.Value = 0
  dbcstrEmissao.Text = ""
  dbcstrInscricaoFinal.Text = ""
  Set dbcstrInscricaoFinal.RowSource = Nothing
  dbcstrInscricaoInicial.Text = ""
  Set dbcstrInscricaoInicial.RowSource = Nothing
  txtExercicio.Text = Year(gstrDataDoSistema)
  lvwParcelas.ListItems.Clear
  lvwParcelas.ListItems.Add , , "Pressione F5 para preencher as parcelas."
  lvwParcelas.Checkboxes = False
  Set dbcstrComposicao.RowSource = Nothing
  dbcstrComposicao.Text = ""
  dbcstrComposicao.SetFocus
End Sub

Private Function strQueryInscricao()
Dim strSql As String
  strSql = ""
  strSql = strSql & "SELECT DISTINCT " & gintPkidFixo & ", "
  strSql = strSql & gstrRIGHT("strInscricao", gintRetornaTamanhoMascara(TYP_IMOBILIARIA)) & " strInscricao "
  strSql = strSql & "FROM " & gstrLancamentoAlfa & " "
  strSql = strSql & "WHERE "
  strSql = strSql & "intComposicaoDaReceita = " & dbcstrComposicao.BoundText & " AND "
  strSql = strSql & "intExercicio = " & txtExercicio.Text & " AND "
  strSql = strSql & "dtmdtCancelamento IS NULL "
  
  strSql = strSql & "ORDER BY strInscricao "
  
  strQueryInscricao = strSql
End Function

Private Function strQueryEmissao()
Dim strSql As String
  strSql = ""
  strSql = strSql & "SELECT DISTINCT " & gintPkidFixo & ", "
  strSql = strSql & "strEmissao "
  strSql = strSql & "FROM " & gstrLancamentoAlfa & " "
  strSql = strSql & "ORDER BY strEmissao "
  
  strQueryEmissao = strSql
End Function

Private Function strQueryComposicao() As String
Dim strSql As String
    
    strSql = "SELECT CO.Pkid,"
    strSql = strSql & gstrCONVERT(CDT_VARCHAR, "CO.intCodigo") & strCONCAT & "' - '" & strCONCAT & _
                      " RTRIM(LTRIM(CO.strDescricao)) Descricao "
    strSql = strSql & "FROM "
    strSql = strSql & gstrComposicaoDaReceita & " CO "
    strSql = strSql & "WHERE "
    strSql = strSql & "CO.Intutilizacao = " & TYP_ECONOMICA & " "
    strSql = strSql & "ORDER BY strDescricao "
    
    strQueryComposicao = strSql

End Function

Private Sub dbcstrComposicao_Change()
  dbcstrInscricaoFinal.Text = ""
  dbcstrInscricaoInicial.Text = ""
  Set dbcstrInscricaoFinal.RowSource = Nothing
  Set dbcstrInscricaoInicial.RowSource = Nothing
End Sub

Private Sub dbcstrInscricaoFinal_Click(Area As Integer)
  DropDownDataCombo dbcstrInscricaoFinal, Me, Area
End Sub

Private Sub dbcstrInscricaoFinal_GotFocus()
  If Trim(dbcstrInscricaoFinal.Text) = "" Then
     dbcstrInscricaoFinal.Text = Trim(dbcstrInscricaoInicial.Text)
  End If
  MarcaCampo dbcstrInscricaoFinal
End Sub

Private Sub dbcstrInscricaoInicial_Click(Area As Integer)
  DropDownDataCombo dbcstrInscricaoInicial, Me, Area
End Sub

Private Sub dbcstrInscricaoInicial_GotFocus()
  If Trim(dbcstrInscricaoInicial.Text) = "" Then
     dbcstrInscricaoInicial.Text = Trim(dbcstrInscricaoFinal.Text)
  End If
  MarcaCampo dbcstrInscricaoInicial
End Sub

Private Sub Form_Load()
  dbcstrEmissao.Tag = strQueryEmissao & ";strEmissao"
  dbcstrComposicao.Tag = strQueryComposicao & ";strDescricao"
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

Private Sub optOpcao_Click(Index As Integer)
  HabilitaFaixa Index
End Sub

Private Sub HabilitaFaixa(Opcao As Integer)
  If Opcao = 0 Then
     lblEmissao.Visible = False
     dbcstrEmissao.Visible = False
     chkTodasEmissoes.Visible = False
     
     lblInicial.Visible = True
     lblFinal.Visible = True
     dbcstrInscricaoInicial.Visible = True
     dbcstrInscricaoFinal.Visible = True
     chkTodasInscricoes.Visible = True
  Else
     lblInicial.Visible = False
     lblFinal.Visible = False
     dbcstrInscricaoInicial.Visible = False
     dbcstrInscricaoFinal.Visible = False
     chkTodasInscricoes.Visible = False
     
     lblEmissao.Visible = True
     dbcstrEmissao.Visible = True
     chkTodasEmissoes.Visible = True
  End If
End Sub

Private Sub txtExercicio_Change()
  dbcstrInscricaoFinal.Text = ""
  dbcstrInscricaoInicial.Text = ""
  Set dbcstrInscricaoFinal.RowSource = Nothing
  Set dbcstrInscricaoInicial.RowSource = Nothing
End Sub

Private Sub txtExercicio_KeyPress(KeyAscii As Integer)
  CaracterValido KeyAscii, "N", txtExercicio
End Sub

Private Function PreencheParcelas()
Dim strSql As String
Dim adoResultado As ADODB.Recordset
  
  lvwParcelas.ListItems.Clear
  
  
  'PARCELAS DO ECONOMICO
  strSql = ""
  strSql = strSql & "SELECT DISTINCT "
  strSql = strSql & "LV.intParcela "
  strSql = strSql & "FROM "
  strSql = strSql & gstrLancamentoValor & " LV, "
  strSql = strSql & gstrLancamentoAlfa & " LA "
  strSql = strSql & "WHERE "
  strSql = strSql & "LA.intComposicaoDaReceita = " & dbcstrComposicao.BoundText & " AND "
  strSql = strSql & "LA.intExercicio = " & Trim(txtExercicio.Text) & " AND "
  strSql = strSql & gstrCONVERT(CDT_INT, "LA.strEmissao") & " = 0 AND "
  strSql = strSql & "LV.intLancamentoAlfa = LA.pkID "
  'strSql = strSql & "LV.bitParcelaValida = 1 "
  
  strSql = strSql & "ORDER BY LV.intParcela "
      
  Set gobjBanco = New clsBanco
  Set adoResultado = New ADODB.Recordset
  If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
     Do While Not adoResultado.EOF
        lvwParcelas.ListItems.Add , , adoResultado(0)
        adoResultado.MoveNext
     Loop
  End If
  Set adoResultado = Nothing

End Function

Private Function strParcelas() As String
Dim strParc As String
Dim intCont As Integer
  
  strParc = ""
  For intCont = 1 To lvwParcelas.ListItems.Count
      If lvwParcelas.ListItems.Item(intCont).Checked = True Then
          strParc = strParc & lvwParcelas.ListItems.Item(intCont).Text & ","
      End If
  
  Next
  
  If Len(strParc) > 0 Then
    strParc = Left(strParc, Len(strParc) - 1)
    strParcelas = strParc
  Else
    strParcelas = "0"
  End If
  
End Function

Private Function strFebraban() As String
Dim adoFebraban As ADODB.Recordset
Dim strSql As String
  
  strSql = ""
  strSql = strSql & "SELECT intFebraban FROM " & gstrEmpresa & " "
  
  Set gobjBanco = New clsBanco
  Set adoFebraban = New ADODB.Recordset
  If gobjBanco.CriaADO(strSql, 5, adoFebraban) Then
     If Not adoFebraban.EOF Then
        strFebraban = adoFebraban(0)
     End If
  End If
  
  Set adoFebraban = Nothing
End Function

Private Function strQueryRelatorio()
Dim strSql As String
Dim strBarra As String
Dim strOpcao As String
Dim intCont As Integer
  
  strOpcao = ""
  If optOpcao(0).Value = True Then 'FAIXA DE INSCRIÇÕES
     If chkTodasInscricoes.Value = 0 Then
        If dbcstrInscricaoInicial.MatchedWithList = True And dbcstrInscricaoFinal.MatchedWithList = True Then
           strOpcao = "LA.strInscricao BETWEEN '" & String(gintLenInscricao - Len(Trim(dbcstrInscricaoInicial.Text)), "0") & Trim(dbcstrInscricaoInicial.Text) & "' AND '"
           strOpcao = strOpcao & String(gintLenInscricao - Len(Trim(dbcstrInscricaoFinal.Text)), "0") & Trim(dbcstrInscricaoFinal.Text) & "' AND "
           
        Else
           If dbcstrInscricaoInicial.MatchedWithList = True Then
              strOpcao = "LA.strInscricao = '" & String(gintLenInscricao - Len(Trim(dbcstrInscricaoInicial.Text)), "0") & Trim(dbcstrInscricaoInicial.Text) & "' AND "
           Else
              strOpcao = "LA.strInscricao = '" & String(gintLenInscricao - Len(Trim(dbcstrInscricaoFinal.Text)), "0") & Trim(dbcstrInscricaoFinal.Text) & "' AND "
           End If
        End If
     End If
  Else
     If chkTodasEmissoes.Value = 0 Then
        strOpcao = "LA.strEmissao = '" & String(gintLenEmissao - Len(Trim(dbcstrEmissao.Text)), "0") & Trim(dbcstrEmissao.Text) & "' AND "
     End If
  End If
  
  
  strSql = ""
    
  strSql = strSql & "SELECT "
      
  strSql = strSql & "LA.pkID pkIDPrincipal, "
  strSql = strSql & gstrRIGHT("LA.strInscricao", gintRetornaTamanhoMascara(TYP_ECONOMICA)) & " strInscricao, "
  strSql = strSql & "LA.intExercicio intExercicio, "
  strSql = strSql & "LA.strComposicaoDaReceita strComposicao, "
  strSql = strSql & "LA.intComposicaoDaReceita intComposicao, "
  strSql = strSql & "LA.strEmissao strEmissao, "
  strSql = strSql & "CR.strSigla strSigla, "
  strSql = strSql & "CR.intUtilizacao intUtilizacao, "
  strSql = strSql & gstrCONVERT(CDT_numeric, "LA.strNumeroAviso") & " strAviso, "
  strSql = strSql & "LA.strNomeProprietario strContribuinte, "
  strSql = strSql & "LA.strLogradouroC strLogradouroC, "
  strSql = strSql & "LA.strNumeroC strNumeroC, "
  strSql = strSql & "LA.strComplementoC strComplementoC, "
  strSql = strSql & "LA.strBairroC strBairroC, "
  strSql = strSql & "LA.strMunicipioC strMunicipioC, "
  strSql = strSql & "LA.strUFC strUFC, "
  strSql = strSql & "LA.intCEPC intCEPC, "
  strSql = strSql & "LA.strLogradouro strLogradouro, "
  strSql = strSql & "LA.strNumero strNumero, "
  strSql = strSql & "LA.strComplemento strComplemento, "
  strSql = strSql & "LA.strBairro strBairro, "
  strSql = strSql & "LA.strMunicipio strMunicipio, "
  strSql = strSql & "LA.strUF strUF, "
  strSql = strSql & "LA.intCEP intCEP, "
  strSql = strSql & "LA.strLogradouro strLogradouro, "
  strSql = strSql & "LA.strNumero strNumero, "
  strSql = strSql & "LA.strComplemento strComplemento, "
  strSql = strSql & "LA.strBairro strBairro, "
  strSql = strSql & "LA.strMunicipio strMunicipio, "
  strSql = strSql & "LA.strUF strUF, "
  strSql = strSql & "LA.intCEP intCEP, "
  strSql = strSql & "LA.strindexador , "
  strSql = strSql & "LA.dblvlIndexador dblvlIndexador, "
  
  strSql = strSql & "LV.intnumeroparcelas intNumeroParcelas, " '--NÚMERO DE PARCELAS intNumeroParcelas,-- "
  strSql = strSql & "LV.dblValorParcela dblValorParcela, " 'VALOR DA 1ª PARCELA dblValorParcela,-- "
  strSql = strSql & "LV.dtmdtVencimentoParcela dtmdtVencimentoParcela, " 'VENCIMENTO DA PARCELA UNICA dtmdtVencimentoParcela,-- "
   
  strSql = strSql & "(CASE WHEN LA.dblvlIndexador IS NULL OR LA.dblvlIndexador = 0 THEN '' END) dblFmpParcela "
   
  strSql = strSql & "FROM "
  strSql = strSql & gstrLancamentoAlfa & " LA, "
  strSql = strSql & gstrComposicaoDaReceita & " CR, "
  
  'LANÇAMENTO VALOR (PARCELA)
  strSql = strSql & "( "
  strSql = strSql & "select "
  strSql = strSql & "lv.pkid, "
  strSql = strSql & "lv.intlancamentoalfa, "
  strSql = strSql & "lv.intparcela, "
  strSql = strSql & "lv.dtmdtvencimento dtmdtVencimentoParcela, "
  strSql = strSql & "lv.dblvalor dblValorParcela, "
  strSql = strSql & "lvg.intnumeroparcelas intnumeroparcelas "
  strSql = strSql & "from "
  strSql = strSql & "( "
  strSql = strSql & "select "
  strSql = strSql & "lv.intlancamentoalfa intlancamentoalfa, "
  strSql = strSql & "min(lv.intparcela) intparcela, "
  strSql = strSql & "count(*) intnumeroparcelas "
  strSql = strSql & "from "
  strSql = strSql & gstrLancamentoValor & " lv, "
  strSql = strSql & gstrLancamentoAlfa & " LA "
  strSql = strSql & "where "
  strSql = strSql & "lv.bitparcelavalida  = 1 AND "
  strSql = strSql & strOpcao
  strSql = strSql & "LV.intLancamentoAlfa IN (LA.pkID) "
  strSql = strSql & "group by "
  strSql = strSql & "lv.intlancamentoalfa "
  strSql = strSql & ") lvg, "
  strSql = strSql & gstrLancamentoValor & " lv, "
  strSql = strSql & gstrLancamentoAlfa & " LA "
  strSql = strSql & "where "
  strSql = strSql & "lv.intparcela        > 0 and "
  strSql = strSql & "lv.intlancamentoalfa = lvg.intlancamentoalfa and "
  strSql = strSql & "lv.intparcela        = lvg.intparcela AND "
  strSql = strSql & "lv.bitparcelavalida  = 1 AND "
  strSql = strSql & strOpcao
  strSql = strSql & "LV.intLancamentoAlfa IN (LA.pkID) "
  strSql = strSql & ") lv, "
  
 
  'VERIFICA SE CONTÉM A PARCELA SELECIONADA
  strSql = strSql & "( "
  strSql = strSql & "SELECT "
  strSql = strSql & "LV.intLancamentoAlfa "
  strSql = strSql & "FROM "
  strSql = strSql & gstrLancamentoValor & " LV, "
  strSql = strSql & gstrLancamentoAlfa & " LA "
  strSql = strSql & "WHERE "
  strSql = strSql & "LV.intParcela IN (" & strParcelas & ") AND "
  strSql = strSql & strOpcao
  strSql = strSql & "LV.intLancamentoAlfa IN (LA.pkID) "
  strSql = strSql & "GROUP BY "
  strSql = strSql & "LV.intLancamentoAlfa "
  strSql = strSql & ") LVG "
  
  strSql = strSql & "WHERE "
    
  If strOpcao <> "" Then 'INSCRIÇÃO OU EMISSÃO
     strSql = strSql & strOpcao
  End If

  strSql = strSql & "LA.dtmdtCancelamento IS NULL AND "
  strSql = strSql & "LV.intLancamentoAlfa " & strOUTJOracle & "=" & strOUTJSQLServer & " LA.pkID AND "
  strSql = strSql & "LA.intComposicaoDaReceita = " & dbcstrComposicao.BoundText & " AND "
  strSql = strSql & "CR.pkID " & strOUTJOracle & "=" & strOUTJSQLServer & " LA.intComposicaoDaReceita AND "
  strSql = strSql & "LA.intExercicio = " & Trim(txtExercicio.Text) & " AND "
  
  If Trim(txtNumeroAviso) <> "" Then
     strSql = strSql & gstrCONVERT(CDT_numeric, "LA.Strnumeroaviso") & " = " & Val(Trim(txtNumeroAviso.Text)) & " AND "
  End If
  
  strSql = strSql & "LVG.intLancamentoAlfa = LA.pkID "
  strSql = strSql & "ORDER BY LA.strInscricao, "
  strSql = strSql & "LV.intParcela "
  
  strQueryRelatorio = strSql
End Function

Private Sub txtNumeroAviso_GotFocus()
    MarcaCampo txtNumeroAviso
End Sub

Private Sub txtNumeroAviso_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtNumeroAviso
End Sub
