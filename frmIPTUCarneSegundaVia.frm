VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Begin VB.Form frmIPTUCarneSegundaVia 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "IPTU - Carnê 2ª Via"
   ClientHeight    =   5175
   ClientLeft      =   4635
   ClientTop       =   2565
   ClientWidth     =   5565
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   5565
   Begin VB.TextBox txtExercicio 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1125
      MaxLength       =   4
      TabIndex        =   1
      Top             =   495
      Width           =   1200
   End
   Begin VB.Frame fra_Mensagem1 
      Caption         =   "Opções de Consulta"
      Height          =   1935
      Left            =   135
      TabIndex        =   8
      Top             =   990
      Width           =   5280
      Begin VB.Frame Frame2 
         Height          =   555
         Left            =   495
         TabIndex        =   14
         Top             =   180
         Width           =   4305
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
      End
      Begin VB.CheckBox chkTodasInscricoes 
         Caption         =   "Selecionar todas as inscrições"
         Height          =   255
         Left            =   1755
         TabIndex        =   6
         Top             =   1575
         Width           =   2835
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
         Top             =   1210
         Width           =   3045
         _ExtentX        =   5371
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dbcstrEmissao 
         Height          =   315
         Left            =   1755
         TabIndex        =   15
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
         TabIndex        =   17
         Top             =   1170
         Visible         =   0   'False
         Width           =   2430
      End
      Begin VB.Label lblEmissao 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Emissão:"
         Height          =   195
         Left            =   1065
         TabIndex        =   16
         Top             =   870
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.Label lblInicial 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Inscrição Inicial:"
         Height          =   195
         Left            =   555
         TabIndex        =   10
         Top             =   870
         Width           =   1140
      End
      Begin VB.Label lblFinal 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Inscrição Final:"
         Height          =   195
         Left            =   555
         TabIndex        =   9
         Top             =   1305
         Width           =   1065
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Parcelas"
      Height          =   2055
      Left            =   135
      TabIndex        =   13
      Top             =   2955
      Width           =   5280
      Begin MSComctlLib.ListView lvwParcelas 
         Height          =   1635
         Left            =   450
         TabIndex        =   7
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
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      Caption         =   "Exercício:"
      Height          =   195
      Left            =   180
      TabIndex        =   12
      Top             =   540
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Composição:"
      Height          =   195
      Left            =   180
      TabIndex        =   11
      Top             =   135
      Width           =   915
   End
End
Attribute VB_Name = "frmIPTUCarneSegundaVia"
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
      If blnDadosOK = True Then
         rptCapaCarneIPTU.strParcelasSelecionadas = strParcelas
         rptCapaCarneIPTU.strEmpresaFebraban = strFebraban
         If strFebraban = "" Then
            ExibeMensagem "Não foi cadastrado o nº Febraban no módulo de Segurança."
            Exit Sub
         End If
         ImprimeRelatorio rptCapaCarneIPTU, strQueryRelatorio
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

Private Function blnDadosOK() As Boolean
Dim intAux As Integer
  blnDadosOK = False
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
     If dbcstrEmissao.MatchedWithList = False Then
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
     
  blnDadosOK = True
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
Dim strSQL As String
  strSQL = ""
  strSQL = strSQL & "SELECT DISTINCT " & gintPkidFixo & ", "
  strSQL = strSQL & gstrRIGHT("strInscricao", gintRetornaTamanhoMascara(TYP_IMOBILIARIA)) & " strInscricao "
  strSQL = strSQL & "FROM " & gstrLancamentoAlfa & " "
  strSQL = strSQL & "WHERE "
  strSQL = strSQL & "intComposicaoDaReceita = " & dbcstrComposicao.BoundText & " AND "
  strSQL = strSQL & "intExercicio = " & txtExercicio.Text & " AND "
  strSQL = strSQL & "dtmdtCancelamento IS NULL "
  
  strSQL = strSQL & "ORDER BY strInscricao "
  
  strQueryInscricao = strSQL
End Function

Private Function strQueryEmissao()
Dim strSQL As String
  strSQL = ""
  strSQL = strSQL & "SELECT DISTINCT " & gintPkidFixo & ", strEmissao "
  strSQL = strSQL & "FROM " & gstrLancamentoAlfa & " "
  strSQL = strSQL & "ORDER BY strEmissao "
  
  strQueryEmissao = strSQL
End Function

Private Function strQueryComposicao() As String

Dim strSQL As String
    
    strSQL = "SELECT CO.Pkid, "
    strSQL = strSQL & gstrCONVERT(CDT_VARCHAR, "CO.intCodigo") & strCONCAT & "' - '" & strCONCAT & _
                      " RTRIM(LTRIM(CO.strDescricao)) Descricao "
    strSQL = strSQL & " FROM "
    strSQL = strSQL & gstrComposicaoDaReceita & " CO "
    strSQL = strSQL & " WHERE "
    strSQL = strSQL & " intUtilizacao IN (1,7) "
    strSQL = strSQL & " ORDER BY strDescricao "
    
    strQueryComposicao = strSQL

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
     dbcstrInscricaoFinal_Click 0
     dbcstrInscricaoFinal.Text = Trim(dbcstrInscricaoInicial.Text)
     MarcaCampo dbcstrInscricaoFinal
  End If
End Sub

Private Sub dbcstrInscricaoInicial_Click(Area As Integer)
  DropDownDataCombo dbcstrInscricaoInicial, Me, Area
End Sub

Private Sub dbcstrInscricaoInicial_GotFocus()
  If Trim(dbcstrInscricaoInicial.Text) = "" Then
     dbcstrInscricaoInicial.Text = Trim(dbcstrInscricaoFinal.Text)
     dbcstrInscricaoInicial_Click 0
     dbcstrInscricaoInicial.Text = Trim(dbcstrInscricaoFinal.Text)
     MarcaCampo dbcstrInscricaoInicial
  End If
End Sub

Private Sub Form_Load()
  txtExercicio.Text = Year(gstrDataDoSistema)
  dbcstrEmissao.Tag = strQueryEmissao & ";strEmissao"
  dbcstrComposicao.Tag = strQueryComposicao & ";strDescricao"
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

Private Sub txtExercicio_GotFocus()
    MarcaCampo txtExercicio
End Sub

Private Sub txtExercicio_KeyPress(KeyAscii As Integer)
  CaracterValido KeyAscii, "N", txtExercicio
End Sub

Private Sub txtExercicio_Change()
  dbcstrInscricaoFinal.Text = ""
  dbcstrInscricaoInicial.Text = ""
  Set dbcstrInscricaoFinal.RowSource = Nothing
  Set dbcstrInscricaoInicial.RowSource = Nothing
End Sub

Private Function PreencheParcelas()
Dim strSQL As String
Dim adoResultado As ADODB.Recordset
  
  lvwParcelas.ListItems.Clear
  
  strSQL = strSQL & "SELECT DISTINCT "
  strSQL = strSQL & "FPV.intParcela "
  strSQL = strSQL & "FROM "
  strSQL = strSQL & gstrParametroIPTU & " PI, "
  strSQL = strSQL & gstrParametroIPTUPagto & " PIP, "
  strSQL = strSQL & gstrFormaPagtoVencimentos & " FPV "
  strSQL = strSQL & "WHERE "
  strSQL = strSQL & "PI.strEmissao = '000' AND "
  
  strSQL = strSQL & "PI.intComposicaoDaReceita = " & dbcstrComposicao.BoundText & " AND "
  strSQL = strSQL & "PI.intExercicio = " & Trim(txtExercicio.Text) & " AND "
  strSQL = strSQL & "PIP.intParametroIptu = PI.pkID AND "
'  strSql = strSql & "PIP.bytParcelado = 1 AND "
  strSQL = strSQL & "FPV.intFormaPagto = PIP.pkID "
  strSQL = strSQL & "ORDER BY FPV.intParcela "
  
  Set gobjBanco = New clsBanco
  Set adoResultado = New ADODB.Recordset
  If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
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
Dim strSQL As String
  
  strSQL = ""
  strSQL = strSQL & "SELECT intFebraban FROM " & gstrEmpresa & " "
  
  Set gobjBanco = New clsBanco
  Set adoFebraban = New ADODB.Recordset
  If gobjBanco.CriaADO(strSQL, 5, adoFebraban) Then
     If Not adoFebraban.EOF Then
        strFebraban = Trim(gstrENulo(adoFebraban(0)))
     End If
  End If
  
End Function

Private Function strQueryRelatorio()
Dim strSQL As String
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
  
  strSQL = ""
    
  strSQL = strSQL & "SELECT "
  strSQL = strSQL & gstrRIGHT("LA.strInscricao", gintRetornaTamanhoMascara(TYP_IMOBILIARIA)) & " strImovel, "
  strSQL = strSQL & "LA.intExercicio intExercicio, "
  strSQL = strSQL & "LA.strComposicaoDaReceita strComposicao, "
  strSQL = strSQL & "LA.intComposicaoDaReceita intComposicao, "
  strSQL = strSQL & "LA.strEmissao strEmissao, "
  strSQL = strSQL & "CR.strSigla strSigla, "
  strSQL = strSQL & "CR.intUtilizacao intUtilizacao, "
  strSQL = strSQL & gstrCONVERT(CDT_numeric, "LA.strNumeroAviso") & " strAviso, "
  strSQL = strSQL & "LA.strNomeProprietario strProprietario, "
  strSQL = strSQL & "LA.strinscricaoauxiliar, "
  strSQL = strSQL & "LA.strPromissario strPromissario, "
  strSQL = strSQL & "LA.strLogradouroC strLogradouroC, "
  strSQL = strSQL & "LA.strNumeroC strNumeroC, "
  strSQL = strSQL & "LA.strComplementoC strComplementoC, "
  strSQL = strSQL & "LA.strBairroC strBairroC, "
  strSQL = strSQL & "LA.strMunicipioC strMunicipioC, "
  strSQL = strSQL & "LA.strUFC strUFC, "
  strSQL = strSQL & "LA.intCEPC intCEPC, "
  strSQL = strSQL & "LA.strLogradouro strLogradouro, "
  strSQL = strSQL & "LA.strNumero strNumero, "
  strSQL = strSQL & "LA.strComplemento strComplemento, "
  strSQL = strSQL & "LA.strBairro strBairro, "
  strSQL = strSQL & "LA.strMunicipio strMunicipio, "
  strSQL = strSQL & "LA.strUF strUF, "
  strSQL = strSQL & "LA.intCEP intCEP, "
  strSQL = strSQL & "LA.strLogradouro strLogradouro, "
  strSQL = strSQL & "LA.strNumero strNumero, "
  strSQL = strSQL & "LA.strComplemento strComplemento, "
  strSQL = strSQL & "LA.strBairro strBairro, "
  strSQL = strSQL & "LA.strMunicipio strMunicipio, "
  strSQL = strSQL & "LA.strUF strUF, "
  strSQL = strSQL & "LA.intCEP intCEP, "
  strSQL = strSQL & gstrISNULL("LA.dblValorTotal", "0") & " dblValorIntegral, "
  strSQL = strSQL & gstrISNULL("LA.dblporcdesconto", "0") & " dblDescontoIntegral, " 'hugo
  strSQL = strSQL & gstrISNULL("LVU.dblValorParcela", "0") & " dblValorTotal, "
  
  strSQL = strSQL & "LI.strLote strLote, "
  strSQL = strSQL & "LI.strQuadra strQuadra, "
  strSQL = strSQL & "LI.strLoteamento strLoteamento, "
  strSQL = strSQL & "LI.dblAreaTerreno dblAreaTerreno, "
  strSQL = strSQL & "LI.dblValorMetro dblValorMetro, "
  strSQL = strSQL & "LI.dblAreaExcedente dblAreaExcedente, "
  strSQL = strSQL & "LI.dblValorTerrenoExcedente dblValorExcedente, "
  strSQL = strSQL & "LI.dblAreaTerreno + LI.dblAreaExcedente dblTotalArea, "
  strSQL = strSQL & "LI.dblValorVenalTerreno dblValorVenalTerreno, "
  strSQL = strSQL & "LI.dblValorTerrenoExcedente dblValorTerrenoExcedente, "
  strSQL = strSQL & "LI.dblValorVenalTerreno  + LI.dblValorTerrenoExcedente dblTotalExcedente, "
  strSQL = strSQL & "LI.dblImpostoTerreno dblImpostoTerreno, "
  strSQL = strSQL & "LI.dblImpostoExcedente dblImpostoExcedente, "
  strSQL = strSQL & "LI.dblImpostoTerreno + LI.dblImpostoExcedente dblTotalImposto, "
  
  strSQL = strSQL & "LP.StrNomePadrao strPadrao, "
  strSQL = strSQL & "LP.StrNomeUso strUso, "
  strSQL = strSQL & "LP.dblTotalAreaPredio dblTotalAreaPredio , "
  strSQL = strSQL & "LP.dblTotalValorVenal dblTotalValorVenal, "
  strSQL = strSQL & "LP.dblFracaoIdeal dblFracaoIdeal, "
  strSQL = strSQL & "LP.dblTotalImpostoPredio dblTotalImpostoPredio, "
  strSQL = strSQL & "LP.dblValorVenalPredial + (LI.dblValorVenalTerreno + (CASE WHEN LI.dblValorTerrenoExcedente IS NULL THEN 0 END)) dblValorTotalImovel, "
  strSQL = strSQL & "LP.dblValorImpostoPredial + (LI.dblImpostoTerreno + (CASE WHEN LI.dblImpostoExcedente IS NULL THEN 0 END)) dblValorTotalImposto, "
  
  strSQL = strSQL & "LV.intnumeroparcelas intNumeroParcelas, " '--NÚMERO DE PARCELAS intNumeroParcelas,-- "
  strSQL = strSQL & "LV.Dblvalorparcela dblValorParcela, " 'VALOR DA 1ª PARCELA dblValorParcela,-- "

  If bytDBType = Oracle Then
    strSQL = strSQL & "TO_CHAR(LV.dtmdtVencimentoParcela,'dd/mm/yyyy') dtmdtVencimentoParcela, " 'VENCIMENTO DA PARCELA UNICA dtmdtVencimentoParcela,-- "
    strSQL = strSQL & "TO_CHAR(LVU.dtmdtVencimentoParcela,'dd/mm/yyyy') dtmdtVencimentoIntegral " 'VENCIMENTO DA PARCELA INTEGRAL dtmdtVencimentoIntegral,-- "
  Else
    strSQL = strSQL & "convert(varchar,LV.dtmdtVencimentoParcela,103) dtmdtVencimentoParcela, " 'VENCIMENTO DA PARCELA UNICA dtmdtVencimentoParcela,-- "
    strSQL = strSQL & "convert(varchar,LVU.dtmdtVencimentoParcela,103) dtmdtVencimentoIntegral " 'VENCIMENTO DA PARCELA INTEGRAL dtmdtVencimentoIntegral,-- "
  End If

  'Valor da única pegando da Lancto Alfa
  'strSql = strSql & "LVU.Dblvalorparcela dblValorIntegral, " ' --VALOR INTEGRAL dblValorIntegral,-- "
  
  strSQL = strSQL & "FROM "
  strSQL = strSQL & gstrLancamentoAlfa & " LA, "
  strSQL = strSQL & gstrLancamentoIPTU & " LI, "
  strSQL = strSQL & gstrComposicaoDaReceita & " CR, "
  
  'LANÇAMENTO PRÉDIO
  strSQL = strSQL & "( "
  strSQL = strSQL & "select "
  strSQL = strSQL & "min(lp.pkid) pkid, "
  strSQL = strSQL & "lp.intlancamentoiptu, "
  strSQL = strSQL & "min(lp.strnomepadrao) strnomepadrao, "
  strSQL = strSQL & "min(lp.strnomeuso) strnomeuso, "
  strSQL = strSQL & "SUM(LP.dblMedidaDaArea) dblTotalAreaPredio, "
  strSQL = strSQL & "SUM(LP.dblValorVenalPredio) dblTotalValorVenal, "
  strSQL = strSQL & "CASE WHEN Sum(LP.dblFracaoIdeal)<=1 THEN Sum(LP.dblFracaoIdeal) ELSE NULL END dblFracaoIdeal, "
  strSQL = strSQL & "SUM(LP.dblImposto) dblTotalImpostoPredio, "
  strSQL = strSQL & "SUM(LP.dblValorVenalPredio) dblValorVenalPredial, "
  strSQL = strSQL & "SUM(LP.dblImposto) dblValorImpostoPredial "
  strSQL = strSQL & "from "
  strSQL = strSQL & gstrLancamentoPredioIPTU & " lp, "
  strSQL = strSQL & gstrLancamentoAlfa & " LA, "
  strSQL = strSQL & gstrLancamentoIPTU & " LI "
  strSQL = strSQL & "WHERE "
  strSQL = strSQL & strOpcao
  strSQL = strSQL & "LI.intLancamentoAlfa = LA.pkID AND "
  strSQL = strSQL & "LA.dtmdtCancelamento IS NULL AND "
  strSQL = strSQL & "LA.intComposicaoDaReceita = " & dbcstrComposicao.BoundText & " AND "
  strSQL = strSQL & "LA.intExercicio = " & Trim(txtExercicio.Text) & " AND "
  strSQL = strSQL & "LP.intLancamentoIptu = LI.pkID "
  strSQL = strSQL & "group by "
  strSQL = strSQL & "lp.intlancamentoiptu "
  strSQL = strSQL & ") lp, "
  
  'LANÇAMENTO VALOR (PARCELA)
  strSQL = strSQL & "( "
  strSQL = strSQL & "select "
  strSQL = strSQL & "lv.pkid, "
  strSQL = strSQL & "lv.intlancamentoalfa, "
  strSQL = strSQL & "lv.intparcela, "
  strSQL = strSQL & "lv.dtmdtvencimento dtmdtVencimentoParcela, "
  strSQL = strSQL & "lv.dblvalor dblValorParcela, "
  strSQL = strSQL & "lvg.intnumeroparcelas intnumeroparcelas, "
  strSQL = strSQL & "lvg.dblValor "
  strSQL = strSQL & "from "
  strSQL = strSQL & "( "
  strSQL = strSQL & "select "
  strSQL = strSQL & "sum(" & gstrISNULL("Lv.Dblvalor", "0") & ") Dblvalor, "
  strSQL = strSQL & "lv.intlancamentoalfa intlancamentoalfa, "
  strSQL = strSQL & "min(lv.intparcela) intparcela, "
  strSQL = strSQL & "count(*) intnumeroparcelas "
  strSQL = strSQL & "from "
  strSQL = strSQL & gstrLancamentoValor & " LV, "
  strSQL = strSQL & gstrLancamentoAlfa & " LA "
  strSQL = strSQL & "where "
  strSQL = strSQL & "lv.bitparcelavalida  = 1 AND "
  strSQL = strSQL & strOpcao
  strSQL = strSQL & "LV.intLancamentoAlfa = LA.pkID AND "
  strSQL = strSQL & "LA.dtmdtCancelamento IS NULL AND "
  strSQL = strSQL & "LA.intComposicaoDaReceita = " & dbcstrComposicao.BoundText & " AND "
  strSQL = strSQL & "LA.intExercicio = " & Trim(txtExercicio.Text) & " AND "
  strSQL = strSQL & "LV.intParcela IN (" & strParcelas & ") "
  strSQL = strSQL & "GROUP BY "
  strSQL = strSQL & "LV.intlancamentoalfa "
  strSQL = strSQL & ") lvg, "
  strSQL = strSQL & gstrLancamentoValor & " lv, "
  strSQL = strSQL & gstrLancamentoAlfa & " LA "
  strSQL = strSQL & "where "
  strSQL = strSQL & "lv.intparcela        > 0 and "
  strSQL = strSQL & "lv.intlancamentoalfa = lvg.intlancamentoalfa and "
  strSQL = strSQL & "lv.intparcela        = lvg.intparcela AND "
  strSQL = strSQL & "lv.bitparcelavalida  = 1 AND "
  strSQL = strSQL & strOpcao
  strSQL = strSQL & "LV.intLancamentoAlfa = LA.pkID AND "
  strSQL = strSQL & "LA.dtmdtCancelamento IS NULL AND "
  strSQL = strSQL & "LA.intComposicaoDaReceita = " & dbcstrComposicao.BoundText & " AND "
  strSQL = strSQL & "LA.intExercicio = " & Trim(txtExercicio.Text) & " "
  strSQL = strSQL & ") lv, "
  
  'VERIFICA SE CONTÉM A PARCELA SELECIONADA
  'strSql = strSql & "( "
  'strSql = strSql & "SELECT "
  'strSql = strSql & "LV.intLancamentoAlfa "
  'strSql = strSql & "FROM "
  'strSql = strSql & gstrLancamentoValor & " LV, "
  'strSql = strSql & gstrLancamentoAlfa & " LA "
  'strSql = strSql & "WHERE "
  'strSql = strSql & "LV.intParcela IN (" & strParcelas & ") AND "
  'strSql = strSql & strOpcao
  'strSql = strSql & "LV.intLancamentoAlfa IN (LA.pkID) "
  'strSql = strSql & "GROUP BY "
  'strSql = strSql & "LV.intLancamentoAlfa "
  'strSql = strSql & ") LVG, "
  
  'LANÇAMENTO VALOR (PARCELA ÚNICA)
  strSQL = strSQL & "( "
  strSQL = strSQL & "select "
  strSQL = strSQL & "lv.pkid, "
  strSQL = strSQL & "lv.intlancamentoalfa, "
  strSQL = strSQL & "lv.intparcela, "
  strSQL = strSQL & "lv.dtmdtvencimento dtmdtVencimentoParcela, "
  strSQL = strSQL & "lv.dblvalor dblValorParcela "
  strSQL = strSQL & "from "
  strSQL = strSQL & "( "
  strSQL = strSQL & "select "
  strSQL = strSQL & "lv.intlancamentoalfa intlancamentoalfa, "
  strSQL = strSQL & "max(lv.intparcela) intparcela "
  strSQL = strSQL & "from "
  strSQL = strSQL & gstrLancamentoValor & " lv, "
  strSQL = strSQL & gstrLancamentoAlfa & " LA "
  strSQL = strSQL & "where "
  strSQL = strSQL & "lv.bitparcelavalida = 0 AND "
  strSQL = strSQL & strOpcao
  strSQL = strSQL & "LV.intLancamentoAlfa = LA.pkID AND "
  strSQL = strSQL & "LA.dtmdtCancelamento IS NULL AND "
  strSQL = strSQL & "LA.intComposicaoDaReceita = " & dbcstrComposicao.BoundText & " AND "
  strSQL = strSQL & "LA.intExercicio = " & Trim(txtExercicio.Text) & " "
  strSQL = strSQL & "group by "
  strSQL = strSQL & "lv.intlancamentoalfa "
  strSQL = strSQL & ") lvu, "
  strSQL = strSQL & gstrLancamentoValor & " lv, "
  strSQL = strSQL & gstrLancamentoAlfa & " LA "
  strSQL = strSQL & "where "
  strSQL = strSQL & "lv.intlancamentoalfa = lvu.intlancamentoalfa and "
  strSQL = strSQL & "lv.intparcela        = lvu.intparcela AND "
  strSQL = strSQL & "lv.bitparcelavalida  = 0 AND "
  strSQL = strSQL & strOpcao
  strSQL = strSQL & "LV.intLancamentoAlfa = LA.pkID AND "
  strSQL = strSQL & "LA.dtmdtCancelamento IS NULL AND "
  strSQL = strSQL & "LA.intComposicaoDaReceita = " & dbcstrComposicao.BoundText & " AND "
  strSQL = strSQL & "LA.intExercicio = " & Trim(txtExercicio.Text) & " "
  strSQL = strSQL & ") lvu "
  
  strSQL = strSQL & "WHERE "

  strSQL = strSQL & "LI.intLancamentoAlfa = LA.Pkid AND "
  strSQL = strSQL & "LP.intLancamentoIptu " & strOUTJOracle & "=" & strOUTJSQLServer & " LI.Pkid AND "
  
  If strOpcao <> "" Then 'INSCRIÇÃO OU EMISSÃO
     strSQL = strSQL & strOpcao & " "
  End If

  strSQL = strSQL & "LA.dtmdtCancelamento IS NULL AND "
  strSQL = strSQL & "LV.intLancamentoAlfa " & strOUTJOracle & "=" & strOUTJSQLServer & " LA.pkID AND "
  strSQL = strSQL & "LVU.intLancamentoAlfa " & strOUTJOracle & "=" & strOUTJSQLServer & " LA.pkID AND "
  strSQL = strSQL & "LA.intComposicaoDaReceita = " & dbcstrComposicao.BoundText & " AND "
  strSQL = strSQL & "CR.pkID " & strOUTJOracle & "=" & strOUTJSQLServer & " LA.intComposicaoDaReceita AND "
  strSQL = strSQL & "LA.intExercicio = " & Trim(txtExercicio.Text) & " "
  strSQL = strSQL & "ORDER BY "
  strSQL = strSQL & gstrCONVERT(CDT_numeric, "LA.strInscricao") & " "
  
  strQueryRelatorio = strSQL
End Function
