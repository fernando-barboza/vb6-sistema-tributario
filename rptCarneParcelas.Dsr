VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptCarneParcelas 
   Caption         =   "rptCarneParcelas (ActiveReport)"
   ClientHeight    =   10815
   ClientLeft      =   0
   ClientTop       =   285
   ClientWidth     =   15240
   MDIChild        =   -1  'True
   _ExtentX        =   26882
   _ExtentY        =   19076
   SectionData     =   "rptCarneParcelas.dsx":0000
End
Attribute VB_Name = "rptCarneParcelas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit

Dim intCont As Byte
Public blnPrimeira              As Boolean 'Não coloca margem superior na 1ª página
Public blnValorEmReal           As Boolean 'Identifica se o valor do boleto esta em Real
Public blnParcelasAtualizadas   As Boolean

Private Sub ActiveReport_Activate()
    If adoDataControl.Recordset.RecordCount = 0 Then
       ExibeMensagem "Não exite nenhuma parcela informada nas inscrições selecionadas."
       Unload Me
       Exit Sub
    End If
    HabilitaDesabilitaBotao1 False, gstrBtnArquivo, True
End Sub

Private Sub ActiveReport_Deactivate()
    HabilitaDesabilitaBotao1 True, gstrBtnArquivo, True
End Sub

Private Sub ActiveReport_QueryClose(Cancel As Integer, CloseMode As Integer)
    HabilitaDesabilitaBotao1 True, gstrBtnArquivo, True
End Sub

Private Sub ActiveReport_ReportStart()
    PadronizaToolBarRelatorio Me
    LeImagemLogotipo imgBrasao, imgLogotipo, txtNomeFantasia
    LeImagemLogotipo imgBrasao, imgLogotipo, txtNomeFantasia1
End Sub

Private Sub ActiveReport_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub ActiveReport_ToolbarClick(ByVal Tool As DDActiveReports2.DDTool)
    Dim vnt As Variant
    If Tool.ID = 14 Then
        ActiveReport_KeyPress 27
    ElseIf Tool.ID = 15 Then
        AbreOpcoesExportacao Me
    ElseIf Tool.ID = 16 Then
        Configura_Relatorio Me, True
    End If
End Sub

Private Sub Detail_Format()
Dim strCodigo As String
Dim strsql    As String
Dim adoResultado As ADODB.Recordset
Dim adoCommand   As ADODB.Command
Dim lngNumeroGuia   As Long
Dim dblValorReal    As Variant
Dim dblValorMult    As Double

'Detail.Height = 4950
  
On Error GoTo Problema_Na_Rotina

ProximoNumeroGuia:

  lngNumeroGuia = glngRetornaProximoNumeroGuia
  If Val(lngNumeroGuia) = 0 Then
      Exit Sub
  End If
  txtGuia.Text = lngNumeroGuia
  txtGuia1.Text = lngNumeroGuia

  dblValorReal = adoDataControl.Recordset("dblValorReal")
  
  'Caso seja acordo com parcelas atualizadas, vamos atualizar o valor das parcelas de acordo com o exercicio solicitado
  If blnParcelasAtualizadas And blnValorEmReal Then
      
      strsql = gstrStoredProcedure("sp_AtualizaParcela", adoDataControl.Recordset("intComposicao").Value & ", " & adoDataControl.Recordset("intExercicio").Value & ", " & adoDataControl.Recordset("intParcela").Value & ", " & gstrConvDtParaSql(adoDataControl.Recordset("Dtmdtvencimento").Value) & ", " & gstrConvDtParaSql(adoDataControl.Recordset("Dtmdtvencimento").Value) & ", " & gstrConvVrParaSql(adoDataControl.Recordset("dblValorReal").Value) & ", " & adoDataControl.Recordset("intMoeda").Value, True)

      Set gobjBanco = New clsBanco

      If gobjBanco.CriaADO(strsql, 80, adoResultado) Then
          dblValorReal = Space$(0) & gstrConvVrDoSql(adoResultado("dblValorPrincipal").Value)
      End If
      
      adoResultado.Close: Set adoResultado = Nothing
      
      'Vamos verificar quantas casas decimais possui o valor para aplicar a multiplicacao correta no codigo de barras
      dblValorMult = Val("1" + String(Len(Trim(dblValorReal)) - InStr(1, dblValorReal, ","), "0"))
  
      'Substitui o número da guia do código de barra pelo o que está no banco
      strCodigo = Left(bcCodigoBarra, 3) & Format$(dblValorReal * dblValorMult, "00000000000") & Mid(bcCodigoBarra, 15, 16) & Format$(lngNumeroGuia, "000000000") & Right(bcCodigoBarra, 4)
      
      txtValorParcela = dblValorReal
      txtValorParcela1 = dblValorReal
      
  Else
      'Substitui o número da guia do código de barra pelo o que está no banco
      strCodigo = Left(bcCodigoBarra, 30) & Format$(lngNumeroGuia, "000000000") & Right(bcCodigoBarra, 4)
  End If
  
  strCodigo = Left(strCodigo, 3) & gstrCalculaDigitoModulo10(strCodigo) & _
              Mid(strCodigo, 4, 40)
   
  bcCodigoBarra.Caption = strCodigo
  
  'Insere o Nº da tblGuia
  Set gobjBanco = New clsBanco
  gobjBanco.ExecutaBeginTrans

  strsql = ""
  'strSQL = IIf((bytDBType = EDatabases.Oracle), "BEGIN ", "")
  'Inserir a guia na tabela TblGuias
  strsql = strsql & "INSERT INTO " & gstrGuias & "("
  'strSQL = strSQL & "pkID, "
  strsql = strsql & "intContaBancaria, "
  strsql = strsql & "intNumero, "
  strsql = strsql & "dtmdtEmissao, "
  strsql = strsql & "dblValor, "
  strsql = strsql & "strCodBarra, "
  strsql = strsql & "dtmdtAtualizacao, "
  strsql = strsql & "lngCodUsr, "
  strsql = strsql & "dtmdtVencimento, "
  strsql = strsql & "STRCODBARRAESP "
  strsql = strsql & ") VALUES ("
  'strSQL = strSQL & lngGuias & ","
  strsql = strsql & "NULL, "
  strsql = strsql & txtGuia.Text & ", "
  strsql = strsql & gstrConvDtParaSql(gstrDataDoSistema) & ", "
  'strSql = strSql & gstrConvVrParaSql(txtValorParcela.Text) & ", '"
  strsql = strsql & gstrConvVrParaSql(dblValorReal) & ", '"
  strsql = strsql & bcCodigoBarra.Caption & "', "
  strsql = strsql & gstrConvDtParaSql(gstrDataDoSistema) & ", "
  strsql = strsql & glngCodUsr & ", "
  strsql = strsql & gstrConvDtParaSql(txtDataVencimento.Text) & ", "
  strsql = strsql & " NULL "
  strsql = strsql & ")"
  'strSQL = strSQL & IIf((bytDBType = EDatabases.Oracle), " ; ", "")
  
  Set adoCommand = New ADODB.Command
  Set adoCommand.ActiveConnection = gcncADOMain
  adoCommand.CommandText = strsql
  adoCommand.Execute strsql, , adExecuteNoRecords
  
  'Inserir a guia na tabela TblLancamentoGuias
  strsql = ""
  strsql = "INSERT INTO " & gstrLancamentoGuias & "("
  strsql = strsql & "intlancamentovalor, "
  strsql = strsql & "intguias, "
  strsql = strsql & "dblvalorprincipal, "
  strsql = strsql & "dblvalormulta, "
  strsql = strsql & "dblvalorjuros, "
  strsql = strsql & "dblvalorcorrecao, "
  strsql = strsql & "dblvalordesconto, "
  strsql = strsql & "dtmdtatualizacao, "
  strsql = strsql & "lngcodusr) "
  strsql = strsql & "Values ("
  strsql = strsql & txtPKId.Text & ", "
  strsql = strsql & glngRetornaPkidTabelaPai("seqTblGuias", gstrGuias) & ", "
  'strSql = strSql & gstrConvVrParaSql(txtValorParcela.Text) & ", "
  strsql = strsql & gstrConvVrParaSql(dblValorReal) & ", "
  strsql = strsql & "0, "
  strsql = strsql & "0, "
  strsql = strsql & "0, "
  strsql = strsql & "0, "
  strsql = strsql & gstrConvDtParaSql(gstrDataDoSistema) & ", "
  strsql = strsql & glngCodUsr & ") "
  'strSQL = strSQL & IIf((bytDBType = EDatabases.Oracle), " ; ", "")
  
  'strSQL = strSQL & IIf((bytDBType = EDatabases.Oracle), "END; ", "")
  
  If gobjBanco.Execute(strsql) Then
      gobjBanco.ExecutaCommitTrans
  Else
      ExibeMensagem "Erro na gravação dos lançamentos da guia. Guia não gravada."
      gobjBanco.ExecutaRollbackTrans
      Exit Sub
  End If

  strCodigo = Left(bcCodigoBarra.Caption, 11)
  txtCaptionCodigo.Text = strCodigo & "-" & gstrCalculaDigitoModulo10(strCodigo) & "  "
  
  strCodigo = Mid(bcCodigoBarra.Caption, 12, 11)
  txtCaptionCodigo.Text = txtCaptionCodigo.Text & strCodigo & "-" & gstrCalculaDigitoModulo10(strCodigo) & "  "
  
  strCodigo = Mid(bcCodigoBarra.Caption, 23, 11)
  txtCaptionCodigo.Text = txtCaptionCodigo.Text & strCodigo & "-" & gstrCalculaDigitoModulo10(strCodigo) & "  "
  
  strCodigo = Mid(bcCodigoBarra.Caption, 34, 11)
  txtCaptionCodigo.Text = txtCaptionCodigo.Text & strCodigo & "-" & gstrCalculaDigitoModulo10(strCodigo)
  
  If Trim(txtUtilizacao.Text) <> "" Then
     txtInscricao.Text = gstrFormataInscricao(txtInscricao.Text, txtUtilizacao.Text)
  End If
  txtInscricao1.Text = txtInscricao.Text
  
  If blnPrimeira = True And adoDataControl.NRecords > 1 Then
     Detail.NewPage = ddNPAfter 'Adiciona nova pagina
     blnPrimeira = False
     GroupFooter1.NewPage = ddNPAfter
  Else
     Detail.NewPage = ddNPNone
  End If
  
  'Carrega as intrucoes da parcela
  lblstrInstrucoes = CarregaInstrucoesParcelas(True, adoDataControl.Recordset("strComposicaoDaReceita"), adoDataControl.Recordset("intExercicio"), adoDataControl.Recordset("intParcela"), adoDataControl.Recordset("bitParcelaValida"), adoDataControl.Recordset("PkidAlfa"))
  
  Exit Sub
  
Problema_Na_Rotina:
   
  If InStr(1, UCase(Err.Description), "UK_TBLGUIAS_INTNUMERODTEMISSAO") > 0 Then
      GoTo ProximoNumeroGuia
  Else
      ExibeDetalheErro Err.Description & "- rptCarneParcelas_Detail_Format"
      gobjBanco.ExecutaRollbackTrans
  End If
  
End Sub

Private Sub GroupHeader1_Format()
  If blnPrimeira = True Then
     GroupHeader1.Height = 0
  Else
     'GroupHeader1.Height = 105
  End If
End Sub
