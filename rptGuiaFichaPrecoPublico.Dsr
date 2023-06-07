VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptGuiaFichaPrecoPublico 
   Caption         =   "rptGuiaFichaPrecoPublico (ActiveReport)"
   ClientHeight    =   11115
   ClientLeft      =   0
   ClientTop       =   285
   ClientWidth     =   15240
   MDIChild        =   -1  'True
   _ExtentX        =   26882
   _ExtentY        =   19606
   SectionData     =   "rptGuiaFichaPrecoPublico.dsx":0000
End
Attribute VB_Name = "rptGuiaFichaPrecoPublico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private ArrayGuia()      As String
Dim iRow                 As Integer

Private Sub ActiveReport_DataInitialize()

    Fields.Add "strNumGuia"
    Fields.Add "dtmRecolher"
    Fields.Add "strContribuinte"
    Fields.Add "strLogradouro"
    Fields.Add "strBairro"
    Fields.Add "strMunicipio"
    Fields.Add "strUf"
    Fields.Add "strQuadra"
    Fields.Add "strLote"
    Fields.Add "strInscricao"
    Fields.Add "strAviso"
    Fields.Add "strReceitas"
    Fields.Add "strHistorico"
    Fields.Add "dblValor"
    Fields.Add "dblCorrecao"
    Fields.Add "dblMulta"
    Fields.Add "dblJuros"
    Fields.Add "dblTotal"
    Fields.Add "dtmEmissao"
    Fields.Add "strFuncionario"
    Fields.Add "dtmVencimento"
    Fields.Add "strCodigoDigitavel"
    Fields.Add "strCodBarras"
    Fields.Add "strProcesso"
    Fields.Add "intCep"
    Fields.Add "intContaBancaria"
    Fields.Add "strNossoNumero"
    
    iRow = LBound(ArrayGuia, 2)
    
End Sub

Private Sub ActiveReport_Activate()
    HabilitaDesabilitaBotao1 False, gstrBtnArquivo, True
End Sub

Private Sub ActiveReport_Deactivate()
    HabilitaDesabilitaBotao1 True, gstrBtnArquivo, True
End Sub

Private Sub ActiveReport_QueryClose(Cancel As Integer, CloseMode As Integer)
    HabilitaDesabilitaBotao1 True, gstrBtnArquivo, True
    HabilitaDesabilitaBotao1 True, gstrMnuArquivo
End Sub

Private Sub ActiveReport_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub ActiveReport_FetchData(EOF As Boolean)

    If iRow > UBound(ArrayGuia, 2) Then
        EOF = True
        Exit Sub
    End If
    
    Fields("strNumGuia") = ArrayGuia(0, iRow)
    Fields("dtmRecolher") = ArrayGuia(1, iRow)
    Fields("strContribuinte") = ArrayGuia(2, iRow)
    Fields("strLogradouro") = ArrayGuia(3, iRow)
    Fields("strBairro") = ArrayGuia(4, iRow)
    Fields("intCep") = ArrayGuia(24, iRow)
    Fields("strMunicipio") = ArrayGuia(21, iRow)
    Fields("strUf") = ArrayGuia(22, iRow)
    Fields("strQuadra") = ArrayGuia(5, iRow)
    Fields("strLote") = ArrayGuia(6, iRow)
    Fields("strInscricao") = ArrayGuia(7, iRow)
    Fields("strAviso") = ArrayGuia(8, iRow)
    Fields("strReceitas") = ArrayGuia(9, iRow)
    Fields("strHistorico") = ArrayGuia(10, iRow)
    Fields("dblValor") = ArrayGuia(11, iRow)
    Fields("dblCorrecao") = ArrayGuia(12, iRow)
    Fields("dblMulta") = ArrayGuia(13, iRow)
    Fields("dblJuros") = ArrayGuia(14, iRow)
    Fields("dblTotal") = ArrayGuia(15, iRow)
    Fields("dtmEmissao") = ArrayGuia(16, iRow)
    Fields("strFuncionario") = ArrayGuia(17, iRow)
    Fields("dtmVencimento") = ArrayGuia(18, iRow)
    Fields("strCodigoDigitavel") = ArrayGuia(19, iRow)
    Fields("strCodBarras") = ArrayGuia(20, iRow)
    Fields("strProcesso") = ArrayGuia(23, iRow)
    Fields("intContaBancaria") = ArrayGuia(25, iRow)
    Fields("strNossoNumero") = ArrayGuia(26, iRow)
    
    EOF = False
    iRow = iRow + 1

End Sub

Private Sub ActiveReport_ReportStart()
    On Error Resume Next
    PadronizaToolBarRelatorio Me
    LeImagemLogotipo imgBrasao, imgLogotipo, txtNomeFantasia, txtEstado
    'LeImagemLogotipo imgBrasao2, imgLogotipo2, txtNomeFantasia2, txtEstado2
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

Public Sub InicializaArray(ArrayCampos() As String)
    ArrayGuia = ArrayCampos
End Sub

Private Sub Detail_Format()

Dim strcnpj As String
Dim strSql      As String
Dim adoBanco  As ADODB.Recordset

    'Vamos atribuir a imagem do banco
    On Error Resume Next
    imgLogoBanco.SizeMode = ddSMZoom

    On Error GoTo 0

    strSql = ""
    strSql = strSql & "SELECT BA.intLogoBanco, BA.intBanco, BA.intDigitoBanco, CB.strCedente, CB.strConta, CB.strDigitoVerificador, AG.strAgencia "
    strSql = strSql & "FROM "
    strSql = strSql & gstrBanco & " BA, " & gstrContaBancaria & " CB, " & gstrAgencia & " AG "
    strSql = strSql & "WHERE BA.PKId " & strOUTJOracle & "=" & strOUTJSQLServer & "CB.intBanco AND " & _
                      "AG.Pkid = CB.intAgencia AND " & _
                      "CB.Pkid = " & txtintConta

    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 5, adoBanco) Then
        With adoBanco
            If .EOF = False Then
                
                LeImagem Val(gstrENulo(!intLogoBanco)), imgLogoBanco
                txtstrCodigoBanco = !intBanco
                
                txtstrCodigoBanco = Format(txtstrCodigoBanco, "000")
                txtstrCodigoBanco = txtstrCodigoBanco & IIf(IsNull(!intDigitoBanco), "", "-" & !intDigitoBanco)
                
                txtstrAgencia = !strAgencia & " " & !strConta & " " & !strDigitoVerificador
                
            End If
        End With
        adoBanco.Close: Set adoBanco = Nothing
    Else
        Exit Sub
    End If
    
  If frmCadLancamentoPrecoPublico.dbcintIdentificacao.MatchedWithList Then
        strcnpj = strcnpj & "SELECT CO.STRCNPJCPF cnpjcpf "
        strcnpj = strcnpj & "FROM "
        strcnpj = strcnpj & gstrContribuinte & " CO, "
        strcnpj = strcnpj & gstrImobiliario & " IM "
        strcnpj = strcnpj & "WHERE CO.Pkid = IM.intContribuinte "
        strcnpj = strcnpj & "AND IM.strInscricao LIKE " & "'%" & frmCadLancamentoPrecoPublico.dbcintIdentificacao.Text & "'"
                    
   Else

        If bytDBType = Oracle Then
            strcnpj = strcnpj & "SELECT LA.STRCNPJCPF cnpjcpf "
        Else
            strcnpj = ""
            strcnpj = strcnpj & "SELECT "
            strcnpj = strcnpj & gstrTOPnSQLServer(1)
            strcnpj = strcnpj & "LA.STRCNPJCPF cnpjcpf "
        End If
        
        strcnpj = strcnpj & " FROM "
        strcnpj = strcnpj & gstrLancamentoAlfa & " LA"
        strcnpj = strcnpj & " WHERE LA.strNomeProprietario =" & "'" & Fields("strContribuinte") & "'"
        strcnpj = strcnpj & " ORDER BY LA.pkid desc"
   
        If bytDBType = Oracle Then
            strcnpj = gstrTOPnOracle(strcnpj, 1)
            
        End If
   
   End If
   
   Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strcnpj, 5, adoBanco) Then
        With adoBanco
            If .EOF = False Then
                
            txtcpfncpj = IIf(IsNull(!cnpjcpf), "", !cnpjcpf)
                
                If Len(txtcpfncpj) = 0 Then
                    lblcpfcnpj.Visible = False
                Else
                      If Len(txtcpfncpj) = 11 Then
                            txtcpfncpj = Left(txtcpfncpj, 3) & "." & Mid(txtcpfncpj, 4, 3) & "." & Mid(txtcpfncpj, 7, 3) & "-" & Right(txtcpfncpj, 2)
                       Else
                            If Len(txtcpfncpj) = 14 Then
                                txtcpfncpj = Left(txtcpfncpj, 2) & "." & Mid(txtcpfncpj, 3, 3) & "." & Mid(txtcpfncpj, 6, 3) & " / " & Mid(txtcpfncpj, 9, 4) & "-" & Right(txtcpfncpj, 2)
                            End If
                      End If
                 End If
                
            End If
        End With
        adoBanco.Close: Set adoBanco = Nothing
    Else
        Exit Sub
    End If
   
   
   
   
   
    
    strSql = ""
    strSql = strSql & "SELECT EM.strNomeFantasia "
    strSql = strSql & "FROM "
    strSql = strSql & gstrEmpresa & " EM "

    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 5, adoBanco) Then
        With adoBanco
            If .EOF = False Then
                
                txtstrCedente = !strNomeFantasia
                
            End If
        End With
        adoBanco.Close: Set adoBanco = Nothing
    Else
        Exit Sub
    End If
    
    'Carrega as intrucoes da parcela
    txtstrInstrucoes = CarregaInstrucoesParcelas(True, ArrayGuia(27, 0), CInt(ArrayGuia(28, 0)), CInt(ArrayGuia(29, 0)), ArrayGuia(30, 0) = True, CLng(ArrayGuia(31, 0)))

    txtintCep1 = gstrCEPFormatado(gstrVerificaCampoNulo(txtintCep1))
    txtdtmDocumento.Text = gstrDataDoSistema
    txtdtmProcessamento.Text = gstrDataDoSistema
    'txtintCep2 = gstrCEPFormatado(gstrVerificaCampoNulo(txtintCep2))
    
    'Alteração feita Tri0747
    txtdblCorrecao1 = ""
    txtdblMulta1 = ""
    txtdblJuros1 = ""
    txtdblTotal1 = ""
    
    txtdblMultaMora = ""
    txtdblAcrescimos1 = ""
    txtValorCobrado = ""
    
End Sub

