VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptGuiaDeArrecadacaoMunicipal 
   Caption         =   "Guias de Arrecadação Municipal"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   MDIChild        =   -1  'True
   WindowState     =   2  'Maximized
   _ExtentX        =   20955
   _ExtentY        =   15161
   SectionData     =   "GuiaDeArrecadacaoMunicipal.dsx":0000
End
Attribute VB_Name = "rptGuiaDeArrecadacaoMunicipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strCGC As String
Public strImposto As String

Private Sub ActiveReport_Activate()
    HabilitaDesabilitaBotao1 False, gstrBtnArquivo, True
End Sub

Private Sub ActiveReport_Deactivate()
    HabilitaDesabilitaBotao1 True, gstrBtnArquivo, True
End Sub

Private Sub ActiveReport_QueryClose(Cancel As Integer, CloseMode As Integer)
    HabilitaDesabilitaBotao1 True, gstrBtnArquivo, True
End Sub

Private Sub ActiveReport_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub ActiveReport_ReportEnd()
Dim i As Integer
    For i = 0 To rptGuiaDeArrecadacaoMunicipal.Pages.Count - 1
        rptGuiaDeArrecadacaoMunicipal.Pages(i).Orientation = ddOLandscape
    Next
End Sub

Private Sub ActiveReport_ReportStart()
    PadronizaToolBarRelatorio Me
    Label1.Caption = Label1.Caption & Chr(13) & strImposto
    Label27.Caption = Label1.Caption
    CGCEmpresa
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

Private Sub Imagens()
    On Error Resume Next
    LeImagemLogotipo imgBrasao, imgLogotipo, txtNomeFantasia, txtEstado
    LeImagemLogotipo imgBrasao, imgLogotipo2, txtNomeFantasia2, txtEstado
End Sub

Private Sub Detail_BeforePrint()

'******************************************************************************************
' Data: 04/04/2003
' Alteração: - Substituição do comando nativo ISNULL() do SQL Server pela variável
'            gstrISNULL.
' Responsável: Everton Bianchini
'------------------------------------------------------------------------------------------
' Data: 04/04/2003
' Alteração: - Substituição do comando nativo de concatenação de strings (+) do SQL Server
'        pela variável strCONCAT.
' Responsável: Everton Bianchini
'------------------------------------------------------------------------------------------
' Data: 07/04/2003
' Alteração: - Substituição do comando nativo de outer join (*= ou =*) do SQL Server pela
'        variável strOUTJSQLServer e inclusão do comando de outer join ((+)) do Oracle,
'        representado pela variável strOUTJOracle.
' Responsável: Everton Bianchini
'******************************************************************************************
    
    Dim strSql As String
    Dim adoRec As ADODB.Recordset
    
    Barcode1.CreatePictureBySize Barcode1.Width, Barcode1.Height
    Image1.Picture = Barcode1.Picture
    
    strSql = ""
'    strSql = strSql & " SELECT I.dblValorEdificacao, I.dblValorTerreno, I.PKId, L.PKId, RTRIM(LTRIM(ISNULL(TL.strSigla, '') + ' ' + ISNULL(U.strDescricao,'') + "
    strSql = strSql & " SELECT I.dblValorEdificacao, I.dblValorTerreno, I.PKId, L.PKId, RTRIM(LTRIM(" & gstrISNULL("TL.strSigla", "''") & strCONCAT & " ' ' " & strCONCAT & gstrISNULL("U.strDescricao", "''") & strCONCAT
'    strSql = strSql & " ' ' + L.strDescricao)) AS Logradouro, I.intNumero , I.strComplemento, B.strDescricao, I.intCep, "
    strSql = strSql & " ' ' " & strCONCAT & " L.strDescricao)) AS Logradouro, I.intNumero , I.strComplemento, B.strDescricao, I.intCep, "
    strSql = strSql & " UF.strSigla "
    strSql = strSql & " FROM "
    strSql = strSql & gstrLogradouro & " L,"
    strSql = strSql & gstrTituloLogradouro & " U,"
    strSql = strSql & gstrTipoLogradouro & " TL ,"
    strSql = strSql & gstrImobiliario & " I,"
    strSql = strSql & gstrBairro & " B,"
    strSql = strSql & gstrUF & " UF "
    strSql = strSql & " WHERE "
'    strSql = strSql & " L.intTituloLogradouro *= U.PKId"
    strSql = strSql & " L.intTituloLogradouro " & strOUTJSQLServer & "= U.PKId" & strOUTJOracle
'    strSql = strSql & " AND L.intTipoLogradouro *= TL.PKId"
    strSql = strSql & " AND L.intTipoLogradouro " & strOUTJSQLServer & "= TL.PKId" & strOUTJOracle
    strSql = strSql & " AND I.intLogradouro = L.PKId"
    strSql = strSql & " AND I.intBairro = B.PKId"
    strSql = strSql & " AND UF.PKId = I.intUf"
    strSql = strSql & " and I.strInscricaoAnterior = '" & txtstrInscricaoCadastral & "'"
    
    txtLogradouroImovel.Text = ""
    txtNumImovel.Text = ""
    txtCompImovel.Text = ""
    txtBairroImovel.Text = ""
    txtCEPImovel.Text = ""
    txtMunUFImovel.Text = ""
    txtdblValorEdificacao.Text = ""
    
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 10, adoRec) Then
        With adoRec
            If Not .EOF Then
                txtLogradouroImovel.Text = gstrENulo(!Logradouro)
                txtNumImovel.Text = gstrENulo(!intNumero)
                txtCompImovel.Text = gstrENulo(!strComplemento)
                txtBairroImovel.Text = gstrENulo(!strDescricao)
                txtCEPImovel.Text = "CEP.: " & gstrENulo(!intCep)
                txtMunUFImovel.Text = gstrCidadeEmpresa & " - " & gstrUFEmpresa
                txtdblValorEdificacao.Text = gstrConvVrDoSql(!dblValorEdificacao)
                txtdblValorTerreno.Text = gstrConvVrDoSql(!dblValorTerreno)
            End If
        End With
    End If
    
    'SomaEMostra
End Sub

Private Sub Detail_Format()

'******************************************************************************************
' Data: 12/05/2003
' Alteração: - Adaptação da string de conexão.
' Responsável: Everton Bianchini
'------------------------------------------------------------------------------------------
' Data: 13/05/2003
' Alteração: - Incluídas chamadas à função CriaADO para criar objetos ADODB.Recordset os
'            quais são passados para os subrelatórios.
' Responsável: Everton Bianchini
'******************************************************************************************

Dim strConnString As String
Dim strSql As String

Dim adoResultSetSubUm As ADODB.Recordset
Dim adoResultSetSubDois As ADODB.Recordset

strConnString = gcncADOMain.ConnectionString

    'ESTE Relatorio possui um sub relatorio
    
Set gobjBanco = New clsBanco
    
Set rptSubUm.object = New rptSubRelatorioGuiaDeArrecadacao
Set rptSubDois.object = New rptSubRelatorioGuiaDeArrecadacao1

    If MDIMenu.Tag = "Ouvidoria" Then
'        rptSubUm.object.adoDataControl.ConnectionString = "driver={SQL Server};Server=" & gstrServidor & ";database=" & gstrDatabase & ";uid=" & gstrLoginUser & ";pwd=" & gstrPwdUser & ";"
        rptSubUm.object.adoDataControl.ConnectionString = strConnString
        strSql = rptSubUm.object.adoDataControl.Source & " and B.intLancamentoCalculo = " & Val(txtPKId) & " GROUP BY A.strSigla, B.dblValorParcela "
    
        Call gobjBanco.CriaADO(strSql, 5, adoResultSetSubUm)
'        rptSubUm.object.adoDataControl.Source = strsql
        Set rptSubUm.object.adoDataControl.Recordset = adoResultSetSubUm
        
'        rptSubDois.object.adoDataControl.ConnectionString = "driver={SQL Server};Server=" & gstrServidor & ";database=" & gstrDatabase & ";uid=" & gstrLoginUser & ";pwd=" & gstrPwdUser & ";"
        rptSubDois.object.adoDataControl.ConnectionString = strConnString
'        rptSubDois.object.adoDataControl.Source = rptSubDois.object.adoDataControl.Source & " and B.intLancamentoCalculo = " & Val(txtPKId) & " GROUP BY A.strSigla, B.dblValorParcela "
        strSql = rptSubDois.object.adoDataControl.Source & " and B.intLancamentoCalculo = " & Val(txtPKId) & " GROUP BY A.strSigla, B.dblValorParcela "
        rptSubDois.object.adoDataControl.Source = strSql
        Call gobjBanco.CriaADO(strSql, 5, adoResultSetSubDois)
        
        Set rptSubDois.object.adoDataControl.Recordset = adoResultSetSubDois
            
        txt_Informacoes.Text = frmGuiaDeArrecadacao.txt_Mensagem1.Text
        txt_Observacao.Text = frmGuiaDeArrecadacao.txt_Mensagem2.Text
        txt_Observacao2.Text = frmGuiaDeArrecadacao.txt_Mensagem2.Text
    Else
'        rptSubUm.object.adoDataControl.ConnectionString = "driver={SQL Server};Server=" & gstrServidor & ";database=" & gstrDatabase & ";uid=" & gstrLoginUser & ";pwd=" & gstrPwdUser & ";"
        rptSubUm.object.adoDataControl.ConnectionString = strConnString
        'rptSubUm.object.adoDataControl.Source = rptSubUm.object.adoDataControl.Source & " and B.intLancamentoCalculo = " & Val(txtPKId) & " AND B.intNumeroParcela BETWEEN " & gfrmFormularioQueEstaImprimindoGuia.txt_intParcelaInicial.Text & " AND " & gfrmFormularioQueEstaImprimindoGuia.txt_intParcelaFinal.Text
'        rptSubUm.object.adoDataControl.Source = rptSubUm.object.adoDataControl.Source & " AND intNumeroParcela = " & Val(txtintNumeroParcela.Text)
        strSql = rptSubUm.object.adoDataControl.Source & " AND intNumeroParcela = " & Val(txtintNumeroParcela.Text)
'        rptSubUm.object.adoDataControl.Source = rptSubUm.object.adoDataControl.Source & " AND intLancamentoCalculo = " & adoDataControl.Recordset!PKId
        strSql = strSql & " AND intLancamentoCalculo = " & adoDataControl.Recordset!Pkid
        rptSubUm.object.adoDataControl.Source = strSql
        
        Call gobjBanco.CriaADO(strSql, 5, adoResultSetSubUm)
        
        Set rptSubUm.object.adoDataControl.Recordset = adoResultSetSubUm
        
'        rptSubDois.object.adoDataControl.ConnectionString = "driver={SQL Server};Server=" & gstrServidor & ";database=" & gstrDatabase & ";uid=" & gstrLoginUser & ";pwd=" & gstrPwdUser & ";"
        rptSubDois.object.adoDataControl.ConnectionString = strConnString
        'rptSubDois.object.adoDataControl.Source = rptSubDois.object.adoDataControl.Source & " and B.intLancamentoCalculo = " & Val(txtPKId) & " AND B.intNumeroParcela BETWEEN " & gfrmFormularioQueEstaImprimindoGuia.txt_intParcelaInicial.Text & " AND " & gfrmFormularioQueEstaImprimindoGuia.txt_intParcelaFinal.Text
'        rptSubDois.object.adoDataControl.Source = rptSubDois.object.adoDataControl.Source & " AND intNumeroParcela = " & Val(txtintNumeroParcela.Text)
        strSql = rptSubDois.object.adoDataControl.Source & " AND intNumeroParcela = " & Val(txtintNumeroParcela.Text)
'        rptSubDois.object.adoDataControl.Source = rptSubDois.object.adoDataControl.Source & " AND intLancamentoCalculo = " & adoDataControl.Recordset!PKId
        strSql = strSql & " AND intLancamentoCalculo = " & adoDataControl.Recordset!Pkid
        rptSubDois.object.adoDataControl.Source = strSql
        
        Call gobjBanco.CriaADO(strSql, 5, adoResultSetSubDois)
        
        Set rptSubDois.object.adoDataControl.Recordset = adoResultSetSubDois
        
        txt_Informacoes.Text = gfrmFormularioQueEstaImprimindoGuia.txt_Mensagem1.Text
        txt_Observacao.Text = gfrmFormularioQueEstaImprimindoGuia.txt_Mensagem2.Text
        txt_Observacao2.Text = gfrmFormularioQueEstaImprimindoGuia.txt_Mensagem2.Text
    End If
    Imagens
    SomaEMostra

    ConfiguraCodigoBarra
End Sub

Private Function SomaEMostra()
Dim strSql As String
Dim adoResultado As ADODB.Recordset
If MDIMenu.Tag = "Ouvidoria" Then
    strSql = ""
    strSql = strSql & " SELECT sum(dblValorParcela) as Total "
    strSql = strSql & " FROM "
    strSql = strSql & gstrParcelaTaxa
    strSql = strSql & " WHERE intLancamentoCalculo = " & Val(txtPKId)
    strSql = strSql & " AND intNumeroParcela = " & Val(txtintNumeroParcela.Text)
ElseIf MDIMenu.Tag = "frmCadDebito" Then
    strSql = ""
    strSql = strSql & " SELECT sum(dblValorParcela) as Total "
    strSql = strSql & " FROM "
    strSql = strSql & gstrParcelaTaxa
    strSql = strSql & " WHERE intLancamentoCalculo = " & Val(txtPKId)
    strSql = strSql & " AND intNumeroParcela = " & Val(txtintNumeroParcela.Text)
Else
    strSql = ""
    strSql = strSql & " SELECT sum(dblValorParcela) as Total "
    strSql = strSql & " FROM "
    strSql = strSql & gstrParcelaTaxa
    strSql = strSql & " WHERE intLancamentoCalculo = " & Val(txtPKId)
    strSql = strSql & " AND intNumeroParcela BETWEEN " & gfrmFormularioQueEstaImprimindoGuia.txt_intParcelaInicial.Text & " AND " & gfrmFormularioQueEstaImprimindoGuia.txt_intParcelaFinal.Text
    strSql = strSql & " AND intNumeroParcela = " & Val(txtintNumeroParcela.Text)
End If
Set gobjBanco = New clsBanco
If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
    If adoResultado.EOF = False Then
        txt_Total1.Text = gstrConvVrDoSql(adoResultado!Total)
        txt_Total2.Text = gstrConvVrDoSql(adoResultado!Total)
        
        If txtJuros.Text = "" Then
            txtJuros.Text = 0
        End If
        If txtMulta.Text = "" Then
            txtMulta.Text = 0
        End If
        If txt_Total1.Text = "" Then
            txt_Total1.Text = 0
        End If

        txtVrTotal.Text = gstrConvVrDoSql(txt_Total1.Text + CDbl(txtJuros.Text) + CDbl(txtMulta.Text))
        
        If txtJuros.Text = 0 Then
            txtJuros.Text = ""
        End If
        If txtMulta.Text = 0 Then
            txtMulta.Text = ""
        End If
        
        If txt_Total1.Text = 0 Then
            txt_Total1.Text = ""
        End If
       
        If txtJuros1.Text = "" Then
            txtJuros1.Text = 0
        End If
        If txtMulta1.Text = "" Then
            txtMulta1.Text = 0
        End If
        
        If txt_Total2.Text = "" Then
            txt_Total2.Text = 0
        End If

        txtVrTotal1.Text = gstrConvVrDoSql(txt_Total2.Text + CDbl(txtJuros1.Text) + CDbl(txtMulta1.Text))
        
        If txtJuros1.Text = 0 Then
            txtJuros1.Text = ""
        End If
        If txtMulta1.Text = 0 Then
            txtMulta1.Text = ""
        End If
        
        If txt_Total2.Text = 0 Then
            txt_Total2.Text = ""
        End If
        
        adoResultado.MoveNext
    End If
End If
End Function

Private Sub ConfiguraCodigoBarra()
    Dim strCampo1 As String
    Dim strCampo2 As String
    Dim strDAC    As String
    Dim strAux    As String
    
    If Trim(txt_Total1) = "" Then
        txt_Total1 = 0
    End If
    
    strCampo1 = ""
    strCampo1 = strCampo1 & "8" 'Arrecadação
    strCampo1 = strCampo1 & "1" 'Prefeituras
    strCampo1 = strCampo1 & "6" 'Valor a ser cobrado efetivamente em reais
    
    strCampo2 = ""
    strCampo2 = strCampo2 & Format(CStr(Trim(gstrConvVrParaSql(txt_Total1))), "00000000000")    'Valor
    strCampo2 = strCampo2 & Mid(strCGC, 1, 8) 'CGC / MF
    strCampo2 = strCampo2 & Trim(txtPKIdParcelaReceita) & String(21 - Len(Trim(txtPKIdParcelaReceita)), "0")   'Uso livre
    
    strDAC = gstrDigitoVerificador(strCampo1 & strCampo2)    'Digito Verificador (Módulo 10)
    
    strAux = CStr(strCampo1 & strDAC & strCampo2)
    Barcode1.Text = strAux
    lblCodigoBarra1 = Mid(strAux, 1, 11) & " " & gstrDigitoVerificador(Mid(strAux, 1, 11))
    lblCodigoBarra2 = Mid(strAux, 12, 11) & " " & gstrDigitoVerificador(Mid(strAux, 12, 11))
    lblCodigoBarra3 = Mid(strAux, 23, 11) & " " & gstrDigitoVerificador(Mid(strAux, 23, 11))
    lblCodigoBarra4 = Mid(strAux, 34, 11) & " " & gstrDigitoVerificador(Mid(strAux, 34, 11))
End Sub

Private Function gstrDigitoVerificador(strNumero As String) As String
    Dim strValores As String
    Dim lngValor   As Long
    Dim i          As Integer
    Dim blnFlag    As Boolean
    
    strValores = ""
    
    blnFlag = True
    For i = 1 To Len(strNumero)
        If blnFlag Then
            strValores = strValores & CStr(Val(Mid(strNumero, i, 1)) * 2)
            blnFlag = False
        Else
            strValores = strValores & CStr(Val(Mid(strNumero, i, 1)) * 1)
            blnFlag = True
        End If
    Next
    
    lngValor = 0
    For i = 1 To Len(strValores)
        lngValor = lngValor + Val(Mid(strValores, i, 1))
    Next
    lngValor = lngValor Mod 10
    If lngValor = 0 Then
        gstrDigitoVerificador = 0
    Else
        gstrDigitoVerificador = 10 - lngValor
    End If
End Function

Private Sub CGCEmpresa()
    Dim adoResultado As ADODB.Recordset
    Dim strSql       As String
    
    strCGC = ""
    
    strSql = ""
    strSql = strSql & "SELECT strCGC "
    strSql = strSql & "FROM " & gstrEmpresa
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        With adoResultado
            If Not (.EOF And .BOF) Then
                strCGC = gstrValorSemMascara(!strCGC)
            End If
        End With
    End If
End Sub
