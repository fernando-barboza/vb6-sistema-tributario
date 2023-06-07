VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptRelProtoclizacaoProcessos 
   Caption         =   "Protocolo - rptRelProtoclizacaoProcessos (ActiveReport)"
   ClientHeight    =   10815
   ClientLeft      =   0
   ClientTop       =   285
   ClientWidth     =   15360
   ControlBox      =   0   'False
   MDIChild        =   -1  'True
   WindowState     =   2  'Maximized
   _ExtentX        =   27093
   _ExtentY        =   19076
   SectionData     =   "rptRelProtocolizacaoProcessos.dsx":0000
End
Attribute VB_Name = "rptRelProtoclizacaoProcessos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim adoResultado As ADODB.Recordset

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

Private Sub ActiveReport_ReportStart()
    On Error Resume Next
    PadronizaToolBarRelatorio Me, lblExercicio
    LeImagemLogotipo imgBrasao, imgLogotipo, txtNomeFantasia, txtEstado
    lblRelatorio = Me.Caption
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

Private Sub PageFooter_Format()
    lblPagina.Caption = "Página " & pageNumber
End Sub

Private Sub PageHeader_Format()
    lblDataHora = gstrDataDoSistema(True, , True)
End Sub

Private Sub Detail_Format()
Dim strSql As String
        
        If Len(txtUnidadeCentroCusto.Text) > 0 Then
            lbl_solicitante.Caption = "Unidade Centro Custo:"
        Else
            lbl_solicitante.Caption = "Requerente:"
        End If
        
        Set gobjBanco = New clsBanco
        lbl_endereco.Caption = ""
        txt_Endereco.Text = ""
        txt_Cidade.Text = ""
        txt_BairroEstado.Text = ""
        'strSql = gstrQueryLogradouro(gstrContribuinte, _
        '                             "PKId = " & Val(gstrENulo(intCodContribuinte)), _
        '                             "intLogradouro")
        
        strSql = "SELECT MU.strDescricao as strCidade, CO.strBairroC, CO.strLogradouroC, CO.intNumeroC, "
        strSql = strSql & " CO.strComplementoC, UF.strSigla AS strSigla, UF.strEstado AS strEstado, CO.intCEPC "
        strSql = strSql & "FROM " & gstrContribuinte & " CO, " ' INNER JOIN"
        strSql = strSql & gstrCidade & " MU, " 'ON CO.intMunicipioC = MU.PKId
        strSql = strSql & gstrUF & " UF "
        strSql = strSql & "WHERE CO.intMunicipioC " & strOUTJSQLServer & "= MU.PKId " & strOUTJOracle & " AND "
        strSql = strSql & "CO.intUFC " & strOUTJSQLServer & "= UF.PKId " & strOUTJOracle & " AND "
        strSql = strSql & "CO.PKID = " & Val(gstrENulo(intCodContribuinte))
        
        If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
            With adoResultado
                Do While Not .EOF
                    lbl_endereco.Caption = "Endereço :"
                    txt_BairroEstado = gstrENulo(!strBairroC) + " - " + Trim(!strCidade) + " - " + gstrENulo(!strEstado) + " - " + gstrENulo(!strsigla)
                    txt_Cep.Text = gstrCEPFormatado(gstrENulo(!intcepc))
                    txt_Cidade.Text = gstrENulo(!strCidade)
                     If IsNull(!strComplementoC) Then
                        txt_Endereco = gstrENulo(!strlogradouroc) + ", " + gstrENulo(!intNumeroC)
                     Else
                        txt_Endereco = gstrENulo(!strlogradouroc) + ", " + gstrENulo(!intNumeroC) + " - " + gstrENulo(!strComplementoC)
                     End If
                    .MoveNext
                Loop
            End With
        End If
        
        Set gobjBanco = New clsBanco
        txtAssunto.Text = ""
        If gobjBanco.CriaADO(strQueryAssunto, 5, adoResultado) Then
            With adoResultado
                Do While Not .EOF
                    txtGrupoAssunto = gstrENulo(!grupoassunto)
                    txtTipoAssunto = gstrENulo(!TipoAssunto)
                    txtCatalogoAssunto = gstrENulo(!CatalogoAssunto)
                    .MoveNext
                Loop
            End With
        End If
        Set gobjBanco = New clsBanco
        txt_CidadePrefeitura.Text = ""
        If gobjBanco.CriaADO(strQueryCidadePrefeitura, 5, adoResultado) Then
            With adoResultado
                Do While Not .EOF
                    txt_CidadePrefeitura.Text = gstrENulo(!strDescricao) + ", " + gstrDataPorExtenso(gstrDataDoSistema)
                    .MoveNext
                Loop
            End With
        End If
        Set gobjBanco = New clsBanco
        txt_centrocusto.Text = ""
        If gobjBanco.CriaADO(strQueryLocais, 5, adoResultado) Then
            With adoResultado
                Do While Not .EOF
                    txt_centrocusto.Text = gstrENulo(!CentrodeCusto)
                    .MoveNext
                Loop
            End With
        End If
        txtAssunto = txtGrupoAssunto + " - " + txtTipoAssunto + " - " + txtCatalogoAssunto
        txttextoprotocolo = "Protocolamos com o número " + txtNumProcesso + ", em " + Format(txtDtmdtdata, "dd/mm/yy") + " às " + Format(txtDtmdtdata, "hh:mm") + ", o requerimento descrito na súmula abaixo:"
'        txt_centrocusto = frmCadProtocolizacaoProcesso.dbcintCodCentroCusto.Text
'       TrocaCorParaZebrado lblSombra
End Sub

Private Function strQuerySolicitacao()
    Dim strSql  As String
    strSql = ""
    strSql = strSql & " SELECT strTextoSolicitacao "
    strSql = strSql & " FROM "
    strSql = strSql & gstrTextoSolicitacao
    strQuerySolicitacao = strSql
End Function

Private Function strQueryCidade()
    Dim strSql  As String
    strSql = ""
    strSql = strSql & " SELECT MU.PKId, MU.strDescricao "
    strSql = strSql & " FROM "
    strSql = strSql & gstrCidade & " MU, "
    strSql = strSql & gstrContribuinte & " CO "
    strSql = strSql & "WHERE "
    strSql = strSql & " MU.PKId = CO.intMunicipio "
    strSql = strSql & " AND MU.intUF IN (SELECT MU.intUF FROM "
    strSql = strSql & gstrCidade & ") "
    strSql = strSql & " AND CO.PKID = " & Val(gstrENulo(intCodContribuinte))  '& frmCertificadoHabilitacaoFornecedor.dbcintFornecedor.BoundText
    strSql = strSql & " OR MU.intUF IS NULL "
    strQueryCidade = strSql
End Function

Private Function strQueryBairroEstado()
Dim strSql As String
    strSql = ""
    strSql = strSql & "SELECT UF.PKId, UF.strEstado, UF.strSigla, CO.intCEP, CO.intBairro, BA.strDescricao "
    strSql = strSql & "FROM " & gstrUF & " UF, "
    strSql = strSql & gstrContribuinte & " CO, "
    strSql = strSql & gstrBairro & " BA "
    strSql = strSql & "WHERE "
    strSql = strSql & " UF.PKId = CO.intUF "
    strSql = strSql & " AND BA.PKId = CO.intBairro "
    strSql = strSql & " AND CO.PKID = " & Val(gstrENulo(intCodContribuinte)) '& frmCertificadoHabilitacaoFornecedor.dbcintFornecedor.BoundText
strQueryBairroEstado = strSql
End Function

Private Function strQueryLocais()
'
'******************************************************************************************
' Data: 24/06/2003
' Alteração: - Retirada a tabela tblUnidadeCentroDeCusto, pois será usada tblLocais
' Responsável: Gustavo Monteiro
'******************************************************************************************

Dim strSql As String
    strSql = ""
    strSql = strSql & " SELECT CC.strDescricao AS CentrodeCusto  "
    strSql = strSql & " FROM "
    'strSql = strSql & gstrUnidadeCentroDeCusto2 & " CC, "
    strSql = strSql & gstrLocais & " CC, "
    strSql = strSql & gstrProtocolizacaoProcesso & " PP, "
    strSql = strSql & gstrContribuinte & " CO "
    strSql = strSql & "WHERE "
    strSql = strSql & " CC.PKId = PP.intCodCentrocusto "
    strSql = strSql & " AND CO.PKId = PP.intCodContribuinte "
    strSql = strSql & " AND CO.PKID = " & Val(gstrENulo(intCodContribuinte)) '& frmCertificadoHabilitacaoFornecedor.dbcintFornecedor.BoundText
strQueryLocais = strSql
End Function

Private Function strQueryCidadePrefeitura()
Dim strSql As String
    strSql = ""
    strSql = strSql & " SELECT EP.PKId, MU.strDescricao "
    strSql = strSql & " FROM " & gstrCidade & " MU, "
    strSql = strSql & gstrEmpresa & " EP "
    strSql = strSql & " WHERE "
    strSql = strSql & " MU.PKId = EP.intCidade "
strQueryCidadePrefeitura = strSql
End Function

Private Function strQueryAssunto()
Dim strSql As String

    strSql = ""

    strSql = strSql & " SELECT "
    strSql = strSql & " GP.strDescricao AS GrupoAssunto, "
    strSql = strSql & " TA.strDescricao AS TipoAssunto, "
    strSql = strSql & " CA.strDescricao AS CatalogoAssunto "

    strSql = strSql & " FROM "
    strSql = strSql & gstrGrupoAssunto & " GP, "
    strSql = strSql & gstrTipoAssunto & " TA, "
    strSql = strSql & gstrCatalogoAssunto & " CA "

    strSql = strSql & " WHERE "
    strSql = strSql & " TA.PKId = CA.intTipoAssunto "
    strSql = strSql & " AND GP.PKId = TA.intGrupoAssunto "
    strSql = strSql & " AND CA.PKId = '" & txtintCodAssunto.Text & "'"
    strQueryAssunto = strSql
End Function

Private Sub ReportFooter_Format()
    MostraEmissorRelatorio Me
End Sub

