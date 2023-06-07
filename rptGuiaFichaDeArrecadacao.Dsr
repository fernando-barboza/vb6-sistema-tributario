VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptGuiaFichaDeArrecadacao 
   Caption         =   "Tributario - rptGuiaFichaDeArrecadacao (ActiveReport)"
   ClientHeight    =   10815
   ClientLeft      =   0
   ClientTop       =   285
   ClientWidth     =   15360
   MDIChild        =   -1  'True
   _ExtentX        =   27093
   _ExtentY        =   19076
   SectionData     =   "rptGuiaFichaDeArrecadacao.dsx":0000
End
Attribute VB_Name = "rptGuiaFichaDeArrecadacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private ArrayGuia()      As String
Dim iRow                 As Integer
Dim blnFieldsExists      As Boolean

 Private Sub ActiveReport_DataInitialize()
    
    If Not blnFieldsExists Then
        
        Fields.Add "strNumGuia"
        Fields.Add "dtmRecolher"
        Fields.Add "strContribuinte"
        Fields.Add "strLogradouro"
        Fields.Add "strBairro"
        Fields.Add "strQuadra"
        Fields.Add "strInscricao"
        Fields.Add "strAviso"
        Fields.Add "strComposicao"
        Fields.Add "strParcelas"
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
        Fields.Add "intContaBancaria"
        Fields.Add "strNossoNumero"
        
        blnFieldsExists = True
        
    End If
    
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
    Fields("strQuadra") = ArrayGuia(5, iRow)
    Fields("strInscricao") = ArrayGuia(6, iRow)
    Fields("strAviso") = ArrayGuia(7, iRow)
    Fields("strComposicao") = ArrayGuia(8, iRow)
    Fields("strParcelas") = ArrayGuia(9, iRow)
    Fields("dblValor") = ArrayGuia(10, iRow)
    Fields("dblCorrecao") = ArrayGuia(11, iRow)
    Fields("dblMulta") = ArrayGuia(12, iRow)
    Fields("dblJuros") = ArrayGuia(13, iRow)
    Fields("dblTotal") = ArrayGuia(14, iRow)
    Fields("dtmEmissao") = ArrayGuia(15, iRow)
    Fields("strFuncionario") = ArrayGuia(16, iRow)
    Fields("dtmVencimento") = ArrayGuia(17, iRow)
    Fields("strCodigoDigitavel") = ArrayGuia(18, iRow)
    Fields("strCodBarras") = ArrayGuia(19, iRow)
    Fields("intContaBancaria") = ArrayGuia(20, iRow)
    Fields("strNossoNumero") = ArrayGuia(21, iRow)
    
    EOF = False
    iRow = iRow + 1

End Sub

Private Sub ActiveReport_ReportStart()
    On Error Resume Next
    PadronizaToolBarRelatorio Me
    LeImagemLogotipo imgBrasao, imgLogotipo, txtNomeFantasia, txtEstado
    'LeImagemLogotipo imgBrasao2, imgLogotipo2, txtNomeFantasia2, txtEstado2
End Sub

Private Sub ActiveReport_ToolbarClick(ByVal tool As DDActiveReports2.DDTool)
    Dim vnt As Variant
    If tool.ID = 14 Then
        ActiveReport_KeyPress 27
    ElseIf tool.ID = 15 Then
        AbreOpcoesExportacao Me
    ElseIf tool.ID = 16 Then
        Configura_Relatorio Me, True
    End If
End Sub

Public Sub InicializaArray(ArrayCampos() As String)
    ArrayGuia = ArrayCampos
End Sub

Private Sub Detail_Format()
Dim strSql      As String
Dim adoBanco  As ADODB.Recordset

    'Vamos atribuir a imagem do banco
    On Error Resume Next
    imgLogoBanco.SizeMode = ddSMZoom

    On Error GoTo 0

    strSql = ""
    strSql = strSql & "SELECT BA.intLogoBanco, BA.intBanco, BA.intDigitoBanco, CB.strCedente, CB.strDigitoVerificador, AG.strAgencia "
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
                
                txtstrAgencia = !strAgencia & " " & !strCedente

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
    
    txtdtmDocumento.Text = gstrDataDoSistema
    txtdtmProcessamento.Text = gstrDataDoSistema
    
End Sub
