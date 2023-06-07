VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRelatorio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public strCarregaBrasao            As String
Public TempTabela                  As String
Public strCarregaLogotipo          As String
Public strCarregaEstado            As String
Public strCarregaNomeFantasia      As String
Public strTelefone                 As String

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' PROSPECTO PARA UTILIZAÇÃO DAS FUNÇÕES EM RELATÓRIOS EXTERNOS               '
'                                                                            '
'   Para Utilizar as Funções e Variaveis do módulo nos relatórios externos   '
'   utilizar o nome do módulo de classe ( ClsRelatorio ) .NomeDaFuncao       '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub MostraEmissorRelatorio(objRelatorio As Object)
Dim objField As Object
On Error Resume Next
    With objRelatorio
        If gbytRelatorioComEmissor = True Then
            .lblEmitido.Visible = True
            .txtEmitido.Visible = True
            .txtEmitido.Text = gstrNomeUsuario
        End If
    End With
     
End Sub
Public Function Modalidad() As Boolean
Modalidad = Modalidade
End Function
'FUNÇÃO USADA PARA ELIMINAR O PRIMEIRO ELEMENTO DO ARRAY MARCAS
Public Function Ajeita_Marca(Codigo As Long)
Dim j As Integer
j = 0
Do While j <= Qtd_Marcas
    If Table_Marcas(j, 1) = Codigo Then
        Ajeita_Marca = Table_Marcas(j, 0)
        Exit Do
    End If
    j = j + 1
Loop
End Function

Public Function InformaPeriodo(strFormulario As String, _
                              strCampoTextoDtInicial As String, _
                              strCampoTextoDtFinal As String) As String
    
    '------------------------------------------------------------------------'
    ' Função usada para mostrar o período inicial e final de uma consulta    '
    ' em um relatório externo                                                '
    '------------------------------------------------------------------------'
    ' Parâmetros:                                                            '
    '                                                                        '
    ' 1 - strFormulario: Nome do formuário de consulta                       '
    ' 2 - strCampoTextoDtInicial: Campo texto onde é digitado a data inicial '
    ' 3 - strCampoTextoDtFinal: Campo texto onde é digitado a data final     '
    '------------------------------------------------------------------------'
    
    Dim strPeriodo As String
    Dim objFormulario As Object
    Dim objCampoTexto As Object
    
    For Each objFormulario In Forms
        If objFormulario.Name = strFormulario Then
            For Each objCampoTexto In objFormulario.Controls
                If objCampoTexto.Name = strCampoTextoDtInicial Then
                    strPeriodo = "Período: " & objCampoTexto & " à "
                End If
            Next
        End If
    Next
    
    For Each objFormulario In Forms
        If objFormulario.Name = strFormulario Then
            For Each objCampoTexto In objFormulario.Controls
                If objCampoTexto.Name = strCampoTextoDtFinal Then
                    strPeriodo = strPeriodo & objCampoTexto
                End If
            Next
        End If
    Next
    
    InformaPeriodo = strPeriodo
    
End Function

Public Function CarregaBrasao() As String
    CarregaBrasao = strCarregaBrasao
End Function

Public Function CarregaLogotipo() As String
    CarregaLogotipo = strCarregaLogotipo
End Function

Public Function CarregaEstado() As String
    CarregaEstado = strCarregaEstado
End Function

Public Function CarregaNomeFantasia() As String
    CarregaNomeFantasia = strCarregaNomeFantasia
End Function

Public Function IIf(Expression, TruePart, FalsePart)
    IIf = VBA.IIf(Expression, TruePart, FalsePart)
End Function

Public Function LenRPX(Expression)
    LenRPX = Len(Expression)
End Function

Public Function Trim(Expression)
    Trim = VBA.Trim(Expression)
End Function

Public Function Val(Str) As Double
    Val = VBA.Val(Str)
End Function

Public Function gstrMascaraEspecifica(objControle As Object, bytTipo As Byte) As String

''''''''''''''''''''''''''''''''''''''''''''''''''
' Mascaras do Orçamento para Relatórios Externos '
'                                                '
' Parametros:                                    '
'    objControle ( Controle em Questão )         '
'    bytTipo ( Tipo da Mascara )                 '
''''''''''''''''''''''''''''''''''''''''''''''''''
        
' 1 - Conta Contabil
' 2 - Código Orçamentário
' 3 - Elemento da Despesa
' 4 - Item da Despesa

    gstrMascaraEspecifica = gvntFormatacaoEspecifica(objControle, bytTipo)
    
End Function

Public Sub LeImagemLogotipoExterno(imgBrasao, imgLogotipo, Optional txtNomeFantasia As Object, Optional txtEstado As Object)

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Carrega as imagens e labels parametrizadas no Sistema.                            '
'                                                                                   '
' Parametros:                                                                       '
'    imgBrasao ( Controle Picture que recebera o Brasão )                           '
'    imgLogoTipo ( Controle Picture para o Logotipo )                               '
'    txtNomeFantasia ( Controle Label ou Txt para receber o Nome Fantasia )         '
'    txtEstado ( Estado do cliente )                                                '
'                                                                                   '
' *** Utilizar no Evento ReportStart dos Relatórios Externos.                       '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


    LeImagemLogotipo imgBrasao, imgLogotipo, txtNomeFantasia, txtEstado

End Sub

Public Sub PadronizaToolBarRelatorioExterno(rptRelatorio As Object, Optional lblExercicio As Object, Optional vntExercicio)
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Carrega as opções e padronizações do Sistema para os relatórios externos. '
'                                                                           '
' Parametros:                                                               '
'    rptRelatorio ( Relatório em questão )                                  '
'    lblExercicio ( Controle Label ou Txt para receber o Exercicio )        '
'    vntExercicio ( Exercicio -  Desejado Default: Exercício Corrente )     '
'                                                                           '
' *** Utilizar no Evento ReportStart dos Relatórios Externos.               '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    PadronizaToolBarRelatorio rptRelatorio, lblExercicio, vntExercicio
    
End Sub

Public Function Format(Expression, sFormat)
    Format = VBA.Format(Expression, sFormat)
End Function

Public Sub ImagemLogotipo()
    
    Dim strSql      As String
    Dim adoEmpresa  As ADODB.Recordset
    
    strCarregaNomeFantasia = ""
    strCarregaEstado = ""
    
    strSql = ""
    strSql = strSql & "SELECT EM.intLogotipo, EM.intBrasao, "
    strSql = strSql & "EM.strNome, UF.strEstado "
    strSql = strSql & "FROM "
    strSql = strSql & gstrEmpresa & " EM, "
    strSql = strSql & gstrUF & " UF "
    strSql = strSql & "WHERE UF.PKId " & strOUTJOracle & "=" & strOUTJSQLServer & "EM.intUF"
    Set gobjBanco = New clsBanco
    
    If gobjBanco.CriaADO(strSql, 5, adoEmpresa) Then
        With adoEmpresa
            If .EOF = False Then
                strCarregaLogotipo = LeImagem(Val(gstrENulo(!intLogotipo)))
                strCarregaBrasao = LeImagem(Val(gstrENulo(!intBrasao)))
                    
                strCarregaNomeFantasia = gstrStringCripitografada(gstrENulo(!STRNOME))
                strCarregaEstado = gstrENulo(!strEstado)
                                    
            End If
        End With
        adoEmpresa.Close
        Set adoEmpresa = Nothing
        Set gobjBanco = Nothing
    Else
        Exit Sub
    End If
End Sub

Public Sub CarregaTitulo(lblRelatorio As Object)
    lblRelatorio = frmVisualizarRelatorio.Caption
End Sub

Private Function LeImagem(intCodigo As Integer) As String

    Dim adoResultado    As ADODB.Recordset
    Dim strSql          As String
    Dim intNumArquivo   As Integer
    Dim intFragmento    As Variant
    Dim bytPedaco()     As Byte
    On Error GoTo ErroLeImagem
    Screen.MousePointer = vbHourglass
    strSql = "SELECT * FROM " & gstrImagem & " WHERE PKId = " & intCodigo
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        With adoResultado
            If .EOF Then
                LeImagem LoadPicture()
            Else
                intNumArquivo = FreeFile
                Open App.Path & "\PicTemp" & intCodigo For Binary Access Write As intNumArquivo
                intFragmento = !imgImagem.ActualSize
                bytPedaco() = !imgImagem.GetChunk(intFragmento)
                Put intNumArquivo, , bytPedaco()
                LeImagem = (App.Path & "\PicTemp" & intCodigo)
                Close intNumArquivo
            End If
        End With
    End If
    Screen.MousePointer = vbDefault
    Exit Function
ErroLeImagem:
    Resume FimLeImagem:

FimLeImagem:
    Screen.MousePointer = vbDefault
End Function

Public Function LeExercicio() As Integer
    
    Dim adoResultado  As ADODB.Recordset
    Dim strSql        As String
    strSql = ""

'    strSql = strSql & "SELECT " & gstrISNULL("MAX(intExercicio)", gstrDATEPART(strYEAR, strGETDATE)) & " intExercicio "
'    strSql = strSql & "FROM " & gstrExercicio & " "
'    strSql = strSql & "WHERE bytSituacao = 1"
    
    strSql = strSql & "SELECT " & gstrISNULL("intExercicio", gstrDATEPART(strYEAR, strGETDATE)) & " intExercicio "
    strSql = strSql & "FROM " & gstrUsuarios
    strSql = strSql & " WHERE PKId = " & glngCodUsr
    
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        If adoResultado.EOF = False Then
            LeExercicio = adoResultado!intExercicio
        End If
    End If
End Function

Public Function MostraAno() As Integer
MostraAno = Year(Date)
End Function
Public Function strStoredProcedure(strStoreProcedure As String, _
                                    Optional strParameters As String = vbNullString, _
                                    Optional blnReturnResultset As Boolean = False, _
                                    Optional lngResultSet As Long = 10000) As String

   strStoredProcedure = gstrStoredProcedure(strStoreProcedure, strParameters, blnReturnResultset, lngResultSet)
   
End Function

Public Function CarregaDataDoSistema(Optional blnTime As Boolean, _
                                     Optional blnNaoFormata As Boolean, _
                                     Optional blnOmiteSegundo As Boolean) As String
    
    CarregaDataDoSistema = gstrDataDoSistema(blnTime, blnNaoFormata, blnOmiteSegundo)
    
End Function



Public Function gstrCGCCPFFormatadoRPX(strCGCCPF As String, _
                           Optional strPFPJ As String) As String
    '------------------------------------------------------------------
    ' FUNÇÃO USADA PARA FORMATAR O CGC OU CPF.
    '------------------------------------------------------------------
    ' PARÂMETRO:
    '
    ' 1 - strCGCCPF(Valor digitado - Tipo String)
    ' 2 - strPFPJ(indica se pessoa jurídica -PJ- or Pessoa Física -PF-
    '------------------------------------------------------------------
    
    strCGCCPF = gstrValorSemMascara(strCGCCPF)
    If UCase(strPFPJ) = "PF" Then
        If gblnCPFOk(strCGCCPF) Then
            gstrCGCCPFFormatadoRPX = Format(strCGCCPF, "000\.000\.000\-00")
        Else
            gstrCGCCPFFormatadoRPX = strCGCCPF
        End If
    ElseIf UCase(strPFPJ) = "PJ" Then
        If gblnCGCOk(strCGCCPF) Then
            gstrCGCCPFFormatadoRPX = Format(strCGCCPF, "00\.000\.000\/0000\-00")
        Else
            gstrCGCCPFFormatadoRPX = strCGCCPF
        End If
    ElseIf gblnCGCOk(strCGCCPF) And Len(strCGCCPF) > 11 Then
        gstrCGCCPFFormatadoRPX = Format(strCGCCPF, "0#\.###\.###\/####\-##")
    
    'por Nino
    'ElseIf gblnCPFOk(strCGCCPF) Then
    '    gstrCGCCPFFormatado = Format(strCGCCPF, "###\.###\.###\-##")
    
    ElseIf gblnCPFOk(strCGCCPF) Then
        gstrCGCCPFFormatadoRPX = Format(strCGCCPF, "0##\.###\.###\-##")
    ElseIf Trim(strCGCCPF) = "0" Then
        gstrCGCCPFFormatadoRPX = ""
    Else
        gstrCGCCPFFormatadoRPX = strCGCCPF
    End If
End Function


Public Function gstrCEPFormatadoRPX(vntCepAux As Variant) As String
    '--------------------------------------------------------------'
    ' FUNÇÃO USADA PARA FORMATAR O VALOR DO CEP INFORMADO.         '
    '--------------------------------------------------------------'
    ' PARÂMETRO:                                                   '
    '                                                              '
    ' 1 - vntCepAux(Valor digitado)                                '
    '--------------------------------------------------------------'
    If IsNumeric(vntCepAux) Then
        If Val(vntCepAux) > 0 Then
            If Len(Trim(vntCepAux)) = 7 Then
                vntCepAux = "0" & Trim(vntCepAux)
            ElseIf Len(Trim(vntCepAux)) < 8 Then
                vntCepAux = vntCepAux & String$(8 - Len(Trim(vntCepAux)), "0")
            End If
            gstrCEPFormatadoRPX = Format(vntCepAux, "00000\-000")
        Else
            gstrCEPFormatadoRPX = ""
        End If
    Else
        gstrCEPFormatadoRPX = gstrENulo(vntCepAux)
    End If
End Function

Public Function strConvVrParaSql(vntValor As Variant) As String
    strConvVrParaSql = gstrConvVrParaSql(vntValor)
End Function

Public Function strConvVrDoSql(vntValor As Variant, Optional intCasaDecimal As Integer) As String
    strConvVrDoSql = gstrConvVrDoSql(vntValor, intCasaDecimal)
End Function


Public Function LeExercicioAtual() As Integer
    LeExercicioAtual = gintExercicio
End Function

Public Function LeTituloRelatorio()
    LeTituloRelatorio = frmVisualizarRelatorio.Caption
End Function


Public Function gstrConvDtParaSqlRPX(vntData As Variant, _
                         Optional blnParaSelect As Boolean) As String
    
    '-------------------------------------------------------------
    ' FUNÇÃO USADA PARA CONVERTER A DATA 'vntData' PARA O FORMATO
    ' 'YYYY/MM/DD' OR RETORNAR NULO PARA O SQL
    '-------------------------------------------------------------
    ' PARAMETRO:
    ' 1 - vntData (Data a ser convertida)
    '----------------------------------------------------------
    
    If gblnDataValida(vntData) Then
        Select Case bytDBType
            Case EDatabases.SQLServer
                If blnParaSelect Then
                    gstrConvDtParaSqlRPX = "#" & Format(vntData, "mm/dd/yyyy hh:mm:ss") & "#"
                Else
                    gstrConvDtParaSqlRPX = "'" & Format(vntData, "yyyy/mm/dd hh:mm:ss") & "'"
                End If
            
            Case EDatabases.Oracle
                gstrConvDtParaSqlRPX = gstrFormataDataOracle(CStr(vntData))
        
        End Select
    Else
        gstrConvDtParaSqlRPX = "NULL"
    End If

End Function

Public Function ConvertePraNumero(Valor As String) As Integer
    ConvertePraNumero = Val(Valor)
End Function

Public Function gstrDataPorExtensoExterno(Optional strData As String, _
                                   Optional blnTime As Boolean, _
                                   Optional blnDiaDaSemana As Boolean) As Variant
    Dim strDataAux  As Variant
    
    If gblnDataValida(strData) = False Then
        strData = gstrDataDoSistema(blnTime)
    End If
    If blnDiaDaSemana Then
        strDataAux = strDataAux & gstrDiaDaSemana(strData) & ", "
    End If
    If Day(strData) = 1 Then
        strDataAux = strDataAux & "1º"
    Else
        strDataAux = strDataAux & Day(strData)
    End If
    strDataAux = strDataAux & " de " & gstrNomeDoMes(Month(strData)) & " de "
    strDataAux = strDataAux & Format(strData, "yyyy")
    If blnTime Then
        strDataAux = strDataAux & " - " & Format(strData, "hh:mm:ss")
    End If
    gstrDataPorExtensoExterno = strDataAux
End Function
Public Function LeUsuarioLogado() As String
     LeUsuarioLogado = gstrUsername
End Function
Public Sub TrocaCorDaSecaoParaZebradoRpx(Band)
    TrocaCorDaSecaoParaZebrado Band
End Sub
Public Function RetornaTabRelDocForn() As String
      RetornaTabRelDocForn = gstrTabAutForn
End Function

Public Function gstrExtensoRpx(vntNumero As Variant, _
                   Optional bytTipoDeMoeda As Byte, _
                   Optional bytFeminino As Byte) As String
                   
    gstrExtensoRpx = gstrExtenso(vntNumero, bytTipoDeMoeda, bytFeminino)

End Function

Public Function LeUsuarioDoSistema() As String
LeUsuarioDoSistema = gstrNomeUsuario
End Function
