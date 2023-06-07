Attribute VB_Name = "modAux"
Option Explicit

    'Controla modos de inclusão e alteração nos cadastros
    Public gintModAtual     As Integer
    Public Const ModInsert = 1
    Public Const ModUpdate = 2

    Public Const m_Endereco = 0
    Public Const m_EnderecoCorr = 1
    Public Const m_Observacoes = 2
    Public gintCaption As Integer
    
    Public gstrRetornaDescricao             As String
    Public gstrComplementoEndereco          As String
    Public gstrComplementoEnderecoCorresp   As String
    Public gstrObservacoes                  As String
    Public glngAreaTerreno                  As Long

    Public Const g_CorDesabilitado = &H8000000F
    Public Const g_Habilitado = &H80000005

    Type NomeTipoCampo
        NomeDoCampo As String
        TipoDoCampo As String
    End Type

    Public glngCodigoImovel             As Long


Public Function CarregaComboAtual(pCombo As Object, pQuery As String, pIndice As Integer)
    '---------------------------------------------------------------'
    ' FUNÇÃO USADA PARA CARREGAR UM COMBO NA TELA.                  '
    '---------------------------------------------------------------'
    ' PARÂMETROS:                                                   '
    '                                                               '
    ' 1 - pCombo(ComboBox - Tipo ComboBox)                          '
    ' 2 - pQuery(Query - Tipo String)                               '
    ' 3 - pIndice(Indice do campo da query que sera carregado no    '
    '             Combo - Tipo Integer)                             '
    '---------------------------------------------------------------'
    Dim ADOTemp As ADODB.Recordset
    
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(pQuery, 5, ADOTemp) Then
        pCombo.Clear
        Do While Not ADOTemp.EOF
            If InStr(UCase(pQuery), "TIPOLOGRADOURO") Or InStr(UCase(pQuery), "UFESTADO") Then
                pCombo.AddItem ADOTemp(pIndice)
                pCombo.Tag = ADOTemp(pIndice)
                ADOTemp.MoveNext
            Else
                pCombo.AddItem ADOTemp(pIndice)
                pCombo.ItemData(pCombo.NewIndex) = ADOTemp(0)
                ADOTemp.MoveNext
            End If
        Loop
        ADOTemp.Close
        Set ADOTemp = Nothing
        Set gobjBanco = Nothing
    End If

End Function

Public Function FieldType(intType As Integer) As String
    Select Case intType
        Case adTinyInt
            FieldType = "adTinyInt"
        Case adSmallInt
            FieldType = "adSmallInt"
        Case adInteger
            FieldType = "adInteger"
        Case adBigInt
            FieldType = "adBigInt"
        Case adUnsignedTinyInt
            FieldType = "adUnsignedTinyInt"
        Case adUnsignedSmallInt
            FieldType = "adUnsignedSmallInt"
        Case adUnsignedInt
            FieldType = "adUnsignedInt"
        Case adUnsignedBigInt
            FieldType = "adUnsignedBigInt"
        Case adSingle
            FieldType = "adSingle"
        Case adDouble
            FieldType = "adDouble"
        Case adCurrency
            FieldType = "adCurrency"
        Case adDecimal
            FieldType = "adDecimal"
        Case adNumeric
            FieldType = "adNumeric"
        Case adBoolean
            FieldType = "adBoolean"
        Case adUserDefined
            FieldType = "adUserDefined"
        Case adVariant
            FieldType = "adVariant"
        Case adGUID
            FieldType = "adGUID"
        Case adDate
            FieldType = "adDate"
        Case adDBDate
            FieldType = "adDBDate"
        Case adDBTime
            FieldType = "adDBTime"
        Case adDBTimeStamp
            FieldType = "adDBTimeStamp"
        Case adBSTR
            FieldType = "adBSTR"
        Case adChar
            FieldType = "adChar"
        Case adVarChar
            FieldType = "adVarChar"
        Case adLongVarChar
            FieldType = "adLongVarChar"
        Case adWChar
            FieldType = "adWChar"
        Case adVarWChar
            FieldType = "adVarWChar"
        Case adLongVarWChar
            FieldType = "adLongVarWChar"
        Case adBinary
            FieldType = "adBinary"
        Case adVarBinary
            FieldType = "adVarBinary"
        Case adLongVarBinary
            FieldType = "adLongVarBinary"
    End Select
End Function

Public Sub HabilitaDesabilitaBotaoMenu(blnFlag As Boolean, _
                                       ParamArray Parametro())
    '--------------------------------------------------------------'
    ' SUB USADA PARA HABILITAR OU DESABILITAR BOTÃO NA BARRA DE    '
    ' FERRAMENTA E MENU NA BARRA DE MENU.                          '
    '--------------------------------------------------------------'
    ' PARÂMETROS:                                                  '
    '                                                              '
    ' 1 - tbl(Objeto ToolBar - Tipo Toolbar)                       '
    ' 2 - mnu(Menu - Tipo Menu)                                    '
    ' 3 - I(Índice que identifica o botão - Tipo Integer)          '
    ' 4 - flg(Flag que indica habilitação ou não - Tipo Boolean)   '
    '--------------------------------------------------------------'
End Sub

Sub ExecutaQueryGeral(strQuery As String, _
                      strTabela As String, _
             Optional blnInsert As Boolean)
    
    Set gobjBanco = New clsBanco
    gobjBanco.Execute strQuery
End Sub



