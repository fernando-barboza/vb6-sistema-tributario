VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBanco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
    Dim mObjErro        As Object
    Dim mobjConeccao    As Object
    
Public Sub ExecutaBeginTrans()
    '--------------------------------------------------------------'
    ' Sub usada para executar o begin transaction no banco.        '
    '--------------------------------------------------------------'
    On Error GoTo Err_BeginTrans
    Screen.MousePointer = vbHourglass
    gcncADOMain.BeginTrans
    
Err_BeginTrans:
    If Err.Number > 0 Then
        Set mObjErro = New clsError
        mObjErro.ExibeErro
        Set mObjErro = Nothing
    End If
    
End Sub

Public Sub ExecutaCommitTrans()
    '--------------------------------------------------------------'
    ' Sub usada para executar o commit transaction no banco.       '
    '--------------------------------------------------------------'
    On Error GoTo Err_CommitTrans
    gcncADOMain.CommitTrans
    
Err_CommitTrans:
    If Err.Number > 0 Then
        Set mObjErro = New clsError
        mObjErro.ExibeErro
        Set mObjErro = Nothing
    End If
    Screen.MousePointer = vbDefault
End Sub

Public Sub ExecutaRollbackTrans()
    '--------------------------------------------------------------'
    ' Sub usada para executar o rollback transaction no banco.     '
    '--------------------------------------------------------------'
    On Error GoTo Err_RollbackTrans
    gcncADOMain.RollbackTrans
Err_RollbackTrans:
    If Err.Number > 0 Then
        Set mObjErro = New clsError
        mObjErro.ExibeErro
        Set mObjErro = Nothing
    End If
    Screen.MousePointer = vbDefault
End Sub

Public Function CriaADO(strQuery As String, _
                        lngQueryTimeOut As Long, _
                        adorsResultado As ADODB.Recordset) As Boolean
    Set mObjErro = New clsError
    On Error GoTo Err_ADO
    
    Dim bytResposta       As Byte
    Dim gcmdADOCmdConMain As ADODB.Command
    Err.Clear
    Set gcmdADOCmdConMain = New ADODB.Command
    
    Set adorsResultado = New ADODB.Recordset
    
    With gcmdADOCmdConMain
        .ActiveConnection = gcncADOMain
        .CommandText = strQuery
        .CommandTimeout = lngQueryTimeOut
        adorsResultado.LockType = adLockReadOnly
        adorsResultado.CursorLocation = adUseClient
        adorsResultado.CursorType = adOpenStatic
        adorsResultado.Open gcmdADOCmdConMain, , , adCmdText
    End With

    CriaADO = True

Err_ADO:
    If Err.Number <> 0 Then
        Select Case InStr(1, UCase(Err.Description), "TIMEOUT EXPIRED")
            Case 0
                If InStr(1, UCase(Err.Description), "ORA-20000") = 0 Then
                    If Err.Number = -2147217871 Then
                        bytResposta = MsgBox(Err.Description & Chr(13) & "Deseja tentar novamente?", vbYesNo)
                        If bytResposta = 6 Then Resume
                    Else
                        ExibeDetalheErro "Retorno dos dados da consulta não completado.", strQuery
                    End If
                Else
                    ExibeMensagem Err.Description
                End If
            Case Else
                ExibeDetalheErro "Congestionamento na rede.", strQuery
        End Select
        Resume FimErroADO
    End If
    Set mObjErro = Nothing
    Set mobjConeccao = Nothing
FimErroADO:
    Err.Clear
End Function

Public Function ExecuteStoredProcedure(strStoredProcedure As String, _
                        lngQueryTimeOut As Long, _
                        Optional adorsResultado As ADODB.Recordset = Nothing, _
                        Optional adopmParameters As ADODB.Parameters = Nothing, _
                        Optional blnExibirMsgErro As Boolean = True) As Boolean

'******************************************************************************************
' Data: 09/04/2003
' Descrição: - strStoredProcedure --> linha de execução completa da stored procedure
'            - lngQueryTimeOut --> tempo limite para execução da stored procedure
'            - adorsResultado --> objeto ADODB.Recordset de retorno da stored procedure
' Alteração: - Implementação da função, a qual tem a função de executar stored procedures
'            porém, para o Banco de Dados Oracle, a função preenche com objetos
'            ADODB.Parameter os parâmetros OUT da stored procedure, uma vez que estes não
'            podem ser opcionais no Oracle ao contrário do SQL Server que permite os mesmos
'            serem opcionais. Caso o parâmetro adorsResultado não seja fornecido a stored
'            procedure será executada sem retorno. Para que a função retorne um objeto
'            ADODB.Recordset o objeto a ser passado no parâmetro adorsResultado deve ser
'            instanciado com a palavra-chave NEW.
' Responsável: Everton Bianchini
'------------------------------------------------------------------------------------------
' Data: 07/05/2003
' Descrição: - adopmParameters --> conjunto de parâmetros retornados pela stored procedure
' Alteração: - Adiconado o parâmetro adopmParameters, o qual tem o objetivo de retornar
'            uma coleção de parâmetros retornados pela stored procedure. Caso a procedure
'            não tenha parâmetros do tipo OUT a variável adopmParameters será nothing.
' Responsável: Everton Bianchini
'******************************************************************************************

    Set mObjErro = New clsError
    On Error GoTo Err_ADO
    
    Dim cmdExecSP As ADODB.Command
    Dim recParamsSP As ADODB.Recordset
    
    Dim strSPTemp As String
    Dim strSPName As String
    Dim strPKGName As String
    Dim strParamResultSet As String
    Dim strArrParams() As String
    
    Dim strSql As String
    
    Dim intInd As Integer
    Dim intIndTmp As Integer
    Dim intCountParameters As Integer
    Dim intPosition As Integer
    Dim blnStop As Boolean

    Dim strParamName As String
    Dim lngParamType As ADODB.DataTypeEnum
    Dim lngParamLength As Long


    Err.Clear

    strSPTemp = LTrim(strStoredProcedure)
    
    Set cmdExecSP = New ADODB.Command
    
'    If (bytDBType = EDatabases.Oracle) Then
    
        If (InStr(1, strSPTemp, " call ", vbTextCompare) > 0) Then
            strSPTemp = Mid(strSPTemp, InStr(1, strSPTemp, "call", vbTextCompare) + 4)
            strSPTemp = StrReverse(strSPTemp)
            strSPTemp = Mid(strSPTemp, InStr(1, strSPTemp, "}", vbTextCompare) + 2)
            strSPTemp = Trim(StrReverse(strSPTemp))
        End If
        
'        strSPName = Trim(Mid(strSPTemp, 1, InStr(1, strSPTemp, "(", vbTextCompare) - 1))
        strSPName = Trim(Mid(strSPTemp, 1, InStr(1, strSPTemp, IIf((bytDBType = EDatabases.Oracle), "(", " "), vbTextCompare) - 1))
        If (InStr(1, strSPName, ".", vbTextCompare) > 0) Then
            strPKGName = Mid(strSPName, 1, InStr(1, strSPName, ".", vbTextCompare))
            strSPName = Mid(strSPName, InStr(1, strSPName, ".", vbTextCompare) + 1)
        End If
'        strSPTemp = Mid(strSPTemp, InStr(1, strSPTemp, "(", vbTextCompare) + 1)
        strSPTemp = Mid(strSPTemp, InStr(1, strSPTemp, IIf((bytDBType = EDatabases.Oracle), "(", " "), vbTextCompare) + 1)
        If (bytDBType = EDatabases.Oracle) Then
            strSPTemp = Trim(Mid(strSPTemp, 1, InStrRev(strSPTemp, ")", , vbTextCompare) - 1))
        Else
            strSPTemp = RTrim(strSPTemp)
        End If
        
'        strArrParams = Split(strSPTemp, ",", , vbTextCompare)
        strArrParams = strarrSplit(strSPTemp)
        
        If (bytDBType = EDatabases.SQLServer) Then
            strSql = "SELECT c.name argument_name, UPPER(t.name) data_type, "
            strSql = strSql & "(" & gstrCASEWHEN("c.isoutparam", "1,'OUT'", "'IN'") & ") in_out, "
            strSql = strSql & "c.colorder position, c.length "
            strSql = strSql & "FROM sysobjects o, syscolumns c, systypes t "
            strSql = strSql & "WHERE t.xtype = c.xtype AND "
            strSql = strSql & "c.id = o.id AND "
            strSql = strSql & "UPPER(t.name) <> 'SYSNAME' AND "
            strSql = strSql & "UPPER(o.name) = '" & UCase(strSPName) & "' "
            strSql = strSql & "ORDER BY c.colorder "
        
        ElseIf (bytDBType = EDatabases.Oracle) Then
            strSql = "select argument_name, data_type, in_out, position From all_arguments "
            strSql = strSql & "where owner = 'CPDMASTER' and "
            strSql = strSql & "object_name = '" & UCase(strSPName) & "' and "
            strSql = strSql & "(not (argument_name is null)) "
            strSql = strSql & "order by sequence "
        
        End If
        
        If CriaADO(strSql, 5, recParamsSP) Then
            
            If Not recParamsSP.EOF Then
                intInd = UBound(strArrParams())
    
                recParamsSP.MoveLast
                intCountParameters = (recParamsSP("position") - 1)
                ReDim Preserve strArrParams(intCountParameters)
                recParamsSP.MoveFirst
                
                Do
                
                    If (recParamsSP("in_out") = "OUT") And _
                        (((recParamsSP("data_type") <> "PL/SQL TABLE") And _
                          (recParamsSP("data_type") <> "REF CURSOR") And _
                          (bytDBType = EDatabases.Oracle)) Or _
                         (bytDBType = EDatabases.SQLServer)) Then
                        
                        intPosition = (recParamsSP("position") - 1)
                        intPosition = IIf((intPosition < 0), 0, intPosition)
                        
                        If (intPosition <= intCountParameters) Then
                            If (LTrim(strArrParams(UBound(strArrParams()))) = "") Then
                                
                                For intIndTmp = UBound(strArrParams()) To LBound(strArrParams()) Step -1
                                    If (LTrim(strArrParams(intIndTmp)) <> "") Then Exit For
                                Next intIndTmp
                                
                                For intIndTmp = intIndTmp To intPosition Step -1
                                    strArrParams(intIndTmp + 1) = strArrParams(intIndTmp)
                                Next
                            End If
                            strArrParams(intPosition) = "?"
                        
                        Else
                            strArrParams(intInd) = "?"
                        
                        End If
                        
                        If (bytDBType = EDatabases.SQLServer) Then
                            Select Case recParamsSP("data_type")
                            
                                Case "BIT"
                                    lngParamType = adBoolean
                                
                                Case "CHAR"
                                    lngParamType = adChar
                                    lngParamLength = recParamsSP("length")
                                
                                Case "DATETIME", "SMALLDATETIME"
                                    lngParamType = adDBTimeStamp
                                
                                Case "DECIMAL"
                                    lngParamType = adDecimal
                                
                                Case "IMAGE"
                                    lngParamType = adLongVarBinary
                                
                                Case "INT"
                                    lngParamType = adInteger
                                
                                Case "MONEY"
                                    lngParamType = adCurrency
                                
                                Case "NTEXT"
                                    lngParamType = adLongVarWChar
                                
                                Case "NUMERIC"
                                    lngParamType = adNumeric
                                
                                Case "NVARCHAR"
                                    lngParamType = adVarWChar
                                    lngParamLength = recParamsSP("length") / 2
                                
                                Case "SMALLINT"
                                    lngParamType = adSmallInt
                                
                                Case "TINYINT"
                                    lngParamType = adUnsignedTinyInt
                                
                                Case "VARCHAR"
                                    lngParamType = adVarChar
                                    lngParamLength = recParamsSP("length")
                                
                                Case Else
                                    strArrParams(recParamsSP("position") - 1) = ""
                                    
                                    lngParamType = -20000
                            
                            End Select
                        
                        ElseIf (bytDBType = EDatabases.Oracle) Then
                            Select Case recParamsSP("data_type")
                                
                                Case "NUMBER", "FLOAT"
                                    lngParamType = adNumeric
                                
                                Case "VARCHAR2"
                                    lngParamType = adVarChar
                                    lngParamLength = 600 'Definido valor fixo pois no oracle nao é possivel declarar tamanho de parametros
                                
                                Case "DATE"
                                    lngParamType = adDate
                                
                                Case Else
                                    strArrParams(recParamsSP("position") - 1) = ""
                                    
                                    lngParamType = -20000
                            
                            End Select
                        
                        End If
                        
                        If (lngParamType <> -20000) Then
                            If (bytDBType = EDatabases.Oracle) Then
                                strParamName = recParamsSP("argument_name")
                            
                            ElseIf (bytDBType = EDatabases.SQLServer) Then
                                'Retira sinal de @
                                strParamName = Mid(recParamsSP("argument_name"), 2)
                            
                            End If
                            
                            Select Case lngParamType
                                
                                Case ADODB.DataTypeEnum.adVarChar, ADODB.DataTypeEnum.adVarWChar, _
                                    ADODB.DataTypeEnum.adChar
                                    cmdExecSP.Parameters.Append cmdExecSP.CreateParameter(strParamName, lngParamType, adParamOutput, lngParamLength)
                                
                                Case Else
                                    cmdExecSP.Parameters.Append cmdExecSP.CreateParameter(strParamName, lngParamType, adParamOutput)
                            
                            End Select
                        
                        End If
                        
                    End If
                
                    recParamsSP.MoveNext
                
                Loop Until recParamsSP.EOF
                
                recParamsSP.Close
                Set recParamsSP = Nothing
                
            End If
            
        End If
        
        strSPTemp = strPKGName & strSPName & "("
        For intInd = LBound(strArrParams()) To UBound(strArrParams())
            If Trim(strArrParams(intInd)) <> "" Then
                If intInd > LBound(strArrParams()) Then
                    strSPTemp = strSPTemp & ", "
                End If
                strSPTemp = strSPTemp & strArrParams(intInd)
            End If
        Next intInd
        strSPTemp = "{ call " & strSPTemp & ") } "
    
'    End If
    
    With cmdExecSP
        .ActiveConnection = gcncADOMain
        .CommandTimeout = lngQueryTimeOut
        .CommandText = strSPTemp
        .CommandType = adCmdText
        If Not (adorsResultado Is Nothing) Then
            Set adorsResultado = .Execute
        Else
            .Execute
        End If
        
        If (.Parameters.Count > 0) Then
            Set adopmParameters = .Parameters
        End If
        
    End With

    ExecuteStoredProcedure = True

Err_ADO:
    If Err.Number <> 0 Then
        Select Case InStr(1, UCase(Err.Description), "TIMEOUT EXPIRED")
            Case 0
                If blnExibirMsgErro Then
                    If InStr(1, UCase(Err.Description), "ORA-20000") = 0 Then
                        ExibeDetalheErro "Execução da Stored Procedure não completado.", strSPTemp
                    Else
                        ExibeMensagem Err.Description
                    End If
                Else
                    gstrErrorInStoredProcedure = Err.Description
                End If
            Case Else
                ExibeDetalheErro "Congestionamento na rede.", strSPTemp
        End Select
        Resume FimErroADO
    End If
    Set mObjErro = Nothing
FimErroADO:
    Err.Clear
End Function
Public Function Execute(strQuery As String, _
               Optional blnNaoExibeMensagem As Boolean) As Boolean
    Dim ErrLoop As ADODB.Error
    Dim strQueryAux As String
    Dim gcmdADOCmdConMain As ADODB.Command
    
    On Error GoTo ErroExecute
    strQueryAux = strQuery
    
    Set gcmdADOCmdConMain = New ADODB.Command
    Set gcmdADOCmdConMain.ActiveConnection = gcncADOMain
    gcmdADOCmdConMain.CommandText = strQuery
    gcmdADOCmdConMain.Execute strQuery, , adExecuteNoRecords
    Execute = True
    
ErroExecute:
    If blnNaoExibeMensagem = False And Err <> 0 Then
        Select Case InStr(1, UCase(Err.Description), "TIMEOUT EXPIRED")
            Case 0
                If InStr(1, UCase(Err.Description), "ORA-20000") = 0 Then
                    ExibeDetalheErro gstrMsgErroADO(Err, strQueryAux), strQueryAux
                Else
                    ExibeMensagem Err.Description
                End If
            Case Else
                ExibeDetalheErro "Congestionamento na rede."
        End Select
        Resume FimExecute
    End If
FimExecute:
End Function

Public Function GetRecordset() As ADODB.Recordset
    Dim adoResultado     As ADODB.Recordset
    
    Set gobjBanco = New clsBanco
    
   'gobjBanco.CriaADO "SELECT * FROM tbCatalogoMaterialServico", 5, adoResultado
    gobjBanco.CriaADO "SELECT * FROM " & gstrCatalogoMaterialServico, 5, adoResultado
    
    Set GetRecordset = adoResultado

End Function

