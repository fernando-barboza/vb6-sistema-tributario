VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsConeccao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
    Private ADOConnectString    As String
    
Public Property Get OpenDataADO() As ADODB.Connection
    Set OpenDataADO = gcncADOMain
End Property

Public Property Get CloseDataADO() As ADODB.Connection
    Set CloseDataADO = gcncADOMain
End Property

Public Sub FechaBancoADO()
    '***********************************************************************
    ' Sub para fechar o banco de dados
    '***********************************************************************
    gcncADOMain.Close
    Set gcncADOMain = Nothing
End Sub
