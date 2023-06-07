VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private errRDO          As ADODB.Error
Private pTextoErro      As String
Dim ErrLoop             As ADODB.Error

Public Property Get MostraErro() As String
    MostraErro = pTextoErro
End Property

Public Sub ExibeErro(ErroVB As ErrObject)
    If ErroVB.Number <> 0 Then
        ExibeDetalheErro ""
    End If
    gcncADOMain.Errors.Clear
    Err.Clear
End Sub





