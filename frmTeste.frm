VERSION 5.00
Begin VB.Form frmTeste 
   Caption         =   "Form1"
   ClientHeight    =   7845
   ClientLeft      =   1680
   ClientTop       =   1935
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7845
   ScaleWidth      =   6585
End
Attribute VB_Name = "frmTeste"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const MODELO As String = "CERTIDÃO POSITIVA"

Private Type Campos
    strNumeroProtocolo As String
    strDataAber        As String
    strDataEnc         As String
    strDiaEnc          As String
    strMesEnc          As String
    strAnoEnc          As String
    strObjeto          As String
    strHoraEnc         As String
    strCentrosDeCusto  As String
    strDiaImp          As String
    strMesImp          As String
    strAnoImp          As String
End Type

 Dim mobjAux                  As Object
 Dim Documentos()             As cWordWrapper
 Dim astrRegistros()          As Campos
'Dim WithEvents obfWordEditor As cWordWrapper

Private Sub OpenWordDocumentOverTemplate()
Dim blpMsg          As Boolean
Dim intFor          As Integer
Dim stpDocument     As String
Dim stpTemplate     As String
Dim objFileSystem   As Scripting.FileSystemObject
Dim stpTemplatePath As String
Dim stpDocumentPath As String
Dim strSql As String
Dim adoResultado As ADODB.Recordset
    
    stpTemplate = gstrDirDocumentos & "Documentos\" & App.ProductName & "\WordModelos\" & MODELO & ".dot"
    stpTemplatePath = gstrDirDocumentos & "Documentos\" & App.ProductName & "\WordModelos\"
        
    stpDocumentPath = gstrDirDocumentos & "Documentos\" & App.ProductName & "\WordGravados\"
   
    Set objFileSystem = New Scripting.FileSystemObject
    
    If objFileSystem.FolderExists(stpTemplatePath) Then
    
        If objFileSystem.FileExists(stpTemplate) Then
    
            If objFileSystem.FolderExists(stpDocumentPath) Then
        
                    strSql = "SELECT TOP 10 Pkid, strNome FROM tblContribuinte"
                    Set gobjBanco = New clsBanco
                    
                    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
                    
                        stpDocument = stpDocumentPath & MODELO & "teste" & ".doc"
                        ReDim Documentos(0)
                        If Not adoResultado.EOF Then
                            Set Documentos(0) = New cWordWrapper
                            
                            Documentos(0).GetContainer
                            Documentos(0).DocumentTemplatePath = stpTemplate
                        
                            Documentos(0).DocumentPath = stpDocument
                            Documentos(0).DocumentFormat = wdOpenFormatDocument
                            Documentos(0).DocumentOpen
                                                                                        
                            Documentos(0).DocumentReplaceField "|Inscrição Municipal|", adoResultado!STRNOME
                            
                            Documentos(0).DocumentInsert "|Tabela|", adoResultado
                        
                        End If
                        
                    
                    
                    
                    End If
                    
                    
                    
                    
                    
                    
                            
                    End If
                
'               Else
'                   MsgBox "O Microsoft Word não está instalado nesta máquina. A operação não pode ser realizada.", vbOKOnly + vbInformation, "Mensagem ao usuário"
'               End If
                
            Else
                MsgBox "A pasta de documentos : " & stpDocumentPath & " não foi localizada. A operação não pode ser realizada.", vbOKOnly + vbInformation, "Mensagem ao usuário"
            End If
        
        Else
            MsgBox "O modelo de documento do Microsoft Word : " & stpTemplate & " não foi localizado. A operação não pode ser realizada.", vbOKOnly + vbInformation, "Mensagem ao usuário"
        End If
    
    
    Set objFileSystem = Nothing

End Sub

Public Sub MantemForm(ByVal strModoOperacao As String)
    
    If UCase(strModoOperacao) = UCase(gstrImprimir) Then
        OpenWordDocumentOverTemplate
    End If
    
End Sub

