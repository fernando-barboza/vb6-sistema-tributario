VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cWordWrapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

#Const WORD_VERSAO_97 = False

Private Declare Function BringWindowToTop Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long

Private Declare Function RegOpenKey Lib "advapi32" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegCloseKey& Lib "advapi32" (ByVal hKey&)

Private Declare Function FindWindow& Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String)

Private Const HKEY_CLASSES_ROOT = &H80000000

Private Const REG_SZ = 1
Private Const REG_EXPAND_SZ = 2
Private Const ERROR_SUCCESS = 0

Private Const WORD_CAPTION = "Documento : "
Private Const WORD_CLASS_97 = "OpusApp"
Private Const WORD_CLASS_2000 = "OpusApp"

Private Const SM_CXForm = 32
Private Const SM_CYForm = 33
Private Const SM_CYCAPTION = 4

Private Type RECT
    Left   As Long
    Top    As Long
    Right  As Long
    Bottom As Long
End Type

Private blfWordClose                  As Boolean
Private blfNewDocument                As Boolean
Private blfNewWindow                  As Boolean
Private stfDocumentPath               As String
Private vrfDocumentFormat             As Variant
Private stfDocumentTemplatePath       As String
Dim WithEvents obfWordWindowContainer As Form
Attribute obfWordWindowContainer.VB_VarHelpID = -1

Private tyfWordWindowRect             As RECT
Private lgfWordWindowHwnd             As Long
Private lgfWordContainerWindowHwnd    As Long

Private blfShowWordWindowRecentFiles  As Boolean
Private blfShowWordWindowStatusBar    As Boolean

Dim obfWordDocument                   As Object
Attribute obfWordDocument.VB_VarHelpID = -1
Dim obfWordApplication                As Object
Attribute obfWordApplication.VB_VarHelpID = -1

Public Event Quit()

Private Sub Class_Initialize()

On Error GoTo Problema_Na_Rotina

   blfNewDocument = False
   stfDocumentPath = Space$(0)
   stfDocumentTemplatePath = Space$(0)
   vrfDocumentFormat = WORDOPENFORMATDOCUMENT
   
   lgfWordWindowHwnd = 0
   lgfWordContainerWindowHwnd = 0
   
   Set obfWordWindowContainer = Nothing
   Set obfWordDocument = Nothing
   Set obfWordApplication = Nothing

   blfWordClose = False

   Exit Sub

Problema_Na_Rotina:
   
'  If RecoverError("Class_Initialize") Then Resume
   
   ExibeDetalheErro "Erro na rotina Class_Initialize."

End Sub

Private Sub Class_Terminate()

On Error GoTo Problema_Na_Rotina

   CloseWordInstance
   
   Exit Sub

Problema_Na_Rotina:

'  If RecoverError("Class_Terminate") Then Resume
   
   ExibeDetalheErro "Erro na rotina Class_Terminate."
   
End Sub

Private Sub ConfigureWordInstance(blpOn As Boolean)
Dim obpWordDocument As Object
    #If WORD_CLASS_2000 Then
        Set obpWordDocument = obfWordApplication.Documents.Add()
    
        If blpOn Then
        
            blfShowWordWindowStatusBar = obfWordApplication.DisplayStatusBar
          
            blfShowWordWindowRecentFiles = obfWordApplication.DisplayRecentFiles
    
            obfWordApplication.DisplayRecentFiles = False
            obfWordApplication.DisplayStatusBar = True
        
            With obfWordApplication
                
                
                .CommandBars("Menu Bar").Reset
                
                .CommandBars("File").Reset
                .CommandBars("File").Controls(1).Visible = False
                .CommandBars("File").Controls(2).Visible = False
                .CommandBars("File").Controls(3).Visible = False
                .CommandBars("File").Controls(5).Visible = False
                .CommandBars("File").Controls(6).Visible = False
                .CommandBars("File").Controls(7).Visible = False
                .CommandBars("File").Controls(8).Visible = False
                .CommandBars("File").Controls(12).Visible = False
                .CommandBars("File").Controls(13).Visible = False
                
                .CommandBars("Menu Bar").Controls(9).Visible = False
                
                .CommandBars("Standard").Controls(1).Visible = False
                .CommandBars("Standard").Controls(2).Visible = False
                .CommandBars("Standard").Controls(4).Visible = False
                
            End With
            
        Else
            
            With obfWordApplication
            
                
            
                .DisplayRecentFiles = blfShowWordWindowRecentFiles
                .DisplayStatusBar = blfShowWordWindowStatusBar
            
                .CommandBars("Menu Bar").Reset
                .CommandBars("File").Reset
                .CommandBars("Standard").Reset
                
            End With
            
        End If
          
        obfWordApplication.NormalTemplate.Saved = True
        
        obpWordDocument.Saved = True: obpWordDocument.Close
   #End If
    Set obpWordDocument = Nothing

End Sub

Private Function GetRegistryString(ByVal lgpKey As Long, ByVal stpSubKey As String, ByVal stpValueName As String) As String
Dim lgpRes     As Long
Dim stpSetting As String
Dim lgpDataLen As Long
   
On Error GoTo Problema_Na_Rotina

   If RegOpenKey(lgpKey, stpSubKey, lgpRes) = ERROR_SUCCESS Then
   
      stpSetting = Space(255): lgpDataLen = Len(stpSetting)
      
      If RegQueryValueEx(lgpRes, stpValueName, ByVal 0, REG_EXPAND_SZ, ByVal stpSetting, lgpDataLen) = ERROR_SUCCESS Then
         If lgpDataLen > 1 Then GetRegistryString = Left(stpSetting, lgpDataLen - 1)
      End If

      If RegCloseKey(lgpRes) <> ERROR_SUCCESS Then GetRegistryString = Space$(0)
         
   End If
   
   Exit Function

Problema_Na_Rotina:

'  If RecoverError("GetRegistryString") Then Resume
   
   ExibeDetalheErro "Erro na rotina GetRegistryString."
   
End Function

Private Sub GetWordInstance()

On Error GoTo Problema_Na_Rotina
   
   If lgfWordWindowHwnd = 0 Or obfWordApplication Is Nothing Then
      
      Set obfWordApplication = Nothing

      Set obfWordApplication = CreateObject("Word.Application") 'New Word.Application
      
      obfWordApplication.Visible = False
      
      #If WORD_VERSAO_97 Then
         lgfWordWindowHwnd = FindWindow(WORD_CLASS_97, vbNullString)
      #Else
         lgfWordWindowHwnd = FindWindow(WORD_CLASS_2000, vbNullString)
      #End If
      
      GetWindowRect lgfWordWindowHwnd, tyfWordWindowRect

      ConfigureWordInstance True
   
    End If
   
    Exit Sub

Problema_Na_Rotina:

'  If RecoverError("GetWordInstance") Then Resume
   
   ExibeDetalheErro "Erro na rotina GetWordInstance."
   
End Sub

Private Sub CloseWordInstance()

On Error GoTo Problema_Na_Rotina

   DeAttachContainerWindow
   
   If Not obfWordWindowContainer Is Nothing Then
      Unload obfWordWindowContainer: Set obfWordWindowContainer = Nothing
   End If
      
   Set obfWordDocument = Nothing
   
   Set obfWordApplication = Nothing

   Exit Sub

Problema_Na_Rotina:

'  If RecoverError("CloseWordInstance") Then Resume
   
   ExibeDetalheErro "Erro na rotina CloseWordInstance."

End Sub

Private Sub AttachContainerWindow()

On Error GoTo Problema_Na_Rotina
   
   If lgfWordContainerWindowHwnd <> 0 Then SetParent lgfWordWindowHwnd, lgfWordContainerWindowHwnd
      
   obfWordApplication.WindowState = 1 'wdWindowStateMaximize
   
   Exit Sub

Problema_Na_Rotina:

'  If RecoverError("AttachContainerWindow") Then Resume
   
   ExibeDetalheErro "Erro na rotina AttachContainerWindow."

End Sub

Private Sub DeAttachContainerWindow()

On Error GoTo Problema_Na_Rotina

   If lgfWordContainerWindowHwnd <> 0 Then

      SetParent lgfWordWindowHwnd, 0
      
      With tyfWordWindowRect
         MoveWindow lgfWordWindowHwnd, .Left, .Top, .Right - .Left, .Bottom - .Top, True
      End With
   
   End If
   
   Exit Sub

Problema_Na_Rotina:

'  If RecoverError("DeAttachContainerWindow") Then Resume
   
   ExibeDetalheErro "Erro na rotina DeAttachContainerWindow."

End Sub

Private Sub ResizeContainerWindow()
Dim lgpX                 As Long
Dim lgpY                 As Long
Dim lgpW                 As Long
Dim lgpH                 As Long
Dim typWordContainerRect As RECT

On Error GoTo Problema_Na_Rotina

   If lgfWordContainerWindowHwnd <> 0 Then
   
      ' Obtem as dimensoes do container do word (o formulario vb ...)
      GetWindowRect lgfWordContainerWindowHwnd, typWordContainerRect
   
      lgpX = -GetSystemMetrics(SM_CXForm)
      lgpY = -GetSystemMetrics(SM_CYForm)
      
      lgpW = (typWordContainerRect.Right - typWordContainerRect.Left - lgpX * 2) - 10
      lgpH = (typWordContainerRect.Bottom - typWordContainerRect.Top - lgpY * 2) - 30
      
      lgpY = lgpY - GetSystemMetrics(SM_CYCAPTION)
      lgpH = lgpH + GetSystemMetrics(SM_CYCAPTION)
      
      ' Coloca a janela da instancia do servidor word em primeiro plano
      BringWindowToTop lgfWordWindowHwnd
         
      ' Faz a janela da instancia do servidor word ocupar toda a area de seu container (o formulario vb ...)
      MoveWindow lgfWordWindowHwnd, lgpX, lgpY, lgpW, lgpH, True
         
      ' Coloca a janela da instancia do servidor do word em primeiro plano dentro de seu container
      SetForegroundWindow lgfWordWindowHwnd
   
   End If
   
   blfWordClose = True
   
   Exit Sub

Problema_Na_Rotina:

'  If RecoverError("ResizeContainerWindow") Then Resume
   
   ExibeDetalheErro "Erro na rotina ResizeContainerWindow."
   
End Sub
   
Public Property Let DocumentPath(ByVal vNewValue As String)

On Error GoTo Problema_Na_Rotina

   stfDocumentPath = vNewValue

   Exit Property

Problema_Na_Rotina:

'  If RecoverError("DocumentPath") Then Resume
   
   ExibeDetalheErro "Erro na rotina DocumentPath."

End Property

Public Property Let DocumentTemplatePath(ByVal vNewValue As String)

On Error GoTo Problema_Na_Rotina

   stfDocumentTemplatePath = vNewValue

   Exit Property

Problema_Na_Rotina:

'  If RecoverError("DocumentTemplatePath") Then Resume
   
   ExibeDetalheErro "Erro na rotina DocumentTemplatePath."

End Property

Public Property Let DocumentFormat(ByVal vNewValue As Variant)

On Error GoTo Problema_Na_Rotina

   vrfDocumentFormat = vNewValue

   Exit Property

Problema_Na_Rotina:

'  If RecoverError("DocumentFormat") Then Resume
   
   ExibeDetalheErro "Erro na rotina DocumentFormat."

End Property

Public Property Get NewDocument() As Boolean

On Error GoTo Problema_Na_Rotina

   NewDocument = blfNewDocument
    
   Exit Property

Problema_Na_Rotina:

'  If RecoverError("NewDocument") Then Resume
   
   ExibeDetalheErro "Erro na rotina NewDocument."
    
End Property

Public Sub GetContainer()

On Error GoTo Problema_Na_Rotina

   Set obfWordWindowContainer = New frmWordWrapper

   lgfWordContainerWindowHwnd = obfWordWindowContainer.hwnd
   
   Exit Sub

Problema_Na_Rotina:

'  If RecoverError("Container") Then Resume
   
   ExibeDetalheErro "Erro na rotina GetContainer."

End Sub

Public Function IsInstalled() As Boolean
Dim stpSubKey    As String
Dim stpValueName As String

On Error GoTo Problema_Na_Rotina

   stpSubKey = "Word.Document\CurVer": stpValueName = Space$(0)
  'stpSubKey = "PowerPoint.Slide\CurVer": stpValueName = Space$(0)
  'stpSubKey = "Excel.Sheet\CurVer": stpValueName = space$(0)
  'stpSubKey = "Access.Database\CurVer": stpValueName = space$(0)
  
   IsInstalled = CBool(Len(GetRegistryString(HKEY_CLASSES_ROOT, stpSubKey, stpValueName)))
  
   Exit Function

Problema_Na_Rotina:

'  If RecoverError("IsInstalled") Then Resume
   
   ExibeDetalheErro "Erro na rotina IsInstalled."
   
End Function

Public Sub DocumentOpen()

On Error GoTo Problema_Na_Rotina
               
   GetWordInstance
   
   If Dir(stfDocumentPath, vbNormal) <> Space$(0) Then
         
     'Set obfWordDocument = obfWordApplication.Documents.Open(stfDocumentPath, , , , , , , , , vrfDocumentFormat)
      Set obfWordDocument = obfWordApplication.Documents.Open(stfDocumentPath)
     
      blfNewDocument = False
      
   Else
            
      If stfDocumentTemplatePath <> Space$(0) Then
                  
         If Dir(stfDocumentTemplatePath, vbNormal) <> Space$(0) Then
         
           'Set obfWordDocument = obfWordApplication.Documents.Add(stfDocumentTemplatePath)
            Set obfWordDocument = obfWordApplication.Documents.Open(stfDocumentTemplatePath)
            
            obfWordDocument.SaveAs stfDocumentPath, vrfDocumentFormat
            
            blfNewDocument = True
         
         Else
            MsgBox "O Modelo " & stfDocumentTemplatePath & " não foi localizado. A operação não pode ser realizada.", vbOKOnly + vbInformation, "Mensagem ao usuário"
            Exit Sub
         End If
                  
      Else
      
        'Set obfWordDocument = obfWordApplication.Documents.Add()
         Set obfWordDocument = obfWordApplication.Documents.Open(stfDocumentPath)
         
        'obfWordDocument.SaveAs stfDocumentPath, vrfDocumentFormat
         
         blfNewDocument = True
         
      End If
            
   End If

   obfWordWindowContainer.Caption = WORD_CAPTION & stfDocumentPath

   obfWordWindowContainer.Visible = True

   obfWordApplication.Visible = True

   AttachContainerWindow

   ResizeContainerWindow
      
   Exit Sub
   
Problema_Na_Rotina:

'  If RecoverError("DocumentOpen") Then Resume
   
   ExibeDetalheErro "Erro na rotina DocumentOpen."
      
End Sub

Public Sub DocumentReplaceField(ByVal stpField As String, ByVal stpValue As String)

On Error GoTo Problema_Na_Rotina

   obfWordApplication.Documents(stfDocumentPath).Activate

   With obfWordApplication.Selection.Find
      
      .ClearFormatting
      
      .Forward = True
      .MatchWholeWord = True
      .MatchCase = False
      .Wrap = 1 'wdFindContinue

        
'      .text = stpField
'      .Replacement.text = stpValue
        
      .Execute FindText:=stpField, ReplaceWith:=stpValue, Format:=True, Replace:=2 ' wdReplaceAll

   End With

   Exit Sub
   
Problema_Na_Rotina:

'  If RecoverError("DocumentReplaceField") Then Resume
   
   ExibeDetalheErro "Erro na rotina DocumentReplaceField."

End Sub

Public Sub Quit()

On Error GoTo Problema_Na_Rotina

   If Not obfWordWindowContainer Is Nothing Then Unload obfWordWindowContainer

   Exit Sub
   
Problema_Na_Rotina:

'  If RecoverError("Quit") Then Resume
   
   ExibeDetalheErro "Erro na rotina Quit."

End Sub

Private Sub obfWordWindowContainer_Load()

On Error GoTo Problema_Na_Rotina
      
   obfWordWindowContainer.Top = 0
   obfWordWindowContainer.Left = 30
      
   obfWordWindowContainer.WindowState = vbMaximized
   
   TrocaInconiDoObj obfWordWindowContainer, 3
   
   Exit Sub

Problema_Na_Rotina:

'  If RecoverError("obfWordWindowContainer_Load") Then Resume
   
   ExibeDetalheErro "Erro na rotina obfWordWindowContainer_Load."

End Sub

Private Sub obfWordWindowContainer_Activate()

On Error GoTo Problema_Na_Rotina

   obfWordWindowContainer.Caption = WORD_CAPTION & stfDocumentPath
   
   HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrNovo, gstrSalvar, gstrDeletar, gstrAplicar, gstrImprimir, gstrLocalizar, gstrPreencherLista
   
   Exit Sub

Problema_Na_Rotina:

'  If RecoverError("obfWordWindowContainer_Activate") Then Resume
   
   ExibeDetalheErro "Erro na rotina obfWordWindowContainer_Activate."

End Sub

Private Sub obfWordWindowContainer_Resize()

On Error GoTo Problema_Na_Rotina

   ResizeContainerWindow
   
   Exit Sub

Problema_Na_Rotina:

'  If RecoverError("obfWordWindowContainer_Resize") Then Resume
   
   ExibeDetalheErro "Erro na rotina obfWordWindowContainer_Resize."
   
End Sub

Private Sub obfWordWindowContainer_Deactivate()

On Error GoTo Problema_Na_Rotina

   HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrNovo, gstrSalvar, gstrAplicar, gstrImprimir, gstrLocalizar, gstrPreencherLista
   
   Exit Sub

Problema_Na_Rotina:

   ExibeDetalheErro "Erro na rotina obfWordWindowContainer_Deactive."

End Sub

Private Sub obfWordWindowContainer_QueryUnload(Cancel As Integer, UnloadMode As Integer)

On Error GoTo Problema_Na_Rotina

   'If UnloadMode <> vbFormCode Then obfWordApplication.Quit

   HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrNovo, gstrSalvar, gstrAplicar, gstrImprimir, gstrLocalizar, gstrPreencherLista
   
   Exit Sub

Problema_Na_Rotina:

'  If RecoverError("obfWordWindowContainer_QueryUnload") Then Resume
   
   ExibeDetalheErro "Erro na rotina obfWordWindowContainer_QueryUnload."

End Sub

Private Sub obfWordDocument_Close()

On Error GoTo Problema_Na_Rotina
    #If WORD_CLASS_2000 Then
        If blfWordClose Then
            obfWordApplication.Quit
        Else
            obfWordApplication_Quit
        End If
    #End If

   Exit Sub

Problema_Na_Rotina:

'  If RecoverError("obfWordDocument_Close") Then Resume
   
   ExibeDetalheErro "Erro na rotina obfWordDocument_Close."

End Sub

Private Sub obfWordApplication_Quit()
  
On Error GoTo Problema_Na_Rotina

   ConfigureWordInstance False
   
   CloseWordInstance
   
   RaiseEvent Quit

   Exit Sub

Problema_Na_Rotina:

'  If RecoverError("obfWordApplication_Quit") Then Resume
   
   ExibeDetalheErro "Erro na rotina obfWordApplication_Quit."
   
End Sub

Public Sub DocumentInsert(ByVal strField As String, Optional ByVal adoRecordset As ADODB.Recordset, _
                            Optional TabelaArray As XArrayDB, Optional AlinhaColunas As XArrayDB, Optional intTabela As Integer = 1, _
                            Optional blnBorda As Boolean = False)
'Método criado por Nino
'AlinhaColunas - Variável que contém o Número do Campo e o tipo de alinhamento para ser aplicado a cada coluna
'da tabela criada.
'EX. 1 - wdAlignParagraphRight
'    2 - wdAlignParagraphLeft

'Vamos concatenar strings pre-definidas de formatacao no inicio do array TabelaArray
'Valores pre definidos para formatacao : #BOLD# (FORMAT_NEGRITO) - Negrito

'intTabela - Variável que define qual tabela será criada
'blnBorda  - Variável que define se devem ser aplicadas ou não bordas na tabela

Dim intCont     As Integer
Dim intColuna   As Integer
Dim intLinha    As Integer

'Vamos armazenar a qtde de caracteres que vamos excluir da string, pois representavam a formatacao
Dim intCaracFormatParaExcluir As Integer

    On Error GoTo Problema_Na_Rotina
    
    intCaracFormatParaExcluir = 1
    
    obfWordApplication.Documents(stfDocumentPath).Activate

    If TabelaArray Is Nothing Then
        If Not adoRecordset.EOF Then
            adoRecordset.MoveFirst
            'Substitui a tabela a ser criada para o campo Field passado
            With obfWordApplication.Selection.Find
                .Forward = True
                .ClearFormatting
                .MatchWholeWord = True
                .MatchCase = False
                .Wrap = 1 'wdFindContinue
                If Not .Execute(FindText:="|" & strField & "|") Then
                    Exit Sub
                End If
            End With
                
            'Cria a Tabela
            obfWordApplication.ActiveDocument.Tables.Add Range:=obfWordApplication.Selection.Range, _
                NumRows:=adoRecordset.RecordCount + 1, NumColumns:=adoRecordset.Fields.Count
            
            
            If blnBorda Then
                obfWordApplication.ActiveDocument.Tables(intTabela).AutoFormat applyborders:=True, Format:=16
            Else
                'Nao podemos aplicar esta formatacao apos preenchido os campos, pois perde a formtacao de fonte
                obfWordApplication.ActiveDocument.Tables(intTabela).AutoFormat applyborders:=False
            End If
               
            'Cria o cabeçalho da Tabela(os nomes dos campos)
            For intCont = 1 To adoRecordset.Fields.Count
                obfWordApplication.ActiveDocument.Tables(intTabela).Cell(Row:=1, Column:=intCont).Range.InsertAfter _
                    adoRecordset.Fields(intCont - 1).Name
            Next
            
            'Preenche a tabela
            Do While Not adoRecordset.EOF
                For intCont = 0 To adoRecordset.Fields.Count - 1
                    
                    obfWordApplication.ActiveDocument.Tables(intTabela).Columns(intCont + 1).Cells(adoRecordset.AbsolutePosition + 1).Select
                    
                    'Vamos verificar se ha formatacao
                    If InStr(1, adoRecordset.Fields(intCont), FORMAT_NEGRITO) > 0 Then
                        obfWordApplication.Selection.Font.Bold = True
                        intCaracFormatParaExcluir = intCaracFormatParaExcluir + Len(FORMAT_NEGRITO)
                    Else
                        obfWordApplication.Selection.Font.Bold = False
                        intCaracFormatParaExcluir = 1
                    End If
                    
                    obfWordApplication.ActiveDocument.Tables(intTabela).Cell(Row:=adoRecordset.AbsolutePosition + 1, _
                        Column:=intCont + 1).Range.InsertAfter Mid(adoRecordset.Fields(intCont), intCaracFormatParaExcluir)
                    
                    'Alinha as determinadas colunas com o array de parâmetro passado.
                    If Not AlinhaColunas Is Nothing Then
                        obfWordApplication.Selection.Paragraphs.Alignment = AlinhaColunas.Value(0, intCont)
                    End If
                    
                Next
                adoRecordset.MoveNext
            Loop
            
            'obfWordApplication.ActiveDocument.Tables(intTabela).AutoFitBehavior (wdAutoFitContent)
            obfWordApplication.ActiveDocument.Tables(intTabela).AutoFormat AutoFit:=True
            
        End If
    Else
        If TabelaArray.Count(1) > 0 Then
            'Substitui a tabela a ser criada para o campo Field passado
            With obfWordApplication.Selection.Find
                .Forward = True
                .ClearFormatting
                .MatchWholeWord = True
                .MatchCase = False
                .Wrap = 1 'wdFindContinue
                If Not .Execute(FindText:=strField) Then
                    Exit Sub
                End If
            End With
                
            'Cria a Tabela
            obfWordApplication.ActiveDocument.Tables.Add Range:=obfWordApplication.Selection.Range, _
                NumRows:=TabelaArray.Count(1), NumColumns:=TabelaArray.Count(2)
                
            
            If blnBorda Then
                obfWordApplication.ActiveDocument.Tables(intTabela).AutoFormat applyborders:=True, AutoFit:=True, Format:=16
            Else
                'Nao podemos aplicar esta formatacao apos preenchido os campos, pois perde a formtacao de fonte
                obfWordApplication.ActiveDocument.Tables(intTabela).AutoFormat applyborders:=False, AutoFit:=True
            End If
            
            'Preenche a tabela
            For intColuna = 1 To TabelaArray.Count(2)
                    For intLinha = 1 To TabelaArray.Count(1)
                                            
                        obfWordApplication.ActiveDocument.Tables(intTabela).Columns(intColuna).Cells(intLinha).Select
                        
                        'Vamos verificar se ha formatacao
                        If InStr(1, TabelaArray.Value(intLinha - 1, intColuna - 1), FORMAT_NEGRITO) > 0 Then
                            obfWordApplication.Selection.Font.Bold = True
                            intCaracFormatParaExcluir = intCaracFormatParaExcluir + Len(FORMAT_NEGRITO)
                        Else
                            obfWordApplication.Selection.Font.Bold = False
                            intCaracFormatParaExcluir = 1
                        End If
                    
                        obfWordApplication.ActiveDocument.Tables(intTabela).Cell(Row:=intLinha, _
                            Column:=intColuna).Range.InsertAfter Mid(TabelaArray.Value(intLinha - 1, intColuna - 1), intCaracFormatParaExcluir)
                            
                        'Alinha as determinadas colunas com o array de parâmetro passado.
                        If Not AlinhaColunas Is Nothing Then
                            obfWordApplication.Selection.Paragraphs.Alignment = AlinhaColunas.Value(0, intColuna - 1)
                        End If
                        
                    Next intLinha
                    
            Next intColuna
            
            'obfWordApplication.ActiveDocument.Tables(intTabela).AutoFitBehavior (wdAutoFitContent)
                obfWordApplication.ActiveDocument.Tables(intTabela).AutoFormat AutoFit:=True
            
        End If
    End If

    Exit Sub
   
Problema_Na_Rotina:
   
   ExibeDetalheErro "Erro na rotina DocumentInsert."

End Sub

Public Sub DocumentSave()

    On Error GoTo Problema_Na_Rotina

    obfWordDocument.SaveAs stfDocumentPath, vrfDocumentFormat
    
    Set obfWordDocument = Nothing
    Set obfWordApplication = Nothing
    
    Exit Sub
   
Problema_Na_Rotina:
   
   ExibeDetalheErro "Erro na rotina DocumentSave."

End Sub

