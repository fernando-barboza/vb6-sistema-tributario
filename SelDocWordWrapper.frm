VERSION 5.00
Begin VB.Form frmSelDocWordWrapper 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Visualizar documento impresso ..."
   ClientHeight    =   5325
   ClientLeft      =   3105
   ClientTop       =   3570
   ClientWidth     =   8790
   Icon            =   "SelDocWordWrapper.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5325
   ScaleWidth      =   8790
   Begin VB.TextBox txtFilterNome 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   30
      TabIndex        =   9
      Top             =   1200
      Width           =   5025
   End
   Begin VB.TextBox txtFilterModificado 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   6930
      TabIndex        =   8
      Top             =   1200
      Width           =   1815
   End
   Begin VB.TextBox txtFilterCriado 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   5070
      TabIndex        =   7
      Top             =   1200
      Width           =   1845
   End
   Begin VB.ComboBox cboModelos 
      Height          =   315
      Left            =   30
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   300
      Width           =   8715
   End
   Begin VB.CommandButton cmdFechar 
      Caption         =   "&Fechar"
      Height          =   375
      Left            =   7650
      TabIndex        =   5
      Top             =   4910
      Width           =   1125
   End
   Begin VB.CommandButton cmdVisualizar 
      Caption         =   "&Visualizar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6480
      TabIndex        =   4
      Top             =   4910
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.ListBox lstDocs 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3270
      IntegralHeight  =   0   'False
      Left            =   30
      Sorted          =   -1  'True
      TabIndex        =   3
      Top             =   1575
      Width           =   8715
   End
   Begin VB.ListBox lstDocsAux 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3270
      IntegralHeight  =   0   'False
      Left            =   30
      Sorted          =   -1  'True
      Style           =   1  'Checkbox
      TabIndex        =   6
      Top             =   1575
      Visible         =   0   'False
      Width           =   6525
   End
   Begin VB.Label lblFilterModificado 
      Caption         =   "Modificado em :"
      Height          =   195
      Left            =   6930
      TabIndex        =   12
      Top             =   960
      Width           =   1830
   End
   Begin VB.Label lblFilterCriado 
      Caption         =   "Criado em :"
      Height          =   195
      Left            =   5070
      TabIndex        =   11
      Top             =   960
      Width           =   1830
   End
   Begin VB.Label lblFilterNome 
      Caption         =   "Nome"
      Height          =   195
      Left            =   30
      TabIndex        =   10
      Top             =   960
      Width           =   5010
   End
   Begin VB.Label lblModelos 
      Caption         =   "Modelos disponíveis ..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   30
      TabIndex        =   0
      Top             =   60
      Width           =   8715
   End
   Begin VB.Label lblDocs 
      Caption         =   "Documentos disponíveis ..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   30
      TabIndex        =   2
      Top             =   720
      Width           =   8700
   End
End
Attribute VB_Name = "frmSelDocWordWrapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private stfDocumentPath     As String
Private stfDocumentTemplate As String

Public Property Get DocumentoSelecionado() As String

On Error GoTo Problema_Na_Rotina

   DocumentoSelecionado = stfDocumentPath
   
   Exit Property

Problema_Na_Rotina:

'  If RecoverError("DocumentoSelecionado") Then Resume
   
   ExibeDetalheErro "Erro na rotina DocumentoSelecionado."
   
End Property

Private Sub CriaRepositorioModelos()
Dim stpFolder     As String
Dim objFolder     As Scripting.Folder
Dim objFileSystem As Scripting.FileSystemObject

On Error GoTo Problema_Na_Rotina

    stpFolder = gstrDirDocumentos & "\Documentos\" & App.ProductName & "\WordModelos"

    Set objFileSystem = New Scripting.FileSystemObject
    
    If Not objFileSystem.FolderExists(gstrDirDocumentos & "\Documentos\" & App.ProductName) Then MkDir gstrDirDocumentos & "\Documentos\" & App.ProductName

    If Not objFileSystem.FolderExists(stpFolder) Then Set objFolder = objFileSystem.CreateFolder(stpFolder)
      
    Set objFolder = Nothing

    Set objFileSystem = Nothing
   
    Exit Sub

Problema_Na_Rotina:

'  If RecoverError("CriaRepositorioModelos") Then Resume
   
   ExibeDetalheErro "Erro na rotina CriaRepositorioModelos."

End Sub

Private Sub CriaRepositorioDocumentos()
Dim stpFolder     As String
Dim objFolder     As Scripting.Folder
Dim objFileSystem As Scripting.FileSystemObject

On Error GoTo Problema_Na_Rotina

   stpFolder = gstrDirDocumentos & "\Documentos\" & App.ProductName & "\WordGravados"

   Set objFileSystem = New Scripting.FileSystemObject

   If Not objFileSystem.FolderExists(gstrDirDocumentos & "\Documentos\" & App.ProductName) Then MkDir gstrDirDocumentos & "\Documentos\" & App.ProductName

   If Not objFileSystem.FolderExists(stpFolder) Then Set objFolder = objFileSystem.CreateFolder(stpFolder)
      
   Set objFolder = Nothing

   Set objFileSystem = Nothing
   
   Exit Sub

Problema_Na_Rotina:

'  If RecoverError("CriaRepositorioDocumentos") Then Resume
   
   ExibeDetalheErro "Erro na rotina CriaRepositorioDocumentos."

End Sub

Private Sub ExibeModelosGravados()
Dim objFile       As Scripting.file
Dim objFiles      As Scripting.Files
Dim stpFolder     As String
Dim objFolder     As Scripting.Folder
Dim objFileSystem As Scripting.FileSystemObject

On Error GoTo Problema_Na_Rotina

   cboModelos.Clear
         
   CriaRepositorioModelos
   
   stpFolder = gstrDirDocumentos & "\Documentos\" & App.ProductName & "\WordModelos"
      
   Set objFileSystem = New Scripting.FileSystemObject
    
   Set objFolder = objFileSystem.GetFolder(stpFolder)
       
   Set objFiles = objFolder.Files
   
   For Each objFile In objFiles
      If UCase$(Right(objFile.Name, 3)) = "DOT" Then cboModelos.AddItem Left(objFile.Name, Len(objFile.Name) - 4)
   Next
    
   Exit Sub

Problema_Na_Rotina:

'  If RecoverError("ExibeModelosGravados") Then Resume
   
   ExibeDetalheErro "Erro na rotina ExibeModelosGravados."

End Sub

Private Sub ExibeDocumentosGravados()
Dim objFile             As Scripting.file
Dim objFiles            As Scripting.Files
Dim stpFolder           As String
Dim objFolder           As Scripting.Folder
Dim stpFileName         As String
Dim stpFileDateC        As String
Dim stpFileDateM        As String
Dim objFileSystem       As Scripting.FileSystemObject
Dim blpAdicionaNaLista  As Boolean
Dim blpAdicionaNaLista1 As Boolean
Dim blpAdicionaNaLista2 As Boolean
Dim blpAdicionaNaLista3 As Boolean

On Error GoTo Problema_Na_Rotina

   lstDocs.Clear: lstDocsAux.Clear
         
   CriaRepositorioDocumentos
   
   stpFolder = gstrDirDocumentos & "\Documentos\" & App.ProductName & "\WordGravados"
      
   Set objFileSystem = New Scripting.FileSystemObject
    
   Set objFolder = objFileSystem.GetFolder(stpFolder)
       
   Set objFiles = objFolder.Files
   
   For Each objFile In objFiles
           
      If InStr(1, objFile.Name, stfDocumentTemplate) > 0 And UCase$(Right(objFile.Name, 3)) = "DOC" Then
         
         blpAdicionaNaLista = True
         
         stpFileName = Mid$(objFile.Name, InStr(1, objFile.Name, "_") + 1)
         stpFileName = Left(stpFileName, Len(stpFileName) - 4)
         
         stpFileDateC = Format$(objFile.DateCreated, "dd/mm/yyyy hh:mm")
         stpFileDateM = Format$(objFile.DateLastModified, "dd/mm/yyyy hh:mm")
         
         blpAdicionaNaLista = False
         blpAdicionaNaLista1 = True
         blpAdicionaNaLista2 = True
         blpAdicionaNaLista3 = True
         
         If txtFilterNome.Text <> Space$(0) Then
            If InStr(1, stpFileName, txtFilterNome.Text) = 0 Then blpAdicionaNaLista1 = False
         End If
        
         If txtFilterCriado.Text <> Space$(0) Then
            If InStr(1, stpFileDateC, txtFilterCriado.Text) = 0 Then blpAdicionaNaLista2 = False
         End If
        
         If txtFilterModificado.Text <> Space$(0) Then
            If InStr(1, stpFileDateM, txtFilterModificado.Text) = 0 Then blpAdicionaNaLista3 = False
         End If
            
         blpAdicionaNaLista = (blpAdicionaNaLista1 And blpAdicionaNaLista2 And blpAdicionaNaLista3)
         
         If blpAdicionaNaLista Then
           'lstDocs.AddItem Left$(stpFileName & Space(25), 25) & stpFileDateC & Space$(2) & stpFileDateM
            lstDocs.AddItem Left$(stpFileName & Space(48), 48) & stpFileDateC & Space$(2) & stpFileDateM
            lstDocsAux.AddItem objFile.Path
        End If
        
      End If
   Next
    
   Exit Sub

Problema_Na_Rotina:

'  If RecoverError("ExibeDocumentosGravados") Then Resume
   
   ExibeDetalheErro "Erro na rotina ExibeDocumentosGravados."

End Sub

Private Sub Form_Load()
   
On Error GoTo Problema_Na_Rotina

   TrocaInconiDoObj Me, 3
   
   stfDocumentTemplate = Space$(0)
   stfDocumentPath = Space$(0)
   
   Exit Sub

Problema_Na_Rotina:

'  If RecoverError("Form_Load") Then Resume
   
   ExibeDetalheErro "Erro na rotina Form_Load."
   
End Sub

Private Sub cboModelos_DropDown()

On Error GoTo Problema_Na_Rotina

   ExibeModelosGravados

   Exit Sub

Problema_Na_Rotina:

'  If RecoverError("cboModelos_DropDown") Then Resume
   
   ExibeDetalheErro "Erro na rotina cboModelos_DropDown."

End Sub

Private Sub cboModelos_Click()

On Error GoTo Problema_Na_Rotina

   stfDocumentTemplate = cboModelos.Text

   ExibeDocumentosGravados

   Exit Sub

Problema_Na_Rotina:

'  If RecoverError("cboModelos_Click") Then Resume
   
   ExibeDetalheErro "Erro na rotina cboModelos_Click."

End Sub

Private Sub lstDocs_ItemCheck(Item As Integer)
Dim itpFor             As Integer
Dim objFile            As Scripting.file
Dim objFiles           As Scripting.Files
Dim stpFolder          As String
Dim objFolder          As Scripting.Folder
Dim objFileSystem      As Scripting.FileSystemObject

Static blpDisableEvent As Boolean

On Error GoTo Problema_Na_Rotina

   stfDocumentPath = Space$(0)

   If Not blpDisableEvent Then

      If lstDocs.Selected(Item) Then
      
         Set objFileSystem = New Scripting.FileSystemObject
         
         cmdVisualizar.Enabled = objFileSystem.FileExists(lstDocsAux.List(Item))
         
         Set objFileSystem = Nothing
         
         blpDisableEvent = True: stfDocumentPath = lstDocsAux.List(Item)
         
         For itpFor = 0 To lstDocs.ListCount - 1
            If itpFor <> Item Then lstDocs.Selected(itpFor) = False
         Next itpFor
         
         blpDisableEvent = False
   
      End If
      
   End If
   
   Exit Sub

Problema_Na_Rotina:

'  If RecoverError("lstDocs_ItemCheck") Then Resume
   
   ExibeDetalheErro "Erro na rotina lstDocs_ItemCheck."

End Sub

Private Sub lstDocs_DblClick()
Dim objFileSystem As Scripting.FileSystemObject

On Error GoTo Problema_Na_Rotina

    If lstDocs.ListIndex > -1 Then
    
        Set objFileSystem = New Scripting.FileSystemObject
             
        If objFileSystem.FileExists(lstDocsAux.List(lstDocs.ListIndex)) Then
            stfDocumentPath = lstDocsAux.List(lstDocs.ListIndex)
            Me.Hide
        Else
            MsgBox "O documento selecionado não foi localizado. A operação não pode ser realizada.", vbOKOnly + vbInformation, "Mensagem ao usuário"
        End If
        
        Set objFileSystem = Nothing
   
    End If
    
   Exit Sub

Problema_Na_Rotina:

'  If RecoverError("cmdVisualizar_Click") Then Resume
   
   ExibeDetalheErro "Erro na rotina lstDocs_DblClick."
   
End Sub

Private Sub cmdVisualizar_Click()

On Error GoTo Problema_Na_Rotina

   Me.Hide
   
   Exit Sub

Problema_Na_Rotina:

'  If RecoverError("cmdVisualizar_Click") Then Resume
   
   ExibeDetalheErro "Erro na rotina cmdVisualizar_Click."

End Sub

Private Sub cmdFechar_Click()

On Error GoTo Problema_Na_Rotina

   stfDocumentPath = Space$(0)

   Me.Hide

   Exit Sub

Problema_Na_Rotina:

'  If RecoverError("cmdFechar_Click") Then Resume
   
   ExibeDetalheErro "Erro na rotina cmdFechar_Click."

End Sub

Private Sub txtFilterNome_Change()

On Error GoTo Problema_Na_Rotina

    txtFilterCriado = Space$(0)
    
    txtFilterModificado = Space$(0)
    
    ExibeDocumentosGravados
    
   Exit Sub

Problema_Na_Rotina:

'  If RecoverError("txtFilterNome_Change") Then Resume
   
   ExibeDetalheErro "Erro na rotina txtFilterNome_Change."
    
End Sub

Private Sub txtFilterCriado_Change()

On Error GoTo Problema_Na_Rotina

    txtFilterNome = Space$(0)
    
    txtFilterModificado = Space$(0)
    
    ExibeDocumentosGravados
    
   Exit Sub

Problema_Na_Rotina:

'  If RecoverError("txtFilterCriado_Change") Then Resume
   
   ExibeDetalheErro "Erro na rotina txtFilterCriado_Change."
    
End Sub

Private Sub txtFilterModificado_Change()

On Error GoTo Problema_Na_Rotina

    txtFilterNome = Space$(0)
    
    txtFilterCriado = Space$(0)
    
    ExibeDocumentosGravados

   Exit Sub

Problema_Na_Rotina:

'  If RecoverError("txtFilterModificado_Change") Then Resume
   
   ExibeDetalheErro "Erro na rotina txtFilterModificado_Change."
    
End Sub
