VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmOpcoesExportacao 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Exportação"
   ClientHeight    =   3300
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6060
   Icon            =   "OpcoesExportacao.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3300
   ScaleWidth      =   6060
   Begin VB.CommandButton cmd_Cancel 
      Caption         =   "&Fechar"
      Height          =   375
      Left            =   4650
      TabIndex        =   10
      Top             =   2850
      Width           =   1335
   End
   Begin VB.CommandButton cmd_OK 
      Caption         =   "&Exportar"
      Height          =   375
      Left            =   3150
      TabIndex        =   9
      Top             =   2850
      Width           =   1335
   End
   Begin TabDlg.SSTab tab_3dOpcoes 
      Height          =   2655
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   4683
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Opções de exportação"
      TabPicture(0)   =   "OpcoesExportacao.frx":1042
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblTag(1)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblTag(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblTag(2)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "chkSuppressBlank"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtDelimitadorCampo"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "chkUnicaode"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtDelimitadorPagina"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cboTipodeArquivo"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      Begin VB.ComboBox cboTipodeArquivo 
         Height          =   315
         ItemData        =   "OpcoesExportacao.frx":105E
         Left            =   1830
         List            =   "OpcoesExportacao.frx":1060
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   600
         Width           =   3345
      End
      Begin VB.TextBox txtDelimitadorPagina 
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         Height          =   305
         Left            =   1830
         TabIndex        =   4
         Top             =   1500
         Width           =   3735
      End
      Begin VB.CheckBox chkUnicaode 
         Caption         =   "Unicode"
         Enabled         =   0   'False
         Height          =   255
         Left            =   1830
         TabIndex        =   3
         Top             =   2250
         Width           =   3735
      End
      Begin VB.TextBox txtDelimitadorCampo 
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         Height          =   305
         Left            =   1830
         TabIndex        =   2
         Top             =   1050
         Width           =   3735
      End
      Begin VB.CheckBox chkSuppressBlank 
         Caption         =   "Ocultar linhas em branco"
         Enabled         =   0   'False
         Height          =   255
         Left            =   1830
         TabIndex        =   1
         Top             =   1890
         Width           =   3735
      End
      Begin VB.Label lblTag 
         AutoSize        =   -1  'True
         Caption         =   "Salvar como tipo:"
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   7
         Top             =   720
         Width           =   1230
      End
      Begin VB.Label lblTag 
         AutoSize        =   -1  'True
         Caption         =   "Delimitador de página:"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   6
         Top             =   1605
         Width           =   1575
      End
      Begin VB.Label lblTag 
         AutoSize        =   -1  'True
         Caption         =   "Delimitador de campo:"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   5
         Top             =   1155
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frmOpcoesExportacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboTipodeArquivo_Click()
    txtDelimitadorCampo = ""
    txtDelimitadorPagina = ""
    chkSuppressBlank.Value = 0
    chkUnicaode.Value = 0
    
    If cboTipodeArquivo.ListIndex = 1 Then
        txtDelimitadorCampo.Enabled = True
        txtDelimitadorPagina.Enabled = True
        chkSuppressBlank.Enabled = True
        chkUnicaode.Enabled = True
        txtDelimitadorCampo.BackColor = &H80000005
        txtDelimitadorPagina.BackColor = &H80000005
    Else
        txtDelimitadorCampo.Enabled = False
        txtDelimitadorPagina.Enabled = False
        chkSuppressBlank.Enabled = False
        chkUnicaode.Enabled = False
        txtDelimitadorCampo.BackColor = &H80000004
        txtDelimitadorPagina.BackColor = &H80000004
    End If
End Sub

Private Sub cmd_Cancel_Click()
    Unload Me
End Sub

Private Sub cmd_OK_Click()
    Select Case cboTipodeArquivo.ListIndex
        Case 0
            Call ExportRTF
        Case 1
            Call ExportTXT
        Case 2
            Call ExportPDF
        Case 3
            Call ExportXLS
    End Select
End Sub

Private Sub Form_Load()
    cboTipodeArquivo.AddItem "Rich Text Format (*.rtf)"
    cboTipodeArquivo.AddItem "Somente texto (*.txt)"
    cboTipodeArquivo.AddItem "(Portable Document Format (*.pdf)"
    cboTipodeArquivo.AddItem "Planilha do Excel (*.xls)"
    cboTipodeArquivo.ListIndex = 0
End Sub

Sub ExportPDF()
    Dim pdf   As New ActiveReportsPDFExport.ARExportPDF
    Dim sFile As String
    Dim bSave As Boolean
    
    On Error GoTo Err_Handle
    bSave = VBGetSaveFileName(sFile, "", True, _
        "Portable Document Format (*.PDF)| *.PDF", , , _
        "Export to PDF", "*.PDF", Me.hWnd, cdlOFNExplorer Or cdlOFNHideReadOnly Or cdlOFNLongNames)
    
    If bSave Then pdf.Filename = sFile Else Exit Sub
    
    If gobjRelatorio.Pages.Count > 0 Then
        pdf.Export gobjRelatorio.Pages
    End If
    
    Set pdf = Nothing
    ExibeMensagem "Exportação completada com sucesso."
    
Exit Sub
Err_Handle:
    ExibeDetalheErro ""
End Sub

Sub ExportRTF()
    Dim rtf   As New ActiveReportsRTFExport.ARExportRTF
    Dim sFile As String
    Dim bSave As Boolean

    On Error GoTo Err_Handle
    
    bSave = VBGetSaveFileName(sFile, "", True, _
        "Rich Text Format (*.RTF)| *.RTF", , , _
        "Export to RTF", "*.RTF", Me.hWnd, cdlOFNExplorer Or cdlOFNHideReadOnly Or cdlOFNLongNames)
    
    If bSave Then rtf.Filename = sFile Else Exit Sub
    
    If gobjRelatorio.Pages.Count > 0 Then
        rtf.Export gobjRelatorio.Pages
    End If
    
    Set rtf = Nothing
    ExibeMensagem "Exportação completada com sucesso."
    
Exit Sub
Err_Handle:
    ExibeDetalheErro ""
End Sub

Sub ExportTXT()
    Dim txt   As New ActiveReportsTextExport.ARExportText
    Dim sFile As String
    Dim bSave As Boolean

    On Error GoTo Err_Handle
    
    bSave = VBGetSaveFileName(sFile, "", True, _
        "Text Files (*.txt)| *.txt;All Files (*.*)| *.*", , , _
        "Export to Text", "*.txt", Me.hWnd, cdlOFNExplorer Or cdlOFNHideReadOnly Or cdlOFNLongNames)
    
    If bSave Then txt.Filename = sFile Else Exit Sub

    txt.PageDelimiter = Trim(txtDelimitadorPagina)
    txt.TextDelimiter = Trim(txtDelimitadorCampo)
    txt.Unicode = chkUnicaode.Value
    txt.SuppressEmptyLines = chkSuppressBlank.Value
    
    If gobjRelatorio.Pages.Count > 0 Then
        txt.Export gobjRelatorio.Pages
    End If
    
    Set txt = Nothing
    ExibeMensagem "Exportação completada com sucesso."
    
Exit Sub
Err_Handle:
    ExibeDetalheErro ""
End Sub

Sub ExportXLS()
    Dim xls   As New ActiveReportsExcelExport.ARExportExcel
    Dim sFile As String
    Dim bSave As Boolean

    On Error GoTo Err_Handle
    
    bSave = VBGetSaveFileName(sFile, "", True, _
        "Excel Format (*.xls)| *.xls", , , _
        "Export to Excel", "*.xls", Me.hWnd, cdlOFNExplorer Or cdlOFNHideReadOnly Or cdlOFNLongNames)
    
    If bSave Then xls.Filename = sFile Else Exit Sub
    
    If gobjRelatorio.Pages.Count > 0 Then
        xls.Export gobjRelatorio.Pages
    End If
    
    Set xls = Nothing
    ExibeMensagem "Exportação completada com sucesso."
    
Exit Sub
Err_Handle:
    ExibeDetalheErro ""
End Sub
