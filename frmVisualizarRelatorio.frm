VERSION 5.00
Object = "{A45D986F-3AAF-4A3B-A003-A6C53E8715A2}#1.0#0"; "ARVIEW2.OCX"
Begin VB.Form frmVisualizarRelatorio 
   Caption         =   "Form1"
   ClientHeight    =   7845
   ClientLeft      =   1965
   ClientTop       =   1935
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7845
   ScaleWidth      =   6585
   Begin DDActiveReportsViewer2Ctl.ARViewer2 ARViewer 
      Height          =   7920
      Left            =   -15
      TabIndex        =   0
      Top             =   -60
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   13970
      SectionData     =   "frmVisualizarRelatorio.frx":0000
   End
End
Attribute VB_Name = "frmVisualizarRelatorio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ARViewer_ToolbarClick(ByVal Tool As DDActiveReportsViewer2Ctl.IDDTool)
    If Tool.ID = 14 Then
        Form_KeyPress 27
    ElseIf Tool.ID = 15 Then
        AbreOpcoesExportacao Me.ARViewer.ReportSource
    ElseIf Tool.ID = 16 Then
        Configura_Relatorio Me.ARViewer.ReportSource, True
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    ARViewer.Left = 0
    ARViewer.Top = 0
    
    ARViewer.Width = Me.Width - 100
    ARViewer.Height = Me.Height - 100

    With ARViewer.Toolbar
              
        .Tools.Item(0).Visible = False
        .Tools.Item(1).Visible = False
        .Tools.Item(3).Visible = False
        .Tools.Item(4).Visible = False
        .Tools.Item(5).Visible = False
                
        '.Tools.Item(6).Visible = False
        .Tools.Item(8).Visible = False
        .Tools.Item(9).Visible = False
        .Tools.Item(10).Visible = False
                
        .Tools.Item(2).Caption = ""
        .Tools.Item(2).Tooltip = "Imprimir"
        .Tools.Item(6).Tooltip = "Localizar"
        .Tools.Item(11).Tooltip = "Diminuir"
        .Tools.Item(12).Tooltip = "Aumentar"
        .Tools.Item(15).Tooltip = "Página anterior"
        .Tools.Item(16).Tooltip = "Página seguinte"
        .Tools.Item(17).Tooltip = "Número da Página"
        .Tools.Item(19).Caption = ""
        .Tools.Item(19).Tooltip = "Histórico anterior"
        .Tools.Item(20).Caption = ""
        .Tools.Item(20).Tooltip = "Histórico seguinte"
        .Tools.Add "&Exportar..."
        .Tools.Item(21).ID = 15
        .Tools.Item(21).Tooltip = "Exportar"
        .Tools.Add "&Fechar"
        .Tools.Item(22).ID = 14
        .Tools.Item(22).Tooltip = "Fechar"
        .Tools.Insert 3, "&Configurar..."
        .Tools.Item(3).ID = 16
        .Tools.Item(3).Tooltip = "Configurar impressão"
    End With
    
    With ARViewer
        .Zoom = 77
    End With
    
End Sub

Private Sub Form_Resize()
    ARViewer.Left = 0
    ARViewer.Top = 0
    
    ARViewer.Width = Me.Width - 120
    ARViewer.Height = Me.Height - 120
End Sub
