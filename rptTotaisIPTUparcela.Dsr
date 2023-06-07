VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptTotaisTPTUparcela 
   Caption         =   "Tributario - rptTotaisTPTUparcela (ActiveReport)"
   ClientHeight    =   5565
   ClientLeft      =   1110
   ClientTop       =   615
   ClientWidth     =   12660
   MDIChild        =   -1  'True
   _ExtentX        =   22331
   _ExtentY        =   9816
   SectionData     =   "rptTotaisIPTUparcela.dsx":0000
End
Attribute VB_Name = "rptTotaisTPTUparcela"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit

Private Sub ActiveReport_Activate()
    HabilitaDesabilitaBotao1 False, gstrBtnArquivo, True
End Sub

Private Sub ActiveReport_Deactivate()
    HabilitaDesabilitaBotao1 True, gstrBtnArquivo, True
End Sub

Private Sub ActiveReport_QueryClose(Cancel As Integer, CloseMode As Integer)
    HabilitaDesabilitaBotao1 True, gstrBtnArquivo, True
End Sub

Private Sub ActiveReport_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        Unload Me
    End If
End Sub
Private Sub Detail_Format()
    txt_Vencimento = gstrDataFormatada(txt_Vencimento)
End Sub
