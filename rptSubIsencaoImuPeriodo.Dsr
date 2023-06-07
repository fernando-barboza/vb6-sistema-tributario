VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptSubIsencaoImuPeriodo 
   Caption         =   "Tributario - rptSubIsencaoImuPeriodo (ActiveReport)"
   ClientHeight    =   10815
   ClientLeft      =   2115
   ClientTop       =   3555
   ClientWidth     =   15240
   MDIChild        =   -1  'True
   _ExtentX        =   26882
   _ExtentY        =   19076
   SectionData     =   "rptSubIsencaoImuPeriodo.dsx":0000
End
Attribute VB_Name = "rptSubIsencaoImuPeriodo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Detail_BeforePrint()
If Trim(txtstrProcesso.Text) = "/-" Then
    txtstrProcesso.Text = ""
End If
End Sub

Private Sub Detail_Format()
  Select Case Val(gstrENulo(txtbytPosicao.Text))
    Case 0
      txtbytPosicao.Text = "Deferido"
    Case 1
      txtbytPosicao.Text = "Indeferido"
    Case 2
      txtbytPosicao.Text = "Em Andamento"
  End Select
  
  Select Case Val(gstrENulo(txtbytCancelamento.Text))
    Case 0
      txtbytCancelamento.Text = "Não cancelado"
    Case 1
      txtbytCancelamento.Text = "Cancelado"
  End Select
  
End Sub

