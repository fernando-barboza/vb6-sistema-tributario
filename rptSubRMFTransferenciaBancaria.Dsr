VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptSubRMFTransferenciaBancaria 
   Caption         =   "prjOrcamentario - rptSubRMFTransferenciaBancaria (ActiveReport)"
   ClientHeight    =   10815
   ClientLeft      =   0
   ClientTop       =   285
   ClientWidth     =   15360
   MDIChild        =   -1  'True
   _ExtentX        =   27093
   _ExtentY        =   19076
   SectionData     =   "rptSubRMFTransferenciaBancaria.dsx":0000
End
Attribute VB_Name = "rptSubRMFTransferenciaBancaria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Detail_Format()


   txtstrBancoConta = adoDataControl.Recordset("strBanco").Value & " - " & adoDataControl.Recordset("strDescricaoConta").Value
   
   If adoDataControl.Recordset("Credito") > 0 Then
      txtstrValorTransferencia = gstrConvVrDoSql(adoDataControl.Recordset("Credito").Value)
      lblSinal.Caption = "( + )"
   Else
      txtstrValorTransferencia = gstrConvVrDoSql(adoDataControl.Recordset("Debito").Value)
      lblSinal.Caption = "( - )"
   End If

End Sub
