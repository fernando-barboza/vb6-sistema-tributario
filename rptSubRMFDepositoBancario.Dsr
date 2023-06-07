VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptSubRMFDepositoBancario 
   Caption         =   "prjOrcamentario - rptSubRMFDepositoBancario (ActiveReport)"
   ClientHeight    =   10815
   ClientLeft      =   0
   ClientTop       =   285
   ClientWidth     =   15360
   MDIChild        =   -1  'True
   _ExtentX        =   27093
   _ExtentY        =   19076
   SectionData     =   "rptSubRMFDepositoBancario.dsx":0000
End
Attribute VB_Name = "rptSubRMFDepositoBancario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dblTotalDepositado As Double

Private Sub ActiveReport_ReportStart()
   dblTotalDepositado = 0
End Sub

Private Sub Detail_Format()

txtdblCredito = gstrConvVrDoSql(txtdblCredito)

dblTotalDepositado = dblTotalDepositado + CDbl(txtdblCredito)

txtdblValorDepositado = gstrConvVrDoSql(dblTotalDepositado)

End Sub
