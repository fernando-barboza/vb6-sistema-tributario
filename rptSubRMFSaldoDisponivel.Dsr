VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptSubRMFSaldoDisponivel 
   Caption         =   "prjOrcamentario - rptSubRMFSaldoDisponivel (ActiveReport)"
   ClientHeight    =   10815
   ClientLeft      =   0
   ClientTop       =   285
   ClientWidth     =   15360
   MDIChild        =   -1  'True
   _ExtentX        =   27093
   _ExtentY        =   19076
   SectionData     =   "rptSubRMFSaldoDisponivel.dsx":0000
End
Attribute VB_Name = "rptSubRMFSaldoDisponivel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim dblTotal As Double

Private Sub ActiveReport_ReportStart()
   dblTotal = 0
End Sub

Private Sub Detail_Format()
Dim dblSaldo As Double

dblSaldo = adoDataControl.Recordset("SaldoInicial").Value + ((adoDataControl.Recordset("CreditoAnterior").Value + adoDataControl.Recordset("Credito").Value) - (adoDataControl.Recordset("DebitoAnterior").Value + adoDataControl.Recordset("Debito").Value))

dblTotal = dblTotal + dblSaldo

txtdblSaldo = gstrConvVrDoSql(dblSaldo)

txtdblTotalGeralSaldosDisponiveis = gstrConvVrDoSql(dblTotal)


End Sub
