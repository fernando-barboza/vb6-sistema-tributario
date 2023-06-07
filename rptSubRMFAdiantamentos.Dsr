VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptSubRMFAdiantamentos 
   Caption         =   "prjOrcamentario - rptSubRMFAdiantamentos (ActiveReport)"
   ClientHeight    =   10815
   ClientLeft      =   0
   ClientTop       =   285
   ClientWidth     =   15360
   MDIChild        =   -1  'True
   _ExtentX        =   27093
   _ExtentY        =   19076
   SectionData     =   "rptSubRMFAdiantamentos.dsx":0000
End
Attribute VB_Name = "rptSubRMFAdiantamentos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dblTotalDeAdiantamentos As Double

Private Sub ActiveReport_ReportStart()

   dblTotalDeAdiantamentos = 0

End Sub

Private Sub Detail_Format()

   txtstrCodigoElementoDespesa = gvntFormatacaoEspecifica(txtstrCodigoElementoDespesa)
   
   txtdblValor = gstrConvVrDoSql(txtdblValor)
   
   dblTotalDeAdiantamentos = dblTotalDeAdiantamentos + CDbl(txtdblValor)
   
   txtdblTotalDeAdiantamentos = gstrConvVrDoSql(dblTotalDeAdiantamentos)

End Sub
