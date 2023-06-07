VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptSubRMFReceitaExtra 
   Caption         =   "prjOrcamentario - rptSubRMFReceitaExtra (ActiveReport)"
   ClientHeight    =   10815
   ClientLeft      =   0
   ClientTop       =   285
   ClientWidth     =   15360
   MDIChild        =   -1  'True
   _ExtentX        =   27093
   _ExtentY        =   19076
   SectionData     =   "rptSubRMFReceitaExtra.dsx":0000
End
Attribute VB_Name = "rptSubRMFReceitaExtra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public dblTotalReceitaExtra As Double




Private Sub ActiveReport_ReportStart()
   dblTotalReceitaExtra = 0
End Sub

Private Sub Detail_Format()

   txtdblValorExtra = gstrConvVrDoSql(adoDataControl.Recordset("dblValor").Value)
   
   dblTotalReceitaExtra = dblTotalReceitaExtra + adoDataControl.Recordset("dblValor").Value
   
   txtdblTotalExtraReceitaOrcamentaria = gstrConvVrDoSql(dblTotalReceitaExtra)
   
   txtstrContaContabil = gvntFormatacaoEspecifica(txtstrContaContabil)
   
End Sub
