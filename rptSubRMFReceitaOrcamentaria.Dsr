VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptSubRMFReceitaOrcamentaria 
   Caption         =   "prjOrcamentario - rptSubRMFReceitaOrcamentaria (ActiveReport)"
   ClientHeight    =   10815
   ClientLeft      =   0
   ClientTop       =   285
   ClientWidth     =   15360
   MDIChild        =   -1  'True
   _ExtentX        =   27093
   _ExtentY        =   19076
   SectionData     =   "rptSubRMFReceitaOrcamentaria.dsx":0000
End
Attribute VB_Name = "rptSubRMFReceitaOrcamentaria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public dblTotalReceitaOrcamentaria As Double



Private Sub ActiveReport_ReportStart()
   dblTotalReceitaOrcamentaria = 0
End Sub

Private Sub Detail_Format()
   Dim strSQL                 As String
   Dim adoResultado           As New ADODB.Recordset
   Dim intInd                 As Integer
      
      Detail.Visible = False
      
      'If adoDataControl.Recordset("dblValorMes").Value > 0 Then
         
         If adoDataControl.Recordset("bytNivel").Value < 3 Or adoDataControl.Recordset("bytNivel").Value = 8 Then
            Detail.Visible = True
         Else
            Detail.Visible = False
         End If
         If adoDataControl.Recordset("bytNivel").Value > 2 Then
            txtDescricao.Visible = False
            txtstrCodigoOrcamentario.Visible = True
            txtstrDescricao.Visible = True
            dblValor.Visible = True
         Else
            txtDescricao.Visible = True
            txtstrCodigoOrcamentario.Visible = False
            txtstrDescricao.Visible = False
            dblValor.Visible = False
         End If
         dblValor.Text = gstrConvVrDoSql(adoDataControl.Recordset("dblValorMes").Value)
         
         txtstrCodigoOrcamentario = gvntFormatacaoEspecifica(txtstrCodigoOrcamentario)
          
         If adoDataControl.Recordset("bytNivel").Value = 1 Then
            txtDescricao.Left = 215
            dblTotalReceitaOrcamentaria = dblTotalReceitaOrcamentaria + adoDataControl.Recordset("dblValorMes").Value
         Else
            txtDescricao.Left = 300
         End If
         
         txtdblTotalReceitaOrcamentaria.Text = gstrConvVrDoSql(dblTotalReceitaOrcamentaria)
       
       'End If
       
End Sub
