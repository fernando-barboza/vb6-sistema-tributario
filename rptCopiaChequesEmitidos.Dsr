VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptCopiaChequesEmitidos 
   Caption         =   "rptCopiaChequesEmitidos (ActiveReport)"
   ClientHeight    =   10815
   ClientLeft      =   0
   ClientTop       =   285
   ClientWidth     =   15360
   ControlBox      =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   WindowState     =   2  'Maximized
   _ExtentX        =   27093
   _ExtentY        =   19076
   SectionData     =   "rptCopiaChequesEmitidos.dsx":0000
End
Attribute VB_Name = "rptCopiaChequesEmitidos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public bytTipoCopia As Byte

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

Private Sub ActiveReport_ReportStart()
    On Error Resume Next
    PadronizaToolBarRelatorio Me, lblExercicio
    LeImagemLogotipo imgBrasao, imgLogotipo, txtNomeFantasia, txtEstado
End Sub

Private Sub ActiveReport_ToolbarClick(ByVal Tool As DDActiveReports2.DDTool)
    Dim vnt As Variant
    If Tool.ID = 14 Then
        ActiveReport_KeyPress 27
    ElseIf Tool.ID = 15 Then
        AbreOpcoesExportacao Me
    ElseIf Tool.ID = 16 Then
        Configura_Relatorio Me, True
    End If
End Sub

Private Sub GroupFooter3_Format()
txtdblValor = gstrConvVrDoSql(txtdblValor)
   txtstrValor = "Pague pelo presente a quantia de: "
   txtstrValor = txtstrValor & gstrExtenso(IIf(txtdblValor = "", 0, txtdblValor))
   
   If adoDataControl.NRecords > 0 Then
      adoDataControl.Recordset.MovePrevious
    
      If Not adoDataControl.Recordset.EOF Then
         
         
         
           txtstrData = gstrCidadeEmpresa & ", " & gstrDataPorExtenso(adoDataControl.Recordset("dtmData"))
           txtstrBancoConta = Trim(adoDataControl.Recordset("strBanco")) & " - " & Trim(adoDataControl.Recordset("strConta"))
    
         If bytTipoCopia = 1 And adoDataControl.Recordset("strHistorico") = "*CHE*" Then
       
            txtstrHistorico.Text = "Referente à(s) OP('s) : " & strMostraOps(adoDataControl.Recordset("strCheque"), adoDataControl.Recordset("intContaContabil"))
       
         End If
       
       
       
      End If
    
      adoDataControl.Recordset.MoveNext
   End If
   
End Sub

Private Sub PageFooter_Format()
    lblPagina.Caption = "Página " & pageNumber
End Sub

Private Sub PageHeader_Format()
    lblDataHora = gstrDataDoSistema(True, , True)
End Sub

Private Sub ReportFooter_Format()
    MostraEmissorRelatorio Me
End Sub
Private Function strMostraOps(Cheque As String, Conta As String) As String
Dim strSql As String
Dim intCheque As String
Dim strTemp As String
Dim i As Integer
Dim adoResultado As New ADODB.Recordset
Dim adoResultado2 As New ADODB.Recordset

strMostraOps = " "
intCheque = 0

strSql = "SELECT CQ.PKID FROM tblCheque CQ, tblPlanoConta PC WHERE CQ.strCheque=" & Cheque & " AND CQ.intContaBancaria=PC.intContaBancaria AND PC.PKID=" & Conta

If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
   If Not adoResultado.EOF Then
      intCheque = adoResultado!Pkid
   End If
   adoResultado.Close
End If

If intCheque <> 0 Then
   strSql = "SELECT intOrdemPagamento FROM tblChequeOP WHERE intCheque=" & intCheque
   If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
      If Not adoResultado.EOF Then
         Do While Not adoResultado.EOF
            strSql = "SELECT intNumero FROM tblOrdemPagamento WHERE PKID=" & adoResultado!intOrdemPagamento
            If gobjBanco.CriaADO(strSql, 5, adoResultado2) Then
               If Not adoResultado2.EOF Then
                  strTemp = strTemp & adoResultado2!INTNUMERO & ", "
               End If
               adoResultado2.Close
            End If
            adoResultado.MoveNext
         Loop
      End If
      adoResultado.Close
   End If
End If

If Not strTemp = "" Then
   If Right(strTemp, 2) = ", " Then
      strTemp = Left(strTemp, Len(strTemp) - 2)
   End If
   If InStr(1, strTemp, ",") <> 0 Then
      For i = Len(strTemp) To 1 Step -1
         If Mid(strTemp, i, 1) = "," Then
            strTemp = Left(strTemp, i - 1) & " e" & Mid(strTemp, i + 1)
         End If
      Next
   End If
End If
      
If strTemp <> "" Then strMostraOps = strTemp
      
End Function
