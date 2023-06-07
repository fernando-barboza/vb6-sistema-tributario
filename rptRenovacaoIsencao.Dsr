VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptRenovacaoIsencao 
   Caption         =   "Tributario - rptRenovacaoIsencao (ActiveReport)"
   ClientHeight    =   10815
   ClientLeft      =   0
   ClientTop       =   285
   ClientWidth     =   15360
   MDIChild        =   -1  'True
   _ExtentX        =   27093
   _ExtentY        =   19076
   SectionData     =   "rptRenovacaoIsencao.dsx":0000
End
Attribute VB_Name = "rptRenovacaoIsencao"
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

Private Sub ActiveReport_ReportStart()
    On Error Resume Next
    PadronizaToolBarRelatorio Me, lblExercicio
    LeImagemLogotipo imgBrasao, imgLogotipo, txtNomeFantasia, txtEstado
    lblRelatorio = Me.Caption
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

Private Sub GroupFooter1_Format()
    txtintCep2 = gstrCEPFormatado(txtintCep2)
    txtstrInscricao2 = gstrFormataInscricao(Right(txtstrInscricao2, gintRetornaTamanhoMascara(TYP_IMOBILIARIA)), TYP_IMOBILIARIA)
    txtData2 = gstrDataPorExtenso(gstrDataDoSistema, False, True)
End Sub

Private Sub GroupHeader1_Format()
Dim adoResultado As New ADODB.Recordset
Dim strSql       As String
Dim strContatos  As String
    
    If Trim(adoDataControl.Recordset("strPromissario")) <> "" Then
        txtstrNome.Text = Trim(adoDataControl.Recordset("strPromissario"))
    Else
        txtstrNome.Text = Trim(adoDataControl.Recordset("strProprietario"))
    End If
    
    strSql = "SELECT strConteudo FROM " & gstrFormaDeComunicacao & " WHERE intContribuinte = " & adoDataControl.Recordset("Pkid").Value
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        Do While Not adoResultado.EOF
            strContatos = strContatos & "  " & adoResultado("strConteudo").Value
            adoResultado.MoveNext
        Loop
    End If
    
    txtintCep = gstrCEPFormatado(txtintCep)
    txtstrInscricao = gstrFormataInscricao(Right(txtstrInscricao, gintRetornaTamanhoMascara(TYP_IMOBILIARIA)), TYP_IMOBILIARIA)
    txtData = gstrDataPorExtenso(gstrDataDoSistema, False, True)
    txtstrContato = strContatos
    
End Sub

Private Sub PageFooter_Format()
    lblPagina.Caption = "Página " & pageNumber
    MostraEmissorRelatorio Me
End Sub

Private Sub PageHeader_Format()
    lblDataHora = gstrDataDoSistema(True, , True)
End Sub

