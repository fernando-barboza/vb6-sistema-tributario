VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptNotaFiscalIss 
   Caption         =   "Tributario - rptNotaFiscalIss (ActiveReport)"
   ClientHeight    =   11115
   ClientLeft      =   0
   ClientTop       =   285
   ClientWidth     =   15240
   MDIChild        =   -1  'True
   _ExtentX        =   26882
   _ExtentY        =   19606
   SectionData     =   "rptNotaFiscalIss.dsx":0000
End
Attribute VB_Name = "rptNotaFiscalIss"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private ArrayNotas()        As String
Dim blnConfig               As Boolean
Dim iRow                    As Integer

Private Sub ActiveReport_DataInitialize()
    If Not blnConfig Then
        Fields.Add "RazaoSocial"
        Fields.Add "Strinscricaoestadual"
        Fields.Add "bytNaturezaJuridica"
        Fields.Add "Strcnpjcpf"
        Fields.Add "strLogradouro"
        Fields.Add "intCep"
        Fields.Add "DTMDTBASE"
        Fields.Add "INTCONTROLENR"
        Fields.Add "STRNOTAFISCALNR"
        Fields.Add "DTMDTLIMITE"
        Fields.Add "strInscricaoCadastral"
        Fields.Add "strnomefantasia"
        Fields.Add "strtelefone"
    End If
    
    
    iRow = LBound(ArrayNotas, 2)
    
End Sub

Private Sub ActiveReport_FetchData(EOF As Boolean)

    If iRow > UBound(ArrayNotas, 2) Then
        EOF = True
        Exit Sub
    End If
    
    'Vamos obter campos que nao existem no array
    Fields("RazaoSocial") = ArrayNotas(0, iRow)
    Fields("Strinscricaoestadual") = ArrayNotas(1, iRow)
    Fields("bytNaturezaJuridica") = ArrayNotas(2, iRow)
    Fields("Strcnpjcpf") = ArrayNotas(3, iRow)
    Fields("strLogradouro") = ArrayNotas(4, iRow)
    Fields("intCep") = ArrayNotas(5, iRow)
    Fields("DTMDTBASE") = ArrayNotas(6, iRow)
    Fields("INTCONTROLENR") = ArrayNotas(7, iRow)
    Fields("STRNOTAFISCALNR") = gstrConvVrDoSql(ArrayNotas(8, iRow), , , True)
    Fields("DTMDTLIMITE") = ArrayNotas(10, iRow)
    Fields("strInscricaoCadastral") = ArrayNotas(11, iRow)
    Fields("strnomefantasia") = ArrayNotas(12, iRow)
    Fields("strtelefone") = ArrayNotas(13, iRow)

    EOF = False
    iRow = iRow + 1

End Sub

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
    PadronizaToolBarRelatorio Me
    'Me.Printer.PaperHeight = 12350
End Sub

Private Sub ActiveReport_ToolbarClick(ByVal Tool As DDActiveReports2.DDTool)
    Dim vnt As Variant
    If Tool.ID = 14 Then
        ActiveReport_KeyPress 27
    ElseIf Tool.ID = 15 Then
        AbreOpcoesExportacao Me
    ElseIf Tool.ID = 16 Then
        Configura_Relatorio Me, True
        blnConfig = True
    End If
End Sub

Private Sub Detail_Format()
'    If Val(txtbytNaturezaJuridica) = 0 Then
'        lblTipo = "CGC"
'    Else
'        lblTipo = "CNPJ"
'    End If
    txtStrInscricaoCadastral = gstrFormataInscricao(Replace(txtStrInscricaoCadastral, ".", ""), TYP_ECONOMICA)
    txtintCep = gstrCEPFormatado(txtintCep)
    txtStrcnpjcpf = gstrCGCCPFFormatado(txtStrcnpjcpf)
    txtstrNotaFiscalNr = CDbl(txtstrNotaFiscalNr)
    TXTSTRNOTAFISCALNR1 = CDbl(TXTSTRNOTAFISCALNR1)
End Sub

Public Sub InicializaArray(ArrayCampos() As String)
    ArrayNotas = ArrayCampos
End Sub
