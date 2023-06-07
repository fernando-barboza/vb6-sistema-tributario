VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptReservaDotacao 
   Caption         =   "prjOrcamentario - rptReservaDotacao (ActiveReport)"
   ClientHeight    =   10815
   ClientLeft      =   0
   ClientTop       =   285
   ClientWidth     =   15360
   MDIChild        =   -1  'True
   WindowState     =   2  'Maximized
   _ExtentX        =   27093
   _ExtentY        =   19076
   SectionData     =   "ReservaDotacao.dsx":0000
End
Attribute VB_Name = "rptReservaDotacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dblTotalReservado As Double
Dim dblTotalCancelado As Double
Dim strProgramaDeTrabalho As String

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

Private Sub Detail_Format()
'    txtValorDoCancelamento.Visible = True
'    txtDataDoCancelamento.Visible = True
'    txtNumeroDoCancelamento.Visible = True
'    lblSombra.Visible = True
    Detail.Visible = True
    If txtNumeroDoCancelamento.Text <> "" Then
        txtNumeroDoCancelamento.Text = "00000" & Trim(txtNumeroDoCancelamento.Text)
    End If
    txtValorDoCancelamento.Text = gstrConvVrDoSql(txtValorDoCancelamento.Text)
    If txtValorDoCancelamento.Text <> "" Then
        dblTotalCancelado = dblTotalCancelado + CDbl(txtValorDoCancelamento.Text)
        TrocaCorParaZebrado lblSombra
    Else
'        txtValorDoCancelamento.Visible = False
'        txtDataDoCancelamento.Visible = False
'        txtNumeroDoCancelamento.Visible = False
'        lblSombra.Visible = False
        Detail.Visible = False
    End If
End Sub

Private Sub GroupHeader1_Format()
strProgramaDeTrabalho = ""
strProgramaDeTrabalho = Trim(txtintCodigoReduzido.Text) & " - " & Trim(txtstrCodigo.Text)
txt_strProgramaDeTrabalho = strProgramaDeTrabalho
TrocaCorParaZebrado lblSombra1
End Sub

Private Sub GroupHeader2_Format()
    If txtNumeroDaReserva.Text <> "" Then
        txtNumeroDaReserva.Text = "00000" & Trim(txtNumeroDaReserva.Text)
    End If
    txtValorReservado.Text = gstrConvVrDoSql(txtValorReservado.Text)
    If txtPKId <> "" Then
        txt_SaldoReservadoCancelado = gstrConvVrDoSql(dblSaldoReservadoCancelado)
    End If
    TrocaCorParaZebrado lblSombra2
End Sub

Private Sub PageFooter_Format()
    lblPagina.Caption = "Página " & pageNumber
End Sub

Private Sub PageHeader_Format()
    lblDataHora = gstrDataDoSistema(True, , True)
End Sub

Private Function dblSaldoReservadoCancelado() As Double
dblSaldoReservadoCancelado = 0
dblTotalReservado = 0
dblTotalCancelado = 0
    Dim strSQL          As String
    Dim adoResultado    As ADODB.Recordset
    strSQL = ""
    strSQL = strSQL & " SELECT SUM(dblValor) AS TotalReservado "
    strSQL = strSQL & " FROM "
    strSQL = strSQL & gstrReservaDotacao
    strSQL = strSQL & " WHERE "
    strSQL = strSQL & " PKId = " & Val(txtPKId)
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
        With adoResultado
            If .EOF = False Then
                If !TotalReservado <> "Null" Then
                    dblTotalReservado = gstrConvVrDoSql(!TotalReservado)
                End If
            End If
        End With
    End If
    adoResultado.Close
    
    strSQL = ""
    strSQL = strSQL & " SELECT SUM(dblValor) AS TotalCancelado "
    strSQL = strSQL & " FROM "
    strSQL = strSQL & gstrReservaDotacaoLiberada
    strSQL = strSQL & " WHERE "
    strSQL = strSQL & " intReservaDotacao = " & Val(txtPKId)
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
        With adoResultado
            If .EOF = False Then
                If !TotalCancelado <> "Null" Then
                    dblTotalCancelado = gstrConvVrDoSql(!TotalCancelado)
                End If
            End If
        End With
    End If
    dblSaldoReservadoCancelado = dblTotalReservado - dblTotalCancelado
End Function

Private Sub ReportFooter_Format()
    MostraEmissorRelatorio Me
End Sub
