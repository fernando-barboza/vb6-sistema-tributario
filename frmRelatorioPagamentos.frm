VERSION 5.00
Begin VB.Form frmRelatorioPagamentos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relatório de Pagamentos"
   ClientHeight    =   1500
   ClientLeft      =   4200
   ClientTop       =   5460
   ClientWidth     =   4995
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1500
   ScaleWidth      =   4995
   Begin VB.Frame fra_Pagamentos 
      Height          =   1275
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4995
      Begin VB.TextBox txtdtmFinal 
         Height          =   285
         Left            =   3750
         MaxLength       =   10
         TabIndex        =   6
         Top             =   780
         Width           =   1035
      End
      Begin VB.TextBox txtdtmInicial 
         Height          =   285
         Left            =   1620
         MaxLength       =   10
         TabIndex        =   4
         Top             =   780
         Width           =   1035
      End
      Begin VB.TextBox txtstrInscricao 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "#.##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Left            =   1620
         MaxLength       =   20
         TabIndex        =   2
         Top             =   300
         Width           =   3180
      End
      Begin VB.Label lblEmpenhoFinal 
         AutoSize        =   -1  'True
         Caption         =   "Fim"
         Height          =   195
         Left            =   3360
         TabIndex        =   5
         Top             =   810
         Width           =   240
      End
      Begin VB.Label lblEmpenhoInicial 
         AutoSize        =   -1  'True
         Caption         =   "Início"
         Height          =   195
         Left            =   1080
         TabIndex        =   3
         Top             =   810
         Width           =   405
      End
      Begin VB.Label lblstrInscricao 
         AutoSize        =   -1  'True
         Caption         =   "Inscrição Cadastral"
         Height          =   195
         Left            =   150
         TabIndex        =   1
         Top             =   375
         Width           =   1350
      End
   End
End
Attribute VB_Name = "frmRelatorioPagamentos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub txtdtmFinal_GotFocus()
    MarcaCampo txtdtmFinal
End Sub

Private Sub txtdtmFinal_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "D", txtdtmFinal
End Sub

Private Sub txtdtmFinal_LostFocus()
    txtdtmFinal = gstrDataFormatada(txtdtmFinal)
End Sub

Private Sub txtdtmInicial_GotFocus()
    MarcaCampo txtdtmInicial
End Sub

Private Sub txtdtmInicial_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "D", txtdtmInicial
End Sub

Private Sub txtdtmInicial_LostFocus()
    txtdtmInicial = gstrDataFormatada(txtdtmInicial)
End Sub

Private Sub txtstrInscricao_GotFocus()
    MarcaCampo txtstrInscricao
End Sub

Private Sub txtstrInscricao_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrInscricao
End Sub
    
Private Function blnPeriodoOk() As Boolean
    
    If Len(txtstrInscricao) = 0 Then
        ExibeMensagem "A Inscrição deve ser informada."
        txtstrInscricao.SetFocus
        Exit Function
    End If
    
    If Len(txtdtmInicial) > 0 Or Len(txtdtmFinal) > 0 Then
        If gblnDataValida(txtdtmInicial) = False Then
            ExibeMensagem "Data inicial incorreta."
            txtdtmInicial.SetFocus
            Exit Function
        ElseIf gblnDataValida(txtdtmFinal) = False Then
            ExibeMensagem "Data final incorreta."
            txtdtmFinal.SetFocus
            Exit Function
        ElseIf CVDate(txtdtmFinal) < CVDate(txtdtmInicial) Then
            ExibeMensagem "Data inicial não poder menor que a data final."
            txtdtmInicial.SetFocus
            Exit Function
        End If
    End If
    
    blnPeriodoOk = True
    
End Function

Private Sub ImprimePagamentos()
Dim strPeriodo             As String
Dim strSql                 As String
Dim adoResultado           As ADODB.Recordset

Dim intFor                 As Integer
Dim vetPagamentos()        As String

Dim strInscricoes          As String
Dim strAcordosParaConsulta As String

    Set gobjBanco = New clsBanco
        
    'Vamos obter os Pkids das inscricoes para fazer consulta de acordos
    strSql = "SELECT  LA.Pkid " & _
             "FROM " & gstrLancamentoAlfa & " LA " & _
             "WHERE LA.strInscricao = '" & String(gintLenInscricao - Len(Trim(txtstrInscricao.Text)), "0") & UCase(txtstrInscricao.Text) & "' AND " & _
             "(LA.INTUTILIZACAO = " & TYP_ACORDO & " OR (LA.INTUTILIZACAO <> " & TYP_ACORDO & " AND LA.bytNaoInscreveda = 0)) "
    
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        If Not adoResultado.EOF Then
            Do While Not adoResultado.EOF
                strAcordosParaConsulta = strAcordosParaConsulta & adoResultado("Pkid").Value & ","
                adoResultado.MoveNext
            Loop
            strAcordosParaConsulta = Mid(strAcordosParaConsulta, 1, Len(strAcordosParaConsulta) - 1)
        Else
            ExibeMensagem "A Inscrição Cadastral não existe."
            Exit Sub
        End If
    End If
    
ConsultarAcordos:

    'Vamos obter os acordos, caso exista
    strSql = "SELECT  LV.intLancamentoAlfaAcordo " & _
             "FROM " & gstrLancamentoValor & " LV, " & gstrLancamentoAlfa & " LA " & _
             "WHERE LV.intLancamentoAlfa = LA.pkid AND " & _
             "LA.Pkid IN (" & strAcordosParaConsulta & ") AND Not LV.intLancamentoAlfaAcordo Is Null " & _
             "GROUP BY LV.intLancamentoAlfaAcordo "
    
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        If Not adoResultado.EOF Then
            strAcordosParaConsulta = Space$(0)
            Do While Not adoResultado.EOF
                strInscricoes = strInscricoes & adoResultado("intlancamentoalfaacordo").Value & ","
                strAcordosParaConsulta = strAcordosParaConsulta & adoResultado("intlancamentoalfaacordo").Value & ","
                adoResultado.MoveNext
            Loop
            strAcordosParaConsulta = Mid(strAcordosParaConsulta, 1, Len(strAcordosParaConsulta) - 1)
            GoTo ConsultarAcordos
        End If
    End If
    
    strSql = ""
    strSql = "SELECT LA.Pkid, LA.strInscricao, LA.intUtilizacao, LA.strComposicaoDaReceita strComposicao, LA.intExercicio, LA.strEmissao, LA.strNomeProprietario strContribuinte, LA.strNumeroAviso, "
    strSql = strSql & "LV.intParcela, LV.dtmDtVencimento, LP.DTMDTMOVIMENTO, LP.DTMDTPAGAMENTO, (LP.DBLVALORPRINCIPAL + LP.DBLVALORMULTA + LP.DBLVALORJUROS + LP.DBLVALORCORRECAO) dblValor, LP.STROBSERVACAO "
    strSql = strSql & "FROM " & gstrLancamentoAlfa & " LA, " & gstrLancamentoValor & " LV, " & gstrLancamentoPagamento & " LP "
    strSql = strSql & "WHERE LA.strInscricao = '" & String(gintLenInscricao - Len(Trim(txtstrInscricao.Text)), "0") & UCase(txtstrInscricao.Text) & "' AND LV.intLancamentoAlfa = LA.Pkid AND LP.intLancamentoValor = LV.Pkid " & _
                      " AND (LA.INTUTILIZACAO = " & TYP_ACORDO & " OR (LA.INTUTILIZACAO <> " & TYP_ACORDO & " AND LA.bytNaoInscreveda = 0)) "

    If Len(txtdtmInicial.Text) > 0 Then
        strSql = strSql & " AND LP.dtmDtPagamento BETWEEN " & gstrConvDtParaSql(txtdtmInicial) & " AND " & gstrConvDtParaSql(txtdtmFinal)
    End If
    
    'Consulta que retorna os acordos
    If Len(strInscricoes) > 0 Then

        strInscricoes = Mid(strInscricoes, 1, Len(strInscricoes) - 1)

        strSql = strSql & " UNION ALL "
        strSql = strSql & "SELECT LA.Pkid, LA.strInscricao, LA.intUtilizacao, LA.strComposicaoDaReceita strComposicao, LA.intExercicio, LA.strEmissao, LA.strNomeProprietario strContribuinte, LA.strNumeroAviso, "
        strSql = strSql & "LV.intParcela, LV.dtmDtVencimento, LP.DTMDTMOVIMENTO, LP.DTMDTPAGAMENTO, (LP.DBLVALORPRINCIPAL + LP.DBLVALORMULTA + LP.DBLVALORJUROS + LP.DBLVALORCORRECAO) dblValor, LP.STROBSERVACAO "
        strSql = strSql & "FROM " & gstrLancamentoAlfa & " LA, " & gstrLancamentoValor & " LV, " & gstrLancamentoPagamento & " LP "
        strSql = strSql & "WHERE LA.Pkid IN (" & strInscricoes & ") AND LV.intLancamentoAlfa = LA.Pkid AND LP.intLancamentoValor = LV.Pkid "

        If Len(txtdtmInicial.Text) > 0 Then
            strSql = strSql & " AND LP.dtmDtPagamento BETWEEN " & gstrConvDtParaSql(txtdtmInicial) & " AND " & gstrConvDtParaSql(txtdtmFinal)
        End If

    End If

    'strsql = strsql & "ORDER BY strInscricao, strEmissao, intUtilizacao, strComposicao, intExercicio, intParcela "
    strSql = strSql & "ORDER BY strComposicao, intExercicio, strNumeroAviso, strEmissao, intUtilizacao, intParcela "
    

    'Vamos preencher o array para a impressao
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then

        If Not adoResultado.EOF Then
        
            For intFor = 0 To adoResultado.RecordCount - 1
        
                ReDim Preserve vetPagamentos(12, intFor)
                            
                vetPagamentos(0, intFor) = Space$(0) & adoResultado("strInscricao").Value
                vetPagamentos(1, intFor) = Space$(0) & adoResultado("intUtilizacao").Value
                vetPagamentos(2, intFor) = Space$(0) & adoResultado("strComposicao").Value
                vetPagamentos(3, intFor) = Space$(0) & adoResultado("intExercicio").Value
                vetPagamentos(4, intFor) = Space$(0) & adoResultado("strEmissao").Value
                vetPagamentos(5, intFor) = Space$(0) & adoResultado("strContribuinte").Value
                vetPagamentos(6, intFor) = Space$(0) & adoResultado("intParcela").Value
                vetPagamentos(7, intFor) = Space$(0) & adoResultado("dtmDtVencimento").Value
                vetPagamentos(8, intFor) = Space$(0) & adoResultado("dtmDtMovimento").Value
                vetPagamentos(9, intFor) = Space$(0) & adoResultado("dtmDtPagamento").Value
                vetPagamentos(10, intFor) = Space$(0) & adoResultado("dblValor").Value
                vetPagamentos(11, intFor) = Space$(0) & adoResultado("strObservacao").Value
                vetPagamentos(12, intFor) = Space$(0) & adoResultado("strNumeroAviso").Value
                
                adoResultado.MoveNext
                
            Next
        Else
            ExibeMensagem "Não existem parcelas pagas para esta Inscrição no período."
            Exit Sub
        End If
    Else
        Exit Sub
    End If
    
    strPeriodo = ""
    
    If Len(txtdtmInicial.Text) > 0 Then
        If txtdtmInicial.Text = txtdtmFinal.Text Then
            strPeriodo = "No dia: " & txtdtmInicial.Text
        Else
            strPeriodo = "No periodo de: " & txtdtmInicial.Text & " até " & txtdtmFinal.Text
        End If
    End If
    
    ImprimeRelatorioPorArray rptRelatorioPagamentos, vetPagamentos, "Relatório de Pagamentos - " & strPeriodo
    
End Sub

Public Sub MantemForm(ByVal strModoOperacao As String)
    Select Case UCase(strModoOperacao)
        Case UCase(gstrImprimir)
            If blnPeriodoOk Then ImprimePagamentos
        Case UCase(gstrNovo)
            LimpaObjeto
    End Select
End Sub

Private Sub LimpaObjeto()
    txtstrInscricao = ""
    txtdtmInicial = ""
    txtdtmFinal = ""
    txtstrInscricao.SetFocus
End Sub

Private Sub Form_Activate()
    gintCodSeguranca = 1259
    HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrImprimir
End Sub

