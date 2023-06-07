VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmRecebeMovBancario 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Receber Movimento Bancário"
   ClientHeight    =   2355
   ClientLeft      =   3735
   ClientTop       =   4935
   ClientWidth     =   6390
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2355
   ScaleWidth      =   6390
   Begin VB.CommandButton cmdProvisorio 
      Caption         =   "Provisório"
      Height          =   300
      Left            =   210
      TabIndex        =   8
      Top             =   2430
      Visible         =   0   'False
      Width           =   6000
   End
   Begin MSComctlLib.ProgressBar pgr_Status 
      Height          =   165
      Left            =   180
      TabIndex        =   7
      Top             =   1950
      Visible         =   0   'False
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Frame fra_Arquivo 
      Caption         =   " Arquivo de leitura "
      Height          =   840
      Left            =   180
      TabIndex        =   3
      Top             =   1050
      Width           =   6000
      Begin VB.TextBox txt_Arquivo 
         Height          =   285
         Left            =   1035
         TabIndex        =   5
         Top             =   315
         Width           =   4410
      End
      Begin VB.CommandButton cmd_Arquivo 
         Caption         =   "..."
         Height          =   300
         Left            =   5460
         Picture         =   "frmRecebeMovBancario.frx":0000
         TabIndex        =   6
         TabStop         =   0   'False
         ToolTipText     =   "Localiza Arquivo de Baixa Automática"
         Top             =   315
         Width           =   345
      End
      Begin VB.Label lbl_Arquivo 
         AutoSize        =   -1  'True
         Caption         =   "Localização"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   405
         Width           =   855
      End
   End
   Begin VB.Frame fra_Dados 
      Caption         =   " Dados para o recebimento "
      Height          =   840
      Left            =   165
      TabIndex        =   0
      Top             =   105
      Width           =   6015
      Begin VB.TextBox txt_DtMovimento 
         Height          =   285
         Left            =   1605
         MaxLength       =   10
         TabIndex        =   2
         Top             =   315
         Width           =   975
      End
      Begin VB.Label lbl_DtMovimento 
         AutoSize        =   -1  'True
         Caption         =   "Data do Movimento"
         Height          =   195
         Left            =   135
         TabIndex        =   1
         Top             =   390
         Width           =   1395
      End
   End
   Begin MSComDlg.CommonDialog dlgArquivo 
      Left            =   0
      Top             =   30
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lbl_Status 
      Alignment       =   2  'Center
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   210
      TabIndex        =   9
      Top             =   2130
      Visible         =   0   'False
      Width           =   5955
   End
End
Attribute VB_Name = "frmRecebeMovBancario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mblnPrimeiraVez      As Boolean
Dim strUltimoCaminho     As String
Dim intCodBaixaAposVcto  As Integer
Dim intCodBaixaAntesVcto As Integer

Public DesContarTarifa As Boolean

Private Sub cmd_arquivo_Click()
    
    dlgArquivo.CancelError = True
    dlgArquivo.DialogTitle = "Selecione o arquivo"
    dlgArquivo.InitDir = strUltimoCaminho
    dlgArquivo.Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
    dlgArquivo.flags = cdlOFNOverwritePrompt + cdlOFNHideReadOnly
    
    On Error GoTo err_cmd_Arquivo_Click
    
    dlgArquivo.ShowOpen
    txt_Arquivo = dlgArquivo.Filename
    strUltimoCaminho = Replace(dlgArquivo.Filename, dlgArquivo.FileTitle, "")
    Exit Sub

err_cmd_Arquivo_Click:
    If Err.Number = 32755 Then
        txt_Arquivo = ""
    End If
    
End Sub

Private Sub Form_Activate()
    gintCodSeguranca = 1154
    HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrLerArquivo
End Sub

Private Sub Form_Deactivate()
    HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrLerArquivo
End Sub

Private Sub Form_Load()
    txt_DtMovimento = gstrDataDoSistema
    strUltimoCaminho = "C:\"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrLerArquivo
    mblnPrimeiraVez = False
End Sub

Private Sub LeMovimentoBancarioFebrabam()
Dim strLinha               As String
Dim lngLinha               As Long
Dim lngSize                As Long
Dim strBanco               As String
Dim strCodEspecifico       As String

Dim blnTipoJaSomado        As Boolean
Dim strSQL                 As String
Dim strSqlSub              As String

Dim adoResultado           As New ADODB.Recordset
Dim adoAtualizacao         As New ADODB.Recordset

Dim bytQtdeTipos           As Byte ' Tipos (A,G,Z)
Dim dblPorcentagemDif      As Double
Dim dblValorDiferenca      As Double
Dim dblValorDoArquivo      As Double

Dim dblSomaDosValores      As Double

Dim dblValorPrincipal      As Double
Dim dblValorMulta          As Double
Dim dblValorJuros          As Double
Dim dblValorCorrecao       As Double
Dim dblValorDesconto       As Double
Dim dblValorCorreto        As Double

Dim lngPkidContaBancaria   As Long
Dim blnExisteGuia          As Boolean
Dim blnLayoutNovo          As Boolean
Dim strLote                As String

    On Error GoTo err_BaixaAutomatica
    
    bytQtdeTipos = 0
    lngLinha = 0
    blnTipoJaSomado = False
    
    Open txt_Arquivo For Input As #1
    
    lngSize = FileLen(txt_Arquivo) / 150
    
    pgr_Status.Value = 0
    pgr_Status.Visible = True
    pgr_Status.Max = Abs(lngSize)
    lbl_Status.Visible = True
    
    Set gobjBanco = New clsBanco
    
    gobjBanco.ExecutaBeginTrans
    
    Do While Not EOF(1)

        Line Input #1, strLinha

        If Len(strLinha) = 0 Then
            GoTo ProximaLinha
        End If
        
        'Vamos verificar se a linha contem 150 posicoes
        If Len(strLinha) <> 150 Then
            ExibeMensagem "Ocorreu um Erro na leitura do arquivo. A operação não foi concluída."
            Screen.MousePointer = vbDefault
            pgr_Status.Visible = False
            lbl_Status.Visible = False
            gobjBanco.ExecutaRollbackTrans
            Close #1
            Exit Sub
        End If
        
        If Mid(strLinha, 1, 1) = "A" Then
            bytQtdeTipos = bytQtdeTipos + 1
            strLote = strLinha
            strBanco = Trim(Mid(strLinha, 43, 3))
        ElseIf Mid(strLinha, 1, 1) = "G" Then
            
            blnLayoutNovo = Len(Trim(Mid(strLinha, 105, 13))) > 0
            
            If Len(strLote) > 4 Then
                strLote = IIf(blnLayoutNovo, Mid(strLote, 76, 4), Mid(strLote, 74, 4))
                
                'Vamos verifcar se este lote ja foi importado
                strSQL = "SELECT Count(*) Total FROM " & gstrResumoBancario & " WHERE dtmData = " & gstrConvDtParaSql(txt_DtMovimento.Text) & " AND intLote = " & strLote
                If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
                    If adoResultado("Total") > 0 Then
                        ExibeMensagem "O lote " & strLote & " já foi importado no dia " & txt_DtMovimento.Text & ". A operação não concluída."
                        Screen.MousePointer = vbDefault
                        pgr_Status.Visible = False
                        lbl_Status.Visible = False
                        gobjBanco.ExecutaRollbackTrans
                        Close #1
                        Exit Sub
                    End If
                End If
                
            End If
            
            If Not blnTipoJaSomado Then
                bytQtdeTipos = bytQtdeTipos + 1
                blnTipoJaSomado = True
                     
                'Vamos buscar o Pkid referente à Conta em tblContaBancaria
                strSQL = "SELECT Pkid FROM " & gstrContaBancaria & " WHERE strContaRetorno = '" & Trim(Mid(strLinha, 2, 20)) & "'"
                If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
                    If adoResultado.EOF Then
                        ExibeMensagem "Não foi encontrada Conta referente à este arquivo. A operação não concluída."
                        Screen.MousePointer = vbDefault
                        pgr_Status.Visible = False
                        lbl_Status.Visible = False
                        gobjBanco.ExecutaRollbackTrans
                        Close #1
                        Exit Sub
                    Else
                        lngPkidContaBancaria = adoResultado("Pkid").Value
                    End If
                End If
            End If
            
            blnExisteGuia = True
                         
            strSqlSub = ""
            strSqlSub = "SELECT " & gstrTOPnSQLServer(1) & " Pkid FROM " & gstrGuias & " WHERE strCodBarra = '" & IIf(blnLayoutNovo, Mid(strLinha, 38, 44), Mid(strLinha, 34, 44)) & "' ORDER BY Pkid Desc"
            strSqlSub = gstrTOPnOracle(strSqlSub, 1)
            
            'Vamos buscar os dados referentes às guias
            strSQL = "SELECT LA.intComposicaoDaReceita, LA.intExercicio, LV.dblValor ValorOrig, LV.dtmDtVencimento, LV.intMoeda, LV.Pkid intLancamentoValor, LV.intParcela, LG.intGuias, LV.DtmDtVencimento, LG.DBLVALORPRINCIPAL, LG.DBLVALORMULTA, LG.DBLVALORJUROS, LG.Dblvalorcorrecao, LG.DBLVALORDESCONTO, (LG.DBLVALORPRINCIPAL + LG.DBLVALORMULTA + LG.DBLVALORJUROS + LG.Dblvalorcorrecao - LG.DBLVALORDESCONTO) dblCorreto, " & _
                            gstrCONVERT(CDT_numeric, "LA.STRNUMEROAVISO") & " strNumeroAviso," & _
                            "(SELECT SUM(LG2.DBLVALORPRINCIPAL + LG2.DBLVALORMULTA + LG2.DBLVALORJUROS + LG2.Dblvalorcorrecao - LG2.DBLVALORDESCONTO) FROM " & gstrLancamentoGuias & " LG2 WHERE LG2.intguias = LG.intGuias) dblTotal " & _
                     "FROM " & gstrLancamentoGuias & " LG, " & _
                             gstrLancamentoValor & " LV, " & _
                             gstrLancamentoAlfa & " LA " & _
                     "WHERE LV.Pkid = LG.INTLANCAMENTOVALOR AND " & _
                     "      LA.Pkid = LV.INTLANCAMENTOALFA AND " & _
                           "LG.intGuias = (" & strSqlSub & ")"
                     
            If gobjBanco.CriaADO(strSQL, 10, adoResultado) Then
                If adoResultado.EOF Then
                
                    strSQL = "Select "
                    strSQL = strSQL & "BAC.intPosInicial, "
                    strSQL = strSQL & "BAC.intPosFinal "
                    strSQL = strSQL & "from "
                    strSQL = strSQL & "tblBaixaAuxiliar BA, "
                    strSQL = strSQL & "tblBaixaAuxiliarCampos BAC "
                    strSQL = strSQL & "Where "
                    strSQL = strSQL & "BA.pkid = BAC.intBaixaAuxiliar and "
                    strSQL = strSQL & "BA.strBanco = '" & strBanco & "' and "
                    strSQL = strSQL & "BA.strContaBancaria = '" & Trim(Mid(strLinha, 2, 20)) & "'"
                    strSQL = strSQL & " Order By BAC.Pkid "
                    
                    strCodEspecifico = ""
                    
                    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
                        If Not adoResultado.EOF Then
                            Do While Not adoResultado.EOF
                                strCodEspecifico = strCodEspecifico & Mid(strLinha, gstrENulo(adoResultado!intPosInicial), gstrENulo(adoResultado!intPosFinal))
                                adoResultado.MoveNext
                            Loop
                        Else
                            strCodEspecifico = IIf(blnLayoutNovo, Mid(strLinha, 63, 14), Mid(strLinha, 59, 14))
                        End If
                    End If

                 
                    strSqlSub = ""
                    strSqlSub = "SELECT " & gstrTOPnSQLServer(1) & " Pkid FROM " & gstrGuias & " WHERE strCodBarraEsp = '" & strCodEspecifico & "' ORDER BY Pkid Desc"
                    strSqlSub = gstrTOPnOracle(strSqlSub, 1)
                 
                    'Vamos buscar os dados referentes às guias apenas com a parte especifica do codigo de barras
                    strSQL = "SELECT LA.intComposicaoDaReceita, LA.intExercicio, LV.dblValor ValorOrig, LV.dtmDtVencimento, LV.intMoeda, LV.Pkid intLancamentoValor, LV.intParcela, LG.intGuias, LV.DtmDtVencimento, LG.DBLVALORPRINCIPAL, LG.DBLVALORMULTA, LG.DBLVALORJUROS, LG.Dblvalorcorrecao, LG.DBLVALORDESCONTO, (LG.DBLVALORPRINCIPAL + LG.DBLVALORMULTA + LG.DBLVALORJUROS + LG.Dblvalorcorrecao - LG.DBLVALORDESCONTO) dblCorreto, " & _
                                       gstrCONVERT(CDT_numeric, "LA.STRNUMEROAVISO") & " strNumeroAviso," & _
                                       "(SELECT SUM(LG2.DBLVALORPRINCIPAL + LG2.DBLVALORMULTA + LG2.DBLVALORJUROS + LG2.Dblvalorcorrecao - LG2.DBLVALORDESCONTO) FROM " & gstrLancamentoGuias & " LG2 WHERE LG2.intguias = LG.intGuias) dblTotal " & _
                             "FROM " & gstrLancamentoGuias & " LG, " & _
                                       gstrLancamentoValor & " LV, " & _
                                       gstrLancamentoAlfa & " LA " & _
                             "WHERE LV.Pkid = LG.INTLANCAMENTOVALOR AND " & _
                             "      LA.Pkid = LV.INTLANCAMENTOALFA AND " & _
                                    "LG.intGuias = (" & strSqlSub & ")"
                                                        
                    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
                        If adoResultado.EOF Then
                 
                            blnExisteGuia = False
                            'ExibeMensagem "Não foi encontrada Conta referente à este arquivo. A operação não concluída."
                            'Screen.MousePointer = vbDefault
                            'pgr_Status.Visible = False
                            'gobjBanco.ExecutaRollbackTrans
                            'Close #1
                            'Exit Sub
                        End If
                    End If
                End If
                
                dblPorcentagemDif = 0
                dblSomaDosValores = 0
                dblValorDiferenca = 0
                dblValorDoArquivo = ConverteValorDoArquivo(IIf(blnLayoutNovo, Mid(strLinha, 82, 12), Mid(strLinha, 78, 12)))
                
                If blnExisteGuia Then 'Verifica se o código de barras vindo do arquivo existe no banco
                
                    'Vamos verificar se existe diferenca de valores
                    If dblValorDoArquivo <> adoResultado("dblTotal").Value Then
                        'dblPorcentagemDif = (dblValorDoArquivo / adoResultado("dblTotal").Value)
                        'PROVISORIO
                        dblPorcentagemDif = dblValorDoArquivo - adoResultado("dblTotal").Value
                    End If
                    
                    'Vamos gravar um registro em Movimento Bancario para cada Guia
                    Do While Not adoResultado.EOF
                        
                        'Caso exista diferenca dos valores, vamos tirar a diferenca proporcionalmente
                        If dblPorcentagemDif <> 0 Then
                            'dblValorPrincipal = gstrConvVrDoSql((adoResultado("dblvalorprincipal") * dblPorcentagemDif), 2)
                            'dblValorMulta = gstrConvVrDoSql((adoResultado("dblvalorMulta") * dblPorcentagemDif), 2)
                            'dblValorJuros = gstrConvVrDoSql((adoResultado("dblvalorJuros") * dblPorcentagemDif), 2)
                            'dblValorCorrecao = gstrConvVrDoSql((adoResultado("dblvalorCorrecao") * dblPorcentagemDif), 2)
                            'dblValorDesconto = gstrConvVrDoSql((adoResultado("dblvalorDesconto") * dblPorcentagemDif), 2)
                            'PROVISORIO
                            dblValorPrincipal = gstrConvVrDoSql((dblValorDoArquivo - dblPorcentagemDif), 2)
                            dblValorMulta = gstrConvVrDoSql(dblPorcentagemDif, 2)
                            dblValorJuros = 0
                            dblValorCorrecao = 0
                            dblValorDesconto = 0
                        Else
                            dblValorPrincipal = adoResultado("dblValorPrincipal").Value 'ConverteValorDoArquivo(IIf(blnLayoutNovo, Mid(strLinha, 82, 12), Mid(strLinha, 78, 12)))
                            dblValorMulta = adoResultado("dblValorMulta").Value
                            dblValorJuros = adoResultado("dblValorJuros").Value
                            dblValorCorrecao = adoResultado("dblValorCorrecao").Value
                            dblValorDesconto = adoResultado("dblValorDesconto").Value
                        End If
                        
                        dblSomaDosValores = dblSomaDosValores + (dblValorPrincipal + dblValorMulta + dblValorJuros + dblValorCorrecao + dblValorDesconto)
                        
                        'Se este for o ultimo registro, vamos verificar se ha diferenca dos valores calculados
                        If adoResultado.RecordCount = adoResultado.AbsolutePosition Then
                            'Caso depois do acerto de valores ainda exista diferenca, a jogaremos no ultimo registro
                            If CCur(dblSomaDosValores) <> CCur(dblValorDoArquivo) Then
                            
                                dblValorDiferenca = dblValorDoArquivo - dblSomaDosValores
                        
                                'strSQL = "UPDATE " & gstrMovimentoBancario
                                If dblValorPrincipal > 0 Then
                                    dblValorPrincipal = dblValorPrincipal + dblValorDiferenca
                                ElseIf dblValorMulta > 0 Then
                                    dblValorMulta = dblValorMulta + dblValorDiferenca
                                ElseIf dblValorJuros > 0 Then
                                    dblValorJuros = dblValorJuros + dblValorDiferenca
                                Else
                                    dblValorCorrecao = dblValorCorrecao + dblValorDiferenca
                                End If
                            End If
                        End If
                        
                        strSQL = gstrStoredProcedure("sp_AtualizaParcela", adoResultado("intComposicaoDaReceita").Value & ", " & adoResultado("intExercicio").Value & ", " & adoResultado("intLancamentoValor").Value & ", " & gstrConvDtParaSql(adoResultado("dtmDtVencimento").Value) & ", " & gstrConvDtParaSql(ConverteDataDoArquivo(IIf(blnLayoutNovo, Mid(strLinha, 22, 8), Mid(strLinha, 22, 6)), blnLayoutNovo)) & ", " & gstrConvVrParaSql(adoResultado("ValorOrig").Value) & ", " & adoResultado("intMoeda").Value, True)
                        If gobjBanco.CriaADO(strSQL, 80, adoAtualizacao) Then
                            If gstrConvVrDoSql(adoResultado("ValorOrig").Value, , , True) = 0 Then
                                dblValorPrincipal = dblValorDoArquivo
                                dblValorCorreto = dblValorDoArquivo
                                dblValorMulta = 0
                            Else
                                dblValorCorreto = Space$(0) & gstrConvVrDoSql(Val(gstrConvVrParaSql(adoAtualizacao("dblValorPrincipal").Value)) + Val(gstrConvVrParaSql(adoAtualizacao("dblValorMulta").Value)) + Val(gstrConvVrParaSql(adoAtualizacao("dblValorJuros").Value)) + Val(gstrConvVrParaSql(adoAtualizacao("dblValorCorrecao").Value)))
                            End If
                        Else
                            Screen.MousePointer = vbDefault
                            pgr_Status.Visible = False
                            lbl_Status.Visible = False
                            gobjBanco.ExecutaRollbackTrans
                            Close #1
                            Exit Sub

                        End If

                            gobjBanco.Execute "INSERT INTO " & gstrMovimentoBancario & "(dtmDtMovimento, intContaBancaria, intLote, dtmDtPagamento, dblPrincipal, dblMulta, dblJuros, dblCorrecao, strCodigoDeBarras, dblCorreto, intLancamentoValor, intGuias, intCodigoBaixa, intDigito, BytTipo, dtmDtAtualizacao, lngCodUsr, bitGuia, bitProcessado) " & _
                                                                                        "VALUES(" & gstrConvDtParaSql(txt_DtMovimento.Text) & "," & lngPkidContaBancaria & "," & strLote & "," & gstrConvDtParaSql(ConverteDataDoArquivo(IIf(blnLayoutNovo, Mid(strLinha, 22, 8), Mid(strLinha, 22, 6)), blnLayoutNovo)) & "," & gstrConvVrParaSql(dblValorPrincipal) & "," & gstrConvVrParaSql(dblValorMulta) & "," & gstrConvVrParaSql(dblValorJuros) & "," & gstrConvVrParaSql(dblValorCorrecao) & ",'" & IIf(blnLayoutNovo, Mid(strLinha, 38, 44), Mid(strLinha, 34, 44)) & "'," & gstrConvVrParaSql(dblValorCorreto) & "," & adoResultado("intLancamentoValor").Value & "," & adoResultado("intGuias").Value & "," & _
                                                                                        IIf(CDate(adoResultado("dtmDtVencimento").Value) < ConverteDataDoArquivo(IIf(blnLayoutNovo, Mid(strLinha, 22, 8), Mid(strLinha, 22, 6)), blnLayoutNovo), intCodBaixaAposVcto, intCodBaixaAntesVcto) & "," & gstrCalculaDigitoModulo10(Trim(adoResultado("strNumeroAviso").Value) & Format$(Trim(adoResultado("intParcela").Value), "000")) & _
                                                                                        "," & IIf(Mid(strLinha, 117, 1) = " ", "9", Mid(strLinha, 117, 1)) & "," & gstrConvDtParaSql(gstrDataDoSistema) & "," & glngCodUsr & ",1,0)"
                        adoResultado.MoveNext
                        
                    Loop
                    
                Else
                    
                   dblValorPrincipal = dblValorDoArquivo
                   dblValorMulta = 0
                   dblValorJuros = 0
                   dblValorCorrecao = 0
                   dblValorDesconto = 0
                    
                   gobjBanco.Execute "INSERT INTO " & gstrMovimentoBancario & "(dtmDtMovimento, intContaBancaria, intLote, dtmDtPagamento, dblPrincipal, dblMulta, dblJuros, dblCorrecao, strCodigoDeBarras, dblCorreto, intLancamentoValor, intGuias, intCodigoBaixa, intDigito, BytTipo, dtmDtAtualizacao, lngCodUsr, bitGuia, bitProcessado) " & _
                                                                                    "VALUES(" & gstrConvDtParaSql(txt_DtMovimento.Text) & "," & lngPkidContaBancaria & "," & strLote & "," & gstrConvDtParaSql(ConverteDataDoArquivo(IIf(blnLayoutNovo, Mid(strLinha, 22, 8), Mid(strLinha, 22, 6)), blnLayoutNovo)) & "," & gstrConvVrParaSql(dblValorPrincipal) & "," & gstrConvVrParaSql(dblValorMulta) & "," & gstrConvVrParaSql(dblValorJuros) & "," & gstrConvVrParaSql(dblValorCorrecao) & ",'" & IIf(blnLayoutNovo, Mid(strLinha, 38, 44), Mid(strLinha, 34, 44)) & "',0,NULL,NULL,NULL,0," & _
                                                                                    IIf(Mid(strLinha, 117, 1) = " ", "9", Mid(strLinha, 117, 1)) & "," & gstrConvDtParaSql(gstrDataDoSistema) & "," & glngCodUsr & ",0,0)"
                
                End If
                
            End If
             
        ElseIf Mid(strLinha, 1, 1) = "Z" Then
            bytQtdeTipos = bytQtdeTipos + 1
                         
            'Vamos verificar se ja existe um registro referente, caso exista vamos somar o valor
            strSQL = "SELECT Pkid, dblValor FROM " & gstrResumoBancario & " WHERE dtmData = " & gstrConvDtParaSql(txt_DtMovimento.Text) & " AND intLote = " & strLote & " AND intContaBancaria = " & lngPkidContaBancaria
            If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
                If Not adoResultado.EOF Then
             
                    'Vamos somar registro na tabela tblResumoBancario
                    strSQL = "UPDATE " & gstrResumoBancario & " SET dblValor = " & gstrConvVrParaSql(ConverteValorDoArquivo(Mid(strLinha, 8, 17)) + adoResultado("dblValor").Value) & " WHERE Pkid = " & adoResultado("Pkid").Value
                     
                Else
                
                    'Vamos gravar registro na tabela tblResumoBancario
                    strSQL = "INSERT INTO " & gstrResumoBancario & "(dtmData, intContaBancaria, intLote, dblValor, dtmDtAtualizacao, lngCodUsr) " & _
                             "VALUES(" & gstrConvDtParaSql(txt_DtMovimento.Text) & "," & lngPkidContaBancaria & "," & strLote & "," & gstrConvVrParaSql(ConverteValorDoArquivo(Mid(strLinha, 8, 17))) & "," & gstrConvDtParaSql(gstrDataDoSistema) & "," & glngCodUsr & ")"
                
                End If
            End If
            
            gobjBanco.Execute strSQL
            
            'Vamos verificar se a qtde gravada é igual ao total de registros
            If lngLinha + 1 <> Val(Mid(strLinha, 2, 6)) Then
                ExibeMensagem "A quantidade de registros processada não coincide com a informada no arquivo. A operação não concluída."
                Screen.MousePointer = vbDefault
                pgr_Status.Visible = False
                lbl_Status.Visible = False
                gobjBanco.ExecutaRollbackTrans
                Close #1
                Exit Sub
            End If
            
        End If
        
        lngLinha = lngLinha + 1
        
        DoEvents
        lbl_Status.Caption = lngLinha & " de " & lngSize
        Me.Refresh
        
        pgr_Status.Value = lngLinha
        
ProximaLinha:

    Loop
        
    Close #1
    
    pgr_Status.Visible = False
    lbl_Status.Visible = False
    
    'Vamos verificar se foram lidos todos os tipos de registro
    If bytQtdeTipos < 3 Then
        ExibeMensagem "O arquivo de leitura está incorreto."
        Screen.MousePointer = vbDefault
        gobjBanco.ExecutaRollbackTrans
        Close #1
        Exit Sub
    End If
    
    gobjBanco.ExecutaCommitTrans
    
    Screen.MousePointer = vbDefault
    
    Exit Sub

err_BaixaAutomatica:
    ExibeMensagem "Ocorreu um Erro na leitura do arquivo. A operação não foi concluída."
    Screen.MousePointer = vbDefault
    pgr_Status.Visible = False
    lbl_Status.Visible = False
    gobjBanco.ExecutaRollbackTrans
    Close #1
    
End Sub

Private Sub LeMovimentoBancarioDebitoAutomatico()
Dim strLinha               As String
Dim lngLinha               As Long
Dim lngSize                As Long
Dim strBanco               As String
Dim strContaBancaria       As String
Dim strCodEspecifico       As String

Dim blnTipoJaSomado        As Boolean
Dim strSQL                 As String
Dim strSqlSub              As String

Dim adoResultado           As New ADODB.Recordset
Dim adoAtualizacao         As New ADODB.Recordset

Dim bytQtdeTipos           As Byte ' Tipos (A,B,F,Z)
Dim dblPorcentagemDif      As Double
Dim dblValorDiferenca      As Double
Dim dblValorDoArquivo      As Double

Dim dblSomaDosValores      As Double

Dim dblValorPrincipal      As Double
Dim dblValorMulta          As Double
Dim dblValorJuros          As Double
Dim dblValorCorrecao       As Double
Dim dblValorDesconto       As Double
Dim dblValorCorreto        As Double

Dim lngPkidContaBancaria   As Long
Dim blnExisteGuia          As Boolean
Dim blnLayoutNovo          As Boolean
Dim strLote                As String
Dim intBanco               As Integer

    On Error GoTo err_BaixaAutomatica
    
    bytQtdeTipos = 0
    lngLinha = 0
    blnTipoJaSomado = False
    
    Open txt_Arquivo For Input As #1
    
    lngSize = FileLen(txt_Arquivo) / 150
    
    pgr_Status.Value = 0
    pgr_Status.Visible = True
    pgr_Status.Max = Abs(lngSize)
    lbl_Status.Visible = True
    
    Set gobjBanco = New clsBanco
    
    gobjBanco.ExecutaBeginTrans
    
    Do While Not EOF(1)

        Line Input #1, strLinha

        If Len(strLinha) = 0 Then
            GoTo ProximaLinha
        End If
        
        'Vamos verificar se a linha contem 150 posicoes
        If Len(strLinha) <> 150 Then
            ExibeMensagem "Ocorreu um Erro na leitura do arquivo. A operação não foi concluída."
            Screen.MousePointer = vbDefault
            pgr_Status.Visible = False
            lbl_Status.Visible = False
            gobjBanco.ExecutaRollbackTrans
            Close #1
            Exit Sub
        End If
        
        If Mid(strLinha, 1, 1) = "A" Then
            
            bytQtdeTipos = bytQtdeTipos + 1
            strLote = Trim(Mid(strLinha, 80, 2))
            strBanco = Trim(Mid(strLinha, 43, 3))
            strContaBancaria = Trim(Mid(strLinha, 3, 20))
            intBanco = Trim$(Mid$(strLinha, 43, 3))
            
            'Vamos buscar o Pkid referente à Conta em tblContaBancaria
            strSQL = "SELECT Pkid FROM " & gstrContaBancaria & " WHERE strContaRetorno = '" & Trim(Mid(strLinha, 3, 20)) & "'"
            If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
                If adoResultado.EOF Then
                    ExibeMensagem "Não foi encontrada Conta referente à este arquivo. A operação não concluída."
                    Screen.MousePointer = vbDefault
                    pgr_Status.Visible = False
                    lbl_Status.Visible = False
                    gobjBanco.ExecutaRollbackTrans
                    Close #1
                    Exit Sub
                Else
                    lngPkidContaBancaria = adoResultado("Pkid").Value
                End If
            End If
        
        ElseIf Mid(strLinha, 1, 1) = "B" Then
            
            If Not blnTipoJaSomado Then
                bytQtdeTipos = bytQtdeTipos + 1
                blnTipoJaSomado = True
            End If
            
            If InStr(1, Trim(Mid(strLinha, 2, 25)), "-") Then
                'critica
            Else
                strSQL = "SELECT * FROM " & gstrDebitoAutomatico & " WHERE strIdentificacaoDebAut = '" & Trim(Mid(strLinha, 2, 25)) & "'"
                
                If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
                    If Not adoResultado.EOF Then
                        If Trim(Mid(strLinha, 150, 1)) = "1" Then
                            gobjBanco.Execute "UPDATE " & gstrDebitoAutomatico & " SET strAgencia = '', strIdentificacaoBanco = '', dtmDtOpcao = '', intBanco = NULL WHERE PKId = " & adoResultado("PKId")
                        Else
                            gobjBanco.Execute "UPDATE " & gstrDebitoAutomatico & " SET strAgencia = '" & Trim(Mid(strLinha, 27, 4)) & "', strIdentificacaoBanco = '" & Trim(Mid(strLinha, 31, 14)) & "', dtmDtOpcao = " & gstrConvDtParaSql(ConverteDataDoArquivo(Trim(Mid(strLinha, 45, 8)), True)) & ", intBanco = " & intBanco & " WHERE PKId = " & adoResultado("PKId")
                        End If
                    End If
                End If
            End If
            
        ElseIf Mid(strLinha, 1, 1) = "F" Then
        
            If blnTipoJaSomado Then
                bytQtdeTipos = bytQtdeTipos + 1
                blnTipoJaSomado = False
            End If
            
            If Len(strLote) > 0 Then

                'Vamos verifcar se este lote ja foi importado
                strSQL = "SELECT Count(*) Total FROM " & gstrResumoBancario & " WHERE dtmData = " & gstrConvDtParaSql(txt_DtMovimento.Text) & " AND intLote = " & strLote
                If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
                    If adoResultado("Total") > 0 Then
                        ExibeMensagem "O lote " & strLote & " já foi importado no dia " & txt_DtMovimento.Text & ". A operação não concluída."
                        Screen.MousePointer = vbDefault
                        pgr_Status.Visible = False
                        lbl_Status.Visible = False
                        gobjBanco.ExecutaRollbackTrans
                        Close #1
                        Exit Sub
                    End If
                End If

            End If
            
            blnExisteGuia = True
                         
            strSqlSub = ""
            strSqlSub = "SELECT " & gstrTOPnSQLServer(1) & " Pkid FROM " & gstrGuias & " WHERE strCodBarra = '" & Trim(Mid(strLinha, 2, 25)) & Trim(Mid(strLinha, 27, 4)) & Trim(Mid(strLinha, 45, 8)) & "' ORDER BY Pkid Desc"
            strSqlSub = gstrTOPnOracle(strSqlSub, 1)
            
            'Vamos buscar os dados referentes às guias
            strSQL = "SELECT LA.intComposicaoDaReceita, LA.intExercicio, LV.dblValor ValorOrig, LV.dtmDtVencimento, LV.intMoeda, LV.Pkid intLancamentoValor, LV.intParcela, LG.intGuias, LV.DtmDtVencimento, LG.DBLVALORPRINCIPAL, LG.DBLVALORMULTA, LG.DBLVALORJUROS, LG.Dblvalorcorrecao, LG.DBLVALORDESCONTO, (LG.DBLVALORPRINCIPAL + LG.DBLVALORMULTA + LG.DBLVALORJUROS + LG.Dblvalorcorrecao - LG.DBLVALORDESCONTO) dblCorreto, " & _
                            gstrCONVERT(CDT_numeric, "LA.STRNUMEROAVISO") & " strNumeroAviso," & _
                            "(SELECT SUM(LG2.DBLVALORPRINCIPAL + LG2.DBLVALORMULTA + LG2.DBLVALORJUROS + LG2.Dblvalorcorrecao - LG2.DBLVALORDESCONTO) FROM " & gstrLancamentoGuias & " LG2 WHERE LG2.intguias = LG.intGuias) dblTotal " & _
                     "FROM " & gstrLancamentoGuias & " LG, " & _
                             gstrLancamentoValor & " LV, " & _
                             gstrLancamentoAlfa & " LA " & _
                     "WHERE LV.Pkid = LG.INTLANCAMENTOVALOR AND " & _
                     "      LA.Pkid = LV.INTLANCAMENTOALFA AND " & _
                           "LG.intGuias = (" & strSqlSub & ")"
                     
            If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
                If adoResultado.EOF Then
                
                   'rotina que despreza o dia do lancamento
                   
                   strSqlSub = ""
                   strSqlSub = "SELECT " & gstrTOPnSQLServer(1) & " Pkid FROM " & gstrGuias & " WHERE " & strSUBSTRING & "(strCodBarra, 1, 26) = '" & Left$(Trim(Mid(strLinha, 2, 25)) & Trim(Mid(strLinha, 27, 4)) & Trim(Mid(strLinha, 45, 8)), 26) & "' ORDER BY Pkid Desc"
                   strSqlSub = gstrTOPnOracle(strSqlSub, 1)
                
                   'Vamos buscar os dados referentes às guias apenas com a parte especifica do codigo de barras
                   strSQL = "SELECT LA.intComposicaoDaReceita, LA.intExercicio, LV.dblValor ValorOrig, LV.dtmDtVencimento, LV.intMoeda, LV.Pkid intLancamentoValor, LV.intParcela, LG.intGuias, LV.DtmDtVencimento, LG.DBLVALORPRINCIPAL, LG.DBLVALORMULTA, LG.DBLVALORJUROS, LG.Dblvalorcorrecao, LG.DBLVALORDESCONTO, (LG.DBLVALORPRINCIPAL + LG.DBLVALORMULTA + LG.DBLVALORJUROS + LG.Dblvalorcorrecao - LG.DBLVALORDESCONTO) dblCorreto, " & _
                                      gstrCONVERT(CDT_numeric, "LA.STRNUMEROAVISO") & " strNumeroAviso," & _
                                      "(SELECT SUM(LG2.DBLVALORPRINCIPAL + LG2.DBLVALORMULTA + LG2.DBLVALORJUROS + LG2.Dblvalorcorrecao - LG2.DBLVALORDESCONTO) FROM " & gstrLancamentoGuias & " LG2 WHERE LG2.intguias = LG.intGuias) dblTotal " & _
                            "FROM " & gstrLancamentoGuias & " LG, " & _
                                      gstrLancamentoValor & " LV, " & _
                                      gstrLancamentoAlfa & " LA " & _
                            "WHERE LV.Pkid = LG.INTLANCAMENTOVALOR AND " & _
                            "      LA.Pkid = LV.INTLANCAMENTOALFA AND " & _
                                   "LG.intGuias = (" & strSqlSub & ")"
                                                       
                   If gobjBanco.CriaADO(strSQL, 30, adoResultado) Then
                       If adoResultado.EOF Then
                
                           strSQL = "Select "
                           strSQL = strSQL & "BAC.intPosInicial, "
                           strSQL = strSQL & "BAC.intPosFinal "
                           strSQL = strSQL & "from "
                           strSQL = strSQL & "tblBaixaAuxiliar BA, "
                           strSQL = strSQL & "tblBaixaAuxiliarCampos BAC "
                           strSQL = strSQL & "Where "
                           strSQL = strSQL & "BA.pkid = BAC.intBaixaAuxiliar and "
                           strSQL = strSQL & "BA.strBanco = '" & strBanco & "' and "
                           strSQL = strSQL & "BA.strContaBancaria = '" & strContaBancaria & "'"
                           strSQL = strSQL & " Order By BAC.Pkid "
                           
                           strCodEspecifico = ""
                           
                           If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
                               If Not adoResultado.EOF Then
                                   Do While Not adoResultado.EOF
                                       strCodEspecifico = strCodEspecifico & Mid(strLinha, gstrENulo(adoResultado!intPosInicial), gstrENulo(adoResultado!intPosFinal))
                                       adoResultado.MoveNext
                                   Loop
                               Else
                                   strCodEspecifico = Trim(Mid(strLinha, 2, 25)) & Trim(Mid(strLinha, 27, 4)) & Trim(Mid(strLinha, 45, 8))
                               End If
                           End If
        
                        
                           strSqlSub = ""
                           strSqlSub = "SELECT " & gstrTOPnSQLServer(1) & " Pkid FROM " & gstrGuias & " WHERE strCodBarraEsp = '" & strCodEspecifico & "' ORDER BY Pkid Desc"
                           strSqlSub = gstrTOPnOracle(strSqlSub, 1)
                        
                           'Vamos buscar os dados referentes às guias apenas com a parte especifica do codigo de barras
                           strSQL = "SELECT LA.intComposicaoDaReceita, LA.intExercicio, LV.dblValor ValorOrig, LV.dtmDtVencimento, LV.intMoeda, LV.Pkid intLancamentoValor, LV.intParcela, LG.intGuias, LV.DtmDtVencimento, LG.DBLVALORPRINCIPAL, LG.DBLVALORMULTA, LG.DBLVALORJUROS, LG.Dblvalorcorrecao, LG.DBLVALORDESCONTO, (LG.DBLVALORPRINCIPAL + LG.DBLVALORMULTA + LG.DBLVALORJUROS + LG.Dblvalorcorrecao - LG.DBLVALORDESCONTO) dblCorreto, " & _
                                              gstrCONVERT(CDT_numeric, "LA.STRNUMEROAVISO") & " strNumeroAviso," & _
                                              "(SELECT SUM(LG2.DBLVALORPRINCIPAL + LG2.DBLVALORMULTA + LG2.DBLVALORJUROS + LG2.Dblvalorcorrecao - LG2.DBLVALORDESCONTO) FROM " & gstrLancamentoGuias & " LG2 WHERE LG2.intguias = LG.intGuias) dblTotal " & _
                                    "FROM " & gstrLancamentoGuias & " LG, " & _
                                              gstrLancamentoValor & " LV, " & _
                                              gstrLancamentoAlfa & " LA " & _
                                    "WHERE LV.Pkid = LG.INTLANCAMENTOVALOR AND " & _
                                    "      LA.Pkid = LV.INTLANCAMENTOALFA AND " & _
                                           "LG.intGuias = (" & strSqlSub & ")"
                                                               
                           If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
                               If adoResultado.EOF Then
                        
                                    blnExisteGuia = False
                                    'ExibeMensagem "Não foi encontrada Conta referente à este arquivo. A operação não concluída."
                                    'Screen.MousePointer = vbDefault
                                    'pgr_Status.Visible = False
                                    'gobjBanco.ExecutaRollbackTrans
                                    'Close #1
                                    'Exit Sub
                               
                               End If
                           End If
                        End If
                    End If
                End If
                
                dblPorcentagemDif = 0
                dblSomaDosValores = 0
                dblValorDiferenca = 0
                dblValorDoArquivo = ConverteValorDoArquivo(Mid(strLinha, 53, 15))
                
                If blnExisteGuia Then 'Verifica se o código de barras vindo do arquivo existe no banco
                
                    'Vamos verificar se existe diferenca de valores
                    If dblValorDoArquivo <> adoResultado("dblTotal").Value Then
                        'dblPorcentagemDif = (dblValorDoArquivo / adoResultado("dblTotal").Value)
                        'PROVISORIO
                        dblPorcentagemDif = dblValorDoArquivo - adoResultado("dblTotal").Value
                    End If
                    
                    'Vamos gravar um registro em Movimento Bancario para cada Guia
                    Do While Not adoResultado.EOF
                        
                        'Caso exista diferenca dos valores, vamos tirar a diferenca proporcionalmente
                        If dblPorcentagemDif <> 0 Then
                            'dblValorPrincipal = gstrConvVrDoSql((adoResultado("dblvalorprincipal") * dblPorcentagemDif), 2)
                            'dblValorMulta = gstrConvVrDoSql((adoResultado("dblvalorMulta") * dblPorcentagemDif), 2)
                            'dblValorJuros = gstrConvVrDoSql((adoResultado("dblvalorJuros") * dblPorcentagemDif), 2)
                            'dblValorCorrecao = gstrConvVrDoSql((adoResultado("dblvalorCorrecao") * dblPorcentagemDif), 2)
                            'dblValorDesconto = gstrConvVrDoSql((adoResultado("dblvalorDesconto") * dblPorcentagemDif), 2)
                            'PROVISORIO
                            dblValorPrincipal = gstrConvVrDoSql((dblValorDoArquivo - dblPorcentagemDif), 2)
                            dblValorMulta = gstrConvVrDoSql(dblPorcentagemDif, 2)
                            dblValorJuros = 0
                            dblValorCorrecao = 0
                            dblValorDesconto = 0
                        Else
                            dblValorPrincipal = adoResultado("dblValorPrincipal").Value 'ConverteValorDoArquivo(Mid(strLinha, 53, 15))
                            dblValorMulta = adoResultado("dblValorMulta").Value
                            dblValorJuros = adoResultado("dblValorJuros").Value
                            dblValorCorrecao = adoResultado("dblValorCorrecao").Value
                            dblValorDesconto = adoResultado("dblValorDesconto").Value
                        End If
                        
                        dblSomaDosValores = dblSomaDosValores + (dblValorPrincipal + dblValorMulta + dblValorJuros + dblValorCorrecao + dblValorDesconto)
                        
                        'Se este for o ultimo registro, vamos verificar se ha diferenca dos valores calculados
                        If adoResultado.RecordCount = adoResultado.AbsolutePosition Then
                            'Caso depois do acerto de valores ainda exista diferenca, a jogaremos no ultimo registro
                            If CCur(dblSomaDosValores) <> CCur(dblValorDoArquivo) Then
                            
                                dblValorDiferenca = dblValorDoArquivo - dblSomaDosValores
                        
                                'strSQL = "UPDATE " & gstrMovimentoBancario
                                If dblValorPrincipal > 0 Then
                                    dblValorPrincipal = dblValorPrincipal + dblValorDiferenca
                                ElseIf dblValorMulta > 0 Then
                                    dblValorMulta = dblValorMulta + dblValorDiferenca
                                ElseIf dblValorJuros > 0 Then
                                    dblValorJuros = dblValorJuros + dblValorDiferenca
                                Else
                                    dblValorCorrecao = dblValorCorrecao + dblValorDiferenca
                                End If
                            End If
                        End If
                        
                        strSQL = gstrStoredProcedure("sp_AtualizaParcela", adoResultado("intComposicaoDaReceita").Value & ", " & adoResultado("intExercicio").Value & ", " & adoResultado("intLancamentoValor").Value & ", " & gstrConvDtParaSql(adoResultado("dtmDtVencimento").Value) & ", " & gstrConvDtParaSql(ConverteDataDoArquivo(Mid(strLinha, 45, 8), True)) & ", " & gstrConvVrParaSql(adoResultado("ValorOrig").Value) & ", " & adoResultado("intMoeda").Value, True)
                        If gobjBanco.CriaADO(strSQL, 80, adoAtualizacao) Then
                            If gstrConvVrDoSql(adoResultado("ValorOrig").Value, , , True) = 0 Then
                                dblValorPrincipal = dblValorDoArquivo
                                dblValorCorreto = dblValorDoArquivo
                                dblValorMulta = 0
                            Else
                                dblValorCorreto = Space$(0) & gstrConvVrDoSql(Val(gstrConvVrParaSql(adoAtualizacao("dblValorPrincipal").Value)) + Val(gstrConvVrParaSql(adoAtualizacao("dblValorMulta").Value)) + Val(gstrConvVrParaSql(adoAtualizacao("dblValorJuros").Value)) + Val(gstrConvVrParaSql(adoAtualizacao("dblValorCorrecao").Value)))
                            End If
                        Else
                            Screen.MousePointer = vbDefault
                            pgr_Status.Visible = False
                            lbl_Status.Visible = False
                            gobjBanco.ExecutaRollbackTrans
                            Close #1
                            Exit Sub

                        End If

                            gobjBanco.Execute "INSERT INTO " & gstrMovimentoBancario & "(dtmDtMovimento, intContaBancaria, intLote, dtmDtPagamento, dblPrincipal, dblMulta, dblJuros, dblCorrecao, strCodigoDeBarras, dblCorreto, intLancamentoValor, intGuias, intCodigoBaixa, intDigito, BytTipo, dtmDtAtualizacao, lngCodUsr, bitGuia, bitProcessado) " & _
                                                                                        "VALUES(" & gstrConvDtParaSql(txt_DtMovimento.Text) & "," & lngPkidContaBancaria & "," & strLote & "," & gstrConvDtParaSql(ConverteDataDoArquivo(Mid(strLinha, 45, 6), True)) & "," & gstrConvVrParaSql(dblValorPrincipal) & "," & gstrConvVrParaSql(dblValorMulta) & "," & gstrConvVrParaSql(dblValorJuros) & "," & gstrConvVrParaSql(dblValorCorrecao) & ",'" & Trim(Mid(strLinha, 2, 25)) & Trim(Mid(strLinha, 27, 4)) & Trim(Mid(strLinha, 45, 8)) & "'," & gstrConvVrParaSql(dblValorCorreto) & "," & adoResultado("intLancamentoValor").Value & "," & adoResultado("intGuias").Value & "," & _
                                                                                        IIf(CDate(adoResultado("dtmDtVencimento").Value) < ConverteDataDoArquivo(Mid(strLinha, 45, 8), True), intCodBaixaAposVcto, intCodBaixaAntesVcto) & "," & gstrCalculaDigitoModulo10(Trim(adoResultado("strNumeroAviso").Value) & Format$(Trim(adoResultado("intParcela").Value), "000")) & _
                                                                                        "," & IIf(Mid(strLinha, 117, 1) = " ", "9", Mid(strLinha, 117, 1)) & "," & gstrConvDtParaSql(gstrDataDoSistema) & "," & glngCodUsr & ",1,0)"
                        adoResultado.MoveNext
                        
                    Loop
                    
                Else
                    
                   dblValorPrincipal = dblValorDoArquivo
                   dblValorMulta = 0
                   dblValorJuros = 0
                   dblValorCorrecao = 0
                   dblValorDesconto = 0
                    
                   gobjBanco.Execute "INSERT INTO " & gstrMovimentoBancario & "(dtmDtMovimento, intContaBancaria, intLote, dtmDtPagamento, dblPrincipal, dblMulta, dblJuros, dblCorrecao, strCodigoDeBarras, dblCorreto, intLancamentoValor, intGuias, intCodigoBaixa, intDigito, BytTipo, dtmDtAtualizacao, lngCodUsr, bitGuia, bitProcessado) " & _
                                                                                    "VALUES(" & gstrConvDtParaSql(txt_DtMovimento.Text) & "," & lngPkidContaBancaria & "," & strLote & "," & gstrConvDtParaSql(ConverteDataDoArquivo(Mid(strLinha, 45, 8), True)) & "," & gstrConvVrParaSql(dblValorPrincipal) & "," & gstrConvVrParaSql(dblValorMulta) & "," & gstrConvVrParaSql(dblValorJuros) & "," & gstrConvVrParaSql(dblValorCorrecao) & ",'" & Trim(Mid(strLinha, 2, 25)) & Trim(Mid(strLinha, 27, 4)) & Trim(Mid(strLinha, 45, 8)) & "',0,NULL,NULL,NULL,0," & _
                                                                                    IIf(Mid(strLinha, 117, 1) = " ", "9", Mid(strLinha, 117, 1)) & "," & gstrConvDtParaSql(gstrDataDoSistema) & "," & glngCodUsr & ",0,0)"
                
                End If
                
            End If
             
        ElseIf Mid(strLinha, 1, 1) = "Z" Then
            bytQtdeTipos = bytQtdeTipos + 1
                         
            'Vamos verificar se ja existe um registro referente, caso exista vamos somar o valor
            strSQL = "SELECT Pkid, dblValor FROM " & gstrResumoBancario & " WHERE dtmData = " & gstrConvDtParaSql(txt_DtMovimento.Text) & " AND intLote = " & strLote & " AND intContaBancaria = " & lngPkidContaBancaria
            If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
                If Not adoResultado.EOF Then
             
                    'Vamos somar registro na tabela tblResumoBancario
                    strSQL = "UPDATE " & gstrResumoBancario & " SET dblValor = " & gstrConvVrParaSql(ConverteValorDoArquivo(Mid(strLinha, 8, 17)) + adoResultado("dblValor").Value) & " WHERE Pkid = " & adoResultado("Pkid").Value
                     
                Else
                
                    'Vamos gravar registro na tabela tblResumoBancario
                    strSQL = "INSERT INTO " & gstrResumoBancario & "(dtmData, intContaBancaria, intLote, dblValor, dtmDtAtualizacao, lngCodUsr) " & _
                             "VALUES(" & gstrConvDtParaSql(txt_DtMovimento.Text) & "," & lngPkidContaBancaria & "," & strLote & "," & gstrConvVrParaSql(ConverteValorDoArquivo(Mid(strLinha, 8, 17))) & "," & gstrConvDtParaSql(gstrDataDoSistema) & "," & glngCodUsr & ")"
                
                End If
            End If
            
            gobjBanco.Execute strSQL
            
            'Vamos verificar se a qtde gravada é igual ao total de registros
            If lngLinha + 1 <> Val(Mid(strLinha, 2, 6)) Then
                ExibeMensagem "A quantidade de registros processada não coincide com a informada no arquivo. A operação não concluída."
                Screen.MousePointer = vbDefault
                pgr_Status.Visible = False
                lbl_Status.Visible = False
                gobjBanco.ExecutaRollbackTrans
                Close #1
                Exit Sub
            End If
            
        End If
        
        lngLinha = lngLinha + 1
        
        DoEvents
        lbl_Status.Caption = lngLinha & " de " & lngSize
        Me.Refresh
        
        pgr_Status.Value = lngLinha
        
ProximaLinha:

    Loop
        
    Close #1
    
    pgr_Status.Visible = False
    lbl_Status.Visible = False
    
    'Vamos verificar se foram lidos todos os tipos de registro
    If bytQtdeTipos < 4 Then
        ExibeMensagem "O arquivo de leitura está incorreto."
        Screen.MousePointer = vbDefault
        gobjBanco.ExecutaRollbackTrans
        Close #1
        Exit Sub
    End If
    
    gobjBanco.ExecutaCommitTrans
    
    Screen.MousePointer = vbDefault
    
    Exit Sub

err_BaixaAutomatica:
    ExibeMensagem "Ocorreu um Erro na leitura do arquivo. A operação não foi concluída."
    Screen.MousePointer = vbDefault
    pgr_Status.Visible = False
    lbl_Status.Visible = False
    gobjBanco.ExecutaRollbackTrans
    Close #1
End Sub

Private Sub LeMovimentoBancarioFichaCompensacaoBanespa()
Dim strLinha               As String
Dim lngLinha               As Long
Dim lngSize                As Long

Dim blnTipoJaSomado        As Boolean
Dim strSQL                 As String
Dim strSqlSub              As String

Dim adoResultado           As New ADODB.Recordset
Dim adoAtualizacao         As New ADODB.Recordset

Dim bytQtdeTipos           As Byte ' Tipos (0,1,9)
Dim dblPorcentagemDif      As Double
Dim dblValorDiferenca      As Double
Dim dblValorDoArquivo      As Double
Dim dblValorTarifa         As Double
Dim dblTotal               As Double

Dim dblSomaDosValores      As Double

Dim dblValorPrincipal      As Double
Dim dblValorMulta          As Double
Dim dblValorJuros          As Double
Dim dblValorCorrecao       As Double
Dim dblValorDesconto       As Double
Dim dblValorCorreto        As Double

Dim lngPkidContaBancaria   As Long
Dim lngPkidContaBancariaMov As Long
Dim blnExisteGuia          As Boolean
Dim blnLayoutNovo          As Boolean
Dim intLote                As Integer

Dim strCaseField           As String
Dim xadbGuiasDuplicadas    As New XArrayDB

    On Error GoTo err_BaixaAutomatica
    
    bytQtdeTipos = 0
    lngLinha = 0
    blnTipoJaSomado = False
    
    Open txt_Arquivo For Input As #1
    
    lngSize = FileLen(txt_Arquivo) / 120
    
    pgr_Status.Value = 0
    pgr_Status.Visible = True
    pgr_Status.Max = Abs(lngSize)
    lbl_Status.Visible = True
    
    Set gobjBanco = New clsBanco
    
    gobjBanco.ExecutaBeginTrans
    
    Do While Not EOF(1)

        Line Input #1, strLinha

        If Len(strLinha) = 0 Then
            GoTo ProximaLinha
        End If

        'Vamos verificar se a linha contem 150 posicoes
        If Len(strLinha) <> 120 Then
            ExibeMensagem "Ocorreu um Erro na leitura do arquivo. A operação não foi concluída."
            Screen.MousePointer = vbDefault
            pgr_Status.Visible = False
            lbl_Status.Visible = False
            gobjBanco.ExecutaRollbackTrans
            Close #1
            Exit Sub
        End If

        If Mid(strLinha, 1, 1) = "0" Then
            bytQtdeTipos = bytQtdeTipos + 1

            'Vamos buscar o Pkid referente à Conta em tblContaBancaria
            strSQL = "SELECT Pkid FROM " & gstrContaBancaria & " WHERE strContaRetorno = '" & Trim(Mid(strLinha, 20, 11)) & "'"
            If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
                If adoResultado.EOF Then
                    ExibeMensagem "Não foi encontrada Conta referente à este arquivo. A operação não concluída."
                    Screen.MousePointer = vbDefault
                    pgr_Status.Visible = False
                    lbl_Status.Visible = False
                    gobjBanco.ExecutaRollbackTrans
                    Close #1
                    Exit Sub
                Else
                    lngPkidContaBancaria = adoResultado("Pkid").Value
                End If
            End If
                        
            'CONSULTA ESPECIFICA PARA DUPLICATAS NO GUARUJA (INICIO) ******************************************************
            'Vamos buscar o Pkid referente à Conta em tblContaBancaria
            strSQL = "SELECT Pkid FROM " & gstrContaBancaria & " WHERE strConta = '" & Trim(Mid(strLinha, 23, 7)) & "'"
            If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
                If adoResultado.EOF Then
                    ExibeMensagem "Não foi encontrada Conta do movimento. A operação não concluída."
                    Screen.MousePointer = vbDefault
                    pgr_Status.Visible = False
                    lbl_Status.Visible = False
                    gobjBanco.ExecutaRollbackTrans
                    Close #1
                    Exit Sub
                Else
                    lngPkidContaBancariaMov = adoResultado("Pkid").Value
                End If
            End If
            'CONSULTA ESPECIFICA PARA DUPLICATAS NO GUARUJA (FIM) ******************************************************
            
            'Vamos atribuir o numero do lote
            'Caso nas posicoes 73 a 77 estiver zerado vamos obter o lote, caso contrario pegaremos o enviado no arquivo
            If Val(Mid(strLinha, 73, 4)) = 0 Then
                strSQL = "SELECT Max(intLote) UltimoLote FROM " & gstrMovimentoBancario & " WHERE intContaBancaria = " & lngPkidContaBancaria & " AND bytTipo <> 0"
                If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
                    intLote = IIf(IsNull(adoResultado("UltimoLote").Value), 1, adoResultado("UltimoLote").Value + 1)
                End If
            Else
                intLote = Val(Mid(strLinha, 73, 4))
            End If
            
        ElseIf Mid(strLinha, 1, 1) = "1" Then
            
            blnLayoutNovo = False 'Len(Trim(Mid(strLinha, 105, 13))) > 0
            
            If Not blnTipoJaSomado Then
                bytQtdeTipos = bytQtdeTipos + 1
                blnTipoJaSomado = True
            End If
            
            blnExisteGuia = True
                         
            'CONSULTA ESPECIFICA PARA DUPLICATAS NO GUARUJA E CAMPOS DO JORDAO (INICIO) ******************************************************
            'Vamos verificar se existe duplicata
            strCaseField = gstrCONVERT(CDT_INT, "Right(" & gstrCONVERT(CDT_VARCHAR, "G.intNumero") & ",7)")
            strSQL = "SELECT G.Pkid, G.dblValor " & _
                     "FROM " & gstrGuias & " G " & strREADPAST & " " & _
                     "WHERE " & gstrCASEWHEN("G.intNumero > 990000000 ", "*strCaseField*", "G.intNumero") & " = " & Mid(strLinha, 16, 7) & " AND " & _
                     "      G.intContaBancaria = " & lngPkidContaBancariaMov
            
            'Colocada condicao para CJD, pois existe duplicidade de CodBarraEsp mas com numero de Guias com 99 para lancamentos 2005
            strSQL = Replace(strSQL, "*strCaseField*", strCaseField)
            
            If gobjBanco.CriaADO(strSQL, 30, adoResultado) Then
                If adoResultado.RecordCount > 1 Then
                    
                    Set xadbGuiasDuplicadas = New XArrayDB
                    xadbGuiasDuplicadas.Clear
                    xadbGuiasDuplicadas.ReDim 0, 0, 0, 1
                    
                    adoResultado.MoveFirst
                    'Vamos armazenar as guias com os valores
                    Do While Not adoResultado.EOF
                        If Not adoResultado.AbsolutePosition = 1 Then
                            xadbGuiasDuplicadas.ReDim 0, xadbGuiasDuplicadas.UpperBound(1) + 1, 0, 1
                        End If
                        xadbGuiasDuplicadas(xadbGuiasDuplicadas.UpperBound(1), 0) = adoResultado("Pkid").Value
                        xadbGuiasDuplicadas(xadbGuiasDuplicadas.UpperBound(1), 1) = Abs(adoResultado("dblValor").Value - (ConverteValorDoArquivo(IIf(blnLayoutNovo, Mid(strLinha, 76, 13), Mid(strLinha, 76, 13))) + ConverteValorDoArquivo(IIf(blnLayoutNovo, Mid(strLinha, 37, 13), Mid(strLinha, 37, 13)))))
                        adoResultado.MoveNext
                    Loop
                    
                    'Vamos procurar encontrar com o valor da guia inteiro
                    strSQL = "SELECT LA.intComposicaoDaReceita, LA.intExercicio, LV.dblValor ValorOrig, LV.dtmDtVencimento, LV.intMoeda, LV.Pkid intLancamentoValor, LV.intParcela, LG.intGuias, LV.DtmDtVencimento, LG.DBLVALORPRINCIPAL, LG.DBLVALORMULTA, LG.DBLVALORJUROS, LG.Dblvalorcorrecao, LG.DBLVALORDESCONTO, (LG.DBLVALORPRINCIPAL + LG.DBLVALORMULTA + LG.DBLVALORJUROS + LG.Dblvalorcorrecao - LG.DBLVALORDESCONTO) dblCorreto, G.strCodBarra, " & _
                                gstrCONVERT(CDT_numeric, "LA.STRNUMEROAVISO") & " strNumeroAviso, " & _
                                "(SELECT SUM(LG2.DBLVALORPRINCIPAL + LG2.DBLVALORMULTA + LG2.DBLVALORJUROS + LG2.Dblvalorcorrecao - LG2.DBLVALORDESCONTO) FROM " & gstrLancamentoGuias & " LG2 WHERE LG2.intguias = LG.intGuias) dblTotal " & _
                             "FROM " & gstrLancamentoGuias & " LG, " & _
                                 gstrGuias & " G, " & _
                                 gstrLancamentoValor & " LV, " & _
                                 gstrLancamentoAlfa & " LA " & _
                            "WHERE LV.Pkid = LG.INTLANCAMENTOVALOR AND " & _
                            "      LG.intGuias = G.Pkid AND " & _
                            "      LA.Pkid = LV.INTLANCAMENTOALFA AND " & _
                                "LG.intGuias = (SELECT Min(Pkid) FROM tblGuias WHERE intNumero =  " & Mid(strLinha, 16, 7) & " And intContaBancaria = " & lngPkidContaBancariaMov & " And " & gstrCONVERT(CDT_INT, "dblValor") & " = " & Val(ConverteValorDoArquivo(IIf(blnLayoutNovo, Mid(strLinha, 76, 13), Mid(strLinha, 76, 13))) + ConverteValorDoArquivo(IIf(blnLayoutNovo, Mid(strLinha, 37, 13), Mid(strLinha, 37, 13)))) & ")"
                     
                    If gobjBanco.CriaADO(strSQL, 30, adoResultado) Then
                        If adoResultado.EOF Then
                            'Vamos colocar no index 0 o a guia com valor mais proximo
                            xadbGuiasDuplicadas.QuickSort 0, xadbGuiasDuplicadas.UpperBound(1), 1, XORDER_ASCEND, XTYPE_INTEGER
                            
                            'Vamos obter a guia com valor mais proximo
                            strSQL = "SELECT LA.intComposicaoDaReceita, LA.intExercicio, LV.dblValor ValorOrig, LV.dtmDtVencimento, LV.intMoeda, LV.Pkid intLancamentoValor, LV.intParcela, LG.intGuias, LV.DtmDtVencimento, LG.DBLVALORPRINCIPAL, LG.DBLVALORMULTA, LG.DBLVALORJUROS, LG.Dblvalorcorrecao, LG.DBLVALORDESCONTO, (LG.DBLVALORPRINCIPAL + LG.DBLVALORMULTA + LG.DBLVALORJUROS + LG.Dblvalorcorrecao - LG.DBLVALORDESCONTO) dblCorreto, G.strCodBarra, " & _
                                        gstrCONVERT(CDT_numeric, "LA.STRNUMEROAVISO") & " strNumeroAviso, " & _
                                        "(SELECT SUM(LG2.DBLVALORPRINCIPAL + LG2.DBLVALORMULTA + LG2.DBLVALORJUROS + LG2.Dblvalorcorrecao - LG2.DBLVALORDESCONTO) FROM " & gstrLancamentoGuias & " LG2 WHERE LG2.intguias = LG.intGuias) dblTotal " & _
                                     "FROM " & gstrLancamentoGuias & " LG, " & _
                                         gstrGuias & " G, " & _
                                         gstrLancamentoValor & " LV, " & _
                                         gstrLancamentoAlfa & " LA " & _
                                    "WHERE LV.Pkid = LG.INTLANCAMENTOVALOR AND " & _
                                    "      LG.intGuias = G.Pkid AND " & _
                                    "      LA.Pkid = LV.INTLANCAMENTOALFA AND " & _
                                        "LG.intGuias = " & xadbGuiasDuplicadas(0, 0)
                             
                            If gobjBanco.CriaADO(strSQL, 30, adoResultado) Then
                                If adoResultado.EOF Then
                                    blnExisteGuia = False
                                Else
                                    GoTo PularProcedimentoDeBusca
                                End If
                            End If
                        Else
                            GoTo PularProcedimentoDeBusca
                        End If
                    End If
                    
                End If
            End If
            'CONSULTA ESPECIFICA PARA DUPLICATAS NO GUARUJA E CAMPOS DO JORDAO (FIM) ******************************************************
            
'            strSqlSub = ""
'            strSqlSub = "SELECT " & gstrTOPnSQLServer(1) & " Pkid FROM " & gstrGuias & " WHERE strCodBarra = '" & IIf(blnLayoutNovo, Mid(strLinha, 2, 21), Mid(strLinha, 2, 21)) & "' ORDER BY Pkid Desc"
'            strSqlSub = gstrTOPnOracle(strSqlSub, 1)

            'Vamos buscar os dados referentes às guias
            strSQL = "SELECT LA.intComposicaoDaReceita, LA.intExercicio, LV.dblValor ValorOrig, LV.dtmDtVencimento, LV.intMoeda, LV.Pkid intLancamentoValor, LV.intParcela, LG.intGuias, LV.DtmDtVencimento, LG.DBLVALORPRINCIPAL, LG.DBLVALORMULTA, LG.DBLVALORJUROS, LG.Dblvalorcorrecao, LG.DBLVALORDESCONTO, (LG.DBLVALORPRINCIPAL + LG.DBLVALORMULTA + LG.DBLVALORJUROS + LG.Dblvalorcorrecao - LG.DBLVALORDESCONTO) dblCorreto, G.strCodBarra, " & _
                            gstrCONVERT(CDT_numeric, "LA.STRNUMEROAVISO") & " strNumeroAviso, " & _
                            "(SELECT SUM(LG2.DBLVALORPRINCIPAL + LG2.DBLVALORMULTA + LG2.DBLVALORJUROS + LG2.Dblvalorcorrecao - LG2.DBLVALORDESCONTO) FROM " & gstrLancamentoGuias & " LG2 WHERE LG2.intguias = LG.intGuias) dblTotal " & _
                     "FROM " & gstrLancamentoGuias & " LG, " & _
                             gstrGuias & " G, " & _
                             gstrLancamentoValor & " LV, " & _
                             gstrLancamentoAlfa & " LA " & _
                     "WHERE LV.Pkid = LG.INTLANCAMENTOVALOR AND " & _
                     "      LA.Pkid = LV.INTLANCAMENTOALFA AND " & _
                     "      LG.intGuias = G.Pkid AND " & _
                           "LG.intGuias = (SELECT Min(Pkid) FROM tblGuias WHERE strCodBarra =  '" & IIf(blnLayoutNovo, Mid(strLinha, 2, 21), Mid(strLinha, 2, 21)) & "')"
                     
            If gobjBanco.CriaADO(strSQL, 30, adoResultado) Then
                If adoResultado.EOF Then
                 
'                    strSqlSub = ""
'                    strSqlSub = "SELECT " & gstrTOPnSQLServer(1) & " Pkid FROM " & gstrGuias & " WHERE strCodBarraEsp = '" & IIf(blnLayoutNovo, Mid(strLinha, 2, 21), Mid(strLinha, 2, 21)) & "' ORDER BY Pkid Desc"
'                    strSqlSub = gstrTOPnOracle(strSqlSub, 1)

                    'Vamos buscar os dados referentes às guias apenas com a parte especifica do codigo de barras
                    strSQL = "SELECT LA.intComposicaoDaReceita, LA.intExercicio, LV.dblValor ValorOrig, LV.dtmDtVencimento, LV.intMoeda, LV.Pkid intLancamentoValor, LV.intParcela, LG.intGuias, LV.DtmDtVencimento, LG.DBLVALORPRINCIPAL, LG.DBLVALORMULTA, LG.DBLVALORJUROS, LG.Dblvalorcorrecao, LG.DBLVALORDESCONTO, (LG.DBLVALORPRINCIPAL + LG.DBLVALORMULTA + LG.DBLVALORJUROS + LG.Dblvalorcorrecao - LG.DBLVALORDESCONTO) dblCorreto, G.strCodBarra, " & _
                                       gstrCONVERT(CDT_numeric, "LA.STRNUMEROAVISO") & " strNumeroAviso, " & _
                                       "(SELECT SUM(LG2.DBLVALORPRINCIPAL + LG2.DBLVALORMULTA + LG2.DBLVALORJUROS + LG2.Dblvalorcorrecao - LG2.DBLVALORDESCONTO) FROM " & gstrLancamentoGuias & " LG2 WHERE LG2.intguias = LG.intGuias) dblTotal " & _
                             "FROM " & gstrLancamentoGuias & " LG, " & _
                                       gstrGuias & " G, " & _
                                       gstrLancamentoValor & " LV, " & _
                                       gstrLancamentoAlfa & " LA " & _
                             "WHERE LV.Pkid = LG.INTLANCAMENTOVALOR AND " & _
                             "      LA.Pkid = LV.INTLANCAMENTOALFA AND " & _
                             "      LG.intGuias = G.Pkid AND " & _
                                    "LG.intGuias = (SELECT Min(Pkid) FROM tblGuias WHERE strCodBarraEsp =  '" & IIf(blnLayoutNovo, Mid(strLinha, 2, 21), Mid(strLinha, 2, 21)) & "')"
                                    
                    If gobjBanco.CriaADO(strSQL, 30, adoResultado) Then
                        If adoResultado.EOF Then

                            blnExisteGuia = False
                            'ExibeMensagem "Não foi encontrada Conta referente à este arquivo. A operação não concluída."
                            'Screen.MousePointer = vbDefault
                            'pgr_Status.Visible = False
                            'gobjBanco.ExecutaRollbackTrans
                            'Close #1
                            'Exit Sub
                        End If
                    End If
                End If

PularProcedimentoDeBusca:

                dblPorcentagemDif = 0
                dblSomaDosValores = 0
                dblValorDiferenca = 0
                dblValorDoArquivo = ConverteValorDoArquivo(IIf(blnLayoutNovo, Mid(strLinha, 76, 13), Mid(strLinha, 76, 13)))
                dblValorTarifa = ConverteValorDoArquivo(IIf(blnLayoutNovo, Mid(strLinha, 37, 13), Mid(strLinha, 37, 13)))
                If DesContarTarifa = True Then
                    dblTotal = dblTotal + dblValorDoArquivo
                Else
                    dblValorDoArquivo = dblValorDoArquivo + dblValorTarifa
                    dblTotal = dblTotal + dblValorDoArquivo
                End If
                
                
                If blnExisteGuia Then 'Verifica se o código de barras vindo do arquivo existe no banco

                    'Vamos verificar se existe diferenca de valores
                    If dblValorDoArquivo <> adoResultado("dblTotal").Value Then
                        'dblPorcentagemDif = (dblValorDoArquivo / adoResultado("dblTotal").Value)
                        'PROVISORIO
                        dblPorcentagemDif = dblValorDoArquivo - adoResultado("dblTotal").Value
                    End If

                    'Vamos gravar um registro em Movimento Bancario para cada Guia
                    Do While Not adoResultado.EOF

                        'Caso exista diferenca dos valores, vamos tirar a diferenca proporcionalmente
                        If dblPorcentagemDif <> 0 Then
                            'dblValorPrincipal = gstrConvVrDoSql((adoResultado("dblvalorprincipal") * dblPorcentagemDif), 2)
                            'dblValorMulta = gstrConvVrDoSql((adoResultado("dblvalorMulta") * dblPorcentagemDif), 2)
                            'dblValorJuros = gstrConvVrDoSql((adoResultado("dblvalorJuros") * dblPorcentagemDif), 2)
                            'dblValorCorrecao = gstrConvVrDoSql((adoResultado("dblvalorCorrecao") * dblPorcentagemDif), 2)
                            'dblValorDesconto = gstrConvVrDoSql((adoResultado("dblvalorDesconto") * dblPorcentagemDif), 2)
                            'PROVISORIO
                            If dblPorcentagemDif > 0 Then
                                dblValorPrincipal = gstrConvVrDoSql((dblValorDoArquivo - dblPorcentagemDif), 2)
                                dblValorMulta = gstrConvVrDoSql(dblPorcentagemDif, 2)
                            Else
                                dblValorPrincipal = gstrConvVrDoSql((dblValorDoArquivo + dblPorcentagemDif), 2)
                                dblValorMulta = 0
                            End If
                            dblValorJuros = 0
                            dblValorCorrecao = 0
                            dblValorDesconto = 0
                        Else
                            dblValorPrincipal = adoResultado("dblValorPrincipal").Value 'ConverteValorDoArquivo(IIf(blnLayoutNovo, Mid(strLinha, 76, 13), Mid(strLinha, 76, 13)))
                            dblValorMulta = adoResultado("dblValorMulta").Value
                            dblValorJuros = adoResultado("dblValorJuros").Value
                            dblValorCorrecao = adoResultado("dblValorCorrecao").Value
                            dblValorDesconto = adoResultado("dblValorDesconto").Value
                        End If

                        dblSomaDosValores = dblSomaDosValores + (dblValorPrincipal + dblValorMulta + dblValorJuros + dblValorCorrecao + dblValorDesconto)

                        'Se este for o ultimo registro, vamos verificar se ha diferenca dos valores calculados
                        If adoResultado.RecordCount = adoResultado.AbsolutePosition Then
                            'Caso depois do acerto de valores ainda exista diferenca, a jogaremos no ultimo registro
                            If CCur(dblSomaDosValores) <> CCur(dblValorDoArquivo) Then

                                dblValorDiferenca = dblValorDoArquivo - dblSomaDosValores

                                'strSQL = "UPDATE " & gstrMovimentoBancario
                                If dblValorPrincipal > 0 Then
                                    dblValorPrincipal = dblValorPrincipal + dblValorDiferenca
                                ElseIf dblValorMulta > 0 Then
                                    dblValorMulta = dblValorMulta + dblValorDiferenca
                                ElseIf dblValorJuros > 0 Then
                                    dblValorJuros = dblValorJuros + dblValorDiferenca
                                Else
                                    dblValorCorrecao = dblValorCorrecao + dblValorDiferenca
                                End If
                            End If
                        End If

                        strSQL = gstrStoredProcedure("sp_AtualizaParcela", adoResultado("intComposicaoDaReceita").Value & ", " & adoResultado("intExercicio").Value & ", " & adoResultado("intLancamentoValor").Value & ", " & gstrConvDtParaSql(adoResultado("dtmDtVencimento").Value) & ", " & gstrConvDtParaSql(Mid(strLinha, 26, 2) & "/" & Mid(strLinha, 28, 2) & "/" & Mid(strLinha, 30, 2)) & ", " & gstrConvVrParaSql(adoResultado("ValorOrig").Value) & ", " & adoResultado("intMoeda").Value, True)
                        If gobjBanco.CriaADO(strSQL, 80, adoAtualizacao) Then
                            If gstrConvVrDoSql(adoResultado("ValorOrig").Value, , , True) = 0 Then
                                dblValorPrincipal = dblValorDoArquivo
                                dblValorCorreto = dblValorDoArquivo
                                dblValorMulta = 0
                            Else
                                dblValorCorreto = Space$(0) & gstrConvVrDoSql(Val(gstrConvVrParaSql(adoAtualizacao("dblValorPrincipal").Value)) + Val(gstrConvVrParaSql(adoAtualizacao("dblValorMulta").Value)) + Val(gstrConvVrParaSql(adoAtualizacao("dblValorJuros").Value)) + Val(gstrConvVrParaSql(adoAtualizacao("dblValorCorrecao").Value)))
                            End If
                        Else
                            Screen.MousePointer = vbDefault
                            pgr_Status.Visible = False
                            lbl_Status.Visible = False
                            gobjBanco.ExecutaRollbackTrans
                            Close #1
                            Exit Sub
                        End If

                        '*****************************************************
                        ' Data          : 20/12/2005                         *
                        ' Criação       : Verificação de Desconto de Tarifa  *
                        ' Responsável   : Fernando Peixoto                   *
                        ' Pendência     : Tri0518                            *
                        ''****************************************************

                        If DesContarTarifa = False Then
                            gobjBanco.Execute "INSERT INTO " & gstrMovimentoBancario & "(dtmDtMovimento, intContaBancaria, intLote, dtmDtPagamento, dblPrincipal, dblMulta, dblJuros, dblCorrecao, dblTarifa, strCodigoDeBarras, dblCorreto, intLancamentoValor, intGuias, intCodigoBaixa, intDigito, BytTipo, dtmDtAtualizacao, lngCodUsr, bitGuia, bitProcessado) " & _
                                                                                        "VALUES(" & gstrConvDtParaSql(txt_DtMovimento.Text) & "," & lngPkidContaBancaria & "," & intLote & "," & gstrConvDtParaSql(Mid(strLinha, 26, 2) & "/" & Mid(strLinha, 28, 2) & "/" & Mid(strLinha, 30, 2)) & "," & gstrConvVrParaSql(dblValorPrincipal) & "," & gstrConvVrParaSql(dblValorMulta) & "," & gstrConvVrParaSql(dblValorJuros) & "," & gstrConvVrParaSql(dblValorCorrecao) & "," & 0 & ",'" & adoResultado("strCodBarra").Value & "'," & gstrConvVrParaSql(dblValorCorreto) & "," & adoResultado("intLancamentoValor").Value & "," & adoResultado("intGuias").Value & "," & _
                                                                                        IIf(CDate(adoResultado("dtmDtVencimento").Value) < CDate(Mid(strLinha, 26, 2) & "/" & Mid(strLinha, 28, 2) & "/" & Mid(strLinha, 30, 2)), gstrENulo(intCodBaixaAposVcto, , True), gstrENulo(intCodBaixaAntesVcto, , True)) & "," & gstrCalculaDigitoModulo10(Trim(adoResultado("strNumeroAviso").Value) & Format$(Trim(adoResultado("intParcela").Value), "000")) & _
                                                                                        ",4" & "," & gstrConvDtParaSql(gstrDataDoSistema) & "," & glngCodUsr & ",1,0)"

                        Else
                            gobjBanco.Execute "INSERT INTO " & gstrMovimentoBancario & "(dtmDtMovimento, intContaBancaria, intLote, dtmDtPagamento, dblPrincipal, dblMulta, dblJuros, dblCorrecao, dblTarifa, strCodigoDeBarras, dblCorreto, intLancamentoValor, intGuias, intCodigoBaixa, intDigito, BytTipo, dtmDtAtualizacao, lngCodUsr, bitGuia, bitProcessado) " & _
                                                                                        "VALUES(" & gstrConvDtParaSql(txt_DtMovimento.Text) & "," & lngPkidContaBancaria & "," & intLote & "," & gstrConvDtParaSql(Mid(strLinha, 26, 2) & "/" & Mid(strLinha, 28, 2) & "/" & Mid(strLinha, 30, 2)) & "," & gstrConvVrParaSql(dblValorPrincipal - dblValorTarifa) & "," & gstrConvVrParaSql(dblValorMulta) & "," & gstrConvVrParaSql(dblValorJuros) & "," & gstrConvVrParaSql(dblValorCorrecao) & "," & gstrConvVrParaSql(dblValorTarifa) & ",'" & adoResultado("strCodBarra").Value & "'," & gstrConvVrParaSql(dblValorCorreto) & "," & adoResultado("intLancamentoValor").Value & "," & adoResultado("intGuias").Value & "," & _
                                                                                        IIf(CDate(adoResultado("dtmDtVencimento").Value) < CDate(Mid(strLinha, 26, 2) & "/" & Mid(strLinha, 28, 2) & "/" & Mid(strLinha, 30, 2)), gstrENulo(intCodBaixaAposVcto, , True), gstrENulo(intCodBaixaAntesVcto, , True)) & "," & gstrCalculaDigitoModulo10(Trim(adoResultado("strNumeroAviso").Value) & Format$(Trim(adoResultado("intParcela").Value), "000")) & _
                                                                                        ",4" & "," & gstrConvDtParaSql(gstrDataDoSistema) & "," & glngCodUsr & ",1,0)"
                        End If

                        adoResultado.MoveNext

                    Loop

                Else

                   dblValorPrincipal = dblValorDoArquivo
                   dblValorMulta = 0
                   dblValorJuros = 0
                   dblValorCorrecao = 0
                   dblValorDesconto = 0
                    '*****************************************************
                    ' Data          : 20/12/2005                         *
                    ' Criação       : Verificação de Desconto de Tarifa  *
                    ' Responsável   : Fernando Peixoto                   *
                    ' Pendência     : Tri0518                            *
                    ''****************************************************

                    If DesContarTarifa = False Then
                       gobjBanco.Execute "INSERT INTO " & gstrMovimentoBancario & "(dtmDtMovimento, intContaBancaria, intLote, dtmDtPagamento, dblPrincipal, dblMulta, dblJuros, dblCorrecao, dblTarifa, strCodigoDeBarras, dblCorreto, intLancamentoValor, intGuias, intCodigoBaixa, intDigito, BytTipo, dtmDtAtualizacao, lngCodUsr, bitGuia, bitProcessado) " & _
                                                                                        "VALUES(" & gstrConvDtParaSql(txt_DtMovimento.Text) & "," & lngPkidContaBancaria & "," & intLote & "," & gstrConvDtParaSql(Mid(strLinha, 26, 2) & "/" & Mid(strLinha, 28, 2) & "/" & Mid(strLinha, 30, 2)) & "," & gstrConvVrParaSql(dblValorPrincipal) & "," & gstrConvVrParaSql(dblValorMulta) & "," & gstrConvVrParaSql(dblValorJuros) & "," & gstrConvVrParaSql(dblValorCorrecao) & "," & 0 & ",'" & IIf(blnLayoutNovo, Mid(strLinha, 2, 21), Mid(strLinha, 2, 21)) & "',0,NULL,NULL,NULL,0," & _
                                                                                        "4," & gstrConvDtParaSql(gstrDataDoSistema) & "," & glngCodUsr & ",0,0)"

                    Else
                       gobjBanco.Execute "INSERT INTO " & gstrMovimentoBancario & "(dtmDtMovimento, intContaBancaria, intLote, dtmDtPagamento, dblPrincipal, dblMulta, dblJuros, dblCorrecao, dblTarifa, strCodigoDeBarras, dblCorreto, intLancamentoValor, intGuias, intCodigoBaixa, intDigito, BytTipo, dtmDtAtualizacao, lngCodUsr, bitGuia, bitProcessado) " & _
                                                                                        "VALUES(" & gstrConvDtParaSql(txt_DtMovimento.Text) & "," & lngPkidContaBancaria & "," & intLote & "," & gstrConvDtParaSql(Mid(strLinha, 26, 2) & "/" & Mid(strLinha, 28, 2) & "/" & Mid(strLinha, 30, 2)) & "," & gstrConvVrParaSql(dblValorPrincipal - dblValorTarifa) & "," & gstrConvVrParaSql(dblValorMulta) & "," & gstrConvVrParaSql(dblValorJuros) & "," & gstrConvVrParaSql(dblValorCorrecao) & "," & gstrConvVrParaSql(dblValorTarifa) & ",'" & IIf(blnLayoutNovo, Mid(strLinha, 2, 21), Mid(strLinha, 2, 21)) & "',0,NULL,NULL,NULL,0," & _
                                                                                        "4," & gstrConvDtParaSql(gstrDataDoSistema) & "," & glngCodUsr & ",0,0)"

                    End If
                End If

            End If

        ElseIf Mid(strLinha, 1, 1) = "9" Then
            bytQtdeTipos = bytQtdeTipos + 1

            'Vamos verificar se ja existe um registro referente, caso exista vamos somar o valor
            strSQL = "SELECT Pkid, dblValor FROM " & gstrResumoBancario & " WHERE dtmData = " & gstrConvDtParaSql(txt_DtMovimento.Text) & " AND intLote = " & intLote & " AND intContaBancaria = " & lngPkidContaBancaria
            If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
                If Not adoResultado.EOF Then

                    'Vamos somar registro na tabela tblResumoBancario
                    If DesContarTarifa = False Then
                        strSQL = "UPDATE " & gstrResumoBancario & " SET dblValor = " & gstrConvVrParaSql(dblTotal) & " WHERE Pkid = " & adoResultado("Pkid").Value
                    Else
                        strSQL = "UPDATE " & gstrResumoBancario & " SET dblValor = " & gstrConvVrParaSql(dblTotal) & " WHERE Pkid = " & adoResultado("Pkid").Value
                    End If
                    
                Else
                    'Vamos gravar registro na tabela tblResumoBancario
                    If DesContarTarifa = False Then
                       strSQL = "INSERT INTO " & gstrResumoBancario & "(dtmData, intContaBancaria, intLote, dblValor, dtmDtAtualizacao, lngCodUsr) " & _
                                 "VALUES(" & gstrConvDtParaSql(txt_DtMovimento.Text) & "," & lngPkidContaBancaria & "," & intLote & "," & Replace(dblTotal, ",", ".") & "," & gstrConvDtParaSql(gstrDataDoSistema) & "," & glngCodUsr & ")"
                    Else
                        strSQL = "INSERT INTO " & gstrResumoBancario & "(dtmData, intContaBancaria, intLote, dblValor, dtmDtAtualizacao, lngCodUsr) " & _
                             "VALUES(" & gstrConvDtParaSql(txt_DtMovimento.Text) & "," & lngPkidContaBancaria & "," & intLote & "," & Replace(dblTotal, ",", ".") & "," & gstrConvDtParaSql(gstrDataDoSistema) & "," & glngCodUsr & ")"
                    End If
                    
                End If
            End If

            gobjBanco.Execute strSQL

            'Vamos verificar se a qtde gravada é igual ao total de registros
            If lngLinha - 1 <> Val(Mid(strLinha, 11, 6)) Then
                ExibeMensagem "A quantidade de registros processada não coincide com a informada no arquivo. A operação não concluída."
                Screen.MousePointer = vbDefault
                pgr_Status.Visible = False
                lbl_Status.Visible = False
                gobjBanco.ExecutaRollbackTrans
                Close #1
                Exit Sub
            End If

        End If

        lngLinha = lngLinha + 1

        DoEvents
        lbl_Status.Caption = lngLinha & " de " & lngSize
        Me.Refresh

        pgr_Status.Value = lngLinha

ProximaLinha:

    Loop
     Close #1
    
    pgr_Status.Visible = False
    lbl_Status.Visible = False
        
    'Vamos verificar se foram lidos todos os tipos de registro
    If bytQtdeTipos < 3 Then
        ExibeMensagem "O arquivo de leitura está incorreto."
        Screen.MousePointer = vbDefault
        gobjBanco.ExecutaRollbackTrans
        Close #1
        Exit Sub
    End If
    
    gobjBanco.ExecutaCommitTrans
    
    Screen.MousePointer = vbDefault
    
    Exit Sub

err_BaixaAutomatica:
    ExibeMensagem "Ocorreu um Erro na leitura do arquivo. A operação não foi concluída."
    Screen.MousePointer = vbDefault
    pgr_Status.Visible = False
    lbl_Status.Visible = False
    gobjBanco.ExecutaRollbackTrans
    Close #1
    
End Sub

Private Sub LeMovimentoBancarioFichaCompensacaoCEF()
Dim strLinha               As String
Dim lngLinha               As Long
Dim lngSize                As Long

Dim blnTipoJaSomado        As Boolean
Dim strSQL                 As String
Dim strSqlSub              As String

Dim adoResultado           As New ADODB.Recordset
Dim adoAtualizacao         As New ADODB.Recordset

Dim bytQtdeTipos           As Byte ' Tipos (0,1,9)
Dim dblPorcentagemDif      As Double
Dim dblValorDiferenca      As Double
Dim dblValorDoArquivo      As Double
Dim dblValorTarifa         As Double
Dim dblValorTotDoArquivo   As Double

Dim dblSomaDosValores      As Double

Dim dblValorPrincipal      As Double
Dim dblValorMulta          As Double
Dim dblValorJuros          As Double
Dim dblValorCorrecao       As Double
Dim dblValorDesconto       As Double
Dim dblValorCorreto        As Double

Dim lngPkidContaBancaria   As Long
Dim blnExisteGuia          As Boolean
Dim blnLayoutNovo          As Boolean
Dim intLote                As Integer

    On Error GoTo err_BaixaAutomatica
    
    bytQtdeTipos = 0
    lngLinha = 0
    blnTipoJaSomado = False
    
    Open txt_Arquivo For Input As #1
    
    lngSize = FileLen(txt_Arquivo) / 400
    
    pgr_Status.Value = 0
    pgr_Status.Visible = True
    pgr_Status.Max = Abs(lngSize)
    lbl_Status.Visible = True
    
    Set gobjBanco = New clsBanco
    
    gobjBanco.ExecutaBeginTrans
    
    Do While Not EOF(1)

        Line Input #1, strLinha

        If Len(strLinha) = 0 Then
            GoTo ProximaLinha
        End If
        
        'Vamos verificar se a linha contem 150 posicoes
        If Len(strLinha) <> 400 Then
            ExibeMensagem "Ocorreu um Erro na leitura do arquivo. A operação não foi concluída."
            Screen.MousePointer = vbDefault
            pgr_Status.Visible = False
            lbl_Status.Visible = False
            gobjBanco.ExecutaRollbackTrans
            Close #1
            Exit Sub
        End If
        
        If Mid(strLinha, 1, 1) = "0" Then
            bytQtdeTipos = bytQtdeTipos + 1
            
            'Vamos buscar o Pkid referente à Conta em tblContaBancaria
            strSQL = "SELECT Pkid FROM " & gstrContaBancaria & " WHERE strContaRetorno = '" & Trim(Mid(strLinha, 34, 8)) & "'"
            If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
                If adoResultado.EOF Then
                    ExibeMensagem "Não foi encontrada Conta referente à este arquivo. A operação não concluída."
                    Screen.MousePointer = vbDefault
                    pgr_Status.Visible = False
                    lbl_Status.Visible = False
                    gobjBanco.ExecutaRollbackTrans
                    Close #1
                    Exit Sub
                Else
                    lngPkidContaBancaria = adoResultado("Pkid").Value
                End If
            End If
            
            'Vamos atribuir o numero do lote
            'Caso nas posicoes 390 a 394 estiver zerado vamos obter o lote, caso contrario pegaremos o enviado no arquivo
            If Val(Mid(strLinha, 390, 5)) = 0 Then
                strSQL = "SELECT Max(intLote) UltimoLote FROM " & gstrMovimentoBancario & " WHERE intContaBancaria = " & lngPkidContaBancaria & " AND bytTipo <> 0"
                If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
                    intLote = IIf(IsNull(adoResultado("UltimoLote").Value), 1, adoResultado("UltimoLote").Value + 1)
                End If
            Else
                intLote = Val(Mid(strLinha, 390, 5))
            End If
            
        ElseIf Mid(strLinha, 1, 1) = "1" Then
            
            blnLayoutNovo = False 'Len(Trim(Mid(strLinha, 105, 13))) > 0
            
            If Not blnTipoJaSomado Then
                bytQtdeTipos = bytQtdeTipos + 1
                blnTipoJaSomado = True
            End If
            
            blnExisteGuia = True
                         
           strSqlSub = ""
           strSqlSub = "SELECT " & gstrTOPnSQLServer(1) & " Pkid FROM " & gstrGuias & " WHERE strCodBarraEsp = '" & IIf(blnLayoutNovo, Mid(strLinha, 63, 11), Mid(strLinha, 63, 11)) & "' ORDER BY Pkid Desc"
           strSqlSub = gstrTOPnOracle(strSqlSub, 1)
        
           'Vamos buscar os dados referentes às guias apenas com a parte especifica do codigo de barras
           strSQL = "SELECT LA.intComposicaoDaReceita, LA.intExercicio, LV.dblValor ValorOrig, LV.dtmDtVencimento, LV.intMoeda, LV.Pkid intLancamentoValor, LV.intParcela, LG.intGuias, LV.DtmDtVencimento, LG.DBLVALORPRINCIPAL, LG.DBLVALORMULTA, LG.DBLVALORJUROS, LG.Dblvalorcorrecao, LG.DBLVALORDESCONTO, (LG.DBLVALORPRINCIPAL + LG.DBLVALORMULTA + LG.DBLVALORJUROS + LG.Dblvalorcorrecao - LG.DBLVALORDESCONTO) dblCorreto, " & _
                              gstrCONVERT(CDT_numeric, "LA.STRNUMEROAVISO") & " strNumeroAviso," & _
                              "(SELECT SUM(LG2.DBLVALORPRINCIPAL + LG2.DBLVALORMULTA + LG2.DBLVALORJUROS + LG2.Dblvalorcorrecao - LG2.DBLVALORDESCONTO) FROM " & gstrLancamentoGuias & " LG2 WHERE LG2.intguias = LG.intGuias) dblTotal " & _
                    "FROM " & gstrLancamentoGuias & " LG, " & _
                              gstrLancamentoValor & " LV, " & _
                              gstrLancamentoAlfa & " LA " & _
                    "WHERE LV.Pkid = LG.INTLANCAMENTOVALOR AND " & _
                    "      LA.Pkid = LV.INTLANCAMENTOALFA AND " & _
                           "LG.intGuias = (" & strSqlSub & ")"
                    
           If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
               If adoResultado.EOF Then
        
                   blnExisteGuia = False
                   'ExibeMensagem "Não foi encontrada Conta referente à este arquivo. A operação não concluída."
                   'Screen.MousePointer = vbDefault
                   'pgr_Status.Visible = False
                   'gobjBanco.ExecutaRollbackTrans
                    'Close #1
                    'Exit Sub
                End If
                
                dblPorcentagemDif = 0
                dblSomaDosValores = 0
                dblValorDiferenca = 0
                '(valor principal + juros + multa) - (abatimento + descontos)
                dblValorDoArquivo = (ConverteValorDoArquivo(Mid(strLinha, 254, 13)) + ConverteValorDoArquivo(Mid(strLinha, 267, 13)) + ConverteValorDoArquivo(Mid(strLinha, 280, 13))) - (ConverteValorDoArquivo(Mid(strLinha, 228, 13)) + ConverteValorDoArquivo(Mid(strLinha, 241, 13)))
                dblValorTarifa = ConverteValorDoArquivo(IIf(blnLayoutNovo, Mid(strLinha, 176, 13), Mid(strLinha, 176, 13)))
                
                dblValorDoArquivo = dblValorDoArquivo + dblValorTarifa
                 
                'Vamos somar o valor totoal do arquivo
                dblValorTotDoArquivo = dblValorTotDoArquivo + dblValorDoArquivo
                
                If blnExisteGuia Then 'Verifica se o código de barras vindo do arquivo existe no banco
                
                    'Vamos verificar se existe diferenca de valores
                    If dblValorDoArquivo <> adoResultado("dblTotal").Value Then
                        'dblPorcentagemDif = (dblValorDoArquivo / adoResultado("dblTotal").Value)
                        'PROVISORIO
                        dblPorcentagemDif = dblValorDoArquivo - adoResultado("dblTotal").Value
                    End If
                    
                    'Vamos gravar um registro em Movimento Bancario para cada Guia
                    Do While Not adoResultado.EOF
                        
                        'Caso exista diferenca dos valores, vamos tirar a diferenca proporcionalmente
                        If dblPorcentagemDif <> 0 Then
                            'dblValorPrincipal = gstrConvVrDoSql((adoResultado("dblvalorprincipal") * dblPorcentagemDif), 2)
                            'dblValorMulta = gstrConvVrDoSql((adoResultado("dblvalorMulta") * dblPorcentagemDif), 2)
                            'dblValorJuros = gstrConvVrDoSql((adoResultado("dblvalorJuros") * dblPorcentagemDif), 2)
                            'dblValorCorrecao = gstrConvVrDoSql((adoResultado("dblvalorCorrecao") * dblPorcentagemDif), 2)
                            'dblValorDesconto = gstrConvVrDoSql((adoResultado("dblvalorDesconto") * dblPorcentagemDif), 2)
                            'PROVISORIO
                            If dblPorcentagemDif > 0 Then
                                dblValorPrincipal = gstrConvVrDoSql((dblValorDoArquivo - dblPorcentagemDif), 2)
                                dblValorMulta = gstrConvVrDoSql(dblPorcentagemDif, 2)
                            Else
                                dblValorPrincipal = gstrConvVrDoSql((dblValorDoArquivo + dblPorcentagemDif), 2)
                                dblValorMulta = 0
                            End If
                            dblValorJuros = 0
                            dblValorCorrecao = 0
                            dblValorDesconto = 0
                        Else
                            dblValorPrincipal = adoResultado("dblValorPrincipal").Value '(ConverteValorDoArquivo(Mid(strLinha, 254, 13)) + ConverteValorDoArquivo(Mid(strLinha, 267, 13)) + ConverteValorDoArquivo(Mid(strLinha, 280, 13))) - (ConverteValorDoArquivo(Mid(strLinha, 228, 13)) + ConverteValorDoArquivo(Mid(strLinha, 241, 13)))
                            dblValorMulta = adoResultado("dblValorMulta").Value
                            dblValorJuros = adoResultado("dblValorJuros").Value
                            dblValorCorrecao = adoResultado("dblValorCorrecao").Value
                            dblValorDesconto = adoResultado("dblValorDesconto").Value
                        End If
                        
                        dblSomaDosValores = dblSomaDosValores + (dblValorPrincipal + dblValorMulta + dblValorJuros + dblValorCorrecao + dblValorDesconto)
                        
                        'Se este for o ultimo registro, vamos verificar se ha diferenca dos valores calculados
                        If adoResultado.RecordCount = adoResultado.AbsolutePosition Then
                            'Caso depois do acerto de valores ainda exista diferenca, a jogaremos no ultimo registro
                            If CCur(dblSomaDosValores) <> CCur(dblValorDoArquivo) Then
                            
                                dblValorDiferenca = dblValorDoArquivo - dblSomaDosValores
                        
                                'strSQL = "UPDATE " & gstrMovimentoBancario
                                If dblValorPrincipal > 0 Then
                                    dblValorPrincipal = dblValorPrincipal + dblValorDiferenca
                                ElseIf dblValorMulta > 0 Then
                                    dblValorMulta = dblValorMulta + dblValorDiferenca
                                ElseIf dblValorJuros > 0 Then
                                    dblValorJuros = dblValorJuros + dblValorDiferenca
                                Else
                                    dblValorCorrecao = dblValorCorrecao + dblValorDiferenca
                                End If
                            End If
                        End If
                        
                        strSQL = gstrStoredProcedure("sp_AtualizaParcela", adoResultado("intComposicaoDaReceita").Value & ", " & adoResultado("intExercicio").Value & ", " & adoResultado("intLancamentoValor").Value & ", " & gstrConvDtParaSql(adoResultado("dtmDtVencimento").Value) & ", " & gstrConvDtParaSql(Mid(strLinha, 111, 2) & "/" & Mid(strLinha, 113, 2) & "/" & Mid(strLinha, 115, 2)) & ", " & gstrConvVrParaSql(adoResultado("ValorOrig").Value) & ", " & adoResultado("intMoeda").Value, True)
                        If gobjBanco.CriaADO(strSQL, 80, adoAtualizacao) Then
                            If gstrConvVrDoSql(adoResultado("ValorOrig").Value, , , True) = 0 Then
                                dblValorPrincipal = dblValorDoArquivo
                                dblValorCorreto = dblValorDoArquivo
                                dblValorMulta = 0
                            Else
                                dblValorCorreto = Space$(0) & gstrConvVrDoSql(Val(gstrConvVrParaSql(adoAtualizacao("dblValorPrincipal").Value)) + Val(gstrConvVrParaSql(adoAtualizacao("dblValorMulta").Value)) + Val(gstrConvVrParaSql(adoAtualizacao("dblValorJuros").Value)) + Val(gstrConvVrParaSql(adoAtualizacao("dblValorCorrecao").Value)))
                            End If
                        Else
                            Screen.MousePointer = vbDefault
                            pgr_Status.Visible = False
                            lbl_Status.Visible = False
                            gobjBanco.ExecutaRollbackTrans
                            Close #1
                            Exit Sub

                        End If
                        '*****************************************************
                        ' Data          : 20/12/2005                         *
                        ' Criação       : Verificação de Desconto de Tarifa  *
                        ' Responsável   : Fernando Peixoto                   *
                        ' Pendência     : Tri0518                            *
                        ''****************************************************
                        
                        If DesContarTarifa = False Then
                            gobjBanco.Execute "INSERT INTO " & gstrMovimentoBancario & "(dtmDtMovimento, intContaBancaria, intLote, dtmDtPagamento, dblPrincipal, dblMulta, dblJuros, dblCorrecao, dblTarifa, strCodigoDeBarras, dblCorreto, intLancamentoValor, intGuias, intCodigoBaixa, intDigito, BytTipo, dtmDtAtualizacao, lngCodUsr, bitGuia, bitProcessado) " & _
                                                                                        "VALUES(" & gstrConvDtParaSql(txt_DtMovimento.Text) & "," & lngPkidContaBancaria & "," & intLote & "," & gstrConvDtParaSql(Mid(strLinha, 111, 2) & "/" & Mid(strLinha, 113, 2) & "/" & Mid(strLinha, 115, 2)) & "," & gstrConvVrParaSql(dblValorPrincipal) & "," & gstrConvVrParaSql(dblValorMulta) & "," & gstrConvVrParaSql(dblValorJuros) & "," & gstrConvVrParaSql(dblValorCorrecao) & "," & 0 & ",'" & Mid(strLinha, 63, 11) & "'," & gstrConvVrParaSql(dblValorCorreto) & "," & adoResultado("intLancamentoValor").Value & "," & adoResultado("intGuias").Value & "," & _
                                                                                        IIf(CDate(adoResultado("dtmDtVencimento").Value) < CDate(Mid(strLinha, 111, 2) & "/" & Mid(strLinha, 113, 2) & "/" & Mid(strLinha, 115, 2)), gstrENulo(intCodBaixaAposVcto, , True), gstrENulo(intCodBaixaAntesVcto, , True)) & "," & gstrCalculaDigitoModulo10(Trim(adoResultado("strNumeroAviso").Value) & Format$(Trim(adoResultado("intParcela").Value), "000")) & _
                                                                                        ",4" & "," & gstrConvDtParaSql(gstrDataDoSistema) & "," & glngCodUsr & ",1,0)"
                        Else
                            gobjBanco.Execute "INSERT INTO " & gstrMovimentoBancario & "(dtmDtMovimento, intContaBancaria, intLote, dtmDtPagamento, dblPrincipal, dblMulta, dblJuros, dblCorrecao, dblTarifa, strCodigoDeBarras, dblCorreto, intLancamentoValor, intGuias, intCodigoBaixa, intDigito, BytTipo, dtmDtAtualizacao, lngCodUsr, bitGuia, bitProcessado) " & _
                                                                                        "VALUES(" & gstrConvDtParaSql(txt_DtMovimento.Text) & "," & lngPkidContaBancaria & "," & intLote & "," & gstrConvDtParaSql(Mid(strLinha, 111, 2) & "/" & Mid(strLinha, 113, 2) & "/" & Mid(strLinha, 115, 2)) & "," & gstrConvVrParaSql(dblValorPrincipal - dblValorTarifa) & "," & gstrConvVrParaSql(dblValorMulta) & "," & gstrConvVrParaSql(dblValorJuros) & "," & gstrConvVrParaSql(dblValorCorrecao) & "," & gstrConvVrParaSql(dblValorTarifa) & ",'" & Mid(strLinha, 63, 11) & "'," & gstrConvVrParaSql(dblValorCorreto) & "," & adoResultado("intLancamentoValor").Value & "," & adoResultado("intGuias").Value & "," & _
                                                                                        IIf(CDate(adoResultado("dtmDtVencimento").Value) < CDate(Mid(strLinha, 111, 2) & "/" & Mid(strLinha, 113, 2) & "/" & Mid(strLinha, 115, 2)), gstrENulo(intCodBaixaAposVcto, , True), gstrENulo(intCodBaixaAntesVcto, , True)) & "," & gstrCalculaDigitoModulo10(Trim(adoResultado("strNumeroAviso").Value) & Format$(Trim(adoResultado("intParcela").Value), "000")) & _
                                                                                        ",4" & "," & gstrConvDtParaSql(gstrDataDoSistema) & "," & glngCodUsr & ",1,0)"
                        End If
                        
                        adoResultado.MoveNext
                        
                    Loop
                    
                Else
                    
                   dblValorPrincipal = dblValorDoArquivo
                   dblValorMulta = 0
                   dblValorJuros = 0
                   dblValorCorrecao = 0
                   dblValorDesconto = 0
                        
                    '*****************************************************
                    ' Data          : 20/12/2005                         *
                    ' Criação       : Verificação de Desconto de Tarifa  *
                    ' Responsável   : Fernando Peixoto                   *
                    ' Pendência     : Tri0518                            *
                    ''****************************************************
                    
                    If DesContarTarifa = False Then
                       gobjBanco.Execute "INSERT INTO " & gstrMovimentoBancario & "(dtmDtMovimento, intContaBancaria, intLote, dtmDtPagamento, dblPrincipal, dblMulta, dblJuros, dblCorrecao, dblTarifa, strCodigoDeBarras, dblCorreto, intLancamentoValor, intGuias, intCodigoBaixa, intDigito, BytTipo, dtmDtAtualizacao, lngCodUsr, bitGuia, bitProcessado) " & _
                                                                                        "VALUES(" & gstrConvDtParaSql(txt_DtMovimento.Text) & "," & lngPkidContaBancaria & "," & intLote & "," & gstrConvDtParaSql(Mid(strLinha, 111, 2) & "/" & Mid(strLinha, 113, 2) & "/" & Mid(strLinha, 115, 2)) & "," & gstrConvVrParaSql(dblValorPrincipal) & "," & gstrConvVrParaSql(dblValorMulta) & "," & gstrConvVrParaSql(dblValorJuros) & "," & gstrConvVrParaSql(dblValorCorrecao) & "," & 0 & ",'" & Mid(strLinha, 63, 11) & "',0,NULL,NULL,NULL,0," & _
                                                                                        "4," & gstrConvDtParaSql(gstrDataDoSistema) & "," & glngCodUsr & ",0,0)"
                    Else
                       gobjBanco.Execute "INSERT INTO " & gstrMovimentoBancario & "(dtmDtMovimento, intContaBancaria, intLote, dtmDtPagamento, dblPrincipal, dblMulta, dblJuros, dblCorrecao, dblTarifa, strCodigoDeBarras, dblCorreto, intLancamentoValor, intGuias, intCodigoBaixa, intDigito, BytTipo, dtmDtAtualizacao, lngCodUsr, bitGuia, bitProcessado) " & _
                                                                                        "VALUES(" & gstrConvDtParaSql(txt_DtMovimento.Text) & "," & lngPkidContaBancaria & "," & intLote & "," & gstrConvDtParaSql(Mid(strLinha, 111, 2) & "/" & Mid(strLinha, 113, 2) & "/" & Mid(strLinha, 115, 2)) & "," & gstrConvVrParaSql(dblValorPrincipal - dblValorTarifa) & "," & gstrConvVrParaSql(dblValorMulta) & "," & gstrConvVrParaSql(dblValorJuros) & "," & gstrConvVrParaSql(dblValorCorrecao) & "," & gstrConvVrParaSql(dblValorTarifa) & ",'" & Mid(strLinha, 63, 11) & "',0,NULL,NULL,NULL,0," & _
                                                                                        "4," & gstrConvDtParaSql(gstrDataDoSistema) & "," & glngCodUsr & ",0,0)"
                    End If
                End If
                
           End If
             
        ElseIf Mid(strLinha, 1, 1) = "9" Then
            bytQtdeTipos = bytQtdeTipos + 1
                         
            'Vamos verificar se ja existe um registro referente, caso exista vamos somar o valor
            strSQL = "SELECT Pkid, dblValor FROM " & gstrResumoBancario & " WHERE dtmData = " & gstrConvDtParaSql(txt_DtMovimento.Text) & " AND intLote = " & intLote & " AND intContaBancaria = " & lngPkidContaBancaria
            If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
                If Not adoResultado.EOF Then
             
                    'Vamos somar registro na tabela tblResumoBancario
                    strSQL = "UPDATE " & gstrResumoBancario & " SET dblValor = " & gstrConvVrParaSql(dblValorTotDoArquivo + adoResultado("dblValor").Value) & " WHERE Pkid = " & adoResultado("Pkid").Value
                     
                Else
                
                    'Vamos gravar registro na tabela tblResumoBancario
                    strSQL = "INSERT INTO " & gstrResumoBancario & "(dtmData, intContaBancaria, intLote, dblValor, dtmDtAtualizacao, lngCodUsr) " & _
                             "VALUES(" & gstrConvDtParaSql(txt_DtMovimento.Text) & "," & lngPkidContaBancaria & "," & intLote & "," & gstrConvVrParaSql(dblValorTotDoArquivo) & "," & gstrConvDtParaSql(gstrDataDoSistema) & "," & glngCodUsr & ")"
                
                End If
            End If
            
            gobjBanco.Execute strSQL
            
            'Vamos verificar se a qtde gravada é igual ao total de registros
            If lngLinha + 1 <> Val(Mid(strLinha, 395, 6)) Then
                ExibeMensagem "A quantidade de registros processada não coincide com a informada no arquivo. A operação não concluída."
                Screen.MousePointer = vbDefault
                pgr_Status.Visible = False
                lbl_Status.Visible = False
                gobjBanco.ExecutaRollbackTrans
                Close #1
                Exit Sub
            End If
            
        End If
        
        lngLinha = lngLinha + 1
        
        DoEvents
        lbl_Status.Caption = lngLinha & " de " & lngSize
        Me.Refresh
        
        pgr_Status.Value = lngLinha
        
ProximaLinha:

    Loop
        
    Close #1
    
    pgr_Status.Visible = False
    lbl_Status.Visible = False
        
    'Vamos verificar se foram lidos todos os tipos de registro
    If bytQtdeTipos < 3 Then
        ExibeMensagem "O arquivo de leitura está incorreto."
        Screen.MousePointer = vbDefault
        gobjBanco.ExecutaRollbackTrans
        Close #1
        Exit Sub
    End If
    
    gobjBanco.ExecutaCommitTrans
    
    Screen.MousePointer = vbDefault
    
    Exit Sub

err_BaixaAutomatica:
    ExibeMensagem "Ocorreu um Erro na leitura do arquivo. A operação não foi concluída."
    Screen.MousePointer = vbDefault
    pgr_Status.Visible = False
    lbl_Status.Visible = False
    gobjBanco.ExecutaRollbackTrans
    Close #1

End Sub

Public Sub MantemForm(ByVal strModoOperacao As String)
Dim strLinha As String

    Select Case strModoOperacao
        
        Case gstrLerArquivo
            
            If Not blnDadosok Then Exit Sub
            
            Screen.MousePointer = vbHourglass
            
            'Desvio para função que verifica se a tarifa bancaria será descontada ou não
            DesContarTarifaBancaria
            
            'Vamos identificar se é FEBRABAM ou FICHA DE COMPENSAÇÃO
            Open txt_Arquivo For Input As #1
        
ProximaLinha:

            Line Input #1, strLinha
    
            If Len(strLinha) = 0 Then
                GoTo ProximaLinha
            End If
            
'            Close #1
            
            'Vamos obter os Codigos de Baixa de vencimento
            If Not blnObtemCodigoBaixa Then Exit Sub
            
            If UCase(Mid(strLinha, 1, 1)) = "A" Then
                
                Line Input #1, strLinha
                Close #1
                
                If UCase(Mid(strLinha, 1, 1)) = "B" Or UCase(Mid(strLinha, 1, 1)) = "F" Then
                    LeMovimentoBancarioDebitoAutomatico
                Else
                    LeMovimentoBancarioFebrabam
                End If
                
            Else
                Close #1
                
                If Mid(strLinha, 80, 15) = "BANCO DO BRASIL" Then
                    LeMovimentoBancarioFichaCompensacaoBB
                ElseIf Len(strLinha) = 400 Then
                    LeMovimentoBancarioFichaCompensacaoCEF
                Else
                    LeMovimentoBancarioFichaCompensacaoBanespa
                End If
            End If
                
        Case gstrFechar
            Unload Me
    
    End Select

End Sub

Private Sub txt_DtMovimento_GotFocus()
    MarcaCampo txt_DtMovimento
End Sub

Private Sub txt_DtMovimento_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "D", txt_DtMovimento
End Sub

Private Sub txt_DtMovimento_LostFocus()
    txt_DtMovimento.Text = gstrDataFormatada(txt_DtMovimento.Text)
End Sub

Private Function ConverteDataDoArquivo(strData As String, blnLayoutNovo As Boolean) As Date
Dim strAux As String
    If blnLayoutNovo Then
        strAux = Right(strData, 2) & "/" & Mid(strData, 5, 2) & "/" & Left(strData, 4)
    Else
        strAux = Right(strData, 2) & "/" & Mid(strData, 3, 2) & "/" & Left(strData, 2)
    End If
    
    ConverteDataDoArquivo = CDate(strAux)
    
End Function

Private Function ConverteValorDoArquivo(strValor As String) As Double
Dim dblAux As Double

    dblAux = Left(strValor, Len(strValor) - 2) & "," & Right(strValor, 2)
    
    ConverteValorDoArquivo = gstrConvVrDoSql(dblAux)
    
End Function

Private Function ConverteValorDoArquivoAcordo(strValor As String) As Double
Dim dblAux As Double

    dblAux = Left(strValor, Len(strValor) - 4) & "," & Right(strValor, 4)
    
    ConverteValorDoArquivoAcordo = gstrConvVrDoSql(dblAux)
    
End Function

Private Sub cmdProvisorio_Click()
Dim strLinha               As String
Dim lngLinha               As Long
Dim lngSize                As Long

Dim blnTipoJaSomado        As Boolean

Dim bytQtdeTipos           As Byte ' Tipos (A,G,Z)
Dim dblValorDoArquivo      As Double
Dim dblValorCobrado        As Double
Dim dblValorMulta          As Double

Dim blnLayoutNovo          As Boolean

    On Error GoTo err_BaixaAutomatica
        
    If Not blnDadosok Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    
    bytQtdeTipos = 0
    lngLinha = 0
    blnTipoJaSomado = False
    
    Open txt_Arquivo For Input As #1
    
    lngSize = FileLen(txt_Arquivo) / 150
    
    pgr_Status.Value = 0
    pgr_Status.Visible = True
    pgr_Status.Max = Abs(lngSize)
    lbl_Status.Visible = True
    
    Set gobjBanco = New clsBanco
    
    gobjBanco.ExecutaBeginTrans
    
    Do While Not EOF(1)
        
ProximaLinha:

        Line Input #1, strLinha
        
        If Len(strLinha) = 0 Then
            GoTo ProximaLinha
        End If
        
        'Vamos verificar se a linha contem 150 posicoes
        If Len(strLinha) <> 150 Then
            ExibeMensagem "Ocorreu um Erro na leitura do arquivo. A operação não foi concluída."
            Screen.MousePointer = vbDefault
            pgr_Status.Visible = False
            lbl_Status.Visible = False
            gobjBanco.ExecutaRollbackTrans
            Close #1
            Exit Sub
        End If
        
        If Mid(strLinha, 1, 1) = "A" Then
            bytQtdeTipos = bytQtdeTipos + 1
            
        ElseIf Mid(strLinha, 1, 1) = "G" Then
            
            blnLayoutNovo = Len(Trim(Mid(strLinha, 105, 13))) > 0
            
            If Not blnTipoJaSomado Then
                bytQtdeTipos = bytQtdeTipos + 1
                blnTipoJaSomado = True
            End If
                
            dblValorDoArquivo = ConverteValorDoArquivo(IIf(blnLayoutNovo, Mid(strLinha, 82, 12), Mid(strLinha, 78, 12)))
            
            'Vamos verificar se é acordo para convertemos o valor
            Select Case IIf(blnLayoutNovo, Mid(strLinha, 63, 4), Mid(strLinha, 59, 4))
                Case Is = "0270", "0217", "0277"
                    dblValorCobrado = ConverteValorDoArquivoAcordo(IIf(blnLayoutNovo, Mid(strLinha, 42, 11), Mid(strLinha, 38, 11)))
                    dblValorCobrado = gstrConvVrDoSql(dblValorCobrado * 1.7595, 2)
                Case Else
                    dblValorCobrado = ConverteValorDoArquivo(IIf(blnLayoutNovo, Mid(strLinha, 42, 11), Mid(strLinha, 38, 11)))
            End Select
            
            dblValorMulta = 0
            
            If Abs(dblValorDoArquivo - dblValorCobrado) > 0.05 Then
                If dblValorDoArquivo - dblValorCobrado > 0 Then
                    dblValorMulta = dblValorDoArquivo - dblValorCobrado
                End If
            End If
            
            gobjBanco.Execute "INSERT INTO QBGRECEITAS (CONTA, SERIE, TRIBUTO, EXERCICIO, Principal, Multa, Juros, Correcao, DATA) " & _
                                                "VALUES('" & Trim(Mid(strLinha, 2, 20)) & "','" & IIf(blnLayoutNovo, Mid(strLinha, 69, 6), Mid(strLinha, 65, 6)) & "','" & IIf(blnLayoutNovo, Mid(strLinha, 63, 4), Mid(strLinha, 59, 4)) & "'," & IIf(blnLayoutNovo, Year("01/01/" & Mid(strLinha, 67, 2)), Year("01/01/" & Mid(strLinha, 63, 2))) & "," & gstrConvVrParaSql(dblValorDoArquivo - dblValorMulta) & "," & gstrConvVrParaSql(dblValorMulta) & ",NULL,NULL" & "," & gstrConvDtParaSql(txt_DtMovimento.Text) & ")"
                
                    
             
        ElseIf Mid(strLinha, 1, 1) = "Z" Then
            bytQtdeTipos = bytQtdeTipos + 1
                         
            'Vamos verificar se a qtde gravada é igual ao total de registros
            If lngLinha + 1 <> Val(Mid(strLinha, 2, 6)) Then
                ExibeMensagem "A quantidade de registros processada não coincide com a informada no arquivo. A operação não concluída."
                Screen.MousePointer = vbDefault
                pgr_Status.Visible = False
                lbl_Status.Visible = False
                gobjBanco.ExecutaRollbackTrans
                Close #1
                Exit Sub
            End If
                        
            Exit Do
        End If
        
        lngLinha = lngLinha + 1
        
        pgr_Status.Value = lngLinha

    Loop
        
    Close #1
    
    pgr_Status.Visible = False
    lbl_Status.Visible = False
        
    'Vamos verificar se foram lidos todos os tipos de registro
    If bytQtdeTipos < 3 Then
        ExibeMensagem "O arquivo de leitura está incorreto."
        Screen.MousePointer = vbDefault
        gobjBanco.ExecutaRollbackTrans
        Close #1
        Exit Sub
    End If
    
    gobjBanco.ExecutaCommitTrans
    
    Screen.MousePointer = vbDefault
    
    Exit Sub

err_BaixaAutomatica:
    ExibeMensagem "Ocorreu um Erro na leitura do arquivo. A operação não foi concluída."
    Screen.MousePointer = vbDefault
    pgr_Status.Visible = False
    lbl_Status.Visible = False
    gobjBanco.ExecutaRollbackTrans
    Close #1
    
End Sub

Private Function blnObtemCodigoBaixa() As Boolean
Dim adoResultado As New ADODB.Recordset
Dim strSQL       As String
    
    blnObtemCodigoBaixa = False
    
    'Vamos buscar o Pkid referente ao Codigo Baixa dos Vencimentos
    strSQL = "SELECT Min(pkid) intCodBaixaAposVcto, bv.intCodBaixaAntesVcto FROM (SELECT Min(pkid) intCodBaixaAntesVcto FROM " & gstrCodigoDeBaixa & " WHERE bytTipo = 0) bv, " & gstrCodigoDeBaixa & " cb WHERE bytTipo = 4 group by bv.intCodBaixaAntesVcto"
    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
        If adoResultado.EOF Then
            ExibeMensagem "Não foram encontrados registro de Código de Baixa. A operação não concluída."
            Screen.MousePointer = vbDefault
            Exit Function
        Else
            If IsNull(adoResultado("intCodBaixaAposVcto").Value) Or IsNull(adoResultado("intCodBaixaAntesVcto").Value) Then
                ExibeMensagem "Não foram encontrados registro de Código de Baixa. A operação não concluída."
                Screen.MousePointer = vbDefault
                Exit Function
            Else
                intCodBaixaAntesVcto = adoResultado("intCodBaixaAntesVcto").Value
                intCodBaixaAposVcto = adoResultado("intCodBaixaAposVcto").Value
            End If
        End If
    End If
    
    blnObtemCodigoBaixa = True
    
End Function
Private Function blnDadosok() As Boolean

    On Error GoTo err_blnDadosOK
    
    If Trim(txt_Arquivo) = "" Then
        ExibeMensagem "Indique a localização do arquivo de retorno."
        Exit Function
    ElseIf Dir(txt_Arquivo) = "" Then
        ExibeMensagem "Arquivo não encontrado no local especificado."
        Exit Function
    ElseIf Not gblnDataValida(txt_DtMovimento) Then
        ExibeMensagem "A Data deve ser preenchida corretamente."
        Exit Function
    End If
    
    blnDadosok = True
    Exit Function
    
err_blnDadosOK:
    blnDadosok = False

End Function

Private Function DesContarTarifaBancaria() As Boolean
'************************************************************************
' Data          : 21/12/2005                                            *
' Ficha         : Tri0518                                               *
' Criação       : Checa o campo blnDescontarTarifa :                    *
'                 0 = Somar tarifa ao valor principal                   *
'                 1 = Não alterar o valor principal                     *
' Responsável   : Fernando Peixoto                                      *
''***********************************************************************
Dim strSQL As String
Dim ADOTemp As ADODB.Recordset


DesContarTarifa = True

On Error GoTo Err_DesContarTarifaBancaria

strSQL = ""
strSQL = "SELECT *"
strSQL = strSQL & " FROM " & gstrParametrosTributario

    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSQL, 5, ADOTemp) Then
        If (Not ADOTemp.BOF And Not ADOTemp.EOF) Then
            With ADOTemp
                .MoveFirst
                If !blnDescontarTarifa = False Then
                    DesContarTarifa = False
                Else
                    DesContarTarifa = True
                End If
            End With
        End If
    End If
    
Exit Function
    
Err_DesContarTarifaBancaria:
    
    ExibeMensagem "Rotina Descontar Tarifa Bancária .: " & Err.Description
    DesContarTarifa = True
    

End Function

Private Sub LeMovimentoBancarioFichaCompensacaoBB()
Dim strLinha               As String
Dim lngLinha               As Long
Dim lngSize                As Long

Dim blnTipoJaSomado        As Boolean
Dim strSQL                 As String
Dim strSqlSub              As String
Dim strSqlAux              As String

Dim adoResultado           As New ADODB.Recordset
Dim adoAtualizacao         As New ADODB.Recordset
Dim adoAux                 As New ADODB.Recordset

Dim bytQtdeTipos           As Byte ' Tipos (0,7,9)
Dim dblPorcentagemDif      As Double
Dim dblValorDiferenca      As Double
Dim dblValorDoArquivo      As Double
Dim dblValorTarifa         As Double
Dim dblValorTotDoArquivo   As Double

Dim dblSomaDosValores      As Double

Dim dblValorPrincipal      As Double
Dim dblValorMulta          As Double
Dim dblValorJuros          As Double
Dim dblValorCorrecao       As Double
Dim dblValorDesconto       As Double
Dim dblValorCorreto        As Double

Dim lngPkidContaBancaria   As Long
Dim blnExisteGuia          As Boolean
Dim blnLayoutNovo          As Boolean
Dim intLote                As Integer

Dim strCodigoBarra         As String

    On Error GoTo err_BaixaAutomatica
    
    bytQtdeTipos = 0
    lngLinha = 0
    blnTipoJaSomado = False
    
    Open txt_Arquivo For Input As #1
    
    lngSize = FileLen(txt_Arquivo) / 400
    
    pgr_Status.Value = 0
    pgr_Status.Visible = True
    pgr_Status.Max = Abs(lngSize)
    lbl_Status.Visible = True
    
    Set gobjBanco = New clsBanco
    
    gobjBanco.ExecutaBeginTrans
    
    Do While Not EOF(1)

        Line Input #1, strLinha

        If Len(strLinha) = 0 Then
            GoTo ProximaLinha
        End If
        
        'Vamos verificar se a linha contem 400 posicoes
        If Len(strLinha) <> 400 Then
            ExibeMensagem "Ocorreu um Erro na leitura do arquivo. A operação não foi concluída."
            Screen.MousePointer = vbDefault
            pgr_Status.Visible = False
            lbl_Status.Visible = False
            gobjBanco.ExecutaRollbackTrans
            Close #1
            Exit Sub
        End If
        
        If Mid(strLinha, 1, 1) = "0" Then
            bytQtdeTipos = bytQtdeTipos + 1
            
            'Vamos buscar o Pkid referente à Conta em tblContaBancaria
            strSQL = "SELECT Pkid FROM " & gstrContaBancaria & " WHERE strContaRetorno = '" & Trim(Mid(strLinha, 32, 8)) & "'"
            If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
                If adoResultado.EOF Then
                    ExibeMensagem "Não foi encontrada Conta referente à este arquivo. A operação não concluída."
                    Screen.MousePointer = vbDefault
                    pgr_Status.Visible = False
                    lbl_Status.Visible = False
                    gobjBanco.ExecutaRollbackTrans
                    Close #1
                    Exit Sub
                Else
                    lngPkidContaBancaria = adoResultado("Pkid").Value
                End If
            End If
            
            'Vamos atribuir o numero do lote
            'Caso nas posicoes 101 a 107 estiver zerado vamos obter o lote, caso contrario pegaremos o enviado no arquivo
            If Val(Mid(strLinha, 101, 7)) = 0 Then
                strSQL = "SELECT Max(intLote) UltimoLote FROM " & gstrMovimentoBancario & " WHERE intContaBancaria = " & lngPkidContaBancaria & " AND bytTipo <> 0"
                If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
                    intLote = IIf(IsNull(adoResultado("UltimoLote").Value), 1, adoResultado("UltimoLote").Value + 1)
                End If
            Else
                intLote = Val(Mid(strLinha, 101, 7))
            End If
            
        ElseIf Mid(strLinha, 1, 1) = "7" Then
            
            blnLayoutNovo = False 'Len(Trim(Mid(strLinha, 105, 13))) > 0
            
            If Not blnTipoJaSomado Then
                bytQtdeTipos = bytQtdeTipos + 1
                blnTipoJaSomado = True
            End If
            
            blnExisteGuia = True
                                                 
           strSqlSub = ""
           strSqlSub = "SELECT " & gstrTOPnSQLServer(1) & " Pkid FROM " & gstrGuias & " WHERE strCodBarraEsp = '" & Mid(strLinha, 64, 17) & "' ORDER BY Pkid Desc"
           strSqlSub = gstrTOPnOracle(strSqlSub, 1)
        
           'Vamos buscar os dados referentes às guias apenas com a parte especifica do codigo de barras
           strSQL = "SELECT LA.intComposicaoDaReceita, LA.intExercicio, LV.dblValor ValorOrig, LV.dtmDtVencimento, LV.intMoeda, LV.Pkid intLancamentoValor, LV.intParcela, LG.intGuias, LV.DtmDtVencimento, LG.DBLVALORPRINCIPAL, LG.DBLVALORMULTA, LG.DBLVALORJUROS, LG.Dblvalorcorrecao, LG.DBLVALORDESCONTO, (LG.DBLVALORPRINCIPAL + LG.DBLVALORMULTA + LG.DBLVALORJUROS + LG.Dblvalorcorrecao - LG.DBLVALORDESCONTO) dblCorreto, " & _
                              gstrCONVERT(CDT_numeric, "LA.STRNUMEROAVISO") & " strNumeroAviso," & _
                              "(SELECT SUM(LG2.DBLVALORPRINCIPAL + LG2.DBLVALORMULTA + LG2.DBLVALORJUROS + LG2.Dblvalorcorrecao - LG2.DBLVALORDESCONTO) FROM " & gstrLancamentoGuias & " LG2 WHERE LG2.intguias = LG.intGuias) dblTotal " & _
                    "FROM " & gstrLancamentoGuias & " LG, " & _
                              gstrLancamentoValor & " LV, " & _
                              gstrLancamentoAlfa & " LA " & _
                    "WHERE LV.Pkid = LG.INTLANCAMENTOVALOR AND " & _
                    "      LA.Pkid = LV.INTLANCAMENTOALFA AND " & _
                           "LG.intGuias = (" & strSqlSub & ")"
                    
           If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
               If adoResultado.EOF Then
        
                   blnExisteGuia = False
                   'ExibeMensagem "Não foi encontrada Conta referente à este arquivo. A operação não concluída."
                   'Screen.MousePointer = vbDefault
                   'pgr_Status.Visible = False
                   'gobjBanco.ExecutaRollbackTrans
                    'Close #1
                    'Exit Sub
                End If
                
                dblPorcentagemDif = 0
                dblSomaDosValores = 0
                dblValorDiferenca = 0
                '(valor principal + juros + multa) - (abatimento + descontos)
                dblValorDoArquivo = (ConverteValorDoArquivo(Mid(strLinha, 153, 13)) + ConverteValorDoArquivo(Mid(strLinha, 267, 13)) + ConverteValorDoArquivo(Mid(strLinha, 280, 13))) - (ConverteValorDoArquivo(Mid(strLinha, 228, 13)) + ConverteValorDoArquivo(Mid(strLinha, 241, 13)))
                dblValorTarifa = ConverteValorDoArquivo(IIf(blnLayoutNovo, Mid(strLinha, 182, 7), Mid(strLinha, 182, 7)))
                
                'Vamos somar o valor totoal do arquivo
                dblValorTotDoArquivo = dblValorTotDoArquivo + dblValorDoArquivo
                
                If blnExisteGuia Then 'Verifica se o código de barras vindo do arquivo existe no banco
                
                    'Vamos verificar se existe diferenca de valores
                    If dblValorDoArquivo <> adoResultado("dblTotal").Value Then
                        'dblPorcentagemDif = (dblValorDoArquivo / adoResultado("dblTotal").Value)
                        'PROVISORIO
                        dblPorcentagemDif = dblValorDoArquivo - adoResultado("dblTotal").Value
                    End If
                    
                    'Vamos gravar um registro em Movimento Bancario para cada Guia
                    Do While Not adoResultado.EOF
                        
                        'Caso exista diferenca dos valores, vamos tirar a diferenca proporcionalmente
                        If dblPorcentagemDif <> 0 Then
                            'dblValorPrincipal = gstrConvVrDoSql((adoResultado("dblvalorprincipal") * dblPorcentagemDif), 2)
                            'dblValorMulta = gstrConvVrDoSql((adoResultado("dblvalorMulta") * dblPorcentagemDif), 2)
                            'dblValorJuros = gstrConvVrDoSql((adoResultado("dblvalorJuros") * dblPorcentagemDif), 2)
                            'dblValorCorrecao = gstrConvVrDoSql((adoResultado("dblvalorCorrecao") * dblPorcentagemDif), 2)
                            'dblValorDesconto = gstrConvVrDoSql((adoResultado("dblvalorDesconto") * dblPorcentagemDif), 2)
                            'PROVISORIO
                            If dblPorcentagemDif > 0 Then
                                dblValorPrincipal = gstrConvVrDoSql((dblValorDoArquivo - dblPorcentagemDif), 2)
                                dblValorMulta = gstrConvVrDoSql(dblPorcentagemDif, 2)
                            Else
                                dblValorPrincipal = gstrConvVrDoSql((dblValorDoArquivo + dblPorcentagemDif), 2)
                                dblValorMulta = 0
                            End If
                            dblValorJuros = 0
                            dblValorCorrecao = 0
                            dblValorDesconto = 0
                        Else
                            dblValorPrincipal = (ConverteValorDoArquivo(Mid(strLinha, 254, 13)) + ConverteValorDoArquivo(Mid(strLinha, 267, 13)) + ConverteValorDoArquivo(Mid(strLinha, 280, 13))) - (ConverteValorDoArquivo(Mid(strLinha, 228, 13)) + ConverteValorDoArquivo(Mid(strLinha, 241, 13)))
                            dblValorMulta = adoResultado("dblValorMulta").Value
                            dblValorJuros = adoResultado("dblValorJuros").Value
                            dblValorCorrecao = adoResultado("dblValorCorrecao").Value
                            dblValorDesconto = adoResultado("dblValorDesconto").Value
                        End If
                        
                        dblSomaDosValores = dblSomaDosValores + (dblValorPrincipal + dblValorMulta + dblValorJuros + dblValorCorrecao + dblValorDesconto)
                        
                        'Se este for o ultimo registro, vamos verificar se ha diferenca dos valores calculados
                        If adoResultado.RecordCount = adoResultado.AbsolutePosition Then
                            'Caso depois do acerto de valores ainda exista diferenca, a jogaremos no ultimo registro
                            If CCur(dblSomaDosValores) <> CCur(dblValorDoArquivo) Then
                            
                                dblValorDiferenca = dblValorDoArquivo - dblSomaDosValores
                        
                                'strSQL = "UPDATE " & gstrMovimentoBancario
                                If dblValorPrincipal > 0 Then
                                    dblValorPrincipal = dblValorPrincipal + dblValorDiferenca
                                ElseIf dblValorMulta > 0 Then
                                    dblValorMulta = dblValorMulta + dblValorDiferenca
                                ElseIf dblValorJuros > 0 Then
                                    dblValorJuros = dblValorJuros + dblValorDiferenca
                                Else
                                    dblValorCorrecao = dblValorCorrecao + dblValorDiferenca
                                End If
                            End If
                        End If
                        
                        strSQL = gstrStoredProcedure("sp_AtualizaParcela", adoResultado("intComposicaoDaReceita").Value & ", " & adoResultado("intExercicio").Value & ", " & adoResultado("intLancamentoValor").Value & ", " & gstrConvDtParaSql(adoResultado("dtmDtVencimento").Value) & ", " & gstrConvDtParaSql(Mid(strLinha, 176, 2) & "/" & Mid(strLinha, 178, 2) & "/" & Mid(strLinha, 180, 2)) & ", " & gstrConvVrParaSql(adoResultado("ValorOrig").Value) & ", " & adoResultado("intMoeda").Value, True)
                        If gobjBanco.CriaADO(strSQL, 80, adoAtualizacao) Then
                            If gstrConvVrDoSql(adoResultado("ValorOrig").Value, , , True) = 0 Then
                                dblValorPrincipal = dblValorDoArquivo
                                dblValorCorreto = dblValorDoArquivo
                                dblValorMulta = 0
                            Else
                                dblValorCorreto = Space$(0) & gstrConvVrDoSql(Val(gstrConvVrParaSql(adoAtualizacao("dblValorPrincipal").Value)) + Val(gstrConvVrParaSql(adoAtualizacao("dblValorMulta").Value)) + Val(gstrConvVrParaSql(adoAtualizacao("dblValorJuros").Value)) + Val(gstrConvVrParaSql(adoAtualizacao("dblValorCorrecao").Value)))
                            End If
                        Else
                            Screen.MousePointer = vbDefault
                            pgr_Status.Visible = False
                            lbl_Status.Visible = False
                            gobjBanco.ExecutaRollbackTrans
                            Close #1
                            Exit Sub

                        End If
                        '*****************************************************
                        ' Data          : 20/12/2005                         *
                        ' Criação       : Verificação de Desconto de Tarifa  *
                        ' Responsável   : Fernando Peixoto                   *
                        ' Pendência     : Tri0518                            *
                        ''****************************************************
                        
                        strSqlAux = ""
                        strSqlAux = "SELECT " & gstrTOPnSQLServer(1) & " strCodBarra FROM " & gstrGuias & " WHERE strCodBarraEsp = '" & Mid(strLinha, 64, 17) & "' ORDER BY Pkid Desc"
                        strSqlAux = gstrTOPnOracle(strSqlAux, 1)
                        
                        Set gobjBanco = New clsBanco
                        If gobjBanco.CriaADO(strSqlAux, 10, adoAux) Then
                            If Not adoAux.EOF Then
                                strCodigoBarra = adoAux("strCodBarra")
                            Else
                                strCodigoBarra = ""
                            End If
                        End If
                        
                        If DesContarTarifa = False Then
                            gobjBanco.Execute "INSERT INTO " & gstrMovimentoBancario & "(dtmDtMovimento, intContaBancaria, intLote, dtmDtPagamento, dblPrincipal, dblMulta, dblJuros, dblCorrecao, dblTarifa, strCodigoDeBarras, dblCorreto, intLancamentoValor, intGuias, intCodigoBaixa, intDigito, BytTipo, dtmDtAtualizacao, lngCodUsr, bitGuia, bitProcessado) " & _
                                                                                        "VALUES(" & gstrConvDtParaSql(txt_DtMovimento.Text) & "," & lngPkidContaBancaria & "," & intLote & "," & gstrConvDtParaSql(Mid(strLinha, 176, 2) & "/" & Mid(strLinha, 178, 2) & "/" & Mid(strLinha, 180, 2)) & "," & gstrConvVrParaSql(dblValorPrincipal) & "," & gstrConvVrParaSql(dblValorMulta) & "," & gstrConvVrParaSql(dblValorJuros) & "," & gstrConvVrParaSql(dblValorCorrecao) & "," & 0 & ",'" & strCodigoBarra & "'," & gstrConvVrParaSql(dblValorCorreto) & "," & adoResultado("intLancamentoValor").Value & "," & adoResultado("intGuias").Value & "," & _
                                                                                        IIf(CDate(adoResultado("dtmDtVencimento").Value) < CDate(Mid(strLinha, 176, 2) & "/" & Mid(strLinha, 178, 2) & "/" & Mid(strLinha, 180, 2)), gstrENulo(intCodBaixaAposVcto, , True), gstrENulo(intCodBaixaAntesVcto, , True)) & "," & gstrCalculaDigitoModulo10(Trim(adoResultado("strNumeroAviso").Value) & Format$(Trim(adoResultado("intParcela").Value), "000")) & _
                                                                                        ",4" & "," & gstrConvDtParaSql(gstrDataDoSistema) & "," & glngCodUsr & ",1,0)"
                        Else
                            gobjBanco.Execute "INSERT INTO " & gstrMovimentoBancario & "(dtmDtMovimento, intContaBancaria, intLote, dtmDtPagamento, dblPrincipal, dblMulta, dblJuros, dblCorrecao, dblTarifa, strCodigoDeBarras, dblCorreto, intLancamentoValor, intGuias, intCodigoBaixa, intDigito, BytTipo, dtmDtAtualizacao, lngCodUsr, bitGuia, bitProcessado) " & _
                                                                                        "VALUES(" & gstrConvDtParaSql(txt_DtMovimento.Text) & "," & lngPkidContaBancaria & "," & intLote & "," & gstrConvDtParaSql(Mid(strLinha, 176, 2) & "/" & Mid(strLinha, 178, 2) & "/" & Mid(strLinha, 180, 2)) & "," & gstrConvVrParaSql(dblValorPrincipal - dblValorTarifa) & "," & gstrConvVrParaSql(dblValorMulta) & "," & gstrConvVrParaSql(dblValorJuros) & "," & gstrConvVrParaSql(dblValorCorrecao) & "," & gstrConvVrParaSql(dblValorTarifa) & ",'" & strCodigoBarra & "'," & gstrConvVrParaSql(dblValorCorreto) & "," & adoResultado("intLancamentoValor").Value & "," & adoResultado("intGuias").Value & "," & _
                                                                                        IIf(CDate(adoResultado("dtmDtVencimento").Value) < CDate(Mid(strLinha, 176, 2) & "/" & Mid(strLinha, 178, 2) & "/" & Mid(strLinha, 180, 2)), gstrENulo(intCodBaixaAposVcto, , True), gstrENulo(intCodBaixaAntesVcto, , True)) & "," & gstrCalculaDigitoModulo10(Trim(adoResultado("strNumeroAviso").Value) & Format$(Trim(adoResultado("intParcela").Value), "000")) & _
                                                                                        ",4" & "," & gstrConvDtParaSql(gstrDataDoSistema) & "," & glngCodUsr & ",1,0)"
                        End If
                        
                        adoResultado.MoveNext
                        
                    Loop
                    
                Else
                    
                   dblValorPrincipal = dblValorDoArquivo
                   dblValorMulta = 0
                   dblValorJuros = 0
                   dblValorCorrecao = 0
                   dblValorDesconto = 0
                        
                    '*****************************************************
                    ' Data          : 20/12/2005                         *
                    ' Criação       : Verificação de Desconto de Tarifa  *
                    ' Responsável   : Fernando Peixoto                   *
                    ' Pendência     : Tri0518                            *
                    ''****************************************************
                    
                    If DesContarTarifa = False Then
                       gobjBanco.Execute "INSERT INTO " & gstrMovimentoBancario & "(dtmDtMovimento, intContaBancaria, intLote, dtmDtPagamento, dblPrincipal, dblMulta, dblJuros, dblCorrecao, dblTarifa, strCodigoDeBarras, dblCorreto, intLancamentoValor, intGuias, intCodigoBaixa, intDigito, BytTipo, dtmDtAtualizacao, lngCodUsr, bitGuia, bitProcessado) " & _
                                                                                        "VALUES(" & gstrConvDtParaSql(txt_DtMovimento.Text) & "," & lngPkidContaBancaria & "," & intLote & "," & gstrConvDtParaSql(Mid(strLinha, 176, 2) & "/" & Mid(strLinha, 178, 2) & "/" & Mid(strLinha, 180, 2)) & "," & gstrConvVrParaSql(dblValorPrincipal) & "," & gstrConvVrParaSql(dblValorMulta) & "," & gstrConvVrParaSql(dblValorJuros) & "," & gstrConvVrParaSql(dblValorCorrecao) & "," & 0 & ",'" & Mid(strLinha, 64, 17) & "',0,NULL,NULL,NULL,0," & _
                                                                                        "4," & gstrConvDtParaSql(gstrDataDoSistema) & "," & glngCodUsr & ",0,0)"
                    Else
                       gobjBanco.Execute "INSERT INTO " & gstrMovimentoBancario & "(dtmDtMovimento, intContaBancaria, intLote, dtmDtPagamento, dblPrincipal, dblMulta, dblJuros, dblCorrecao, dblTarifa, strCodigoDeBarras, dblCorreto, intLancamentoValor, intGuias, intCodigoBaixa, intDigito, BytTipo, dtmDtAtualizacao, lngCodUsr, bitGuia, bitProcessado) " & _
                                                                                        "VALUES(" & gstrConvDtParaSql(txt_DtMovimento.Text) & "," & lngPkidContaBancaria & "," & intLote & "," & gstrConvDtParaSql(Mid(strLinha, 176, 2) & "/" & Mid(strLinha, 178, 2) & "/" & Mid(strLinha, 180, 2)) & "," & gstrConvVrParaSql(dblValorPrincipal - dblValorTarifa) & "," & gstrConvVrParaSql(dblValorMulta) & "," & gstrConvVrParaSql(dblValorJuros) & "," & gstrConvVrParaSql(dblValorCorrecao) & "," & gstrConvVrParaSql(dblValorTarifa) & ",'" & Mid(strLinha, 64, 17) & "',0,NULL,NULL,NULL,0," & _
                                                                                        "4," & gstrConvDtParaSql(gstrDataDoSistema) & "," & glngCodUsr & ",0,0)"
                    End If
                End If
                
           End If
             
        ElseIf Mid(strLinha, 1, 1) = "9" Then
            bytQtdeTipos = bytQtdeTipos + 1
                         
            'Vamos verificar se ja existe um registro referente, caso exista vamos somar o valor
            strSQL = "SELECT Pkid, dblValor FROM " & gstrResumoBancario & " WHERE dtmData = " & gstrConvDtParaSql(txt_DtMovimento.Text) & " AND intLote = " & intLote & " AND intContaBancaria = " & lngPkidContaBancaria
            If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
                If Not adoResultado.EOF Then
             
                    'Vamos somar registro na tabela tblResumoBancario
                    strSQL = "UPDATE " & gstrResumoBancario & " SET dblValor = " & gstrConvVrParaSql(dblValorTotDoArquivo + adoResultado("dblValor").Value) & " WHERE Pkid = " & adoResultado("Pkid").Value
                     
                Else
                
                    'Vamos gravar registro na tabela tblResumoBancario
                    strSQL = "INSERT INTO " & gstrResumoBancario & "(dtmData, intContaBancaria, intLote, dblValor, dtmDtAtualizacao, lngCodUsr) " & _
                             "VALUES(" & gstrConvDtParaSql(txt_DtMovimento.Text) & "," & lngPkidContaBancaria & "," & intLote & "," & gstrConvVrParaSql(dblValorTotDoArquivo) & "," & gstrConvDtParaSql(gstrDataDoSistema) & "," & glngCodUsr & ")"
                
                End If
            End If
            
            gobjBanco.Execute strSQL
            
            'Vamos verificar se a qtde gravada é igual ao total de registros
            If lngLinha + 1 <> Val(Mid(strLinha, 395, 6)) Then
                ExibeMensagem "A quantidade de registros processada não coincide com a informada no arquivo. A operação não concluída."
                Screen.MousePointer = vbDefault
                pgr_Status.Visible = False
                lbl_Status.Visible = False
                gobjBanco.ExecutaRollbackTrans
                Close #1
                Exit Sub
            End If
            
        End If
        
        lngLinha = lngLinha + 1
        
        DoEvents
        lbl_Status.Caption = lngLinha & " de " & lngSize
        Me.Refresh
        
        pgr_Status.Value = lngLinha
        
ProximaLinha:

    Loop
        
    Close #1
    
    pgr_Status.Visible = False
    lbl_Status.Visible = False
        
    'Vamos verificar se foram lidos todos os tipos de registro
    If bytQtdeTipos < 3 Then
        ExibeMensagem "O arquivo de leitura está incorreto."
        Screen.MousePointer = vbDefault
        gobjBanco.ExecutaRollbackTrans
        Close #1
        Exit Sub
    End If
    
    gobjBanco.ExecutaCommitTrans
    
    Screen.MousePointer = vbDefault
    
    Exit Sub

err_BaixaAutomatica:
    ExibeMensagem "Ocorreu um Erro na leitura do arquivo. A operação não foi concluída."
    Screen.MousePointer = vbDefault
    pgr_Status.Visible = False
    lbl_Status.Visible = False
    gobjBanco.ExecutaRollbackTrans
    Close #1
End Sub
