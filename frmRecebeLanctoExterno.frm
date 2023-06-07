VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmRecebeLanctoExterno 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Receber Lançamento Externo"
   ClientHeight    =   1995
   ClientLeft      =   2475
   ClientTop       =   2685
   ClientWidth     =   6240
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1995
   ScaleWidth      =   6240
   Begin VB.Frame fra_ComposicaoDaReceita 
      Caption         =   "Composição da Receita"
      Height          =   750
      Left            =   90
      TabIndex        =   5
      Top             =   0
      Width           =   6015
      Begin VB.CommandButton cmd_Composicao 
         Height          =   300
         Left            =   5460
         Picture         =   "frmRecebeLanctoExterno.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   6
         TabStop         =   0   'False
         Tag             =   "590"
         ToolTipText     =   "Ativa Cadastro de Composição da Receita"
         Top             =   270
         Width           =   360
      End
      Begin MSDataListLib.DataCombo dbc_intComposicao 
         Height          =   315
         Left            =   1110
         TabIndex        =   7
         Top             =   270
         Width           =   4320
         _ExtentX        =   7620
         _ExtentY        =   556
         _Version        =   393216
         IntegralHeight  =   0   'False
         Text            =   ""
      End
      Begin VB.Label lbl_Composicao 
         AutoSize        =   -1  'True
         Caption         =   "Composição"
         Height          =   195
         Left            =   180
         TabIndex        =   8
         Top             =   315
         Width           =   870
      End
   End
   Begin VB.Frame fra_Arquivo 
      Caption         =   " Arquivo de leitura "
      Height          =   810
      Left            =   105
      TabIndex        =   1
      Top             =   810
      Width           =   6000
      Begin VB.CommandButton cmd_Arquivo 
         Caption         =   "..."
         Height          =   300
         Left            =   5460
         Picture         =   "frmRecebeLanctoExterno.frx":011E
         TabIndex        =   3
         TabStop         =   0   'False
         ToolTipText     =   "Localiza Arquivo de Baixa Automática"
         Top             =   315
         Width           =   345
      End
      Begin VB.TextBox txt_Arquivo 
         Height          =   285
         Left            =   1095
         TabIndex        =   2
         Top             =   315
         Width           =   4335
      End
      Begin VB.Label lbl_Arquivo 
         AutoSize        =   -1  'True
         Caption         =   "Localização"
         Height          =   195
         Left            =   180
         TabIndex        =   4
         Top             =   360
         Width           =   855
      End
   End
   Begin MSComctlLib.ProgressBar pgr_Status 
      Height          =   165
      Left            =   105
      TabIndex        =   0
      Top             =   1710
      Visible         =   0   'False
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin MSComDlg.CommonDialog dlgArquivo 
      Left            =   0
      Top             =   45
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmRecebeLanctoExterno"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mblnPrimeiraVez  As Boolean
Dim strUltimoCaminho As String

Dim blnCriticado     As Boolean

Private Sub cmd_Arquivo_Click()
    
    dlgArquivo.CancelError = True
    dlgArquivo.DialogTitle = "Selecione o arquivo"
    dlgArquivo.InitDir = strUltimoCaminho
    dlgArquivo.Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
    dlgArquivo.flags = cdlOFNOverwritePrompt + cdlOFNHideReadOnly
    
    On Error GoTo err_cmd_Arquivo_Click
    
    dlgArquivo.ShowOpen
    txt_Arquivo = dlgArquivo.filename
    strUltimoCaminho = Replace(dlgArquivo.filename, dlgArquivo.FileTitle, "")
    Exit Sub

err_cmd_Arquivo_Click:
    If Err.Number = 32755 Then
        txt_Arquivo = ""
    End If
    
End Sub

Private Sub cmd_Composicao_Click()
    ChamaFormCadastro frmCadComposicaoDaReceita, dbc_intComposicao
End Sub

Private Sub dbc_intComposicao_Click(Area As Integer)
    If Area = 0 Then DropDownDataCombo dbc_intComposicao, Me, Area
End Sub

Private Sub dbc_intComposicao_GotFocus()
    MarcaCampo dbc_intComposicao
End Sub

Private Sub dbc_intComposicao_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbc_intComposicao, Me, , KeyCode, Shift
End Sub

Private Sub dbc_intComposicao_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbc_intComposicao
End Sub

Private Sub Form_Activate()
    gintCodSeguranca = 1291
    HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrLerArquivo
End Sub

Private Sub Form_Deactivate()
    HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrLerArquivo
End Sub

Private Sub Form_Load()
    strUltimoCaminho = "C:\"
    
    dbc_intComposicao.Tag = strQueryComposicao & ";strDescricao"
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrLerArquivo
    mblnPrimeiraVez = False
End Sub

Private Sub LeMovimentoBancario()
Dim strLinha               As String
Dim strLinhaAnterior       As String
Dim lngLinha               As Long
Dim lngSize                As Long

Dim strsql                 As String

Dim STRMUNICIPIO           As String
Dim STRUF                  As String
Dim lngMoedaAtual          As Long

Dim strInscricaoCorrente   As String

Dim lngPkidLancamentoAlfa  As Long
Dim lngPkidLancamentoValor As Long
Dim lngPkidGuia            As Long

Dim adoResultado           As New ADODB.Recordset

Dim vetReceitas()          As String
Dim lngReceita             As Long
Dim intFor                 As Integer

Dim blnProximaInscricao    As Boolean
Dim blnParcelaAtualizada   As Boolean
Dim varRegistroCriticas    As Variant

    On Error GoTo err_BaixaAutomatica
        
    If Not blnDadosOk Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    
    blnCriticado = False
    
    lngLinha = 0
    
    Open txt_Arquivo For Input As #1
    
    lngSize = FileLen(txt_Arquivo) / 150
    
    pgr_Status.Value = 0
    pgr_Status.Visible = True
    pgr_Status.Max = Abs(lngSize)
    
    Set gobjBanco = New clsBanco
    
    If gobjBanco.CriaADO("SELECT E.intMoeda, M.strDescricao Cidade, u.strSigla UF from tblempresa E, tblmunicipio M, tblUf u Where M.pkid " & strOUTJOracle & " =" & strOUTJSQLServer & " E.intcidade and U.pkid " & strOUTJOracle & " =" & strOUTJSQLServer & " E.intUf", 30, adoResultado) Then
        If Not adoResultado.EOF Then
            STRMUNICIPIO = adoResultado("Cidade")
            STRUF = adoResultado("UF")
            lngMoedaAtual = adoResultado("intMoeda")
        Else
            STRMUNICIPIO = ""
            STRUF = ""
            lngMoedaAtual = Null
        End If
    End If
    
    Line Input #1, strLinha
    
    Do While Not EOF(1)

        If Len(strLinha) = 0 Then
            lngLinha = lngLinha + 1
            pgr_Status.Value = lngLinha
            Line Input #1, strLinha
            GoTo ProximaLinha
        End If
        
        If Mid(strLinha, 1, 2) = "50" Then
            lngLinha = lngLinha + 1
            pgr_Status.Value = lngLinha
            Line Input #1, strLinha
            GoTo ProximaLinha
        End If
   
        gobjBanco.ExecutaBeginTrans

        ReDim vetReceitas(1, 0)
        
        strInscricaoCorrente = Trim(Mid(strLinha, 3, 15))
        
        '**************** LANCAMENTO ALFA ************************
        'Vamos verificar se ja existe em LancamentoAlfa
        'gobjBanco.CriaADO "SELECT Pkid FROM " & gstrLancamentoAlfa & " WHERE strInscricao = '" & (String(gintLenInscricao - Len(Trim(Mid(strLinha, 3, 15))), "0") & Mid(Trim(Mid(strLinha, 3, 15)), 1, Len(Trim(Mid(strLinha, 3, 15))) - 1)) & "'", 10, adoResultado
        
        'If adoResultado.RecordCount = 0 Then
            
            'Tipo 51
            If Mid(strLinha, 1, 2) = "51" Then
        
                'Vamos gravar a tabela LancamentoAlfa
                strsql = "INSERT INTO " & gstrLancamentoAlfa & "(Strinscricao, Strcomposicaodareceita, Strocorrencia, Strnomeproprietario, Strcnpjcpf," & _
                                                "Stridentidade, Strlogradouro, Strnumero, Strcomplemento, Strbairro, Strmunicipio," & _
                                                "Struf, Intcep, Strlogradouroc, Strnumeroc, Strcomplementoc, Strbairroc, Strmunicipioc," & _
                                                "Strufc, Intcepc, Strnumeroaviso, Stremissao, Intexercicio, Intcomposicaodareceita, intUtilizacao, dtmDtAtualizacao, lngCodUsr, bytnaoinscreveda)"
        
                strsql = strsql & "SELECT EC.STRINSCRICAOCADASTRAL, '" & dbc_intComposicao.Text & "', OC.STRDESCRICAO, CO.STRNOME, CO.STRCNPJCPF," & _
                         "CO.stridentidade, TPL.STRSIGLA " & strCONCAT & "' '" & strCONCAT & " TTL.STRdescricao " & strCONCAT & "' '" & strCONCAT & " LO.STRDESCRICAO, EC.INTNUMERO, EC.Strcomplemento, Ba.Strdescricao, '" & STRMUNICIPIO & "', '" & _
                         STRUF & "', EC.Intcep, CO.strlogradouroc, CO.Intnumeroc, CO.Strcomplementoc, CO.Strbairroc, MU.STRDESCRICAO, " & _
                         "UF.strsigla , CO.Intcepc, '" & Mid(strLinha, 28, 6) & "', '999', " & Mid(strLinha, 18, 4) & ", " & dbc_intComposicao.BoundText & ", CR.intUtilizacao, " & gstrConvDtParaSql(gstrDataDoSistema) & ", " & glngCodUsr & ", 0"
                strsql = strsql & " FROM tblEconomico EC, Tblocorrencia OC, tblcontribuinte CO, tbllogradouro LO, " & _
                         "tbltipologradouro TPL, tbltitulologradouro TTL, tblbairro BA, tblmunicipio MU, tbluf UF, tblComposicaoDaReceita CR "
                strsql = strsql & " WHERE EC.STRINSCRICAOCADASTRAL = '" & (String(gintLenInscricao - Len(Trim(Mid(strLinha, 3, 15))) + 1, "0") & Mid(Trim(Mid(strLinha, 3, 15)), 1, Len(Trim(Mid(strLinha, 3, 15))) - 1)) & "' and " & _
                         "EC.Intocorrencia = OC.Pkid and " & _
                         "EC.INTCONTRIBUINTE = CO.PKID and " & _
                         "Lo.Pkid " & strOUTJOracle & " =" & strOUTJSQLServer & " EC.Intlogradouro and " & _
                         "TPL.PKID " & strOUTJOracle & " =" & strOUTJSQLServer & " LO.INTTIPOLOGRADOURO and " & _
                         "TTL.Pkid " & strOUTJOracle & " =" & strOUTJSQLServer & " LO.INTTITULOLOGRADOURO and " & _
                         "BA.PKID " & strOUTJOracle & " =" & strOUTJSQLServer & " EC.Intbairro and " & _
                         "MU.PKID " & strOUTJOracle & " =" & strOUTJSQLServer & " CO.Intmunicipioc and " & _
                         "UF.PKID " & strOUTJOracle & " =" & strOUTJSQLServer & " Co.Intufc and " & _
                         "CR.Pkid = " & dbc_intComposicao.BoundText
            
                If Not gobjBanco.Execute(strsql, False) Then
                    'ExibeMensagem "Não foi possível criar o Lançamento Alfa. A operação foi cancelada."
                    gobjBanco.ExecutaRollbackTrans
                    
                    'Vamos pular para a proxima inscricao
                    blnProximaInscricao = False
                    Do While Not blnProximaInscricao
                        
                        GeraCriticas strLinha & " Não foi possível criar o LancamentoAlfa."
                        
                        lngLinha = lngLinha + 1
                        pgr_Status.Value = lngLinha
                        Line Input #1, strLinha
                    
                        blnProximaInscricao = strInscricaoCorrente <> Trim(Mid(strLinha, 3, 15))
                        
                    Loop
                    
                    GoTo ProximaLinha
                    
                End If
            
                lngPkidLancamentoAlfa = glngRetornaPkidTabelaPai("seqtblLancamentoAlfa", gstrLancamentoAlfa)
                
            End If
            
        'Else
        '    lngPkidLancamentoAlfa = adoResultado("Pkid").Value
        'End If
        
        '**************** LANCAMENTO ECONOMICO ************************
        
        'Tipo 51
        If Mid(strLinha, 1, 2) = "51" Then
        
            strsql = "INSERT INTO " & gstrLancamentoEconomico & " (intLancamentoAlfa, strNomeFantasia, strInscricaoEstadual, strAtividadeBasica, strNaturezaJuridica, " & _
                                        "dtmDataAbertura, dblAreaOcupada, blnMicroEmpresa, dblNumeroEmpregados, dtmDtAtualizacao, lngCodUsr) "
            strsql = strsql & "(SELECT " & lngPkidLancamentoAlfa & " , CO.STRNOMEFANTASIA, CO.STRINSCRICAOESTADUAL, AB.STRDESCRICAO, CASE CO.Bytnaturezajuridica WHEN 0 THEN 'Fisica' WHEN 1 THEN 'Juridica' WHEN 2 THEN 'SC' ELSE 'Outros' END, " & _
                              " EC.Dtmdataabertura , EC.dblAreaOcupada, EC.Blnmicroempresa, EC.Intnumdeempregados, " & gstrConvDtParaSql(gstrDataDoSistema) & ", " & glngCodUsr
            strsql = strsql & " FROM Tbleconomico EC, Tblcontribuinte CO, Tblatividadebasica AB "
            strsql = strsql & " WHERE CO.Pkid (+) = EC.INTCONTRIBUINTE and " & _
                              " AB.Pkid (+) = EC.INTATIVIDADEBASICA and " & _
                              " EC.STRINSCRICAOCADASTRAL = '" & (String(gintLenInscricao - Len(Trim(Mid(strLinha, 3, 15))) + 1, "0") & Mid(Trim(Mid(strLinha, 3, 15)), 1, Len(Trim(Mid(strLinha, 3, 15))) - 1)) & "')"
        
            If Not gobjBanco.Execute(strsql, False) Then
                
                gobjBanco.ExecutaRollbackTrans
                'ExibeMensagem "Não foi possível criar o Lançamento Economico. A operação foi cancelada."
                
                'Vamos pular para a proxima inscricao
                blnProximaInscricao = False
                Do While Not blnProximaInscricao
                    
                    GeraCriticas strLinha & " Não foi possível criar o LancamentoEconomico."
                    
                    lngLinha = lngLinha + 1
                    pgr_Status.Value = lngLinha
                    Line Input #1, strLinha
                
                    blnProximaInscricao = strInscricaoCorrente <> Trim(Mid(strLinha, 3, 15))
                    
                Loop
                
                GoTo ProximaLinha

            End If
        
            '**************** LANCAMENTO VALOR ************************
        
            'Vamos gravar a parcela na tabela tblLancamentoValor
            strsql = "INSERT INTO " & gstrLancamentoValor & " " & _
                     "(intLancamentoAlfa, intParcela, dtmDtVencimento, dblValor, intMoeda, bitParcelaValida, dtmDtAtualizacao, lngCodUsr)" & _
                     " VALUES " & _
                     "(" & lngPkidLancamentoAlfa & "," & Mid(strLinha, 34, 2) & " ," & gstrConvDtParaSql(Mid(strLinha, 63, 2) & "/" & Mid(strLinha, 61, 2) & "/" & Mid(strLinha, 57, 4)) & "," & gstrConvVrParaSql(Mid(strLinha, 44, 11) & "," & Mid(strLinha, 55, 2)) & "," & lngMoedaAtual & ",1," & gstrConvDtParaSql(gstrDataDoSistema) & "," & glngCodUsr & ")"
                                      
            If Not gobjBanco.Execute(strsql, False) Then
                
                gobjBanco.ExecutaRollbackTrans
                'ExibeMensagem "Não foi possível criar o Lançamento Valor. A operação foi cancelada."
                
                'Vamos pular para a proxima inscricao
                blnProximaInscricao = False
                Do While Not blnProximaInscricao
                    
                    GeraCriticas strLinha & " Não foi possível criar o LancamentoValor."
                    
                    lngLinha = lngLinha + 1
                    pgr_Status.Value = lngLinha
                    Line Input #1, strLinha
                
                    blnProximaInscricao = strInscricaoCorrente <> Trim(Mid(strLinha, 3, 15))
                    
                Loop
                
                GoTo ProximaLinha

            End If
                
            lngPkidLancamentoValor = glngRetornaPkidTabelaPai("seqtblLancamentoValor", gstrLancamentoValor)
        
            '**************** GUIAS ************************
            strsql = "INSERT INTO " & gstrGuias & "("
            strsql = strsql & "intContaBancaria, "
            strsql = strsql & "intNumero, "
            strsql = strsql & "dtmdtEmissao, "
            strsql = strsql & "dblValor, "
            strsql = strsql & "strCodBarra, "
            strsql = strsql & "dtmdtAtualizacao, "
            strsql = strsql & "lngCodUsr, "
            strsql = strsql & "dtmdtVencimento "
            strsql = strsql & ") VALUES ("
            strsql = strsql & "NULL, "
            strsql = strsql & Mid(strLinha, 28, 6) & ", "
            strsql = strsql & gstrConvDtParaSql(Mid(strLinha, 42, 2) & "/" & Mid(strLinha, 40, 2) & "/" & Mid(strLinha, 36, 4)) & ", "
            strsql = strsql & gstrConvVrParaSql(Mid(strLinha, 44, 11) & "," & Mid(strLinha, 55, 2)) & ", '"
            strsql = strsql & Mid(strLinha, 75, 44) & "', "
            strsql = strsql & gstrConvDtParaSql(gstrDataDoSistema) & ", "
            strsql = strsql & glngCodUsr & ", "
            strsql = strsql & gstrConvDtParaSql(Mid(strLinha, 63, 2) & "/" & Mid(strLinha, 61, 2) & "/" & Mid(strLinha, 57, 4))
            strsql = strsql & ")"
      
            If Not gobjBanco.Execute(strsql, False) Then
                
                gobjBanco.ExecutaRollbackTrans
                                
                'Vamos pular para a proxima inscricao
                blnProximaInscricao = False
                Do While Not blnProximaInscricao
                    
                    GeraCriticas strLinha & " Não foi possível criar a Guia."
                    
                    lngLinha = lngLinha + 1
                    pgr_Status.Value = lngLinha
                    Line Input #1, strLinha
                
                    blnProximaInscricao = strInscricaoCorrente <> Trim(Mid(strLinha, 3, 15))
                    
                Loop
                
                GoTo ProximaLinha

            End If
                
            lngPkidGuia = glngRetornaPkidTabelaPai("seqtblGuias", gstrGuias)
        
            '**************** LANCAMENTO GUIAS ************************
            strsql = "INSERT INTO " & gstrLancamentoGuias & "("
            strsql = strsql & "intlancamentovalor, "
            strsql = strsql & "intguias, "
            strsql = strsql & "dblvalorprincipal, "
            strsql = strsql & "dblvalormulta, "
            strsql = strsql & "dblvalorjuros, "
            strsql = strsql & "dblvalorcorrecao, "
            strsql = strsql & "dblvalordesconto, "
            strsql = strsql & "dtmdtatualizacao, "
            strsql = strsql & "lngcodusr) "
            strsql = strsql & "Values ("
            strsql = strsql & lngPkidLancamentoValor & ", "
            strsql = strsql & lngPkidGuia & ", "
            strsql = strsql & gstrConvVrParaSql(Mid(strLinha, 44, 11) & "," & Mid(strLinha, 55, 2)) & ", "
            strsql = strsql & "0, "
            strsql = strsql & "0, "
            strsql = strsql & "0, "
            strsql = strsql & "0, "
            strsql = strsql & gstrConvDtParaSql(gstrDataDoSistema) & ", "
            strsql = strsql & glngCodUsr & ") "
        
            If Not gobjBanco.Execute(strsql, False) Then
                
                gobjBanco.ExecutaRollbackTrans
                
                'Vamos pular para a proxima inscricao
                blnProximaInscricao = False
                Do While Not blnProximaInscricao
                    
                    GeraCriticas strLinha & " Não foi possível criar o LancamentoGuia."
                    
                    lngLinha = lngLinha + 1
                    pgr_Status.Value = lngLinha
                    Line Input #1, strLinha
                
                    blnProximaInscricao = strInscricaoCorrente <> Trim(Mid(strLinha, 3, 15))
                    
                Loop
                
                GoTo ProximaLinha

            End If
        
        End If
        
        'Vamos armazenar o registro da inscricao do tipo 51, para o caso de critica
        varRegistroCriticas = strLinha & Chr(13) & Chr(10)
        
ProximaReceita:

        lngLinha = lngLinha + 1
        
        pgr_Status.Value = lngLinha
        
        'Vamos verificar se nao existe o registro tipo 52
        If Mid(strLinha, 1, 2) = "51" Then
        
            strLinhaAnterior = strLinha
            
            'Vamos para a proxima linha e verificar se é do tipo 52
            Line Input #1, strLinha
            
            If Mid(strLinha, 1, 2) = "51" Then
                
                gobjBanco.ExecutaRollbackTrans
                
                'Vamos pular para a proxima inscricao
                GeraCriticas strLinhaAnterior & " Não foi encontrado registro do tipo 52."
               
                blnProximaInscricao = strInscricaoCorrente <> Trim(Mid(strLinha, 3, 15))
                
                GoTo ProximaLinha
            
            End If
            
        Else
            'Vamos para a proxima linha e verificar se é do tipo 52
            Line Input #1, strLinha
        End If
        
        '**************** LANCAMENTO RECEITA ************************
        
        'Tipo 52
        If Mid(strLinha, 1, 2) = "52" And strInscricaoCorrente = Trim(Mid(strLinha, 3, 15)) Then
            
            If Not blnParcelaAtualizada Then
                
                'Vamos atualizar a parcela na tabela tblLancamentoValor com a data de vencimento do tipo 52
                strsql = "UPDATE " & gstrLancamentoValor & " " & _
                         "SET dtmDtVencimento = " & gstrConvDtParaSql(Mid(strLinha, 111, 2) & "/" & Mid(strLinha, 109, 2) & "/" & Mid(strLinha, 105, 4)) & _
                         " WHERE Pkid = " & lngPkidLancamentoValor
                
                If gobjBanco.Execute(strsql, False) Then
                    blnParcelaAtualizada = True
                End If

            End If
                        
            'Vamos obter o Pkid da Receita
            Set gobjBanco = New clsBanco
            
            If gobjBanco.CriaADO("SELECT I.intReceita FROM tblIntegraLancamentoExterno I WHERE I.intCodExterno = " & Mid(strLinha, 113, 12), 10, adoResultado) Then
                If Not adoResultado.EOF Then
                    lngReceita = adoResultado("intReceita")
                Else
                    
                    gobjBanco.ExecutaRollbackTrans
                    
                    'Vamos pular para a proxima inscricao
                    blnProximaInscricao = False
                    Do While Not blnProximaInscricao
                        
                        GeraCriticas strLinha & " Não foi possível retornar a Receita."
                        
                        lngLinha = lngLinha + 1
                        pgr_Status.Value = lngLinha
                        Line Input #1, strLinha
                    
                        blnProximaInscricao = strInscricaoCorrente <> Trim(Mid(strLinha, 3, 15))
                        
                    Loop
                    
                    GoTo ProximaLinha
                
                End If
            End If
            
            'Vamos agrupar as receitas
            For intFor = 0 To UBound(vetReceitas, 2)
                If Val(vetReceitas(0, intFor)) = lngReceita Then
                    vetReceitas(1, intFor) = vetReceitas(1, intFor) + CCur(Mid(strLinha, 92, 11) & "," & Mid(strLinha, 103, 2))
                    Exit For
                End If
                'Caso ainda nao exista no array vamos criar
                If intFor = UBound(vetReceitas, 2) Then
                    'Caso nao seja a primeira passagem, vamos criar um novo index no array
                    If vetReceitas(1, 0) <> "" Then ReDim Preserve vetReceitas(1, intFor + 1)
                    vetReceitas(0, UBound(vetReceitas, 2)) = lngReceita
                    vetReceitas(1, UBound(vetReceitas, 2)) = CCur(Mid(strLinha, 92, 11) & "," & Mid(strLinha, 103, 2))
                End If
            Next
            
            'Vamos armazenar todos os registros da inscricao do tipo 52, para o caso de critica
            varRegistroCriticas = varRegistroCriticas & strLinha & Chr(13) & Chr(10)

            GoTo ProximaReceita
            
        Else
            
            'Vamos gravar as receitas agrupadas no array na tabela de lancamento receitas
            For intFor = 0 To UBound(vetReceitas, 2)
            
                strsql = "INSERT INTO " & gstrLancamentoReceita & " " & _
                         "(intLancamentoValor, intReceita, dblValor, dtmDtAtualizacao, lngCodUsr)" & _
                         " VALUES " & _
                         "(" & lngPkidLancamentoValor & "," & vetReceitas(0, intFor) & "," & gstrConvVrParaSql(vetReceitas(1, intFor)) & "," & gstrConvDtParaSql(gstrDataDoSistema) & "," & glngCodUsr & ")"
            
                If Not gobjBanco.Execute(strsql, False) Then
                    
                    gobjBanco.ExecutaRollbackTrans
                    
                    'Vamos pular para a proxima inscricao
                    GeraCriticas varRegistroCriticas & strLinha & " Não foi possível criar o LançamentoReceita."
                        
                    varRegistroCriticas = Space$(0)
                    
                    GoTo ProximaLinha
                    
                End If
            
            Next
            
        End If

        gobjBanco.ExecutaCommitTrans
        
ProximaLinha:
    
        varRegistroCriticas = Space$(0)
        blnParcelaAtualizada = False
    
    Loop
        
    Close #1
    
    pgr_Status.Visible = False
    
    Screen.MousePointer = vbDefault
    
    If blnCriticado Then Close #2
    
    Exit Sub

err_BaixaAutomatica:
    ExibeMensagem "Ocorreu um Erro na leitura do arquivo. A operação não foi concluída."
    Resume Next
    Screen.MousePointer = vbDefault
    
    pgr_Status.Visible = False
    gobjBanco.ExecutaRollbackTrans
    Close #1
    If blnCriticado Then Close #2
    
End Sub

Private Function blnDadosOk() As Boolean
    On Error GoTo err_blnDadosOK
    
    If Not dbc_intComposicao.MatchedWithList Then
        ExibeMensagem "Selecione uma Composição da Receita válida."
        Exit Function
    ElseIf Trim(txt_Arquivo) = "" Then
        ExibeMensagem "Indique a localização do arquivo de retorno."
        Exit Function
    ElseIf Dir(txt_Arquivo) = "" Then
        ExibeMensagem "Arquivo não encontrado no local especificado."
        Exit Function
    End If
    
    blnDadosOk = True
    Exit Function
    
err_blnDadosOK:
    blnDadosOk = False
End Function

Public Sub MantemForm(ByVal strModoOperacao As String)
    
    Select Case strModoOperacao
        
        Case UCase(gstrPreencherLista)
            PreencherListaDeOpcoes Me.ActiveControl
        
        Case UCase(gstrLerArquivo)
            LeMovimentoBancario
            
        Case UCase(gstrFechar)
            Unload Me
    
    End Select

End Sub

Private Function strQueryComposicao() As String
Dim strsql As String

    strsql = "SELECT Pkid,"
    'strSql = strSql & gstrCONVERT(CDT_VARCHAR, "intCodigo") & strCONCAT & "' - '" & strCONCAT & " strDescricao Descricao"
    strsql = strsql & " strDescricao Descricao"
    strsql = strsql & " FROM "
    strsql = strsql & gstrComposicaoDaReceita
    strsql = strsql & " WHERE"
    strsql = strsql & " intUtilizacao = " & TYP_ECONOMICA
    strsql = strsql & " ORDER BY intCodigo"

    strQueryComposicao = strsql

End Function

Private Sub GeraCriticas(strLinha As String)
    
    If Not blnCriticado Then Open Mid(txt_Arquivo, 1, Len(txt_Arquivo) - 4) & "Critica.txt" For Output Access Write As #2
    
    Print #2, strLinha
    
    blnCriticado = True
    
End Sub
