VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MsDatLst.ocx"
Begin VB.Form frmCadDividaAtivaComposicao 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Inscrição Divida Ativa por Composição e Exercício"
   ClientHeight    =   2475
   ClientLeft      =   2505
   ClientTop       =   2940
   ClientWidth     =   9810
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2475
   ScaleWidth      =   9810
   Begin TabDlg.SSTab tab_3dPasta 
      Height          =   2430
      Left            =   30
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   30
      Width           =   9750
      _ExtentX        =   17198
      _ExtentY        =   4286
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Dívida Ativa"
      TabPicture(0)   =   "frmCadDividaAtivaComposicao.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Fra_Titulo"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin VB.Frame Fra_Titulo 
         Height          =   1965
         Left            =   150
         TabIndex        =   1
         Top             =   360
         Width           =   9480
         Begin VB.CheckBox chk_InscricaoAcordo 
            Caption         =   "Exibir inscrições em acordo"
            Height          =   195
            Left            =   1890
            TabIndex        =   18
            Top             =   1350
            Width           =   2865
         End
         Begin MSComctlLib.ProgressBar prgImportacaoDativa 
            Height          =   225
            Left            =   1890
            TabIndex        =   19
            Top             =   1650
            Visible         =   0   'False
            Width           =   7455
            _ExtentX        =   13150
            _ExtentY        =   397
            _Version        =   393216
            Appearance      =   1
            Scrolling       =   1
         End
         Begin VB.TextBox txtdtmdtinscricao 
            Height          =   285
            Left            =   1890
            TabIndex        =   10
            Top             =   660
            Width           =   1005
         End
         Begin VB.CommandButton cmd_TabelaComposicaoDaReceita 
            Height          =   315
            Left            =   5325
            Picture         =   "frmCadDividaAtivaComposicao.frx":001C
            Style           =   1  'Graphical
            TabIndex        =   4
            TabStop         =   0   'False
            Tag             =   "617"
            ToolTipText     =   "Ativa Cadastro de Composição Da Receita"
            Top             =   255
            Width           =   360
         End
         Begin VB.TextBox txtintcertidao 
            Height          =   285
            Left            =   4335
            TabIndex        =   12
            Top             =   675
            Width           =   1365
         End
         Begin VB.TextBox txtcadastro 
            Height          =   285
            Left            =   6480
            TabIndex        =   6
            Top             =   255
            Width           =   1320
         End
         Begin VB.TextBox txtintfolha 
            Height          =   285
            Left            =   7245
            MaxLength       =   4
            TabIndex        =   14
            Top             =   675
            Width           =   555
         End
         Begin VB.TextBox txtintlivro 
            Height          =   285
            Left            =   8370
            MaxLength       =   8
            TabIndex        =   16
            Top             =   675
            Width           =   945
         End
         Begin VB.TextBox txtintExercicio 
            Height          =   285
            Left            =   8805
            MaxLength       =   8
            TabIndex        =   8
            Top             =   255
            Width           =   495
         End
         Begin VB.CheckBox chk_Atualizacao 
            Caption         =   "Valores atualizados"
            Enabled         =   0   'False
            Height          =   195
            Left            =   1890
            TabIndex        =   17
            Top             =   1050
            Visible         =   0   'False
            Width           =   1710
         End
         Begin MSDataListLib.DataCombo dbc_intReceita 
            Height          =   315
            Left            =   1890
            TabIndex        =   3
            Top             =   255
            Width           =   3435
            _ExtentX        =   6059
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin VB.Label lblStatusContagem 
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   1920
            TabIndex        =   21
            Top             =   1650
            Width           =   1005
         End
         Begin VB.Label lblStatusTotal 
            Alignment       =   1  'Right Justify
            Height          =   225
            Left            =   8430
            TabIndex        =   20
            Top             =   1650
            Width           =   825
         End
         Begin VB.Label lbl_inscricao 
            AutoSize        =   -1  'True
            Caption         =   "Data de Inscrição"
            Height          =   195
            Left            =   555
            TabIndex        =   9
            Top             =   750
            Width           =   1260
         End
         Begin VB.Label lbl_compreceita 
            AutoSize        =   -1  'True
            Caption         =   "Composição da receita"
            Height          =   195
            Left            =   195
            TabIndex        =   2
            Top             =   345
            Width           =   1620
         End
         Begin VB.Label lbl_certidao 
            AutoSize        =   -1  'True
            Caption         =   "Certidão"
            Height          =   195
            Left            =   3660
            TabIndex        =   11
            Top             =   750
            Width           =   585
         End
         Begin VB.Label lbl_cadastro 
            AutoSize        =   -1  'True
            Caption         =   "Cadastro"
            Height          =   195
            Left            =   5775
            TabIndex        =   5
            Top             =   345
            Width           =   630
         End
         Begin VB.Label lbl_folha 
            AutoSize        =   -1  'True
            Caption         =   "Folha"
            Height          =   195
            Left            =   6765
            TabIndex        =   13
            Top             =   765
            Width           =   390
         End
         Begin VB.Label lbl_livro 
            AutoSize        =   -1  'True
            Caption         =   "Livro"
            Height          =   195
            Left            =   7950
            TabIndex        =   15
            Top             =   765
            Width           =   345
         End
         Begin VB.Label lbl_exercicio 
            AutoSize        =   -1  'True
            Caption         =   "Exercício"
            Height          =   195
            Left            =   8055
            TabIndex        =   7
            Top             =   345
            Width           =   675
         End
      End
   End
End
Attribute VB_Name = "frmCadDividaAtivaComposicao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mobjAux                     As Object
Dim mblnSelecionou              As Boolean
Dim blnVerificaProximoLivro     As Boolean
Dim vetParcelas(0, 6)           As String
Dim dblValorTaxa                As Double
Dim dblValorImposto             As Double

Private Sub cmd_TabelaComposicaoDaReceita_Click()
    ChamaFormCadastro frmCadComposicaoDaReceita, dbc_intReceita
End Sub

Private Sub dbc_intReceita_Change()
    If dbc_intReceita.MatchedWithList = True And dbc_intReceita.BoundText <> "" Then
        PreencheCadastro CLng(dbc_intReceita.BoundText)
    End If
End Sub

Private Sub dbc_intReceita_Click(Area As Integer)
    DropDownDataCombo dbc_intReceita, Me, Area
End Sub

Private Sub dbc_intReceita_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbc_intReceita, Me, , KeyCode, Shift
End Sub

Private Sub txtcadastro_GotFocus()
    MarcaCampo txtcadastro
End Sub

Private Sub txtcadastro_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtcadastro
End Sub

Private Sub txtdtmdtinscricao_GotFocus()
    If txtdtmdtinscricao = "" Then
        txtdtmdtinscricao = gstrDataFormatada(Date)
    End If
    MarcaCampo txtdtmdtinscricao
End Sub

Private Sub txtdtmdtinscricao_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "D", txtdtmdtinscricao
End Sub

Private Sub txtdtmdtinscricao_LostFocus()
    txtdtmdtinscricao = gstrDataFormatada(txtdtmdtinscricao)
End Sub

Private Sub txtintCertidao_GotFocus()
    MarcaCampo txtintCertidao
End Sub

Private Sub txtintcertidao_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtintCertidao
End Sub

Private Sub txtintExercicio_GotFocus()
    MarcaCampo txtintExercicio
End Sub

Private Sub txtintExercicio_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtintExercicio
End Sub

Private Sub txtintFolha_GotFocus()
    MarcaCampo txtintfolha
End Sub

Private Sub txtintfolha_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtintfolha
End Sub

Private Sub txtintLivro_GotFocus()
    MarcaCampo txtintlivro
End Sub

Private Sub txtintlivro_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtintlivro
End Sub

Private Sub dbc_intReceita_GotFocus()
    MarcaCampo dbc_intReceita
End Sub

Private Sub dbc_intReceita_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbc_intReceita
End Sub

Private Sub Form_Activate()
    gintCodSeguranca = 1296

    If mblnSelecionou Then
        HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrDeletar
    Else
        HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrDeletar
    End If
    
    If mobjAux Is Nothing Then
        HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar
    Else
        HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrAplicar
    End If
    
End Sub

Private Sub Form_Deactivate()
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar
End Sub

Private Sub Form_Load()
    
    VerificaObjParaAplicar mobjAux
    TrocaCorObjeto txtcadastro, True
    TrocaCorObjeto txtintCertidao, True
    TrocaCorObjeto txtintfolha, True
    TrocaCorObjeto txtintlivro, True
    
    dbc_intReceita.Tag = strQueryComposicaoReceita & ";strDescricao"
      
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar
    mblnSelecionou = False
End Sub

Function strQueryComposicaoReceita()
    Dim strSQL As String
    
    strSQL = ""
    strSQL = strSQL & "SELECT PKId," & gstrCONVERT(CDT_VARCHAR, "intCodigo") & strCONCAT & "' - '" & strCONCAT & " Ltrim(Rtrim(strDescricao)) as strDescricao "
    strSQL = strSQL & "FROM " & gstrComposicaoDaReceita & " "
    strSQL = strSQL & "WHERE bytDividaAtiva = 1 " ' And ""
    'strSQL = strSQL & "intUtilizacao <> 3 "
    strSQL = strSQL & "ORDER BY strDescricao"
    
    strQueryComposicaoReceita = strSQL
    
End Function

Public Sub MantemForm(ByVal strModoOperacao As String)
    
    If UCase(gstrImprimir) = UCase(strModoOperacao) Then

    ElseIf UCase(strModoOperacao) = gstrPreencherLista Then
        PreencherListaDeOpcoes Me.ActiveControl
    
    ElseIf UCase(gstrFechar) = UCase(strModoOperacao) Then
        Unload Me
    
    ElseIf UCase(gstrNovo) = UCase(strModoOperacao) Then
        Limpa_Controles Me, True, True, True, True, True
        prgImportacaoDativa.Visible = True
        prgImportacaoDativa.Min = 0
        prgImportacaoDativa.Max = 1
        prgImportacaoDativa.Value = 0
        lblStatusContagem.Caption = ""
        lblStatusTotal.Caption = ""
        dbc_intReceita.SetFocus
    
    ElseIf UCase(gstrSalvar) = UCase(strModoOperacao) Then
        If blnDadosOk Then
            If gblnExclusaoGravacaoOk("SALVAR", "Deseja realmente Inscrever em Dívida Ativa") Then
                Set gobjBanco = New clsBanco
                gobjBanco.ExecutaBeginTrans
                If blnSalvarLancamentoDA Then
                    gobjBanco.ExecutaCommitTrans
                    Limpa_Controles Me, True, True, True, True, True
                    dbc_intReceita.SetFocus
                Else
                    ExibeMensagem "Não foi possível gravar Inscrição de Dívida Ativa"
                    gobjBanco.ExecutaRollbackTrans
                End If
            End If
        End If
    Else
        
    End If
End Sub

Private Sub PreencheCadastro(lngPkid As Long)
Dim strSQL          As String
Dim adoResultado    As ADODB.Recordset
    
    strSQL = ""
    strSQL = strSQL & "Select * From " & gstrComposicaoDaReceita & " Where pkid = " & lngPkid
    
    Set gobjBanco = New clsBanco
    
    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
        If Not adoResultado.EOF Then
            Select Case gstrENulo(adoResultado!intUtilizacao)
                Case 1
                    txtcadastro = "Imobiliário"
                Case 2
                    txtcadastro = "Econômico"
                Case 3
                    txtcadastro = "Dívida Ativa"
                Case 4
                    txtcadastro = "Acordo"
                Case 5
                    txtcadastro = "Preco Público"
                Case 6
                    txtcadastro = "ISS Construção"
                Case Else
                    txtcadastro = ""
            End Select
            PreencheCertidaoLivroFolha adoResultado!Pkid
        End If
    End If
    
End Sub

Private Function blnDadosAtualizacao() As Boolean
    
    blnDadosAtualizacao = False
    
    If dbc_intReceita.MatchedWithList = False Then
        ExibeMensagem "O campo de Composição da Receita é obrigatório."
        dbc_intReceita.SetFocus
        Exit Function
    ElseIf Trim(Len(txtintExercicio)) <> 4 Then
        ExibeMensagem "O campo de exercício deve ser preenchido corretamente."
        txtintExercicio.SetFocus
        Exit Function
    ElseIf Trim(txtdtmdtinscricao) = "" Then
        ExibeMensagem "O campo de data da inscrição deve ser preenchido corretamente."
        txtdtmdtinscricao.SetFocus
        Exit Function
    ElseIf Trim(txtdtmdtinscricao) <> "" Then
        If gblnDataValida(txtdtmdtinscricao, True) = False Then
            txtdtmdtinscricao.SetFocus
            Exit Function
        End If
    End If
    
    blnDadosAtualizacao = True
    
End Function

Private Function blnDadosOk() As Boolean
    
    blnDadosOk = False
    
    If Not dbc_intReceita.MatchedWithList Then
        ExibeMensagem "O campo de Composição da Receita é obrigatório."
        dbc_intReceita.SetFocus
        Exit Function
    ElseIf Not blnExisteParametroAtualizacao Then
        ExibeMensagem "Não foi encontrado PARAMETROS DE DÍVIDA ATÍVA para inscrever em dívida atíva esta inscrição cadastral!"
        Exit Function
    ElseIf Trim(txtintExercicio.Text) = "" Then
        ExibeMensagem "O campo de Exercício é obrigatório."
        txtintExercicio.SetFocus
        Exit Function
    ElseIf Trim(txtdtmdtinscricao) = "" Then
        ExibeMensagem "O campo de Data da Inscrição deve ser preenchido corretamente."
        txtdtmdtinscricao.SetFocus
        Exit Function
    ElseIf Trim(txtdtmdtinscricao) <> "" Then
        If gblnDataValida(txtdtmdtinscricao, True) = False Then
            txtdtmdtinscricao.SetFocus
            Exit Function
        End If
    End If
    
    blnDadosOk = True
    
End Function

Private Function blnSalvarLancamentoDA() As Boolean
    Dim strSQL           As String
    Dim intFor           As Integer
    Dim adoResultado     As ADODB.Recordset
    Dim adoConsulta      As ADODB.Recordset
    Dim adoParcelas      As ADODB.Recordset
    Dim adoAtualizadas   As ADODB.Recordset
    Dim adoAux           As ADODB.Recordset
    Dim strIDLanctoValor As String
    
    blnSalvarLancamentoDA = False

    If bnlVerificaParametros = False Then Exit Function
    
    strSQL = ""
    strSQL = strSQL & "SELECT "
    
    If bytDBType = Oracle Then
        strSQL = strSQL & "/*+ index(A) */ " 'Parâmetro adicional inserido para otimizar a consulta à pedido do DBA
    End If
    
    strSQL = strSQL & "LA.Pkid, "
    strSQL = strSQL & "LA.Intcomposicaodareceita, "
    strSQL = strSQL & "da.dblvalorimposto, "
    strSQL = strSQL & "da.dblvalortaxas, "
    strSQL = strSQL & "CR.INTUTILIZACAO, "
    strSQL = strSQL & gstrRIGHT("LA.strInscricao", gintRetornaTamanhoMascara(TYP_IMOBILIARIA)) & " strInscricao, "
    strSQL = strSQL & "LA.strNumeroAviso, "
    strSQL = strSQL & "LA.intExercicio, "
    strSQL = strSQL & "LA.strnomeproprietario, "
    strSQL = strSQL & "PDA.Intlivro, "
    strSQL = strSQL & "pda.intfolha, "
    strSQL = strSQL & "pda.intcertidao, "
    strSQL = strSQL & "LA.strcnpjcpf, "
    strSQL = strSQL & "LA.stridentidade, "
    strSQL = strSQL & "LA.strlogradouro, "
    strSQL = strSQL & "LA.strnumero, "
    strSQL = strSQL & "LA.strcomplemento, "
    strSQL = strSQL & "LA.strbairro, "
    strSQL = strSQL & "LA.strmunicipio, "
    strSQL = strSQL & "LA.struf, "
    strSQL = strSQL & "LA.intcep, "
    strSQL = strSQL & "LA.strlogradouroc, "
    strSQL = strSQL & "LA.strnumeroc, "
    strSQL = strSQL & "LA.strcomplementoc, "
    strSQL = strSQL & "LA.strbairroc, "
    strSQL = strSQL & "LA.strmunicipioc, "
    strSQL = strSQL & "LA.strufc, "
    strSQL = strSQL & "LA.intcepc, "
    strSQL = strSQL & "LA.strpromissario, "
    strSQL = strSQL & "LA.strindexador, "
    strSQL = strSQL & "LA.dblvlindexador "
        
    If bytDBType = Oracle Then
        strSQL = strSQL & "FROM "
        strSQL = strSQL & gstrLancamentoAlfa & " LA, "
        strSQL = strSQL & gstrComposicaoDaReceita & " CR, "
        strSQL = strSQL & gstrParametroDividaAtiva & " PDA, "
        strSQL = strSQL & gstrLancamentoValor & " LV, "
        strSQL = strSQL & gstrLancamentoPagamento & " LP, "
        strSQL = strSQL & gstrDativa & " DA "
        strSQL = strSQL & "WHERE "
        strSQL = strSQL & "CR.Pkid = LA.Intcomposicaodareceita AND "
        strSQL = strSQL & "CR.Pkid = PDA.Intcomposicaodareceita " & strOUTJOracle & " AND "
        strSQL = strSQL & "LA.pkid = LV.Intlancamentoalfa" & strOUTJOracle & " AND "
        strSQL = strSQL & "LV.PKID " & strOUTJSQLServer & "= LP.INTLANCAMENTOVALOR " & strOUTJOracle & " AND "
        strSQL = strSQL & "LP.PKID Is Null AND LA.Dtmdtcancelamento is null AND "
        'strsql = strsql & "(select sum(" & gstrISNULL("dblvalor", 0) & ") from " & gstrLancamentoValor & " where intlancamentoalfa = la.pkid AND LV.BITPARCELAVALIDA = 1) > 0 AND "
        strSQL = strSQL & "LV.BITPARCELAVALIDA = 1 AND "
        strSQL = strSQL & "LV.IntlancamentoalfaDativa Is Null AND "
            
        If chk_InscricaoAcordo.Value = vbUnchecked Then
            strSQL = strSQL & "LV.INTLANCAMENTOALFAACORDO IS NULL AND "
        End If
        
        strSQL = strSQL & "LA.pkid " & strOUTJSQLServer & "= DA.intLancamentoAlfa " & strOUTJOracle & " AND DA.intLancamentoAlfa is null AND "
        strSQL = strSQL & "CR.bytDividaAtiva = 1 "
        strSQL = strSQL & " AND LA.intComposicaoDaReceita = " & dbc_intReceita.BoundText
        strSQL = strSQL & " AND LA.intExercicio = " & txtintExercicio.Text
        'strSql = strSql & " AND LA.strinscricao between '00000000000001001001' and '00000000000001999999' "
    Else
        strSQL = strSQL & "FROM " & gstrLancamentoAlfa & " LA INNER JOIN "
        strSQL = strSQL & gstrComposicaoDaReceita & " CR ON LA.INTCOMPOSICAODARECEITA = CR.PKId LEFT JOIN "
        strSQL = strSQL & gstrParametroDividaAtiva & " PDA ON CR.PKId = PDA.INTCOMPOSICAODARECEITA INNER JOIN "
        strSQL = strSQL & gstrLancamentoValor & " LV ON LA.PKId = LV.intLancamentoAlfa LEFT OUTER JOIN "
        strSQL = strSQL & gstrLancamentoPagamento & " LP ON LV.PKId = LP.INTLANCAMENTOVALOR LEFT OUTER JOIN "
        strSQL = strSQL & gstrDativa & " DA ON LA.PKId = DA.INTLANCAMENTOALFA "
        strSQL = strSQL & "WHERE (LP.PKID IS NULL) AND (LA.dtmDtCancelamento IS NULL) AND (LV.bitParcelaValida = 1) AND (LV.intLancamentoAlfaDAtiva IS NULL) "
        'strsql = strsql & "((SELECT SUM(ISNULL(dblvalor, 0)) FROM " & gstrLancamentoValor & _
        '                  " WHERE intlancamentoalfa = la.pkid AND LV.BITPARCELAVALIDA = 1) > 0) AND "
                          
        If chk_InscricaoAcordo.Value = vbUnchecked Then
            strSQL = strSQL & " AND (LV.intLancamentoAlfaAcordo IS NULL) "
        End If
        
        strSQL = strSQL & " AND (DA.INTLANCAMENTOALFA IS NULL) AND (CR.bytDividaAtiva = 1) AND "
        strSQL = strSQL & " (LA.INTCOMPOSICAODARECEITA = " & dbc_intReceita.BoundText & ") AND (LA.intExercicio = " & txtintExercicio.Text & ") "
    End If
    strSQL = strSQL & " GROUP BY LA.Pkid, LA.Intcomposicaodareceita, da.dblvalorimposto, da.dblvalortaxas, CR.INTUTILIZACAO, "
    strSQL = strSQL & "LA.strInscricao, LA.strNumeroAviso, LA.intExercicio, LA.strnomeproprietario, PDA.Intlivro, pda.intfolha, "
    strSQL = strSQL & "pda.intcertidao, LA.strcnpjcpf, LA.stridentidade, LA.strlogradouro, LA.strnumero, LA.strcomplemento, "
    strSQL = strSQL & "LA.strbairro, LA.strmunicipio, LA.struf, LA.intcep, LA.strlogradouroc, LA.strnumeroc, LA.strcomplementoc, "
    strSQL = strSQL & "LA.strbairroc, LA.strmunicipioc, LA.strufc, LA.intcepc, LA.strpromissario, LA.strindexador, LA.dblvlindexador Order By LA.strInscricao "
    
    DoEvents
   
    Set gobjBanco = New clsBanco
    
    If gobjBanco.CriaADO(strSQL, 200, adoResultado) Then
        With adoResultado
        
        If .RecordCount > 0 Then
            prgImportacaoDativa.Visible = True
            prgImportacaoDativa.Min = 0
            prgImportacaoDativa.Max = .RecordCount
            prgImportacaoDativa.Value = 0
        End If
            
            Do While Not .EOF
                
                'Vamos consultar se a somas das parcelas é maior que Zero
                If bytDBType = Oracle Then
                    strSQL = " SELECT SUM(" & gstrISNULL("LV.dblvalor", 0) & ") dblTotal " & _
                             " FROM " & gstrLancamentoValor & " LV, " & gstrLancamentoPagamento & " LP " & _
                             " WHERE LV.intlancamentoalfa = " & adoResultado("pkid").Value & " AND LV.BITPARCELAVALIDA = 1 " & _
                             " AND LV.PKID " & strOUTJSQLServer & "= LP.INTLANCAMENTOVALOR " & strOUTJOracle & " AND LP.Pkid IS NULL "
                Else
                    strSQL = " SELECT SUM(" & gstrISNULL("LV.dblvalor", 0) & ") dblTotal " & _
                             " FROM " & gstrLancamentoValor & " LV LEFT JOIN " & gstrLancamentoPagamento & " LP ON LV.Pkid = LP.intLancamentoValor " & _
                             " WHERE LV.intlancamentoalfa = " & adoResultado("pkid").Value & " AND LP.Pkid IS NULL AND LV.BITPARCELAVALIDA = 1 "
                End If
                If gobjBanco.CriaADO(strSQL, 10, adoConsulta) Then
                    If adoConsulta("dblTotal").Value > 0 Then
                    
                        AdicionaTaxasImpostos (adoResultado!Pkid), adoConsulta("dblTotal").Value
                        strIDLanctoValor = ""
                        strSQL = ""
                        strSQL = strSQL & " INSERT INTO " & gstrDativa
                        strSQL = strSQL & "(intlancamentoalfa, intfolha, intlivro, dtmdtinscricao, strobservacao, strnomeproprietario, strcnpjcpf, stridentidade, strlogradouro, strnumero, strcomplemento, strbairro, strmunicipio, struf, intcep, strlogradouroc, strnumeroc, strcomplementoc, strbairroc, strmunicipioc, strufc, intcepc, strpromissario, strindexador, dblvlindexador, dtmdtatualizacao, lngcodusr, intcertidao, dblValorImposto, dblValorTaxas)"
                        strSQL = strSQL & " Values "
                        strSQL = strSQL & "(" & adoResultado!Pkid & ", "
                        strSQL = strSQL & gstrENulo(txtintfolha, , True) & ", "
                        strSQL = strSQL & gstrENulo(txtintlivro, , True) & ", "
                        strSQL = strSQL & gstrConvDtParaSql(gstrENulo(txtdtmdtinscricao, , True)) & ", "
                        strSQL = strSQL & "'', "
                        strSQL = strSQL & "'" & Replace(gstrENulo(adoResultado!strnomeproprietario), "'", " ") & "', "
                        strSQL = strSQL & "'" & gstrENulo(adoResultado!StrCnpjCpf) & "', "
                        strSQL = strSQL & "'" & gstrENulo(adoResultado!STRIDENTIDADE) & "', "
                        strSQL = strSQL & "'" & Replace(LTrim(RTrim(gstrENulo(adoResultado!strLogradouro))), "'", " ") & "', "
                        strSQL = strSQL & "'" & LTrim(RTrim(gstrENulo(adoResultado!strNumero))) & "', "
                        strSQL = strSQL & "'" & Replace(LTrim(RTrim(gstrENulo(adoResultado!STRCOMPLEMENTO))), "'", " ") & "', "
                        strSQL = strSQL & "'" & Replace(Trim(gstrENulo(adoResultado!strBairro)), "'", "''") & "', "
                        strSQL = strSQL & "'" & Replace(LTrim(RTrim(gstrENulo(adoResultado!STRMUNICIPIO))), "'", " ") & "', "
                        strSQL = strSQL & "'" & IIf(LTrim(RTrim(gstrENulo(adoResultado!STRUF))) = 0, " ", LTrim(RTrim(gstrENulo(adoResultado!STRUF)))) & "', "
                        strSQL = strSQL & gstrENulo(adoResultado!INTCEP, , True) & ", "
                        
                        strSQL = strSQL & "'" & Replace(IIf(Len(LTrim(RTrim(gstrENulo(adoResultado!strlogradouroc)))) = 0, " ", LTrim(RTrim(gstrENulo(adoResultado!strlogradouroc)))), "'", " ") & "', "
                        strSQL = strSQL & "'" & Replace(IIf(Len(LTrim(RTrim(gstrENulo(adoResultado!strNumeroC)))) = 0, " ", LTrim(RTrim(gstrENulo(adoResultado!strNumeroC)))), "'", " ") & "', "
                        strSQL = strSQL & "'" & Replace(IIf(Len(LTrim(RTrim(gstrENulo(adoResultado!strComplementoC)))) = 0, " ", LTrim(RTrim(gstrENulo(adoResultado!strComplementoC)))), "'", " ") & "', "
                        strSQL = strSQL & "'" & Replace(IIf(Len(Trim(gstrENulo(adoResultado!strBairroC))) = 0, " ", (Replace(Trim(gstrENulo(adoResultado!strBairroC)), "'", "''"))), "'", " ") & "', "
                        strSQL = strSQL & "'" & Replace(IIf(Len(LTrim(RTrim(gstrENulo(adoResultado!strMunicipioC)))) = 0, " ", LTrim(RTrim(gstrENulo(adoResultado!strMunicipioC)))), "'", " ") & "', "
                        strSQL = strSQL & "'" & IIf(Len(LTrim(RTrim(gstrENulo(adoResultado!strUFC)))) = 0, " ", LTrim(RTrim(gstrENulo(adoResultado!strUFC)))) & "', "
                        strSQL = strSQL & IIf(IsNull(adoResultado!intcepc), 0, gstrENulo(adoResultado!intcepc, , True)) & ", "
                        
                        strSQL = strSQL & "'" & Replace(gstrENulo(adoResultado!strpromissario), "'", " ") & "', "
                        strSQL = strSQL & "'" & gstrENulo(adoResultado!Strindexador) & "', "
                        strSQL = strSQL & gstrENulo(gstrConvVrParaSql(gstrConvVrDoSql(adoResultado!dblvlIndexador)), , True) & ", "
                        strSQL = strSQL & gstrConvDtParaSql(gstrDataDoSistema) & ", "
                        strSQL = strSQL & glngCodUsr & ", "
                        strSQL = strSQL & gstrENulo(txtintCertidao, , True) & ", "
                        strSQL = strSQL & gstrConvVrParaSql(dblValorImposto) & ", "
                        strSQL = strSQL & gstrConvVrParaSql(dblValorTaxa) & ")"
                        
                        If Not gobjBanco.Execute(strSQL) Then
                            prgImportacaoDativa.Visible = False
                            lblStatusContagem.Caption = ""
                            lblStatusTotal.Caption = ""
                            Exit Function
                        End If
                        
                        strSQL = ""
                        If gobjBanco.CriaADO(strQueryParcela(adoResultado("Pkid")), 5, adoParcelas) Then
                            strSQL = IIf(bytDBType = Oracle, "Begin ", "")
                            Do While Not adoParcelas.EOF
                                
                                If Not chk_Atualizacao.Value Then
                                    vetParcelas(0, 0) = adoParcelas!intNumeroParcela
                                    vetParcelas(0, 1) = adoParcelas!dblValor
                                    vetParcelas(0, 2) = adoParcelas!dtmVencimento
                                    vetParcelas(0, 3) = 0
                                    vetParcelas(0, 4) = 0
                                    vetParcelas(0, 5) = 0
                                    vetParcelas(0, 6) = adoParcelas!intMoeda
                                Else
                                    strSQL = gstrStoredProcedure("sp_AtualizaParcela", dbc_intReceita.BoundText & ", " & txtintExercicio & ", " & gstrENulo(adoParcelas!intNumeroParcela) & ", " & gstrConvDtParaSql(adoParcelas!dtmVencimento) & ", " & gstrConvDtParaSql(txtdtmdtinscricao) & ", " & gstrConvVrParaSql(gstrConvVrDoSql(gstrENulo(adoParcelas!dblValor), 2, , True)) & ", " & Val(gstrENulo(adoParcelas!intMoeda)), True)
        
                                    If gobjBanco.CriaADO(strSQL, 80, adoAtualizadas) Then
                                        With adoAtualizadas
                                            If Not .EOF Then
                                            
                                                vetParcelas(0, 0) = Space$(0) & adoParcelas!intNumeroParcela                                  'Número das parcelas
                                                vetParcelas(0, 1) = Space$(0) & CCur(gstrConvVrDoSql(adoAtualizadas("dblValorPrincipal").Value))  'Valor da parcela
                                                vetParcelas(0, 2) = Space$(0) & gstrDataFormatada(adoParcelas!dtmVencimento)                  'Vencimento da parcela
                                                vetParcelas(0, 3) = 0
                                                vetParcelas(0, 4) = 0
                                                vetParcelas(0, 5) = CCur(gstrConvVrDoSql(adoAtualizadas("dblValorCorrecao").Value))               'Valor de Correção Atualizado
                                                vetParcelas(0, 6) = Space$(0) & adoParcelas!intMoeda                                          'Abreviatura da Moeda
                                            
                                            End If
                                        End With
                                    Else
                                        Exit Function
                                    End If
                                    
                                End If
                                
                                strSQL = strSQL & " INSERT INTO " & gstrDaParcel
                                strSQL = strSQL & "(intdativa, intparcela, dblvalor, dtmdtvencimento, dblmulta, dblcorrecaomonet, dbljuros, intmoeda, dtmdtatualizacao, lngcodusr) "
                                strSQL = strSQL & "Values("
                                strSQL = strSQL & glngRetornaPkidTabelaPai("seqTBLDATIVA", gstrDativa) & ", " 'gstrENulo(adoAux!Pkid)
                                strSQL = strSQL & vetParcelas(0, 0) & ", "
                                strSQL = strSQL & gstrConvVrParaSql(vetParcelas(0, 1)) & ", "
                                strSQL = strSQL & gstrConvDtParaSql(vetParcelas(0, 2)) & ", "
                                strSQL = strSQL & gstrConvVrParaSql(vetParcelas(0, 3)) & ", "
                                strSQL = strSQL & gstrConvVrParaSql(vetParcelas(0, 4)) & ", "
                                strSQL = strSQL & gstrConvVrParaSql(vetParcelas(0, 5)) & ", "
                                strSQL = strSQL & vetParcelas(0, 6) & ", "
                                strSQL = strSQL & gstrConvDtParaSql(gstrDataDoSistema) & ", "
                                strSQL = strSQL & glngCodUsr
                                strSQL = strSQL & ")" & IIf(bytDBType = Oracle, ";", "")
                                
                                strIDLanctoValor = strIDLanctoValor & adoParcelas!Pkid & ","
                                
                                adoParcelas.MoveNext
                                
                            Loop
                        End If
                        
                        strSQL = strSQL & strParametrosDividaAtiva(dbc_intReceita.BoundText) & IIf(bytDBType = Oracle, ";", "")
                        
                        If Len(strIDLanctoValor) > 0 Then
                            strIDLanctoValor = Mid(strIDLanctoValor, 1, Len(strIDLanctoValor) - 1)
                            strSQL = strSQL & " UPDATE " & gstrLancamentoValor & " SET INTLANCAMENTOALFADATIVA = " & adoResultado!Pkid & " WHERE pkid in(" & strIDLanctoValor & ")"
                            strSQL = strSQL & IIf(bytDBType = Oracle, ";", "")
                        End If
                        
                        strSQL = strSQL & IIf(bytDBType = Oracle, "End;", "")
                        
                        If .RecordCount > 0 Then
                            prgImportacaoDativa.Value = .AbsolutePosition
                            lblStatusContagem.Caption = prgImportacaoDativa.Value
                            lblStatusTotal.Caption = .RecordCount
                        End If
                        
                        DoEvents
                
                        If Not gobjBanco.Execute(strSQL) Then
                            prgImportacaoDativa.Visible = False
                            lblStatusContagem.Caption = ""
                            lblStatusTotal.Caption = ""
                            Exit Function
                        End If
                        
                        PreencheCertidaoLivroFolha dbc_intReceita.BoundText, False
                    
                    End If
                End If
                
                .MoveNext
            
            Loop
        
        End With
        
    End If
    
    prgImportacaoDativa.Visible = False
    lblStatusContagem.Caption = ""
    lblStatusTotal.Caption = ""
    
    blnSalvarLancamentoDA = True
    
End Function

Private Function PreencheCertidaoLivroFolha(intComposicaoDaReceita As Double, Optional blnVerificaLivro As Boolean = True) As String

Dim strSQL                  As String
Dim adoResultado            As New ADODB.Recordset
Dim intQtdCertidaoUltFolha  As Integer
Dim blnProximoLivro         As Boolean
Dim intFolha                As Integer
Dim intLivro                As Integer
Dim intCertidao             As Long

    strSQL = "SELECT "
    strSQL = strSQL & "PDA.INTCERTIDAO, "
    strSQL = strSQL & "PDA.INTFOLHA, "
    strSQL = strSQL & "PDA.INTLIVRO, "
    strSQL = strSQL & "pda.intfolhaporlivro, "
    strSQL = strSQL & "pda.intcertidaoporfolha, "
    strSQL = strSQL & "intQtdCertidaoUltFolha "
    strSQL = strSQL & "FROM "
    strSQL = strSQL & gstrParametroDividaAtiva & " PDA "
    strSQL = strSQL & "Where "
    strSQL = strSQL & "PDA.intComposicaoDaReceita = " & intComposicaoDaReceita
    
    Set gobjBanco = New clsBanco
    
    If gobjBanco.CriaADO(strSQL, 10, adoResultado) Then
        With adoResultado
            If Not .EOF Then
                
                intCertidao = IIf(IsNull(!intCertidao), 1, !intCertidao + 1)
                intFolha = !intFolha
                intLivro = !intLivro
                intQtdCertidaoUltFolha = !intQtdCertidaoUltFolha
                
                'Vamos se chegou a fim do livro
                If (intQtdCertidaoUltFolha Mod !intCertidaoPorFolha = 0) And intQtdCertidaoUltFolha <> 0 Then
                   intQtdCertidaoUltFolha = 0
                   If (intFolha Mod !intFolhaPorLivro = 0) And intFolha <> 0 Then
                       intFolha = 1
                       intLivro = intLivro + 1
                   Else
                      intFolha = intFolha + 1
                   End If
                End If
                
                txtintCertidao = intCertidao
                txtintfolha = intFolha
                txtintlivro = intLivro
                
            Else
                strSQL = "SELECT "
                strSQL = strSQL & "PDA.INTCERTIDAO, "
                strSQL = strSQL & "PDA.INTFOLHA, "
                strSQL = strSQL & "PDA.INTLIVRO, "
                strSQL = strSQL & "pda.intfolhaporlivro, "
                strSQL = strSQL & "pda.intcertidaoporfolha, "
                strSQL = strSQL & "intQtdCertidaoUltFolha "
                strSQL = strSQL & "FROM "
                strSQL = strSQL & gstrParametroDividaAtiva & " PDA "
                strSQL = strSQL & "Where "
                strSQL = strSQL & "PDA.intComposicaoDaReceita Is null "
                
                Set gobjBanco = New clsBanco
                Set adoResultado = New ADODB.Recordset
                
                If gobjBanco.CriaADO(strSQL, 10, adoResultado) Then
                    If Not adoResultado.EOF Then
                    
                        intCertidao = IIf(IsNull(adoResultado!intCertidao), 1, adoResultado!intCertidao + 1)
                        intFolha = adoResultado!intFolha
                        intLivro = adoResultado!intLivro
                        intQtdCertidaoUltFolha = adoResultado!intQtdCertidaoUltFolha
                        
                        'Vamos se chegou a fim do livro
                        If (intQtdCertidaoUltFolha Mod adoResultado!intCertidaoPorFolha = 0) And intQtdCertidaoUltFolha <> 0 Then
                           intQtdCertidaoUltFolha = 0
                           If (intFolha Mod adoResultado!intFolhaPorLivro = 0) And intFolha <> 0 Then
                              intFolha = 1
                              intLivro = intLivro + 1
                           Else
                              intFolha = intFolha + 1
                           End If
                        End If
                        
                        txtintCertidao = intCertidao
                        txtintfolha = intFolha
                        txtintlivro = intLivro
                    
                    Else
                        txtintCertidao = ""
                        txtintfolha = ""
                        txtintlivro = ""
                        ExibeMensagem "Não há parâmetros para inscrição em Dívida Ativa."
                    End If
                End If
            End If
        End With
    End If
    
End Function

Private Function strParametrosDividaAtiva(intComposicaoDaReceita As Double) As String

    Dim adoResultado    As New ADODB.Recordset
    Dim blnComposicao   As Boolean
    Dim strSQL          As String
    Dim intCertidao     As Long
    Dim intFolha        As Integer
    Dim intLivro        As Integer
    Dim intQtdCertidaoUltFolha As Integer
    Dim blnNovaFolha    As Boolean
    Dim blnProximoLivro As Boolean
    
    blnNovaFolha = False
    blnComposicao = False
    
    strSQL = "SELECT "
    strSQL = strSQL & "PDA.INTCERTIDAO, "
    strSQL = strSQL & "PDA.INTFOLHA, "
    strSQL = strSQL & "PDA.INTLIVRO, "
    strSQL = strSQL & "pda.intfolhaporlivro, "
    strSQL = strSQL & "pda.intcertidaoporfolha, "
    strSQL = strSQL & "pda.intQtdCertidaoUltFolha "
    strSQL = strSQL & "FROM "
    strSQL = strSQL & gstrParametroDividaAtiva & " PDA "
    strSQL = strSQL & "Where "
    strSQL = strSQL & "PDA.intComposicaoDaReceita = " & intComposicaoDaReceita
    
    If gobjBanco.CriaADO(strSQL, 10, adoResultado) Then
        If Not adoResultado.EOF Then
            With adoResultado
                blnComposicao = True
                If !intFolha < !intFolhaPorLivro Then
                    intCertidao = !intCertidao + 1
                    intLivro = !intLivro
                    If (!intQtdCertidaoUltFolha + 1) >= !intCertidaoPorFolha Then
                        intFolha = !intFolha + 1
                        intQtdCertidaoUltFolha = 0
                    Else
                        intQtdCertidaoUltFolha = !intQtdCertidaoUltFolha + 1
                        intFolha = !intFolha
                    End If
                Else
                    If !intFolha = !intFolhaPorLivro Then
                        intCertidao = !intCertidao + 1
                        intLivro = !intLivro
                        If (!intQtdCertidaoUltFolha + 1) > !intCertidaoPorFolha Then
                            intLivro = !intLivro + 1
                            intFolha = 1
                            intQtdCertidaoUltFolha = 1
                        Else
                            intQtdCertidaoUltFolha = !intQtdCertidaoUltFolha + 1
                            intFolha = !intFolha
                        End If
                    End If
                End If
            End With
        Else
            strSQL = "SELECT "
            strSQL = strSQL & "PDA.INTCERTIDAO, "
            strSQL = strSQL & "PDA.INTFOLHA, "
            strSQL = strSQL & "PDA.INTLIVRO, "
            strSQL = strSQL & "pda.intfolhaporlivro, "
            strSQL = strSQL & "pda.intcertidaoporfolha, "
            strSQL = strSQL & "pda.intQtdCertidaoUltFolha "
            strSQL = strSQL & "FROM "
            strSQL = strSQL & gstrParametroDividaAtiva & " PDA "
            strSQL = strSQL & "Where "
            strSQL = strSQL & "PDA.intComposicaoDaReceita Is Null "
            
            If gobjBanco.CriaADO(strSQL, 10, adoResultado) Then
                If Not adoResultado.EOF Then
                    With adoResultado
                        If !intFolha < !intFolhaPorLivro Then
                            intCertidao = !intCertidao + 1
                            intLivro = !intLivro
                            If (!intQtdCertidaoUltFolha + 1) >= !intCertidaoPorFolha Then
                                intFolha = !intFolha + 1
                                intQtdCertidaoUltFolha = 0
                            Else
                                intQtdCertidaoUltFolha = !intQtdCertidaoUltFolha + 1
                                intFolha = !intFolha
                            End If
                        Else
                            If !intFolha = !intFolhaPorLivro Then
                                intCertidao = !intCertidao + 1
                                intLivro = !intLivro
                                If (!intQtdCertidaoUltFolha + 1) > !intCertidaoPorFolha Then
                                    intLivro = !intLivro + 1
                                    intFolha = 1
                                    intQtdCertidaoUltFolha = 1
                                Else
                                    intQtdCertidaoUltFolha = !intQtdCertidaoUltFolha + 1
                                    intFolha = !intFolha
                                End If
                            End If
                        End If
                    End With
                Else
                    ExibeMensagem "Não há parâmetros para inscrição em Dívida Ativa."
                End If
            End If
        End If
    End If
    Set gobjBanco = New clsBanco
    strSQL = "UPDATE " & gstrParametroDividaAtiva
    strSQL = strSQL & " SET intCertidao = " & intCertidao & ", intFolha = " & intFolha & ", intLivro = " & intLivro & ", intQtdCertidaoUltFolha = " & intQtdCertidaoUltFolha
    
    If blnComposicao Then
        strSQL = strSQL & " WHERE intComposicaoDaReceita = " & intComposicaoDaReceita
    Else
        strSQL = strSQL & " WHERE intComposicaoDaReceita IS NULL"
    End If
    
    strParametrosDividaAtiva = strSQL
    
End Function

Private Function blnExisteParametroAtualizacao() As Boolean

    Dim strSQL As String
    Dim adoRec As New ADODB.Recordset
    
    blnExisteParametroAtualizacao = False
    
    strSQL = "SELECT intComposicaoDaReceita "
    strSQL = strSQL & "FROM " & gstrParametroDividaAtiva
    strSQL = strSQL & " WHERE intComposicaoDaReceita = " & dbc_intReceita.BoundText
    
    Set gobjBanco = New clsBanco
    
    If gobjBanco.CriaADO(strSQL, 10, adoRec) Then
        If adoRec.EOF Then
            strSQL = "SELECT intComposicaoDaReceita "
            strSQL = strSQL & "FROM " & gstrParametroDividaAtiva
            strSQL = strSQL & " WHERE intComposicaoDaReceita is Null"
            If gobjBanco.CriaADO(strSQL, 10, adoRec) Then
                If adoRec.EOF Then
                    Exit Function
                End If
            End If
        End If
    End If
    
    blnExisteParametroAtualizacao = True
    
End Function

Private Function bnlVerificaParametros() As Boolean
    Dim strSQL          As String
    Dim adoResultado    As ADODB.Recordset
    Dim adoResultado1    As ADODB.Recordset
    
    bnlVerificaParametros = False
    
    strSQL = "SELECT "
    strSQL = strSQL & "PDA.INTCERTIDAO, "
    strSQL = strSQL & "PDA.INTFOLHA, "
    strSQL = strSQL & "PDA.INTLIVRO, "
    strSQL = strSQL & "pda.intfolhaporlivro, "
    strSQL = strSQL & "pda.intcertidaoporfolha, "
    strSQL = strSQL & "intQtdCertidaoUltFolha "
    strSQL = strSQL & "FROM "
    strSQL = strSQL & gstrParametroDividaAtiva & " PDA "
    strSQL = strSQL & "Where "
    strSQL = strSQL & "PDA.intComposicaoDaReceita = " & dbc_intReceita.BoundText & " AND "
    strSQL = strSQL & "PDA.INTCERTIDAO =" & Trim(txtintCertidao) & " AND "
    strSQL = strSQL & "PDA.INTFOLHA =" & Trim(txtintfolha) & " AND "
    strSQL = strSQL & "PDA.INTLIVRO =" & Trim(txtintlivro)
    
    Set gobjBanco = New clsBanco
    
    If gobjBanco.CriaADO(strSQL, 10, adoResultado) Then
        With adoResultado
            If Not .EOF Then
                strSQL = "SELECT "
                strSQL = strSQL & "PDA.INTCERTIDAO, "
                strSQL = strSQL & "PDA.INTFOLHA, "
                strSQL = strSQL & "PDA.INTLIVRO, "
                strSQL = strSQL & "pda.intfolhaporlivro, "
                strSQL = strSQL & "pda.intcertidaoporfolha, "
                strSQL = strSQL & "intQtdCertidaoUltFolha "
                strSQL = strSQL & "FROM "
                strSQL = strSQL & gstrParametroDividaAtiva & " PDA "
                strSQL = strSQL & "Where "
                strSQL = strSQL & "PDA.intComposicaoDaReceita = " & dbc_intReceita.BoundText
                If gobjBanco.CriaADO(strSQL, 10, adoResultado1) Then
                    With adoResultado1
                        If Not .EOF Then
                            If gblnExclusaoGravacaoOk("A", "Número de certidão " & txtintCertidao & " já se encontra cadastrado." & Chr(13) & _
                                                    "Deseja salvar D.A. com o número de certidão " & (!intCertidao + 1), True) Then
                                bnlVerificaParametros = True
                                txtintCertidao = !intCertidao + 1
                                txtintfolha = !intFolha
                                txtintlivro = !intLivro
                            Else
                                bnlVerificaParametros = False
                                Exit Function
                            End If
                        End If
                    End With
                End If
            Else
                strSQL = "SELECT "
                strSQL = strSQL & "PDA.INTCERTIDAO, "
                strSQL = strSQL & "PDA.INTFOLHA, "
                strSQL = strSQL & "PDA.INTLIVRO, "
                strSQL = strSQL & "pda.intfolhaporlivro, "
                strSQL = strSQL & "pda.intcertidaoporfolha, "
                strSQL = strSQL & "intQtdCertidaoUltFolha "
                strSQL = strSQL & "FROM "
                strSQL = strSQL & gstrParametroDividaAtiva & " PDA "
                strSQL = strSQL & "Where "
                strSQL = strSQL & "PDA.intComposicaoDaReceita Is null AND "
                strSQL = strSQL & "PDA.INTCERTIDAO =" & Trim(txtintCertidao) & " AND "
                strSQL = strSQL & "PDA.INTFOLHA =" & Trim(txtintfolha) & " AND "
                strSQL = strSQL & "PDA.INTLIVRO =" & Trim(txtintlivro)
                
                Set gobjBanco = New clsBanco
                Set adoResultado = New ADODB.Recordset
                
                If gobjBanco.CriaADO(strSQL, 10, adoResultado) Then
                    With adoResultado
                        If Not .EOF Then
                            strSQL = "SELECT "
                            strSQL = strSQL & "PDA.INTCERTIDAO, "
                            strSQL = strSQL & "PDA.INTFOLHA, "
                            strSQL = strSQL & "PDA.INTLIVRO, "
                            strSQL = strSQL & "pda.intfolhaporlivro, "
                            strSQL = strSQL & "pda.intcertidaoporfolha, "
                            strSQL = strSQL & "intQtdCertidaoUltFolha "
                            strSQL = strSQL & "FROM "
                            strSQL = strSQL & gstrParametroDividaAtiva & " PDA "
                            strSQL = strSQL & "Where "
                            strSQL = strSQL & "PDA.intComposicaoDaReceita Is null "
                            If gobjBanco.CriaADO(strSQL, 10, adoResultado1) Then
                                With adoResultado1
                                    If Not .EOF Then
                                        If gblnExclusaoGravacaoOk("A", "Número de certidão " & txtintCertidao & " já se encontra cadastrado." & Chr(13) & _
                                                                "Deseja salvar D.A. com o número de certidão " & (!intCertidao + 1), True) Then
                                            bnlVerificaParametros = True
                                            txtintCertidao = !intCertidao + 1
                                            txtintfolha = !intFolha
                                            txtintlivro = !intLivro
                                        Else
                                            bnlVerificaParametros = False
                                            Exit Function
                                        End If
                                    End If
                                End With
                            End If
                        End If
                    End With
                End If
            End If
        End With
    End If
    bnlVerificaParametros = True
    
End Function

Private Function strQueryParcela(lngLancamentoAlfa As Long) As String
Dim strSQL As String
    
    strSQL = ""
    strSQL = strSQL & "SELECT "
    strSQL = strSQL & "LV.Pkid, "
    strSQL = strSQL & "LV.Intparcela as intNumeroParcela, "
    strSQL = strSQL & "M.Strabreviatura as strMoeda, "
    strSQL = strSQL & "LV.Dtmdtvencimento as dtmVencimento, "
    strSQL = strSQL & gstrISNULL("LV.Dblvalor", 0) & " as dblValor, "
    strSQL = strSQL & gstrISNULL("LV.DBLVALOR", 0) & " as dblTotal, "
    strSQL = strSQL & "M.Pkid as intMoeda "
        
    If bytDBType = Oracle Then
        strSQL = strSQL & "FROM "
        strSQL = strSQL & gstrLancamentoAlfa & " LA, "
        strSQL = strSQL & gstrLancamentoValor & " LV, "
        strSQL = strSQL & gstrLancamentoPagamento & " LP, "
        strSQL = strSQL & gstrMoedas & " M "
        strSQL = strSQL & "WHERE "
        strSQL = strSQL & "LA.Pkid = LV.INTLANCAMENTOALFA AND "
        strSQL = strSQL & "LV.INTLANCAMENTOALFADATIVA IS NULL AND "
        strSQL = strSQL & "LV.bitParcelaValida = 1 AND "
        strSQL = strSQL & "LV.Pkid" & strOUTJSQLServer & "= LP.Intlancamentovalor" & strOUTJOracle & " AND "
        strSQL = strSQL & "LP.Intlancamentovalor is null AND "
        strSQL = strSQL & "M.Pkid" & strOUTJOracle & "=" & strOUTJSQLServer & "LV.Intmoeda AND "
        strSQL = strSQL & "LV.Dtmdtvencimento < " & gstrConvDtParaSql(IIf(txtdtmdtinscricao <> "", txtdtmdtinscricao, Date)) & " AND "
        strSQL = strSQL & "LA.Pkid = " & lngLancamentoAlfa
    Else
        strSQL = strSQL & "FROM " & gstrLancamentoAlfa & " LA INNER JOIN "
        strSQL = strSQL & gstrLancamentoValor & " LV ON LA.PKId = LV.intLancamentoAlfa LEFT OUTER JOIN "
        strSQL = strSQL & gstrLancamentoPagamento & " LP ON LV.PKId = LP.INTLANCAMENTOVALOR LEFT OUTER JOIN "
        strSQL = strSQL & gstrMoedas & " M ON LV.INTMOEDA = M.PKID "
        strSQL = strSQL & "WHERE (LV.intLancamentoAlfaDAtiva IS NULL) AND "
        strSQL = strSQL & "(LV.bitParcelaValida = 1) AND "
        strSQL = strSQL & "(LP.INTLANCAMENTOVALOR IS NULL) AND "
        strSQL = strSQL & "LV.dtmDtVencimento < " & gstrConvDtParaSql(IIf(txtdtmdtinscricao <> "", txtdtmdtinscricao, Date)) & " AND "
        strSQL = strSQL & "LA.Pkid = " & lngLancamentoAlfa
    End If
    
    strSQL = strSQL & " Order By LV.Intparcela"
    
    strQueryParcela = strSQL
    
End Function

Private Sub AdicionaTaxasImpostos(lngPkid As Long, dblTotalParcelas As Double)
Dim strSQL          As String
Dim adoResultado    As New ADODB.Recordset
Dim dblPorcImpostos As Double
Dim dblPorcTaxas    As Double

    'Query para buscar o valor dos impostos
    strSQL = "Select "
    strSQL = strSQL & "Sum(" & gstrISNULL("TT.dblImposto", "0") & ") As  dblImposto, "
    strSQL = strSQL & "Sum(" & gstrISNULL("TT.dblTaxa", "0") & "+" & gstrISNULL("TT.dblTaxa1", "0") & ") As dblTaxa "
    strSQL = strSQL & "From "
    strSQL = strSQL & "(SELECT "
    strSQL = strSQL & "LV.INTPARCELA, "
    strSQL = strSQL & "SUM(LR.DBLVALOR) dblImposto, "
    strSQL = strSQL & "0 dblTaxa, "
    strSQL = strSQL & "0 dblTaxa1"
    If bytDBType = Oracle Then
        strSQL = strSQL & " From "
        strSQL = strSQL & gstrLancamentoAlfa & " LA,"
        strSQL = strSQL & gstrLancamentoValor & " LV, "
        strSQL = strSQL & gstrLancamentoPagamento & " LP, "
        strSQL = strSQL & gstrLancamentoReceita & " LR, "
        strSQL = strSQL & gstrReceita & " RC "
        strSQL = strSQL & "Where "
        strSQL = strSQL & "LA.Pkid = LV.intLancamentoAlfa"
        strSQL = strSQL & " AND LV.PKID = LR.INTLANCAMENTOVALOR"
        strSQL = strSQL & " AND LV.Pkid " & strOUTJSQLServer & "= LP.Intlancamentovalor" & strOUTJOracle
        strSQL = strSQL & " AND LP.Intlancamentovalor is null"
        strSQL = strSQL & " AND RC.PKID = LR.INTRECEITA"
        strSQL = strSQL & " AND LA.PKID = " & Trim(lngPkid)
        strSQL = strSQL & " AND RC.BYTTIPO in(2,1,5,6)"
        strSQL = strSQL & " AND LV.BITPARCELAVALIDA = 1"
        strSQL = strSQL & " AND LV.Dtmdtvencimento < " & gstrConvDtParaSql(IIf(txtdtmdtinscricao <> "", txtdtmdtinscricao, Date))
    Else
        strSQL = strSQL & " FROM " & gstrLancamentoAlfa & " LA INNER JOIN "
        strSQL = strSQL & gstrLancamentoValor & " LV ON LA.PKId = LV.intLancamentoAlfa INNER JOIN "
        strSQL = strSQL & gstrLancamentoReceita & " LR ON LV.PKId = LR.intlancamentoValor INNER JOIN "
        strSQL = strSQL & gstrReceita & " RC ON LR.intReceita = RC.PKId LEFT OUTER JOIN "
        strSQL = strSQL & gstrLancamentoPagamento & " LP ON LV.PKId = LP.INTLANCAMENTOVALOR "
        strSQL = strSQL & " WHERE LP.INTLANCAMENTOVALOR IS NULL "
        strSQL = strSQL & " AND LA.PKId = " & Trim(lngPkid)
        strSQL = strSQL & " AND RC.bytTipo IN (2,1,5,6) "
        strSQL = strSQL & " AND LV.Dtmdtvencimento < " & gstrConvDtParaSql(IIf(txtdtmdtinscricao <> "", txtdtmdtinscricao, Date))
        strSQL = strSQL & " AND LV.bitParcelaValida = 1"
    End If
    strSQL = strSQL & " Group By LV.intParcela , LV.DBLVALOR"

    
    strSQL = strSQL & " UNION "
    
    'Query para buscar o valor das taxas
    strSQL = strSQL & "SELECT "
    strSQL = strSQL & "LV.INTPARCELA, "
    strSQL = strSQL & "0 dblImposto, "
    strSQL = strSQL & "SUM(LR.DBLVALOR) dblTaxa,"
    strSQL = strSQL & "0 dblTaxa1"
    If bytDBType = Oracle Then
        strSQL = strSQL & " From "
        strSQL = strSQL & gstrLancamentoAlfa & " LA,"
        strSQL = strSQL & gstrLancamentoValor & " LV, "
        strSQL = strSQL & gstrLancamentoPagamento & " LP, "
        strSQL = strSQL & gstrLancamentoReceita & " LR, "
        strSQL = strSQL & gstrReceita & " RC "
        strSQL = strSQL & "Where "
        strSQL = strSQL & "LA.Pkid = LV.intLancamentoAlfa"
        strSQL = strSQL & " AND LV.PKID = LR.INTLANCAMENTOVALOR"
        strSQL = strSQL & " AND LV.Pkid " & strOUTJSQLServer & "= LP.Intlancamentovalor" & strOUTJOracle
        strSQL = strSQL & " AND LP.Intlancamentovalor is null"
        strSQL = strSQL & " AND RC.PKID = LR.INTRECEITA"
        strSQL = strSQL & " AND LA.PKID = " & Trim(lngPkid)
        strSQL = strSQL & " AND RC.BYTTIPO in(3,4)"
        strSQL = strSQL & " AND LV.BITPARCELAVALIDA = 1"
        strSQL = strSQL & " AND LV.Dtmdtvencimento <  " & gstrConvDtParaSql(IIf(txtdtmdtinscricao <> "", txtdtmdtinscricao, Date))
    Else
        strSQL = strSQL & " FROM " & gstrLancamentoAlfa & " LA INNER JOIN "
        strSQL = strSQL & gstrLancamentoValor & " LV ON LA.PKId = LV.intLancamentoAlfa INNER JOIN "
        strSQL = strSQL & gstrLancamentoReceita & " LR ON LV.PKId = LR.intlancamentoValor INNER JOIN "
        strSQL = strSQL & gstrReceita & " RC ON LR.intReceita = RC.PKId LEFT OUTER JOIN "
        strSQL = strSQL & gstrLancamentoPagamento & " LP ON LV.PKId = LP.INTLANCAMENTOVALOR "
        strSQL = strSQL & " WHERE LP.INTLANCAMENTOVALOR IS NULL "
        strSQL = strSQL & " AND LA.PKID = " & Trim(lngPkid)
        strSQL = strSQL & " AND RC.bytTipo IN (3, 4) "
        strSQL = strSQL & " AND LV.BITPARCELAVALIDA = 1 "
        strSQL = strSQL & " AND LV.Dtmdtvencimento <  " & gstrConvDtParaSql(IIf(txtdtmdtinscricao <> "", txtdtmdtinscricao, Date))
    End If
    strSQL = strSQL & " Group By LV.intParcela , LV.DBLVALOR"
    
    strSQL = strSQL & " UNION "
    
    'Query para buscar o valor das taxas que os tipos de receitas não sejam (1,2,3,4,5,6)
    strSQL = strSQL & "SELECT "
    strSQL = strSQL & "LV.INTPARCELA, "
    strSQL = strSQL & "0 dblImposto, "
    strSQL = strSQL & "0 dblTaxa,"
    strSQL = strSQL & "SUM(LR.DBLVALOR) dblTaxa"
    If bytDBType = Oracle Then
        strSQL = strSQL & " From "
        strSQL = strSQL & gstrLancamentoAlfa & " LA,"
        strSQL = strSQL & gstrLancamentoValor & " LV, "
        strSQL = strSQL & gstrLancamentoPagamento & " LP, "
        strSQL = strSQL & gstrLancamentoReceita & " LR, "
        strSQL = strSQL & gstrReceita & " RC "
        strSQL = strSQL & "Where "
        strSQL = strSQL & "LA.Pkid = LV.intLancamentoAlfa"
        strSQL = strSQL & " AND LV.PKID = LR.INTLANCAMENTOVALOR"
        strSQL = strSQL & " AND LV.Pkid " & strOUTJSQLServer & "= LP.Intlancamentovalor" & strOUTJOracle
        strSQL = strSQL & " AND LP.Intlancamentovalor is null"
        strSQL = strSQL & " AND RC.PKID = LR.INTRECEITA"
        strSQL = strSQL & " AND LA.PKID = " & Trim(lngPkid)
        strSQL = strSQL & " AND not RC.BYTTIPO in(1,2,3,4,5,6)"
        strSQL = strSQL & " AND LV.BITPARCELAVALIDA = 1"
        strSQL = strSQL & " AND LV.Dtmdtvencimento <  " & gstrConvDtParaSql(IIf(txtdtmdtinscricao <> "", txtdtmdtinscricao, Date))
    Else
        strSQL = strSQL & " FROM " & gstrLancamentoAlfa & " LA INNER JOIN "
        strSQL = strSQL & gstrLancamentoValor & " LV ON LA.PKId = LV.intLancamentoAlfa INNER JOIN "
        strSQL = strSQL & gstrLancamentoReceita & " LR ON LV.PKId = LR.intlancamentoValor INNER JOIN "
        strSQL = strSQL & gstrReceita & " RC ON LR.intReceita = RC.PKId LEFT OUTER JOIN "
        strSQL = strSQL & gstrLancamentoPagamento & " LP ON LV.PKId = LP.INTLANCAMENTOVALOR "
        strSQL = strSQL & " WHERE LP.INTLANCAMENTOVALOR IS NULL "
        strSQL = strSQL & " AND LA.PKID = " & Trim(lngPkid)
        strSQL = strSQL & " AND NOT (RC.bytTipo IN (1, 2, 3, 4, 5, 6)) "
        strSQL = strSQL & " AND LV.BITPARCELAVALIDA = 1 "
        strSQL = strSQL & " AND LV.Dtmdtvencimento <  " & gstrConvDtParaSql(IIf(txtdtmdtinscricao <> "", txtdtmdtinscricao, Date))
    End If
    strSQL = strSQL & " Group By LV.intParcela , LV.DBLVALOR ) TT"
    
    Set gobjBanco = New clsBanco
    
    If gobjBanco.CriaADO(strSQL, 15, adoResultado) Then
        With adoResultado
            If Not .EOF Then
                dblValorImposto = gstrConvVrDoSql(gstrENulo(!dblImposto), 2, , True)
                dblValorTaxa = gstrConvVrDoSql(gstrENulo(!dblTaxa), 2, , True)
            End If
        End With
    End If
    
    'Obter as porcentagens de impostos e taxas
    If dblValorImposto = 0 Then
        dblPorcImpostos = 0
    Else
        dblPorcImpostos = gstrConvVrDoSql(gstrENulo(dblValorImposto / (dblValorImposto + dblValorTaxa)), 2, , True)
    End If
    
    If dblValorTaxa = 0 Then
        dblPorcTaxas = 0
    Else
        dblPorcTaxas = gstrConvVrDoSql(gstrENulo(dblValorTaxa / (dblValorImposto + dblValorTaxa)), 2, , True)
    End If
    
    'Vamos aplicar a proporcao obtida no valor da soma das parcelas
    dblValorImposto = gstrConvVrDoSql(dblTotalParcelas * dblPorcImpostos, 2, , True)
    dblValorTaxa = gstrConvVrDoSql(dblTotalParcelas * dblPorcTaxas, 2, , True)
    
End Sub

