VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmGeraDebitoAutomatico 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Geração do Arquivo de Débito Atutomático"
   ClientHeight    =   4050
   ClientLeft      =   2835
   ClientTop       =   4200
   ClientWidth     =   7410
   Icon            =   "frmGeraDebitoAutomatico.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4050
   ScaleWidth      =   7410
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab tab_Parametros 
      Height          =   3960
      Left            =   60
      TabIndex        =   21
      Top             =   60
      Width           =   7305
      _ExtentX        =   12885
      _ExtentY        =   6985
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Parâmetros"
      TabPicture(0)   =   "frmGeraDebitoAutomatico.frx":1042
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "pgr_Status"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "fra_ComposicaoDaReceita"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "CmdBanco"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      Begin MSComDlg.CommonDialog CmdBanco 
         Left            =   6570
         Top             =   630
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Frame fra_ComposicaoDaReceita 
         Height          =   2820
         Left            =   150
         TabIndex        =   0
         Top             =   480
         Width           =   6990
         Begin VB.Frame fra_Vencimento 
            Caption         =   "Vencimento"
            Height          =   675
            Left            =   1140
            TabIndex        =   10
            Top             =   1590
            Width           =   5325
            Begin VB.TextBox txt_VencFinal 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   3060
               MaxLength       =   10
               TabIndex        =   14
               Top             =   270
               Width           =   1200
            End
            Begin VB.TextBox txt_VencInicial 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   1320
               MaxLength       =   10
               TabIndex        =   12
               Top             =   270
               Width           =   1200
            End
            Begin VB.Label lblInicial 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Inicial"
               Height          =   195
               Left            =   825
               TabIndex        =   11
               Top             =   330
               Width           =   405
            End
            Begin VB.Label lblFinal 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Final:"
               Height          =   195
               Left            =   2625
               TabIndex        =   13
               Top             =   345
               Width           =   375
            End
         End
         Begin VB.CommandButton cmd_arquivo 
            Height          =   300
            Left            =   6510
            Picture         =   "frmGeraDebitoAutomatico.frx":105E
            Style           =   1  'Graphical
            TabIndex        =   17
            TabStop         =   0   'False
            Tag             =   "590"
            ToolTipText     =   "Ativa tela para escolha do Local para criação do arquivo."
            Top             =   2370
            Width           =   360
         End
         Begin VB.TextBox txt_Arquivo 
            Height          =   285
            Left            =   1080
            MaxLength       =   500
            TabIndex        =   16
            Top             =   2370
            Width           =   5370
         End
         Begin VB.CheckBox chk_AllCombos 
            Caption         =   "Todas as Composições"
            Height          =   195
            Left            =   1140
            TabIndex        =   9
            Top             =   1380
            Width           =   3315
         End
         Begin VB.TextBox txt_intSequencialDebAut 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1140
            MaxLength       =   6
            TabIndex        =   5
            Top             =   600
            Width           =   1200
         End
         Begin VB.CommandButton cmd_Banco 
            Height          =   300
            Left            =   4980
            Picture         =   "frmGeraDebitoAutomatico.frx":117C
            Style           =   1  'Graphical
            TabIndex        =   3
            TabStop         =   0   'False
            Tag             =   "590"
            ToolTipText     =   "Ativa Cadastro de Bancos"
            Top             =   210
            Width           =   405
         End
         Begin VB.CommandButton cmd_Composicao 
            Height          =   300
            Left            =   6480
            Picture         =   "frmGeraDebitoAutomatico.frx":129A
            Style           =   1  'Graphical
            TabIndex        =   8
            TabStop         =   0   'False
            Tag             =   "590"
            ToolTipText     =   "Ativa Cadastro de Composição da Receita"
            Top             =   1005
            Width           =   360
         End
         Begin MSDataListLib.DataCombo dbc_intComposicao 
            Height          =   315
            Left            =   1140
            TabIndex        =   7
            Top             =   1005
            Width           =   5295
            _ExtentX        =   9340
            _ExtentY        =   556
            _Version        =   393216
            IntegralHeight  =   0   'False
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dbcintBanco 
            Height          =   315
            Left            =   1140
            TabIndex        =   2
            Top             =   210
            Width           =   3810
            _ExtentX        =   6720
            _ExtentY        =   556
            _Version        =   393216
            MatchEntry      =   -1  'True
            Text            =   ""
         End
         Begin VB.Label lbl_arquivo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Arquivo"
            Height          =   195
            Left            =   495
            TabIndex        =   15
            Top             =   2445
            Width           =   540
         End
         Begin VB.Label lbl_intSequencialDebAut 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Sequêncial"
            Height          =   195
            Left            =   240
            TabIndex        =   4
            Top             =   690
            Width           =   795
         End
         Begin VB.Label lblintBanco 
            AutoSize        =   -1  'True
            Caption         =   "Banco"
            Height          =   195
            Left            =   570
            TabIndex        =   1
            Top             =   330
            Width           =   465
         End
         Begin VB.Label lbl_Composicao 
            AutoSize        =   -1  'True
            Caption         =   "Composição"
            Height          =   195
            Left            =   165
            TabIndex        =   6
            Top             =   1125
            Width           =   870
         End
      End
      Begin MSComctlLib.ProgressBar pgr_Status 
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   3360
         Visible         =   0   'False
         Width           =   6990
         _ExtentX        =   12330
         _ExtentY        =   344
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Height          =   225
         Left            =   6015
         TabIndex        =   20
         Top             =   3630
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   3630
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmGeraDebitoAutomatico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim mobjAux           As Object
    Dim mblnSelecionou    As Boolean
    Dim strWord           As String

Private Sub chk_AllCombos_Click()
    If chk_AllCombos.Value Then
        dbc_intComposicao.Text = ""
        TrocaCorObjeto dbc_intComposicao, True
    Else
        TrocaCorObjeto dbc_intComposicao, False
    End If
End Sub

Private Sub cmd_arquivo_Click()
    CmdBanco.ShowSave
    txt_Arquivo = CmdBanco.Filename
End Sub

Private Sub cmd_Banco_Click()
    CarregaForm frmCadBanco, dbcintBanco
End Sub

Private Sub cmd_Composicao_Click()
    CarregaForm frmCadComposicaoDaReceita, dbc_intComposicao
End Sub

Private Sub dbc_intComposicao_GotFocus()
    MarcaCampo dbc_intComposicao
End Sub

Private Sub dbc_intComposicao_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbc_intComposicao, Me, , , Shift
End Sub

Private Sub dbc_intComposicao_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbc_intComposicao
End Sub

Private Sub dbcintBanco_GotFocus()
    MarcaCampo dbcintBanco
End Sub

Private Sub Form_Activate()

    gintCodSeguranca = 1419
    
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
    
    CmdBanco.Filter = "Arquivo Texto | *.txt"
    
    dbcintBanco.Tag = gstrQueryDataComboBanco & ";strDescricao"
    dbc_intComposicao.Tag = strQueryComposicao & ";strDescricao"
    TrocaCorObjeto txt_intSequencialDebAut, True
    txt_intSequencialDebAut = intNumeroSequencial
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar
    mblnSelecionou = False
End Sub

Public Sub MantemForm(ByVal strModoOperacao As String)
    If UCase(strModoOperacao) = gstrSalvar Then
        If gblnExclusaoGravacaoOk("I", "Deseja realmente gerar Débito Automático", True) Then
            If blnDadosok Then
                GeraArquivo
                MantemForm gstrNovo
                Label1.Caption = ""
                Label2.Caption = ""
                dbcintBanco.SetFocus
            End If
        End If
    ElseIf UCase(strModoOperacao) = gstrPreencherLista Then
        PreencherListaDeOpcoes Me.ActiveControl
    ElseIf UCase(strModoOperacao) = gstrNovo Then
        Limpa_Controles Me, True, True, True, True, True
        txt_intSequencialDebAut = intNumeroSequencial
    End If
End Sub

Private Function blnDadosok() As Boolean
    
    
    blnDadosok = False
    
    If Not dbcintBanco.MatchedWithList Then
        ExibeMensagem "É necessário informar o Banco."
        dbcintBanco.SetFocus
        Exit Function
    ElseIf Val(gstrConvVrDoSql(txt_intSequencialDebAut, , , True)) <= 0 Then
        ExibeMensagem "É necessário informar o número sequêncial."
        Exit Function
    ElseIf chk_AllCombos.Value = 0 Then
        If Not dbc_intComposicao.MatchedWithList Then
            ExibeMensagem "É necessário informar a Composição da Receita."
            dbc_intComposicao.SetFocus
            Exit Function
        End If
    End If
    If Not gblnDataValida(txt_VencInicial, False) Then
        ExibeMensagem "É necessário informar uma data inicial válida."
        txt_VencInicial.SetFocus
        Exit Function
    ElseIf Not gblnDataValida(txt_VencFinal, False) Then
        ExibeMensagem "É necessário informar uma data final válida."
        txt_VencFinal.SetFocus
        Exit Function
    ElseIf CDate(txt_VencInicial) > CDate(txt_VencFinal) Then
        ExibeMensagem "É necessário informar uma data inicial menor que a data final."
        txt_VencInicial.SetFocus
        Exit Function
    End If
    
    

    If Trim(txt_Arquivo) = "" Then
        ExibeMensagem "É necessário informar um caminho e nome do arquivo para geração."
        txt_Arquivo.SetFocus
        Exit Function
    End If
    
    blnDadosok = True
    
End Function

Private Sub GeraArquivo()
    Dim strSQL              As String
    Dim adoResultado        As ADODB.Recordset
    Dim intContador         As Integer
    Dim strEmpresaAbrev     As String
    Dim dblTotal            As Double
    
'Variaveis do Banco
    Dim intBanco            As String
    Dim strsigla            As String
    Dim strDescricao        As String
    Dim strConvenioDebAut   As String

    
    pgr_Status.Value = 0
    strWord = ""
    intContador = 0
    dblTotal = 0
    
On Error GoTo Gravar
    
    Screen.MousePointer = vbArrow
    
    strSQL = "Select "
    strSQL = strSQL & " strEmpresaAbrev From " & gstrEmpresa
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSQL, 40, adoResultado) Then
        If Not adoResultado.EOF Then
            strEmpresaAbrev = gstrENulo(adoResultado!strEmpresaAbrev)
        Else
            ExibeMensagem "Não foi encontrado registro referente a prefeitura."
            Exit Sub
        End If
    End If
    
    strSQL = "Select "
    strSQL = strSQL & " intBanco, strsigla, strDescricao, strConvenioDebAut  From " & gstrBanco & " where Pkid = " & dbcintBanco.BoundText
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSQL, 40, adoResultado) Then
        If Not adoResultado.EOF Then
            intBanco = gstrENulo(adoResultado!intBanco)
            strsigla = gstrENulo(adoResultado!strsigla)
            strDescricao = gstrENulo(adoResultado!strDescricao)
            strConvenioDebAut = gstrENulo(adoResultado!strConvenioDebAut)
        Else
            ExibeMensagem "Não foi encontrado registro referentes ao banco informado."
            Exit Sub
        End If
    End If
    
    strSQL = "Select "
    strSQL = strSQL & "LA.pkid, "
    strSQL = strSQL & "LA.intExercicio, "
    strSQL = strSQL & "LA.intComposicaoDaReceita, "
    strSQL = strSQL & "LA.strEmissao, "
    strSQL = strSQL & "LA.strNumeroAviso, "
    strSQL = strSQL & "LV.pkid, "
    strSQL = strSQL & "Lv.intParcela, "
    strSQL = strSQL & "LV.dtmDtVencimento, "
    strSQL = strSQL & "LV.dblValor, "
    strSQL = strSQL & "DA.strIdentificacaoDebAut, strIdentificacaoBanco, "
    strSQL = strSQL & "DA.dtmDtOpcao, "
    strSQL = strSQL & "DA.strAgencia, "
    strSQL = strSQL & "G.STRCODBARRA "
    strSQL = strSQL & "From "
    strSQL = strSQL & gstrLancamentoAlfa & " LA " & strREADPAST & ", "
    strSQL = strSQL & gstrLancamentoValor & " LV " & strREADPAST & ", "
    strSQL = strSQL & gstrDebitoAutomatico & " DA " & strREADPAST & ", "
    strSQL = strSQL & gstrLancamentoGuias & " LG " & strREADPAST & ", "
    strSQL = strSQL & gstrGuias & " G " & strREADPAST & " "
    strSQL = strSQL & "Where "
    strSQL = strSQL & "La.pkid = LV.intLancamentoAlfa and "
    If chk_AllCombos.Value = 0 Then
        strSQL = strSQL & "LA.intComposicaoDaReceita = " & dbc_intComposicao.BoundText & " and "
    End If
    strSQL = strSQL & "Lv.pkid = LG.intLancamentovalor and "
    strSQL = strSQL & "G.pkid = LG.intguias and "
    strSQL = strSQL & "DA.strInscricaoCadastral = LA.strinscricao and "
    strSQL = strSQL & "Not DA.strAgencia is null and "
    strSQL = strSQL & "DA.intComposicaoDaReceita = LA.intComposicaoDaReceita and "
    strSQL = strSQL & strLen & "(Ltrim(Rtrim(G.STRCODBARRA))) = 28 and "
    strSQL = strSQL & "La.intExercicio = " & Year(gstrDataDoSistema) & " and "
    strSQL = strSQL & "Lv.dtmdtVencimento BetWeen " & gstrConvDtParaSql(txt_VencInicial) & " and " & gstrConvDtParaSql(txt_VencFinal) & " and "
    strSQL = strSQL & "Da.intBanco = " & intBanco & " "
    'strSQL = strSQL & "and strInscricao = '00000000000034001157' "
    strSQL = strSQL & "Order By La.strInscricao, LV.intPArcela "
    
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSQL, 60, adoResultado) Then
        With adoResultado
            If Not .EOF Then
            
                Open txt_Arquivo.Text For Output As #1
                
                pgr_Status.Visible = True
                pgr_Status.Max = Abs(.RecordCount)
                Label2.Caption = adoResultado.RecordCount
                
'INICIO DE GERAÇÃO
                'Header
                strWord = "A"
                'Código de Remessa
                strWord = strWord & "1"
                'Código do Convenio
                strWord = strWord & gstrENulo(strConvenioDebAut) & Space$(20 - Len(gstrENulo(strConvenioDebAut)))
                'Nome da Empresa
                strWord = strWord & strEmpresaAbrev & Space$(20 - Len(strEmpresaAbrev))
                'Código do Banco
                strWord = strWord & String(3 - Len(gstrENulo(intBanco)), "0") & gstrENulo(intBanco)
                'Nome do Banco
                strWord = strWord & gstrENulo(strsigla) & Space$(20 - Len(gstrENulo(strsigla)))
                'Data da Geração do Banco
                strWord = strWord & Format$(gstrDataDoSistema, "YYYYMMDD")
                'Numero Sequencial do Arquivo
                strWord = strWord & String$(6 - Len(txt_intSequencialDebAut), "0") & gstrENulo(txt_intSequencialDebAut)
                'Versão do Layout
                strWord = strWord & "04"
                'Identificação do Serviço
                strWord = strWord & "DEBITO AUTOMATICO"
                'Reservado para o Futuro
                strWord = strWord & Space$(52)
                
                Print #1, strWord
                
                Do While Not .EOF
                    'Débito em Conta Corrente
                    strWord = "E"
                    'Identificacao do Cliente na Empresa
                    strWord = strWord & gstrENulo(!strIdentificacaoDebAut) & Space$(25 - Len(gstrENulo(!strIdentificacaoDebAut)))
                    'Agencia para Débito
                    strWord = strWord & gstrENulo(!strAgencia) & Space$(4 - Len(gstrENulo(!strAgencia)))
                    'Identificacao do Cliente no Banco
                    strWord = strWord & gstrENulo(!strIdentificacaoBanco) & Space$(14 - Len(gstrENulo(!strIdentificacaoBanco)))
                    'Data do Vencimento
                    strWord = strWord & Format$(gstrENulo(!dtmDtOpcao), "YYYYMMDD")
                    'Valor do Debito
                    strWord = strWord & String$(15 - Len(Trim(CDbl(gstrConvVrDoSql(gstrENulo(!dblValor), , , True)) * 100)), "0") & Trim(CDbl(gstrConvVrDoSql(gstrENulo(!dblValor), , , True)) * 100)
                    dblTotal = Format$(dblTotal + CDbl(gstrConvVrDoSql(gstrENulo(!dblValor), , , True)), "0.00")
                    'Codigo da Moeda 01 = UFIR / 03 = REAL
                    strWord = strWord & "03"
                    'Uso da Empresa
                    strWord = strWord & gstrENulo(!STRCODBARRA) & Space$(60 - Len(gstrENulo(!STRCODBARRA)))
                    'Reservado para Futuro(filter)
                    strWord = strWord & Space(20)
                    'Codigo do Movimento
                    strWord = strWord & "0"
                    
                    Print #1, strWord
                    DoEvents
                    pgr_Status.Value = .AbsolutePosition
                    Label1.Caption = .AbsolutePosition
                    intContador = intContador + 1
                    .MoveNext
                Loop
                intContador = intContador + 2
                'Trailler
                strWord = "Z"
                'Total de Registros no Arquivo inclusive Header e Trailler
                strWord = strWord & String$(6 - Len(Trim(intContador)), "0") & intContador
                'Total de dos registros do Arquivo
                strWord = strWord & String(17 - Len(Trim(dblTotal * 100)), "0") & dblTotal * 100
                'Reservado para O Futuro(Filter)
                strWord = strWord & Space(126)
                
                Print #1, strWord
'FIM DE GERAÇÃO
                Close #1
                
                strSQL = "update " & gstrBanco & " Set intSequencialDebAut = " & txt_intSequencialDebAut.Text & " Where pkid = " & dbcintBanco.BoundText
                
                If Not gobjBanco.Execute(strSQL) Then
                    ExibeMensagem "Não foi possível atualizar o número sequêncial do Débito Automático."
                End If
                
            Else
                ExibeMensagem "Não foram encontrados registros com esses parâmetros."
                GoTo Gravar
            End If
        End With
    Else
        GoTo Gravar
    End If
    
    Screen.MousePointer = vbDefault
    
    If intContador >= 1 Then
        ExibeMensagem "Arquivo gerado com sucesso com " & (intContador - 2) & " boleto(s)."
    End If
    
    pgr_Status.Value = 0
    Exit Sub
    
Gravar:
    If Len(Err.Description) > 0 Then MsgBox Err.Description
    
    Close #1
    Screen.MousePointer = vbDefault
End Sub

Private Function strQueryComposicao() As String
    Dim strSQL As String

    strSQL = "SELECT CR.Pkid,"
    strSQL = strSQL & gstrCONVERT(CDT_VARCHAR, "CR.intCodigo") & strCONCAT & "' - '" & strCONCAT & " CR.strDescricao Descricao"
    strSQL = strSQL & " FROM "
    strSQL = strSQL & gstrComposicaoDaReceita & " CR, "
    strSQL = strSQL & "(select Distinct intComposicaoDaReceita from " & gstrDebitoAutomatico & ") DA "
    strSQL = strSQL & "Where "
    strSQL = strSQL & "CR.Pkid = DA.intComposicaoDaReceita "
    strSQL = strSQL & "Order By "
    strSQL = strSQL & "CR.intCodigo "

    strQueryComposicao = strSQL

End Function

Private Sub dbcintBanco_Click(Area As Integer)
    DropDownDataCombo dbcintBanco, Me, Area
End Sub

Private Sub dbcintBanco_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbcintBanco, Me, , KeyCode, Shift
End Sub

Private Sub dbcintBanco_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbcintBanco
End Sub

Public Function gstrQueryDataComboBanco()
    Dim strSQL As String
    
    strSQL = ""
    strSQL = strSQL & "SELECT PKId, strDescricao "
    strSQL = strSQL & "FROM " & gstrBanco & " "
    strSQL = strSQL & "Where not strConvenioDebAut Is null "
    strSQL = strSQL & "ORDER BY strDescricao"
    
    gstrQueryDataComboBanco = strSQL
    
End Function

Private Sub txt_VencInicial_GotFocus()
    If Trim(txt_VencInicial) = "" Then txt_VencInicial = gstrDataDoSistema
    MarcaCampo txt_VencInicial
End Sub

Private Sub txt_VencInicial_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "D", txt_VencInicial
End Sub

Private Sub txt_VencInicial_LostFocus()
    txt_VencInicial = gstrDataFormatada(txt_VencInicial)
End Sub

Private Sub txt_VencFinal_GotFocus()
    If Trim(txt_VencFinal) = "" Then txt_VencFinal = txt_VencInicial
    MarcaCampo txt_VencFinal
End Sub

Private Sub txt_VencFinal_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "D", txt_VencFinal
End Sub

Private Sub txt_VencFinal_LostFocus()
    txt_VencFinal = gstrDataFormatada(txt_VencFinal)
End Sub

Private Function intNumeroSequencial() As Long
    Dim strSQL          As String
    Dim adoResultado    As ADODB.Recordset
    
    strSQL = ""
    strSQL = strSQL & "Select MAx(" & gstrISNULL("intSequencialDebAut", 0) & ") + 1 intSequencialDebAut from " & gstrBanco
    
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSQL, 10, adoResultado) Then
        If Not adoResultado.EOF Then
            intNumeroSequencial = CLng(gstrENulo(adoResultado!intSequencialDebAut))
        Else
            intNumeroSequencial = 0
            ExibeMensagem "Não foi encontrado resgistro para o número sequêncial."
        End If
    End If
    
End Function
