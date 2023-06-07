VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Begin VB.Form frmGeracaoSpoolCobAmigavel 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Geração de Spool para Cobrança Amigável"
   ClientHeight    =   4920
   ClientLeft      =   3870
   ClientTop       =   3045
   ClientWidth     =   5985
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   5985
   Begin TabDlg.SSTab tab_Parametros 
      Height          =   4770
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   5835
      _ExtentX        =   10292
      _ExtentY        =   8414
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Parâmetros"
      TabPicture(0)   =   "frmGeracaoSpoolCobAmigavel.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblStatusInicial"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblStatusFinal"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblLote"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "dbcintLote"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "pgr_Status"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "CmdBanco"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "fra_Parametros"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "fra_Cobranca"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "fraOrdenacao"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "chk_GerarPeloCEP(1)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "chk_GerarPeloCEP(0)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).ControlCount=   11
      Begin VB.CheckBox chk_GerarPeloCEP 
         Caption         =   "Fora do município"
         Height          =   195
         Index           =   0
         Left            =   3030
         TabIndex        =   3
         Top             =   450
         Value           =   1  'Checked
         Width           =   2595
      End
      Begin VB.CheckBox chk_GerarPeloCEP 
         Caption         =   "Dentro do município"
         Height          =   195
         Index           =   1
         Left            =   3030
         TabIndex        =   4
         Top             =   690
         Value           =   1  'Checked
         Width           =   2595
      End
      Begin VB.Frame fraOrdenacao 
         Caption         =   "Ordenação"
         Height          =   795
         Left            =   450
         TabIndex        =   17
         Top             =   3330
         Width           =   4965
         Begin VB.OptionButton opt_Ordenacao 
            Caption         =   "Ordenação por CEP"
            Height          =   195
            Index           =   1
            Left            =   90
            TabIndex        =   19
            Top             =   510
            Width           =   3075
         End
         Begin VB.OptionButton opt_Ordenacao 
            Caption         =   "Ordenação por Identificação"
            Height          =   195
            Index           =   0
            Left            =   90
            TabIndex        =   18
            Top             =   240
            Width           =   3075
         End
      End
      Begin VB.Frame fra_Cobranca 
         Caption         =   "Dados para Cobrança"
         Height          =   1215
         Left            =   450
         TabIndex        =   10
         Top             =   2010
         Width           =   4935
         Begin VB.OptionButton opt_Cobranca 
            Caption         =   "Ficha Compensação"
            Height          =   195
            Index           =   1
            Left            =   90
            TabIndex        =   12
            Top             =   510
            Width           =   3075
         End
         Begin VB.OptionButton opt_Cobranca 
            Caption         =   "Febraban"
            Height          =   195
            Index           =   0
            Left            =   90
            TabIndex        =   11
            Top             =   240
            Width           =   3075
         End
         Begin MSDataListLib.DataCombo dbcintBanco 
            Height          =   315
            Left            =   630
            TabIndex        =   14
            Top             =   780
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dbcstrAgencia 
            Height          =   315
            Left            =   3000
            TabIndex        =   16
            Top             =   780
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin VB.Label lblAgencia 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Agência-Conta"
            Height          =   195
            Left            =   1905
            TabIndex        =   15
            Top             =   870
            Width           =   1050
         End
         Begin VB.Label lblBanco 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Banco"
            Height          =   195
            Left            =   90
            TabIndex        =   13
            Top             =   870
            Width           =   465
         End
      End
      Begin VB.Frame fra_Parametros 
         Caption         =   "Faixa de nº de Cobrança Amigável"
         Height          =   840
         Left            =   450
         TabIndex        =   5
         Top             =   1050
         Width           =   4935
         Begin VB.TextBox txt_strAvisoF 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   3180
            MaxLength       =   9
            TabIndex        =   9
            Top             =   330
            Width           =   1200
         End
         Begin VB.TextBox txt_strAvisoI 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1080
            MaxLength       =   9
            TabIndex        =   7
            Top             =   330
            Width           =   1200
         End
         Begin VB.Label lblInicial 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Inicial"
            Height          =   195
            Left            =   615
            TabIndex        =   6
            Top             =   390
            Width           =   405
         End
         Begin VB.Label lblFinal 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Final"
            Height          =   195
            Left            =   2790
            TabIndex        =   8
            Top             =   405
            Width           =   330
         End
      End
      Begin MSComDlg.CommonDialog CmdBanco 
         Left            =   4890
         Top             =   1500
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSComctlLib.ProgressBar pgr_Status 
         Height          =   195
         Left            =   450
         TabIndex        =   20
         Top             =   4260
         Visible         =   0   'False
         Width           =   4965
         _ExtentX        =   8758
         _ExtentY        =   344
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin MSDataListLib.DataCombo dbcintLote 
         Height          =   315
         Left            =   1350
         TabIndex        =   2
         Top             =   510
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin VB.Label lblLote 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Lote"
         Height          =   195
         Left            =   930
         TabIndex        =   1
         Top             =   600
         Width           =   315
      End
      Begin VB.Label lblStatusFinal 
         Alignment       =   1  'Right Justify
         Height          =   225
         Left            =   4320
         TabIndex        =   22
         Top             =   4470
         Width           =   1095
      End
      Begin VB.Label lblStatusInicial 
         Height          =   225
         Left            =   480
         TabIndex        =   21
         Top             =   4470
         Width           =   915
      End
   End
End
Attribute VB_Name = "frmGeracaoSpoolCobAmigavel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strWord           As String

Dim intCepInicial       As Long
Dim intCepFinal         As Long


Private Sub dbcintBanco_Change()
    
    dbcstrAgencia.Text = ""
    dbcstrAgencia.ListField = ""

    If dbcintBanco.MatchedWithList Then
        LeDaTabelaParaObj gstrContaBancaria, dbcstrAgencia, strQueryAgencias(dbcintBanco.BoundText)
    End If
    
End Sub

Private Sub dbcintBanco_Click(Area As Integer)
    If Area = 0 Then DropDownDataCombo dbcintBanco, Me, Area
End Sub

Private Sub dbcintBanco_GotFocus()
    MarcaCampo dbcintBanco
End Sub

Private Sub dbcintBanco_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbcintBanco, Me, , KeyCode, Shift
End Sub

Private Sub dbcintBanco_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbcintBanco
End Sub

Private Sub Form_Activate()
    
    gintCodSeguranca = 1440
    
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo
    
    HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrSalvar
    
End Sub

Private Sub Form_Deactivate()
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar
End Sub

Private Sub Form_Load()
    
    CmdBanco.Filter = "Arquivo Texto | *.txt"
    
    opt_Cobranca(0).Value = True
    opt_Ordenacao(0).Value = True
    
    dbcintLote.Tag = "SELECT 0, intLote FROM tblLancamentoCobAmigavelAlfa WHERE intGuias IS NULL GROUP BY intLote ORDER BY intLote DESC;intLote"
    dbcintBanco.Tag = "SELECT Pkid, strSigla FROM " & gstrBanco & " ORDER BY strSigla;strSigla"
    
    ObterFaixaDeCeps
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar
End Sub

Public Sub MantemForm(ByVal strModoOperacao As String)
    
    Select Case strModoOperacao
    
        Case Is = UCase(gstrPreencherLista)
                PreencherListaDeOpcoes Me.ActiveControl
        Case Is = UCase(gstrSalvar)
            If blnDadosOk Then
                GeraArquivo
            End If
            
    End Select
    
End Sub

Private Function blnDadosOk() As Boolean
    
    blnDadosOk = False
    If Not dbcintLote.MatchedWithList Then
        ExibeMensagem "É necessário selecionar algum Lote"
        dbcintLote.SetFocus
        Exit Function
    ElseIf Trim(txt_strAvisoI) = "" Then
        ExibeMensagem "Número de aviso inicial foi preenchido incorretamente."
        txt_strAvisoI.SetFocus
        Exit Function
    ElseIf CDbl(txt_strAvisoI) <= 0 Then
        ExibeMensagem "Número de aviso inicial foi preenchido incorretamente."
        txt_strAvisoI.SetFocus
        Exit Function
    ElseIf Trim(txt_strAvisoF) = "" Then
        ExibeMensagem "Número de aviso final foi preenchido incorretamente."
        txt_strAvisoF.SetFocus
        Exit Function
    ElseIf CDbl(txt_strAvisoF) <= 0 Then
        ExibeMensagem "Número de aviso final foi preenchido incorretamente."
        txt_strAvisoF.SetFocus
        Exit Function
    ElseIf opt_Ordenacao(0).Value = False And opt_Ordenacao(1).Value = False Then
        ExibeMensagem "É necessário preenchido da ordenação."
        opt_Ordenacao(0).SetFocus
        Exit Function
    ElseIf chk_GerarPeloCEP(0).Value = 0 And chk_GerarPeloCEP(1).Value = 0 Then
        ExibeMensagem "É necessário escolher uma opção de munícipio."
        chk_GerarPeloCEP(0).SetFocus
        Exit Function
    End If
    
    If opt_Cobranca(1).Value = True Then
        If Not dbcintBanco.MatchedWithList Then
            ExibeMensagem "É necessário selecionar algum Banco"
            dbcintBanco.SetFocus
            Exit Function
        ElseIf Not dbcstrAgencia.MatchedWithList Then
            ExibeMensagem "É necessário selecionar alguma Agência/Conta"
            dbcstrAgencia.SetFocus
            Exit Function
        End If
    End If
    
    blnDadosOk = True
    
End Function

Private Sub GeraArquivo()

Dim intFebraban         As Integer

Dim strSql              As String

Dim adoLancamentoAlfa   As ADODB.Recordset
Dim adoResultado        As ADODB.Recordset
Dim adoBanco            As ADODB.Recordset
Dim adoParcelas         As ADODB.Recordset
Dim adoParcelasCob      As ADODB.Recordset

Dim lngLinha            As Long
Dim intContador         As Double

Dim strComposicaoReceita As String
Dim intExercicio        As Integer
Dim strParcelas         As String
Dim dblValorPrincipal   As Variant
Dim dblValorMulta       As Variant
Dim dblValorJuros       As Variant
Dim dblValorCorrecao    As Variant

Dim INTNUMERO           As Long
Dim strCodBarras        As String
Dim strNumeroBoleto1    As String

Dim lngGuias            As Double
Dim lngPkidAlfa         As Long

    pgr_Status.Value = 0
    
    Screen.MousePointer = vbArrow
    
On Error GoTo Gravar
    
    strSql = ""
    strSql = " SELECT LCA.Pkid PkidCAAlfa, LCA.strNome, LCA.intNumeroCA, LCA.intLote, LCA.dtmDtVencimentoCoaAmigavel, LCA.strLogradouro, LCA.strNumero, LCA.strComplemento, LCA.strBairro, LCA.strMunicipio, " & _
             " LCA.strUF, LCA.intCep, LCA.strLogradouroC, LCA.strNumeroC, LCA.strComplementoC, LCA.strBairroC, LCA.strMunicipioC, LCA.strUFC, LCA.intCepC, " & _
             " LA.strInscricao, LCA.dblVlTotalPrincipal, LCA.dblVlTotalMulta, LCA.dblVlTotalJuros, LCA.dblVlTotalCorrecao "
    strSql = strSql & " FROM tblLancamentoCobAmigavelAlfa LCA, tblLancamentoCobAmigavelValor LCV, " & gstrLancamentoValor & " LV, " & gstrLancamentoAlfa & " LA "
    strSql = strSql & " WHERE LCA.intLote = " & dbcintLote.Text
    strSql = strSql & " AND LCA.intNumeroCA BETWEEN " & txt_strAvisoI & " And " & txt_strAvisoF
    strSql = strSql & " AND LCA.IntGuias IS NULL "
    strSql = strSql & " AND LCV.IntCobAmigavel = LCA.Pkid "
    strSql = strSql & " AND LCV.intLancamentoValor = LV.Pkid "
    strSql = strSql & " AND LV.intLancamentoAlfa = LA.Pkid "
    If chk_GerarPeloCEP(0).Value = vbChecked And chk_GerarPeloCEP(1).Value = vbUnchecked Then
        strSql = strSql & " AND LCA.intCepC NOT BETWEEN " & intCepInicial & " AND " & intCepFinal
    ElseIf chk_GerarPeloCEP(1).Value = vbChecked And chk_GerarPeloCEP(0).Value = vbUnchecked Then
        strSql = strSql & " AND LCA.intCepC BETWEEN " & intCepInicial & " AND " & intCepFinal
    End If
    strSql = strSql & " GROUP BY LCA.Pkid, LCA.strNome, LCA.intNumeroCA, LCA.intLote, LCA.dtmDtVencimentoCoaAmigavel, LCA.strLogradouro, LCA.strNumero, LCA.strComplemento, LCA.strBairro, LCA.strMunicipio, " & _
             " LCA.strUF, LCA.intCep, LCA.strLogradouroC, LCA.strNumeroC, LCA.strComplementoC, LCA.strBairroC, LCA.strMunicipioC, LCA.strUFC, LCA.intCepC, " & _
             " LA.strInscricao, LCA.dblVlTotalPrincipal, LCA.dblVlTotalMulta, LCA.dblVlTotalJuros, LCA.dblVlTotalCorrecao "
    
    If opt_Ordenacao(0).Value Then
        strSql = strSql & " ORDER BY LA.strInscricao "
    Else
        strSql = strSql & " ORDER BY LCA.intCepC, LA.strInscricao "
    End If
    
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 40, adoLancamentoAlfa) Then
            
        If Not adoLancamentoAlfa.EOF Then
            
            CmdBanco.ShowSave
            If Trim(CmdBanco.Filename) = "" Then Exit Sub
            
            Open CmdBanco.Filename For Output As #1
            pgr_Status.Visible = True
            pgr_Status.Max = Abs(adoLancamentoAlfa.RecordCount)
            lblStatusFinal.Caption = adoLancamentoAlfa.RecordCount
            
            'Query utilizada para retornar dados da empresa
            strSql = "SELECT EM.strNomeFantasia, EM.intFebraban, " & _
                            " TL.strSigla " & strCONCAT & "' '" & strCONCAT & " LO.strDescricao strLogradouro, " & _
                            " EM.strNumero, " & _
                            " BA.strDescricao strBairro, " & _
                            " EM.strComplemento, " & _
                            " EM.intCep, " & _
                            " MU.strDescricao strCidade, " & _
                            " UF.strsigla STRUF "
            strSql = strSql & " FROM " & gstrEmpresa & " EM, " & _
                            gstrTipoLogradouro & " TL, " & _
                            gstrLogradouro & " LO, " & _
                            gstrBairro & " BA, " & _
                            gstrCidade & " MU, " & _
                            gstrUF & " UF "
            strSql = strSql & " WHERE TL.Pkid = EM.intTipoLogradouro And " & _
                            " LO.Pkid = EM.intLogradouro And " & _
                            " BA.Pkid = EM.intBairro And " & _
                            " MU.Pkid = EM.intCidade And " & _
                            " UF.Pkid = EM.intUF "

            Set gobjBanco = New clsBanco
            If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
                If gstrENulo(adoResultado!intFebraban) <> "" Then
                    intFebraban = gstrENulo(adoResultado!intFebraban)
                Else
                    ExibeMensagem "Código Febraban não encontrado."
                    GoTo Gravar
                End If
            End If
            
            With adoResultado
            
            lngLinha = 1
            
            '********** REGISTRO TIPO H **********
            strWord = ""
            'Tipo do Registro
            strWord = strWord & "H"
            'Nome da Prefeitura
            strWord = strWord & Left(gstrENulo(!strNomeFantasia), 60) & Space$(60 - Len(Left(gstrENulo(!strNomeFantasia), 60)))
            'Logradouro da Prefeitura
            strWord = strWord & Left(gstrENulo(!strLogradouro), 100) & Space$(100 - Len(Left(gstrENulo(!strLogradouro), 100)))
            'Numero da Prefeitura
            strWord = strWord & Left(gstrENulo(!strNumero), 10) & Space$(10 - Len(Left(gstrENulo(!strNumero), 10)))
            'Bairro da Prefeitura
            strWord = strWord & Left(gstrENulo(!strBairro), 50) & Space$(50 - Len(Left(gstrENulo(!strBairro), 50)))
            'Complemento da Prefeitura
            strWord = strWord & Left(gstrENulo(!STRCOMPLEMENTO), 20) & Space$(20 - Len(Left(gstrENulo(!STRCOMPLEMENTO), 20)))
            'Cep da Prefeitura
            strWord = strWord & Format$(IIf(Trim(gstrENulo(!INTCEP)) = "", "0", Trim(gstrENulo(!INTCEP))), "00000000")
            'Municipio da Prefeitura
            strWord = strWord & Left(gstrENulo(!strCidade), 50) & Space$(50 - Len(Left(gstrENulo(!strCidade), 50)))
            'UF da Prefeitura
            strWord = strWord & Left(gstrENulo(!STRUF), 2) & Space$(2 - Len(Left(gstrENulo(!STRUF), 2)))
            'Data Geracao Spool
            strWord = strWord & Format$(gstrDataDoSistema, "dd/mm/yyyy")
            'Numero do Lote
            strWord = strWord & Format$(dbcintLote.Text, "00000")
            'Cobranca Amigavel Inicial
            strWord = strWord & Format$(txt_strAvisoI.Text, "000000000")
            'Cobranca Amigavel Final
            strWord = strWord & Format$(txt_strAvisoF.Text, "000000000")
                        
            If opt_Cobranca(1).Value Then
                If gobjBanco.CriaADO("SELECT intBanco, intDigitoBanco FROM " & gstrBanco & " WHERE Pkid = " & dbcintBanco.BoundText, 5, adoBanco) Then
                    'Numero do Banco
                    strWord = strWord & Format$(IIf(Trim(gstrENulo(adoBanco("intBanco").Value)) = "", "0", Trim(gstrENulo(adoBanco("intBanco").Value))), "000")
                    'Digito do Banco
                    strWord = strWord & Format$(IIf(Trim(gstrENulo(adoBanco("intDigitoBanco").Value)) = "", "0", Trim(gstrENulo(adoBanco("intDigitoBanco").Value))), "0")
                End If
                adoBanco.Close: Set adoBanco = Nothing
            Else
                'Numero do Banco
                strWord = strWord & "000"
                'Digito do Banco
                strWord = strWord & "0"
            End If
            
            'Filler
            strWord = strWord & String(447, " ")
            'Numero da Linha
            strWord = strWord & Format$(lngLinha, "00000")
            
            End With
            
            adoResultado.Close: Set adoResultado = Nothing
            
            Print #1, strWord
            
            Do While Not adoLancamentoAlfa.EOF
                                    
                With adoLancamentoAlfa
                
                lngLinha = lngLinha + 1
                
                '********** REGISTRO TIPO 1 **********
                strWord = ""
                'Tipo do Registro
                strWord = strWord & "1"
                'Inscricao
                strWord = strWord & Left(gstrENulo(!strInscricao), 20) & Space$(20 - Len(Left(gstrENulo(!strInscricao), 20)))
                'Proprietario
                strWord = strWord & Left(gstrENulo(!STRNOME), 100) & Space$(100 - Len(Left(gstrENulo(!STRNOME), 100)))
                'Logradouro
                strWord = strWord & Left(gstrENulo(!strLogradouro), 100) & Space$(100 - Len(Left(gstrENulo(!strLogradouro), 100)))
                'Numero
                strWord = strWord & Left(gstrENulo(!strNumero), 10) & Space$(10 - Len(Left(gstrENulo(!strNumero), 10)))
                'Bairro
                strWord = strWord & Left(gstrENulo(!strBairro), 50) & Space$(50 - Len(Left(gstrENulo(!strBairro), 50)))
                'Complemento
                strWord = strWord & Left(gstrENulo(!STRCOMPLEMENTO), 20) & Space$(20 - Len(Left(gstrENulo(!STRCOMPLEMENTO), 20)))
                'Cep
                strWord = strWord & Format$(IIf(Trim(gstrENulo(!INTCEP)) = "", "0", Trim(gstrENulo(!INTCEP))), "00000000")
                'Logradouro Correspondencia
                strWord = strWord & Left(gstrENulo(!strLogradouroC), 100) & Space$(100 - Len(Left(gstrENulo(!strLogradouroC), 100)))
                'Numero Correspondencia
                strWord = strWord & Left(gstrENulo(!strNumeroC), 10) & Space$(10 - Len(Left(gstrENulo(!strNumeroC), 10)))
                'Bairro Correspondencia
                strWord = strWord & Left(gstrENulo(!strBairroC), 50) & Space$(50 - Len(Left(gstrENulo(!strBairroC), 50)))
                'Complemento Correspondencia
                strWord = strWord & Left(gstrENulo(!strComplementoC), 30) & Space$(30 - Len(Left(gstrENulo(!strComplementoC), 30)))
                'Cep Correspondencia
                strWord = strWord & Format$(IIf(Trim(gstrENulo(!INTCEPC)) = "", "0", Trim(gstrENulo(!INTCEPC))), "00000000")
                'Municipio Correspondencia
                strWord = strWord & Left(gstrENulo(!strMunicipioC), 50) & Space$(50 - Len(Left(gstrENulo(!strMunicipioC), 50)))
                'UF Correspondencia
                strWord = strWord & Left(gstrENulo(!strUFC), 2) & Space$(2 - Len(Left(gstrENulo(!strUFC), 2)))
                        
                INTNUMERO = glngRetornaProximoNumeroGuia
                
                'Vamos definir o codigo de barras
                strCodBarras = gstrMontaCodigoBarras(IIf(opt_Cobranca(0).Value, FEBRABAN, FICHA_COMPENSACAO), IIf(opt_Cobranca(1).Value, dbcstrAgencia.BoundText, 0), gstrConvVrDoSql(!dblVlTotalPrincipal + !dblVlTotalMulta + !dblVlTotalJuros + !dblVlTotalCorrecao, 2), adoLancamentoAlfa!DtmdtVencimentoCoaAmigavel, intFebraban, INTNUMERO, True, True)
                
                If Len(strCodBarras) = 0 Then
                    gobjBanco.ExecutaRollbackTrans
                    GoTo Gravar
                End If
                'Vamos definir a linha digitavel
                strNumeroBoleto1 = gstrMontaLinhaDigitavel(IIf(opt_Cobranca(0).Value, FEBRABAN, FICHA_COMPENSACAO), strCodBarras)
                strNumeroBoleto1 = Replace(strNumeroBoleto1, "-", "")
                
                'Vamos tratar a linha digitavel de acordo com o tipo de codigo de barras
                If opt_Cobranca(1).Value Then
                    strNumeroBoleto1 = Replace(Replace(strNumeroBoleto1, " ", ""), ".", "")
                End If
                
                'Vamos inserir a Tblguias e TbllancamentoGuia
                
                gobjBanco.ExecutaBeginTrans
                
                'Vamos inserir a guia na tabela TblGuias
                strSql = ""
                strSql = strSql & "Insert Into " & gstrGuias & "("
                strSql = strSql & "Intcontabancaria, "
                strSql = strSql & "Intnumero, "
                strSql = strSql & "Dtmdtemissao, "
                strSql = strSql & "Dblvalor, "
                strSql = strSql & "Strcodbarra, "
                strSql = strSql & "Dtmdtatualizacao, "
                strSql = strSql & "Lngcodusr, "
                strSql = strSql & "Dtmdtvencimento "
                strSql = strSql & ") Values("
                strSql = strSql & gstrENulo(IIf(opt_Cobranca(1).Value, dbcintBanco.BoundText, Null), , True) & ", "
                strSql = strSql & INTNUMERO & ", "
                strSql = strSql & gstrConvDtParaSql(gstrDataDoSistema) & ", "
                strSql = strSql & gstrConvVrParaSql(!dblVlTotalPrincipal + !dblVlTotalMulta + !dblVlTotalJuros + !dblVlTotalCorrecao) & ", '"
                strSql = strSql & strCodBarras & "', "
                strSql = strSql & gstrConvDtParaSql(gstrDataDoSistema) & ", "
                strSql = strSql & glngCodUsr & ", "
                strSql = strSql & gstrConvDtParaSql(!DtmdtVencimentoCoaAmigavel)
                strSql = strSql & ")"
               
                If Not gobjBanco.Execute(strSql) Then
                    ExibeMensagem "Erro na gravação da guia."
                    gobjBanco.ExecutaRollbackTrans
                    GoTo Gravar
                End If

                'Vamos inserir as parcelas na tabela TblLancamentoGuias
                
                lngGuias = glngRetornaPkidTabelaPai("seqtblGuias", "tblGuias")
                
                'Vamos preencher as colunas de parcelas
                strSql = ""
                strSql = strSql & " SELECT  LCV.dblVlPrincipal, LCV.dblVlMulta, LCV.dblVlJuros, LCV.dblVlCorrecao, LCV.intLancamentoValor "
                strSql = strSql & " FROM "
                strSql = strSql & " tblLancamentoCobAmigavelValor LCV " & strREADPAST
                strSql = strSql & " WHERE "
                strSql = strSql & "LCV.IntCobAmigavel = " & !PkidCAAlfa
                
                Set gobjBanco = New clsBanco
                If gobjBanco.CriaADO(strSql, 10, adoParcelasCob) Then
                    If Not .EOF Then
                        
                        Set gobjBanco = New clsBanco
                        
                        Do While Not adoParcelasCob.EOF
                            
                            strSql = "Insert Into " & gstrLancamentoGuias & "("
                            strSql = strSql & "intlancamentovalor, "
                            strSql = strSql & "intguias, "
                            strSql = strSql & "dblvalorprincipal, "
                            strSql = strSql & "dblvalormulta, "
                            strSql = strSql & "dblvalorjuros, "
                            strSql = strSql & "dblvalorcorrecao, "
                            strSql = strSql & "dblvalordesconto, "
                            strSql = strSql & "dtmdtatualizacao, "
                            strSql = strSql & "lngcodusr) "
                            strSql = strSql & "Values ("
                            strSql = strSql & adoParcelasCob!intLancamentoValor & ", "
                            strSql = strSql & lngGuias & ","
                            strSql = strSql & gstrConvVrParaSql(adoParcelasCob!dblVlPrincipal) & ", "
                            strSql = strSql & gstrConvVrParaSql(adoParcelasCob!dblVlMulta) & ", "
                            strSql = strSql & gstrConvVrParaSql(adoParcelasCob!dblVlJuros) & ", "
                            strSql = strSql & gstrConvVrParaSql(adoParcelasCob!dblVlCorrecao) & ", "
                            strSql = strSql & "0.00" & ", "
                            strSql = strSql & strGETDATE & ", "
                            strSql = strSql & glngCodUsr & ") "
                            
                            If Not gobjBanco.Execute(strSql) Then
                                ExibeMensagem "Erro na gravação das parcelas da guia."
                                gobjBanco.ExecutaRollbackTrans
                                GoTo Gravar
                            End If
                            
                            adoParcelasCob.MoveNext
                        
                        Loop
                        
                    Else
                        ExibeMensagem "Não foram encontradas parcelas para algum dos Lançamentos."
                        GoTo Gravar
                    End If
                End If
                        
                adoParcelasCob.Close: Set adoParcelasCob = Nothing
                
                strSql = "UPDATE tblLancamentoCobAmigavelAlfa SET intGuias = " & lngGuias & " WHERE Pkid = " & !PkidCAAlfa
                
                If Not gobjBanco.Execute(strSql) Then
                    ExibeMensagem "Erro na gravação da guia."
                    gobjBanco.ExecutaRollbackTrans
                    GoTo Gravar
                End If
                
                gobjBanco.ExecutaCommitTrans
                
                'Codigo de Barras
                strWord = strWord & strCodBarras
                'Linha Digitavel
                strWord = strWord & Left(gstrENulo(strNumeroBoleto1), 51) & Space$(51 - Len(Left(gstrENulo(strNumeroBoleto1), 51)))
                'Filler
                strWord = strWord & String(131, " ")
                'Numero da Linha
                strWord = strWord & Format$(lngLinha, "00000")
                
                Print #1, strWord
                
                'Vamos preencher as colunas de parcelas
                strSql = ""
                strSql = strSql & " SELECT  LCV.*, "
                strSql = strSql & " LV.intParcela, LA.strComposicaoDaReceita, LA.intExercicio, LV.intLancamentoAlfa PkidAlfa "
                strSql = strSql & " FROM "
                strSql = strSql & " tblLancamentoCobAmigavelValor LCV " & strREADPAST & ", "
                strSql = strSql & gstrLancamentoValor & " LV " & strREADPAST & ", "
                strSql = strSql & gstrLancamentoAlfa & " LA " & strREADPAST
                strSql = strSql & " WHERE "
                strSql = strSql & "LCV.IntCobAmigavel = " & !PkidCAAlfa
                strSql = strSql & " AND LCV.intLancamentoValor = LV.Pkid "
                strSql = strSql & " AND LV.intLancamentoAlfa = LA.Pkid "
                strSql = strSql & " ORDER BY "
                strSql = strSql & "LV.Intlancamentoalfa, "
                strSql = strSql & "LV.Bitparcelavalida, "
                strSql = strSql & "LV.Intparcela "
                
                Set gobjBanco = New clsBanco
                If gobjBanco.CriaADO(strSql, 10, adoParcelas) Then
                    If Not .EOF Then
                        
                        Set gobjBanco = New clsBanco
                        
ProximoAlfa:
                        
                        strComposicaoReceita = adoParcelas("strComposicaoDaReceita").Value
                        intExercicio = adoParcelas("intExercicio").Value

                        lngPkidAlfa = adoParcelas("PkidAlfa").Value
                        strParcelas = ""
                        dblValorPrincipal = 0
                        dblValorMulta = 0
                        dblValorJuros = 0
                        dblValorCorrecao = 0
                        
                        Do While lngPkidAlfa = adoParcelas("PkidAlfa").Value
                            
                            'Vamos agregando as parcelas
                            strParcelas = strParcelas & Format(adoParcelas("intParcela").Value, "000") & "/"
                            
                            dblValorPrincipal = dblValorPrincipal + adoParcelas!dblVlPrincipal
                            dblValorMulta = dblValorMulta + adoParcelas!dblVlMulta
                            dblValorJuros = dblValorJuros + adoParcelas!dblVlJuros
                            dblValorCorrecao = dblValorCorrecao + adoParcelas!dblVlCorrecao
                            
                            adoParcelas.MoveNext
                            
                            If adoParcelas.EOF Then Exit Do
                            
                        Loop
                        
                        strParcelas = Mid(strParcelas, 1, Len(strParcelas) - 1)
                        
                        lngLinha = lngLinha + 1
                        
                        '********** REGISTRO TIPO 2 **********
                        strWord = ""
                        'Tipo do Registro
                        strWord = strWord & "2"
                        'Composição da Receita
                        strWord = strWord & Left(gstrENulo(strComposicaoReceita), 100) & Space$(100 - Len(Left(gstrENulo(strComposicaoReceita), 100)))
                        'Exercicio
                        strWord = strWord & Format$(IIf(Trim(gstrENulo(intExercicio)) = "", "0", Trim(gstrENulo(intExercicio))), "0000")
                        'Parcelas
                        strWord = strWord & Left(gstrENulo(strParcelas), 47) & Space$(47 - Len(Left(gstrENulo(strParcelas), 47)))
                        'Valor Principal
                        strWord = strWord & Space$(19 - Len(Left(gstrConvVrDoSql(dblValorPrincipal, 2), 19))) & Left(gstrConvVrDoSql(dblValorPrincipal, 2), 19)
                        'Valor Multa
                        strWord = strWord & Space$(19 - Len(Left(gstrConvVrDoSql(dblValorMulta, 2), 19))) & Left(gstrConvVrDoSql(dblValorMulta, 2), 19)
                        'Valor Juros
                        strWord = strWord & Space$(19 - Len(Left(gstrConvVrDoSql(dblValorJuros, 2), 19))) & Left(gstrConvVrDoSql(dblValorJuros, 2), 19)
                        'Valor Correcao
                        strWord = strWord & Space$(19 - Len(Left(gstrConvVrDoSql(dblValorCorrecao, 2), 19))) & Left(gstrConvVrDoSql(dblValorCorrecao, 2), 19)
                        'Valor Total
                        strWord = strWord & Space$(19 - Len(Left(gstrConvVrDoSql(dblValorPrincipal + dblValorMulta + dblValorJuros + dblValorCorrecao, 2), 19))) & Left(gstrConvVrDoSql(dblValorPrincipal + dblValorMulta + dblValorJuros + dblValorCorrecao, 2), 19)
                        'Data Vencimento
                        strWord = strWord & Format$(!DtmdtVencimentoCoaAmigavel, "dd/mm/yyyy")
                        'Lote
                        strWord = strWord & Format$(IIf(Trim(gstrENulo(!intLote)) = "", "0", Trim(gstrENulo(!intLote))), "00000")
                        'Numero Cobranca
                        strWord = strWord & Format$(IIf(Trim(gstrENulo(!intNumeroCA)) = "", "0", Trim(gstrENulo(!intNumeroCA))), "000000000")
                        'Filler
                        strWord = strWord & String(514, " ")
                        'Numero da Linha
                        strWord = strWord & Format$(lngLinha, "00000")
                                
                        Print #1, strWord
                        
                        If Not adoParcelas.EOF Then GoTo ProximoAlfa
                        
                    Else
                        ExibeMensagem "Não foram encontradas parcelas para algum dos Lançamentos."
                        GoTo Gravar
                    End If
                End If
                
                adoParcelas.Close: Set adoParcelas = Nothing
                
                DoEvents
                pgr_Status.Value = adoLancamentoAlfa.AbsolutePosition
                lblStatusInicial.Caption = adoLancamentoAlfa.AbsolutePosition
                intContador = intContador + 1
                gobjBanco.ExecutaCommitTrans
                adoLancamentoAlfa.MoveNext
                    
                End With
                
            Loop
            
            lngLinha = lngLinha + 1
            
            '********** REGISTRO TIPO Z **********
            strWord = ""
            'Tipo do Registro
            strWord = strWord & "Z"
            'Qtde Cobrancas
            strWord = strWord & Format$(IIf(Trim(gstrENulo(intContador)) = "", "0", Trim(gstrENulo(intContador))), "0000000")
            'Qtde Linhas
            strWord = strWord & Format$(IIf(Trim(gstrENulo(lngLinha)) = "", "0", Trim(gstrENulo(lngLinha))), "00000000")
            'Filler
            strWord = strWord & String(769, " ")
            'Numero da Linha
            strWord = strWord & Format$(lngLinha, "00000")
            
            Print #1, strWord
            
            Close #1
            
        Else
            ExibeMensagem "Não foram encontrados registros com esses parâmetros."
            GoTo Gravar
        End If

    Else
        GoTo Gravar
    End If
    
    Screen.MousePointer = vbDefault
    
    If intContador >= 1 Then
        ExibeMensagem "Arquivo gerado com sucesso com " & intContador & " boleto(s)."
    End If
    
    pgr_Status.Value = 0
    pgr_Status.Visible = False
    lblStatusInicial.Caption = ""
    lblStatusFinal.Caption = ""
    
    Exit Sub
    
Gravar:
    gobjBanco.ExecutaRollbackTrans
    If Len(Err.Description) > 0 Then MsgBox Err.Description
    Close #1
    pgr_Status.Visible = False
    lblStatusInicial.Caption = ""
    lblStatusFinal.Caption = ""
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub opt_Cobranca_Click(Index As Integer)
    TrocaCorObjeto dbcintBanco, Index = 0, True
    TrocaCorObjeto dbcstrAgencia, Index = 0, True
End Sub

Private Sub txt_strAvisoI_GotFocus()
    MarcaCampo txt_strAvisoI
End Sub

Private Sub txt_strAvisoI_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txt_strAvisoI
End Sub

Private Sub txt_strAvisoF_GotFocus()
    MarcaCampo txt_strAvisoF
End Sub

Private Sub txt_strAvisoF_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txt_strAvisoF
End Sub

Private Function TruncaValores(strValor As String, bytCasasDecimais As Byte) As Double
Dim bytPos   As Byte

    bytPos = (Len(strValor) - InStr(strValor, ",")) - bytCasasDecimais
    
    TruncaValores = Mid(strValor, 1, Len(strValor) - bytPos)
    
End Function

Private Function strQueryAgencias(intBanco As Long)
    
    strQueryAgencias = " SELECT CB.Pkid, LTRIM(RTRIM(AG.strAgencia)) " & strCONCAT & "'/'" & strCONCAT & " LTRIM(RTRIM(CB.strConta)) strAgencia " & _
                       " FROM " & gstrContaBancaria & " CB, " & _
                        gstrAgencia & " AG " & _
                       " WHERE AG.intBanco = " & intBanco & " AND " & _
                       " CB.intAgencia = AG.Pkid "
                       
End Function

Private Sub ObterFaixaDeCeps()
Dim adoResultado As New ADODB.Recordset

    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO("SELECT " & gstrISNULL("intCepInicial", "0") & " intCepInicial, " & gstrISNULL("intCepFinal", "0") & " intCepFinal FROM " & gstrEmpresa, 10, adoResultado) Then
        intCepInicial = adoResultado("intCepInicial").Value
        intCepFinal = adoResultado("intCepInicial").Value
    End If
    adoResultado.Close: Set adoResultado = Nothing
    
    If intCepInicial = 0 Or intCepFinal = 0 Then
        ExibeMensagem "A faixa de Ceps do município não é válida."
    End If
    
End Sub
