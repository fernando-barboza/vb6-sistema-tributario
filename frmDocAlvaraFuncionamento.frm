VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmDocAlvaraFuncionamento 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Alvará de Funcionamento"
   ClientHeight    =   2130
   ClientLeft      =   3750
   ClientTop       =   3315
   ClientWidth     =   3840
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2130
   ScaleWidth      =   3840
   Begin TabDlg.SSTab tab_3DPasta 
      Height          =   2085
      Left            =   60
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   30
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   3678
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   529
      TabCaption(0)   =   "Inscrição Cadastral"
      TabPicture(0)   =   "frmDocAlvaraFuncionamento.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblInscricaoInicial"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "dbc_Inscricao"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txt_dtmVencimento"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "fra_OpcaoConsulta"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      Begin VB.Frame fra_OpcaoConsulta 
         Caption         =   "Opções Para Consulta"
         Height          =   795
         Left            =   30
         TabIndex        =   5
         Top             =   1200
         Width           =   3615
         Begin MSDataListLib.DataCombo dbc_Obs 
            Height          =   315
            Left            =   1140
            TabIndex        =   6
            Top             =   360
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   556
            _Version        =   393216
            IntegralHeight  =   0   'False
            Text            =   ""
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Observação"
            Height          =   195
            Left            =   90
            TabIndex        =   7
            Top             =   480
            Width           =   870
         End
      End
      Begin VB.TextBox txt_dtmVencimento 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#.##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   1500
         MaxLength       =   30
         TabIndex        =   4
         Top             =   810
         Width           =   1155
      End
      Begin MSDataListLib.DataCombo dbc_Inscricao 
         Height          =   315
         Left            =   1500
         TabIndex        =   1
         Top             =   420
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   556
         _Version        =   393216
         IntegralHeight  =   0   'False
         Text            =   ""
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Data Vencimento"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   870
         Width           =   1230
      End
      Begin VB.Label lblInscricaoInicial 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Inscrição:"
         Height          =   195
         Left            =   105
         TabIndex        =   2
         Top             =   510
         Width           =   690
      End
   End
End
Attribute VB_Name = "frmDocAlvaraFuncionamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSql                      As String
Dim lngPkid                     As Long
Dim XArrayAlinhaColunas         As XArrayDB
Dim XValores                    As XArrayDB
Dim XArrayAlinhaColunasHorario  As XArrayDB
Dim XArrayTabelaHorario         As XArrayDB


Private Sub dbc_Inscricao_GotFocus()
    MarcaCampo dbc_Inscricao
End Sub

Private Sub dbc_Inscricao_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbc_Inscricao
End Sub

Private Sub dbc_Obs_GotFocus()
    MarcaCampo dbc_Obs
End Sub

Private Sub dbc_Obs_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbc_Obs
End Sub

Private Sub Form_Activate()
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrDeletar, gstrAplicar, gstrSalvar
    HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrImprimir, gstrPreencherLista
End Sub

Private Sub Form_Load()
    dbc_Inscricao.Tag = strQueryInscricaoCadastral & ";strInscricaoCadastral"
    dbc_Obs.Tag = strQueryTextoLivre & ";strDescricao"
End Sub

Private Function strQueryTextoLivre() As String
    
    strSql = ""
    strSql = "SELECT Pkid, strDescricao"
    strSql = strSql & " FROM "
    strSql = strSql & gstrTextoLivre
    strSql = strSql & " ORDER BY strdescricao"
    strQueryTextoLivre = strSql
End Function
Private Sub txt_dtmVencimento_GotFocus()
    MarcaCampo txt_dtmVencimento
End Sub

Private Sub txt_dtmVencimento_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "D", txt_dtmVencimento
End Sub

Private Function strQueryInscricaoCadastral() As String
    
    strSql = ""
    strSql = strSql & "SELECT "
    strSql = strSql & "EC.Pkid, "
    strSql = strSql & gstrRIGHT("EC.Strinscricaocadastral", gintRetornaTamanhoMascara(TYP_ECONOMICA)) & " Strinscricaocadastral "
    strSql = strSql & " FROM "
    strSql = strSql & gstrEconomico & " EC "
    strSql = strSql & " WHERE EC.dtmdataencerramento IS NULL "
    
    If gintCodSeguranca = 1191 Then
        strSql = strSql & "AND Ec.bitdefinitivo = 1 "
    Else
        strSql = strSql & "AND " & gstrISNULL("Ec.bitdefinitivo", 0) & " = 0 "
    End If
    
    strSql = strSql & "ORDER BY "
    strSql = strSql & "strInscricaoCadastral"

strQueryInscricaoCadastral = strSql
End Function

Private Function strQuery() As String

        strSql = ""
        strSql = strSql & "SELECT"
        strSql = strSql & " EC.PKID Pkid,"
        strSql = strSql & gstrRIGHT("EC.Strinscricaocadastral", gintRetornaTamanhoMascara(TYP_ECONOMICA)) & " InsCad, "
        strSql = strSql & " EC.dtmdataabertura,"
        strSql = strSql & gstrCONVERT(CDT_VARCHAR, "EC.Intnumero") & strCONCAT & "'  '" & strCONCAT & gstrISNULL("EC.strcomplemento", "''") & " Num,"
        strSql = strSql & " EC.Intnumdeempregados Empregados,"
        strSql = strSql & " EC.Dblareaocupada AreaOcupada,"
        strSql = strSql & " CO.Strnome RazaoSocial,"
        strSql = strSql & " EC.Dtmrazaoinicio,"
        strSql = strSql & " CO.Strcnpjcpf CNPJCPF,"
        strSql = strSql & " CO.Strinscricaoestadual IE,"
        strSql = strSql & " TL.Strsigla Sigla,"
        
        strSql = strSql & gstrISNULL("tl.strsigla", "''") & strCONCAT & "' '" & strCONCAT & gstrISNULL("tlg.strsigla", "''") & strCONCAT & "' '" & strCONCAT & gstrISNULL("LG.Strdescricao", "''") & " Logradouro, "
        
        'strSQL = strSQL & " LG.Strdescricao Logradouro,"
        strSql = strSql & " EC.Dtmenderecoinicio,"
        strSql = strSql & " BA.strDescricao Bairro,"
        strSql = strSql & " EC.strcodprocesso " & strCONCAT & "'/'" & strCONCAT & gstrCONVERT(CDT_VARCHAR, "EC.intexerprocesso") & strCONCAT & "'-'" & strCONCAT & gstrCONVERT(CDT_VARCHAR, "ec.bitdigprocesso") & " Processo,"
        strSql = strSql & " OE.Strdescricao OcorrenciaProcesso,"
        strSql = strSql & " EC.bitDefinitivo,"
        strSql = strSql & " EC.Inthorariofuncionamento "
    strSql = strSql & " FROM "
        strSql = strSql & gstrEconomico & " EC, "
        strSql = strSql & gstrContribuinte & " CO, "
        strSql = strSql & gstrTituloLogradouro & " TLG, "
        strSql = strSql & gstrTipoLogradouro & " TL, "
        strSql = strSql & gstrLogradouro & " LG, "
        strSql = strSql & gstrBairro & " BA, "
        strSql = strSql & gstrOcorrenciaDoEconomico & " OE"
    strSql = strSql & " WHERE"
        strSql = strSql & " EC.strInscricaoCadastral = '" & String(gintLenInscricao - Len(dbc_Inscricao.Text), "0") & dbc_Inscricao.Text & "' And "
        strSql = strSql & " EC.Intcontribuinte " & strOUTJSQLServer & "= CO.pkid " & strOUTJOracle & " And "
        strSql = strSql & " EC.Intlogradouro = LG.Pkid and "
        strSql = strSql & " LG.INTTIPOLOGRADOURO " & strOUTJSQLServer & "= TL.Pkid" & strOUTJOracle & " And "
        strSql = strSql & " LG.Inttitulologradouro " & strOUTJSQLServer & "= TLG.Pkid" & strOUTJOracle & " And "
        strSql = strSql & " EC.Intbairro " & strOUTJSQLServer & "= BA.Pkid " & strOUTJOracle & " And"
        strSql = strSql & " EC.Intocorrenciadoeconomico " & strOUTJSQLServer & "= OE.Pkid" & strOUTJOracle & ""
    strSql = strSql & " ORDER BY"
          strSql = strSql & " EC.strInscricaoCadastral"
    strQuery = strSql
    
End Function

Private Sub ImprimirDocumento()
    Dim strNumero           As String
    Dim adoResultado        As ADODB.Recordset
    Dim strObs              As String
    
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strQuery, 5, adoResultado) Then
        If Not adoResultado.EOF Then
            With adoResultado
                lngPkid = gstrENulo(!Pkid)
                
                'CHAMA ARRAY RESPONSAVEL POR TRAZER AS ATIVIDADES ECONOMICAS
                PreencheCampos lngPkid
                
                PreencheCamposHorario lngPkid, Val(gstrENulo(!Inthorariofuncionamento))
                
                'Número do Alvará
                strNumero = glngRetornaProximoNumeroGuia(gstrEmpresa, "intNumeroAlvaraFuncionamento")
                
                'Query utlizada para pegar a observacao
                strObs = ""
                If dbc_Obs.MatchedWithList Then
                    strSql = ""
                    strSql = strSql & "select * from " & gstrTextoLivre & " WHERE PKID = " & dbc_Obs.BoundText & ""
                    Set gobjBanco = New clsBanco
                    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
                        If adoResultado.RecordCount >= 0 Then strObs = gstrENulo(adoResultado!strTexto)
                    End If
                End If
                
                AlinhaCampos
                AlinhaCamposHorario
                
                OpenWordDocumentAlvaraFuncionamento strNumero, gstrDataFormatada(gstrENulo(!dtmDataAbertura)), txt_dtmVencimento.Text, gstrENulo(!InsCad), gstrENulo(!RazaoSocial), _
                gstrENulo(!Sigla), gstrENulo(!Logradouro), gstrENulo(!Num), gstrENulo(!Bairro), IIf(IsNull((!Empregados)), 0, (!Empregados)), _
                gstrConvVrDoSql(gstrENulo(!AreaOcupada), , , True), "", gstrENulo(!Processo), gstrCGCCPFFormatado(gstrENulo(!CNPJCPF)), gstrENulo(!IE), _
                "'NumeroJucesp'", strObs, XValores, XArrayAlinhaColunas, Val(gstrENulo(!BitDefinitivo)), gstrDataFormatada(gstrENulo(!Dtmrazaoinicio)), gstrDataFormatada(gstrENulo(!Dtmenderecoinicio)), gstrENulo(!OcorrenciaProcesso), XArrayTabelaHorario, XArrayAlinhaColunasHorario
                
            End With
        Else
            ExibeMensagem "Nada foi encontrado com nesse intervalo de Inscrições"
        End If
    End If
    Set gobjBanco = Nothing
End Sub

Private Sub PreencheCampos(intPkid As Long)
    Dim intPosition     As Integer
    Dim varAux          As Variant
    Dim adoResultado    As ADODB.Recordset
    
    Set XValores = New XArrayDB
    
    XValores.Clear
    
    strSql = ""
    strSql = strSql & "SELECT "
    strSql = strSql & gstrCONVERT(CDT_VARCHAR, "AEC.Intcodigo") & strCONCAT & " '   ' " & strCONCAT & " Case When AEM.blnPrincipal = 1 then 'P' Else 'S' End " & strCONCAT & " '   ' " & strCONCAT & " AEC.strdescricao Atividade, "
    strSql = strSql & "AEM.dtmatividadeinicio "
    strSql = strSql & "FROM "
    strSql = strSql & gstrEconomico & " EC, "
    strSql = strSql & gstrAtividadeDaEmpresa & " AEM, "
    strSql = strSql & gstrAtividadeEC & " AEC "
    strSql = strSql & "WHERE "
    strSql = strSql & "EC.PKID = " & intPkid & " and "
    strSql = strSql & "AEM.INTECONOMICO = EC.Pkid and "
    strSql = strSql & "AEM.Intatividade = AEC.pkid And "
    strSql = strSql & "AEM.dtmatividadefim Is Null "
    strSql = strSql & "ORDER BY "
    strSql = strSql & "AEM.blnPrincipal Desc"
    intPosition = 0
    
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        If Not adoResultado.EOF Then
            With adoResultado
                Do While Not .EOF
                    XValores.ReDim 0, intPosition, 0, 1
                    
                    varAux = gstrENulo(!Atividade)
                    XValores(intPosition, 0) = varAux
                    
                    varAux = "em " & gstrDataFormatada(gstrENulo(!dtmAtividadeInicio))
                    XValores(intPosition, 1) = varAux

                    .MoveNext
                    intPosition = intPosition + 1
                Loop
            End With
        End If
    End If
    
End Sub


Private Function blnDadosOk()
    blnDadosOk = False
    
    If Not dbc_Inscricao.MatchedWithList Then
        ExibeMensagem "O número da inscrição deve ser selecionado."
        dbc_Inscricao.SetFocus
        Exit Function
    ElseIf Trim(txt_dtmVencimento) <> "" Then
        If gblnDataValida(txt_dtmVencimento, False, False) = False Then
            ExibeMensagem "Data inválida."
            txt_dtmVencimento.SetFocus
            Exit Function
        End If
    End If
    
    blnDadosOk = True
End Function

Public Sub MantemForm(ByVal strModoOperacao As String)
    If UCase(strModoOperacao) = gstrImprimir Then
        If blnDadosOk Then ImprimirDocumento
        ElseIf UCase(strModoOperacao) = gstrNovo Then
            dbc_Inscricao.Text = ""
            txt_dtmVencimento.Text = ""
            dbc_Obs.Text = ""
        ElseIf UCase(strModoOperacao) = gstrPreencherLista Then PreencherListaDeOpcoes Me.ActiveControl
    End If
End Sub

Private Sub AlinhaCampos()

    Set XArrayAlinhaColunas = New XArrayDB
    
    With XArrayAlinhaColunas 'Alinhamento
        .Clear
        .ReDim 0, 0, 0, 1
        .Value(0, 0) = WORDALIGNPARAGRAPHLEFT
        .Value(0, 1) = WORDALIGNPARAGRAPHRIGHT
    End With
    
End Sub

Private Sub txt_dtmVencimento_LostFocus()
    txt_dtmVencimento = gstrDataFormatada(txt_dtmVencimento)
End Sub

Private Sub PreencheCamposHorario(intPkid As Long, intHorario As Long)
    Dim intPosition     As Integer
    Dim varAux          As Variant
    Dim adoResultado    As ADODB.Recordset
    
    Set XArrayTabelaHorario = New XArrayDB
    
    XArrayTabelaHorario.Clear
    
    If intHorario > 0 Then
        strSql = "Select "
        strSql = strSql & "HF.Strdescricao str1, Null str2 "
        strSql = strSql & "from "
        strSql = strSql & gstrEconomico & " EC, "
        strSql = strSql & "tblhorariofuncionamento HF "
        strSql = strSql & "Where "
        strSql = strSql & "EC.Inthorariofuncionamento = HF.PKID AND "
        strSql = strSql & "EC.Pkid = " & intPkid
    Else
        strSql = "select strManhaDe str1, strManhaAte str2 from tbleconomico where pkid = " & intPkid & " and not strManhaDe is null and not strManhaAte is null Union "
        strSql = strSql & "select strTardeDe, strTardeAte from tbleconomico Where pkid = " & intPkid & " and not strTardeDe is null and not strTardeAte is null Union "
        strSql = strSql & "select strNoiteDe, strNoiteAte from tbleconomico Where pkid = " & intPkid & " and not strNoiteDe is null and not strNoiteAte is null Union "
        strSql = strSql & "select strMadrugadaDe, strMadrugadaAte from tbleconomico Where pkid = " & intPkid & " and not strMadrugadaDe is null and not strMadrugadaAte is null "
    End If
    
    intPosition = 0
    
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        If Not adoResultado.EOF Then
            With adoResultado
                Do While Not .EOF
                    XArrayTabelaHorario.ReDim 0, intPosition, 0, 1
                    
                    varAux = gstrENulo(!str1)
                    XArrayTabelaHorario(intPosition, 0) = varAux
                    
                    varAux = gstrENulo(!str2)
                    XArrayTabelaHorario(intPosition, 1) = varAux

                    .MoveNext
                    intPosition = intPosition + 1
                Loop
            End With
        End If
    End If
    
End Sub

Private Sub AlinhaCamposHorario()

    Set XArrayAlinhaColunasHorario = New XArrayDB
    
    With XArrayAlinhaColunasHorario 'Alinhamento
        .Clear
        .ReDim 0, 0, 0, 1
        .Value(0, 0) = WORDALIGNPARAGRAPHLEFT
        .Value(0, 1) = WORDALIGNPARAGRAPHRIGHT
    End With
    
End Sub

