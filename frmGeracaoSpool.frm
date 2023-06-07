VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmGeracaoSpool 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Geração de Spool"
   ClientHeight    =   4260
   ClientLeft      =   2130
   ClientTop       =   2550
   ClientWidth     =   6090
   Icon            =   "frmGeracaoSpool.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4260
   ScaleWidth      =   6090
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab tab_3dPasta 
      Height          =   4185
      Left            =   60
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   60
      Width           =   5985
      _ExtentX        =   10557
      _ExtentY        =   7382
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Taxa de Licença"
      TabPicture(0)   =   "frmGeracaoSpool.frx":1042
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lbl_progress2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lbl_progress1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "pgr_Status"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "fra_Parametros"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      Begin VB.Frame fra_Parametros 
         Caption         =   "Faixa de nº de Aviso"
         Height          =   2940
         Left            =   120
         TabIndex        =   1
         Top             =   420
         Width           =   5745
         Begin VB.TextBox txt_strAvisoI 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1065
            MaxLength       =   6
            TabIndex        =   5
            Top             =   810
            Width           =   1200
         End
         Begin VB.TextBox txt_strAvisoF 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1065
            MaxLength       =   6
            TabIndex        =   7
            Top             =   1230
            Width           =   1200
         End
         Begin VB.TextBox txt_intExercicio 
            Height          =   285
            Left            =   1065
            MaxLength       =   4
            TabIndex        =   9
            Top             =   1590
            Width           =   540
         End
         Begin VB.OptionButton opt_Ordenacao 
            Caption         =   "Ordenação por Identificação"
            Height          =   195
            Index           =   0
            Left            =   1065
            TabIndex        =   12
            Top             =   2355
            Width           =   3075
         End
         Begin VB.OptionButton opt_Ordenacao 
            Caption         =   "Ordenação por CEP"
            Height          =   195
            Index           =   1
            Left            =   1065
            TabIndex        =   13
            Top             =   2625
            Width           =   3075
         End
         Begin VB.TextBox txt_dtmDtBaixa 
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
            Left            =   1065
            MaxLength       =   10
            TabIndex        =   11
            Top             =   1965
            Width           =   1125
         End
         Begin MSDataListLib.DataCombo dbc_intComposicao 
            Height          =   315
            Left            =   1065
            TabIndex        =   3
            Top             =   390
            Width           =   4575
            _ExtentX        =   8070
            _ExtentY        =   556
            _Version        =   393216
            IntegralHeight  =   0   'False
            Text            =   ""
         End
         Begin MSComDlg.CommonDialog CmdBanco 
            Left            =   5160
            Top             =   2340
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Composição:"
            Height          =   195
            Left            =   90
            TabIndex        =   2
            Top             =   435
            Width           =   915
         End
         Begin VB.Label lblFinal 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Aviso Final:"
            Height          =   195
            Left            =   195
            TabIndex        =   6
            Top             =   1305
            Width           =   810
         End
         Begin VB.Label lblInicial 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Aviso Inicial:"
            Height          =   195
            Left            =   120
            TabIndex        =   4
            Top             =   870
            Width           =   885
         End
         Begin VB.Label lbl_exercicio 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Exercício:"
            Height          =   195
            Left            =   285
            TabIndex        =   8
            Top             =   1665
            Width           =   720
         End
         Begin VB.Label lblData 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Data base:"
            Height          =   195
            Left            =   225
            TabIndex        =   10
            Top             =   2040
            Width           =   780
         End
      End
      Begin MSComctlLib.ProgressBar pgr_Status 
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   3465
         Visible         =   0   'False
         Width           =   5745
         _ExtentX        =   10134
         _ExtentY        =   344
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin VB.Label lbl_progress1 
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   3720
         Width           =   735
      End
      Begin VB.Label lbl_progress2 
         Alignment       =   1  'Right Justify
         Height          =   255
         Left            =   4770
         TabIndex        =   16
         Top             =   3720
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmGeracaoSpool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim strWord           As String
    Dim strWordAux        As String

Private Function strQuery() As String
    
    Dim strsql  As String
    
    strsql = ""
    
    strsql = strsql & " SELECT PKId FROM "
   
    strQuery = strsql
    
End Function

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
    gintCodSeguranca = 1310
    
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo
    
    HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrSalvar
    
End Sub

Private Sub Form_Deactivate()
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar
End Sub

Private Sub Form_Load()
    CmdBanco.Filter = "Arquivo Texto | *.txt"
    dbc_intComposicao.Tag = strQueryComposicao & ";Strdescricao"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar
End Sub

Public Sub MantemForm(ByVal strModoOperacao As String)
        
    If UCase(strModoOperacao) = gstrSalvar Then
        If Not blnDadosOk Then Exit Sub
        GeraArquivo
    ElseIf UCase(strModoOperacao) = UCase(gstrPreencherLista) Then
        PreencherListaDeOpcoes Me.ActiveControl
    ElseIf UCase(strModoOperacao) = UCase(gstrNovo) Then
        Limpa_Controles Me, True, True, True, True, True
        dbc_intComposicao.SetFocus
    End If
                 
End Sub

Private Function strQueryComposicao() As String
    Dim strsql As String
    
    strsql = "SELECT CO.Pkid,"
    strsql = strsql & gstrCONVERT(CDT_VARCHAR, "CO.intCodigo") & strCONCAT & "' - '" & strCONCAT & _
                      " RTRIM(LTRIM(CO.strDescricao)) Descricao "
    strsql = strsql & "FROM "
    strsql = strsql & gstrComposicaoDaReceita & " CO "
    strsql = strsql & "WHERE "
    strsql = strsql & "CO.Intutilizacao = " & TYP_ECONOMICA & " "
    strsql = strsql & "ORDER BY strDescricao "
    
    strQueryComposicao = strsql

End Function

Private Function blnDadosOk()
    
    blnDadosOk = False
    If Not dbc_intComposicao.MatchedWithList Then
        ExibeMensagem "A composição da receita foi preenchida incorretamente."
        dbc_intComposicao.SetFocus
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
    ElseIf Trim(txt_intExercicio) = "" Or Len(txt_intExercicio) > 4 Then
        ExibeMensagem "Número do exercício foi preenchido incorretamente."
        txt_intExercicio.SetFocus
        Exit Function
    ElseIf Not gblnDataValida(txt_dtmDtBaixa) Then
        ExibeMensagem "A Data informada não é válida."
        txt_dtmDtBaixa.SetFocus
        Exit Function
    ElseIf opt_Ordenacao(0).Value = False And opt_Ordenacao(1).Value = False Then
        ExibeMensagem "É necessário preenchido da ordenação."
        opt_Ordenacao(0).SetFocus
        Exit Function
    End If
    
    blnDadosOk = True
    
End Function

Private Sub GeraArquivo()
    Dim adoResultado    As ADODB.Recordset
    Dim adoAux          As ADODB.Recordset
    Dim strsql          As String
    Dim intCont         As Integer
    Dim intFebraban     As Long
    Dim adoCodBarras    As ADODB.Recordset
    Dim INTNUMERO       As Long
    Dim bytDigito       As Integer
    Dim strCodBarras    As String
    Dim strNumeroBoleto1 As String
    Dim lngGuias        As Long
    Dim strNomeFantasia As String
    
    pgr_Status.Value = 0
    strWord = ""

On Error GoTo Gravar
    
    CmdBanco.ShowSave
    Screen.MousePointer = vbArrow
    
    If Trim(CmdBanco.filename) = "" Then Exit Sub
    
    'Query utilizada para pegar o Nome Fantasia e o Codigo Febraban da tblEmpresa
    strsql = ""
    strsql = strsql & "Select strNomeFantasia, intFebraban From " & gstrEmpresa
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strsql, 5, adoAux) Then
        If adoAux.RecordCount > 0 Then
        
            strNomeFantasia = gstrENulo(adoAux!strNomeFantasia)
            
            If gstrENulo(adoAux!intFebraban) <> "" Then
                intFebraban = gstrENulo(adoAux!intFebraban)
            Else
                ExibeMensagem "Código Febraban não encontrado."
                GoTo Gravar
            End If
        Else
            ExibeMensagem "Código Febraban não encontrado."
            GoTo Gravar
        End If
    End If
    
    'Vamos trazer os lançamentos de ISS
    strsql = "SELECT "
    strsql = strsql & "LA.PKID PkidLA, "
    strsql = strsql & "LE.PKID PkidLE, "
    strsql = strsql & gstrRIGHT("LA.strInscricao", gintRetornaTamanhoMascara(TYP_ECONOMICA)) & " strInscricao, "
    strsql = strsql & gstrISNULL("Ltrim(Rtrim(LA.strComposicaoDaReceita))", "''") & " strComposicao, "
    strsql = strsql & gstrISNULL("LA.intComposicaoDaReceita", "''") & " intComposicao, "
    strsql = strsql & gstrISNULL("LA.intExercicio", "0") & " intExercicio, "
    strsql = strsql & gstrCONVERT(cdt_numeric, gstrISNULL("LA.strNumeroAviso", "''")) & " strAviso, '"
    strsql = strsql & strNomeFantasia & "' Prefeitura, "
    strsql = strsql & "'Divisão de Rendas Mobiliárias' Secretaria, "
    strsql = strsql & "'Departamento' Departamento,"
    strsql = strsql & "'Divisão' Divisao, "
    strsql = strsql & gstrISNULL("Ltrim(Rtrim(LA.strNomeProprietario))", "' '") & " strContribuinte, "
    strsql = strsql & gstrISNULL("Ltrim(Rtrim(LA.strLogradouroC))", "' '") & " strLogradouroC, "
    strsql = strsql & gstrISNULL("Ltrim(Rtrim(LA.strNumeroC))", "' '") & " strNumeroC, "
    strsql = strsql & gstrISNULL("Ltrim(Rtrim(LA.strComplementoC))", "' '") & " strComplementoC, "
    strsql = strsql & gstrISNULL("Ltrim(Rtrim(LA.strBairroC))", "' '") & " strBairroC, "
    strsql = strsql & gstrISNULL("Ltrim(Rtrim(LA.strMunicipioC))", "' '") & " strMunicipioC, "
    strsql = strsql & gstrISNULL("Ltrim(Rtrim(LA.strUFC))", "' '") & " strUFC, "
    strsql = strsql & gstrISNULL("Ltrim(Rtrim(LA.intCEPC))", "'1'") & " intCEPC, "
    strsql = strsql & gstrISNULL("Ltrim(Rtrim(LA.strLogradouro))", "' '") & " strLogradouro, "
    strsql = strsql & gstrISNULL("Ltrim(Rtrim(LA.strNumero))", "' '") & " strNumero, "
    strsql = strsql & gstrISNULL("Ltrim(Rtrim(LA.strComplemento))", "' '") & " strComplemento, "
    strsql = strsql & gstrISNULL("Ltrim(Rtrim(LA.strBairro))", "' '") & " strBairro, "
    strsql = strsql & gstrISNULL("Ltrim(Rtrim(LA.strMunicipio))", "' '") & " strMunicipio, "
    strsql = strsql & gstrISNULL("Ltrim(Rtrim(LA.strUF))", "' '") & " strUF, "
    strsql = strsql & gstrISNULL("Ltrim(Rtrim(LA.intCEP))", "'1'") & " intCEP, "
    strsql = strsql & "'Sigla' strSigla, "
    strsql = strsql & gstrISNULL("Ltrim(Rtrim(LA.Strindexador))", "''") & " Strindexador, "
    strsql = strsql & gstrISNULL("LA.dblvlIndexador", "'1'") & " dblIndexador "
    strsql = strsql & "FROM "
    strsql = strsql & gstrLancamentoAlfa & " LA, "
    strsql = strsql & gstrLancamentoEconomico & " LE "
    strsql = strsql & "WHERE "
    strsql = strsql & "LA.Pkid = LE.Intlancamentoalfa AND "
    strsql = strsql & "LA.dtmdtCancelamento IS NULL AND "
    strsql = strsql & "LA.intComposicaoDaReceita = " & Val(dbc_intComposicao.BoundText) & " AND "
    strsql = strsql & "LA.Strnumeroaviso BETWEEN '" & String(gintLenNumAviso - Len(Trim(txt_strAvisoI.Text)), "0") & txt_strAvisoI.Text & "' AND '" & String(gintLenNumAviso - Len(Trim(txt_strAvisoF.Text)), "0") & txt_strAvisoF.Text & "' AND "
    strsql = strsql & "LA.intExercicio = " & txt_intExercicio
    If opt_Ordenacao(0).Value = True Then
        strsql = strsql & " Order By La.strinscricao "
    Else
        strsql = strsql & " Order By La.intcepc, La.strinscricao"
    End If
    
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strsql, 10, adoResultado) Then
        With adoResultado
            If Not .EOF Then
                Open CmdBanco.filename For Output As #1
                pgr_Status.Visible = True
                pgr_Status.Max = Abs(.RecordCount)
                lbl_progress2.Caption = adoResultado.RecordCount
                Do While Not .EOF
                
                    strWord = ""
                    'Inscrição
                    strWord = strWord & Format$(!strInscricao, "000,000")
                    'Composição da Receita
                    strWord = strWord & Left(!strComposicao, 40) & Space$(40 - Len(Left(!strComposicao, 40)))
                    'Exercício
                    strWord = strWord & !intExercicio & Space$(4 - Len(!intExercicio))
                    'Nº do Aviso
                    strWord = strWord & Format$(!strAviso, "000,000")
                    'Prefeitura
                    If Len(!Prefeitura) > 31 Then
                        strWord = strWord & Left(!Prefeitura, 31)
                    Else
                        strWord = strWord & Left(!Prefeitura, 31) & Space$(31 - Len(!Prefeitura))
                    End If
                    'Secretaria
                    strWord = strWord & Left(!Secretaria, 32) & Space$(32 - Len(!Secretaria))
                    'Departamento
                    strWord = strWord & Left(!Departamento, 23) & Space$(23 - Len(!Departamento))
                    'Divisao
                    strWord = strWord & Left(!Divisao, 23) & Space$(23 - Len(!Divisao))
                    'Nome do Contribuinte
                    strWord = strWord & Left(!strContribuinte, 100) & Space$(100 - Len(Left(!strContribuinte, 100)))
                    'Nome do Logradouro de Correspondencia
                    strWord = strWord & Left(!strlogradouroc, 100) & Space$(100 - Len(Left(!strlogradouroc, 100)))
                    'Número de Correspondencia
                    strWord = strWord & Space$(23 - Len(!strNumeroC)) & !strNumeroC
                    'Nome do Complemento de Correspondencia
                    strWord = strWord & Left(!strComplementoC, 30) & Space$(30 - Len(Left(!strComplementoC, 30)))
                    'Nome do Bairro de Correspondencia
                    strWord = strWord & Left(!strBairroC, 50) & Space$(50 - Len(Left(!strBairroC, 50)))
                    'Nome do Municipio de Correspondencia
                    strWord = strWord & Left(!strMunicipioC, 50) & Space$(50 - Len(Left(!strMunicipioC, 50)))
                    'Nome da UF de Correspondencia
                    strWord = strWord & Left(!strUFC, 2) & Space$(50 - Len(Left(!strUFC, 2)))
                    'CEP de Correspondencia
                    strWord = strWord & gstrCEPFormatado(Format$(IIf(Trim(!INTCEP) = "", "0", Trim(!INTCEP)), "00000000"))
                    'Nome do Logradouro
                    strWord = strWord & Left(!strLogradouro, 100) & Space$(100 - Len(Left(!strLogradouro, 100)))
                    'Número
                    strWord = strWord & Space$(23 - Len(!strNumero)) & !strNumero
                    'Nome do Complemento
                    strWord = strWord & Left(!STRCOMPLEMENTO, 30) & Space$(30 - Len(Left(!STRCOMPLEMENTO, 30)))
                    'Nome do Bairro
                    strWord = strWord & Left(!STRBAIRRO, 50) & Space$(50 - Len(Left(!STRBAIRRO, 50)))
                    'Nome do Municipio
                    strWord = strWord & Left(!STRMUNICIPIO, 50) & Space$(50 - Len(Left(!STRMUNICIPIO, 50)))
                    'Nome da UF
                    strWord = strWord & Left(!STRUF, 2) & Space$(50 - Len(Left(!STRUF, 2)))
                    'CEP
                    strWord = strWord & gstrCEPFormatado(Format$(IIf(Trim(!INTCEP) = "", "0", Trim(!INTCEP)), "00000000"))
                    'Sigla
                    strWord = strWord & Left(!strsigla, 15) & Space$(15 - Len(!strsigla))
                                            
                    'Vamos achar as atividades
                    
                    strsql = "Select " & gstrCASEWHEN("LEA.blnPrincipal", "1,'P',0,'S'") & strCONCAT & "' '" & strCONCAT & " Rtrim(Ltrim(" & gstrISNULL("LEA.Strdescricaoatividade", "''") & ")) Strdescricao "
                    strsql = strsql & "From " & gstrLctEconomicoAtividade & " LEA "
                    strsql = strsql & "Where LEA.Intlancamentoeconomico = " & !PkidLE
                    
                    strWordAux = ""
                    intCont = 0
                    
                    Set gobjBanco = New clsBanco
                    If gobjBanco.CriaADO(strsql, 10, adoAux) Then
                        If Not adoAux.EOF Then
                            Do While Not adoAux.EOF
                                If Not intCont > 6 Then
                                    strWordAux = strWordAux & Left(adoAux!strDescricao, 50) & Space$(50 - Len(Left(adoAux!strDescricao, 50)))
                                    intCont = intCont + 1
                                End If
                                adoAux.MoveNext
                            Loop
                        Else
                            ExibeMensagem "Não foram encontrados atividades para a inscrição: " & !strInscricao & " Aviso: " & !strAviso
                            GoTo Gravar
                        End If
                    Else
                        GoTo Gravar
                    End If
                    
                    'Atividades
                    If Len(Trim(strWordAux)) > 0 Then
                        strWord = strWord & Left(strWordAux, 300) & Space$(300 - Len(Left(strWordAux, 300)))
                    Else
                        strWord = strWord & Space(300)
                    End If
                    
                    'Vamos achar os tributos devidos
                    strsql = "Select "
                    strsql = strsql & gstrISNULL("R.Strsigla", "''") & " strReceita, "
                    strsql = strsql & "Sum(" & gstrISNULL("LR.DblValor", "0") & ") as dblValorReceita "
                    strsql = strsql & "From "
                    strsql = strsql & gstrLancamentoEconomico & " LE, "
                    strsql = strsql & gstrLancamentoAlfa & " LA, "
                    strsql = strsql & gstrLancamentoValor & " LV, "
                    strsql = strsql & gstrLancamentoReceita & " LR, "
                    strsql = strsql & gstrReceita & " R "
                    strsql = strsql & "Where "
                    strsql = strsql & "LA.Pkid = LE.Intlancamentoalfa And "
                    strsql = strsql & "LA.Pkid = LV.Intlancamentoalfa And "
                    strsql = strsql & "LV.Pkid = LR.Intlancamentovalor And "
                    strsql = strsql & "R.Pkid = LR.Intreceita And "
                    strsql = strsql & "LV.bitParcelaValida = 1 And "
                    strsql = strsql & "LE.Pkid = " & !PkidLE & " "
                    strsql = strsql & "Group By "
                    strsql = strsql & "R.pkid, "
                    strsql = strsql & "R.Strsigla, "
                    strsql = strsql & "LR.dblvalor "
                    
                    strWordAux = ""
                    intCont = 0
                    
                    Set gobjBanco = New clsBanco
                    If gobjBanco.CriaADO(strsql, 10, adoAux) Then
                        If Not adoAux.EOF Then
                            Do While Not adoAux.EOF
                                If Not intCont > 5 Then
                                    strWordAux = strWordAux & Left(adoAux!strReceita, 10) & Space$(10 - Len(Left(adoAux!strReceita, 10)))
                                    strWordAux = strWordAux & Space$(16 - Len(Left(gstrConvVrDoSql(adoAux!dblValorReceita, , , True), 16))) & Left(gstrConvVrDoSql(adoAux!dblValorReceita, , , True), 16)
                                    'Debug.Print strWordAux
                                    intCont = intCont + 1
                                End If
                                adoAux.MoveNext
                            Loop
                        Else
                            ExibeMensagem "Não foram encontrados tributos para a inscrição: " & !strInscricao & " Aviso: " & !strAviso
                            GoTo Gravar
                        End If
                    Else
                        GoTo Gravar
                    End If
                    
                    'Receitas
                    If Len(Trim(strWordAux)) > 0 Then
                        strWord = strWord & Left(strWordAux, 130) & Space$(130 - Len(Left(strWordAux, 130)))
                    Else
                        strWord = strWord & Space(130)
                    End If
                    
                    'Vamos achar os valores de Lançamentos
                    strsql = "Select " & IIf(bytDBType = SQLServer, "Top 1 ", "")
                    strsql = strsql & "LV.Dtmdtvencimento, "
                    strsql = strsql & gstrISNULL("LV.DblValor", 0) & " DblParcela, "
                    strsql = strsql & gstrISNULL("LV.DblValor", 0) & " / " & gstrConvVrParaSql(!dblIndexador) & " DblParcelaFMP, "
                    strsql = strsql & "DblTotal, "
                    strsql = strsql & "DblTotalFMP, "
                    strsql = strsql & "DblTotParcela, "
                    strsql = strsql & "(Select " & gstrISNULL("M.STRABREVIATURA", "''") & " From " & gstrMoedas & " M where pkid = LV.Intmoeda) strMoeda,"
                    strsql = strsql & "(Select " & gstrISNULL("M.dblValorCorte", 0) & " From " & gstrMoedas & " M where pkid = LV.Intmoeda) dblMoeda "
                    strsql = strsql & "From "
                    strsql = strsql & gstrLancamentoValor & " LV, "
                    strsql = strsql & "(SELECT Sum(" & gstrISNULL("LV.Dblvalor", 0) & ") dblTotal, "
                    strsql = strsql & "Sum(" & gstrISNULL("LV.Dblvalor", 0) & ")" & " / " & gstrConvVrParaSql(!dblIndexador) & " as DblTotalFMP, "
                    strsql = strsql & "Count(*) dblTotParcela "
                    strsql = strsql & "FROM " & gstrLancamentoValor & " LV "
                    strsql = strsql & "WHERE LV.bitparcelavalida  = 1 AND "
                    strsql = strsql & "LV.Intlancamentoalfa = " & !PkidLA & ") LV1 "
                    strsql = strsql & "Where "
                    strsql = strsql & "LV.Intlancamentoalfa = " & !PkidLA & " AND "
                    strsql = strsql & "LV.bitparcelavalida  = 1 " & IIf(bytDBType = Oracle, " AND RowNum < 2 ", "")
                    
                    Set gobjBanco = New clsBanco
                    If gobjBanco.CriaADO(strsql, 10, adoAux) Then
                        If Not adoAux.EOF Then
                            'Total do Lançamento
                            strWord = strWord & Space$(16 - Len(Left(gstrConvVrDoSql(adoAux!dblTotal, , , True), 16))) & Left(gstrConvVrDoSql(adoAux!dblTotal, , , True), 16)
                            'Total do Lançamento em FMP
                            strWord = strWord & Space$(11 - Len(Left(gstrConvVrDoSql(adoAux!DblTotalFMP, 4, , True), 11))) & Left(gstrConvVrDoSql(adoAux!DblTotalFMP, 4, , True), 11)
                            'Valor da Parcela
                            strWord = strWord & Space$(16 - Len(Left(gstrConvVrDoSql(adoAux!DblParcela, , , True), 16))) & Left(gstrConvVrDoSql(adoAux!DblParcela, , , True), 16)
                            'Valor da Parcela em FMP
                            strWord = strWord & Space$(11 - Len(Left(gstrConvVrDoSql(adoAux!DblParcelaFMP, 4, , True), 11))) & Left(gstrConvVrDoSql(adoAux!DblParcelaFMP, 4, , True), 11)
                            'Nº de parcelas
                            strWord = strWord & Format$(adoAux!DblTotParcela, "000")
                            'Data de Vencimento
                            strWord = strWord & Format$(adoAux!Dtmdtvencimento, "dd/mm/yyyy")
                            'Descrição do Indexador
                            strWord = strWord & Left(!Strindexador, 10) & Space$(10 - Len(!Strindexador))
                            'Valor do Indexador
                            'strWord = strWord & Space$(6 - Len(Left(gstrConvVrDoSql(adoAux!dblMoeda, 4, , True), 6))) & Left(gstrConvVrDoSql(adoAux!dblMoeda, 4, , True), 6)
                            strWord = strWord & Space$(6 - Len(Left(gstrConvVrDoSql(!dblIndexador, 4, , True), 6))) & Left(gstrConvVrDoSql(!dblIndexador, 4, , True), 6)
                        Else
                            ExibeMensagem "Não foram encontrados valores de lançamentos para a inscrição: " & !strInscricao & " Aviso: " & !strAviso
                            GoTo Gravar
                        End If
                    Else
                        GoTo Gravar
                    End If
'-----------------------------------------------------------------------------------------------------------------------
                    strsql = ""
                    strsql = strsql & "Select "
                    strsql = strsql & "LV.Pkid, LV.Intlancamentoalfa, "
                    strsql = strsql & "LV.INTPARCELA, "
                    strsql = strsql & "LV.Dtmdtvencimento, "
                    strsql = strsql & gstrISNULL("LV.Dblvalor", 0) & " Dblvalor, "
                    strsql = strsql & "LV.Bitparcelavalida "
                    strsql = strsql & "From "
                    strsql = strsql & gstrLancamentoValor & " LV "
                    strsql = strsql & "Where "
                    strsql = strsql & "LV.Intlancamentoalfa = " & !PkidLA
                    strsql = strsql & "Order By "
                    strsql = strsql & "LV.Intlancamentoalfa, "
                    strsql = strsql & "LV.Bitparcelavalida, "
                    strsql = strsql & "LV.Intparcela "
                    
                    Set gobjBanco = New clsBanco
                    If gobjBanco.CriaADO(strsql, 10, adoAux) Then
                        If Not .EOF Then
                            
                            Set gobjBanco = New clsBanco
                            gobjBanco.ExecutaBeginTrans
                            
                            Do While Not adoAux.EOF
                            
                                'Query Utilizada para pegar o último número da guia na seqNumeroGuia
                                INTNUMERO = glngRetornaProximoNumeroGuia
                            
                                'Nº da parcela
                                strWord = strWord & Format$(gstrENulo(adoAux!intParcela), "000")
                                'Nº da Guia
                                strWord = strWord & Format$(INTNUMERO, "0000000000")
                                'Data de Vencimento
                                strWord = strWord & Format$(gstrENulo(adoAux!Dtmdtvencimento), "dd/mm/yyyy")
                                'Valor das parcelas
                                strWord = strWord & Space$(16 - Len(Left(gstrConvVrDoSql(gstrENulo(adoAux!dblValor), 2, , True), 16))) & Left(gstrConvVrDoSql(gstrENulo(adoAux!dblValor), 2, , True), 16)
                                
    'Vamos criar o Código de Barras
                        
                                strCodBarras = ""
                                strCodBarras = IIf(Abs(CInt(CBool(gstrENulo(adoAux!bitParcelaValida)))) = 0, "816", "817") 'Digito fixo
                                strCodBarras = strCodBarras & Format$((adoAux!dblValor * 100), "00000000000")   'Valor da guia
                                strCodBarras = strCodBarras & Format$(intFebraban, "0000")               'Codigo do Febraban
                                strCodBarras = strCodBarras & Replace(Format$(adoAux!Dtmdtvencimento, "YYYY/MM/DD"), "/", "") 'Vencimento da Guia
                                strCodBarras = strCodBarras & "0000"                                     'Conta bancaria tipo nulo que é a do Febrabraban
                                strCodBarras = strCodBarras & Format$(INTNUMERO, "000000000")            'Número sequencial da guia
                                strCodBarras = strCodBarras & Year(gstrDataDoSistema)                      'Exercício corrente
                                
                                bytDigito = gstrCalculaDigitoModulo10(strCodBarras) 'Calcula o digito
                                strCodBarras = Mid(strCodBarras, 1, 3) & bytDigito & Mid(strCodBarras, 4, Len(strCodBarras)) 'Adiciona o digito ao codigo de barras
                                
                                strNumeroBoleto1 = Mid(strCodBarras, 1, 11) & "-" & gstrCalculaDigitoModulo10(Mid(strCodBarras, 1, 11)) & " "
                                strNumeroBoleto1 = strNumeroBoleto1 & Mid(strCodBarras, 12, 11) & "-" & gstrCalculaDigitoModulo10(Mid(strCodBarras, 12, 11)) & " "
                                strNumeroBoleto1 = strNumeroBoleto1 & Mid(strCodBarras, 23, 11) & "-" & gstrCalculaDigitoModulo10(Mid(strCodBarras, 23, 11)) & " "
                                strNumeroBoleto1 = strNumeroBoleto1 & Mid(strCodBarras, 34, 11) & "-" & gstrCalculaDigitoModulo10(Mid(strCodBarras, 34, 11))
                                
                                strWord = strWord & strCodBarras
                                strWord = strWord & strNumeroBoleto1
                                
                                'Vamos inserir a guia na tabela TblGuias
                                strsql = ""
                                strsql = strsql & "Insert Into " & gstrGuias & "("
                                strsql = strsql & "Intcontabancaria, "
                                strsql = strsql & "Intnumero, "
                                strsql = strsql & "Dtmdtemissao, "
                                strsql = strsql & "Dblvalor, "
                                strsql = strsql & "Strcodbarra, "
                                strsql = strsql & "Dtmdtatualizacao, "
                                strsql = strsql & "Lngcodusr, "
                                strsql = strsql & "Dtmdtvencimento "
                                strsql = strsql & ") Values("
                                
                                strsql = strsql & "Null, "
                                strsql = strsql & INTNUMERO & ", "
                                strsql = strsql & gstrConvDtParaSql(gstrDataDoSistema) & ", "
                                strsql = strsql & gstrConvVrParaSql(adoAux!dblValor) & ", '"
                                strsql = strsql & strCodBarras & "', "
                                strsql = strsql & gstrConvDtParaSql(gstrDataDoSistema) & ", "
                                strsql = strsql & glngCodUsr & ", "
                                strsql = strsql & gstrConvDtParaSql(adoAux!Dtmdtvencimento)
                                strsql = strsql & ")"
                                'strSql = strSql & IIf((bytDBType = EDatabases.Oracle), " ; ", "")
                                
                                If Not gobjBanco.Execute(strsql) Then
                                    ExibeMensagem "Erro na gravação da guia."
                                    gobjBanco.ExecutaRollbackTrans
                                    GoTo Gravar
                                End If

                                'Vamos inserir as parcelas na tabela TblLancamentoGuias
                                
                                lngGuias = glngRetornaPkidTabelaPai("seqtblGuias", "tblGuias")
                                
                                strsql = "Insert Into " & gstrLancamentoGuias & "("
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
                                strsql = strsql & adoAux!Pkid & ", "
                                strsql = strsql & lngGuias & ","
                                strsql = strsql & gstrConvVrParaSql(adoAux!dblValor) & ", "
                                strsql = strsql & "0.00" & ", "
                                strsql = strsql & "0.00" & ", "
                                strsql = strsql & "0.00" & ", "
                                strsql = strsql & "0.00" & ", "
                                strsql = strsql & strGETDATE & ", "
                                strsql = strsql & glngCodUsr & ") "
                                
                                If Not gobjBanco.Execute(strsql) Then
                                    ExibeMensagem "Erro na gravação da guia."
                                    gobjBanco.ExecutaRollbackTrans
                                    GoTo Gravar
                                End If
                                
                                adoAux.MoveNext
                                
                            Loop
                        Else
                            ExibeMensagem "Não foram encontrados registros com esses parâmetros."
                            GoTo Gravar
                        End If
                    End If


                    Print #1, strWord
                    DoEvents
                    pgr_Status.Value = .AbsolutePosition
                    lbl_progress1.Caption = .AbsolutePosition
                    .MoveNext

                Loop
            Else
               ExibeMensagem "Não foram encontrados registros com esses parâmetros."
               GoTo Gravar
            End If
        End With
    Else
        GoTo Gravar
    End If
    
    gobjBanco.ExecutaCommitTrans
    Close #1
    Screen.MousePointer = vbDefault
    Exit Sub
    
Gravar:
    gobjBanco.ExecutaRollbackTrans
    Close #1
    'Open CmdBanco.filename For Output As #1
    'Close #1
    Screen.MousePointer = vbDefault

End Sub

Private Sub txt_dtmDtBaixa_GotFocus()
    MarcaCampo txt_dtmDtBaixa
End Sub

Private Sub txt_dtmDtBaixa_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "D", txt_dtmDtBaixa
End Sub

Private Sub txt_dtmDtBaixa_LostFocus()
    txt_dtmDtBaixa = gstrDataFormatada(txt_dtmDtBaixa)
End Sub

Private Sub txt_intExercicio_GotFocus()
    MarcaCampo txt_intExercicio
End Sub

Private Sub txt_intExercicio_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txt_intExercicio
End Sub

Private Sub txt_strAvisoF_GotFocus()
    MarcaCampo txt_strAvisoF
End Sub

Private Sub txt_strAvisoF_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txt_strAvisoF
End Sub

Private Sub txt_strAvisoI_GotFocus()
    MarcaCampo txt_strAvisoI
End Sub

Private Sub txt_strAvisoI_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txt_strAvisoI
End Sub
