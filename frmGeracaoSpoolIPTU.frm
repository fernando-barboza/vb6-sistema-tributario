VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmGeracaoSpoolIPTU 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Geração de Spool para Impressão IPTU"
   ClientHeight    =   7695
   ClientLeft      =   2490
   ClientTop       =   2085
   ClientWidth     =   7770
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7695
   ScaleWidth      =   7770
   Begin TabDlg.SSTab tab_Parametros 
      Height          =   6930
      Left            =   90
      TabIndex        =   0
      Top             =   60
      Width           =   7605
      _ExtentX        =   13414
      _ExtentY        =   12224
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Parâmetros"
      TabPicture(0)   =   "frmGeracaoSpoolIPTU.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "pgr_Status"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "CmdBanco"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "fra_Parametros"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "fra_GerarPeloEndereco"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "fra_GerarPorCEP"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      Begin VB.Frame fra_GerarPorCEP 
         Caption         =   "Gerar lançamentos com entrega"
         Height          =   1695
         Left            =   1320
         TabIndex        =   28
         Top             =   4800
         Width           =   4935
         Begin VB.CheckBox chk_GerarPeloCEP 
            Caption         =   "Dentro do município"
            Height          =   195
            Index           =   1
            Left            =   240
            TabIndex        =   13
            Top             =   600
            Width           =   4515
         End
         Begin VB.CheckBox chk_GerarPeloCEP 
            Caption         =   "Fora do município"
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   12
            Top             =   360
            Width           =   4515
         End
         Begin VB.TextBox txt_strCEP 
            Height          =   285
            Index           =   1
            Left            =   2520
            TabIndex        =   15
            Top             =   1200
            Width           =   1575
         End
         Begin VB.TextBox txt_strCEP 
            Height          =   285
            Index           =   0
            Left            =   480
            TabIndex        =   14
            Top             =   1200
            Width           =   1575
         End
         Begin VB.Label Label5 
            Caption         =   "até"
            Height          =   255
            Left            =   2160
            TabIndex        =   31
            Top             =   1230
            Width           =   255
         End
         Begin VB.Label Label4 
            Caption         =   "de"
            Height          =   255
            Left            =   240
            TabIndex        =   30
            Top             =   1230
            Width           =   255
         End
         Begin VB.Label Label3 
            Caption         =   "Faixa de CEP do município"
            Height          =   255
            Left            =   240
            TabIndex        =   29
            Top             =   960
            Width           =   4455
         End
      End
      Begin VB.Frame fra_GerarPeloEndereco 
         Caption         =   "Geração pela existência de endereço no cadastro imobiliário"
         Height          =   975
         Left            =   1320
         TabIndex        =   17
         Top             =   3720
         Width           =   4935
         Begin VB.OptionButton opt_GerarPeloEndereco 
            Caption         =   "Sem endereço de notificação no cadastro imobiliário"
            Height          =   195
            Index           =   1
            Left            =   240
            TabIndex        =   11
            Top             =   600
            Width           =   4515
         End
         Begin VB.OptionButton opt_GerarPeloEndereco 
            Caption         =   "Com endereço de notificação no cadastro imobiliário"
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   10
            Top             =   360
            Width           =   4515
         End
      End
      Begin VB.Frame fra_Parametros 
         Caption         =   "Faixa de nº de Aviso"
         Height          =   3270
         Left            =   1320
         TabIndex        =   16
         Top             =   360
         Width           =   4935
         Begin VB.ComboBox cboComposicaoReceita 
            Height          =   315
            Left            =   240
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   2760
            Width           =   4545
         End
         Begin VB.TextBox txt_strAvisoI 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1350
            MaxLength       =   6
            TabIndex        =   1
            Top             =   330
            Width           =   1200
         End
         Begin VB.TextBox txt_strAvisoF 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   3120
            MaxLength       =   6
            TabIndex        =   2
            Top             =   330
            Width           =   1200
         End
         Begin VB.TextBox txt_intExercicio 
            Height          =   285
            Left            =   1365
            MaxLength       =   4
            TabIndex        =   3
            Top             =   690
            Width           =   540
         End
         Begin VB.OptionButton opt_Ordenacao 
            Caption         =   "Ordenação por Identificação"
            Height          =   195
            Index           =   0
            Left            =   1680
            TabIndex        =   7
            Top             =   1875
            Width           =   3075
         End
         Begin VB.OptionButton opt_Ordenacao 
            Caption         =   "Ordenação por CEP"
            Height          =   195
            Index           =   1
            Left            =   1680
            TabIndex        =   8
            Top             =   2145
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
            Left            =   1365
            MaxLength       =   10
            TabIndex        =   4
            Top             =   1065
            Width           =   1125
         End
         Begin VB.ComboBox cboCodBarras 
            Height          =   315
            ItemData        =   "frmGeracaoSpoolIPTU.frx":001C
            Left            =   1365
            List            =   "frmGeracaoSpoolIPTU.frx":001E
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   1440
            Width           =   2985
         End
         Begin VB.CheckBox chkDebitos 
            Caption         =   "Pesquisar débitos"
            Height          =   315
            Left            =   2760
            TabIndex        =   5
            Top             =   1080
            Width           =   1635
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Composição da Receita"
            Height          =   195
            Left            =   240
            TabIndex        =   32
            Top             =   2520
            Width           =   1695
         End
         Begin VB.Label lblFinal 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Final:"
            Height          =   195
            Left            =   2685
            TabIndex        =   19
            Top             =   405
            Width           =   375
         End
         Begin VB.Label lblInicial 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Inicial:"
            Height          =   195
            Left            =   840
            TabIndex        =   18
            Top             =   390
            Width           =   450
         End
         Begin VB.Label lbl_exercicio 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Exercício:"
            Height          =   195
            Left            =   585
            TabIndex        =   20
            Top             =   765
            Width           =   720
         End
         Begin VB.Label lblData 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Data base:"
            Height          =   195
            Left            =   525
            TabIndex        =   21
            Top             =   1140
            Width           =   780
         End
         Begin VB.Label lblCodBarras 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Código Barras:"
            Height          =   195
            Left            =   255
            TabIndex        =   22
            Top             =   1530
            Width           =   1035
         End
      End
      Begin MSComDlg.CommonDialog CmdBanco 
         Left            =   5610
         Top             =   1260
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSComctlLib.ProgressBar pgr_Status 
         Height          =   195
         Left            =   1320
         TabIndex        =   23
         Top             =   6480
         Visible         =   0   'False
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   344
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin VB.Label Label1 
         Height          =   225
         Left            =   1320
         TabIndex        =   25
         Top             =   6660
         Width           =   915
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Height          =   225
         Left            =   5160
         TabIndex        =   24
         Top             =   6660
         Width           =   1095
      End
   End
   Begin MSComctlLib.ProgressBar prgStatus 
      Height          =   300
      Left            =   30
      TabIndex        =   26
      Top             =   7335
      Visible         =   0   'False
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   529
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label lblMensagens 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   285
      Left            =   30
      TabIndex        =   27
      Top             =   6990
      Width           =   7695
   End
End
Attribute VB_Name = "frmGeracaoSpoolIPTU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mobjAux           As Object
Dim mblnSelecionou    As Boolean
Dim strWord           As String

Private Sub Form_Activate()
    
    gintCodSeguranca = 1439
    
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo
    
    HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrSalvar
    
End Sub

Private Sub Form_Deactivate()
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar
End Sub

Private Sub Form_Load()
    
    CmdBanco.Filter = "Arquivo Texto | *.txt"
    
    cboCodBarras.AddItem "Febraban "
    cboCodBarras.ItemData(cboCodBarras.NewIndex) = "0"
    cboCodBarras.AddItem "Ficha Compensação "
    cboCodBarras.ItemData(cboCodBarras.NewIndex) = "1"
    
    CarregaComboComposicaoReceita
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar
End Sub

Public Sub MantemForm(ByVal strModoOperacao As String)

    If UCase(strModoOperacao) = gstrSalvar Then
        If blnDadosOk Then
            GeraArquivo
        End If
    End If
End Sub

Private Function blnDadosOk() As Boolean
    
    blnDadosOk = False
    If Trim(txt_strAvisoI) = "" Then
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
    ElseIf cboCodBarras.ListIndex < 0 Then
        ExibeMensagem "É necessário selecionar algum tipo de código de barras."
        cboCodBarras.SetFocus
        Exit Function
    ElseIf chk_GerarPeloCEP(0).Value = 0 And chk_GerarPeloCEP(1).Value = 0 Then
        ExibeMensagem "É necessário escolher uma opção de munícipio."
        chk_GerarPeloCEP(0).SetFocus
        Exit Function
    ElseIf (chk_GerarPeloCEP(0) = 1 Or chk_GerarPeloCEP(1) = 1) And (txt_strCEP(0).Text = "" Or txt_strCEP(1).Text = "") And Not (chk_GerarPeloCEP(0) = 1 And chk_GerarPeloCEP(1).Value = 1) Then
        ExibeMensagem "É necessário preencher o intervalo de CEPs."
        txt_strCEP(0).SetFocus
        Exit Function
    ElseIf CDbl(Replace(IIf(txt_strCEP(0).Text = "", 0, txt_strCEP(0).Text), "-", "")) > CDbl(Replace(IIf(txt_strCEP(1).Text = "", 0, txt_strCEP(1).Text), "-", "")) Then
        ExibeMensagem "O CEP inicial não pode ser maior que o CEP final."
        txt_strCEP(0).SetFocus
        Exit Function
    ElseIf cboComposicaoReceita.ListIndex = -1 Then
        ExibeMensagem "É necessário escolher uma composição da receita."
        cboComposicaoReceita.SetFocus
        Exit Function
    End If
    
    blnDadosOk = True
    
End Function

Private Sub GeraArquivo()
Dim strSql          As String
Dim strSQLSpool     As String
Dim strAux          As String
Dim adoLancamentoAlfa As ADODB.Recordset
Dim adoResultado    As ADODB.Recordset
Dim adoAtualizacao  As ADODB.Recordset
Dim adoParcelas     As ADODB.Recordset
Dim adoAux          As ADODB.Recordset
Dim intFebraban     As Integer
Dim INTNUMERO       As Long
Dim strCodBarras    As String
Dim strNumeroBoleto1    As String
Dim bytDigito       As Integer
Dim lngGuias        As Double
Dim intContador     As Double
Dim vetDebitos()    As String
Dim ValorParcela    As Variant

Dim PKID_ALFA       As String
Dim PKID_IPTU       As String

'Variáveis usadas para Débitos
Dim intForParcelas  As Integer
Dim intForTaxas     As Integer
Dim intCont         As Integer
Dim dblTotal        As Double
Dim strsigla        As String
Dim intExercicio    As Integer
Dim strAviso        As String
Dim strPkidAlfa     As String
Dim blnPrimeiraVez  As Boolean

Dim strInscricao    As String

Dim dblFMPCotaUnica As Double
Dim dblFMPParcelas  As Double
Dim strDtmUltimaReforma As String

    pgr_Status.Value = 0
    strWord = ""
    blnPrimeiraVez = False
    dblTotal = 0

On Error GoTo Gravar
    
    CmdBanco.ShowSave
    
    Screen.MousePointer = vbArrow
    If Trim(CmdBanco.Filename) = "" Then Exit Sub
    
    'Query utilizada para pegar o Codigo Febraban da tblEmpresa
    strSql = "SELECT intFebraban FROM " & gstrEmpresa
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        If gstrENulo(adoResultado!intFebraban) <> "" Then
            intFebraban = gstrENulo(adoResultado!intFebraban)
        Else
            ExibeMensagem "Código Febraban não encontrado."
            GoTo Gravar
        End If
    End If
    adoResultado.Close: Set adoResultado = Nothing
    
    PKID_ALFA = "(0)"
    PKID_IPTU = "(1)"

    strSQLSpool = ""
    strSQLSpool = strSQLSpool & "Select "
    strSQLSpool = strSQLSpool & "LA.pkid As IntLancamentoAlfa, LA.strinscricao, LA.strcomposicaodareceita, LA.strInscricaoAuxiliar, LA.strnomeproprietario, LA.strlogradouro, LA.strnumero, LA.strcomplemento, LA.strbairro, LA.strmunicipio, LA.struf, LA.intcep, LA.strlogradouroc, LA.strnumeroc, LA.strcomplementoc, LA.strbairroc, LA.strmunicipioc, LA.strufc, LA.intcepc, LA.strnumeroaviso, LA.strpromissario, LA.intexercicio, LA.intlancamentoalfa, LA.dblvlindexador dblIndexadorAlfa, "
    strSQLSpool = strSQLSpool & "LI.dblAreaExcedente, LI.dblValorTerrenoExcedente, LI.dblImpostoExcedente, LI.strSetor, "
    strSQLSpool = strSQLSpool & "((" & gstrISNULL("LI.dblvalorvenalterreno", "0") & ") /" & gstrISNULL("LI.dblAreaTerreno", "1") & ") as dblValorM2Terreno, "
    strSQLSpool = strSQLSpool & "(" & gstrISNULL("LI.dblareaterreno", "0") & " + " & gstrISNULL("LI.dblareaexcedente", "0") & ") as dblAreaTotalTerreno, "
    strSQLSpool = strSQLSpool & "((" & gstrISNULL("LI.dblimpostoterreno", "0") & " + " & gstrISNULL("LI.dblimpostoexcedente", "0") & ") /" & gstrISNULL("LA.dblvlindexador", "1") & ") as dblimpostoterreno, "
    strSQLSpool = strSQLSpool & "((" & gstrISNULL("LI.dblvalorterrenoexcedente", "0") & " + " & gstrISNULL("LI.dblvalorvenalterreno", "0") & ") /" & gstrISNULL("LA.dblvlindexador", "1") & ") as dblvalorvenalterreno, "
    strSQLSpool = strSQLSpool & "LI.pkid, LI.intlancamentoalfa, LI.strlote, LI.strquadra, LI.strloteamento, "
    strSQLSpool = strSQLSpool & gstrISNULL("LV.QtdeParcelas", "0") & " As QtdeParcelas, "
    strSQLSpool = strSQLSpool & "((" & gstrISNULL("LI.dblimpostoterreno", "0") & " + " & gstrISNULL("LI.dblimpostoexcedente", "0") & ") /" & gstrISNULL("LA.dblvlindexador", "1") & ") + (" & gstrISNULL("LPI.Dblimposto", "0") & "/" & gstrISNULL("LA.dblvlindexador", "1") & ") as dblTotTributo, "
    strSQLSpool = strSQLSpool & "LV.dblTotTributo dblTotTributo1, "
    strSQLSpool = strSQLSpool & "LV.dtmPrimeiroVenc, PA.INTCONTABANCARIA, "
    strSQLSpool = strSQLSpool & "(" & gstrISNULL("LPI.Dblvalorvenalpredio", "0") & "/" & gstrISNULL("LA.dblvlindexador", "1") & ") As Dblvalorvenalpredio, "
    strSQLSpool = strSQLSpool & "((" & gstrISNULL("LPI.dblvalorvenalpredio", "0") & ") /" & gstrISNULL("LPI.dblmedidadaarea", "1") & ") as dblValorM2Predio, "
    strSQLSpool = strSQLSpool & gstrISNULL("LPI.Dblmedidadaarea", "0") & " As Dblmedidadaarea, "
    strSQLSpool = strSQLSpool & gstrISNULL("LPI.Dblfatorobsolescencia", "0") & " As Dblfatorobsolescencia, "
    strSQLSpool = strSQLSpool & gstrISNULL("LPI.Dblimposto", "0") & "/" & gstrISNULL("LA.dblvlindexador", "1") & " As Dblimposto, "
    strSQLSpool = strSQLSpool & gstrISNULL("LPI.intNPavimento", "0") & " As intNPavimento, "
    strSQLSpool = strSQLSpool & gstrISNULL("LPI.intAndar", "0") & " As intAndar, "
    strSQLSpool = strSQLSpool & gstrISNULL("LPI.intElevador", "0") & " As intElevador, "
    strSQLSpool = strSQLSpool & gstrISNULL("LPI.intGaragem", "0") & " As intGaragem, "
    strSQLSpool = strSQLSpool & gstrISNULL("LPI.intSuite", "0") & " As intSuite, "
    strSQLSpool = strSQLSpool & gstrISNULL("LPI.intAmbientes", "0") & " As intAmbientes, "
    strSQLSpool = strSQLSpool & " LTI.strMedidaDaTestada As TestadaPrincipal, "
    strSQLSpool = strSQLSpool & "(Select " & gstrISNULL("Ltrim(Rtrim(FQ.strSetor))", "''") & strCONCAT & "' '" & strCONCAT & gstrISNULL("Ltrim(Rtrim(FQ.strQuadra))", "''") & strCONCAT & "' '" & strCONCAT & gstrISNULL("LO.strCodigo", "''") & " FROM " & gstrFaceDeQuadra & " FQ, " & gstrLogradouro & " LO Where FQ.Pkid = LTI.intFacedequadra and LO.Pkid " & strOUTJOracle & " =" & strOUTJSQLServer & " FQ.intLogradouro) As CodFaceDeQuadra,"
    strSQLSpool = strSQLSpool & "(Select dblFator FROM " & gstrLancamentoFatores & " LF Where LF.intLancamentoIPTU = LI.Pkid and UPPER(LF.strDescricao) Like '%TOPOGR%') As dblFatorTopografia, "
    strSQLSpool = strSQLSpool & "(Select dblFator FROM " & gstrLancamentoFatores & " LF Where LF.intLancamentoIPTU = LI.Pkid and UPPER(LF.strDescricao) Like '%EQUIPA%') As dblFatorEquipamentos, "
    strSQLSpool = strSQLSpool & "(Select dblFator FROM " & gstrLancamentoFatores & " LF Where LF.intLancamentoIPTU = LI.Pkid and UPPER(LF.strDescricao) Like '%SUPERF%') As dblFatorSuperficie, "
    strSQLSpool = strSQLSpool & "(Select dblFator FROM " & gstrLancamentoFatores & " LF Where LF.intLancamentoIPTU = LI.Pkid and UPPER(LF.strDescricao) Like '%SITUAÇ%') As dblFatorSituacao, "
    strSQLSpool = strSQLSpool & "(Select dblFator FROM " & gstrLancamentoFatores & " LF Where LF.intLancamentoIPTU = LI.Pkid and UPPER(LF.strDescricao) Like '%ACESSI%') As dblFatorAcessibilidade, "
    strSQLSpool = strSQLSpool & "(Select dblFator FROM " & gstrLancamentoFatores & " LF Where LF.intLancamentoIPTU = LI.Pkid and UPPER(LF.strDescricao) Like '%GLEBA%') As dblFatorGleba, "
    strSQLSpool = strSQLSpool & "(Select dblFator FROM " & gstrLancamentoFatores & " LF Where LF.intLancamentoIPTU = LI.Pkid and UPPER(LF.strDescricao) Like '%PROFUN%') As dblFatorProfundidade, "
    
    'Depois colocar no padrao com gstrConvert, mas é preciso criar a constante para money
    If bytDBType = SQLServer Then
        strSQLSpool = strSQLSpool & "(Select Sum(CONVERT(money, LTI2.strMedidaDaTestada)) FROM tblLancamentoTestadasIPTU LTI2, " & gstrTipoDeTestada & " TT WHERE LTI2.intLancamentoIPTU = LI.Pkid and LTI2.intTipoDeTestada = TT.Pkid and TT.bytPrincipal = 0) As DemaisTestadas, "
    Else
        strSQLSpool = strSQLSpool & "(Select Sum(" & gstrCONVERT(CDT_numeric, "LTI2.strMedidaDaTestada") & ") FROM tblLancamentoTestadasIPTU LTI2, " & gstrTipoDeTestada & " TT WHERE LTI2.intLancamentoIPTU = LI.Pkid and LTI2.intTipoDeTestada = TT.Pkid and TT.bytPrincipal = 0) As DemaisTestadas, "
    End If
    
    strSQLSpool = strSQLSpool & "(" & gstrISNULL("LI.dblvalorterrenoexcedente", "0") & " + " & gstrISNULL("LPI.Dblvalorvenalpredio", "0") & " + " & gstrISNULL("LI.Dblvalorvenalterreno", "0") & ") / " & gstrISNULL("LA.dblvlindexador", "1") & " as dblvalorVenalTotal "
    strSQLSpool = strSQLSpool & "From "
    strSQLSpool = strSQLSpool & gstrLancamentoAlfa & " LA, "
    strSQLSpool = strSQLSpool & gstrLancamentoIPTU & " LI, "
    strSQLSpool = strSQLSpool & gstrParametroAtualizacao & " PA, "
    strSQLSpool = strSQLSpool & gstrImobiliario & " I, "
    strSQLSpool = strSQLSpool & "( "
    strSQLSpool = strSQLSpool & "Select LPI.Intlancamentoiptu, "
    strSQLSpool = strSQLSpool & "Sum(" & gstrISNULL("LPI.Dblvalorvenalpredio", "0") & ") As Dblvalorvenalpredio, "
    strSQLSpool = strSQLSpool & "Sum(" & gstrISNULL("LPI.Dblmedidadaarea", "0") & ") As Dblmedidadaarea, "
    strSQLSpool = strSQLSpool & "Sum(" & gstrISNULL("LPI.Dblfatorobsolescencia", "0") & ") As Dblfatorobsolescencia, "
    strSQLSpool = strSQLSpool & "Sum(" & gstrISNULL("LPI.Dblimposto", "0") & ") As Dblimposto, "
    strSQLSpool = strSQLSpool & "Sum(" & gstrISNULL("LPI.intNPavimento", "0") & ") As intNPavimento, "
    strSQLSpool = strSQLSpool & "Sum(" & gstrISNULL("LPI.intAndar", "0") & ") As intAndar, "
    strSQLSpool = strSQLSpool & "Sum(" & gstrISNULL("LPI.intElevador", "0") & ") As intElevador, "
    strSQLSpool = strSQLSpool & "Sum(" & gstrISNULL("LPI.intGaragem", "0") & ") As intGaragem, "
    strSQLSpool = strSQLSpool & "Sum(" & gstrISNULL("LPI.intSuite", "0") & ") As intSuite, "
    strSQLSpool = strSQLSpool & "Sum(" & gstrISNULL("LPI.intQuarto", "0") & " + " & gstrISNULL("LPI.intSala", "0") & " + " & gstrISNULL("LPI.intCozinha", "0") & " + " & gstrISNULL("LPI.intBanheiro", "0") & ") As intAmbientes "
    strSQLSpool = strSQLSpool & "From " & gstrLancamentoPredioIPTU & " LPI "
    strSQLSpool = strSQLSpool & "Where "
    strSQLSpool = strSQLSpool & "LPI.Intlancamentoiptu = (" & PKID_IPTU & ") "
    strSQLSpool = strSQLSpool & "Group By LPI.Intlancamentoiptu "
    strSQLSpool = strSQLSpool & ") LPI, "
    
    strSQLSpool = strSQLSpool & "( "
    strSQLSpool = strSQLSpool & "Select LTI.strMedidaDaTestada, LTI.intFacedequadra, LTI.Intlancamentoiptu "
    strSQLSpool = strSQLSpool & "From tblLancamentoTestadasIPTU LTI, " & gstrTipoDeTestada & " TT "
    strSQLSpool = strSQLSpool & "Where TT.Pkid = LTI.intTipoDeTestada And "
    strSQLSpool = strSQLSpool & "TT.bytPrincipal = 1 "
    strSQLSpool = strSQLSpool & ") LTI, "
    
    strSQLSpool = strSQLSpool & "( "
    strSQLSpool = strSQLSpool & "Select count(LV.Intparcela) As QtdeParcelas, LV.Intlancamentoalfa, Sum(lv.dblvalor) dblTotTributo, Min(lv.dtmdtvencimento) dtmPrimeiroVenc "
    strSQLSpool = strSQLSpool & "From " & gstrLancamentoValor & " LV "
    strSQLSpool = strSQLSpool & "Where LV.Intlancamentoalfa = (" & PKID_ALFA & ") And "
    strSQLSpool = strSQLSpool & "LV.Bitparcelavalida = 1 "
    strSQLSpool = strSQLSpool & "Group By LV.Intlancamentoalfa "
    strSQLSpool = strSQLSpool & ") LV "

    strSQLSpool = strSQLSpool & "Where "
    strSQLSpool = strSQLSpool & "LA.Pkid = Li.Intlancamentoalfa And "
    strSQLSpool = strSQLSpool & "LI.Pkid " & strOUTJSQLServer & "= LPI.Intlancamentoiptu" & strOUTJOracle & " And "
    strSQLSpool = strSQLSpool & "LI.Pkid " & strOUTJSQLServer & "= LTI.Intlancamentoiptu" & strOUTJOracle & " And "
    strSQLSpool = strSQLSpool & "LA.Pkid " & strOUTJSQLServer & "= LV.IntlancamentoAlfa" & strOUTJOracle & " And "
    strSQLSpool = strSQLSpool & " PA.intComposicaoReceita = LA.intComposicaoDaReceita "
    strSQLSpool = strSQLSpool & " AND PA.intExercicio = LA.intExercicio "
    strSQLSpool = strSQLSpool & " AND I.strInscricao = LA.strInscricao "
    
    strSQLSpool = strSQLSpool & " AND LA.Pkid = (" & PKID_ALFA & ") "
    
    Dim A As Integer
    
    'Vamos primeiro obter todas as chaves, para nao realizar os filtros com todos os campos, somente com os registros necessarios
    strSql = ""
    strSql = strSql & "Select LA.pkid As IntLancamentoAlfa, LI.Pkid intLancamentoIPTU "
    strSql = strSql & "From "
    strSql = strSql & gstrLancamentoAlfa & " LA, "
    strSql = strSql & gstrLancamentoIPTU & " LI, "
    strSql = strSql & gstrImobiliario & " I "
    strSql = strSql & "Where "
    strSql = strSql & "LA.Pkid = Li.Intlancamentoalfa And "
    strSql = strSql & "LA.intExercicio = " & txt_intExercicio & " And "
    strSql = strSql & "LA.dtmdtcancelamento is null" & " And "
    strSql = strSql & "LA.strNumeroAviso Between '" & String(gintLenNumAviso - Len(Trim(txt_strAvisoI.Text)), "0") & txt_strAvisoI.Text & "' And '" & String(gintLenNumAviso - Len(Trim(txt_strAvisoF.Text)), "0") & txt_strAvisoF.Text & "' "
    strSql = strSql & " AND I.strInscricao = LA.strInscricao "
    If opt_GerarPeloEndereco(0).Value = True Or opt_GerarPeloEndereco(1).Value = True Then
        strSql = strSql & " AND I.strLogradouroC IS " & IIf(opt_GerarPeloEndereco(0), "NOT NULL ", "NULL ")
    End If
    If chk_GerarPeloCEP(0).Value = 0 Or chk_GerarPeloCEP(1).Value = 0 Then
        strSql = strSql & " AND I.intCEPC "
        If chk_GerarPeloCEP(0).Value = 1 Then
            strSql = strSql & "NOT "
        End If
        strSql = strSql & "BETWEEN " & Replace(txt_strCEP(0), "-", "") & " AND " & Replace(txt_strCEP(1), "-", "")
    End If
    strSql = strSql & " AND LA.intComposicaoDaReceita = " & cboComposicaoReceita.ItemData(cboComposicaoReceita.ListIndex)
    
    If opt_Ordenacao(0).Value = True Then
        strSql = strSql & " Order By La.strinscricao"
    Else
        strSql = strSql & " Order By La.intcepc,la.strinscricao"
    End If
    
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 10, adoLancamentoAlfa) Then
'        With adoResultado
            If Not adoLancamentoAlfa.EOF Then
                Open CmdBanco.Filename For Output As #1
                pgr_Status.Visible = True
                pgr_Status.Max = Abs(adoLancamentoAlfa.RecordCount)
                Label2.Caption = adoLancamentoAlfa.RecordCount
                
                Do While Not adoLancamentoAlfa.EOF
                                        
                    'Vamos substituir os campos variaveis pre definidos, com os pkids
                    strSQLSpool = Replace(strSQLSpool, PKID_ALFA, "(" & adoLancamentoAlfa("intLancamentoAlfa") & ")")
                    strSQLSpool = Replace(strSQLSpool, PKID_IPTU, "(" & adoLancamentoAlfa("intLancamentoIptu") & ")")
                    
                    PKID_ALFA = "(" & adoLancamentoAlfa("intLancamentoAlfa") & ")"
                    PKID_IPTU = "(" & adoLancamentoAlfa("intLancamentoIptu") & ")"
                    
                    If gobjBanco.CriaADO(strSQLSpool, 10, adoResultado) Then
                        With adoResultado
                        
                        'Vamos verificar se a conta bancaria esta definida para esta composicao
                        If IsNull(!intContaBancaria) And cboCodBarras.ItemData(cboCodBarras.ListIndex) = FICHA_COMPENSACAO Then
                            ExibeMensagem "Não foi encontrada Conta Bancária para a Composição " & !strComposicaoDaReceita & " no Exercício de " & !intExercicio & "."
                            GoTo Gravar
                        End If
                        
                        If CDbl(gstrConvVrDoSql(gstrENulo(!dblTotTributo1), , , True)) > 0 Then
                            strWord = ""
                            'Código do tributo
                            strWord = strWord & Space$(3 - Len("10")) & "10"
                            'Descrição do Tributo
                            strWord = strWord & Left(gstrENulo(!strComposicaoDaReceita), 40) & Space$(40 - Len(Left(gstrENulo(!strComposicaoDaReceita), 40)))
                            'Exercício
                            strWord = strWord & gstrENulo(!intExercicio) & Space$(4 - Len(gstrENulo(!intExercicio)))
                            'Nº do Aviso
                            strWord = strWord & Format$(gstrENulo(Right(!strNumeroAviso, 6)), "000,000")
                            'Nº de Parcelas
                            strWord = strWord & Format$(!QtdeParcelas, "00")
                            'Tipo de Debito
                            strWord = strWord & " 0"
                            'Código de Atraso
                            strWord = strWord & "1"
                            'Data do Dia
                            strWord = strWord & Format$(!dtmPrimeiroVenc, "dd/mm/yyyy")
                            'Nome do Proprietario
                            strWord = strWord & Left(gstrENulo(!strnomeproprietario), 40) & Space$(40 - Len(Left(gstrENulo(!strnomeproprietario), 40)))
                            'Inscricao
                            strInscricao = gstrFormataInscricao(Right(gstrENulo(!strInscricao), gintRetornaTamanhoMascara(TYP_IMOBILIARIA)))
                            strWord = strWord & Left(gstrENulo(strInscricao), 10) & Space$(10 - Len(Left(gstrENulo(strInscricao), 10)))
                            'Nome do Promissário
                            strWord = strWord & Left(gstrENulo(!strpromissario), 40) & Space$(40 - Len(Left(gstrENulo(!strpromissario), 40)))
                            'Endereço do Local
                            strAux = Trim(gstrENulo(!strLogradouro)) & " " & Trim(gstrENulo(!strNumero)) & " " & Trim(gstrENulo(!STRCOMPLEMENTO)) & " " & Trim(gstrENulo(!strBairro))
                            strWord = strWord & Left(strAux, 73) & Space$(73 - Len(Left(strAux, 73)))
                            strAux = ""
                            'Nº CEP
                            strWord = strWord & Format$(IIf(Trim(gstrENulo(!INTCEP)) = "", "0", Trim(gstrENulo(!INTCEP))), "00000000")
                            'Endereço de correspondência
                            strAux = Trim(gstrENulo(!strLogradouroC)) & " " & Trim(gstrENulo(!strNumeroC)) & " " & Trim(gstrENulo(!strComplementoC)) & " " & Trim(gstrENulo(!strBairroC))
                            strWord = strWord & Left(strAux, 73) & Space$(73 - Len(Left(strAux, 73)))
                            strAux = ""
                            'Endereço de correspondência
                            strAux = Trim(gstrENulo(!strMunicipioC)) & " - " & Trim(gstrENulo(!strUFC)) & " CEP : " & Trim(gstrENulo(!strComplementoC)) & " " & gstrCEPFormatado(Trim(gstrENulo(!intcepc)))
                            strWord = strWord & Left(strAux, 73) & Space$(73 - Len(Left(strAux, 73)))
                            strAux = ""
                            'Área do terreno
                            strWord = strWord & Space$(15 - Len(Left(gstrConvVrDoSql(gstrENulo(!DblAreaTotalTerreno), , , True), 15))) & Left(gstrConvVrDoSql(gstrENulo(!DblAreaTotalTerreno), , , True), 15)
                            'Valor Venal terreno
                            strWord = strWord & Space$(18 - Len(Left(gstrConvVrDoSql(gstrENulo(!dblValorVenalTerreno), 4, , True), 18))) & Left(gstrConvVrDoSql(gstrENulo(!dblValorVenalTerreno), 4, , True), 18)
                            'Valor Imposto Territorial
                            strWord = strWord & Space$(18 - Len(Left(gstrConvVrDoSql(gstrENulo(!dblimpostoterreno), 4, , True), 18))) & Left(gstrConvVrDoSql(gstrENulo(!dblimpostoterreno), 4, , True), 18)
                            'Área dos Prédios
                            strWord = strWord & Space$(14 - Len(Left(gstrConvVrDoSql(gstrENulo(!Dblmedidadaarea), 2, , True), 14))) & Left(gstrConvVrDoSql(gstrENulo(!Dblmedidadaarea), 2, , True), 14)
                            'Valor venal Prédio
                            strWord = strWord & Space$(18 - Len(Left(gstrConvVrDoSql(gstrENulo(!Dblvalorvenalpredio), 4, , True), 18))) & Left(gstrConvVrDoSql(gstrENulo(!Dblvalorvenalpredio), 4, , True), 18)
                            'Valor Imposto Predial
                            strWord = strWord & Space$(18 - Len(Left(gstrConvVrDoSql(gstrENulo(!dblImposto), 4, , True), 18))) & Left(gstrConvVrDoSql(gstrENulo(!dblImposto), 4, , True), 18)
                            'Valor Venal Total
                            strWord = strWord & Space$(18 - Len(Left(gstrConvVrDoSql(gstrENulo(!dblvalorVenalTotal), 4, , True), 18))) & Left(gstrConvVrDoSql(gstrENulo(!dblvalorVenalTotal), 4, , True), 18)
                            'Valor Total do Lancamento
                            strWord = strWord & Space$(18 - Len(Left(gstrConvVrDoSql(gstrENulo(!dblTotTributo), 4, , True), 18))) & Left(gstrConvVrDoSql(gstrENulo(!dblTotTributo), 4, , True), 18)
                            
                            'Vamos preencher as colunas de parcelas
                            strSql = ""
                            strSql = strSql & "Select "
                            strSql = strSql & "LV.Pkid, LV.Intlancamentoalfa, "
                            strSql = strSql & "LV.INTPARCELA, "
                            strSql = strSql & "LV.Dtmdtvencimento, "
                            strSql = strSql & "LV.Dblvalor, "
                            strSql = strSql & "LV.Bitparcelavalida "
                            strSql = strSql & "From "
                            strSql = strSql & gstrLancamentoValor & " LV " & strREADPAST
                            strSql = strSql & " Where "
                            strSql = strSql & "LV.Intlancamentoalfa = " & !intLancamentoAlfa
                            strSql = strSql & "Order By "
                            strSql = strSql & "LV.Intlancamentoalfa, "
                            strSql = strSql & "LV.Bitparcelavalida, "
                            strSql = strSql & "LV.Intparcela "
                            
                            Set gobjBanco = New clsBanco
                            If gobjBanco.CriaADO(strSql, 10, adoAux) Then
                                If Not .EOF Then
                                    
                                    Set gobjBanco = New clsBanco
                                    gobjBanco.ExecutaBeginTrans
                                    
                                    'Do While Not adoAux.EOF
                                    For intForParcelas = 1 To 13 '12 parcelas + 1 unica
                                    
                                        If Not adoAux.EOF Then
                                            If intForParcelas = 1 Then
                                                dblFMPCotaUnica = gstrConvVrDoSql(gstrENulo(adoAux!dblValor), 2, , True)
                                            ElseIf intForParcelas = 2 Then
                                                dblFMPParcelas = gstrConvVrDoSql(gstrENulo(adoAux!dblValor), 2, , True)
                                            End If
                                            
                                            'Nº da parcela
                                            strWord = strWord & Format$(gstrENulo(adoAux!intParcela), "00")
                                            'Data de Vencimento
                                            strWord = strWord & Format$(gstrENulo(adoAux!Dtmdtvencimento), "dd/mm/yyyy")
                                            'Valor das parcelas
                                            strWord = strWord & Space$(18 - Len(Left(gstrConvVrDoSql(gstrENulo(adoAux!dblValor), 2, , True), 18))) & Left(gstrConvVrDoSql(gstrENulo(adoAux!dblValor), 2, , True), 18)
                                            
                                            INTNUMERO = glngRetornaProximoNumeroGuia
                                            
                                            ValorParcela = adoAux!dblValor
                                            If InStr(ValorParcela, ",") = 0 Then
                                                ValorParcela = gstrConvVrDoSql(ValorParcela)
                                            Else
                                                If Len(ValorParcela) - InStr(ValorParcela, ",") < 2 Then
                                                    ValorParcela = gstrConvVrDoSql(ValorParcela)
                                                End If
                                            End If
                                            
                                            'Vamos definir o codigo de barras
                                            strCodBarras = gstrMontaCodigoBarras(cboCodBarras.ItemData(cboCodBarras.ListIndex), IIf(IsNull(!intContaBancaria), 0, !intContaBancaria), ValorParcela, adoAux!Dtmdtvencimento, intFebraban, INTNUMERO, True, adoAux!bitParcelaValida <> 0)
                                            'Alteração Feita em 20/12/2005 para geração do Spool de GRJ
                                            'strCodBarras = gstrMontaCodigoBarras(cboCodBarras.ItemData(cboCodBarras.ListIndex), !intContaBancaria, adoAux!dblValor, adoAux!Dtmdtvencimento, intFebraban, INTNUMERO, True, 1)
                                            
                                            If Len(strCodBarras) = 0 Then
                                                gobjBanco.ExecutaRollbackTrans
                                                GoTo Gravar
                                            End If
                                            'Vamos definir a linha digitavel
                                            strNumeroBoleto1 = gstrMontaLinhaDigitavel(cboCodBarras.ItemData(cboCodBarras.ListIndex), strCodBarras)
                                            strNumeroBoleto1 = Replace(strNumeroBoleto1, "-", "")
                                            
                                            'Vamos tratar a linha digitavel de acordo com o tipo de codigo de barras
                                            If cboCodBarras.ItemData(cboCodBarras.ListIndex) = FICHA_COMPENSACAO Then
                                                strNumeroBoleto1 = Replace(Replace(strNumeroBoleto1, " ", ""), ".", "")
                                            End If
                                            
                                            strWord = strWord & strCodBarras
                                            strWord = strWord & Left(gstrENulo(strNumeroBoleto1), 51) & Space$(51 - Len(Left(gstrENulo(strNumeroBoleto1), 51)))
                                            
                                            'Vamos inserir a Tblguias e TbllancamentoGuia
            
                                            'Vamos inserir a guia na tabela TblGuias
                                            strSql = ""
                                            'strSql = IIf((bytDBType = EDatabases.Oracle), "BEGIN ", "")
                                            strSql = strSql & "Insert Into " & gstrGuias & "("
                                            'strSql = strSql & "Pkid, "
                                            strSql = strSql & "Intcontabancaria, "
                                            strSql = strSql & "Intnumero, "
                                            strSql = strSql & "Dtmdtemissao, "
                                            strSql = strSql & "Dblvalor, "
                                            strSql = strSql & "Strcodbarra, "
                                            strSql = strSql & "Dtmdtatualizacao, "
                                            strSql = strSql & "Lngcodusr, "
                                            strSql = strSql & "Dtmdtvencimento "
                                            strSql = strSql & ") Values("
                                            'strSql = strSql & lngGuias & ", "
                                            strSql = strSql & gstrENulo(!intContaBancaria, , True) & ", "
                                            strSql = strSql & INTNUMERO & ", "
                                            strSql = strSql & gstrConvDtParaSql(gstrDataDoSistema) & ", "
                                            strSql = strSql & gstrConvVrParaSql(adoAux!dblValor) & ", '"
                                            strSql = strSql & strCodBarras & "', "
                                            strSql = strSql & gstrConvDtParaSql(gstrDataDoSistema) & ", "
                                            strSql = strSql & glngCodUsr & ", "
                                            strSql = strSql & gstrConvDtParaSql(adoAux!Dtmdtvencimento)
                                            strSql = strSql & ")"
                                            'strSql = strSql & IIf((bytDBType = EDatabases.Oracle), " ; ", "")
                                            
                                            If Not gobjBanco.Execute(strSql) Then
                                                ExibeMensagem "Erro na gravação da guia."
                                                gobjBanco.ExecutaRollbackTrans
                                                GoTo Gravar
                                            End If
            
                                            'Vamos inserir as parcelas na tabela TblLancamentoGuias
                                            
                                            lngGuias = glngRetornaPkidTabelaPai("seqtblGuias", "tblGuias")
                                            
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
                                            strSql = strSql & adoAux!Pkid & ", "
                                            strSql = strSql & lngGuias & ","
                                            strSql = strSql & gstrConvVrParaSql(adoAux!dblValor) & ", "
                                            strSql = strSql & "0.00" & ", "
                                            strSql = strSql & "0.00" & ", "
                                            strSql = strSql & "0.00" & ", "
                                            strSql = strSql & "0.00" & ", "
                                            strSql = strSql & strGETDATE & ", "
                                            strSql = strSql & glngCodUsr & ") "
                                            'strSql = strSql & IIf((bytDBType = EDatabases.Oracle), ";", "")
                                        
                                            'strSql = strSql & IIf((bytDBType = EDatabases.Oracle), " END; ", "")
                                            
                                            If Not gobjBanco.Execute(strSql) Then
                                                ExibeMensagem "Erro na gravação da guia."
                                                gobjBanco.ExecutaRollbackTrans
                                                GoTo Gravar
                                            End If
                                            
                                            adoAux.MoveNext
                                        
                                        Else
                                            'Nº da parcela
                                            strWord = strWord & "00"
                                            'Data de Vencimento
                                            strWord = strWord & "00/00/0000"
                                            'Valor das parcelas
                                            strWord = strWord & String(18, " ")
                                            'Codigo de barras
                                            strWord = strWord & String(44, "0")
                                            strWord = strWord & String(51, "0")
                                        End If
                                    'Loop
                                    Next
                                    
                                Else
                                    ExibeMensagem "Não foram encontrados registros com esses parâmetros."
                                    GoTo Gravar
                                End If
                            End If
                            
                            'Caso seja necessario pesquisar debitos
                            If chkDebitos.Value = vbChecked Then
                            
                                'Vamos verificar se há débitos para a inscrição
                                If gobjBanco.CriaADO(strAtualizacaoDeDebitosSpool(!strInscricao), 10, adoAtualizacao) Then
                                    If Not adoAtualizacao.EOF Then
                                        With adoAtualizacao
                                        
                                            Do While Not adoAtualizacao.EOF
                                                
                                                'Caso tenha acordo vamos desprezar
                                                If IsNull(adoAtualizacao!intlancamentoalfaacordo) Then
                            
                                                    strSql = gstrStoredProcedure("sp_AtualizaParcela", adoAtualizacao!intComposicaoDaReceita & ", " & adoAtualizacao!intExercicio & ", " & adoAtualizacao!intParcela & ", " & gstrConvDtParaSql(adoAtualizacao!Dtmdtvencimento) & ", " & gstrConvDtParaSql(txt_dtmDtBaixa) & ", " & gstrConvVrParaSql(adoAtualizacao!ValorOrig) & ", " & adoAtualizacao!intMoeda, True)
                                                    
                                                    Set gobjBanco = New clsBanco
                                                    If gobjBanco.CriaADO(strSql, 80, adoParcelas) Then
                                                    
                                                        If intCont < 14 Then
                                                            If adoAtualizacao("PkidLV").Value > 0 Then
                                                                If adoAtualizacao("bitParcelaValida").Value = 1 Then
                                                                    If Not blnPrimeiraVez Then
                                                                    
                                                                        dblTotal = dblTotal + CDbl(gstrConvVrDoSql(CCur(gstrConvVrDoSql(adoParcelas("dblValorPrincipal").Value)) + CCur(gstrConvVrDoSql(adoParcelas("dblValorMulta").Value)) + CCur(gstrConvVrDoSql(adoParcelas("dblValorJuros").Value)) + CCur(gstrConvVrDoSql(adoParcelas("dblValorCorrecao").Value)), 4, , True))
                                                                        strsigla = Mid(Trim(gstrENulo(adoAtualizacao("strSigla").Value)), 1, 4)
                                                                        intExercicio = adoAtualizacao("intExercicio").Value
                                                                        strAviso = Right(adoAtualizacao("strNumeroAviso").Value, 6)
                                                                        strPkidAlfa = adoAtualizacao("intLancamentoAlfa").Value
                                                                        blnPrimeiraVez = True
                                                                    Else
                                                                        If strPkidAlfa = adoAtualizacao("intLancamentoAlfa").Value Then
                                                                            dblTotal = dblTotal + CDbl(gstrConvVrDoSql(CCur(gstrConvVrDoSql(adoParcelas("dblValorPrincipal").Value)) + CCur(gstrConvVrDoSql(adoParcelas("dblValorMulta").Value)) + CCur(gstrConvVrDoSql(adoParcelas("dblValorJuros").Value)) + CCur(gstrConvVrDoSql(adoParcelas("dblValorCorrecao").Value)), 4, , True))
                                                                        Else
                                                                            strWord = strWord & strsigla        'Sigla da Composição da Receita
                                                                            strWord = strWord & intExercicio    'Exercicio
                                                                            strWord = strWord & Format$(strAviso, "000,000") 'Aviso
                                                                            strWord = strWord & Space$(15 - Len(Left(gstrConvVrDoSql(dblTotal), 15))) & Left(gstrConvVrDoSql(dblTotal), 15) 'Valor da Divida
                                                                            intCont = intCont + 1
                                                                            dblTotal = 0
                                                                            dblTotal = dblTotal + CDbl(gstrConvVrDoSql(CCur(gstrConvVrDoSql(adoParcelas("dblValorPrincipal").Value)) + CCur(gstrConvVrDoSql(adoParcelas("dblValorMulta").Value)) + CCur(gstrConvVrDoSql(adoParcelas("dblValorJuros").Value)) + CCur(gstrConvVrDoSql(adoParcelas("dblValorCorrecao").Value)), 4, , True))
                                                                            strsigla = Mid(Trim(gstrENulo(adoAtualizacao("strSigla").Value)), 1, 4)
                                                                            intExercicio = adoAtualizacao("intExercicio").Value
                                                                            strAviso = Right(adoAtualizacao("strNumeroAviso").Value, 6)
                                                                            strPkidAlfa = adoAtualizacao("intLancamentoAlfa").Value
                                                                        End If
                                                                    End If
                                                                End If
                                                            End If
                                                        Else
                                                            Exit Do
                                                        End If
                                                    
                                                        adoAtualizacao.MoveNext
                                                    End If
                                                Else
                                                    adoAtualizacao.MoveNext
                                                End If
                                                
                                            Loop
                                            
                                            strWord = strWord & strsigla
                                            strWord = strWord & intExercicio
                                            strWord = strWord & Format$(strAviso, "000,000")
                                            strWord = strWord & Space$(15 - Len(Left(gstrConvVrDoSql(dblTotal), 15))) & Left(gstrConvVrDoSql(dblTotal), 15)
                                            
                                            dblTotal = 0
                                            blnPrimeiraVez = False
                                            
                                            If intCont < 14 Then
                                                intCont = intCont + 1
                                            End If
                                            
                                        End With
                                    End If
                                End If
                                
                            End If
                            
                            'Vamos preencher em branco os espacos reservados para as dividas nao preenchidas
                            strWord = strWord & String(420 - (intCont * 30), " ")
                            
                            strWord = strWord & String(145, " ")
                            
                            'Valor cota unica fator FMP
                            dblFMPCotaUnica = TruncaValores(gstrConvVrDoSql((dblFMPCotaUnica / IIf(IsNull(!dblIndexadorAlfa), 1, !dblIndexadorAlfa)), 4, , True), 4)
                            strWord = strWord & Space$(17 - Len(Left(gstrConvVrDoSql(dblFMPCotaUnica, 4), 17))) & Left(gstrConvVrDoSql(dblFMPCotaUnica, 4), 17)
                            'Valor parcelas fator FMP
                            dblFMPParcelas = TruncaValores(gstrConvVrDoSql((dblFMPParcelas / IIf(IsNull(!dblIndexadorAlfa), 1, !dblIndexadorAlfa)), 4, , True), 4)
                            strWord = strWord & Space$(17 - Len(Left(gstrConvVrDoSql(dblFMPParcelas, 4), 17))) & Left(gstrConvVrDoSql(dblFMPParcelas, 4), 17)
                            
                            strSql = " Select " & gstrTOPnSQLServer(1) & " LPI.dtmUltimaReforma, LPI.strNomeUso  From " & gstrLancamentoPredioIPTU & " LPI Where " & Val(gstrENulo(!Pkid)) & " = LPI.Intlancamentoiptu Order By LPI.Pkid "
                            strSql = gstrTOPnOracle(strSql, 1)
                            
                            Set gobjBanco = New clsBanco
                            If gobjBanco.CriaADO(strSql, 10, adoAux) Then
                                If Not .EOF Then
                                    'Ano ultima reforma
                                    If IsDate(adoAux!dtmUltimaReforma) Then
                                        strWord = strWord & Left(gstrENulo(Year(adoAux!dtmUltimaReforma)), 4) & Space$(4 - Len(Left(gstrENulo(Year(adoAux!dtmUltimaReforma)), 4)))
                                        strDtmUltimaReforma = Left(gstrENulo(Year(adoAux!dtmUltimaReforma)), 4) & Space$(4 - Len(Left(gstrENulo(Year(adoAux!dtmUltimaReforma)), 4)))
                                    Else
                                        strWord = strWord & "0000"
                                        strDtmUltimaReforma = "0000"
                                    End If
                                    'Categoria
                                    strWord = strWord & Left(gstrENulo(adoAux!strNomeUso), 25) & Space$(25 - Len(Left(gstrENulo(adoAux!strNomeUso), 25)))
                                Else
                                    strWord = strWord & "0000"
                                    strWord = strWord & Space(25)
                                End If
                            End If
                            
                            'Area terreno excedente
                            strWord = strWord & Space$(15 - Len(Left(gstrConvVrDoSql(!dblAreaExcedente), 15))) & Left(gstrConvVrDoSql(!dblAreaExcedente), 15)
                            'Valor venal excedente
                            strWord = strWord & Space$(18 - Len(Left(gstrConvVrDoSql(!dblValorTerrenoExcedente), 18))) & Left(gstrConvVrDoSql(!dblValorTerrenoExcedente), 18)
                            'Imposto territorial excedente
                            strWord = strWord & Space$(18 - Len(Left(gstrConvVrDoSql(!dblImpostoExcedente), 18))) & Left(gstrConvVrDoSql(!dblImpostoExcedente), 18)
                            'Lote
                            strWord = strWord & Left(gstrENulo(!strLote), 10) & Space$(10 - Len(Left(gstrENulo(!strLote), 10)))
                            'Quadra
                            strWord = strWord & Left(gstrENulo(!strQuadra), 10) & Space$(10 - Len(Left(gstrENulo(!strQuadra), 10)))
                            'Loteamento
                            strWord = strWord & Left(gstrENulo(!strLoteamento), 20) & Space$(20 - Len(Left(gstrENulo(!strLoteamento), 20)))
                            'Valor M2 terreno
                            strWord = strWord & Space$(15 - Len(Left(gstrConvVrDoSql(!dblValorM2Terreno), 15))) & Left(gstrConvVrDoSql(!dblValorM2Terreno), 15)
                            'Valor M2 construcao
                            strWord = strWord & Space$(15 - Len(Left(gstrConvVrDoSql(!dblValorM2Predio), 15))) & Left(gstrConvVrDoSql(!dblValorM2Predio), 15)
                            'Tipo cod. barras (Febraban = 1, Ficha C. = 2)
                            strWord = strWord & cboCodBarras.ItemData(cboCodBarras.ListIndex) + 1
                            'Inscricao Cadastral
                            strWord = strWord & Left(gstrENulo(strInscricao), 20) & Space$(20 - Len(Left(gstrENulo(strInscricao), 20)))
                            'Cod Face de quadra
                            strWord = strWord & Left(gstrENulo(!CodFaceDeQuadra), 20) & Space$(20 - Len(Left(gstrENulo(!CodFaceDeQuadra), 20)))
                            'Testada principal
                            strWord = strWord & Space$(15 - Len(Left(gstrConvVrDoSql(!TestadaPrincipal), 15))) & Left(gstrConvVrDoSql(!TestadaPrincipal), 15)
                            'Demais testadas
                            strWord = strWord & Space$(15 - Len(Left(gstrConvVrDoSql(!DemaisTestadas), 15))) & Left(gstrConvVrDoSql(!DemaisTestadas), 15)
                            
                            'Vamos preencher as colunas de taxas
                            strSql = ""
                            strSql = strSql & "Select "
                            strSql = strSql & "SUM(LR.dblValor) dblValor, R.strDescricao "
                            strSql = strSql & "From "
                            strSql = strSql & gstrLancamentoValor & " LV " & strREADPAST & " , "
                            strSql = strSql & gstrLancamentoReceita & " LR " & strREADPAST & " , "
                            strSql = strSql & gstrReceita & " R " & strREADPAST
                            strSql = strSql & " Where "
                            strSql = strSql & "LV.Intlancamentoalfa = " & !intLancamentoAlfa & " AND "
                            strSql = strSql & "LR.intLancamentoValor = LV.Pkid AND "
                            strSql = strSql & "R.Pkid = LR.intReceita AND "
                            strSql = strSql & "LV.bitParcelaValida = 1 AND "
                            strSql = strSql & "R.bytTipo = 3 " 'Tipo Taxa
                            strSql = strSql & "Group By R.strDescricao "
                            
                            Set gobjBanco = New clsBanco
                            If gobjBanco.CriaADO(strSql, 10, adoAux) Then
                                    
                                'Do While Not adoAux.EOF
                                For intForTaxas = 1 To 8
                                
                                    If Not adoAux.EOF Then
                                        'Descricao Taxa
                                        strWord = strWord & Left(gstrENulo(adoAux!strDescricao), 15) & Space$(15 - Len(Left(gstrENulo(adoAux!strDescricao), 15)))
                                        'Valor da Taxa
                                        strWord = strWord & Space$(18 - Len(Left(gstrConvVrDoSql(gstrENulo(adoAux!dblValor), 2, , True), 18))) & Left(gstrConvVrDoSql(gstrENulo(adoAux!dblValor), 2, , True), 18)
                                        
                                        adoAux.MoveNext
                                    Else
                                        'Descricao Taxa
                                        strWord = strWord & String(15, " ")
                                        'Valor da Taxa
                                        strWord = strWord & String(18, " ")
                                    End If
                                Next
                            
                            End If
                            
                            'Predio Tipo
                            strWord = strWord & String(3, " ")
                            'Predio Pavimentos
                            strWord = strWord & Left(gstrENulo(!intNPavimento), 3) & Space$(3 - Len(Left(gstrENulo(!intNPavimento), 3)))
                            'Predio Ultima Reforma
                            strWord = strWord & strDtmUltimaReforma
                            'Predio Elevador
                            strWord = strWord & Left(gstrENulo(!intElevador), 3) & Space$(3 - Len(Left(gstrENulo(!intElevador), 3)))
                            'Predio Garagem
                            strWord = strWord & Left(gstrENulo(!intGaragem), 3) & Space$(3 - Len(Left(gstrENulo(!intGaragem), 3)))
                            'Predio Suite
                            strWord = strWord & Left(gstrENulo(!intSuite), 3) & Space$(3 - Len(Left(gstrENulo(!intSuite), 3)))
                            'Predio Ambientes
                            strWord = strWord & Left(gstrENulo(!intAmbientes), 3) & Space$(3 - Len(Left(gstrENulo(!intAmbientes), 3)))
                            'Predio Andares
                            strWord = strWord & Left(gstrENulo(!intAndar), 3) & Space$(3 - Len(Left(gstrENulo(!intAndar), 3)))
                            
                            'Fator obsolescencia
                            strWord = strWord & Space$(5 - Len(Left(gstrConvVrDoSql(gstrENulo(!dblFatorObsolescencia), 2, , True), 5))) & Left(gstrConvVrDoSql(gstrENulo(!dblFatorObsolescencia), 2, , True), 5)
                            
                            'Fator topografia
                            strWord = strWord & Space$(5 - Len(Left(gstrConvVrDoSql(gstrENulo(!dblFatorTopografia), 2, , True), 5))) & Left(gstrConvVrDoSql(gstrENulo(!dblFatorTopografia), 2, , True), 5)
                            'Fator equipamentos
                            strWord = strWord & Space$(5 - Len(Left(gstrConvVrDoSql(gstrENulo(!dblFatorEquipamentos), 2, , True), 5))) & Left(gstrConvVrDoSql(gstrENulo(!dblFatorEquipamentos), 2, , True), 5)
                            'Fator superficie
                            strWord = strWord & Space$(5 - Len(Left(gstrConvVrDoSql(gstrENulo(!dblFatorSuperficie), 2, , True), 5))) & Left(gstrConvVrDoSql(gstrENulo(!dblFatorSuperficie), 2, , True), 5)
                            'Fator situação
                            strWord = strWord & Space$(5 - Len(Left(gstrConvVrDoSql(gstrENulo(!dblFatorSituacao), 2, , True), 5))) & Left(gstrConvVrDoSql(gstrENulo(!dblFatorSituacao), 2, , True), 5)
                            'Fator acessibilidade
                            strWord = strWord & Space$(5 - Len(Left(gstrConvVrDoSql(gstrENulo(!dblFatorAcessibilidade), 2, , True), 5))) & Left(gstrConvVrDoSql(gstrENulo(!dblFatorAcessibilidade), 2, , True), 5)
                            'Fator gleba
                            strWord = strWord & Space$(5 - Len(Left(gstrConvVrDoSql(gstrENulo(!dblFatorGleba), 2, , True), 5))) & Left(gstrConvVrDoSql(gstrENulo(!dblFatorGleba), 2, , True), 5)
                            'Fator profundidade
                            strWord = strWord & Space$(5 - Len(Left(gstrConvVrDoSql(gstrENulo(!dblFatorProfundidade), 2, , True), 5))) & Left(gstrConvVrDoSql(gstrENulo(!dblFatorProfundidade), 2, , True), 5)
                            
                            'Codigo valor m2 terreno
                            strWord = strWord & Left(gstrENulo(!strSetor), 2) & Space$(2 - Len(Left(gstrENulo(!strSetor), 2)))
                            'Inscricao Auxiliar
                            strWord = strWord & Left(gstrENulo(!strInscricaoAuxiliar), 25) & Space$(25 - Len(Left(gstrENulo(!strInscricaoAuxiliar), 25)))
                            
                            Print #1, strWord
                            DoEvents
                            pgr_Status.Value = adoLancamentoAlfa.AbsolutePosition
                            Label1.Caption = adoLancamentoAlfa.AbsolutePosition
                            intContador = intContador + 1
                            gobjBanco.ExecutaCommitTrans
                            adoLancamentoAlfa.MoveNext
                        
                        Else
                            DoEvents
                            pgr_Status.Value = adoLancamentoAlfa.AbsolutePosition
                            Label1.Caption = adoLancamentoAlfa.AbsolutePosition
                            Me.Refresh
                            adoLancamentoAlfa.MoveNext
                        End If
                        
                    End With
                    End If
                    
                    intCont = 0
                    
                Loop
                Close #1
            Else
                ExibeMensagem "Não foram encontrados registros com esses parâmetros."
                GoTo Gravar
            End If
        'End With
    Else
        GoTo Gravar
    End If
    
    Screen.MousePointer = vbDefault
    
    If intContador >= 1 Then
        ExibeMensagem "Arquivo gerado com sucesso com " & intContador & " boleto(s)."
    End If
    
    pgr_Status.Value = 0
    Exit Sub
    
Gravar:
    gobjBanco.ExecutaRollbackTrans
    If Len(Err.Description) > 0 Then MsgBox Err.Description
    Close #1
    'Open CmdBanco.filename For Output As #1
    'Close #1
    Screen.MousePointer = vbDefault
End Sub

Private Sub tab_Parametros_DblClick()
   frm_Arquiv_Banco.Show
End Sub

Private Sub txt_dtmDtBaixa_LostFocus()
    txt_dtmDtBaixa = gstrDataFormatada(txt_dtmDtBaixa)
End Sub

Private Sub txt_dtmDtBaixa_GotFocus()
    If txt_dtmDtBaixa.Text = "" Then txt_dtmDtBaixa = gstrDataDoSistema
    MarcaCampo txt_dtmDtBaixa
End Sub

Private Sub txt_dtmDtBaixa_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "D", txt_dtmDtBaixa
End Sub

Private Sub txt_intExercicio_GotFocus()
    MarcaCampo txt_intExercicio
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

Public Function strAtualizacaoDeDebitosSpool(strInscricao As String) As String
Dim adoResultado            As ADODB.Recordset
'Dim adoParcelas             As ADODB.Recordset
Dim strSql                  As String
Dim strAcordosParaConsulta  As String
Dim strInscricoes           As String
'Dim intFor                  As Integer
    
    'Vamos obter os Pkids das inscricoes para fazer consulta de acordos
    strSql = "SELECT  LA.Pkid " & _
             "FROM " & gstrLancamentoAlfa & " LA " & _
             "WHERE LA.strInscricao = '" & String(gintLenInscricao - Len(Trim(strInscricao)), "0") & UCase(strInscricao) & "' AND " & _
             "(LA.INTUTILIZACAO = " & TYP_ACORDO & " OR (LA.INTUTILIZACAO <> " & TYP_ACORDO & " AND LA.bytNaoInscreveda = 0)) "
    
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        If Not adoResultado.EOF Then
            Do While Not adoResultado.EOF
                strAcordosParaConsulta = strAcordosParaConsulta & adoResultado("Pkid").Value & ","
                adoResultado.MoveNext
            Loop
            strAcordosParaConsulta = Mid(strAcordosParaConsulta, 1, Len(strAcordosParaConsulta) - 1)
        End If
    Else
        Exit Function
    End If

ConsultarAcordos:

    'Vamos obter os acordos, caso exista, para exibir no grid Pai
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
    
    strSql = "SELECT LV.Pkid PkidLV, LV.bitParcelaValida, LA.intExercicio, LV.intLancamentoAlfa, LV.intParcela, "
    strSql = strSql & "LV.dblValor ValorOrig, LV.dtmDtVencimento, LV.intMoeda, LV.intLancamentoAlfaAcordo, LV.intLancamentoAlfaDAtiva, "
    strSql = strSql & "LA.strInscricao, " & gstrCONVERT(CDT_numeric, "LA.strNumeroAviso") & " strNumeroAviso, "
    strSql = strSql & "LA.intComposicaoDaReceita, CR.strSigla, LA.strComposicaoDaReceita, " & strSUBSTRING & "(LAA.strInscricao, " & gintLenInscricao - gintRetornaTamanhoMascara(TYP_ACORDO) + 1 & "," & gintRetornaTamanhoMascara(TYP_ACORDO) - 4 & ")" & strCONCAT & " '/' " & strCONCAT & " " & gstrRIGHT("LAA.strInscricao", 4) & " Acordo, LA.intUtilizacao "
    strSql = strSql & "FROM " & gstrLancamentoValor & " LV, "
    strSql = strSql & gstrLancamentoAlfa & " LA, "
    strSql = strSql & gstrAcordo & " AC, "
    strSql = strSql & gstrLancamentoAlfa & " LAA, "
    strSql = strSql & gstrLancamentoPagamento & " LP, "
    strSql = strSql & gstrComposicaoDaReceita & " CR "
    strSql = strSql & "WHERE LV.intLancamentoAlfa = LA.pkid AND "
    strSql = strSql & "LV.intLancamentoAlfaAcordo " & strOUTJSQLServer & "= AC.intLancamentoAlfa " & strOUTJOracle & " And "
    strSql = strSql & "LA.intcomposicaodareceita " & strOUTJSQLServer & "= CR.Pkid " & strOUTJOracle & " And "
    strSql = strSql & "LV.intLancamentoAlfaAcordo " & strOUTJSQLServer & "= LAA.Pkid " & strOUTJOracle & " And "
    strSql = strSql & "LV.Pkid" & strOUTJSQLServer & "= LP.Intlancamentovalor " & strOUTJOracle & " And "
    strSql = strSql & "LV.dtmDtVencimento <= " & gstrConvDtParaSql(txt_dtmDtBaixa) & " And "
    strSql = strSql & "LV.dblValor > 0 And "
    strSql = strSql & "LP.Intlancamentovalor Is Null And "
    strSql = strSql & "LA.strInscricao = '" & String(gintLenInscricao - Len(Trim(strInscricao)), "0") & UCase(strInscricao) & "' And "
    strSql = strSql & "(LA.INTUTILIZACAO = " & TYP_ACORDO & " OR (LA.INTUTILIZACAO <> " & TYP_ACORDO & " AND LA.bytNaoInscreveda = 0)) "
             
    'Consulta que retorna os acordos
    If Len(strInscricoes) > 0 Then
        
        strInscricoes = Mid(strInscricoes, 1, Len(strInscricoes) - 1)
        
        strSql = strSql & " UNION ALL "
        strSql = strSql & "SELECT LV.Pkid PkidLV, LV.bitParcelaValida, LA.intExercicio, LV.intLancamentoAlfa, LV.intParcela, LV.dblValor ValorOrig, LV.dtmDtVencimento, LV.intMoeda, LV.intLancamentoAlfaAcordo, LV.intLancamentoAlfaDAtiva, " & _
                 "LA.strInscricao, " & gstrCONVERT(CDT_numeric, "LA.strNumeroAviso") & " strNumeroAviso, LA.intComposicaoDaReceita, CR.strSigla, LA.strComposicaoDaReceita, " & strSUBSTRING & "(LAA.strInscricao, " & gintLenInscricao - gintRetornaTamanhoMascara(TYP_ACORDO) + 1 & "," & gintRetornaTamanhoMascara(TYP_ACORDO) - 4 & ")" & strCONCAT & " '/' " & strCONCAT & " " & gstrRIGHT("LAA.strInscricao", 4) & " Acordo, LA.intUtilizacao " & _
                 "FROM " & gstrLancamentoValor & " LV, " & gstrLancamentoAlfa & " LA, " & gstrAcordo & " AC, " & gstrLancamentoAlfa & " LAA, " & gstrComposicaoDaReceita & " CR " & _
                 "WHERE LV.intLancamentoAlfa = LA.pkid AND LV.intLancamentoAlfaAcordo " & strOUTJSQLServer & "= AC.intLancamentoAlfa " & strOUTJOracle
                 strSql = strSql & " AND LV.intLancamentoAlfaAcordo " & strOUTJSQLServer & "= LAA.Pkid " & strOUTJOracle
                 strSql = strSql & " AND LA.intcomposicaodareceita " & strOUTJSQLServer & "= CR.Pkid " & strOUTJOracle
                 strSql = strSql & " AND LV.dtmDtVencimento <= " & gstrConvDtParaSql(txt_dtmDtBaixa)
                 strSql = strSql & " AND LV.Pkid not in(Select Intlancamentovalor From " & gstrLancamentoPagamento & ")" & _
                                   " AND LV.dblValor > 0 AND LA.Pkid IN (" & strInscricoes & ") "
    
    End If

    If bytDBType = EDatabases.Oracle Then
       strSql = strSql & " ORDER BY intLancamentoAlfa, intParcela"
    Else
       strSql = strSql & " ORDER BY LV.intLancamentoAlfa, LV.intParcela"
    End If

    strAtualizacaoDeDebitosSpool = strSql

End Function

Private Function TruncaValores(strValor As String, bytCasasDecimais As Byte) As Double
Dim bytPos   As Byte

    bytPos = (Len(strValor) - InStr(strValor, ",")) - bytCasasDecimais
    
    TruncaValores = Mid(strValor, 1, Len(strValor) - bytPos)
    
End Function


Private Function Gravar()
    Dim strSql As String
    
    strSql = "Update tbllancamentovalor set intlancamentoalfaacordo = Null where Pkid in "
    strSql = strSql & "( "
    strSql = strSql & "Select "
    strSql = strSql & "Lv.Pkid "
    strSql = strSql & "From "
    strSql = strSql & "tbllancamentoalfa LA, "
    strSql = strSql & "tbllancamentovalor LV "
    strSql = strSql & "Where "
    strSql = strSql & "La.Pkid = LV.Intlancamentoalfa And "
    strSql = strSql & "LV.Intlancamentoalfaacordo = (Select Pkid From tbllancamentoalfa LA Where LA.Intutilizacao = 4 and LA.Strinscricao = '00000000001850312000') And "
    strSql = strSql & "LA.Strinscricao = '00000000000001048012' "
    strSql = strSql & ") "
    
    strSql = "Delete From tblacordodebitos where pkid in "
    strSql = strSql & "( "
    strSql = strSql & "Select "
    strSql = strSql & "Ad.Pkid "
    strSql = strSql & "From "
    strSql = strSql & "Tblacordo A, "
    strSql = strSql & "tblacordodebitos AD "
    strSql = strSql & "Where "
    strSql = strSql & "A.Pkid = AD.Intacordo AND "
    strSql = strSql & "A.Intlancamentoalfa = (Select Pkid From tbllancamentoalfa LA Where LA.Strinscricao = '00000000002236212001') AND "
    strSql = strSql & "AD.Strnumeroaviso = '008334' AND "
    strSql = strSql & "AD.Stridentificacao Like '%05012020' "
    strSql = strSql & ") "

    strSql = "Delete From tblcriticabaixa where pkid in ( "
    strSql = strSql & "Select "
    strSql = strSql & "CB.Pkid "
    strSql = strSql & "From "
    strSql = strSql & "Tblmovimentobancario MB, "
    strSql = strSql & "Tblcriticabaixa CB "
    strSql = strSql & "Where "
    strSql = strSql & "MB.Pkid = CB.Intmovimentobancario and "
    strSql = strSql & "MB.Dtmdtmovimento = '10/03/2005'    ) "
        
    
    Set gobjBanco = New clsBanco
    gobjBanco.ExecutaBeginTrans
    gobjBanco.Execute strSql
    gobjBanco.ExecutaRollbackTrans
    gobjBanco.ExecutaCommitTrans
    
End Function

Private Sub txt_strCEP_GotFocus(Index As Integer)
    MarcaCampo txt_strCEP(Index)
End Sub

Private Sub txt_strCEP_KeyPress(Index As Integer, KeyAscii As Integer)
    CaracterValido KeyAscii, "E", txt_strCEP(Index)
End Sub

Private Sub CarregaComboComposicaoReceita()
Dim strSql As String
Dim adoResultado As New ADODB.Recordset

strSql = "SELECT PKID, strDescricao FROM " & gstrComposicaoDaReceita & " ORDER BY strDescricao"

Set gobjBanco = New clsBanco
If gobjBanco.CriaADO(strSql, 10, adoResultado) Then
    If Not adoResultado.EOF Then
        Do While Not adoResultado.EOF
            cboComposicaoReceita.AddItem adoResultado!strDescricao
            cboComposicaoReceita.ItemData(cboComposicaoReceita.NewIndex) = adoResultado!Pkid
            adoResultado.MoveNext
        Loop
    End If
End If

End Sub

Private Sub GeraArquivoCorrigido()
Dim adoLancamentoAlfa As New ADODB.Recordset
Dim strWord           As String
Dim intContador       As Integer

    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO("SELECT * FROM grjspoolbomb40001a50000", 10, adoLancamentoAlfa) Then

            If Not adoLancamentoAlfa.EOF Then
                CmdBanco.ShowSave
                Open CmdBanco.Filename For Output As #1
                pgr_Status.Visible = True
                pgr_Status.Max = Abs(adoLancamentoAlfa.RecordCount)
                Label2.Caption = adoLancamentoAlfa.RecordCount
                
                Do While Not adoLancamentoAlfa.EOF
                                        
                            strWord = ""

                            strWord = adoLancamentoAlfa("lixo1") + adoLancamentoAlfa("areapredio") + adoLancamentoAlfa("lixo2") + adoLancamentoAlfa("inscricao") + adoLancamentoAlfa("lixo3")




                            
                            Print #1, strWord
                            DoEvents
                            pgr_Status.Value = adoLancamentoAlfa.AbsolutePosition
                            Label1.Caption = adoLancamentoAlfa.AbsolutePosition
                            intContador = intContador + 1
                            gobjBanco.ExecutaCommitTrans
                            adoLancamentoAlfa.MoveNext
                        
                        
                    
                   
                Loop
                Close #1
            Else
                ExibeMensagem "Não foram encontrados registros com esses parâmetros."
            End If
    End If
    
    If intContador >= 1 Then
        ExibeMensagem "Arquivo gerado com sucesso com " & intContador & " boleto(s)."
    End If
    
    Screen.MousePointer = vbDefault
    
    pgr_Status.Value = 0
    Exit Sub

End Sub
