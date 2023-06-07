VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Begin VB.Form frmGeracaoSpoolNet 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Geração de Spool para impressão para Internet"
   ClientHeight    =   4620
   ClientLeft      =   1725
   ClientTop       =   3135
   ClientWidth     =   7725
   Icon            =   "frmGeracaoSpoolNet.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4620
   ScaleWidth      =   7725
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab tab_Parametros 
      Height          =   4500
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   7605
      _ExtentX        =   13414
      _ExtentY        =   7938
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Parâmetros"
      TabPicture(0)   =   "frmGeracaoSpoolNet.frx":1042
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "pgr_Status"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "CmdBanco"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "fra_Parametros"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "fra_ComposicaoDaReceita"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      Begin VB.Frame fra_ComposicaoDaReceita 
         Caption         =   "Composição da Receita"
         Height          =   780
         Left            =   1140
         TabIndex        =   1
         Top             =   480
         Width           =   5460
         Begin VB.CommandButton cmd_Composicao 
            Height          =   300
            Left            =   4965
            Picture         =   "frmGeracaoSpoolNet.frx":105E
            Style           =   1  'Graphical
            TabIndex        =   4
            TabStop         =   0   'False
            Tag             =   "590"
            ToolTipText     =   "Ativa Cadastro de Composição da Receita"
            Top             =   315
            Width           =   360
         End
         Begin MSDataListLib.DataCombo dbc_intComposicao 
            Height          =   315
            Left            =   1140
            TabIndex        =   3
            Top             =   315
            Width           =   3795
            _ExtentX        =   6694
            _ExtentY        =   556
            _Version        =   393216
            IntegralHeight  =   0   'False
            Text            =   ""
         End
         Begin VB.Label lbl_Composicao 
            AutoSize        =   -1  'True
            Caption         =   "Composição"
            Height          =   195
            Left            =   165
            TabIndex        =   2
            Top             =   360
            Width           =   870
         End
      End
      Begin VB.Frame fra_Parametros 
         Caption         =   "Faixa de nº de Aviso"
         Height          =   2370
         Left            =   1140
         TabIndex        =   5
         Top             =   1320
         Width           =   5460
         Begin VB.CheckBox chkDebitos 
            Caption         =   "Pesquisar débitos"
            Height          =   315
            Left            =   3000
            TabIndex        =   14
            Top             =   990
            Width           =   1635
         End
         Begin VB.ComboBox cboCodBarras 
            Height          =   315
            ItemData        =   "frmGeracaoSpoolNet.frx":117C
            Left            =   1605
            List            =   "frmGeracaoSpoolNet.frx":117E
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   1350
            Width           =   2985
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
            Left            =   1605
            MaxLength       =   10
            TabIndex        =   13
            Top             =   975
            Width           =   1125
         End
         Begin VB.OptionButton opt_Ordenacao 
            Caption         =   "Ordenação por CEP"
            Height          =   195
            Index           =   1
            Left            =   1920
            TabIndex        =   18
            Top             =   2055
            Width           =   3075
         End
         Begin VB.OptionButton opt_Ordenacao 
            Caption         =   "Ordenação por Identificação"
            Height          =   195
            Index           =   0
            Left            =   1920
            TabIndex        =   17
            Top             =   1785
            Width           =   3075
         End
         Begin VB.TextBox txt_intExercicio 
            Height          =   285
            Left            =   1605
            MaxLength       =   4
            TabIndex        =   11
            Top             =   600
            Width           =   540
         End
         Begin VB.TextBox txt_strAvisoF 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   3360
            MaxLength       =   6
            TabIndex        =   9
            Top             =   240
            Width           =   1200
         End
         Begin VB.TextBox txt_strAvisoI 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1590
            MaxLength       =   6
            TabIndex        =   7
            Top             =   240
            Width           =   1200
         End
         Begin VB.Label lblCodBarras 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Código Barras:"
            Height          =   195
            Left            =   495
            TabIndex        =   15
            Top             =   1440
            Width           =   1035
         End
         Begin VB.Label lblData 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Data base:"
            Height          =   195
            Left            =   765
            TabIndex        =   12
            Top             =   1050
            Width           =   780
         End
         Begin VB.Label lbl_exercicio 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Exercício:"
            Height          =   195
            Left            =   825
            TabIndex        =   10
            Top             =   675
            Width           =   720
         End
         Begin VB.Label lblInicial 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Inicial:"
            Height          =   195
            Left            =   1080
            TabIndex        =   6
            Top             =   300
            Width           =   450
         End
         Begin VB.Label lblFinal 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Final:"
            Height          =   195
            Left            =   2925
            TabIndex        =   8
            Top             =   315
            Width           =   375
         End
      End
      Begin MSComDlg.CommonDialog CmdBanco 
         Left            =   7020
         Top             =   390
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSComctlLib.ProgressBar pgr_Status 
         Height          =   195
         Left            =   1140
         TabIndex        =   19
         Top             =   3750
         Visible         =   0   'False
         Width           =   5460
         _ExtentX        =   9631
         _ExtentY        =   344
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Height          =   225
         Left            =   5505
         TabIndex        =   21
         Top             =   3990
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   1140
         TabIndex        =   20
         Top             =   3960
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmGeracaoSpoolNet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim mobjAux           As Object
    Dim mblnSelecionou    As Boolean
    Dim strWord           As String

Private Sub dbc_intComposicao_GotFocus()
    MarcaCampo dbc_intComposicao
End Sub

Private Sub dbc_intComposicao_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbc_intComposicao, Me, , , Shift
End Sub

Private Sub dbc_intComposicao_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbc_intComposicao
End Sub

Private Sub Form_Activate()
    
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
    
    cboCodBarras.AddItem "Febraban "
    cboCodBarras.ItemData(cboCodBarras.NewIndex) = "0"
    cboCodBarras.AddItem "Ficha Compensação "
    cboCodBarras.ItemData(cboCodBarras.NewIndex) = "1"
    
    dbc_intComposicao.Tag = strQueryComposicao & ";strDescricao"
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar
    mblnSelecionou = False
End Sub

Public Sub MantemForm(ByVal strModoOperacao As String)
    If UCase(strModoOperacao) = gstrSalvar Then
        If blnDadosOk Then
            GeraArquivo
        End If
    ElseIf UCase(strModoOperacao) = gstrPreencherLista Then
        PreencherListaDeOpcoes Me.ActiveControl
    End If
End Sub

Private Function blnDadosOk() As Boolean
    
    blnDadosOk = False
    If Not dbc_intComposicao.MatchedWithList Then
        ExibeMensagem "A composição foi preenchida incorretamente."
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
    ElseIf cboCodBarras.ListIndex < 0 Then
        ExibeMensagem "É necessário selecionar algum tipo de código de barras."
        cboCodBarras.SetFocus
        Exit Function
    End If
    
    blnDadosOk = True
    
End Function

Private Sub GeraArquivo()
    Dim strsql          As String
    Dim strAux          As String
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
    strsql = "SELECT intFebraban FROM " & gstrEmpresa
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strsql, 5, adoResultado) Then
        If gstrENulo(adoResultado!intFebraban) <> "" Then
            intFebraban = gstrENulo(adoResultado!intFebraban)
        Else
            ExibeMensagem "Código Febraban não encontrado."
            GoTo Gravar
        End If
    End If
    adoResultado.Close: Set adoResultado = Nothing
    
    strsql = ""
    strsql = strsql & "Select "
    strsql = strsql & "LA.pkid As IntLancamentoAlfa, LA.strinscricao, LA.strcomposicaodareceita, LA.strocorrencia, LA.strnomeproprietario, LA.strcnpjcpf, LA.stridentidade, LA.strlogradouro, LA.strnumero, LA.strcomplemento, LA.strbairro, LA.strmunicipio, LA.struf, LA.intcep, LA.strlogradouroc, LA.strnumeroc, LA.strcomplementoc, LA.strbairroc, LA.strmunicipioc, LA.strufc, LA.intcepc, LA.strnumeroaviso, LA.strpromissario, LA.stremissao, LA.intexercicio, LA.dtmdtatualizacao, LA.lngcodusr, LA.intcomposicaodareceita, LA.dtmdtcancelamento, LA.bytnaoinscreveda, LA.strindexador, LA.dblvalorcompensacao, LA.intlancamentoalfa, LA.dblvalortotal, LA.dblporcdesconto, LA.dbltotalcomdesconto, LA.dblvlindexador dblIndexadorAlfa, LA.intutilizacao, "
    strsql = strsql & "LI.dblAreaExcedente, LI.dblValorTerrenoExcedente, LI.dblImpostoExcedente, "
    strsql = strsql & "((" & gstrISNULL("LI.dblvalorvenalterreno", "0") & ") /" & gstrISNULL("LI.dblAreaTerreno", "1") & ") as dblValorM2Terreno, "
    strsql = strsql & "(" & gstrISNULL("LI.dblareaterreno", "0") & " + " & gstrISNULL("LI.dblareaexcedente", "0") & ") as dblAreaTotalTerreno, "
    strsql = strsql & "((" & gstrISNULL("LI.dblimpostoterreno", "0") & " + " & gstrISNULL("LI.dblimpostoexcedente", "0") & ") /" & gstrISNULL("LA.dblvlindexador", "1") & ") as dblimpostoterreno, "
    strsql = strsql & "((" & gstrISNULL("LI.dblvalorterrenoexcedente", "0") & " + " & gstrISNULL("LI.dblvalorvenalterreno", "0") & ") /" & gstrISNULL("LA.dblvlindexador", "1") & ") as dblvalorvenalterreno, "
    strsql = strsql & "LI.pkid, LI.intlancamentoalfa, LI.strlote, LI.strquadra, LI.strsequenciadeface, LI.strloteamento, LI.dblaliquotaterreno, LI.dblaliquotaexcedente, LI.intisencao, LI.dtmdtatualizacao, LI.lngcodusr, LI.intlogradouro, LI.dbldesconto, LI.dblvalorreferencia, LI.strindexador, LI.dblvlindexador, "
    strsql = strsql & gstrISNULL("LV.QtdeParcelas", "0") & " As QtdeParcelas, "
    strsql = strsql & "((" & gstrISNULL("LI.dblimpostoterreno", "0") & " + " & gstrISNULL("LI.dblimpostoexcedente", "0") & ") /" & gstrISNULL("LA.dblvlindexador", "1") & ") + (" & gstrISNULL("LPI.Dblimposto", "0") & "/" & gstrISNULL("LA.dblvlindexador", "1") & ") as dblTotTributo, "
    strsql = strsql & "LV.dblTotTributo dblTotTributo1, "
    strsql = strsql & "LV.dtmPrimeiroVenc, PA.INTCONTABANCARIA, "
    strsql = strsql & "(" & gstrISNULL("LPI.Dblvalorvenalpredio", "0") & "/" & gstrISNULL("LA.dblvlindexador", "1") & ") As Dblvalorvenalpredio, "
    strsql = strsql & "Case When " & gstrISNULL("LPI.dblmedidadaarea", "1") & " = 0 Then " & "((" & gstrISNULL("LPI.dblvalorvenalpredio", "0") & ") / 1) Else " & "((" & gstrISNULL("LPI.dblvalorvenalpredio", "0") & ") /" & gstrISNULL("LPI.dblmedidadaarea", "1") & ") End as dblValorM2Predio, "
    strsql = strsql & gstrISNULL("LPI.Dblmedidadaarea", "0") & " As Dblmedidadaarea, "
    strsql = strsql & gstrISNULL("LPI.Dblpontos", "0") & " As Dblpontos, "
    strsql = strsql & gstrISNULL("LPI.Dblvalormetro", "0") & " As Dblvalormetro, "
    strsql = strsql & gstrISNULL("LPI.Dblfatorobsolescencia", "0") & " As Dblfatorobsolescencia, "
    strsql = strsql & gstrISNULL("LPI.Dblaliquota", "0") & " As Dblaliquota, "
    strsql = strsql & gstrISNULL("LPI.Dblimposto", "0") & "/" & gstrISNULL("LA.dblvlindexador", "1") & " As Dblimposto, "
    strsql = strsql & gstrISNULL("LPI.intNPavimento", "0") & " As intNPavimento, "
    strsql = strsql & gstrISNULL("LPI.intAndar", "0") & " As intAndar, "
    strsql = strsql & gstrISNULL("LPI.intElevador", "0") & " As intElevador, "
    strsql = strsql & gstrISNULL("LPI.intGaragem", "0") & " As intGaragem, "
    strsql = strsql & gstrISNULL("LPI.intSuite", "0") & " As intSuite, "
    strsql = strsql & gstrISNULL("LPI.intAmbientes", "0") & " As intAmbientes, "
    strsql = strsql & gstrISNULL("LPI.dblFracaoIdeal", "0") & " As dblFracaoIdeal, "
    strsql = strsql & " LTI.strMedidaDaTestada As TestadaPrincipal, "
    strsql = strsql & "(Select " & gstrISNULL("Ltrim(Rtrim(FQ.strSetor))", "''") & strCONCAT & "' '" & strCONCAT & gstrISNULL("Ltrim(Rtrim(FQ.strQuadra))", "''") & strCONCAT & "' '" & strCONCAT & gstrISNULL("LO.strCodigo", "''") & " FROM " & gstrFaceDeQuadra & " FQ, " & gstrLogradouro & " LO Where FQ.Pkid = LTI.intFacedequadra and LO.Pkid " & strOUTJOracle & " =" & strOUTJSQLServer & " FQ.intLogradouro) As CodFaceDeQuadra,"
    'Depois colocar no padrao com gstrConvert, mas é preciso criar a constante para money
    strsql = strsql & "(Select " & IIf(bytDBType = SQLServer, "Sum(CONVERT(money, LTI2.strMedidaDaTestada))", "Sum(LTI2.strMedidaDaTestada)") & " FROM tblLancamentoTestadasIPTU LTI2 WHERE LTI2.intLancamentoIPTU = LI.Pkid and LTI2.intTipoDeTestada = 3) As DemaisTestadas, "
    
    'strSql = strSql & gstrTOPnOracle(" Select " & gstrTOPnSQLServer(1) & " LPI.strNomeUso From " & gstrLancamentoPredioIPTU & " LPI Where LI.Pkid = LPI.Intlancamentoiptu Order By LPI.Pkid ", 1) & ")  strCategoria, ("
    'strSql = strSql & gstrTOPnOracle(" Select " & gstrTOPnSQLServer(1) & " LPI.dtmUltimaReforma From " & gstrLancamentoPredioIPTU & " LPI Where LI.Pkid = LPI.Intlancamentoiptu Order By LPI.Pkid ", 1) & ") dtmUltimaReforma, "
    
    strsql = strsql & "(" & gstrISNULL("LI.dblvalorterrenoexcedente", "0") & " + " & gstrISNULL("LPI.Dblvalorvenalpredio", "0") & " + " & gstrISNULL("LI.Dblvalorvenalterreno", "0") & ") / " & gstrISNULL("LA.dblvlindexador", "1") & " as dblvalorVenalTotal "
    strsql = strsql & "From "
    strsql = strsql & gstrLancamentoAlfa & " LA, "
    strsql = strsql & gstrLancamentoIPTU & " LI, "
    strsql = strsql & gstrParametroAtualizacao & " PA, "
    strsql = strsql & " tblLancamentoTestadasIPTU LTI, "
    strsql = strsql & "( "
    strsql = strsql & "Select LPI.Intlancamentoiptu, "
    strsql = strsql & "Sum(" & gstrISNULL("LPI.Dblvalorvenalpredio", "0") & ") As Dblvalorvenalpredio, "
    strsql = strsql & "Sum(" & gstrISNULL("LPI.Dblmedidadaarea", "0") & ") As Dblmedidadaarea, "
    strsql = strsql & "Sum(" & gstrISNULL("LPI.Dblpontos", "0") & ") As Dblpontos, "
    strsql = strsql & "Sum(" & gstrISNULL("LPI.Dblvalormetro", "0") & ") As Dblvalormetro, "
    strsql = strsql & "Sum(" & gstrISNULL("LPI.Dblfatorobsolescencia", "0") & ") As Dblfatorobsolescencia, "
    strsql = strsql & "Sum(" & gstrISNULL("LPI.Dblaliquota", "0") & ") As Dblaliquota, "
    strsql = strsql & "Sum(" & gstrISNULL("LPI.Dblimposto", "0") & ") As Dblimposto, "
    strsql = strsql & "Sum(" & gstrISNULL("LPI.intNPavimento", "0") & ") As intNPavimento, "
    strsql = strsql & "Sum(" & gstrISNULL("LPI.intAndar", "0") & ") As intAndar, "
    strsql = strsql & "Sum(" & gstrISNULL("LPI.intElevador", "0") & ") As intElevador, "
    strsql = strsql & "Sum(" & gstrISNULL("LPI.intGaragem", "0") & ") As intGaragem, "
    strsql = strsql & "Sum(" & gstrISNULL("LPI.intSuite", "0") & ") As intSuite, "
    strsql = strsql & "Sum(" & gstrISNULL("LPI.intQuarto", "0") & " + " & gstrISNULL("LPI.intSala", "0") & " + " & gstrISNULL("LPI.intCozinha", "0") & " + " & gstrISNULL("LPI.intBanheiro", "0") & ") As intAmbientes, "
    strsql = strsql & "Sum(" & gstrISNULL("LPI.dblFracaoIdeal", "0") & ") As dblFracaoIdeal "
    strsql = strsql & "From " & gstrLancamentoAlfa & " La," & gstrLancamentoIPTU & " LI," & gstrLancamentoPredioIPTU & " LPI "
    strsql = strsql & "Where "
    strsql = strsql & "LA.Pkid = Li.Intlancamentoalfa  And "
    strsql = strsql & "LI.Pkid = LPI.Intlancamentoiptu And "
    strsql = strsql & "LA.intExercicio = " & txt_intExercicio & " And "
    strsql = strsql & gstrCONVERT(CDT_numeric, "LA.strNumeroAviso") & " Between " & Val(txt_strAvisoI.Text) & " And " & Val(txt_strAvisoF.Text) & " "
    strsql = strsql & "Group By LPI.Intlancamentoiptu "
    strsql = strsql & ") LPI, "
    strsql = strsql & "( "
    strsql = strsql & "Select count(LV.Intparcela) As QtdeParcelas, LV.Intlancamentoalfa, Sum(lv.dblvalor) dblTotTributo, Min(lv.dtmdtvencimento) dtmPrimeiroVenc "
    strsql = strsql & "From " & gstrLancamentoAlfa & " La," & gstrLancamentoValor & " LV "
    strsql = strsql & "Where LA.Pkid = LV.Intlancamentoalfa And "
    strsql = strsql & "LA.intExercicio = " & txt_intExercicio & " And "
    strsql = strsql & "LV.Bitparcelavalida = 1 And "
    strsql = strsql & gstrCONVERT(CDT_numeric, "LA.strNumeroAviso") & " Between " & Val(txt_strAvisoI.Text) & " And " & Val(txt_strAvisoF.Text) & " "
    strsql = strsql & "Group By LV.Intlancamentoalfa "
    strsql = strsql & ") LV "
    strsql = strsql & "Where "
    strsql = strsql & "LA.Pkid = Li.Intlancamentoalfa And "
    strsql = strsql & "LI.Pkid " & strOUTJSQLServer & "= LPI.Intlancamentoiptu" & strOUTJOracle & " And "
    strsql = strsql & "LI.Pkid " & strOUTJSQLServer & "= LTI.Intlancamentoiptu" & strOUTJOracle & " And "
    strsql = strsql & "LTI.intTipoDeTestada = 2 " & " And "
    strsql = strsql & "LA.Pkid " & strOUTJSQLServer & "= LV.IntlancamentoAlfa" & strOUTJOracle & " And "
    strsql = strsql & "LA.intExercicio = " & txt_intExercicio & " And "
    strsql = strsql & "LA.dtmdtcancelamento is null" & " And "
    strsql = strsql & gstrCONVERT(CDT_numeric, "LA.strNumeroAviso") & " Between " & Val(txt_strAvisoI.Text) & " And " & Val(txt_strAvisoF.Text) & " "
    strsql = strsql & " AND PA.intComposicaoReceita = LA.intComposicaoDaReceita "
    strsql = strsql & " AND LA.intComposicaoDaReceita = " & Val(dbc_intComposicao.BoundText)
    strsql = strsql & " AND PA.intExercicio = LA.intExercicio "

    If opt_Ordenacao(0).Value = True Then
        strsql = strsql & " Order By La.strNumeroAviso"
    Else
        strsql = strsql & " Order By La.intcepc,la.strinscricao"
    End If
    
    Dim A As Integer
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strsql, 40, adoResultado) Then
        With adoResultado
            If Not .EOF Then
                Open CmdBanco.Filename For Output As #1
                pgr_Status.Visible = True
                pgr_Status.Max = Abs(.RecordCount)
                Label2.Caption = adoResultado.RecordCount
                Do While Not .EOF
                    
                    'Vamos verificar se a conta bancaria esta definida para esta composicao
'                    If IsNull(!intContaBancaria) Then
'                        ExibeMensagem "Não foi encontrada Conta Bancária para a Composição " & !strComposicaoDaReceita & " no Exercício de " & !intExercicio & "."
'                        GoTo Gravar
'                    End If
                    
                    If CDbl(gstrConvVrDoSql(gstrENulo(!dblTotTributo1), , , True)) > 0 Then
                        strWord = ""
                        'Código do tributo
                        strWord = strWord & Space$(3 - Len("10")) & "10"
                        'Descrição do Tributo
                        strWord = strWord & Left(gstrENulo(!strComposicaoDaReceita), 40) & Space$(40 - Len(Left(gstrENulo(!strComposicaoDaReceita), 40)))
                        'Exercício
                        strWord = strWord & gstrENulo(!intExercicio) & Space$(4 - Len(gstrENulo(!intExercicio)))
                        'Nº do Aviso
                        strWord = strWord & Format$(gstrENulo(!strNumeroAviso), "000,000")
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
                        strAux = Trim(gstrENulo(!strlogradouroc)) & " " & Trim(gstrENulo(!strNumeroC)) & " " & Trim(gstrENulo(!strComplementoC)) & " " & Trim(gstrENulo(!strBairroC))
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
                        strsql = ""
                        strsql = strsql & "Select "
                        strsql = strsql & "LV.Pkid, LV.Intlancamentoalfa, "
                        strsql = strsql & "LV.INTPARCELA, "
                        strsql = strsql & "LV.Dtmdtvencimento, "
                        strsql = strsql & "LV.Dblvalor, "
                        strsql = strsql & "LV.Bitparcelavalida, "
                        strsql = strsql & "G.Strcodbarra "
                        strsql = strsql & "From "
                        strsql = strsql & gstrLancamentoValor & " LV, "
                        strsql = strsql & gstrLancamentoGuias & " LG, "
                        strsql = strsql & gstrGuias & " G, "
                        strsql = strsql & "(Select LV.Pkid, Min(G.Pkid) PkidG "
                        strsql = strsql & "From tbllancamentovalor LV, tbllancamentoguias LG, tblGuias G "
                        strsql = strsql & "Where LV.PKID = LG.Intlancamentovalor "
                        strsql = strsql & "AND LG.Intguias = G.Pkid "
                        strsql = strsql & "AND LV.Intlancamentoalfa = " & !intLancamentoAlfa & " "
                        strsql = strsql & "Group By LV.Pkid) G1 "
                        strsql = strsql & "Where "
                        strsql = strsql & "LV.pkid = LG.intLancamentoValor AND "
                        strsql = strsql & "G.Pkid = LG.intGuias AND "
                        strsql = strsql & "G.pkid = G1.pkidG AND "
                        strsql = strsql & "LV.Intlancamentoalfa = " & !intLancamentoAlfa & " "
                        strsql = strsql & "Group By "
                        strsql = strsql & "LV.PKID, "
                        strsql = strsql & "LV.INTLANCAMENTOALFA, "
                        strsql = strsql & "LV.INTPARCELA, "
                        strsql = strsql & "LV.DTMDTVENCIMENTO, "
                        strsql = strsql & "LV.DBLVALOR, "
                        strsql = strsql & "LV.BITPARCELAVALIDA, "
                        strsql = strsql & "G.Strcodbarra "
                        strsql = strsql & " Order By "
                        strsql = strsql & "LV.Intlancamentoalfa, "
                        strsql = strsql & "LV.Bitparcelavalida, "
                        strsql = strsql & "LV.Intparcela "
                        
                        Set gobjBanco = New clsBanco
                        If gobjBanco.CriaADO(strsql, 10, adoAux) Then
                            If Not .EOF Then
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
                                       
                                        'Vamos definir o codigo de barras
                                        strCodBarras = adoAux!Strcodbarra
                                        
                                        If Len(strCodBarras) = 0 Then
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
                        
                                                strsql = gstrStoredProcedure("sp_AtualizaParcela", adoAtualizacao!intComposicaoDaReceita & ", " & adoAtualizacao!intExercicio & ", " & adoAtualizacao!intParcela & ", " & gstrConvDtParaSql(adoAtualizacao!Dtmdtvencimento) & ", " & gstrConvDtParaSql(txt_dtmDtBaixa) & ", " & gstrConvVrParaSql(adoAtualizacao!ValorOrig) & ", " & adoAtualizacao!intMoeda, True)
                                                
                                                Set gobjBanco = New clsBanco
                                                If gobjBanco.CriaADO(strsql, 80, adoParcelas) Then
                                                
                                                    If intCont < 14 Then
                                                        If adoAtualizacao("PkidLV").Value > 0 Then
                                                            If adoAtualizacao("bitParcelaValida").Value = 1 Then
                                                                If Not blnPrimeiraVez Then
                                                                
                                                                    dblTotal = dblTotal + CDbl(gstrConvVrDoSql(CCur(gstrConvVrDoSql(adoParcelas("dblValorPrincipal").Value)) + CCur(gstrConvVrDoSql(adoParcelas("dblValorMulta").Value)) + CCur(gstrConvVrDoSql(adoParcelas("dblValorJuros").Value)) + CCur(gstrConvVrDoSql(adoParcelas("dblValorCorrecao").Value)), 4, , True))
                                                                    strsigla = Mid(Trim(gstrENulo(adoAtualizacao("strSigla").Value)), 1, 4)
                                                                    intExercicio = adoAtualizacao("intExercicio").Value
                                                                    strAviso = adoAtualizacao("strNumeroAviso").Value
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
                                                                        strAviso = adoAtualizacao("strNumeroAviso").Value
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
                        
                        
                        strsql = " Select " & gstrTOPnSQLServer(1) & " LPI.dtmUltimaReforma From " & gstrLancamentoPredioIPTU & " LPI Where " & Val(gstrENulo(!Pkid)) & " = LPI.Intlancamentoiptu Order By LPI.Pkid "
                        strsql = gstrTOPnOracle(strsql, 1)
                        
                        Set gobjBanco = New clsBanco
                        If gobjBanco.CriaADO(strsql, 10, adoAux) Then
                            If Not .EOF Then
                                'Ano ultima reforma
                                If IsDate(adoAux!dtmUltimaReforma) Then
                                    strWord = strWord & Left(gstrENulo(Year(adoAux!dtmUltimaReforma)), 4) & Space$(4 - Len(Left(gstrENulo(Year(adoAux!dtmUltimaReforma)), 4)))
                                    strDtmUltimaReforma = Left(gstrENulo(Year(adoAux!dtmUltimaReforma)), 4) & Space$(4 - Len(Left(gstrENulo(Year(adoAux!dtmUltimaReforma)), 4)))
                                Else
                                    strWord = strWord & "0000"
                                    strDtmUltimaReforma = "0000"
                                End If
                            Else
                                strWord = strWord & "0000"
                            End If
                        End If
                        
                        strsql = " Select " & gstrTOPnSQLServer(1) & " LPI.strNomeUso From " & gstrLancamentoPredioIPTU & " LPI Where " & Val(gstrENulo(!Pkid)) & " = LPI.Intlancamentoiptu Order By LPI.Pkid "
                        strsql = gstrTOPnOracle(strsql, 1)
                        
                        Set gobjBanco = New clsBanco
                        If gobjBanco.CriaADO(strsql, 10, adoAux) Then
                            If Not .EOF Then
                                'Categoria
                                strWord = strWord & Left(gstrENulo(adoAux!strNomeUso), 25) & Space$(25 - Len(Left(gstrENulo(adoAux!strNomeUso), 25)))
                            Else
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
                        strInscricao = gstrFormataInscricao(Right(gstrENulo(!strInscricao), gintRetornaTamanhoMascara(TYP_IMOBILIARIA)))
                        strWord = strWord & Left(gstrENulo(strInscricao), 20) & Space$(20 - Len(Left(gstrENulo(strInscricao), 20)))
                        'Cod Face de quadra
                        strWord = strWord & Left(gstrENulo(!CodFaceDeQuadra), 20) & Space$(20 - Len(Left(gstrENulo(!CodFaceDeQuadra), 20)))
                        'Testada principal
                        strWord = strWord & Space$(15 - Len(Left(gstrConvVrDoSql(!TestadaPrincipal), 15))) & Left(gstrConvVrDoSql(!TestadaPrincipal), 15)
                        'Demais testadas
                        strWord = strWord & Space$(15 - Len(Left(gstrConvVrDoSql(!DemaisTestadas), 15))) & Left(gstrConvVrDoSql(!DemaisTestadas), 15)
                        
                        'Vamos preencher as colunas de taxas
                        strsql = ""
                        strsql = strsql & "Select "
                        strsql = strsql & "SUM(LR.dblValor) dblValor, R.strDescricao "
                        strsql = strsql & "From "
                        strsql = strsql & gstrLancamentoValor & " LV, "
                        strsql = strsql & gstrLancamentoReceita & " LR, "
                        strsql = strsql & gstrReceita & " R "
                        strsql = strsql & "Where "
                        strsql = strsql & "LV.Intlancamentoalfa = " & !intLancamentoAlfa & " AND "
                        strsql = strsql & "LR.intLancamentoValor = LV.Pkid AND "
                        strsql = strsql & "R.Pkid = LR.intReceita AND "
                        strsql = strsql & "LV.bitParcelaValida = 1 AND "
                        strsql = strsql & "R.bytTipo = 3 " 'Tipo Taxa
                        strsql = strsql & "Group By R.strDescricao "
                        
                        Set gobjBanco = New clsBanco
                        If gobjBanco.CriaADO(strsql, 10, adoAux) Then
                                
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
                        
                        Print #1, strWord
                        DoEvents
                        pgr_Status.Value = .AbsolutePosition
                        Label1.Caption = .AbsolutePosition
                        intContador = intContador + 1
                        .MoveNext
                    Else
                        DoEvents
                        pgr_Status.Value = .AbsolutePosition
                        Label1.Caption = .AbsolutePosition
                        .MoveNext
                    End If
                    
                    intCont = 0
                    
                Loop
                Close #1
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
        ExibeMensagem "Arquivo gerado com sucesso com " & intContador & " boleto(s)."
    End If
    
    pgr_Status.Value = 0
    Exit Sub
    
Gravar:
    If Len(Err.Description) > 0 Then MsgBox Err.Description
    
    Close #1
    Screen.MousePointer = vbDefault
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

Private Sub txt_intExercicio_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txt_intExercicio
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
Dim strsql                  As String
Dim strAcordosParaConsulta  As String
Dim strInscricoes           As String
    
    'Vamos obter os Pkids das inscricoes para fazer consulta de acordos
    strsql = "SELECT  LA.Pkid " & _
             "FROM " & gstrLancamentoAlfa & " LA " & _
             "WHERE LA.strInscricao = '" & String(gintLenInscricao - Len(Trim(strInscricao)), "0") & UCase(strInscricao) & "' AND " & _
             "(LA.INTUTILIZACAO = " & TYP_ACORDO & " OR (LA.INTUTILIZACAO <> " & TYP_ACORDO & " AND LA.bytNaoInscreveda = 0)) "
    
    If gobjBanco.CriaADO(strsql, 5, adoResultado) Then
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
    strsql = "SELECT  LV.intLancamentoAlfaAcordo " & _
             "FROM " & gstrLancamentoValor & " LV, " & gstrLancamentoAlfa & " LA " & _
             "WHERE LV.intLancamentoAlfa = LA.pkid AND " & _
             "LA.Pkid IN (" & strAcordosParaConsulta & ") AND Not LV.intLancamentoAlfaAcordo Is Null " & _
             "GROUP BY LV.intLancamentoAlfaAcordo "
    
    If gobjBanco.CriaADO(strsql, 5, adoResultado) Then
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
    
    strsql = "SELECT LV.Pkid PkidLV, LV.bitParcelaValida, LA.intExercicio, LV.intLancamentoAlfa, LV.intParcela, "
    strsql = strsql & "LV.dblValor ValorOrig, LV.dtmDtVencimento, LV.intMoeda, LV.intLancamentoAlfaAcordo, LV.intLancamentoAlfaDAtiva, "
    strsql = strsql & "LA.strInscricao, " & gstrCONVERT(CDT_numeric, "LA.strNumeroAviso") & " strNumeroAviso, "
    strsql = strsql & "LA.intComposicaoDaReceita, CR.strSigla, LA.strComposicaoDaReceita, " & strSUBSTRING & "(LAA.strInscricao, " & gintLenInscricao - gintRetornaTamanhoMascara(TYP_ACORDO) + 1 & "," & gintRetornaTamanhoMascara(TYP_ACORDO) - 4 & ")" & strCONCAT & " '/' " & strCONCAT & " " & gstrRIGHT("LAA.strInscricao", 4) & " Acordo, LA.intUtilizacao "
    strsql = strsql & "FROM " & gstrLancamentoValor & " LV, "
    strsql = strsql & gstrLancamentoAlfa & " LA, "
    strsql = strsql & gstrAcordo & " AC, "
    strsql = strsql & gstrLancamentoAlfa & " LAA, "
    strsql = strsql & gstrLancamentoPagamento & " LP, "
    strsql = strsql & gstrComposicaoDaReceita & " CR "
    strsql = strsql & "WHERE LV.intLancamentoAlfa = LA.pkid AND "
    strsql = strsql & "LV.intLancamentoAlfaAcordo " & strOUTJSQLServer & "= AC.intLancamentoAlfa " & strOUTJOracle & " And "
    strsql = strsql & "LA.intcomposicaodareceita " & strOUTJSQLServer & "= CR.Pkid " & strOUTJOracle & " And "
    strsql = strsql & "LV.intLancamentoAlfaAcordo " & strOUTJSQLServer & "= LAA.Pkid " & strOUTJOracle & " And "
    strsql = strsql & "LV.Pkid" & strOUTJSQLServer & "= LP.Intlancamentovalor " & strOUTJOracle & " And "
    strsql = strsql & "LV.dtmDtVencimento <= " & gstrConvDtParaSql(txt_dtmDtBaixa) & " And "
    strsql = strsql & "LV.dblValor > 0 And "
    strsql = strsql & "LP.Intlancamentovalor Is Null And "
    strsql = strsql & "LA.strInscricao = '" & String(gintLenInscricao - Len(Trim(strInscricao)), "0") & UCase(strInscricao) & "' And "
    strsql = strsql & "(LA.INTUTILIZACAO = " & TYP_ACORDO & " OR (LA.INTUTILIZACAO <> " & TYP_ACORDO & " AND LA.bytNaoInscreveda = 0)) "
             
    'Consulta que retorna os acordos
    If Len(strInscricoes) > 0 Then
        
        strInscricoes = Mid(strInscricoes, 1, Len(strInscricoes) - 1)
        
        strsql = strsql & " UNION ALL "
        strsql = strsql & "SELECT LV.Pkid PkidLV, LV.bitParcelaValida, LA.intExercicio, LV.intLancamentoAlfa, LV.intParcela, LV.dblValor ValorOrig, LV.dtmDtVencimento, LV.intMoeda, LV.intLancamentoAlfaAcordo, LV.intLancamentoAlfaDAtiva, " & _
                 "LA.strInscricao, " & gstrCONVERT(CDT_numeric, "LA.strNumeroAviso") & " strNumeroAviso, LA.intComposicaoDaReceita, CR.strSigla, LA.strComposicaoDaReceita, " & strSUBSTRING & "(LAA.strInscricao, " & gintLenInscricao - gintRetornaTamanhoMascara(TYP_ACORDO) + 1 & "," & gintRetornaTamanhoMascara(TYP_ACORDO) - 4 & ")" & strCONCAT & " '/' " & strCONCAT & " " & gstrRIGHT("LAA.strInscricao", 4) & " Acordo, LA.intUtilizacao " & _
                 "FROM " & gstrLancamentoValor & " LV, " & gstrLancamentoAlfa & " LA, " & gstrAcordo & " AC, " & gstrLancamentoAlfa & " LAA, " & gstrComposicaoDaReceita & " CR " & _
                 "WHERE LV.intLancamentoAlfa = LA.pkid AND LV.intLancamentoAlfaAcordo " & strOUTJSQLServer & "= AC.intLancamentoAlfa " & strOUTJOracle
                 strsql = strsql & " AND LV.intLancamentoAlfaAcordo " & strOUTJSQLServer & "= LAA.Pkid " & strOUTJOracle
                 strsql = strsql & " AND LA.intcomposicaodareceita " & strOUTJSQLServer & "= CR.Pkid " & strOUTJOracle
                 strsql = strsql & " AND LV.dtmDtVencimento <= " & gstrConvDtParaSql(txt_dtmDtBaixa)
                 strsql = strsql & " AND LV.Pkid not in(Select Intlancamentovalor From " & gstrLancamentoPagamento & ")" & _
                                   " AND LV.dblValor > 0 AND LA.Pkid IN (" & strInscricoes & ") "
    
    End If

    If bytDBType = EDatabases.Oracle Then
       strsql = strsql & " ORDER BY intLancamentoAlfa, intParcela"
    Else
       strsql = strsql & " ORDER BY LV.intLancamentoAlfa, LV.intParcela"
    End If

    strAtualizacaoDeDebitosSpool = strsql

End Function

Private Function TruncaValores(strValor As String, bytCasasDecimais As Byte) As Double
Dim bytPos   As Byte

    bytPos = (Len(strValor) - InStr(strValor, ",")) - bytCasasDecimais
    
    TruncaValores = Mid(strValor, 1, Len(strValor) - bytPos)
    
End Function

Private Function strQueryComposicao() As String
    Dim strsql As String

    strsql = "SELECT Pkid,"
    strsql = strsql & gstrCONVERT(CDT_VARCHAR, "intCodigo") & strCONCAT & "' - '" & strCONCAT & " strDescricao Descricao"
    strsql = strsql & " FROM "
    strsql = strsql & gstrComposicaoDaReceita
    strsql = strsql & " WHERE"
    strsql = strsql & " intUtilizacao in (1) "
    strsql = strsql & " ORDER BY intCodigo"

    strQueryComposicao = strsql

End Function

