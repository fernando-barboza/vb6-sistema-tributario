VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MsDatLst.ocx"
Begin VB.Form frmRelMovimentoBancario 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Movimento Bancário"
   ClientHeight    =   3435
   ClientLeft      =   4560
   ClientTop       =   4005
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3435
   ScaleWidth      =   6585
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab tab_3dPasta 
      Height          =   3285
      Left            =   60
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   60
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   5794
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Movimento"
      TabPicture(0)   =   "frmRelMovimentoBancario.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblLote"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblBanco"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblAgencia"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblContaBancaria"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblDtMovimento"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "dbc_intLote"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "dbcintContaBancaria"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txt_strAgencia"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txt_strBanco"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "chk_TodosLotes"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "chk_TodosContas"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txt_DataMovimento"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).ControlCount=   12
      Begin VB.TextBox txt_DataMovimento 
         Height          =   285
         Left            =   1680
         MaxLength       =   11
         TabIndex        =   5
         Top             =   540
         Width           =   1065
      End
      Begin VB.CheckBox chk_TodosContas 
         Caption         =   "Selecionar todas as Contas"
         Height          =   195
         Left            =   1680
         TabIndex        =   4
         Top             =   1290
         Width           =   2865
      End
      Begin VB.CheckBox chk_TodosLotes 
         Caption         =   "Selecionar todos os Lotes"
         Height          =   195
         Left            =   1680
         TabIndex        =   3
         Top             =   2790
         Width           =   2865
      End
      Begin VB.TextBox txt_strBanco 
         Height          =   315
         Left            =   1680
         TabIndex        =   2
         Top             =   1995
         Width           =   2655
      End
      Begin VB.TextBox txt_strAgencia 
         Height          =   315
         Left            =   1680
         TabIndex        =   1
         Top             =   1560
         Width           =   1515
      End
      Begin MSDataListLib.DataCombo dbcintContaBancaria 
         Height          =   315
         Left            =   1680
         TabIndex        =   6
         Top             =   930
         Width           =   3045
         _ExtentX        =   5371
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dbc_intLote 
         Height          =   315
         Left            =   1680
         TabIndex        =   7
         Top             =   2400
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin VB.Label lblDtMovimento 
         AutoSize        =   -1  'True
         Caption         =   "Data do Movimento"
         Height          =   195
         Left            =   210
         TabIndex        =   12
         Top             =   600
         Width           =   1395
      End
      Begin VB.Label lblContaBancaria 
         AutoSize        =   -1  'True
         Caption         =   "Conta Bancária"
         Height          =   195
         Left            =   510
         TabIndex        =   11
         Top             =   1020
         Width           =   1095
      End
      Begin VB.Label lblAgencia 
         AutoSize        =   -1  'True
         Caption         =   "Agência"
         Height          =   195
         Left            =   1020
         TabIndex        =   10
         Top             =   1680
         Width           =   585
      End
      Begin VB.Label lblBanco 
         AutoSize        =   -1  'True
         Caption         =   "Banco"
         Height          =   195
         Left            =   1140
         TabIndex        =   9
         Top             =   2085
         Width           =   465
      End
      Begin VB.Label lblLote 
         AutoSize        =   -1  'True
         Caption         =   "Lote"
         Height          =   195
         Left            =   1290
         TabIndex        =   8
         Top             =   2430
         Width           =   315
      End
   End
End
Attribute VB_Name = "frmRelMovimentoBancario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim mblnAlterando    As Boolean
    Dim mobjAux          As Object
    Dim mblnSelecionou   As Boolean
    Dim mblnPrimeiraVez  As Boolean

Private Sub chk_TodosContas_Click()
    If chk_TodosContas.Value Then
        TrocaCorObjeto dbcintContaBancaria, True
        TrocaCorObjeto dbc_intLote, True
        chk_TodosLotes.Value = 1
        chk_TodosLotes.Enabled = False
        txt_strAgencia.Text = ""
        txt_strBanco.Text = ""
        dbcintContaBancaria.Text = ""
        dbc_intLote.Text = ""
    Else
        TrocaCorObjeto dbcintContaBancaria, False
        TrocaCorObjeto dbc_intLote, False
        chk_TodosLotes.Value = 0
        chk_TodosLotes.Enabled = True
    End If
    
End Sub

Private Sub chk_TodosLotes_Click()
    If chk_TodosLotes.Value Then
        TrocaCorObjeto dbc_intLote, True
        chk_TodosLotes.Value = 1
        dbc_intLote.Text = ""
    Else
        TrocaCorObjeto dbc_intLote, False
        chk_TodosLotes.Value = 0
    End If
End Sub

Private Sub dbcintcontabancaria_Change()
    If dbcintContaBancaria.MatchedWithList Then
        PreencheAgBanco (dbcintContaBancaria.BoundText)
    End If
End Sub

Private Sub dbcintContaBancaria_GotFocus()
    MarcaCampo dbcintContaBancaria
    dbcintContaBancaria.Tag = strQueryContaCorrente & ";strConta"
End Sub

Private Sub dbc_intLote_GotFocus()
    MarcaCampo dbc_intLote
    dbc_intLote.Tag = strQueryLotes & ";intLote"
End Sub

Private Sub Form_Load()
    HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrSalvar, gstrDeletar
    TrocaCorObjeto txt_strAgencia, True
    TrocaCorObjeto txt_strBanco, True
    
End Sub
Private Sub Form_Activate()
    gintCodSeguranca = 1418
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

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar
    mblnSelecionou = False
    mblnPrimeiraVez = False
End Sub

Public Sub MantemForm(ByVal strModoOperacao As String)
On Error Resume Next
    If UCase(strModoOperacao) = UCase(gstrImprimir) Then
        If blnDadosOk Then
            ImprimeRelatorio rptMovimentoBancario, strQueryRelatorio
        End If
    End If
    
    If UCase(strModoOperacao) = UCase(gstrNovo) Then
        Limpa_Controles frmRelMovimentoBancario, True, False, True, True, False
        
        dbcintContaBancaria.ListField = ""
        Set dbc_intLote.DataSource = Nothing
        Set dbc_intLote.RowSource = Nothing
    End If
    
    If UCase(strModoOperacao) = UCase(gstrFechar) Then
        Unload Me
    End If
    If UCase(strModoOperacao) = UCase(gstrPreencherLista) Then
        If Me.ActiveControl.Name = "dbcintContaBancaria" Then
            If Trim(txt_DataMovimento) = "" Then
                Exit Sub
            End If
        End If
        If Me.ActiveControl.Name = "dbc_intLote" Then
            If dbcintContaBancaria.MatchedWithList = False Then
                Exit Sub
            End If
        End If
        PreencherListaDeOpcoes Me.ActiveControl
    End If
    
End Sub

Private Function strQueryRelatorio() As String
Dim strsql As String
    
    strsql = ""
    strsql = strsql & "SELECT "
    strsql = strsql & "MB.Dtmdtmovimento, "
    strsql = strsql & "B.STRDESCRICAO, "
    If bytDBType = Oracle Then
        strsql = strsql & "Trim(" & gstrCONVERT(CDT_VARCHAR, "CB.strConta") & ")" & strCONCAT & "'-'" & strCONCAT & " Trim(CB.strDigitoVerificador) ContaCorrente, "
    Else
        strsql = strsql & "LTrim(RTrim(" & gstrCONVERT(CDT_VARCHAR, "CB.strConta") & "))" & strCONCAT & "'-'" & strCONCAT & " LTrim(RTrim(CB.strDigitoVerificador)) ContaCorrente, "
    End If
    strsql = strsql & "A.STRDESCRICAO AS strAgencia, "
    strsql = strsql & "MB.INTLOTE, "
    strsql = strsql & "LA.Strinscricao, "
    strsql = strsql & "LA.Strcomposicaodareceita, "
    strsql = strsql & "LA.Intexercicio, "
    strsql = strsql & "LA.strNumeroAviso , "
    strsql = strsql & "LA.intUtilizacao, "
    strsql = strsql & "MB.Dtmdtpagamento, "
    strsql = strsql & "CRB.INTTIPOCRITICA, "
    strsql = strsql & "LV.intParcela, "
    strsql = strsql & "MB.strCodigoDeBarras, "
    strsql = strsql & "(" & gstrISNULL("MB.Dblprincipal", "0") & " + " & gstrISNULL("MB.Dblmulta", "0") & " + " & _
        gstrISNULL("MB.Dbljuros", "0") & " + " & gstrISNULL("MB.Dblcorrecao", "0") & ") AS DblValor "
    If bytDBType = SQLServer Then
        strsql = strsql & " FROM tblCriticaBaixa CRB INNER JOIN "
        strsql = strsql & gstrMovimentoBancario & " MB ON CRB.INTMOVIMENTOBANCARIO = MB.Pkid INNER JOIN "
        strsql = strsql & gstrContaBancaria & " CB ON MB.intContaBancaria = CB.PKId INNER JOIN "
        strsql = strsql & gstrBanco & " B ON CB.intBanco = B.PKId INNER JOIN "
        strsql = strsql & gstrAgencia & " A ON CB.intAgencia = A.PKId INNER JOIN "
        strsql = strsql & gstrLancamentoValor & " LV ON MB.intlancamentovalor = LV.PKId LEFT OUTER JOIN "
        strsql = strsql & gstrLancamentoAlfa & " LA ON LV.intLancamentoAlfa = LA.PKId "
        strsql = strsql & " WHERE MB.Dtmdtmovimento = " & gstrConvDtParaSql(txt_DataMovimento.Text)
    Else
        strsql = strsql & " FROM "
        strsql = strsql & gstrCriticaBaixa & " CRB, "
        strsql = strsql & gstrMovimentoBancario & " MB, "
        strsql = strsql & gstrLancamentoValor & " LV, "
        strsql = strsql & gstrLancamentoAlfa & " LA, "
        strsql = strsql & gstrContaBancaria & " CB, "
        strsql = strsql & gstrBanco & " B, "
        strsql = strsql & gstrAgencia & " A "
        strsql = strsql & " WHERE "
        strsql = strsql & "MB.Pkid = CRB.intMovimentoBancario AND "
        strsql = strsql & "LV.Pkid " & strOUTJOracle & " =" & strOUTJSQLServer & " MB.Intlancamentovalor AND "
        strsql = strsql & "LA.Pkid " & strOUTJOracle & " =" & strOUTJSQLServer & " LV.Intlancamentoalfa AND "
        strsql = strsql & "CB.Pkid = MB.Intcontabancaria AND "
        strsql = strsql & "B.Pkid = CB.Intbanco AND "
        strsql = strsql & "A.Pkid = CB.INTAGENCIA AND "
        strsql = strsql & "MB.Dtmdtmovimento = " & gstrConvDtParaSql(txt_DataMovimento.Text)
    End If
    
    If chk_TodosContas.Value = 0 Then
        strsql = strsql & " AND CB.Pkid =" & dbcintContaBancaria.BoundText
        If chk_TodosLotes.Value = 0 Then
            strsql = strsql & " AND MB.INTLOTE =" & dbc_intLote.BoundText
        End If
    End If
    strsql = strsql & " GROUP BY "
    strsql = strsql & "MB.Dtmdtmovimento, "
    strsql = strsql & "B.STRDESCRICAO, "
    strsql = strsql & "CB.strConta, "
    strsql = strsql & "CB.strDigitoVerificador, "
    strsql = strsql & "A.STRDESCRICAO, "
    strsql = strsql & "MB.INTLOTE, LA.Strinscricao, "
    strsql = strsql & "LA.Strcomposicaodareceita, "
    strsql = strsql & "LA.Intexercicio, "
    strsql = strsql & "LA.strNumeroAviso , "
    strsql = strsql & "LA.intUtilizacao, "
    strsql = strsql & "MB.Dtmdtpagamento, "
    strsql = strsql & "CRB.INTTIPOCRITICA, "
    strsql = strsql & "LV.intParcela, "
    strsql = strsql & "MB.strCodigoDeBarras, "
    strsql = strsql & "( " & gstrISNULL("MB.Dblprincipal", "0") & "  +  " & gstrISNULL("MB.Dblmulta", "0") & "  +  " & gstrISNULL("MB.Dbljuros", "0") & "  +  " & gstrISNULL("MB.Dblcorrecao", "0") & " ) "
    strsql = strsql & " ORDER BY "
    strsql = strsql & "MB.Dtmdtmovimento, "
    strsql = strsql & "B.STRDESCRICAO, "
    strsql = strsql & "CB.STRCONTA, "
    strsql = strsql & "MB.INTLOTE, "
    
    strsql = strsql & "LA.strComposicaoDaReceita, "
    strsql = strsql & "LA.strInscricao, "
    strsql = strsql & "LA.intExercicio, "
    strsql = strsql & "LV.intParcela "
    
    strQueryRelatorio = strsql
    
End Function

Private Function strQueryContaCorrente() As String
    Dim strsql As String

    strsql = "SELECT CB.Pkid, "
    strsql = strsql & gstrCONVERT(CDT_VARCHAR, "CB.strConta") & strCONCAT & "'-'" & strCONCAT & " CB.strDigitoVerificador ContaCorrente"
    strsql = strsql & " FROM " & gstrContaBancaria & " CB, "
    strsql = strsql & gstrResumoBancario & " RB"
    strsql = strsql & " WHERE"
    strsql = strsql & " RB.intContaBancaria = CB.Pkid "
    If Trim(txt_DataMovimento) <> "" Then
        strsql = strsql & " AND RB.dtmData = " & gstrConvDtParaSql(txt_DataMovimento.Text)
    End If
    strsql = strsql & " ORDER BY intNumeroConta, strDigitoVerificador"

    strQueryContaCorrente = strsql

End Function

Private Sub PreencheAgBanco(lngPkidContaBancaria As Long)
Dim adoResultado    As ADODB.Recordset
Dim strsql          As String

    strsql = "SELECT BA.strDescricao Banco,"
    strsql = strsql & " AG.strDescricao Agencia"
    strsql = strsql & " FROM "
    strsql = strsql & gstrContaBancaria & " CB, "
    strsql = strsql & gstrBanco & " BA, "
    strsql = strsql & gstrAgencia & " AG"
    strsql = strsql & " WHERE"
    strsql = strsql & " BA.Pkid " & strOUTJOracle & "=" & strOUTJSQLServer & " CB.intBanco AND"
    strsql = strsql & " AG.Pkid " & strOUTJOracle & "=" & strOUTJSQLServer & " CB.intAgencia AND"
    strsql = strsql & " CB.Pkid = " & lngPkidContaBancaria
    
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strsql, 5, adoResultado) Then
        If Not adoResultado.EOF Then
            txt_strAgencia.Text = gstrENulo(adoResultado!Agencia)
            txt_strBanco.Text = gstrENulo(adoResultado!Banco)
        Else
            txt_strAgencia.Text = ""
            txt_strBanco.Text = ""
        End If
    End If
End Sub

Private Sub txt_DataMovimento_GotFocus()
    MarcaCampo txt_DataMovimento
End Sub

Private Sub txt_DataMovimento_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "D", txt_DataMovimento
End Sub

Private Sub txt_DataMovimento_LostFocus()
    txt_DataMovimento = gstrDataFormatada(txt_DataMovimento)
End Sub

Private Function strQueryLotes() As String
Dim strsql As String

    strsql = "SELECT RB.intLote,"
    strsql = strsql & " RB.intLote"
    strsql = strsql & " FROM "
    strsql = strsql & gstrContaBancaria & " CB, "
    strsql = strsql & gstrResumoBancario & " RB"
    strsql = strsql & " WHERE"
    strsql = strsql & " RB.intContaBancaria " & strOUTJSQLServer & "= CB.Pkid " & strOUTJOracle & " AND"
    strsql = strsql & " RB.dtmData = " & gstrConvDtParaSql(txt_DataMovimento)
    strsql = strsql & " AND RB.intContaBancaria = " & dbcintContaBancaria.BoundText
    strsql = strsql & " GROUP BY RB.intContaBancaria, RB.intLote"
    strsql = strsql & " ORDER BY RB.intLote"

    strQueryLotes = strsql
    
End Function

Private Function blnDadosOk() As Boolean
    blnDadosOk = False

    If Trim(txt_DataMovimento.Text) = "" Then
        ExibeMensagem "A data do Movimento deve ser preenchido Corretamente."
        txt_DataMovimento.SetFocus
        Exit Function
    End If
    
    If chk_TodosContas.Value = 0 Then
        If dbcintContaBancaria.MatchedWithList = False Then
            ExibeMensagem "A Conta Bancária deve ser preenchida Corretamente."
            dbcintContaBancaria.SetFocus
            Exit Function
        End If
        If chk_TodosLotes.Value = 0 Then
            If dbc_intLote.MatchedWithList = False Then
                ExibeMensagem "A Lote deve ser preenchido Corretamente."
                dbcintContaBancaria.SetFocus
                Exit Function
            End If
        End If
    End If
    
    blnDadosOk = True
End Function


