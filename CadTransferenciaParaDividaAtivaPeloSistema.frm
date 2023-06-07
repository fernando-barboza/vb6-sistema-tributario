VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmCadTransferenciaParaDividaAtivaPeloSistema 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Transferência de Débitos para Dívida Ativa - Gerados pelo Sistema"
   ClientHeight    =   4005
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7965
   Icon            =   "CadTransferenciaParaDividaAtivaPeloSistema.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4005
   ScaleWidth      =   7965
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab tab_3dPasta 
      Height          =   3645
      Left            =   180
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   150
      Width           =   7590
      _ExtentX        =   13388
      _ExtentY        =   6429
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Transferência de Débitos para Dívida Ativa"
      TabPicture(0)   =   "CadTransferenciaParaDividaAtivaPeloSistema.frx":1042
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lbl_NumeroPagina"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lbl_NumeroLivroInscricao"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lbl_NumeroInscricao"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lbl_DataInscricao"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lbl_NumeroParcelaInicial"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lbl_Exercicio"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lbl_DataVencimentoInicial"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lbl_NumeroParcelaFinal"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lbl_DataVencimentoFinal"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "lbl_strComposicaoDividaAtiva"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "lbl_strComposicaoReceita"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "dbc_strComposicaoDividaAtiva"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "dbc_strComposicaoReceita"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txt_NumeroPaginaInscricao"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txt_NumeroLivroInscricao"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txt_DataInscricao"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "txt_Exercicio"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "txt_NumeroInscricao"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "txt_NumeroParcelaInicial"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "txt_VencimentoInicial"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "txt_NumeroParcelaFinal"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "txt_VencimentoFinal"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).ControlCount=   22
      Begin VB.TextBox txt_VencimentoFinal 
         Height          =   285
         Left            =   6240
         TabIndex        =   6
         Top             =   2220
         Width           =   1035
      End
      Begin VB.TextBox txt_NumeroParcelaFinal 
         Height          =   285
         Left            =   6240
         MaxLength       =   3
         TabIndex        =   4
         Top             =   1860
         Width           =   1035
      End
      Begin VB.TextBox txt_VencimentoInicial 
         Height          =   285
         Left            =   2340
         TabIndex        =   5
         Top             =   2220
         Width           =   1035
      End
      Begin VB.TextBox txt_NumeroParcelaInicial 
         Height          =   285
         Left            =   2340
         MaxLength       =   3
         TabIndex        =   3
         Top             =   1860
         Width           =   1035
      End
      Begin VB.TextBox txt_NumeroInscricao 
         Height          =   285
         Left            =   2340
         MaxLength       =   8
         TabIndex        =   9
         Top             =   2940
         Width           =   1035
      End
      Begin VB.TextBox txt_Exercicio 
         Height          =   285
         Left            =   2340
         MaxLength       =   4
         TabIndex        =   2
         Top             =   1500
         Width           =   705
      End
      Begin VB.TextBox txt_DataInscricao 
         Height          =   285
         Left            =   6240
         TabIndex        =   10
         Top             =   2940
         Width           =   1035
      End
      Begin VB.TextBox txt_NumeroLivroInscricao 
         Height          =   285
         Left            =   2340
         MaxLength       =   8
         TabIndex        =   7
         Top             =   2580
         Width           =   1035
      End
      Begin VB.TextBox txt_NumeroPaginaInscricao 
         Height          =   285
         Left            =   6240
         MaxLength       =   8
         TabIndex        =   8
         Top             =   2580
         Width           =   1035
      End
      Begin MSDataListLib.DataCombo dbc_strComposicaoReceita 
         Height          =   315
         Left            =   2340
         TabIndex        =   0
         Top             =   690
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   556
         _Version        =   393216
         IntegralHeight  =   0   'False
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dbc_strComposicaoDividaAtiva 
         Height          =   315
         Left            =   2340
         TabIndex        =   1
         Top             =   1095
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   556
         _Version        =   393216
         IntegralHeight  =   0   'False
         Style           =   2
         Text            =   ""
      End
      Begin VB.Label lbl_strComposicaoReceita 
         AutoSize        =   -1  'True
         Caption         =   "Composição da Receita"
         Height          =   195
         Left            =   570
         TabIndex        =   22
         Top             =   735
         Width           =   1695
      End
      Begin VB.Label lbl_strComposicaoDividaAtiva 
         AutoSize        =   -1  'True
         Caption         =   "Composição da Dívida Ativa"
         Height          =   195
         Left            =   240
         TabIndex        =   21
         Top             =   1140
         Width           =   2025
      End
      Begin VB.Label lbl_DataVencimentoFinal 
         AutoSize        =   -1  'True
         Caption         =   "Data de Vencimento Final"
         Height          =   195
         Left            =   4320
         TabIndex        =   20
         Top             =   2310
         Width           =   1830
      End
      Begin VB.Label lbl_NumeroParcelaFinal 
         AutoSize        =   -1  'True
         Caption         =   "Nº da Parcela Final"
         Height          =   195
         Left            =   4785
         TabIndex        =   19
         Top             =   1950
         Width           =   1365
      End
      Begin VB.Label lbl_DataVencimentoInicial 
         AutoSize        =   -1  'True
         Caption         =   "Data de Vencimento Inicial"
         Height          =   195
         Left            =   360
         TabIndex        =   18
         Top             =   2280
         Width           =   1905
      End
      Begin VB.Label lbl_Exercicio 
         AutoSize        =   -1  'True
         Caption         =   "Exercício"
         Height          =   195
         Left            =   1590
         TabIndex        =   17
         Top             =   1560
         Width           =   675
      End
      Begin VB.Label lbl_NumeroParcelaInicial 
         AutoSize        =   -1  'True
         Caption         =   "Nº da Parcela Incial"
         Height          =   195
         Left            =   855
         TabIndex        =   16
         Top             =   1920
         Width           =   1410
      End
      Begin VB.Label lbl_DataInscricao 
         AutoSize        =   -1  'True
         Caption         =   "Data da Inscrição"
         Height          =   195
         Left            =   4890
         TabIndex        =   15
         Top             =   3030
         Width           =   1260
      End
      Begin VB.Label lbl_NumeroInscricao 
         AutoSize        =   -1  'True
         Caption         =   "Número de Inscrição"
         Height          =   195
         Left            =   795
         TabIndex        =   14
         Top             =   3000
         Width           =   1470
      End
      Begin VB.Label lbl_NumeroLivroInscricao 
         AutoSize        =   -1  'True
         Caption         =   "Nº do Livro"
         Height          =   195
         Left            =   1470
         TabIndex        =   13
         Top             =   2640
         Width           =   795
      End
      Begin VB.Label lbl_NumeroPagina 
         AutoSize        =   -1  'True
         Caption         =   "Primeira Página"
         Height          =   195
         Left            =   5055
         TabIndex        =   12
         Top             =   2670
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmCadTransferenciaParaDividaAtivaPeloSistema"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mblnAlterando               As Boolean
Dim mobjAux                     As Object
Dim mblnSelecionou              As Boolean
Dim mblnPrimeiraVez             As Boolean
Dim blnexistecontribuinte       As Boolean
Dim adoRecDados                 As ADODB.Recordset
Dim adoExisteContribuinte       As ADODB.Recordset

Private Sub dbc_strComposicaoDividaAtiva_Click(Area As Integer)
    DropDownDataCombo dbc_strComposicaoDividaAtiva, Me, Area
End Sub

Private Sub dbc_strComposicaoDividaAtiva_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbc_strComposicaoDividaAtiva, Me, , KeyCode, Shift
End Sub

Private Sub dbc_strComposicaoDividaAtiva_LostFocus()
    If dbc_strComposicaoReceita.BoundText <> "" Then
        HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrCalcularReajuste
    End If
End Sub

Private Sub dbc_strComposicaoReceita_Click(Area As Integer)
    DropDownDataCombo dbc_strComposicaoReceita, Me, Area
End Sub

Private Sub dbc_strComposicaoReceita_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbc_strComposicaoReceita, Me, , KeyCode, Shift
End Sub

Private Sub dbc_strComposicaoReceita_LostFocus()
    If dbc_strComposicaoDividaAtiva.BoundText <> "" Then
        HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrCalcularReajuste
    End If
End Sub

Private Sub Form_Activate()
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrNovo
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrSalvar
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrDeletar
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar
End Sub

Private Sub Form_Load()
    mblnAlterando = False
    LeDaTabelaParaObj gstrComposicaoDaReceita, dbc_strComposicaoReceita, QueryComposicaoReceita
    LeDaTabelaParaObj gstrComposicaoDaReceita, dbc_strComposicaoDividaAtiva, QueryComposicaoDividaAtiva
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrCalcularReajuste
    mblnSelecionou = False
    mblnPrimeiraVez = False
End Sub

Public Sub MantemForm(ByVal strModoOperacao As String)
Dim strSql As String

    If UCase(strModoOperacao) = UCase(gstrFechar) Then
        Unload Me
    End If
    
    If strModoOperacao = gstrCalcularReajuste Then
        EfetuaTransferenciaParaDividaAtiva
    End If
        
End Sub

Private Function QueryComposicaoReceita() As String
    Dim strSql As String
    strSql = ""
    strSql = strSql & "SELECT PKId, strDescricao FROM " & gstrComposicaoDaReceita
    strSql = strSql & " WHERE bytDividaAtiva <> 1 "
    strSql = strSql & " ORDER BY strDescricao "
    QueryComposicaoReceita = strSql
End Function

Private Function QueryComposicaoDividaAtiva() As String
    Dim strSql As String
    strSql = ""
    strSql = strSql & "SELECT PKId, strDescricao FROM " & gstrComposicaoDaReceita
    strSql = strSql & " WHERE bytDividaAtiva = 1 " 'só Dívida Ativa
    strSql = strSql & " ORDER BY strDescricao "
    QueryComposicaoDividaAtiva = strSql
End Function

Private Sub EfetuaTransferenciaParaDividaAtiva()

'******************************************************************************************
' Data: 02/05/2003
' Alteração: - Substituição do comando nativo GETDATE() do SQL Server pela variável
'            strGETDATE.
' Responsável: Everton Bianchini
'------------------------------------------------------------------------------------------
' Data: 09/05/2003
' Alteração: - Incluídos comandos de inicialização e finalização de bloco PL/SQL permitindo
'            , assim, a execução de múltiplos comandos SQL de uma única vez.
' Responsável: Everton Bianchini
'******************************************************************************************
    
Dim strSql                  As String
Dim Contribuinte            As Integer
Dim NumeroInscricao         As Integer
Dim ContaPagina             As Integer
Dim NumeroPagina            As Integer

Set gobjBanco = New clsBanco

ContaPagina = 1
NumeroPagina = Val(txt_NumeroPaginaInscricao.Text)
NumeroInscricao = Val(txt_NumeroInscricao.Text)

If blnDadosOk Then

    BuscaDados
    If adoRecDados.RecordCount = 0 Then
        ExibeMensagem "Não existem dados a serem " & Chr(13) & " transferidos neste período!"
        Exit Sub
    End If
    
    adoRecDados.MoveFirst
    strSql = ""
    
    strSql = strSql & IIf((bytDBType = EDatabases.Oracle), "BEGIN ", "")

    gobjBanco.ExecutaBeginTrans
    Screen.MousePointer = vbHourglass

    Do While Not adoRecDados.EOF
        Contribuinte = adoRecDados!intContribuinte
    
        If Not ExisteContribuinte Then
        
            strSql = strSql & " INSERT INTO " & gstrDividaAtiva
            strSql = strSql & " (intContribuinte, dtmDtAtualizacao, lngCodUsr ) VALUES ( "
            strSql = strSql & adoRecDados!intContribuinte
'            strSql = strSql & ", GETDATE()"
            strSql = strSql & ", " & strGETDATE
            strSql = strSql & ", " & glngCodUsr & " )"
            
            strSql = strSql & IIf((bytDBType = EDatabases.Oracle), "; ", " ")
        
        End If
        
        Do While adoRecDados!intContribuinte = Contribuinte

            If ContaPagina <= 60 Then
                NumeroPagina = NumeroPagina
            Else
                NumeroPagina = NumeroPagina + 1
                ContaPagina = 1
            End If
            
            strSql = strSql & " INSERT INTO " & gstrDetalheDividaAtiva
            strSql = strSql & " (intDividaAtiva, strInscricaoCadastral, intExercicio, dtmVencimento, "
            strSql = strSql & " intNumeroParcela, dtmInscricao, intComposicaoReceita, intOcorrencia, "
            strSql = strSql & " dblValorOriginal, dblValorAtual, bytOrigem, bytDebitoGeradoManualmente, "
            strSql = strSql & " bytSituacao, intNumeroLivroInscricao, intNumeroPaginaInscricao, intNumeroInscricao, "
            strSql = strSql & " dtmDtAtualizacao, lngCodUsr ) "
            If blnexistecontribuinte Then
                strSql = strSql & " VALUES (" & adoExisteContribuinte!PKId
            Else
                strSql = strSql & " (SELECT MAX(PKId) "
            End If
            strSql = strSql & ", '" & adoRecDados!strInscricaoCadastral
            strSql = strSql & "', " & adoRecDados!intExercicio
            strSql = strSql & ", " & gstrConvDtParaSql(adoRecDados!dtmDataVencimento)
            strSql = strSql & ", " & adoRecDados!intNumeroParcela
            strSql = strSql & ", " & gstrConvDtParaSql(txt_DataInscricao.Text)
            strSql = strSql & ", " & dbc_strComposicaoReceita.BoundText 'troca composicao receita por composicao dívida ativa
            strSql = strSql & ", " & adoRecDados!intOcorrencia
            strSql = strSql & ", " & gstrConvVrParaSql(adoRecDados!dblValorParcela)
            strSql = strSql & ", " & gstrConvVrParaSql(adoRecDados!dblValorParcela)
            strSql = strSql & ", " & adoRecDados!bytOrigem
            strSql = strSql & ", 0 " 'zero é débito gerado pelo sistema - 1 é débito gerado manualmente
            strSql = strSql & ", 2 " 'situacao "Em Aberto"
            strSql = strSql & ", " & txt_NumeroLivroInscricao.Text
            strSql = strSql & ", " & NumeroPagina
            strSql = strSql & ", " & NumeroInscricao
'            strSql = strSql & ", GETDATE()"
            strSql = strSql & ", " & strGETDATE
            strSql = strSql & ", " & glngCodUsr
            If blnexistecontribuinte Then
                strSql = strSql & ")"
            Else
                strSql = strSql & " FROM " & gstrDividaAtiva
                strSql = strSql & ")"
            End If
            
            strSql = strSql & IIf((bytDBType = EDatabases.Oracle), "; ", " ")
            
            strSql = strSql & " UPDATE " & gstrParcelaReceita & " SET "
            strSql = strSql & " bytAtiva = 1 "
            strSql = strSql & " WHERE PKId = " & adoRecDados!PKId
            
            strSql = strSql & IIf((bytDBType = EDatabases.Oracle), "; ", "")
            
            ContaPagina = ContaPagina + 1
            NumeroInscricao = NumeroInscricao + 1
            
            adoRecDados.MoveNext

            If adoRecDados.EOF Then
                Exit Do
            End If
            
        Loop
        
    Loop
    
    strSql = strSql & IIf((bytDBType = EDatabases.Oracle), "END; ", "")
    
    Set gobjBanco = New clsBanco
    If gobjBanco.Execute(strSql, False) Then
        gobjBanco.ExecutaCommitTrans
        Screen.MousePointer = vbNormal
        ExibeMensagem "Tranferência efetuada com sucesso!"
    Else
        Screen.MousePointer = vbNormal
        gobjBanco.ExecutaRollbackTrans
    End If

End If
End Sub

Private Function ExisteContribuinte() As Boolean
Dim strSql                  As String

blnexistecontribuinte = False
    
    strSql = ""
    strSql = strSql & " SELECT PKId, intContribuinte "
    strSql = strSql & " FROM " & gstrDividaAtiva
    strSql = strSql & " WHERE intcontribuinte = " & adoRecDados!intContribuinte
    
Set gobjBanco = New clsBanco
gobjBanco.CriaADO strSql, 5, adoExisteContribuinte
    
If adoExisteContribuinte.RecordCount > 0 Then
    ExisteContribuinte = True
    blnexistecontribuinte = True
End If

End Function

Private Sub BuscaDados()
Dim strSql          As String

strSql = ""
strSql = strSql & " SELECT PR.PKId, LC.intContribuinte, LC.strInscricaoCadastral, LC.intExercicio, "
strSql = strSql & " PR.dtmDataVencimento, PR.intNumeroParcela, LC.intOcorrencia, PR.dblValorParcela, "
strSql = strSql & " LC.BytOrigem "
strSql = strSql & " FROM " & gstrLancamentoCalculo & " LC, "
''strSql = strSql & gstrParcelaTaxa & " PT, "
strSql = strSql & gstrParcelaReceita & " PR "
strSql = strSql & " WHERE LC.PKId = PR.intLancamentoCalculo "
''strSql = strSql & " AND LC.PKId = PT.intLancamentoCalculo "
strSql = strSql & " AND PR.dtmDataVencimento BETWEEN " & gstrConvDtParaSql(txt_VencimentoInicial.Text) & " AND " & gstrConvDtParaSql(txt_VencimentoFinal.Text)
strSql = strSql & " AND PR.intNumeroParcela BETWEEN " & Val(txt_NumeroParcelaInicial.Text) & " AND " & Val(txt_NumeroParcelaFinal.Text)
strSql = strSql & " AND LC.intExercicio = " & txt_Exercicio.Text
strSql = strSql & " AND LC.intComposicaoReceita = " & dbc_strComposicaoReceita.BoundText
strSql = strSql & " AND PR.bytAtiva = 0 "
strSql = strSql & " AND PR.bytSuspensaoDeExigencia = 0 "
strSql = strSql & " ORDER BY intContribuinte "

Set gobjBanco = New clsBanco
gobjBanco.CriaADO strSql, 5, adoRecDados

End Sub

Private Function blnDadosOk() As Boolean
    blnDadosOk = False
    If txt_Exercicio.Text = "" Then
        ExibeMensagem "O campo " & lbl_Exercicio.Caption & " não pode ser em branco."
        txt_Exercicio.SetFocus
        Exit Function
    End If
    If txt_NumeroParcelaInicial.Text = "" Then
        ExibeMensagem "O campo " & lbl_NumeroParcelaInicial.Caption & " não pode ser em branco."
        txt_NumeroParcelaInicial.SetFocus
        Exit Function
    End If
    If txt_NumeroParcelaFinal.Text = "" Then
        ExibeMensagem "O campo " & lbl_NumeroParcelaFinal.Caption & " não pode ser em branco."
        txt_NumeroParcelaFinal.SetFocus
        Exit Function
    End If
    If Val(txt_NumeroParcelaInicial.Text) > Val(txt_NumeroParcelaFinal.Text) Then
        ExibeMensagem "A Parcela Inicial não pode ser superior a Parcela Final."
        txt_NumeroParcelaFinal.SetFocus
        Exit Function
    End If
    If txt_VencimentoInicial.Text = "" Then
        ExibeMensagem "O campo " & lbl_DataVencimentoInicial.Caption & " não pode ser em branco."
        txt_VencimentoInicial.SetFocus
        Exit Function
    End If
    If txt_VencimentoFinal.Text = "" Then
        ExibeMensagem "O campo " & lbl_DataVencimentoFinal.Caption & " não pode ser em branco."
        txt_VencimentoFinal.SetFocus
        Exit Function
    End If
    If Not gblnDataValida(txt_VencimentoInicial) Then
        ExibeMensagem lbl_DataVencimentoInicial.Caption & " é inválido."
        txt_VencimentoInicial.SetFocus
        Exit Function
    End If
    If Not gblnDataValida(txt_VencimentoFinal) Then
        ExibeMensagem lbl_DataVencimentoFinal.Caption & " é inválido."
        txt_VencimentoFinal.SetFocus
        Exit Function
    End If
    If CDate(txt_VencimentoInicial) > CDate(txt_VencimentoFinal) Then
        ExibeMensagem "A Data de Vencimento Inicial não pode ser superior a Data de Vencimento Final."
        txt_VencimentoFinal.SetFocus
        Exit Function
    End If
    If txt_NumeroLivroInscricao.Text = "" Then
        ExibeMensagem "O campo " & lbl_NumeroLivroInscricao.Caption & " não pode ser em branco."
        txt_NumeroLivroInscricao.SetFocus
        Exit Function
    End If
    If txt_NumeroPaginaInscricao.Text = "" Then
        ExibeMensagem "O campo " & lbl_NumeroPagina.Caption & " não pode ser em branco."
        txt_NumeroPaginaInscricao.SetFocus
        Exit Function
    End If
    If txt_NumeroInscricao.Text = "" Then
        ExibeMensagem "O campo " & lbl_NumeroInscricao.Caption & " não pode ser em branco."
        txt_NumeroInscricao.SetFocus
        Exit Function
    End If
    If txt_DataInscricao.Text = "" Then
        ExibeMensagem "O campo " & lbl_DataInscricao.Caption & " não pode ser em branco."
        txt_DataInscricao.SetFocus
        Exit Function
    ElseIf Not gblnDataValida(txt_DataInscricao) Then
        ExibeMensagem lbl_DataInscricao.Caption & " é inválido."
        txt_DataInscricao.SetFocus
        Exit Function
    ElseIf Not gblnDataValida(txt_VencimentoFinal) Then
        ExibeMensagem lbl_DataVencimentoFinal.Caption & " é inválido."
        txt_VencimentoFinal.SetFocus
        Exit Function
    ElseIf CDate(txt_DataInscricao) < CDate(txt_VencimentoFinal) Then
        ExibeMensagem "O campo " & lbl_DataInscricao.Caption & " tem que ser maior que o campo " & lbl_DataVencimentoFinal.Caption & "."
        txt_DataInscricao.SetFocus
        Exit Function
    End If
    blnDadosOk = True
End Function

'###################### CARACTER VÁLIDO E MARCA CAMPO ##############################

Private Sub dbc_strComposicaoReceita_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "", dbc_strComposicaoReceita
End Sub

Private Sub dbc_strComposicaoDividaAtiva_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "", dbc_strComposicaoDividaAtiva
End Sub

Private Sub txt_Exercicio_GotFocus()
    MarcaCampo txt_Exercicio
End Sub

Private Sub txt_Exercicio_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txt_Exercicio
End Sub

Private Sub txt_NumeroParcelaInicial_GotFocus()
    MarcaCampo txt_NumeroParcelaInicial
End Sub

Private Sub txt_NumeroParcelaInicial_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txt_NumeroParcelaInicial
End Sub

Private Sub txt_NumeroParcelaFinal_GotFocus()
    MarcaCampo txt_NumeroParcelaFinal
End Sub

Private Sub txt_NumeroParcelaFinal_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txt_NumeroParcelaFinal
End Sub

Private Sub txt_VencimentoInicial_GotFocus()
    MarcaCampo txt_VencimentoInicial
End Sub

Private Sub txt_VencimentoInicial_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "D", txt_VencimentoInicial
End Sub

Private Sub txt_VencimentoInicial_LostFocus()
    txt_VencimentoInicial = gstrDataFormatada(txt_VencimentoInicial)
End Sub

Private Sub txt_VencimentoFinal_GotFocus()
    MarcaCampo txt_VencimentoFinal
End Sub

Private Sub txt_VencimentoFinal_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "D", txt_VencimentoFinal
End Sub

Private Sub txt_VencimentoFinal_LostFocus()
    txt_VencimentoFinal = gstrDataFormatada(txt_VencimentoFinal)
End Sub

Private Sub txt_NumeroLivroInscricao_GotFocus()
    MarcaCampo txt_NumeroLivroInscricao
End Sub

Private Sub txt_NumeroLivroInscricao_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txt_NumeroLivroInscricao
End Sub

Private Sub txt_NumeroPaginaInscricao_GotFocus()
    MarcaCampo txt_NumeroPaginaInscricao
End Sub

Private Sub txt_NumeroPaginaInscricao_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txt_NumeroPaginaInscricao
End Sub

Private Sub txt_NumeroInscricao_GotFocus()
    MarcaCampo txt_NumeroInscricao
End Sub

Private Sub txt_NumeroInscricao_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txt_NumeroInscricao
End Sub

Private Sub txt_DataInscricao_GotFocus()
    MarcaCampo txt_DataInscricao
End Sub

Private Sub txt_DataInscricao_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "D", txt_DataInscricao
End Sub

Private Sub txt_DataInscricao_LostFocus()
    txt_DataInscricao = gstrDataFormatada(txt_DataInscricao)
End Sub

''''L
