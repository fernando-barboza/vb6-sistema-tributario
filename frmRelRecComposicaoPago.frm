VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmRelRecComposicaoPago 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Receita de Composição - Lançado Pago"
   ClientHeight    =   2715
   ClientLeft      =   3465
   ClientTop       =   3510
   ClientWidth     =   6630
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2715
   ScaleWidth      =   6630
   Begin TabDlg.SSTab tab_3dPasta 
      Height          =   2625
      Left            =   60
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   30
      Width           =   6525
      _ExtentX        =   11509
      _ExtentY        =   4630
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Receita de Composição - Lançado Pago"
      TabPicture(0)   =   "frmRelRecComposicaoPago.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Fra_Pagamentos"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin VB.Frame Fra_Pagamentos 
         Height          =   2145
         Left            =   90
         TabIndex        =   6
         Top             =   360
         Width           =   6315
         Begin VB.CheckBox chkListarPagamentos 
            Caption         =   "Listar Somente Lançamentos Pagos"
            Height          =   255
            Left            =   2160
            TabIndex        =   11
            Top             =   1800
            Width           =   3345
         End
         Begin VB.TextBox txtdtmInicial 
            Height          =   285
            Left            =   2160
            MaxLength       =   10
            TabIndex        =   3
            Top             =   1470
            Width           =   1035
         End
         Begin VB.TextBox txtdtmFinal 
            Height          =   285
            Left            =   3780
            MaxLength       =   10
            TabIndex        =   4
            Top             =   1470
            Width           =   1035
         End
         Begin VB.TextBox txt_intExercicio 
            Height          =   315
            Left            =   2160
            MaxLength       =   4
            TabIndex        =   2
            Top             =   1080
            Width           =   525
         End
         Begin MSDataListLib.DataCombo dbc_intComposicao 
            Height          =   315
            Left            =   2160
            TabIndex        =   0
            Top             =   300
            Width           =   3795
            _ExtentX        =   6694
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dbc_intReceita 
            Height          =   315
            Left            =   2160
            TabIndex        =   1
            Top             =   690
            Width           =   3795
            _ExtentX        =   6694
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Receita"
            Height          =   195
            Left            =   1500
            TabIndex        =   13
            Top             =   750
            Width           =   555
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Data de Vencimento:"
            Height          =   195
            Left            =   120
            TabIndex        =   12
            Top             =   1500
            Width           =   1500
         End
         Begin VB.Label lblEmpenhoInicial 
            AutoSize        =   -1  'True
            Caption         =   "Início"
            Height          =   195
            Left            =   1680
            TabIndex        =   10
            Top             =   1500
            Width           =   405
         End
         Begin VB.Label lblEmpenhoFinal 
            AutoSize        =   -1  'True
            Caption         =   "Fim"
            Height          =   195
            Left            =   3390
            TabIndex        =   9
            Top             =   1500
            Width           =   240
         End
         Begin VB.Label lbl_Exercicio 
            AutoSize        =   -1  'True
            Caption         =   "Exercício"
            Height          =   195
            Left            =   1410
            TabIndex        =   8
            Top             =   1140
            Width           =   675
         End
         Begin VB.Label lbl_Composicao 
            AutoSize        =   -1  'True
            Caption         =   "Composição da Receita"
            Height          =   195
            Left            =   390
            TabIndex        =   7
            Top             =   390
            Width           =   1695
         End
      End
   End
End
Attribute VB_Name = "frmRelRecComposicaoPago"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim mobjAux          As Object
    Dim mblnSelecionou   As Boolean


Private Sub dbc_intComposicao_Click(Area As Integer)
    'Set dbc_intReceita.DataSource = Nothing
    'Set dbc_intReceita.RowSource = Nothing
    'dbc_intReceita.BoundText = ""
    dbc_intReceita.Tag = strQueryGrid
    PreencherListaDeOpcoes dbc_intReceita
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

Private Sub Form_Load()
    HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrSalvar, gstrDeletar
    dbc_intComposicao.Tag = strQueryComposicao & ";strDescricao"
End Sub
Private Sub Form_Activate()
    
    gintCodSeguranca = 1392
    HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrDeletar, gstrSalvar, gstrAplicar
    HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrNovo, gstrImprimir

End Sub
Private Sub Form_Deactivate()
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar
    mblnSelecionou = False
End Sub

Private Sub txtdtmInicial_GotFocus()
    MarcaCampo txtdtmInicial
End Sub

Private Sub txtdtmInicial_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "D", txtdtmInicial
End Sub

Private Sub txtdtmInicial_LostFocus()
    txtdtmInicial = gstrDataFormatada(txtdtmInicial)
End Sub

Private Sub txtdtmFinal_GotFocus()
    MarcaCampo txtdtmFinal
End Sub

Private Sub txtdtmFinal_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "D", txtdtmFinal
End Sub

Private Sub txtdtmFinal_LostFocus()
    txtdtmFinal = gstrDataFormatada(txtdtmFinal)
End Sub


Public Sub MantemForm(ByVal strModoOperacao As String)
    Dim strIntervalo As String

    If UCase(strModoOperacao) = UCase(gstrImprimir) Then
        If Len(txtdtmInicial) > 0 Then
            If txtdtmInicial = txtdtmFinal Then
                strIntervalo = " - Vencimento: " & txtdtmInicial
            Else
                strIntervalo = " - Vencto: " & txtdtmInicial & " até " & txtdtmFinal
            End If
            
        End If
        If blnDadosOk Then
            rptRecComposicaoPago.intUtilizacao = gintUtilizacao(gstrItemData(dbc_intComposicao))
            rptRecComposicaoPago.lblComposicaoReceita.Caption = "Composição: " & dbc_intComposicao.Text
            rptRecComposicaoPago.lblReceita.Caption = "Receita: " & dbc_intReceita.Text & " - " & txt_intExercicio.Text & strIntervalo
            ImprimeRelatorio rptRecComposicaoPago, strQueryRelatorio, , 90
        End If
    End If
    
    If UCase(strModoOperacao) = UCase(gstrNovo) Then
        Limpa_Controles Me, True, False, True, True, False
        dbc_intComposicao.SetFocus
    End If
    
    If UCase(strModoOperacao) = UCase(gstrFechar) Then
        Unload Me
    End If
    
    If UCase(strModoOperacao) = UCase(gstrPreencherLista) Then
        
        If Me.ActiveControl.Name = "dbc_intReceita" Then
            If Not dbc_intComposicao.MatchedWithList Then
                ExibeMensagem "É necessário preencher a Composição da Receita para escolher a Receita."
                If dbc_intReceita.Enabled Then dbc_intReceita.SetFocus
            Else
                dbc_intReceita.Tag = strQueryGrid
                PreencherListaDeOpcoes Me.ActiveControl
            End If
            
        Else
            PreencherListaDeOpcoes Me.ActiveControl
        End If
    
        
    End If
    
End Sub

Private Sub dbc_intReceita_GotFocus()
    MarcaCampo dbc_intReceita
End Sub

Private Sub dbc_intReceita_KeyDown(KeyCode As Integer, Shift As Integer)
        If Not dbc_intComposicao.MatchedWithList Then
                ExibeMensagem "É necessário preencher a Composição da Receita para escolher a Receita."
                If dbc_intReceita.Enabled Then dbc_intReceita.SetFocus
        Else
            dbc_intReceita.Tag = strQueryGrid
            DropDownDataCombo dbc_intReceita, Me, , KeyCode, Shift
        End If

End Sub

Private Sub dbc_intReceita_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbc_intReceita
End Sub

Private Function strQueryRelatorio() As String
    Dim strSql As String
    
    strSql = ""
    strSql = strSql & "SELECT "
    strSql = strSql & gstrRIGHT("LA.strinscricao", gintRetornaTamanhoMascara(gintUtilizacao(gstrItemData(dbc_intComposicao)))) & " strinscricao, "
    strSql = strSql & "LA.strnomeproprietario, "
    strSql = strSql & "LV.intParcela, "
    strSql = strSql & "LV.dtmdtvencimento, "
    strSql = strSql & "LV.dblvalor dblLancamentoValor, "
    strSql = strSql & "LR.dblvalor dblValorLancado, "
    strSql = strSql & "(LP.dblvalorprincipal + LP.DBLVALORMULTA + LP.DBLVALORJUROS )dblvalorprincipal , "
    strSql = strSql & "LP.dtmdtpagamento "
    
    strSql = strSql & "FROM "
    strSql = strSql & gstrLancamentoAlfa & " LA, "
    strSql = strSql & gstrLancamentoValor & " LV, "
    strSql = strSql & gstrLancamentoReceita & " LR, "
    strSql = strSql & gstrLancamentoPagamento & " LP "
    strSql = strSql & "Where "
    strSql = strSql & "LA.Pkid = LV.Intlancamentoalfa       AND "
    strSql = strSql & "LV.Pkid = LR.intlancamentovalor       AND "
    
    If chkListarPagamentos.Value = vbChecked Then
        strSql = strSql & "LV.Pkid = LP.Intlancamentovalor  AND "
    Else
        strSql = strSql & " LP.Intlancamentovalor " & strOUTJOracle & " =" & strOUTJSQLServer & " LV.Pkid AND "
    End If
    strSql = strSql & "LA.Intexercicio = " & txt_intExercicio & " AND "
    strSql = strSql & "LA.Intcomposicaodareceita = " & dbc_intComposicao.BoundText & " AND "
    strSql = strSql & "LR.Intreceita = " & dbc_intReceita.BoundText
    'strSql = strSql & " and la.strinscricao = '00000000000002020046' "
    If Len(txtdtmInicial) > 0 Then
        strSql = strSql & " AND LV.dtmdtvencimento BETWEEN " & gstrConvDtParaSql(txtdtmInicial)
        strSql = strSql & " AND " & gstrConvDtParaSql(txtdtmFinal)
    End If
    
    strSql = strSql & " ORDER BY LA.strinscricao,"
    strSql = strSql & " LA.strnomeproprietario, "
    strSql = strSql & " LV.dtmdtvencimento, "
    strSql = strSql & " LP.dtmdtpagamento "
    strQueryRelatorio = strSql
End Function

Private Function blnDadosOk() As Boolean
    blnDadosOk = False
    If Not dbc_intComposicao.MatchedWithList Then
        ExibeMensagem "O campo de composição da receita não foi preenchido corretamente."
        dbc_intComposicao.SetFocus
        Exit Function
    ElseIf Not dbc_intReceita.MatchedWithList Then
        ExibeMensagem "O campo de receita não foi preenchido corretamente."
        If dbc_intReceita.Enabled Then dbc_intReceita.SetFocus
        Exit Function
    
    ElseIf Trim(txt_intExercicio) = "" Then
        ExibeMensagem "O campo de exercício não foi preenchido corretamente."
        txt_intExercicio.SetFocus
        Exit Function
    End If
    If Len(txtdtmInicial) > 0 Or Len(txtdtmFinal) > 0 Then
        If gblnDataValida(txtdtmInicial) = False Then
            ExibeMensagem "Data inicial incorreta."
            txtdtmInicial.SetFocus
            Exit Function
        ElseIf gblnDataValida(txtdtmFinal) = False Then
            ExibeMensagem "Data final incorreta."
            txtdtmFinal.SetFocus
            Exit Function
        ElseIf CVDate(txtdtmFinal) < CVDate(txtdtmInicial) Then
            ExibeMensagem "Data inicial não poder menor que a data final."
            txtdtmInicial.SetFocus
            Exit Function
        End If
    Else
        ExibeMensagem "É necessário informar o intervalo de datas."
        txtdtmInicial.SetFocus
        Exit Function
    End If
    blnDadosOk = True
End Function

Private Function strQueryComposicao() As String
    Dim strSql As String
    
    strSql = "SELECT Pkid,"
    strSql = strSql & gstrCONVERT(CDT_VARCHAR, "intCodigo") & strCONCAT & "' - '" & strCONCAT & " strDescricao Descricao"
    strSql = strSql & " FROM "
    strSql = strSql & gstrComposicaoDaReceita
    strSql = strSql & " ORDER BY intCodigo"

    strQueryComposicao = strSql

End Function

Private Sub txt_intExercicio_GotFocus()
    MarcaCampo txt_intExercicio
End Sub

Private Sub txt_intExercicio_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txt_intExercicio
End Sub

Private Function gintUtilizacao(ByVal strpkidCompReceita As String) As Integer
    Dim strSql       As String
    Dim adoResultado As ADODB.Recordset
    
    strSql = strSql & " SELECT intutilizacao FROM "
    strSql = strSql & gstrComposicaoDaReceita & " CP "
    strSql = strSql & " WHERE "
    strSql = strSql & " pkid = " & strpkidCompReceita
    
    
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        If Not adoResultado.EOF Then
            gintUtilizacao = adoResultado!intUtilizacao
        End If
    End If

End Function


Function strQueryGrid() As String

'******************************************************************************************
' Data: 27/03/2003
' Alteração: - Retirado o comando CONVERT da cláusula ORDER BY uma vez que este não era
'            necessário.
' Responsável: Everton Bianchini
'******************************************************************************************
    
Dim strSql As String

   strSql = ""
   'strSql = strSql & "SELECT VL.intReceita, RC.strDescricao, RC.strSigla FROM "
   strSql = strSql & "SELECT VL.intReceita, RC.strDescricao FROM "
   strSql = strSql & gstrValorCompoRec & " VL,"
   strSql = strSql & gstrReceita & " RC "
   strSql = strSql & "WHERE RC.PKId = VL.intReceita "
   
   If dbc_intComposicao.MatchedWithList Then
      strSql = strSql & "AND VL.intComposicaoDaReceita = " & gstrItemData(dbc_intComposicao)
   End If
       
   strSql = strSql & "ORDER BY RC.strDescricao;strDescricao"
      
   
   strQueryGrid = strSql
   
End Function


