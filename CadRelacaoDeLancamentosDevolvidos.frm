VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCadRelacaoDeLancamentosDevolvidos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relação de Lançamentos Devolvidos"
   ClientHeight    =   2550
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6525
   Icon            =   "CadRelacaoDeLancamentosDevolvidos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2550
   ScaleWidth      =   6525
   Begin TabDlg.SSTab tab_3dPasta 
      Height          =   2415
      Left            =   60
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   120
      Width           =   6435
      _ExtentX        =   11351
      _ExtentY        =   4260
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Lançamentos Devolvidos"
      TabPicture(0)   =   "CadRelacaoDeLancamentosDevolvidos.frx":1042
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lbl_Devolucao"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lbl_Tal"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lbl_Tal2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblBal2"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lbl_Bal"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lin_linha"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "fra_Tipoo"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txt_Devolucao"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txt_Inicial"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txt_Final"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txt_Ate"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Fra_Analitico"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txt_BairroFinal"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txt_BairroInicial"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).ControlCount=   15
      Begin VB.TextBox txt_BairroInicial 
         Height          =   285
         Left            =   3420
         MaxLength       =   10
         TabIndex        =   5
         Top             =   1590
         Width           =   975
      End
      Begin VB.TextBox txt_BairroFinal 
         Height          =   285
         Left            =   5310
         MaxLength       =   10
         TabIndex        =   6
         Top             =   1560
         Width           =   975
      End
      Begin VB.Frame Fra_Analitico 
         Height          =   885
         Left            =   120
         TabIndex        =   16
         Top             =   1460
         Width           =   1425
         Begin VB.OptionButton opt_Analitico 
            Caption         =   "Bairro"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   4
            Top             =   240
            Width           =   1035
         End
         Begin VB.OptionButton opt_Analitico 
            Caption         =   "Contribuinte"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   7
            Top             =   600
            Width           =   1155
         End
      End
      Begin VB.TextBox txt_Ate 
         Height          =   285
         Left            =   4230
         MaxLength       =   10
         TabIndex        =   3
         Top             =   1065
         Width           =   975
      End
      Begin VB.TextBox txt_Final 
         Height          =   285
         Left            =   5310
         MaxLength       =   10
         TabIndex        =   9
         Top             =   1890
         Width           =   975
      End
      Begin VB.TextBox txt_Inicial 
         Height          =   285
         Left            =   3420
         MaxLength       =   10
         TabIndex        =   8
         Top             =   1920
         Width           =   975
      End
      Begin VB.TextBox txt_Devolucao 
         Height          =   285
         Left            =   2340
         MaxLength       =   10
         TabIndex        =   2
         Top             =   1065
         Width           =   975
      End
      Begin VB.Frame fra_Tipoo 
         Height          =   645
         Left            =   1800
         TabIndex        =   12
         Top             =   310
         Width           =   2865
         Begin VB.OptionButton opt_Tipo 
            Caption         =   "Sintético"
            Height          =   195
            Index           =   1
            Left            =   1620
            TabIndex        =   1
            Top             =   270
            Width           =   1035
         End
         Begin VB.OptionButton opt_Tipo 
            Caption         =   "Analítico"
            Height          =   195
            Index           =   0
            Left            =   270
            TabIndex        =   0
            Top             =   270
            Width           =   1035
         End
      End
      Begin VB.Line lin_linha 
         BorderColor     =   &H80000005&
         X1              =   120
         X2              =   6360
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Label lbl_Bal 
         AutoSize        =   -1  'True
         Caption         =   "Cod. inicial bairro"
         Height          =   195
         Left            =   2040
         TabIndex        =   18
         Top             =   1680
         Width           =   1200
      End
      Begin VB.Label lblBal2 
         AutoSize        =   -1  'True
         Caption         =   "Cod. final"
         Height          =   195
         Left            =   4500
         TabIndex        =   17
         Top             =   1650
         Width           =   660
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Até"
         Height          =   195
         Left            =   3720
         TabIndex        =   15
         Top             =   1065
         Width           =   240
      End
      Begin VB.Label lbl_Tal2 
         AutoSize        =   -1  'True
         Caption         =   "Cod. final"
         Height          =   195
         Left            =   4500
         TabIndex        =   14
         Top             =   1980
         Width           =   660
      End
      Begin VB.Label lbl_Tal 
         AutoSize        =   -1  'True
         Caption         =   "Cod. inicial contribuinte"
         Height          =   195
         Left            =   1650
         TabIndex        =   13
         Top             =   2010
         Width           =   1635
      End
      Begin VB.Label lbl_Devolucao 
         AutoSize        =   -1  'True
         Caption         =   "Data de Devolução"
         Height          =   195
         Left            =   810
         TabIndex        =   11
         Top             =   1065
         Width           =   1395
      End
   End
End
Attribute VB_Name = "frmCadRelacaoDeLancamentosDevolvidos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim mblnAlterando    As Boolean
    Dim mobjAux          As Object
    Dim mblnSelecionou   As Boolean
    Dim mblnPrimeiraVez  As Boolean

Private Sub Form_Activate()
    gintCodSeguranca = 684
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
    HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrSalvar, gstrDeletar
    txt_Inicial.Enabled = False
    TrocaCorObjeto txt_Inicial, True
    txt_Final.Enabled = False
    TrocaCorObjeto txt_Final, True
    txt_Ate = gstrDataFormatada(gstrDataDoSistema)
    opt_Tipo(0).Value = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar
    mblnSelecionou = False
    mblnPrimeiraVez = False
End Sub

Public Sub MantemForm(ByVal strModoOperacao As String)
Dim adoRelatorio   As ADODB.Recordset
On Error GoTo ErroImprimeRelatorio

    If UCase(strModoOperacao) = UCase(gstrImprimir) And opt_Tipo(1).Value = True Then
        If blnDadosOk = False Then
            Exit Sub
        End If
        ImprimeRelatorio rptLancamentoSintetico, strQuerrySintetico
    End If
    
    If UCase(strModoOperacao) = UCase(gstrImprimir) And opt_Tipo(0).Value = True And opt_Analitico(0).Value = True Then
        If blnDadosOK2 = False Then
            Exit Sub
        End If
        Screen.MousePointer = vbHourglass
        Set gobjBanco = New clsBanco
        If gobjBanco.CriaADO(strQuerryAnaliticoBairro, 5, adoRelatorio) Then
            Set rptLancamentoAnaliticoBairro.adoDataControl.Recordset = adoRelatorio
            rptLancamentoAnaliticoBairro.Show
        End If
        Screen.MousePointer = vbDefault
        Exit Sub
    ElseIf UCase(strModoOperacao) = UCase(gstrImprimir) And opt_Tipo(0).Value = True And opt_Analitico(1).Value = True Then
        If blnDadosOK3 = False Then
            Exit Sub
        End If
        Screen.MousePointer = vbHourglass
        Set gobjBanco = New clsBanco
        If gobjBanco.CriaADO(strQuerryAnaliticoContribuinte, 5, adoRelatorio) Then
            Set rptLancamentoAnaliticoContribuinte.adoDataControl.Recordset = adoRelatorio
            rptLancamentoAnaliticoContribuinte.Show
        End If
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    If UCase(strModoOperacao) = UCase(gstrNovo) Then
        LimpaObjetos
    End If
    If UCase(strModoOperacao) = UCase(gstrFechar) Then
        Unload Me
    End If
Screen.MousePointer = vbDefault

ErroImprimeRelatorio:
    GoTo FimImprimeRelatorio
FimImprimeRelatorio:

End Sub

Private Function blnDadosOk() As Boolean
blnDadosOk = False
        
    If txt_Devolucao = "" Then
        ExibeMensagem "A data de devolução tem que ser digitada."
        txt_Devolucao.SetFocus
        Exit Function
    Else
        If gblnDataValida(txt_Devolucao.Text) = False Then
            ExibeMensagem "A data de devolução não é válida."
            txt_Devolucao.SetFocus
            Exit Function
        End If
    End If
    
    If txt_Ate = "" Then
        ExibeMensagem "A data final tem que ser digitada."
        txt_Ate.SetFocus
        Exit Function
    Else
        If gblnDataValida(txt_Ate.Text) = False Then
            ExibeMensagem "A data final não é válida."
            txt_Ate.SetFocus
            Exit Function
        End If
    End If
    
    If CVDate(txt_Inicial) > CVDate(txt_Ate) Then
        ExibeMensagem "A data inicial tem que ser anterior à data final."
        txt_Inicial.SetFocus
        Exit Function
    End If

blnDadosOk = True
End Function


Private Function blnDadosOK2() As Boolean
blnDadosOK2 = False
    
    If txt_BairroInicial = "" Then
        ExibeMensagem "O código inicial do bairro tem que ser digitado."
        Exit Function
    End If
    
    If txt_BairroFinal = "" Then
        ExibeMensagem "O código final do bairro tem que ser digitado."
        Exit Function
    End If
    
    If Val(txt_BairroInicial) > Val(txt_BairroFinal) Then
        ExibeMensagem "O código inicial tem que ser menor que o código final."
        Exit Function
    End If
        
    If txt_Devolucao = "" Then
        ExibeMensagem "A data de devolução tem que ser digitada."
        txt_Devolucao.SetFocus
        Exit Function
    Else
        If gblnDataValida(txt_Devolucao.Text) = False Then
            ExibeMensagem "A data de devolução não é válida."
            txt_Devolucao.SetFocus
            Exit Function
        End If
    End If
    
    If txt_Ate = "" Then
        ExibeMensagem "A data final tem que ser digitada."
        txt_Ate.SetFocus
        Exit Function
    Else
        If gblnDataValida(txt_Ate.Text) = False Then
            ExibeMensagem "A data final não é válida."
            txt_Ate.SetFocus
            Exit Function
        End If
    End If
    
    If CVDate(txt_Inicial) > CVDate(txt_Ate) Then
        ExibeMensagem "A data inicial tem que ser anterior à data final."
        txt_Inicial.SetFocus
        Exit Function
    End If

blnDadosOK2 = True
End Function

Private Function blnDadosOK3() As Boolean
blnDadosOK3 = False
    
    If txt_Inicial = "" Then
        ExibeMensagem "O código inicial do contribuinte tem que ser digitado."
        Exit Function
    End If
    If txt_Final = "" Then
        ExibeMensagem "O código final do contribuinte tem que ser digitado."
        Exit Function
    End If
    If Val(txt_Inicial) > Val(txt_Final) Then
        ExibeMensagem "O código inicial tem que ser menor que o código final."
        Exit Function
    End If

    If txt_Devolucao = "" Then
        ExibeMensagem "A data de devolução tem que ser digitada."
        txt_Devolucao.SetFocus
        Exit Function
    Else
        If gblnDataValida(txt_Devolucao.Text) = False Then
            ExibeMensagem "A data de devolução não é válida."
            txt_Devolucao.SetFocus
            Exit Function
        End If
    End If
    
    If txt_Ate = "" Then
        ExibeMensagem "A data final tem que ser digitada."
        txt_Ate.SetFocus
        Exit Function
    Else
        If gblnDataValida(txt_Ate.Text) = False Then
            ExibeMensagem "A data final não é válida."
            txt_Ate.SetFocus
            Exit Function
        End If
    End If
    
    If CVDate(txt_Inicial) > CVDate(txt_Ate) Then
        ExibeMensagem "A data inicial tem que ser anterior à data final."
        txt_Inicial.SetFocus
        Exit Function
    End If
    
blnDadosOK3 = True
End Function


Sub LimpaObjetos()
    txt_Devolucao = ""
    txt_Ate = ""
    txt_Inicial = ""
    txt_Final = ""
    opt_Tipo(0).Value = True
End Sub

Private Function strQuerryAnaliticoContribuinte() As String
Dim strSql As String
Dim dtInicial  As Date
Dim dtFinal    As Date
Dim codInicial As Double
Dim codFinal   As Double
dtInicial = CVDate(txt_Devolucao)
dtFinal = CVDate(txt_Ate)
codInicial = Val(txt_Inicial)
codFinal = Val(txt_Final)

    strSql = ""
    strSql = strSql & " SELECT COUNT(*) as TOTAL , DV.strInscricao, DV.intContribuinte, CO.strNome, "
    strSql = strSql & " DE.strDescricao DOCUMENTOSEMITIDOS, OC.strDescricao OCORRENCIA, DV.intDocumentosEmitidos, DV.dtmDevolucao "
    strSql = strSql & " FROM " & gstrDevolucao & " DV, "
    strSql = strSql & gstrContribuinte & " CO, " & gstrOcorrencia & " OC, "
    strSql = strSql & gstrDocumentoEmitido & " DE "
    strSql = strSql & " WHERE DV.intContribuinte = CO.PKId "
    strSql = strSql & " AND DV.intOcorrencia = OC.PKId "
    strSql = strSql & " AND DV.intDocumentosEmitidos = DE.PKId "
    strSql = strSql & " AND DV.dtmDevolucao BETWEEN " & gstrConvDtParaSql(dtInicial) & " AND " & gstrConvDtParaSql(dtFinal)
    strSql = strSql & " AND DV.intContribuinte BETWEEN " & codInicial & " AND " & codFinal
    strSql = strSql & " GROUP BY DV.intDocumentosEmitidos, DE.strDescricao, CO.strNome,DV.strInscricao, "
    strSql = strSql & " DV.intContribuinte, OC.strDescricao, DV.dtmDevolucao "
    strSql = strSql & " ORDER BY CO.strNome "
strQuerryAnaliticoContribuinte = strSql

End Function

Private Function strQuerryAnaliticoBairro() As String

'******************************************************************************************
' Data: 05/05/2003
' Alteração: - Substituição do comando nativo de outer join (*= ou =*) do SQL Server pela
'            variável strOUTJSQLServer e inclusão do comando de outer join ((+)) do Oracle,
'            representado pela variável strOUTJOracle.
' Responsável: Everton Bianchini
'******************************************************************************************
    
Dim strSql     As String
Dim dtInicial  As Date
Dim dtFinal    As Date
Dim codInicial As Double
Dim codFinal   As Double
dtInicial = CVDate(txt_Devolucao)
dtFinal = CVDate(txt_Ate)
codInicial = Val(txt_BairroInicial)
codFinal = Val(txt_BairroFinal)

    strSql = ""
    strSql = strSql & " SELECT COUNT(*) as TOTAL , BA.strDescricao, DV.strInscricao, DV.intContribuinte, CO.strNome, CO.intBairro, "
    strSql = strSql & " DE.strDescricao DOCUMENTOSEMITIDOS, OC.strDescricao OCORRENCIA, DV.intDocumentosEmitidos, DV.dtmDevolucao "
    strSql = strSql & " FROM " & gstrDevolucao & " DV, "
    strSql = strSql & gstrContribuinte & " CO, " & gstrOcorrencia & " OC, "
    strSql = strSql & gstrDocumentoEmitido & " DE, " & gstrBairro & " BA "
    strSql = strSql & " WHERE DV.intContribuinte = CO.PKId "
'    strSql = strSql & " AND CO.intBairro *= BA.PKID "
    strSql = strSql & " AND CO.intBairro " & strOUTJOracle & strOUTJSQLServer & "= BA.PKID "
    strSql = strSql & " AND DV.intOcorrencia = OC.PKId "
    strSql = strSql & " AND DV.intDocumentosEmitidos = DE.PKId "
    strSql = strSql & " AND DV.dtmDevolucao BETWEEN " & gstrConvDtParaSql(dtInicial) & " AND " & gstrConvDtParaSql(dtFinal)
    strSql = strSql & " AND CO.intBairro BETWEEN " & codInicial & " AND " & codFinal
    strSql = strSql & " GROUP BY DV.intDocumentosEmitidos, DE.strDescricao, CO.strNome,DV.strInscricao, "
    strSql = strSql & " DV.intContribuinte, OC.strDescricao, DV.dtmDevolucao, CO.intBairro, BA.strDescricao "
    strSql = strSql & " ORDER BY BA.strDescricao "
strQuerryAnaliticoBairro = strSql

End Function

Private Function strQuerrySintetico() As String

'******************************************************************************************
' Data: 05/05/2003
' Alteração: - Deslocamento dos JOINS utilizados na cláusula FROM para a cláusula WHERE.
' Responsável: Everton Bianchini
'******************************************************************************************

Dim strSql    As String
Dim Inicial   As Date
Dim Final     As Date
    Inicial = CVDate(txt_Devolucao)
    Final = CVDate(txt_Ate)
    strSql = ""
    strSql = strSql & " SELECT BA.strdescricao, COUNT(DV.intdocumentosEmitidos) AS QTDDOCUMENTOS, "
    strSql = strSql & " COUNT(DV.intOcorrencia) AS QTDOCORRENCIAS, DE.strDescricao AS DOCUMENTO, "
    strSql = strSql & " OC.strDescricao AS OCORRENCIA "
'    strSql = strSql & " FROM " & gstrContribuinte & " CO " & "INNER JOIN " & gstrBairro & " BA "
    strSql = strSql & " FROM " & gstrContribuinte & " CO, " & gstrBairro & " BA "
'    strSql = strSql & " ON CO.intBairro = BA.PKID INNER JOIN " & gstrDevolucao & " DV "
    strSql = strSql & ", " & gstrDevolucao & " DV "
'    strSql = strSql & " ON CO.PKId = DV.intcontribuinte LEFT OUTER JOIN " & gstrDocumentoEmitido & " DE "
    strSql = strSql & ", " & gstrDocumentoEmitido & " DE "
'    strSql = strSql & " ON DV.intDocumentosEmitidos = DE.PKId "
'    strSql = strSql & " LEFT OUTER JOIN " & gstrOcorrencia & " OC "
    strSql = strSql & ", " & gstrOcorrencia & " OC "
'    strSql = strSql & " ON DV.intOcorrencia = OC.PKId "
    strSql = strSql & " WHERE DV.intOcorrencia = OC.PKId "
    
    strSql = strSql & " AND CO.intBairro = BA.PKID "
    strSql = strSql & " AND CO.PKId = DV.intcontribuinte "
    strSql = strSql & " AND DV.intDocumentosEmitidos " & strOUTJOracle & strOUTJSQLServer & "= DE.PKId "
    
    strSql = strSql & " AND DV.dtmDevolucao BETWEEN " & gstrConvDtParaSql(Inicial) & " AND " & gstrConvDtParaSql(Final)
    strSql = strSql & " GROUP BY BA.strDescricao, OC.strDescricao, DE.strDescricao "
    strSql = strSql & " ORDER BY BA.strDescricao "
strQuerrySintetico = strSql
End Function

Private Sub opt_Analitico_Click(Index As Integer)
    If Index = 0 Then
        txt_Inicial.Enabled = False
        txt_BairroInicial.Enabled = True
        TrocaCorObjeto txt_Inicial, True
        TrocaCorObjeto txt_BairroInicial, False
        txt_Final.Enabled = False
        txt_BairroFinal.Enabled = True
        TrocaCorObjeto txt_Final, True
        TrocaCorObjeto txt_BairroFinal, False
        txt_Inicial = ""
        txt_Final = ""
        txt_BairroInicial = ""
        txt_BairroFinal = ""

    Else
        txt_Inicial.Enabled = True
        txt_BairroInicial.Enabled = False
        TrocaCorObjeto txt_Inicial, False
        TrocaCorObjeto txt_BairroInicial, True
        txt_Final.Enabled = True
        txt_BairroFinal.Enabled = False
        TrocaCorObjeto txt_Final, False
        TrocaCorObjeto txt_BairroFinal, True
        txt_Inicial = ""
        txt_Final = ""
        txt_BairroInicial = ""
        txt_BairroFinal = ""
    End If

End Sub

Private Sub opt_Tipo_Click(Index As Integer)
    If Index = 0 Then
        Fra_Analitico.Enabled = True
        opt_Analitico(0).Value = True
        txt_Inicial.Enabled = False
        txt_BairroInicial.Enabled = True
        TrocaCorObjeto txt_Inicial, True
        TrocaCorObjeto txt_BairroInicial, False
        txt_Final.Enabled = False
        txt_BairroFinal.Enabled = True
        TrocaCorObjeto txt_Final, True
        TrocaCorObjeto txt_BairroFinal, False
    Else
        Fra_Analitico.Enabled = False
        opt_Analitico(0).Value = False
        opt_Analitico(1).Value = False
        txt_Inicial.Enabled = False
        txt_BairroInicial.Enabled = False
        TrocaCorObjeto txt_Inicial, True
        TrocaCorObjeto txt_BairroInicial, True
        txt_Final.Enabled = False
        txt_BairroFinal.Enabled = False
        TrocaCorObjeto txt_Final, True
        TrocaCorObjeto txt_BairroFinal, True
        txt_Inicial = ""
        txt_Final = ""
        txt_BairroInicial = ""
        txt_BairroFinal = ""
    End If
End Sub

Private Sub opt_Tipo_KeyPress(Index As Integer, KeyAscii As Integer)
CaracterValido KeyAscii, "A", opt_Tipo
End Sub

Private Sub opt_Analitico_KeyPress(Index As Integer, KeyAscii As Integer)
CaracterValido KeyAscii, "A", opt_Analitico
End Sub

Private Sub txt_Ate_GotFocus()
    MarcaCampo txt_Ate
End Sub

Private Sub txt_Ate_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "D", txt_Ate
End Sub

Private Sub txt_BairroFinal_GotFocus()
    MarcaCampo txt_BairroFinal
End Sub

Private Sub txt_BairroInicial_GotFocus()
    MarcaCampo txt_BairroInicial
End Sub

Private Sub txt_Devolucao_GotFocus()
    MarcaCampo txt_Devolucao
End Sub

Private Sub txt_Devolucao_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "D", txt_Devolucao
End Sub

Private Sub txt_Inicial_GotFocus()
    MarcaCampo txt_Inicial
End Sub

Private Sub txt_Inicial_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txt_Inicial
End Sub

Private Sub txt_bairroInicial_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txt_BairroInicial
End Sub

Private Sub txt_bairrofinal_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txt_BairroFinal
End Sub

Private Sub txt_Final_GotFocus()
    MarcaCampo txt_Final
End Sub

Private Sub txt_Final_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txt_Final
End Sub

