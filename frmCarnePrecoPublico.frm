VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCarnePrecoPublico 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Carne de Preço Público"
   ClientHeight    =   2160
   ClientLeft      =   3555
   ClientTop       =   2520
   ClientWidth     =   3225
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2160
   ScaleWidth      =   3225
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab tab_3dPasta 
      Height          =   2055
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   3165
      _ExtentX        =   5583
      _ExtentY        =   3625
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "2º Via de Preço Público"
      TabPicture(0)   =   "frmCarnePrecoPublico.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fra_DividaAtiva"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin VB.Frame fra_DividaAtiva 
         Height          =   1605
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   2925
         Begin VB.TextBox txtstrInscricao 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "#.##0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
            EndProperty
            Height          =   285
            Left            =   150
            MaxLength       =   20
            TabIndex        =   2
            Top             =   450
            Width           =   2100
         End
         Begin VB.TextBox txtstrNumeroAviso 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   150
            MaxLength       =   10
            TabIndex        =   3
            Top             =   990
            Width           =   1260
         End
         Begin VB.Label lblstrInscricao 
            AutoSize        =   -1  'True
            Caption         =   "Inscrição Cadastral"
            Height          =   195
            Left            =   150
            TabIndex        =   5
            Top             =   195
            Width           =   1350
         End
         Begin VB.Label lbl_Emissao 
            AutoSize        =   -1  'True
            Caption         =   "Aviso"
            Height          =   195
            Index           =   0
            Left            =   150
            TabIndex        =   4
            Top             =   750
            Width           =   390
         End
      End
   End
End
Attribute VB_Name = "frmCarnePrecoPublico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim mblnSelecionou  As Boolean
    Dim vetGuiaPrecoPublico() As String
    Dim blnFebraban           As Boolean

Private Sub txtstrInscricao_Change()
    txtstrNumeroAviso.Text = ""
End Sub

Private Sub txtstrInscricao_GotFocus()
    MarcaCampo txtstrInscricao
End Sub

Private Sub txtstrInscricao_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtstrInscricao
End Sub

Private Sub Form_Load()
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrSalvar, gstrDeletar
    mblnSelecionou = True
    
End Sub

Private Sub Form_Activate()
    gintCodSeguranca = 1268
    If mblnSelecionou Then
       HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrImprimir
       HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrSalvar, gstrDeletar, gstrAplicar
    Else
       HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrImprimir
    End If
End Sub

Private Sub Form_Deactivate()
    HabilitaDesabilitaBotao1 True, gstrMnuArquivo
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    HabilitaDesabilitaBotao1 True, gstrMnuArquivo
    mblnSelecionou = False
End Sub

Public Sub MantemForm(ByVal strModoOperacao As String)
    Select Case strModoOperacao
        Case gstrImprimir
            If blnDadosOk Then
                If blnFebraban Then
                    ImprimeRelatorioPorArray rptGuiaPrecoPublico, vetGuiaPrecoPublico, "Guia de Arrecadação"
                Else
                    ImprimeRelatorioPorArray rptGuiaFichaPrecoPublico, vetGuiaPrecoPublico, "Guia de Arrecadação"
                End If
            End If
        Case gstrNovo
            LimpaPPublico
            txtstrInscricao.SetFocus
        Case gstrPreencherLista
            PreencherListaDeOpcoes Me.ActiveControl
    End Select
End Sub

Private Function blnDadosOk() As Boolean
    blnDadosOk = False
    If Trim(txtstrInscricao.Text) = "" And Trim(txtstrNumeroAviso.Text) = "" Then
        ExibeMensagem "É necessário preencher algum dos dos campos."
        txtstrInscricao.SetFocus
        Exit Function
    ElseIf Not bnlPreencheVetor Then
        Exit Function
    End If
    blnDadosOk = True
End Function

Private Function bnlPreencheVetor() As Boolean
    Dim adoResultado        As ADODB.Recordset
    Dim adoResultado1       As ADODB.Recordset
    Dim adoRec              As ADODB.Recordset
    Dim strSql              As String
    Dim strNumeroBoleto     As String
    Dim strsigla            As String
    Dim lngContaBancaria    As Long
    Dim strNossoNumero      As String
    
    ReDim vetGuiaPrecoPublico(31, 0)
    bnlPreencheVetor = False
    strsigla = ""
    
    strSql = ""
    strSql = strSql & "Select "
    strSql = strSql & "LA.Pkid, LA.intExercicio, LA.strComposicaoDaReceita, "
    strSql = strSql & "LPP.Strinscricao, "
    strSql = strSql & "LA.strNumeroAviso  strNumeroAviso, "
    strSql = strSql & " LPP.strCodigo " & strCONCAT & "'/'" & strCONCAT
    strSql = strSql & gstrCONVERT(CDT_VARCHAR, "LPP.intExercicio ") & strCONCAT & " '-' " & strCONCAT
    strSql = strSql & gstrCONVERT(CDT_NVARCHAR, "LPP.bitDigito") & " AS strProcesso, "
    strSql = strSql & "LA.strnomeproprietario, "
    strSql = strSql & "LA.strlogradouro, "
    strSql = strSql & "LA.strnumero, "
    strSql = strSql & "LA.strcomplemento, "
    strSql = strSql & "LA.strbairro, "
    strSql = strSql & "LA.strmunicipio, "
    strSql = strSql & "LA.intCep, "
    strSql = strSql & "LA.struf, "
    strSql = strSql & "LPP.Strhistorico, "
    strSql = strSql & "LPP.Dblvalor, "
    strSql = strSql & "LPP.Dblmulta, "
    strSql = strSql & "LPP.Dblcorrecaomonet, "
    strSql = strSql & "LPP.dblJuros, "
    strSql = strSql & "LV.intParcela, LV.bitParcelaValida, "
    strSql = strSql & "G.INTNUMERO, "
    strSql = strSql & "US.strLogin, "
    strSql = strSql & "LA.Intexercicio, "
    strSql = strSql & "LA.intUtilizacao, "
    strSql = strSql & "LA.intComposicaoDaReceita, "
    strSql = strSql & "LPP.Dtmdtvencimento, "
    strSql = strSql & "G.DTMDTEMISSAO, "
    strSql = strSql & "G.Strcodbarra "
    strSql = strSql & "From "
    strSql = strSql & gstrLancamentoAlfa & " LA, "
    strSql = strSql & gstrLancamentoPPublico & " LPP, "
    strSql = strSql & gstrLancamentoValor & " LV, "
    strSql = strSql & gstrLancamentoGuias & " LG, "
    strSql = strSql & gstrGuias & " G, "
    strSql = strSql & gstrUsuarios & " US "
    strSql = strSql & "Where "
    strSql = strSql & "LA.Pkid = LPP.Intlancamentoalfa AND "
    strSql = strSql & "LA.Pkid = LV.Intlancamentoalfa AND "
    strSql = strSql & "LV.Pkid = LG.Intlancamentovalor AND "
    strSql = strSql & "G.Pkid = LG.Intguias AND "
    strSql = strSql & "LPP.lngCodUsr = US.Pkid  "
    strSql = strSql & IIf(Trim(txtstrInscricao) <> "", " AND LA.Strinscricao = '" & String(gintLenInscricao - Len(Trim(txtstrInscricao.Text)), "0") & UCase(txtstrInscricao.Text) & "' ", "")
    strSql = strSql & IIf(Trim(txtstrNumeroAviso) <> "", " AND LA.strNumeroAviso = '" & UCase(String(gintLenNumAviso - Len(txtstrNumeroAviso), "0") & txtstrNumeroAviso.Text) & "' ", "")
    
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 10, adoResultado) Then
        With adoResultado
            If .RecordCount >= 1 Then
                strSql = ""
                strSql = strSql & "Select R.strSigla From "
                strSql = strSql & gstrLancamentoValor & " LV,"
                strSql = strSql & gstrLancamentoReceita & " LR, "
                strSql = strSql & gstrReceita & " R "
                strSql = strSql & "Where LV.pkid = LR.Intlancamentovalor AND R.Pkid = LR.Intreceita AND "
                strSql = strSql & "LV.intLancamentoAlfa = " & gstrENulo(!Pkid)
                Set gobjBanco = New clsBanco
                If gobjBanco.CriaADO(strSql, 5, adoResultado1) Then
                    If .RecordCount >= 1 Then
                        Do While Not adoResultado1.EOF
                            strsigla = strsigla & gstrENulo(adoResultado1!strsigla) & " / "
                            adoResultado1.MoveNext
                        Loop
                    Else
                        ExibeMensagem "Não foi possível localizar as receitas desta guia."
                        Exit Function
                    End If
                End If
            
                'Vamos obter a conta bancaria da composicao
                strSql = "Select PA.intContaBancaria From " & gstrParametroAtualizacao & " PA Where PA.intComposicaoReceita  = " & !intComposicaoDaReceita & " And PA.intExercicio = " & !intExercicio
                
                If gobjBanco.CriaADO(strSql, 10, adoRec) Then
                
                    With adoRec
                        If Not (.BOF And .EOF) Then
                            lngContaBancaria = IIf(IsNull(adoRec("intContaBancaria").Value), 0, adoRec("intContaBancaria").Value)
                        End If
                    End With
                    
                End If
                
                blnFebraban = lngContaBancaria = 0
                
                'Vamos definir a linha digitavel
                strNumeroBoleto = gstrMontaLinhaDigitavel(IIf(blnFebraban, FEBRABAN, FICHA_COMPENSACAO), !strCodBarra)
                'Vamos definir o nosso numero
                If Not blnFebraban Then
                    strNossoNumero = gstrMontaNossoNumero(lngContaBancaria, !INTNUMERO)
                End If
            
                'Vamos preencher o vetor
                vetGuiaPrecoPublico(0, 0) = gstrENulo(!INTNUMERO) & "/" & gstrENulo(!intExercicio)
                vetGuiaPrecoPublico(1, 0) = gstrDataFormatada(gstrENulo(!Dtmdtvencimento))
                vetGuiaPrecoPublico(2, 0) = gstrENulo(!strnomeproprietario)
                vetGuiaPrecoPublico(3, 0) = gstrENulo(!strLogradouro) & IIf(Trim(gstrENulo(!strNumero)) <> "", "," & gstrENulo(!strNumero), "") & IIf(Trim(gstrENulo(!STRCOMPLEMENTO)) <> "", " - " & gstrENulo(!STRCOMPLEMENTO), "")
                vetGuiaPrecoPublico(4, 0) = gstrENulo(!strBairro)
                vetGuiaPrecoPublico(5, 0) = ""
                vetGuiaPrecoPublico(6, 0) = ""
                If IsNull(!strInscricao) = True Or Trim(!strInscricao) = "" Then
                    vetGuiaPrecoPublico(7, 0) = gstrENulo(!INTNUMERO)
                Else
                    vetGuiaPrecoPublico(7, 0) = gstrFormataInscricao(Right(gstrENulo(!strInscricao), gintRetornaTamanhoMascara(gstrENulo(!intUtilizacao))), gstrENulo(!intUtilizacao))
                End If
                vetGuiaPrecoPublico(8, 0) = gstrENulo(!strNumeroAviso)
                vetGuiaPrecoPublico(9, 0) = strsigla
                vetGuiaPrecoPublico(10, 0) = gstrENulo(!strHistorico)
                vetGuiaPrecoPublico(11, 0) = gstrConvVrDoSql(gstrENulo(!dblValor), , , True)
                vetGuiaPrecoPublico(12, 0) = gstrConvVrDoSql(gstrENulo(!Dblcorrecaomonet), , , True)
                vetGuiaPrecoPublico(13, 0) = gstrConvVrDoSql(gstrENulo(!dblMulta), , , True)
                vetGuiaPrecoPublico(14, 0) = gstrConvVrDoSql(gstrENulo(!dblJuros), , , True)
                vetGuiaPrecoPublico(15, 0) = gstrConvVrDoSql(gstrENulo(!dblValor), , , True)
                vetGuiaPrecoPublico(16, 0) = gstrDataFormatada(gstrENulo(!dtmDtEmissao))
                vetGuiaPrecoPublico(17, 0) = gstrENulo(!strLogin)
                vetGuiaPrecoPublico(18, 0) = gstrDataFormatada(gstrENulo(!Dtmdtvencimento))
                vetGuiaPrecoPublico(19, 0) = strNumeroBoleto
                vetGuiaPrecoPublico(20, 0) = gstrENulo(!strCodBarra)
                vetGuiaPrecoPublico(21, 0) = gstrENulo(!STRMUNICIPIO)
                vetGuiaPrecoPublico(22, 0) = gstrENulo(!STRUF)
                vetGuiaPrecoPublico(23, 0) = gstrENulo(!strProcesso)
                vetGuiaPrecoPublico(24, 0) = gstrENulo(!INTCEP)
                vetGuiaPrecoPublico(25, 0) = lngContaBancaria
                vetGuiaPrecoPublico(26, 0) = strNossoNumero
                vetGuiaPrecoPublico(27, 0) = gstrENulo(!strComposicaoDaReceita)
                vetGuiaPrecoPublico(28, 0) = gstrENulo(!intExercicio)
                vetGuiaPrecoPublico(29, 0) = gstrENulo(!intParcela)
                vetGuiaPrecoPublico(30, 0) = gstrENulo(!bitParcelaValida)
                vetGuiaPrecoPublico(31, 0) = gstrENulo(!Pkid)
                
                bnlPreencheVetor = True
            Else
                ExibeMensagem "Não existe guia com esses parâmetros."
                Exit Function
            End If
        End With
    End If
    
    
End Function

Private Sub LimpaPPublico()
    txtstrInscricao.Text = ""
    txtstrNumeroAviso.Text = ""
End Sub

Private Sub txtstrNumeroAviso_GotFocus()
    MarcaCampo txtstrNumeroAviso
End Sub

Private Sub txtstrNumeroAviso_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtstrNumeroAviso
End Sub

