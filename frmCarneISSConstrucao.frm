VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmCarneISSConstrucao 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Carne de ISS Construção"
   ClientHeight    =   2190
   ClientLeft      =   3555
   ClientTop       =   2520
   ClientWidth     =   4530
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2190
   ScaleWidth      =   4530
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab tab_3dPasta 
      Height          =   2055
      Left            =   60
      TabIndex        =   8
      Top             =   60
      Width           =   4425
      _ExtentX        =   7805
      _ExtentY        =   3625
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Carne de ISS Construção "
      TabPicture(0)   =   "frmCarneISSConstrucao.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fra_DividaAtiva"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin VB.Frame fra_DividaAtiva 
         Height          =   1605
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   4215
         Begin VB.ComboBox cbo_IntExercicio 
            Height          =   315
            Left            =   2250
            TabIndex        =   6
            Top             =   960
            Width           =   1605
         End
         Begin VB.ComboBox cbo_intArea 
            Height          =   315
            Left            =   150
            TabIndex        =   4
            Top             =   960
            Width           =   1725
         End
         Begin VB.TextBox txtintExercicio 
            Height          =   300
            Left            =   2250
            MaxLength       =   4
            TabIndex        =   2
            Top             =   420
            Width           =   570
         End
         Begin MSMask.MaskEdBox mskstrInscricao 
            Height          =   300
            Left            =   150
            TabIndex        =   0
            Top             =   420
            Width           =   1710
            _ExtentX        =   3016
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   24
            PromptChar      =   " "
         End
         Begin VB.Label lbl_Emissao 
            AutoSize        =   -1  'True
            Caption         =   "Área"
            Height          =   195
            Index           =   1
            Left            =   150
            TabIndex        =   3
            Top             =   750
            Width           =   330
         End
         Begin VB.Label lbl_Emissao 
            AutoSize        =   -1  'True
            Caption         =   "Exercício"
            Height          =   195
            Index           =   0
            Left            =   2250
            TabIndex        =   1
            Top             =   180
            Width           =   675
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Vencimento da 1º Parcela"
            Height          =   195
            Left            =   2250
            TabIndex        =   5
            Top             =   750
            Width           =   1845
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Identificação"
            Height          =   195
            Left            =   150
            TabIndex        =   9
            Top             =   180
            Width           =   915
         End
      End
   End
End
Attribute VB_Name = "frmCarneISSConstrucao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim mblnSelecionou  As Boolean
    Dim lngpkID         As Long
    Dim blnAuto         As Boolean

Private Sub cbo_intArea_Click()
    Dim adoResultado As ADODB.Recordset
    Dim strSql As String
    
    If gstrItemData(cbo_intArea) > 0 Then
        cbo_IntExercicio.Clear
        'Vamos preencher a combo de vencimento
        strSql = ""
        strSql = strSql & "Select " & gstrTOPnSQLServer(1) & " "
        strSql = strSql & "LA.intExercicio, "
        strSql = strSql & "LV.Pkid, "
        strSql = strSql & "LV.Dtmdtvencimento "
        strSql = strSql & "From "
        strSql = strSql & gstrLancamentoAlfa & " LA, "
        strSql = strSql & gstrLancamentoValor & " LV, "
        strSql = strSql & gstrLanctoIssConstrucao & " LI "
        strSql = strSql & "Where "
        strSql = strSql & "LA.Pkid = LV.Intlancamentoalfa AND "
        strSql = strSql & "LA.Pkid = LI.Intlancamentoalfa AND "
        strSql = strSql & "LI.Pkid = " & gstrItemData(cbo_intArea) & " AND "
        strSql = strSql & "LV.Bitparcelavalida = 1 "
        strSql = strSql & "Order by "
        strSql = strSql & "LV.DTMDTVENCIMENTO"
        strSql = gstrTOPnOracle(strSql, 1)
        
        If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
            With adoResultado
                If .RecordCount >= 1 Then
                    txtintExercicio = gstrENulo(!intExercicio)
                    Do While Not .EOF
                        cbo_IntExercicio.AddItem gstrDataFormatada(!Dtmdtvencimento)
                        cbo_IntExercicio.ItemData(cbo_IntExercicio.NewIndex) = gstrENulo(!Pkid)
                        .MoveNext
                    Loop
                End If
            End With
        End If
        
        If cbo_IntExercicio.ListCount > 0 Then
            cbo_IntExercicio.ListIndex = 0
        End If

        
    End If
End Sub

Private Sub cbo_intArea_GotFocus()
    MarcaCampo cbo_intArea
End Sub

Private Sub cbo_intArea_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "V", cbo_intArea
End Sub

Private Sub cbo_IntExercicio_GotFocus()
    MarcaCampo cbo_IntExercicio
End Sub

Private Sub cbo_IntExercicio_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", cbo_IntExercicio
End Sub

Private Sub mskstrInscricao_Change()
    txtintExercicio.Text = ""
    cbo_intArea.Clear
    cbo_IntExercicio.Clear
End Sub

Private Sub mskstrInscricao_LostFocus()
    VerificaPreenchimento mskstrInscricao
End Sub

Private Sub mskstrInscricao_GotFocus()
    MarcaCampo mskstrInscricao
End Sub

Private Sub mskstrInscricao_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", mskstrInscricao
End Sub

Private Sub Form_Load()
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrSalvar, gstrDeletar
    mblnSelecionou = True
    VerificaMascaraInscricao
End Sub

Private Sub Form_Activate()
    gintCodSeguranca = 1243
    If mblnSelecionou Then
       HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrImprimir
       HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrSalvar, gstrDeletar, gstrAplicar
    Else
       HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrImprimir
    End If
End Sub

Private Sub Form_Deactivate()
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar
    mblnSelecionou = False
End Sub

Public Sub MantemForm(ByVal strModoOperacao As String)
    Select Case strModoOperacao
        Case gstrImprimir
            If blnDadosOK Then
                ImprimeRelatorio rptCapaCarneISSConstrucao, strQueryCarneISSConstrucao(lngpkID), "Carne de ISS Construção."
            End If
        Case gstrNovo
            LimpaISS
            mskstrInscricao.SetFocus
            blnAuto = False
        Case gstrPreencherLista
            PreencherListaDeOpcoes Me.ActiveControl
    End Select
End Sub

Private Function strQueryRelatorio() As String
    Dim strSql  As String
    Dim strSql1 As String
    
    strSql = ""
    strSql = strSql & "SELECT "
    strSql = strSql & "LI.Pkid, "
    strSql = strSql & "LA.Pkid as IntLancamentoAlfa, "
    strSql = strSql & gstrRIGHT("LA.strInscricao", gintRetornaTamanhoMascara(TYP_IMOBILIARIA)) & " strInscricao, "
    strSql = strSql & "LA.strComposicaoDaReceita, "
    strSql = strSql & "LA.intExercicio, "
    strSql = strSql & gstrRIGHT("LA.strNumeroAviso", gintLenNumAviso) & " strNumeroAviso, "
    strSql = strSql & gstrRIGHT("LA.strEmissao", gintLenEmissao) & " strEmissao, "
    strSql = strSql & "LA.strNomeProprietario, "
    strSql = strSql & "LA.Strpromissario, "
    strSql = strSql & "LA.strInscricao, "
    strSql = strSql & "LA.strLogradouro, "
    strSql = strSql & "LA.strNumero, "
    strSql = strSql & "LA.strComplemento, "
    strSql = strSql & "LA.strBairro, "
    strSql = strSql & "LA.strMunicipio, "
    strSql = strSql & "LA.strUf, "
    strSql = strSql & "LA.intCep, "
    strSql = strSql & "LA.strLogradouroC, "
    strSql = strSql & "LA.strNumeroC, "
    strSql = strSql & "LA.strComplementoC, "
    strSql = strSql & "LA.strBairroC, "
    strSql = strSql & "LA.strMunicipioC, "
    strSql = strSql & "LA.strUfC, "
    strSql = strSql & "LA.intCepC, "
    strSql = strSql & "LA.Strindexador, "
    strSql = strSql & "LA.Dblvlindexador, "
    strSql = strSql & "LI.strCodigoProcesso" & strCONCAT & "'/'" & strCONCAT & "LI.intExercicioProcesso" & strCONCAT & "'-'" & strCONCAT & "LI.bitDigitoProcesso as strProcesso, "
    strSql = strSql & "LI.strObservacoes, "
    strSql = strSql & "LI.dtmLancamento, "
    strSql = strSql & "LV.TotParcela, "
    strSql = strSql & "LV1.dbl1valor, "
    strSql = strSql & "LV1.dtmdtvencimentoParcela, "
    strSql = strSql & "LIC1.dblPorcDemolicao, "
    strSql = strSql & "LIC1.dblarealancada, "
    strSql = strSql & "LIC1.dblvalorm2, "
    strSql = strSql & "LIC1.dblvalorservico, "
    strSql = strSql & "LIC1.dblaliquotaiss, "
    strSql = strSql & "LIC1.dblvalorlancto, "
    strSql = strSql & "LIC1.dblvalorabatido, "
    strSql = strSql & "LIC1.dblSaldo "
    strSql = strSql & "FROM "
    strSql = strSql & gstrLancamentoAlfa & " LA, "
    strSql = strSql & gstrLanctoIssConstrucao & " LI, "
    
    'Select para trazer somatória da tabela de prédios ISS
    strSql = strSql & "(Select "
    strSql = strSql & "Sum(LIC.dblPorcDemolicao) as dblPorcDemolicao, "
    strSql = strSql & "Sum(LIC.dblarealancada) as dblarealancada, "
    strSql = strSql & "Sum(LIC.dblvalorm2) as dblvalorm2, "
    strSql = strSql & "Sum(LIC.dblvalorservico) as dblvalorservico, "
    strSql = strSql & "LIC.dblaliquotaiss, "
    strSql = strSql & "Sum(LIC.dblvalorlancto) as dblvalorlancto, "
    strSql = strSql & "Sum(LIC.dblvalorabatido) as dblvalorabatido, "
    'strSQL = strSQL & "(Sum(LIC.dblvalorlancto)  -  Sum(LIC.dblvalorabatido)) dblSaldo "
    strSql = strSql & "((CASE WHEN SUM(LIC.dblValorLancto) IS NULL THEN 0 ELSE SUM(LIC.dblValorLancto) END) - "
    strSql = strSql & "(CASE WHEN SUM(LIC.dblValorAbatido)IS NULL THEN 0 ELSE SUM(LIC.dblValorAbatido) END)) dblSaldo "
    strSql = strSql & "From " & gstrLanctoIssConstrucao & " LI," & gstrLanctoIssConstrucaoPredios & " LIC "
    strSql = strSql & "Where LI.Pkid" & strOUTJOracle & "=" & strOUTJSQLServer & "LIC.INTLANCTOISSCONSTRUCAO AND LI.Intlancamentoalfa = " & lngpkID & " Group by LIC.dblaliquotaiss) LIC1, "
    
    'Select para trazer Qtde de parcelas
    strSql = strSql & "(Select Count(intParcela) as TotParcela From "
    strSql = strSql & gstrLancamentoValor & " Where Intlancamentoalfa =" & lngpkID & " ) LV, "
    
    'Select para trazer 1º Vencimento e 1º Valor de parcela
    strSql1 = ""
    strSql1 = strSql1 & "Select " & gstrTOPnSQLServer(1) & "dblvalor as dbl1valor, dtmdtvencimento as dtmdtvencimentoParcela From "
    strSql1 = strSql1 & gstrLancamentoValor & " Where intLancamentoAlfa = " & lngpkID & " Order by dtmdtvencimento"
    strSql1 = "(" & gstrTOPnOracle(strSql1, 1) & ") LV1 "
    
    strSql = strSql & strSql1

    strSql = strSql & "WHERE LA.Pkid = " & lngpkID & " AND LI.intLancamentoAlfa = LA.Pkid "
    
    strQueryRelatorio = strSql

End Function

Private Function blnDadosOK() As Boolean
    blnDadosOK = False
    If Trim(mskstrInscricao.Text) = "" Then
        ExibeMensagem "É necessário preencher o campo de Inscrição."
        mskstrInscricao.SetFocus
        Exit Function
    ElseIf Trim(txtintExercicio.Text) = "" Then
        ExibeMensagem "É necessário preencher o campo de Exercício."
        txtintExercicio.SetFocus
        Exit Function
    ElseIf gstrItemData(cbo_intArea) <= 0 Then
        ExibeMensagem "É necessário selecionar alguma área."
        cbo_intArea.SetFocus
        Exit Function
    ElseIf gstrItemData(cbo_IntExercicio) <= 0 Then
        ExibeMensagem "É necessário selecionar algum vencimento."
        cbo_IntExercicio.SetFocus
        Exit Function
    ElseIf Not blnInscricaoOk Then
        ExibeMensagem "Inscrição não está válida."
        mskstrInscricao.SetFocus
        Exit Function
    End If
    blnDadosOK = True
End Function

Private Sub VerificaMascaraInscricao()
    Dim strSql As String
    Dim adoResultado As ADODB.Recordset
    Dim strMascara   As String
    
    strMascara = ""
    
    strSql = ""
    strSql = strSql & "Select * From " & gstrCampoDeInscricao & " "
    strSql = strSql & "Where intTipoDeInscricao = " & TYP_IMOBILIARIA
    strSql = strSql & "Order By intSequencia"
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        With adoResultado
            Do While Not .EOF
                strMascara = strMascara & String(!intTamanho, "#") & gstrVerificaCampoNulo(!strSeparador)
                .MoveNext
            Loop
        End With
    End If
    mskstrInscricao.Mask = strMascara
End Sub

Private Function strQueryInscricao() As String
    Dim strSql As String
    
    
    strQueryInscricao = strSql
End Function

Private Function strQueryExercicio() As String
    Dim strSql As String
    
    strSql = ""
    strSql = strSql & "SELECT"
    strSql = strSql & " LA.Pkid,"
    strSql = strSql & " LA.IntExercicio "
    strSql = strSql & " FROM "
    strSql = strSql & gstrLancamentoAlfa & " LA, "
    strSql = strSql & gstrLanctoIssConstrucao & " LI "
    strSql = strSql & "WHERE "
    strSql = strSql & "LI.intLancamentoAlfa = LA.Pkid "
    strSql = strSql & "LI.intLancamentoAlfa = LA.Pkid "
    strSql = strSql & "Order By LA.strInscricao"
    
    strQueryExercicio = strSql
End Function

Private Sub VerificaPreenchimento(Optional strInscricao As String, Optional intExercicio As Integer)
    Dim adoResultado    As ADODB.Recordset
    Dim strSql          As String
    Dim lngPkidAlfa     As String
    
    Dim strExercicio    As String
    
    If strInscricao = "" Then Exit Sub
     
    strSql = ""
    strSql = strSql & "SELECT"
    strSql = strSql & " LA.Pkid,"
    strSql = strSql & " LA.strInscricao, "
    strSql = strSql & " LA.intExercicio "
    strSql = strSql & " FROM "
    strSql = strSql & gstrLancamentoAlfa & " LA, "
    strSql = strSql & gstrLanctoIssConstrucao & " LI "
    strSql = strSql & "WHERE "
    strSql = strSql & "LI.intLancamentoAlfa = LA.Pkid AND "
    strSql = strSql & "LA.strinscricao ='" & (String(gintLenInscricao - Len(Trim(mskstrInscricao.Text)), "0") & Trim(mskstrInscricao.Text)) & "' "
    strSql = strSql & "Order By LA.strInscricao"
    
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        If Not adoResultado.EOF And Not adoResultado.BOF Then
            If adoResultado.RecordCount = 1 Then
                blnAuto = True
                lngPkidAlfa = CLng(gstrENulo(adoResultado!Pkid))
                strExercicio = CInt(gstrENulo(adoResultado!intExercicio))
                txtintExercicio.Text = strExercicio
                
                'Vamos preencher a combo de vencimento
                strSql = ""
                strSql = strSql & "Select " & gstrTOPnSQLServer(1) & " "
                strSql = strSql & "LV.Pkid, "
                strSql = strSql & "LV.Dtmdtvencimento "
                strSql = strSql & "From "
                strSql = strSql & gstrLancamentoValor & " LV "
                strSql = strSql & "where "
                strSql = strSql & "LV.intLancamentoAlfa = " & lngPkidAlfa & " AND "
                strSql = strSql & "LV.Bitparcelavalida = 1 "
                strSql = strSql & "Order by "
                strSql = strSql & "LV.DTMDTVENCIMENTO"
                strSql = gstrTOPnOracle(strSql, 1)
                
                Set gobjBanco = New clsBanco
                If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
                    With adoResultado
                        If .RecordCount >= 1 Then
                            Do While Not .EOF
                                cbo_IntExercicio.AddItem gstrDataFormatada(!Dtmdtvencimento)
                                cbo_IntExercicio.ItemData(cbo_IntExercicio.NewIndex) = gstrENulo(!Pkid)
                                .MoveNext
                            Loop
                        End If
                    End With
                End If
                If cbo_IntExercicio.ListCount > 0 Then
                    cbo_IntExercicio.ListIndex = 0
                End If
                
                'Vamos preencher a combo de Área
                strSql = ""
                strSql = strSql & "Select "
                strSql = strSql & "LI.Pkid, "
                strSql = strSql & "Sum(LIP.DBLAREALANCADA) as dblValor  "
                strSql = strSql & "From "
                strSql = strSql & gstrLancamentoAlfa & " LA, "
                strSql = strSql & gstrLanctoIssConstrucao & " LI, "
                strSql = strSql & gstrLanctoIssConstrucaoPredios & " LIP "
                strSql = strSql & "Where "
                strSql = strSql & "La.Pkid = LI.Intlancamentoalfa  AND "
                strSql = strSql & "LI.Pkid = LIP.INTLANCTOISSCONSTRUCAO AND "
                strSql = strSql & "LA.Pkid = " & lngPkidAlfa & " "
                strSql = strSql & "Group by "
                strSql = strSql & "LI.Pkid, "
                strSql = strSql & "LA.Strnumeroaviso "
                Set gobjBanco = New clsBanco
                If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
                    With adoResultado
                        If .RecordCount >= 1 Then
                            Do While Not .EOF
                                cbo_intArea.AddItem gstrConvVrDoSql(!DBLVALOR, 2)
                                cbo_intArea.ItemData(cbo_intArea.NewIndex) = gstrENulo(!Pkid)
                                .MoveNext
                            Loop
                        End If
                    End With
                End If
                If cbo_intArea.ListCount > 0 Then
                    cbo_intArea.ListIndex = 0
                End If
            Else
'                'Vamos preencher a combo de Área
'                If Len(Trim(mskstrInscricao.Text)) <> 8 Then Exit Sub
                blnAuto = False
'                strSql = ""
'                strSql = strSql & "Select "
'                strSql = strSql & "LI.Pkid, "
'                strSql = strSql & "Sum(LIP.DBLAREALANCADA) as dblValor  "
'                strSql = strSql & "From "
'                strSql = strSql & gstrLancamentoAlfa & " LA, "
'                strSql = strSql & gstrLanctoIssConstrucao & " LI, "
'                strSql = strSql & gstrLanctoIssConstrucaoPredios & " LIP "
'                strSql = strSql & "Where "
'                strSql = strSql & "La.Pkid = LI.Intlancamentoalfa  AND "
'                strSql = strSql & "LI.Pkid = LIP.INTLANCTOISSCONSTRUCAO AND "
'                strSql = strSql & "LA.strInscricao = '" & mskstrInscricao.Text & "' "
'                If Len(Trim(txtintExercicio)) = 4 Then
'                    strSql = strSql & "AND LA.IntExercicio = " & txtintExercicio.Text & " "
'                End If
'                strSql = strSql & "Group by "
'                strSql = strSql & "LI.Pkid, "
'                strSql = strSql & "LA.Strnumeroaviso "
'
'                Set gobjBanco = New clsBanco
'                If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
'                    With adoResultado
'                        If .RecordCount >= 1 Then
'                            Do While Not .EOF
'                                cbo_intArea.AddItem gstrConvVrDoSql(!DBLVALOR, 2)
'                                cbo_intArea.ItemData(cbo_intArea.NewIndex) = gstrENulo(!Pkid)
'                                .MoveNext
'                            Loop
'                        End If
'                    End With
'                End If
                
            End If
        End If
    End If
End Sub

Private Sub LimpaISS()
    mskstrInscricao.Text = ""
    txtintExercicio.Text = ""
    cbo_intArea.Clear
    cbo_IntExercicio.Clear
End Sub

Private Sub txtintExercicio_GotFocus()
    MarcaCampo txtintExercicio
End Sub

Private Sub txtintExercicio_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtintExercicio
End Sub

Private Sub txtintexercicio_LostFocus()
    Dim adoResultado    As ADODB.Recordset
    Dim strSql          As String
    
    If (cbo_intArea.ListIndex >= 0) And (cbo_IntExercicio.ListIndex >= 0) Then Exit Sub
    
    cbo_intArea.Clear
    cbo_IntExercicio.Clear

    If Len(Trim(mskstrInscricao.Text)) <> 8 Then Exit Sub
    If blnAuto = True Then Exit Sub
    If Not Len(Trim(txtintExercicio)) = 4 Then Exit Sub
    
    strSql = ""
    strSql = strSql & "Select "
    strSql = strSql & "LI.Pkid, "
    strSql = strSql & "Sum(LIP.DBLAREALANCADA) as dblValor  "
    strSql = strSql & "From "
    strSql = strSql & gstrLancamentoAlfa & " LA, "
    strSql = strSql & gstrLanctoIssConstrucao & " LI, "
    strSql = strSql & gstrLanctoIssConstrucaoPredios & " LIP "
    strSql = strSql & "Where "
    strSql = strSql & "La.Pkid = LI.Intlancamentoalfa  AND "
    strSql = strSql & "LI.Pkid = LIP.INTLANCTOISSCONSTRUCAO AND "
    strSql = strSql & "LA.strinscricao ='" & (String(gintLenInscricao - Len(Trim(mskstrInscricao.Text)), "0") & Trim(mskstrInscricao.Text)) & "' "
    If Len(Trim(txtintExercicio)) = 4 Then
        strSql = strSql & "AND LA.IntExercicio = " & txtintExercicio.Text & " "
    End If
    strSql = strSql & "Group by "
    strSql = strSql & "LI.Pkid, "
    strSql = strSql & "LA.Strnumeroaviso "
    
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
    With adoResultado
        If .RecordCount >= 1 Then
            Do While Not .EOF
                cbo_intArea.AddItem gstrConvVrDoSql(!DBLVALOR, 2)
                cbo_intArea.ItemData(cbo_intArea.NewIndex) = gstrENulo(!Pkid)
                .MoveNext
            Loop
        End If
    End With
    End If
End Sub

Private Function blnInscricaoOk() As Boolean
    Dim adoResultado    As ADODB.Recordset
    Dim strSql          As String

    If Len(Trim(mskstrInscricao.Text)) <> 8 Then Exit Function
    
    blnInscricaoOk = False
    
    strSql = ""
    strSql = strSql & "Select "
    strSql = strSql & "Distinct LA.Pkid "
    strSql = strSql & "From "
    strSql = strSql & gstrLancamentoAlfa & " LA, "
    strSql = strSql & gstrLancamentoValor & " LV, "
    strSql = strSql & gstrLanctoIssConstrucao & " LI, "
    strSql = strSql & gstrLanctoIssConstrucaoPredios & " LIP "
    strSql = strSql & "Where "
    strSql = strSql & "LA.Pkid = LV.Intlancamentoalfa AND "
    strSql = strSql & "LA.Pkid = LI.Intlancamentoalfa AND "
    strSql = strSql & "LI.Pkid = " & gstrItemData(cbo_intArea) & " AND "
    strSql = strSql & "LA.strinscricao ='" & (String(gintLenInscricao - Len(Trim(mskstrInscricao.Text)), "0") & Trim(mskstrInscricao.Text)) & "' AND "
    strSql = strSql & "LA.Intexercicio = " & txtintExercicio

    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        With adoResultado
            If .RecordCount >= 1 Then
                blnInscricaoOk = True
                lngpkID = gstrENulo(!Pkid)
            End If
        End With
    End If
    
End Function
