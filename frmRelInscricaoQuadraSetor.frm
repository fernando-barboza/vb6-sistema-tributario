VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmRelInscricaoQuadraSetor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Inscrição"
   ClientHeight    =   2175
   ClientLeft      =   4650
   ClientTop       =   3135
   ClientWidth     =   6540
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2175
   ScaleWidth      =   6540
   Begin TabDlg.SSTab tab_3dPasta 
      Height          =   2145
      Left            =   60
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   30
      Width           =   6405
      _ExtentX        =   11298
      _ExtentY        =   3784
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Bairros"
      TabPicture(0)   =   "frmRelInscricaoQuadraSetor.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraLogradouro"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraQuadraSetor"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      Begin VB.Frame fraQuadraSetor 
         Caption         =   "Procura por "
         Height          =   765
         Left            =   150
         TabIndex        =   3
         Top             =   360
         Width           =   6045
         Begin VB.TextBox txtquadra 
            Height          =   285
            Left            =   4290
            MaxLength       =   3
            TabIndex        =   5
            Top             =   300
            Width           =   645
         End
         Begin VB.TextBox txtSetor 
            Height          =   285
            Left            =   1920
            MaxLength       =   2
            TabIndex        =   4
            Top             =   300
            Width           =   645
         End
         Begin VB.Label lblquadra 
            AutoSize        =   -1  'True
            Caption         =   "Quadra"
            Height          =   195
            Left            =   3690
            TabIndex        =   7
            Top             =   390
            Width           =   525
         End
         Begin VB.Label lblsetor 
            AutoSize        =   -1  'True
            Caption         =   "Setor"
            Height          =   195
            Left            =   1500
            TabIndex        =   6
            Top             =   390
            Width           =   375
         End
      End
      Begin VB.Frame fraLogradouro 
         Caption         =   "Procura por"
         Height          =   795
         Left            =   150
         TabIndex        =   1
         Top             =   1200
         Width           =   6045
         Begin MSDataListLib.DataCombo dbc_intlogradouro 
            Height          =   315
            Left            =   1050
            TabIndex        =   8
            Top             =   330
            Width           =   4875
            _ExtentX        =   8599
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin VB.Label lbllogradouro 
            AutoSize        =   -1  'True
            Caption         =   "Logradouro"
            Height          =   195
            Left            =   180
            TabIndex        =   2
            Top             =   450
            Width           =   810
         End
      End
   End
End
Attribute VB_Name = "frmRelInscricaoQuadraSetor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub dbc_intlogradouro_GotFocus()
    MarcaCampo dbc_intlogradouro
End Sub

Private Sub dbc_intlogradouro_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbc_intlogradouro, True
End Sub

Private Sub txtquadra_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtquadra
End Sub

Private Sub txtquadra_Gotfocus()
    MarcaCampo txtquadra
End Sub

Private Sub txtsetor_Keypress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtSetor
End Sub

Private Sub txtsetor_gotfocus()
    MarcaCampo txtSetor
End Sub

Private Sub Form_Activate()
    gintCodSeguranca = 1169
    HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrImprimir
End Sub

Private Sub Form_Load()
    dbc_intlogradouro.Tag = strQueryLogradouro & ";L.strDescricao"
    fraQuadraSetor_Click
End Sub

Private Sub fraLogradouro_Click()
    txtquadra.Text = Empty
    txtSetor.Text = Empty
    lblsetor.Enabled = False
    txtquadra.Enabled = False
    lblquadra.Enabled = False
    txtSetor.Enabled = False
    lbllogradouro.Enabled = True
    dbc_intlogradouro.Enabled = True
End Sub

Private Sub fraQuadraSetor_Click()
    dbc_intlogradouro.Text = Empty
    lblsetor.Enabled = True
    txtquadra.Enabled = True
    lblquadra.Enabled = True
    txtSetor.Enabled = True
    lbllogradouro.Enabled = False
    dbc_intlogradouro.Enabled = False
    
End Sub


Private Function strQueryLogradouro() As String
    Dim strSQL As String
    
'    strSQL = ""
'    strSQL = strSQL & "Select Pkid, strDescricao FROM " & gstrLogradouro & " Order by strDescricao "
    strSQL = ""
    
    strSQL = strSQL & "SELECT L.PKId as Pkid, "
    strSQL = strSQL & " RTRIM(LTRIM(L.strDescricao)) " & strCONCAT & gstrISNULL("TL.strSigla", "''", "', '") & strCONCAT & " RTRIM(LTRIM(" & gstrISNULL("TL.strSigla", "''") & _
             strCONCAT & gstrISNULL("U.strDescricao", "' '", "', '") & strCONCAT & gstrISNULL("U.strDescricao", "''") & ")) " & strCONCAT & "' -> '" & strCONCAT & gstrISNULL("BA.strDescricao", "''") & " AS strdescricao "
    strSQL = strSQL & "FROM " & gstrLogradouro & " L, "
    strSQL = strSQL & gstrTituloLogradouro & " U, "
    strSQL = strSQL & gstrTipoLogradouro & " TL, "
    strSQL = strSQL & gstrBairro & " BA "
    strSQL = strSQL & " WHERE L.intTituloLogradouro " & strOUTJSQLServer & "= U.PKId " & strOUTJOracle
    strSQL = strSQL & " AND L.Dtmdtexclusao is null "
    strSQL = strSQL & " AND L.intTipoLogradouro " & strOUTJSQLServer & "= TL.PKId " & strOUTJOracle
    strSQL = strSQL & " AND L.intBairro " & strOUTJSQLServer & "= BA.PKId " & strOUTJOracle
    strSQL = strSQL & " ORDER BY L.strDescricao "

    strQueryLogradouro = strSQL
    
End Function

Private Function strQueryRelatorio() As String
    Dim strSQL As String

    strSQL = ""
    strSQL = strSQL & " Select"
    strSQL = strSQL & " LG.Pkid ID_Logradouro, "
    strSQL = strSQL & " LG.strdescricao Logradouro,"
   
    strSQL = strSQL & strSUBSTRING & "(" & gstrRIGHT("IM.strInscricao", gintRetornaTamanhoMascara(TYP_IMOBILIARIA)) & ",1,2) Setor, "
    strSQL = strSQL & strSUBSTRING & "(" & gstrRIGHT("IM.strInscricao", gintRetornaTamanhoMascara(TYP_IMOBILIARIA)) & ",3,3) Quadra, "
    strSQL = strSQL & strSUBSTRING & "(" & gstrRIGHT("IM.strInscricao", gintRetornaTamanhoMascara(TYP_IMOBILIARIA)) & ",6,3) Lote, "
    
    strSQL = strSQL & " IM.dblarea Terreno, "
    strSQL = strSQL & " Sum(AI.INTMEDIDADAAREA) Construido, "
    strSQL = strSQL & " TI.STRMEDIDADATESTADA TestadaP, "
    
    'FATOR_TOPOGRAFIA
    strSQL = strSQL & " (Select "
    strSQL = strSQL & " TV.DBLVALOR "
    strSQL = strSQL & " From "
    strSQL = strSQL & gstrCaracteristicaDoImovel & " CI, "
    strSQL = strSQL & gstrCaracteristicaGeral & " CG, "
    strSQL = strSQL & gstrDetalheDaCaracteristica & " DC, "
    strSQL = strSQL & " tbltabeladevalor TV "
    strSQL = strSQL & " Where "
    strSQL = strSQL & " CI.Intcodigoimobiliario = IM.pkid and "
    strSQL = strSQL & " CI.INTCODIGOCARACTERISTICAGERAL = CG.PKID and "
    strSQL = strSQL & " CG.Intcategoriaconstrucao = " & vetCategoriaConstrucao.ImobiliarioTerreno & " and "
    strSQL = strSQL & " CI.Intcodigodetalhedacaracteristi  = DC.PKID and "
    strSQL = strSQL & " DC.INTTABELADEVALORES = TV.PKID  and "
    strSQL = strSQL & " CG.intUtilizacaoDaCaracteristica = 2 and " 'CI.Intcodigoutilizacaodatabeladev = 2  Alterado Rafael 21/10/04
    strSQL = strSQL & " DC.INTREFERENCIATRIBUTO = " & FATOR_TOPOGRAFIA & ") FT, "
    
    'FATOR_PEDOLOGIA
    strSQL = strSQL & " (Select "
    strSQL = strSQL & " TV.DBLVALOR "
    strSQL = strSQL & " From "
    strSQL = strSQL & gstrCaracteristicaDoImovel & " CI, "
    strSQL = strSQL & gstrCaracteristicaGeral & " CG, "
    strSQL = strSQL & gstrDetalheDaCaracteristica & " DC, "
    strSQL = strSQL & gstrTabelaDeValor & " TV "
    strSQL = strSQL & " Where "
    strSQL = strSQL & " CI.Intcodigoimobiliario = IM.pkid and "
    strSQL = strSQL & " CI.INTCODIGOCARACTERISTICAGERAL = CG.PKID and "
    strSQL = strSQL & " CG.Intcategoriaconstrucao = " & vetCategoriaConstrucao.ImobiliarioTerreno & " and "
    strSQL = strSQL & " CI.Intcodigodetalhedacaracteristi  = DC.PKID and "
    strSQL = strSQL & " DC.INTTABELADEVALORES = TV.PKID  and "
    strSQL = strSQL & " CG.intUtilizacaoDaCaracteristica = 2 and " 'CI.Intcodigoutilizacaodatabeladev = 2  Alterado Rafael 21/10/04
    strSQL = strSQL & " DC.INTREFERENCIATRIBUTO = " & FATOR_PEDOLOGIA & ") FP, "
    
    'FATOR_SITUACAO
    strSQL = strSQL & " (Select "
    strSQL = strSQL & " TV.DBLVALOR "
    strSQL = strSQL & " From "
    strSQL = strSQL & gstrCaracteristicaDoImovel & " CI, "
    strSQL = strSQL & gstrCaracteristicaGeral & " CG, "
    strSQL = strSQL & gstrDetalheDaCaracteristica & " DC, "
    strSQL = strSQL & gstrTabelaDeValor & " TV "
    strSQL = strSQL & " Where "
    strSQL = strSQL & " CI.Intcodigoimobiliario = IM.pkid and "
    strSQL = strSQL & " CI.INTCODIGOCARACTERISTICAGERAL = CG.PKID and "
    strSQL = strSQL & " CG.Intcategoriaconstrucao = " & vetCategoriaConstrucao.ImobiliarioTerreno & " and "
    strSQL = strSQL & " CI.Intcodigodetalhedacaracteristi  = DC.PKID and "
    strSQL = strSQL & " DC.INTTABELADEVALORES = TV.PKID  and "
    strSQL = strSQL & " CG.intUtilizacaoDaCaracteristica = 2 and " 'CI.Intcodigoutilizacaodatabeladev = 2  Alterado Rafael 21/10/04
    strSQL = strSQL & " DC.INTREFERENCIATRIBUTO = " & FATOR_SITUACAO & ") FS, "
    
    'FATOR_ZONEAMENTO
    strSQL = strSQL & " (Select "
    strSQL = strSQL & " TV.DBLVALOR "
    strSQL = strSQL & " From "
    strSQL = strSQL & gstrCaracteristicaDoImovel & " CI, "
    strSQL = strSQL & gstrCaracteristicaGeral & " CG, "
    strSQL = strSQL & gstrDetalheDaCaracteristica & " DC, "
    strSQL = strSQL & gstrTabelaDeValor & " TV "
    strSQL = strSQL & " Where "
    strSQL = strSQL & " CI.Intcodigoimobiliario = IM.pkid and "
    strSQL = strSQL & " CI.INTCODIGOCARACTERISTICAGERAL = CG.PKID and "
    strSQL = strSQL & " CG.Intcategoriaconstrucao = " & vetCategoriaConstrucao.ImobiliarioTerreno & " and "
    strSQL = strSQL & " CI.Intcodigodetalhedacaracteristi  = DC.PKID and "
    strSQL = strSQL & " DC.INTTABELADEVALORES = TV.PKID  and "
    strSQL = strSQL & " CG.intUtilizacaoDaCaracteristica = 2 and " 'CI.Intcodigoutilizacaodatabeladev = 2  Alterado Rafael 21/10/04
    strSQL = strSQL & " DC.INTREFERENCIATRIBUTO = " & FATOR_ZONEAMENTO & ") FZ, "
    
    'FATOR_DESVIO_FERROVIARIO
    strSQL = strSQL & " (Select "
    strSQL = strSQL & " TV.DBLVALOR "
    strSQL = strSQL & " From "
    strSQL = strSQL & gstrCaracteristicaDoImovel & " CI, "
    strSQL = strSQL & gstrCaracteristicaGeral & " CG, "
    strSQL = strSQL & gstrDetalheDaCaracteristica & " DC, "
    strSQL = strSQL & gstrTabelaDeValor & " TV "
    strSQL = strSQL & " Where "
    strSQL = strSQL & " CI.Intcodigoimobiliario = IM.pkid and "
    strSQL = strSQL & " CI.INTCODIGOCARACTERISTICAGERAL = CG.PKID and "
    strSQL = strSQL & " CG.Intcategoriaconstrucao = " & vetCategoriaConstrucao.ImobiliarioTerreno & " and "
    strSQL = strSQL & " CI.Intcodigodetalhedacaracteristi  = DC.PKID and "
    strSQL = strSQL & " DC.INTTABELADEVALORES = TV.PKID  And "
    strSQL = strSQL & " CG.intUtilizacaoDaCaracteristica = 2 and " 'CI.Intcodigoutilizacaodatabeladev = 2  Alterado Rafael 21/10/04
    strSQL = strSQL & " DC.INTREFERENCIATRIBUTO = " & FATOR_DESVIO_FERROVIARIO & ") FDF, "
    
    'FATOR_CORREGO
    strSQL = strSQL & " (Select "
    strSQL = strSQL & " TV.DBLVALOR "
    strSQL = strSQL & " From "
    strSQL = strSQL & gstrCaracteristicaDoImovel & " CI, "
    strSQL = strSQL & gstrCaracteristicaGeral & " CG, "
    strSQL = strSQL & gstrDetalheDaCaracteristica & " DC, "
    strSQL = strSQL & gstrTabelaDeValor & " TV "
    strSQL = strSQL & " Where "
    strSQL = strSQL & " CI.Intcodigoimobiliario = IM.pkid and "
    strSQL = strSQL & " CI.INTCODIGOCARACTERISTICAGERAL = CG.PKID and "
    strSQL = strSQL & " CG.Intcategoriaconstrucao = " & vetCategoriaConstrucao.ImobiliarioTerreno & " and "
    strSQL = strSQL & " CI.Intcodigodetalhedacaracteristi  = DC.PKID and "
    strSQL = strSQL & " DC.INTTABELADEVALORES = TV.PKID  and "
    strSQL = strSQL & " CG.intUtilizacaoDaCaracteristica = 2 and " 'CI.Intcodigoutilizacaodatabeladev = 2  Alterado Rafael 21/10/04
    strSQL = strSQL & " DC.INTREFERENCIATRIBUTO = " & FATOR_CORREGO & " ) FC "
    
    strSQL = strSQL & " From "
    strSQL = strSQL & gstrImobiliario & "  IM, "
    strSQL = strSQL & gstrAreaImobiliario & "  AI, "
    strSQL = strSQL & gstrTestadaImobiliario & "  TI, "
    strSQL = strSQL & gstrLogradouro & " LG, "
    strSQL = strSQL & "(Select Pkid From " & gstrTipoDeTestada & " Where bytprincipal = 1) TT "
    
    strSQL = strSQL & " Where "
    strSQL = strSQL & " IM.Intlogradouro = LG.Pkid and "
    strSQL = strSQL & " IM.pkid = AI.Intimobiliario and "
    strSQL = strSQL & " IM.pkid = TI.INTIMOBILIARIO and "
    strSQL = strSQL & " TI.Inttipodetestada = TT.Pkid and "
    
    If txtquadra.Enabled = True Then
        strSQL = strSQL & strSUBSTRING & "(" & gstrRIGHT("IM.strInscricao", gintRetornaTamanhoMascara(TYP_IMOBILIARIA)) & ",1,2) = '" & Format$(txtSetor.Text, "00") & "' And "
        strSQL = strSQL & strSUBSTRING & "(" & gstrRIGHT("IM.strInscricao", gintRetornaTamanhoMascara(TYP_IMOBILIARIA)) & ",3,3) = '" & Format$(txtquadra.Text, "000") & "'"
    Else
        strSQL = strSQL & " IM.Intlogradouro = " & dbc_intlogradouro.BoundText
    End If
    
    strSQL = strSQL & " Group By "
    strSQL = strSQL & " IM.pkid, "
    strSQL = strSQL & " IM.strInscricao, "
    strSQL = strSQL & " IM.dblArea, "
    strSQL = strSQL & " TI.strmedidadatestada, "
    strSQL = strSQL & " LG.Pkid, "
    strSQL = strSQL & " LG.strDescricao "
    
    strSQL = strSQL & " ORDER BY LG.strDescricao, SETOR, QUADRA, LOTE "
    strQueryRelatorio = strSQL
End Function

Private Function blnDadosOk() As Boolean
    blnDadosOk = False
    If txtquadra.Enabled = True Then
        If Trim(txtSetor.Text) = Empty Then
            ExibeMensagem "O Setor deve ser preenchido Corretamente."
            txtSetor.SetFocus
            Exit Function
        Else
            If Trim(txtquadra.Text) = Empty Then
                ExibeMensagem "A Quadra deve ser preenchida Corretamente."
                txtquadra.SetFocus
                Exit Function
            End If
        End If
    Else
        If dbc_intlogradouro.MatchedWithList = False Then
            ExibeMensagem "O Logradouro deve ser preenchido Corretamente."
            dbc_intlogradouro.SetFocus
            Exit Function
        End If
    End If
    blnDadosOk = True
End Function



Public Sub MantemForm(ByVal strModoOperacao As String)
On Error Resume Next
    If UCase(strModoOperacao) = UCase(gstrImprimir) Then
        If blnDadosOk Then
            If strModoOperacao = UCase("IMPRIMIR") Then
                ImprimeRelatorio rptInsSetorQuadraLogradouro, strQueryRelatorio
                Exit Sub
            End If
        End If
    End If

    If UCase(strModoOperacao) = UCase(gstrFechar) Then
        Unload Me
    End If
    If UCase(strModoOperacao) = UCase(gstrPreencherLista) Then
            PreencherListaDeOpcoes Me.ActiveControl
        Exit Sub
    End If
End Sub
