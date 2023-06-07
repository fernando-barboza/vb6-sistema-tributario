VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmFichaCadastroEconomico 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ficha de Cadastro Econômico"
   ClientHeight    =   2700
   ClientLeft      =   4125
   ClientTop       =   3870
   ClientWidth     =   6210
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2700
   ScaleWidth      =   6210
   Begin TabDlg.SSTab tab_3dPasta 
      Height          =   2475
      Left            =   120
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   5955
      _ExtentX        =   10504
      _ExtentY        =   4366
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Ficha de Cadastro Econômico"
      TabPicture(0)   =   "frmFichaCadastroEconomico.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fra_Contribuinte"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fra_Inscricao"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      Begin VB.Frame fra_Inscricao 
         Caption         =   "Inscrição"
         Height          =   825
         Left            =   240
         TabIndex        =   5
         Top             =   420
         Width           =   5415
         Begin MSDataListLib.DataCombo dbc_strInscricao 
            Height          =   315
            Left            =   1260
            TabIndex        =   1
            Top             =   330
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin VB.Label lblDtMovimento 
            AutoSize        =   -1  'True
            Caption         =   "Inscrição"
            Height          =   195
            Left            =   525
            TabIndex        =   6
            Top             =   390
            Width           =   645
         End
      End
      Begin VB.Frame fra_Contribuinte 
         Caption         =   "Contribuinte"
         Height          =   915
         Left            =   240
         TabIndex        =   3
         Top             =   1350
         Width           =   5415
         Begin MSDataListLib.DataCombo dbc_strContribuinte 
            Height          =   315
            Left            =   1260
            TabIndex        =   2
            Top             =   360
            Width           =   3975
            _ExtentX        =   7011
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Contribuinte"
            Height          =   195
            Left            =   315
            TabIndex        =   4
            Top             =   420
            Width           =   840
         End
      End
   End
End
Attribute VB_Name = "frmFichaCadastroEconomico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim blnInscricao As Boolean
Dim blnContribuinte As Boolean

Public Sub MantemForm(ByVal strModoOperacao As String)

    Select Case UCase(strModoOperacao)
      Case UCase(gstrPreencherLista)
        PreencherListaDeOpcoes Me.ActiveControl
        
      Case UCase(gstrImprimir)
        If blnDadosok Then
          ImprimeRelatorio rptFichaCadastroEconomico, strQueryRelatorio, "Ficha de Cadastro Econômico"
        End If
      
      Case UCase(gstrNovo)
        dbc_strInscricao.Text = ""
        DesabilitaContribuinte
        dbc_strInscricao.SetFocus
      Case gstrPreencherLista
        PreencherListaDeOpcoes Me.ActiveControl
    End Select
      
End Sub
Public Function strQueryRelatorio() As String

Dim strSQL As String

    strSQL = "SELECT " & _
             "EC.pkid PkidEconomico, EC.strEmissao, " & _
             gstrRIGHT("EC.strInscricaoImobiliaria", gintRetornaTamanhoMascara(TYP_IMOBILIARIA)) & " strInscricaoImobiliaria, " & _
             gstrRIGHT("EC.strInscricaoCadastral", gintRetornaTamanhoMascara(TYP_ECONOMICA)) & " strInscricaoCadastral, " & _
             gstrISNULL("TPL.strSigla", "''") & strCONCAT & "' '" & strCONCAT & gstrISNULL("TTL.strSigla", "''") & strCONCAT & "' '" & strCONCAT & " LO.strDescricao strLogradouro, " & _
             "EC.dtmdataabertura, EC.strcodprocabertura " & strCONCAT & "'/'" & strCONCAT & gstrCONVERT(CDT_VARCHAR, "EC.intexerprocabertura") & strCONCAT & "' '" & strCONCAT & gstrCONVERT(CDT_VARCHAR, "EC.bitdigprocabertura") & " strProcessoAbertura, " & _
             "EC.dtmdataencerramento, EC.strcodprocencerramento " & strCONCAT & "'/'" & strCONCAT & gstrCONVERT(CDT_VARCHAR, "EC.intexerprocencerramento") & strCONCAT & "' '" & strCONCAT & gstrCONVERT(CDT_VARCHAR, "ec.bitdigprocencerramento") & " strProcessoEncerramento, " & _
             "EC.dtmdataprocesso, EC.strcodprocesso " & strCONCAT & "'/'" & strCONCAT & gstrCONVERT(CDT_VARCHAR, "EC.intexerprocesso") & strCONCAT & "' '" & strCONCAT & gstrCONVERT(CDT_VARCHAR, "EC.bitdigprocesso") & " strProcessoOcor, " & _
             "OE.strDescricao strOcorrenciaProc, EC.strhistoricoprocesso, EC.dblvalorestimado, EC.dtmdataestimativa," & _
             "TI.strdescricao strTipoIss, LS.STRDESCRICAO strListaServico, OC.STRDESCRICAO strOcorrencia, EC.Intnumdeempregados, " & _
             "HF.STRDESCRICAO strHorarioFuncionamento, EC.Strmanhade, EC.Strmanhaate, EC.Strtardede, EC.Strtardeate, EC.Strnoitede, EC.strnoiteate, EC.Strmadrugadade, EC.Strmadrugadaate, " & _
             "EC.dblareaocupada, EC.dblareaanuncio, " & _
             "EC.intNumero, EC.strComplemento, EC.intCep, BA.strDescricao strBairro, " & _
             gstrCONVERT(CDT_VARCHAR, "CO.PKID") & strCONCAT & "' - '" & strCONCAT & " CO.STRNOME strNome, CO.Strnomefantasia, CO.STRINSCRICAOESTADUAL, CO.Strcnpjcpf, CO.bytNaturezaJuridica, " & _
             "AEC.STRDESCRICAO strAtividade, TR.STRDESCRICAO strTributo, TT.STRDESCRICAO strTipoTributo, " & _
             "AB.strDescricao strAtividadeBasica, AEC.intcodigo "
             
    If bytDBType = Oracle Then
        strSQL = strSQL & "FROM "
        strSQL = strSQL & gstrEconomico & " EC, "
        strSQL = strSQL & gstrLogradouro & " LO, "
        strSQL = strSQL & gstrTipoLogradouro & " TPL, "
        strSQL = strSQL & gstrTituloLogradouro & " TTL, "
        strSQL = strSQL & gstrBairro & " BA, "
        strSQL = strSQL & gstrContribuinte & " CO, "
        strSQL = strSQL & gstrAtividadeBasica & " AB, "
        strSQL = strSQL & gstrOcorrenciaDoEconomico & " OE, "
        strSQL = strSQL & "tblHorarioFuncionamento" & " HF, "
        strSQL = strSQL & gstrTipoIss & " TI, "
        strSQL = strSQL & gstrListaServico & " LS, "
        strSQL = strSQL & gstrOcorrencia & " OC, "
        strSQL = strSQL & gstrAtividadeDaEmpresa & " AE, "
        strSQL = strSQL & gstrAtivEmpresaTributo & " AET, "
        strSQL = strSQL & gstrAtividadeEC & " AEC, "
        strSQL = strSQL & gstrTributo & " TR, "
        strSQL = strSQL & gstrTributoTipo & " TT "
        strSQL = strSQL & "WHERE "
        strSQL = strSQL & "LO.Pkid " & strOUTJOracle & " =" & strOUTJSQLServer & " EC.intLogradouro AND "
        strSQL = strSQL & "TPL.Pkid " & strOUTJOracle & " =" & strOUTJSQLServer & " LO.intTipoLogradouro AND "
        strSQL = strSQL & "TTL.Pkid " & strOUTJOracle & " =" & strOUTJSQLServer & " LO.intTituloLogradouro AND "
        strSQL = strSQL & "BA.Pkid " & strOUTJOracle & " =" & strOUTJSQLServer & " EC.intBairro AND "
        strSQL = strSQL & "CO.Pkid " & strOUTJOracle & " =" & strOUTJSQLServer & " EC.intContribuinte AND "
        strSQL = strSQL & "AB.Pkid " & strOUTJOracle & " =" & strOUTJSQLServer & " EC.intAtividadeBasica AND "
        strSQL = strSQL & "OE.Pkid " & strOUTJOracle & " =" & strOUTJSQLServer & " EC.intOcorrenciaDoEconomico AND "
        strSQL = strSQL & "HF.Pkid " & strOUTJOracle & " =" & strOUTJSQLServer & " EC.intHorarioFuncionamento AND "
        strSQL = strSQL & "TI.Pkid " & strOUTJOracle & " =" & strOUTJSQLServer & " EC.intTipoIss AND "
        strSQL = strSQL & "LS.Pkid " & strOUTJOracle & " =" & strOUTJSQLServer & " EC.intListaServico AND "
        strSQL = strSQL & "OC.Pkid " & strOUTJOracle & " =" & strOUTJSQLServer & " EC.intOcorrencia AND "
        strSQL = strSQL & "AE.intEconomico " & strOUTJOracle & " =" & strOUTJSQLServer & " EC.Pkid AND "
        strSQL = strSQL & "AET.INTATIVIDADEDAEMPRESA " & strOUTJOracle & " =" & strOUTJSQLServer & " AE.Pkid AND "
        strSQL = strSQL & "AEC.PKID " & strOUTJOracle & " =" & strOUTJSQLServer & " AE.INTATIVIDADE AND "
        strSQL = strSQL & "TR.PKID  " & strOUTJOracle & " =" & strOUTJSQLServer & " AET.INTTRIBUTO AND "
        strSQL = strSQL & "TT.PKID  " & strOUTJOracle & " =" & strOUTJSQLServer & " TR.INTTRIBUTOTIPO AND "
    Else
        strSQL = strSQL & " FROM tblEconomico EC LEFT OUTER JOIN " & _
                      " tblLogradouro LO ON EC.intLogradouro = LO.PKId LEFT OUTER JOIN " & _
                      " tblTipoLogradouro TPL ON LO.intTipoLogradouro = TPL.PKId LEFT OUTER JOIN " & _
                      " tblTituloLogradouro TTL ON LO.intTituloLogradouro = TTL.PKId LEFT OUTER JOIN " & _
                      " tblBairro BA ON EC.intBairro = BA.PKId LEFT OUTER JOIN " & _
                      " tblContribuinte CO ON EC.intContribuinte = CO.PKId LEFT OUTER JOIN " & _
                      " tblAtividadeBasica AB ON EC.intAtividadeBasica = AB.PKId LEFT OUTER JOIN " & _
                      " TBLOCORRENCIADOECONOMICO OE ON EC.INTOCORRENCIADOECONOMICO = OE.PKID LEFT OUTER JOIN " & _
                      " tblHorarioFuncionamento HF ON EC.intHorarioFuncionamento = HF.PKId LEFT OUTER JOIN " & _
                      " TBLTIPOISS TI ON EC.INTTIPOISS = TI.PKID LEFT OUTER JOIN " & _
                      " TBLLISTASERVICO LS ON EC.INTLISTASERVICO = LS.PKID LEFT OUTER JOIN " & _
                      " tblOcorrencia OC ON EC.intOcorrencia = OC.PKId LEFT OUTER JOIN " & _
                      " tblAtividadeDaEmpresa AE ON EC.PKId = AE.intEconomico LEFT OUTER JOIN " & _
                      " TBLATIVEMPRESATRIBUTO AET ON AE.PKId = AET.INTATIVIDADEDAEMPRESA LEFT OUTER JOIN " & _
                      " tblAtividadeEC AEC ON AE.intAtividade = AEC.PKId LEFT OUTER JOIN " & _
                      " TBLTRIBUTO TR ON AET.INTTRIBUTO = TR.PKID LEFT OUTER JOIN " & _
                      " TBLTRIBUTOTIPO TT ON TR.INTTRIBUTOTIPO = TT.PKID " & _
                      " WHERE "
    End If
    
    If blnInscricao Then
        strSQL = strSQL & " EC.strInscricaoCadastral = '" & String(gintLenInscricao - Len(Trim(dbc_strInscricao.Text)), "0") & dbc_strInscricao.Text & "' "
    Else
        strSQL = strSQL & " EC.intContribuinte = " & dbc_strContribuinte.BoundText
    End If
    
    strSQL = strSQL & " ORDER BY EC.strInscricaoCadastral "

    strQueryRelatorio = strSQL

End Function
Private Function blnDadosok() As Boolean

    blnDadosok = False
    If blnInscricao Then
       If dbc_strInscricao.MatchedWithList = False Then
          ExibeMensagem "A Inscrição deve ser informada."
          dbc_strInscricao.SetFocus
          Exit Function
       End If
    Else
       If dbc_strContribuinte.MatchedWithList = False Then
           ExibeMensagem "O Contribuinte deve ser informado."
           dbc_strContribuinte.SetFocus
           Exit Function
       End If
    End If

    blnDadosok = True
    
End Function

Private Sub DesabilitaInscricao()
    dbc_strInscricao.Text = ""
    Set dbc_strInscricao.RowSource = Nothing
    TrocaCorObjeto dbc_strInscricao, True
    TrocaCorObjeto dbc_strContribuinte, False
    
    blnInscricao = False
End Sub

Private Sub DesabilitaContribuinte()
    dbc_strContribuinte.Text = ""
    Set dbc_strContribuinte.RowSource = Nothing
    TrocaCorObjeto dbc_strInscricao, False
    TrocaCorObjeto dbc_strContribuinte, True
        
    blnInscricao = True
End Sub

Private Sub dbc_strContribuinte_Click(Area As Integer)
    DropDownDataCombo dbc_strContribuinte, Me, Area
End Sub

Private Sub dbc_strContribuinte_GotFocus()
    MarcaCampo dbc_strContribuinte
End Sub

Private Sub dbc_strInscricao_Click(Area As Integer)
    DropDownDataCombo dbc_strInscricao, Me, Area
End Sub

Private Sub dbc_strInscricao_GotFocus()
    MarcaCampo dbc_strInscricao
End Sub

Private Sub Form_Load()
    dbc_strContribuinte.Tag = strQueryContribuintes & ";strNome"
    dbc_strInscricao.Tag = "SELECT pkid, " & gstrRIGHT("strInscricaoCadastral", gintRetornaTamanhoMascara(TYP_ECONOMICA)) & " strInscricao FROM " & gstrEconomico & " ORDER BY strInscricaoCadastral;strInscricaoCadastral"
    DesabilitaContribuinte
End Sub

Private Sub fra_Contribuinte_Click()
    DesabilitaInscricao
    dbc_strContribuinte.SetFocus
End Sub
Private Sub fra_Inscricao_Click()
    DesabilitaContribuinte
    dbc_strInscricao.SetFocus
End Sub

Private Function strQueryContribuintes() As String

Dim strSQL As String

    strSQL = "SELECT ct.pkid, ct.strnome "
    
    strSQL = strSQL & "FROM " & gstrModuloContribuinte & " MCT , " & _
             gstrItens & " IT, " & _
             gstrContribuinte & " CT "
    
    strSQL = strSQL & "WHERE MCT.intItem = IT.Pkid " & _
             "AND it.strcoditem = 'J' " & _
             "AND mct.intcontribuinte = ct.pkid "

    strSQL = strSQL & "ORDER BY strNome"
    
strQueryContribuintes = strSQL

End Function


