VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmFichaCadastroImobiliario 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ficha de cadastro imobiliário"
   ClientHeight    =   2580
   ClientLeft      =   3690
   ClientTop       =   4110
   ClientWidth     =   6015
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2580
   ScaleWidth      =   6015
   Begin TabDlg.SSTab tab_3dPasta 
      Height          =   2475
      Left            =   60
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   30
      Width           =   5955
      _ExtentX        =   10504
      _ExtentY        =   4366
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Ficha de Cadastro Imobiliário"
      TabPicture(0)   =   "frmFichaCadastroImobiliario.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fra_Inscricao"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fra_Contribuinte"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
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
            TabIndex        =   5
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
      Begin VB.Frame fra_Inscricao 
         Caption         =   "Inscrição"
         Height          =   825
         Left            =   240
         TabIndex        =   1
         Top             =   420
         Width           =   5415
         Begin MSDataListLib.DataCombo dbc_strInscricao 
            Height          =   315
            Left            =   1260
            TabIndex        =   6
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
            TabIndex        =   2
            Top             =   390
            Width           =   645
         End
      End
   End
End
Attribute VB_Name = "frmFichaCadastroImobiliario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim blnInscricao As Boolean
Dim blnContribuinte As Boolean

Private Function MontaDataCombo(Param As String)
    Dim StrSQL As String
    
    StrSQL = "(SELECT pkid, "
    StrSQL = StrSQL & gstrRIGHT("strInscricao", gintRetornaTamanhoMascara(TYP_IMOBILIARIA))
    StrSQL = StrSQL & " strInscricao FROM "
    StrSQL = StrSQL & gstrImobiliario
    If Not Param = "" Then
        StrSQL = StrSQL & " WHERE " & gstrRIGHT("strInscricao", gintRetornaTamanhoMascara(TYP_IMOBILIARIA))
        StrSQL = StrSQL & " Like '" & Param & "%'"
    End If
    If bytDBType = Oracle Then
       StrSQL = StrSQL & " ORDER BY strInscricao)"
    Else
       StrSQL = StrSQL & ")"
    End If
    LeDaTabelaParaObj "", dbc_strInscricao, StrSQL
    
End Function
Public Sub MantemForm(ByVal strModoOperacao As String)
    
    Select Case UCase(strModoOperacao)
      Case UCase(gstrPreencherLista)
           MontaDataCombo (dbc_strInscricao.Text)
      Case UCase(gstrImprimir)
        If blnDadosOk Then
          ImprimeRelatorio rptFichaCadastroImobiliario, strQueryRelatorio, "Ficha de Cadastro Imobiliário"
        End If
      
      Case UCase(gstrNovo)
        dbc_strInscricao.Text = ""
        DesabilitaContribuinte
        dbc_strInscricao.SetFocus
        
    End Select
      
End Sub
Public Function strQueryRelatorio() As String
    Dim StrSQL As String
    
    StrSQL = "SELECT "
    StrSQL = StrSQL & "IM.pkid IDImobiliario, "
    StrSQL = StrSQL & "IM.strmatricula, "
    StrSQL = StrSQL & "IM.strcartorio, "
    StrSQL = StrSQL & "IM.intfolha, "
    StrSQL = StrSQL & "IM.dtmdtmatricula, "
    StrSQL = StrSQL & "IM.dtmdtescritura, "
    StrSQL = StrSQL & "IM.intLivro, "
    StrSQL = StrSQL & "IM.Dblarea, "
    StrSQL = StrSQL & gstrRIGHT("strInscricao", gintRetornaTamanhoMascara(TYP_IMOBILIARIA)) & " strInscricao, "
    StrSQL = StrSQL & "IM.strInscricao, "
    StrSQL = StrSQL & gstrISNULL("LG.strCodigo", "''") & strCONCAT & " ' - '" & strCONCAT & gstrISNULL("tlg.strsigla", "''") & strCONCAT & "' '" & strCONCAT & gstrISNULL("ttlg.strsigla", "''") & strCONCAT & "' '" & strCONCAT & gstrISNULL("LG.Strdescricao", "''") & " strLogradouro, "
    StrSQL = StrSQL & "IM.intcep, "
    StrSQL = StrSQL & "IM.intnumero, "
    StrSQL = StrSQL & "IM.strcomplemento, "
    StrSQL = StrSQL & "BR.strcodigo " & strCONCAT & "' - '" & strCONCAT & " br.strdescricao strBairro, "
    StrSQL = StrSQL & "IM.strquadra, "
    StrSQL = StrSQL & "IM.strlote, "
    StrSQL = StrSQL & "lt.strnome strLoteamento, "
    StrSQL = StrSQL & "IM.intNumeroC, "
    StrSQL = StrSQL & "IM.strlogradouroc strLogradouroC, "
    StrSQL = StrSQL & "IM.strcomplementoc, "
    StrSQL = StrSQL & "IM.intNumeroc, "
    StrSQL = StrSQL & "IM.intCepC, "
    StrSQL = StrSQL & "IM.strbairroc, "
    StrSQL = StrSQL & "MU.strDescricao strMunicipioC, "
    StrSQL = StrSQL & "CTPP.Strnome strProprietario, "
    StrSQL = StrSQL & "CTPP.Strcnpjcpf strCnpjCpfPP, "
    StrSQL = StrSQL & "CTPM.Strnome strPromissario, "
    StrSQL = StrSQL & "CTPM.Strcnpjcpf strCnpjCpfPM, "
    StrSQL = StrSQL & "IM.Bytedificado, "
    StrSQL = StrSQL & "UF.strSigla strSiglaC, "
    StrSQL = StrSQL & "IM.intContribuinte,"
    StrSQL = StrSQL & " IM.DTMDTCANCELAMENTO "
    StrSQL = StrSQL & "FROM "
    StrSQL = StrSQL & gstrImobiliario & " IM, "
    StrSQL = StrSQL & gstrLogradouro & " LG, "
    StrSQL = StrSQL & gstrContribuinte & " CTPP, "
    StrSQL = StrSQL & gstrTipoLogradouro & " TLG, "
    StrSQL = StrSQL & gstrTituloLogradouro & " TTLG, "
    StrSQL = StrSQL & gstrBairro & " BR, "
    StrSQL = StrSQL & gstrLoteamento & " LT, "
    StrSQL = StrSQL & gstrCidade & " MU, "
    StrSQL = StrSQL & gstrUF & " UF, "
    StrSQL = StrSQL & gstrContribuinte & " CTPM "
    StrSQL = StrSQL & "WHERE "
    StrSQL = StrSQL & " lg.Pkid = IM.intlogradouro AND "
    StrSQL = StrSQL & " IM.intcontribuinte = ctpp.pkid  AND "
    StrSQL = StrSQL & " IM.inttipologradouro " & strOUTJSQLServer & "= tlg.pkid " & strOUTJOracle & " AND "
    StrSQL = StrSQL & " lg.inttitulologradouro " & strOUTJSQLServer & "= ttlg.pkid " & strOUTJOracle & " AND "
    StrSQL = StrSQL & " lg.intbairro " & strOUTJSQLServer & "= br.pkid " & strOUTJOracle & " AND "
    StrSQL = StrSQL & " IM.intloteamento " & strOUTJSQLServer & "= lt.pkid " & strOUTJOracle & " AND "
    StrSQL = StrSQL & " IM.intmunicipioc " & strOUTJSQLServer & "= mu.pkid " & strOUTJOracle & " AND "
    StrSQL = StrSQL & " IM.intufc " & strOUTJSQLServer & "= uf.pkid " & strOUTJOracle & " AND "
    StrSQL = StrSQL & " IM.intpromissario " & strOUTJSQLServer & "= ctpm.pkid " & strOUTJOracle
    
    If blnInscricao Then
        StrSQL = StrSQL & " AND IM.strInscricao = '" & String(gintLenInscricao - Len(Trim(dbc_strInscricao.Text)), "0") & dbc_strInscricao.Text & "' "
    Else
        StrSQL = StrSQL & " AND CTPP.pkID = '" & dbc_strContribuinte.BoundText & "' "
    End If
    
    StrSQL = StrSQL & " ORDER BY IM.strInscricao "

    strQueryRelatorio = StrSQL

End Function
Private Function blnDadosOk() As Boolean

    blnDadosOk = False
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

    blnDadosOk = True
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
    dbc_strInscricao.Tag = "SELECT pkid, " & gstrRIGHT("strInscricao", gintRetornaTamanhoMascara(TYP_IMOBILIARIA)) & " strInscricao FROM " & gstrImobiliario & " ORDER BY strInscricao;strInscricao"
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

Dim StrSQL As String

    StrSQL = "SELECT ct.pkid, ct.strnome "
    
    StrSQL = StrSQL & "FROM " & gstrModuloContribuinte & " MCT , " & _
             gstrItens & " IT, " & _
             gstrContribuinte & " CT "
    
    StrSQL = StrSQL & "WHERE MCT.intItem = IT.Pkid " & _
             "AND it.strcoditem = 'J' " & _
             "AND mct.intcontribuinte = ct.pkid "

    StrSQL = StrSQL & "ORDER BY strNome"
    
strQueryContribuintes = StrSQL

End Function
