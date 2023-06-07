VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmContribuintesPorAtividades 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Contribuintes por Atividades"
   ClientHeight    =   3270
   ClientLeft      =   4455
   ClientTop       =   3930
   ClientWidth     =   5100
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3270
   ScaleWidth      =   5100
   Begin VB.CheckBox chkReduzido 
      Caption         =   "Reduzido"
      Height          =   255
      Left            =   900
      TabIndex        =   11
      Top             =   1440
      Width           =   1035
   End
   Begin Threed.SSFrame SSFrame2 
      Height          =   705
      Left            =   60
      TabIndex        =   10
      Top             =   1740
      Width           =   4965
      _Version        =   65536
      _ExtentX        =   8758
      _ExtentY        =   1244
      _StockProps     =   14
      Caption         =   "Ocorrências"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin MSDataListLib.DataCombo dbcOcorrencias 
         Height          =   315
         Left            =   90
         TabIndex        =   12
         Top             =   270
         Width           =   4755
         _ExtentX        =   8387
         _ExtentY        =   556
         _Version        =   393216
         IntegralHeight  =   0   'False
         Text            =   ""
      End
   End
   Begin VB.CheckBox chkTodas 
      Caption         =   "Selecionar todas as atividades"
      Height          =   255
      Left            =   900
      TabIndex        =   6
      Top             =   1170
      Width           =   2520
   End
   Begin MSDataListLib.DataCombo dbcGrupo 
      Height          =   315
      Left            =   885
      TabIndex        =   0
      Top             =   90
      Width           =   4155
      _ExtentX        =   7329
      _ExtentY        =   556
      _Version        =   393216
      IntegralHeight  =   0   'False
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo dbcSubGrupo 
      Height          =   315
      Left            =   885
      TabIndex        =   1
      Top             =   450
      Width           =   4155
      _ExtentX        =   7329
      _ExtentY        =   556
      _Version        =   393216
      IntegralHeight  =   0   'False
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo dbcAtividade 
      Height          =   315
      Left            =   885
      TabIndex        =   2
      Top             =   810
      Width           =   4155
      _ExtentX        =   7329
      _ExtentY        =   556
      _Version        =   393216
      IntegralHeight  =   0   'False
      Text            =   ""
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   645
      Left            =   60
      TabIndex        =   7
      Top             =   2520
      Width           =   4980
      _Version        =   65536
      _ExtentX        =   8784
      _ExtentY        =   1138
      _StockProps     =   14
      Caption         =   "Ordenação"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.OptionButton optInscricao 
         Caption         =   "Inscrição"
         Height          =   285
         Left            =   1260
         TabIndex        =   9
         Top             =   225
         Value           =   -1  'True
         Width           =   1050
      End
      Begin VB.OptionButton optRazaoSocial 
         Caption         =   "Razão Social"
         Height          =   285
         Left            =   2475
         TabIndex        =   8
         Top             =   225
         Width           =   1320
      End
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Atividade:"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   900
      Width           =   705
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Subgrupo:"
      Height          =   195
      Left            =   90
      TabIndex        =   4
      Top             =   540
      Width           =   735
   End
   Begin VB.Label lbl_CodigoDaUtilizacao 
      AutoSize        =   -1  'True
      Caption         =   "Grupo:"
      Height          =   195
      Left            =   345
      TabIndex        =   3
      Top             =   180
      Width           =   480
   End
End
Attribute VB_Name = "frmContribuintesPorAtividades"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bytHabilita As Byte

Public Sub MantemForm(ByVal strModoOperacao As String)
  
  Select Case strModoOperacao
    Case UCase(gstrNovo)
      LimpaObjetos
      
    Case UCase(gstrImprimir)
      If blnDadosOk = True Then
         ImprimeRelatorio rptContribuintesPorAtividades, strQueryRelatorio, "Relatório de Contribuíntes por Atividades"
      End If
      
    Case UCase(gstrPreencherLista)
      PreencherListaDeOpcoes Me.ActiveControl
  
    Case UCase(gstrLocalizar)
      
      If ActiveControl.Name = dbcOcorrencias.Name Then
        If Trim(dbcOcorrencias.Text) = "" Then
            dbcOcorrencias.SetFocus
            Exit Sub
        Else
            LeDaTabelaParaObj gstrOcorrenciaDoEconomico, dbcOcorrencias, "Select pkid,strdescricao from " & gstrOcorrencia & " where strdescricao like '" & dbcOcorrencias.Text & "' and intUtilizacaoDaOcorrencia = 5 ORDER BY strDescricao"
        End If
      End If
      
  End Select
  
End Sub

Private Sub chkTodas_Click()
  If chkTodas.Value = 1 Then
     If dbcGrupo.Enabled = True Then
        bytHabilita = 1
        TrocaCorObjeto dbcGrupo, True
     Else
        If dbcSubGrupo.Enabled = True Then
           bytHabilita = 2
           TrocaCorObjeto dbcSubGrupo, True
        Else
           bytHabilita = 3
           TrocaCorObjeto dbcAtividade, True
        End If
     End If
  Else
     Select Case bytHabilita
       Case 1
         TrocaCorObjeto dbcGrupo, False
         dbcGrupo.SetFocus
       Case 2
         TrocaCorObjeto dbcSubGrupo, False
         dbcSubGrupo.SetFocus
       Case 3
         TrocaCorObjeto dbcAtividade, True
         TrocaCorObjeto dbcAtividade, False
         dbcAtividade.SetFocus
     End Select
  End If
End Sub


Private Sub dbcGrupo_Change()

dbcSubGrupo.BoundText = ""
TrocaCorObjeto dbcSubGrupo, True
dbcAtividade.BoundText = ""
TrocaCorObjeto dbcAtividade, True

End Sub

Private Sub dbcGrupo_Click(Area As Integer)
  If Trim(dbcGrupo.Text) <> "" And Not IsNull(dbcGrupo.SelectedItem) And Area = 2 Then
     Set dbcSubGrupo.RowSource = Nothing
     dbcSubGrupo.Tag = strTagSubGrupo
     TrocaCorObjeto dbcSubGrupo, False
     dbcSubGrupo.SetFocus
  End If
End Sub

Private Sub dbcGrupo_LostFocus()
If Trim(dbcGrupo.Text) <> "" And Not IsNull(dbcGrupo.SelectedItem) Then
     Set dbcSubGrupo.RowSource = Nothing
     dbcSubGrupo.Tag = strTagSubGrupo
     TrocaCorObjeto dbcSubGrupo, False
     dbcSubGrupo.SetFocus
  End If
End Sub


Private Sub dbcOcorrencias_Click(Area As Integer)
   DropDownDataCombo dbcOcorrencias, Me, Area
End Sub

Private Sub dbcOcorrencias_GotFocus()
MarcaCampo dbcOcorrencias
End Sub

Private Sub dbcOcorrencias_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbcOcorrencias, Me, , KeyCode, Shift
End Sub

Private Sub dbcOcorrencias_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbcOcorrencias
End Sub

Private Sub dbcSubGrupo_Change()

dbcAtividade.BoundText = ""
TrocaCorObjeto dbcAtividade, True

End Sub

Private Sub dbcSubGrupo_Click(Area As Integer)
  If Trim(dbcSubGrupo.Text) <> "" And Not IsNull(dbcSubGrupo.SelectedItem) And Area = 2 Then
     dbcAtividade.Text = ""
     Set dbcAtividade.RowSource = Nothing
     dbcAtividade.Tag = strTagAtividade
     TrocaCorObjeto dbcAtividade, False
     dbcAtividade.SetFocus
  End If
End Sub


Private Function blnDadosOk() As Boolean
  blnDadosOk = False
  
  If Trim(dbcGrupo.Text) = "" And chkTodas.Value = 0 Then
     ExibeMensagem "O grupo deve ser informado."
     dbcGrupo.SetFocus
     Exit Function
  End If
  
  If Trim(dbcOcorrencias.Text) = "" Then
     ExibeMensagem "A ocorrência deve ser informada."
     dbcOcorrencias.SetFocus
     Exit Function
  End If
  
  
  blnDadosOk = True
End Function

Private Function LimpaObjetos()
  dbcGrupo.Text = ""
  dbcSubGrupo.Text = ""
  dbcOcorrencias.Text = ""
  dbcAtividade.Text = ""
  Set dbcGrupo.RowSource = Nothing
  Set dbcSubGrupo.RowSource = Nothing
  Set dbcOcorrencias.RowSource = Nothing
  Set dbcAtividade.RowSource = Nothing
  TrocaCorObjeto dbcGrupo, False
  TrocaCorObjeto dbcSubGrupo, True
  TrocaCorObjeto dbcOcorrencias, True
  TrocaCorObjeto dbcAtividade, True
  optInscricao.Value = True
  dbcGrupo.SetFocus
End Function

'Verifica dados preenchidos no form e elabora o Select do Grupo
Private Function strTagGrupo()
Dim strSql As String

  strSql = ""
  strSql = strSql & "SELECT "
  strSql = strSql & "GA.pkID, GA.strNomeDoGrupo "
  strSql = strSql & "FROM "
  strSql = strSql & gstrGrupoDeAtividade & " GA "
     
  strSql = strSql & "ORDER BY GA.strNomeDoGrupo; GA.strNomeDoGrupo "
  
  strTagGrupo = strSql
End Function

'Verifica dados preenchidos no form e elabora o Select do SubGrupo
Private Function strTagSubGrupo()
Dim strSql As String
  
  strSql = ""
  strSql = strSql & "SELECT "
  strSql = strSql & "SGA.pkID, SGA.strNomeDoSubGrupo "
  strSql = strSql & "FROM "
  strSql = strSql & gstrGrupoDeAtividade & " GA, "
  strSql = strSql & gstrSubGrupoDeAtividade & " SGA "
  strSql = strSql & "WHERE "
  strSql = strSql & "SGA.intCodigoDoGrupo = GA.pkID AND "
  strSql = strSql & "GA.pkID = " & dbcGrupo.BoundText & " "
      
  strSql = strSql & "ORDER BY SGA.strNomeDoSubGrupo; SGA.strNomeDoSubGrupo "
    
  strTagSubGrupo = strSql
End Function

'Verifica dados preenchidos no form e elabora o Select da Atividade
Private Function strTagAtividade()
Dim strSql As String
  
  strSql = ""
  strSql = strSql & "SELECT "
  strSql = strSql & "AEC.pkID, AEC.strDescricao "
  strSql = strSql & "FROM "
  strSql = strSql & gstrAtividadeEC & " AEC, "
  strSql = strSql & gstrSubGrupoDeAtividade & " SGA, "
  strSql = strSql & gstrGrupoDeAtividade & " GA "
  strSql = strSql & "WHERE "
  strSql = strSql & "AEC.intGrupo = GA.pkID AND "
  strSql = strSql & "AEC.intSubGrupo = SGA.pkID AND "
  strSql = strSql & "GA.pkID = " & dbcGrupo.BoundText & " AND "
  strSql = strSql & "SGA.pkID = " & dbcSubGrupo.BoundText & " "
  
  strSql = strSql & "ORDER BY AEC.strDescricao; AEC.strDescricao "
 
  strTagAtividade = strSql
End Function

Private Function strQueryRelatorio()
Dim strSql As String
  
  strSql = ""
  strSql = strSql & "SELECT "
  strSql = strSql & gstrRIGHT("EC.strInscricaoCadastral", gintRetornaTamanhoMascara(TYP_ECONOMICA)) & " strInscricao, "
  strSql = strSql & "CT.strNome strRazaoSocial, "
  strSql = strSql & "GA.strNomeDoGrupo strGrupo, SGA.strNomeDoSubGrupo strSubGrupo, "
  strSql = strSql & "AEC.strDescricao strAtividade, " & gstrCASEWHEN("AE.blnPrincipal", "1,'P',0,'S'") & " blnPrincipal, "
  strSql = strSql & " Oe.strdescricao strOcorrencia," 'gstrCASEWHEN("CT.blnInativo", "1, 'Inativo', 0, 'Ativo'") & " blnInativo, "
  strSql = strSql & "LO.strDescricao strLogradouro, EC.intNumero intNumero, "
  strSql = strSql & "BA.strDescricao strBairro, EC.strComplemento strComplemento, "
  strSql = strSql & "CT.intCep intCep "
  strSql = strSql & "FROM "
  strSql = strSql & gstrEconomico & " EC, "
  strSql = strSql & gstrContribuinte & " CT, "
  strSql = strSql & gstrAtividadeDaEmpresa & " AE, "
  strSql = strSql & gstrAtividadeEC & " AEC, "
  strSql = strSql & gstrGrupoDeAtividade & " GA, "
  strSql = strSql & gstrSubGrupoDeAtividade & " SGA, "
  strSql = strSql & gstrLogradouro & " LO, "
  strSql = strSql & gstrBairro & " BA, "
  strSql = strSql & gstrOcorrencia & " Oe "
  strSql = strSql & " WHERE "
  strSql = strSql & " Oe.pkid = Ec.intocorrencia And "
  strSql = strSql & " Ec.intocorrencia = " & Val(dbcOcorrencias.BoundText) & " And "
    
  If chkTodas.Value = 0 Then
    If dbcGrupo.MatchedWithList = True Then
       strSql = strSql & " GA.pkID = " & dbcGrupo.BoundText & " AND "
    End If
    
    If dbcSubGrupo.MatchedWithList = True Then
       strSql = strSql & " SGA.pkID = " & dbcSubGrupo.BoundText & " AND "
    End If
    
    If dbcAtividade.MatchedWithList = True Then
       strSql = strSql & " AEC.pkID = " & dbcAtividade.BoundText & " AND "
    End If
  End If
  
  strSql = strSql & " AEC.intGrupo = GA.pkID AND "
  strSql = strSql & " SGA.intCodigoDoGrupo = GA.pkID AND "
  strSql = strSql & " AEC.intSubGrupo = SGA.pkID AND "

  strSql = strSql & " AE.intAtividade = AEC.pkID AND "
  strSql = strSql & " AE.intEconomico = EC.pkID AND "
  strSql = strSql & " EC.intContribuinte = CT.PkID AND "

  strSql = strSql & " EC.intLogradouro = LO.pkID AND "
  strSql = strSql & " EC.intBairro = BA.pkID "
  
'  If optAtivo.Value = True Then
'      strSql = strSql & " AND blnInativo = 0 "
'  End If
'
'  If optInativo.Value = True Then
'      strSql = strSql & " AND blnInativo = 1 "
'  End If
  
  strSql = strSql & " ORDER BY strGrupo, strSubGrupo, strAtividade, "
  
  If optInscricao.Value = True Then
     strSql = strSql & " strInscricao "
  Else
     strSql = strSql & " strRazaoSocial "
  End If
  
  strSql = strSql & ", blnPrincipal "
  
  strQueryRelatorio = strSql
  
End Function

Private Sub Form_Load()
'***************************************
' Data          : 24/02/2006           *
' Criação       : Pesquisa ocorrencias *
' Responsável   : Fernando Peixoto     *
' Pendência     : Tri0637              *
''**************************************

  dbcGrupo.Tag = strTagGrupo
  dbcOcorrencias.Tag = strOcorrencias
  TrocaCorObjeto dbcSubGrupo, True
  TrocaCorObjeto dbcAtividade, True
  
End Sub

Private Function strOcorrencias() As String
'***************************************
' Data          : 24/02/2006           *
' Criação       : Pesquisa ocorrencias *
' Responsável   : Fernando Peixoto     *
' Pendência     : Tri0637              *
''**************************************
Dim strSql As String

    strSql = ""
    strSql = strSql & " SELECT PkID, "
    strSql = strSql & " StrDescricao "
    strSql = strSql & " FROM " & gstrOcorrencia
    strSql = strSql & " WHERE intUtilizacaoDaOcorrencia = 5 "
    strSql = strSql & " ORDER BY strDescricao"
    strSql = strSql & "; StrDescricao"
    
    strOcorrencias = strSql

End Function
