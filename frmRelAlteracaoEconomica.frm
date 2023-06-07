VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmRelAlteracaoEconomica 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Alterações Cadastrais Econômicas"
   ClientHeight    =   1860
   ClientLeft      =   3030
   ClientTop       =   3990
   ClientWidth     =   5055
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1860
   ScaleWidth      =   5055
   Begin TabDlg.SSTab SSTab1 
      Height          =   1755
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   3096
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Opções de Consulta"
      TabPicture(0)   =   "frmRelAlteracaoEconomica.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fra_Mensagem1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin VB.Frame fra_Mensagem1 
         Height          =   1215
         Left            =   120
         TabIndex        =   1
         Top             =   390
         Width           =   4650
         Begin MSDataListLib.DataCombo dbcintInscricaoInicial 
            Height          =   315
            Left            =   1425
            TabIndex        =   2
            Top             =   270
            Width           =   3045
            _ExtentX        =   5371
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dbcintInscricaoFinal 
            Height          =   315
            Left            =   1425
            TabIndex        =   3
            Top             =   675
            Width           =   3045
            _ExtentX        =   5371
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin VB.Label lblInicial 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Inscrição Inicial:"
            Height          =   195
            Left            =   225
            TabIndex        =   5
            Top             =   330
            Width           =   1140
         End
         Begin VB.Label lblFinal 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Inscrição Final:"
            Height          =   195
            Left            =   225
            TabIndex        =   4
            Top             =   765
            Width           =   1065
         End
      End
   End
End
Attribute VB_Name = "frmRelAlteracaoEconomica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub MantemForm(ByVal strModoOperacao As String)

  Select Case UCase(strModoOperacao)
  
    Case UCase(gstrImprimir)
      If blnDadosOk Then
         ImprimeRelatorio rptAlteracaoEconomico, strQueryRelatorio, "Alterações Cadastrais"
      End If
      
    Case UCase(gstrNovo)
      dbcintInscricaoInicial.Text = ""
      Set dbcintInscricaoInicial.RowSource = Nothing
      dbcintInscricaoFinal.Text = ""
      Set dbcintInscricaoFinal.RowSource = Nothing
      dbcintInscricaoInicial.SetFocus
      
    Case UCase(gstrPreencherLista)
         PreencherListaDeOpcoes Me.ActiveControl
  End Select
  
End Sub

Private Function blnDadosOk() As Boolean
  blnDadosOk = False
  
    If dbcintInscricaoInicial.MatchedWithList = False Then
        ExibeMensagem "A inscrição Inicial deve ser informada."
        dbcintInscricaoInicial.SetFocus
        Exit Function
    ElseIf dbcintInscricaoFinal.MatchedWithList = False Then
        ExibeMensagem "A inscrição Final deve ser informada."
        dbcintInscricaoInicial.SetFocus
        Exit Function
    ElseIf Int(dbcintInscricaoFinal.Text) < Int(dbcintInscricaoInicial.Text) Then
        ExibeMensagem "A inscrição inicial não pode ser maior que a inscrição final."
        dbcintInscricaoFinal.SetFocus
        Exit Function
    End If
  
  blnDadosOk = True
End Function

Private Sub dbcintInscricaoFinal_Click(Area As Integer)
  DropDownDataCombo dbcintInscricaoFinal, Me, Area
End Sub

Private Sub dbcintInscricaoFinal_GotFocus()
  If Trim(dbcintInscricaoFinal.Text) = "" Then
     dbcintInscricaoFinal.Text = Trim(dbcintInscricaoInicial.Text)
     dbcintInscricaoFinal_Click 0
     dbcintInscricaoFinal.Text = Trim(dbcintInscricaoInicial.Text)
     MarcaCampo dbcintInscricaoFinal
  End If
End Sub

Private Sub dbcintInscricaoInicial_Click(Area As Integer)
  DropDownDataCombo dbcintInscricaoInicial, Me, Area
End Sub

Private Sub dbcintInscricaoInicial_GotFocus()
  If Trim(dbcintInscricaoInicial.Text) = "" Then
     dbcintInscricaoInicial.Text = Trim(dbcintInscricaoFinal.Text)
     dbcintInscricaoInicial_Click 0
     dbcintInscricaoInicial.Text = Trim(dbcintInscricaoFinal.Text)
     MarcaCampo dbcintInscricaoInicial
  End If
End Sub

Private Sub Form_Activate()
    gintCodSeguranca = 1361
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrDeletar, gstrAplicar, gstrSalvar
    HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrImprimir, gstrPreencherLista
End Sub

Private Function strQueryInscricao() As String
    Dim strSql As String
    
    strSql = ""
    strSql = strSql & "SELECT "
    strSql = strSql & "Pkid, " & gstrRIGHT("Strinscricaocadastral", gintRetornaTamanhoMascara(TYP_ECONOMICA)) & " Strinscricaocadastral "
    strSql = strSql & "FROM "
    strSql = strSql & gstrEconomico & " ORDER BY Strinscricaocadastral "
    
    strQueryInscricao = strSql

End Function

Private Sub Form_Load()
  dbcintInscricaoInicial.Tag = strQueryInscricao & ";Strinscricaocadastral "
  dbcintInscricaoFinal.Tag = strQueryInscricao & ";Strinscricaocadastral "
End Sub

Private Function strQueryRelatorio() As String
    Dim strSql As String
    
    strSql = "Select EC.Pkid, "
    strSql = strSql & gstrRIGHT("EC.strinscricaocadastral", gintRetornaTamanhoMascara(TYP_ECONOMICA)) & " strinscricaocadastral, "
    strSql = strSql & "HEV.Pkid IntHistorico, "
    strSql = strSql & "1 Tipo, "
    strSql = strSql & "'Razão Social' as strTipo, "
    strSql = strSql & "Replace(Ltrim(rTrim(HEV.Strdescricao)),'Fant:','        Fant: ') Ocorrencias, "
    strSql = strSql & "HEV.Dtmdtinicial DtInicio, "
    strSql = strSql & "HEV.Dtmdtfinal DtFim "
    strSql = strSql & "From " & gstrEconomico & " EC, " & gstrHistoricoEconVariavel & " HEV "
    strSql = strSql & "Where EC.Pkid = HEV.Inteconomico "
    strSql = strSql & "AND HEV.Byttipohistorico = 1 "
    strSql = strSql & "AND strinscricaocadastral Between '" & String(gintLenInscricao - Len(Trim(dbcintInscricaoInicial.Text)), "0") & Trim(dbcintInscricaoInicial.Text) & "' and '" & String(gintLenInscricao - Len(Trim(dbcintInscricaoFinal.Text)), "0") & Trim(dbcintInscricaoFinal.Text) & "' "
    
    strSql = strSql & "Union "
    
    strSql = strSql & "Select EC.Pkid, "
    strSql = strSql & gstrRIGHT("EC.strinscricaocadastral", gintRetornaTamanhoMascara(TYP_ECONOMICA)) & " strinscricaocadastral, "
    strSql = strSql & "HEV.Pkid IntHistorico, "
    strSql = strSql & "2 Tipo, "
    strSql = strSql & "'Endereço' as strTipo, "
    strSql = strSql & "replace(Replace(Ltrim(rTrim(HEV.Strdescricao)),'Compl:','   Compl: '),'CEP:','   CEP: ') Ocorrencias, "
    strSql = strSql & "HEV.Dtmdtinicial DtInicio, "
    strSql = strSql & "HEV.Dtmdtfinal DtFim "
    strSql = strSql & "From " & gstrEconomico & " EC, " & gstrHistoricoEconVariavel & " HEV "
    strSql = strSql & "Where EC.Pkid = HEV.Inteconomico "
    strSql = strSql & "AND HEV.Byttipohistorico = 2 "
    strSql = strSql & "AND strinscricaocadastral Between '" & String(gintLenInscricao - Len(Trim(dbcintInscricaoInicial.Text)), "0") & Trim(dbcintInscricaoInicial.Text) & "' and '" & String(gintLenInscricao - Len(Trim(dbcintInscricaoFinal.Text)), "0") & Trim(dbcintInscricaoFinal.Text) & "' "
    
    strSql = strSql & "Union "
    
    strSql = strSql & "Select EC.Pkid, "
    strSql = strSql & gstrRIGHT("EC.strinscricaocadastral", gintRetornaTamanhoMascara(TYP_ECONOMICA)) & " strinscricaocadastral, "
    strSql = strSql & "HEV.Pkid IntHistorico, "
    strSql = strSql & "3 Tipo, "
    strSql = strSql & "'Atividade' as strTipo, "
    strSql = strSql & "Ltrim(rTrim(HEV.Strdescricao)) Ocorrencias, "
    strSql = strSql & "HEV.Dtmdtinicial DtInicio, "
    strSql = strSql & "HEV.Dtmdtfinal DtFim "
    strSql = strSql & "From " & gstrEconomico & " EC, " & gstrHistoricoEconVariavel & " HEV "
    strSql = strSql & "Where EC.Pkid = HEV.Inteconomico "
    strSql = strSql & "AND HEV.Byttipohistorico = 3 "
    strSql = strSql & "AND strinscricaocadastral Between '" & String(gintLenInscricao - Len(Trim(dbcintInscricaoInicial.Text)), "0") & Trim(dbcintInscricaoInicial.Text) & "' and '" & String(gintLenInscricao - Len(Trim(dbcintInscricaoFinal.Text)), "0") & Trim(dbcintInscricaoFinal.Text) & "' "
    
    strSql = strSql & "Union "
    
    strSql = strSql & "Select EC.Pkid, "
    strSql = strSql & gstrRIGHT("EC.strinscricaocadastral", gintRetornaTamanhoMascara(TYP_ECONOMICA)) & " strinscricaocadastral, "
    strSql = strSql & "HEV.Pkid IntHistorico, "
    strSql = strSql & "4 Tipo, "
    strSql = strSql & "'Sócio' as strTipo, "
    strSql = strSql & "Ltrim(rTrim(HEV.Strdescricao)) Ocorrencias, "
    strSql = strSql & "HEV.Dtmdtinicial DtInicio, "
    strSql = strSql & "HEV.Dtmdtfinal DtFim "
    strSql = strSql & "From " & gstrEconomico & " EC, " & gstrHistoricoEconVariavel & " HEV "
    strSql = strSql & "Where EC.Pkid = HEV.Inteconomico "
    strSql = strSql & "AND HEV.Byttipohistorico = 4 "
    strSql = strSql & "AND strinscricaocadastral Between '" & String(gintLenInscricao - Len(Trim(dbcintInscricaoInicial.Text)), "0") & Trim(dbcintInscricaoInicial.Text) & "' and '" & String(gintLenInscricao - Len(Trim(dbcintInscricaoFinal.Text)), "0") & Trim(dbcintInscricaoFinal.Text) & "' "
    
    strSql = strSql & "Union "
    
    strSql = strSql & "Select EC.Pkid, "
    strSql = strSql & gstrRIGHT("EC.strinscricaocadastral", gintRetornaTamanhoMascara(TYP_ECONOMICA)) & " strinscricaocadastral, "
    strSql = strSql & "HEV.Pkid IntHistorico, "
    strSql = strSql & "5 Tipo, "
    strSql = strSql & "'Ocorrência' as strTipo, "
    strSql = strSql & "Ltrim(rTrim(HEV.Strdescricao)) Ocorrencias, "
    strSql = strSql & "HEV.Dtmdtinicial DtInicio, "
    strSql = strSql & "HEV.Dtmdtfinal DtFim "
    strSql = strSql & "From " & gstrEconomico & " EC, " & gstrHistoricoEconVariavel & " HEV "
    strSql = strSql & "Where EC.Pkid = HEV.Inteconomico "
    strSql = strSql & "AND HEV.Byttipohistorico = 5 "
    strSql = strSql & "AND strinscricaocadastral Between '" & String(gintLenInscricao - Len(Trim(dbcintInscricaoInicial.Text)), "0") & Trim(dbcintInscricaoInicial.Text) & "' and '" & String(gintLenInscricao - Len(Trim(dbcintInscricaoFinal.Text)), "0") & Trim(dbcintInscricaoFinal.Text) & "' "
    
    strSql = strSql & "Union "
    
    strSql = strSql & "Select EC.Pkid, "
    strSql = strSql & gstrRIGHT("EC.strinscricaocadastral", gintRetornaTamanhoMascara(TYP_ECONOMICA)) & " strinscricaocadastral, "
    strSql = strSql & "HEV.Pkid IntHistorico, "
    strSql = strSql & "6 Tipo, "
    strSql = strSql & "'Publicidade' as strTipo, "
    strSql = strSql & "Ltrim(rTrim(HEV.Strdescricao)) Ocorrencias, "
    strSql = strSql & "HEV.Dtmdtinicial DtInicio, "
    strSql = strSql & "HEV.Dtmdtfinal DtFim "
    strSql = strSql & "From " & gstrEconomico & " EC, " & gstrHistoricoEconVariavel & " HEV "
    strSql = strSql & "Where EC.Pkid = HEV.Inteconomico "
    strSql = strSql & "AND HEV.Byttipohistorico = 6 "
    strSql = strSql & "AND strinscricaocadastral Between '" & String(gintLenInscricao - Len(Trim(dbcintInscricaoInicial.Text)), "0") & Trim(dbcintInscricaoInicial.Text) & "' and '" & String(gintLenInscricao - Len(Trim(dbcintInscricaoFinal.Text)), "0") & Trim(dbcintInscricaoFinal.Text) & "' "
    
    strSql = strSql & "Union "
    
    strSql = strSql & "Select EC.Pkid, "
    strSql = strSql & gstrRIGHT("EC.strinscricaocadastral", gintRetornaTamanhoMascara(TYP_ECONOMICA)) & " strinscricaocadastral, "
    strSql = strSql & "HEV.Pkid IntHistorico, "
    strSql = strSql & "7 Tipo, "
    strSql = strSql & "'ISSQN' as strTipo, "
    strSql = strSql & "Ltrim(rTrim(HEV.Strdescricao)) Ocorrencias, "
    strSql = strSql & "HEV.Dtmdtinicial DtInicio, "
    strSql = strSql & "HEV.Dtmdtfinal DtFim "
    strSql = strSql & "From " & gstrEconomico & " EC, " & gstrHistoricoEconVariavel & " HEV "
    strSql = strSql & "Where EC.Pkid = HEV.Inteconomico "
    strSql = strSql & "AND HEV.Byttipohistorico = 7 "
    strSql = strSql & "AND strinscricaocadastral Between '" & String(gintLenInscricao - Len(Trim(dbcintInscricaoInicial.Text)), "0") & Trim(dbcintInscricaoInicial.Text) & "' and '" & String(gintLenInscricao - Len(Trim(dbcintInscricaoFinal.Text)), "0") & Trim(dbcintInscricaoFinal.Text) & "' "
    strSql = strSql & "Order By strinscricaocadastral, Tipo Asc, DtInicio Asc, dtfim asc, intHistorico Desc "
    
    strQueryRelatorio = strSql

End Function

