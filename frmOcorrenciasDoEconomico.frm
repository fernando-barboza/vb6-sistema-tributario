VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmOcorrenciasDoEconomico 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ocorrências do Econômico"
   ClientHeight    =   3360
   ClientLeft      =   3210
   ClientTop       =   1695
   ClientWidth     =   6450
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   6450
   Begin TabDlg.SSTab tab_3dPasta 
      Height          =   3075
      Left            =   150
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   150
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   5424
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Ocorrências do Econômico"
      TabPicture(0)   =   "frmOcorrenciasDoEconomico.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fra_Inscricao"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin VB.Frame fra_Inscricao 
         Height          =   2295
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   5655
         Begin MSDataListLib.DataCombo dbcstrInscricaoInicial 
            Height          =   315
            Left            =   2250
            TabIndex        =   2
            Top             =   810
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dbcstrInscricaoFinal 
            Height          =   315
            Left            =   2250
            TabIndex        =   4
            Top             =   1320
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin VB.Label lblInscricaoFinal 
            AutoSize        =   -1  'True
            Caption         =   "Inscrição Final:"
            Height          =   195
            Left            =   870
            TabIndex        =   5
            Top             =   1380
            Width           =   1065
         End
         Begin VB.Label lblInscricaoInicial 
            AutoSize        =   -1  'True
            Caption         =   "Inscrição Inicial:"
            Height          =   195
            Left            =   870
            TabIndex        =   3
            Top             =   870
            Width           =   1200
         End
      End
   End
End
Attribute VB_Name = "frmOcorrenciasDoEconomico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Function strQueryInscricaoInicial() As String

    Dim strSql As String

    strSql = ""
    strSql = strSql & "SELECT "
    strSql = strSql & " eco.PKID, "
    strSql = strSql & gstrRIGHT("eco.strInscricaoCadastral", gintRetornaTamanhoMascara(TYP_ECONOMICA)) & "strInscricaoCadastral "
    strSql = strSql & "FROM "
    strSql = strSql & gstrEconomico & " eco "
    
    If Trim(dbcstrInscricaoInicial.Text) <> "" Then
        strSql = strSql & "WHERE "
        strSql = strSql & gstrRIGHT("eco.strInscricaoCadastral", gintRetornaTamanhoMascara(TYP_ECONOMICA)) & "LIKE '%" & dbcstrInscricaoInicial.Text & "%'"
    End If
    
    strSql = strSql & " ORDER BY eco.strInscricaoCadastral "
       
    strQueryInscricaoInicial = strSql
       
End Function

Private Function strQueryInscricaoFinal() As String

    Dim strSql As String

    strSql = ""
    strSql = strSql & "SELECT eco.PKID, "
    strSql = strSql & gstrRIGHT("eco.strInscricaoCadastral", gintRetornaTamanhoMascara(TYP_ECONOMICA)) & "strInscricaoCadastral "
    strSql = strSql & "  FROM " & gstrEconomico & " eco "
                      
    If Trim(dbcstrInscricaoFinal.Text) <> "" Then
        strSql = strSql & " WHERE "
        strSql = strSql & " eco.strInscricaoCadastral = " & dbcstrInscricaoFinal.Text
    End If
    
    strSql = strSql & " ORDER BY eco.strInscricaoCadastral "
       
    strQueryInscricaoFinal = strSql
       
End Function

Private Function strQueryRelatorio() As String

    Dim strSql As String

    strSql = ""
    strSql = strSql & "SELECT "
    strSql = strSql & gstrRIGHT("eco.STRINSCRICAOCADASTRAL", gintRetornaTamanhoMascara(TYP_ECONOMICA)) & "STRINSCRICAOCADASTRAL, "
    strSql = strSql & " con.STRNOME, "
    strSql = strSql & " hev.PKID, "
    strSql = strSql & " hev.INTECONOMICO, "
    strSql = strSql & " hev.BYTTIPOHISTORICO, "
    strSql = strSql & " hev.DTMDTINICIAL, "
    strSql = strSql & " hev.DTMDTFINAL, "
    strSql = strSql & " hev.STRDESCRICAO, "
    strSql = strSql & " hev.DTMDTATUALIZACAO, "
    strSql = strSql & " hev.LNGCODUSR, "
    strSql = strSql & " hev.INTIDENTIFICACAO "
    strSql = strSql & "FROM tblContribuinte con, "
    strSql = strSql & " tblEconomico eco, "
    strSql = strSql & " tblHistoricoEconVariavel hev "
    strSql = strSql & "WHERE eco.Pkid = hev.intEconomico "
    strSql = strSql & "AND eco.INTCONTRIBUINTE = con.PKID "
    strSql = strSql & "AND hev.BYTTIPOHISTORICO = 5 "
    strSql = strSql & "AND " & gstrRIGHT("eco.STRINSCRICAOCADASTRAL", gintRetornaTamanhoMascara(TYP_ECONOMICA))
    strSql = strSql & "BETWEEN '" & dbcstrInscricaoInicial.Text & "'"
    strSql = strSql & "    AND '" & dbcstrInscricaoFinal.Text & "'"
    strSql = strSql & "ORDER BY STRINSCRICAOCADASTRAL "
   
    strQueryRelatorio = strSql

End Function

Private Function blnDadosOk() As Boolean
    
    blnDadosOk = False
    
    If Not dbcstrInscricaoInicial.MatchedWithList Then
        ExibeMensagem "A Inscrição Inicial não está preenchida corretamente."
        dbcstrInscricaoInicial.SetFocus
        Exit Function
    End If
    
    If Not dbcstrInscricaoFinal.MatchedWithList Then
        ExibeMensagem "A Inscrição Final não está preenchida corretamente."
        dbcstrInscricaoFinal.SetFocus
        Exit Function
    End If
    
    blnDadosOk = True
    
End Function

Public Sub MantemForm(ByVal strModoOperacao As String)

    Select Case strModoOperacao
        Case UCase(gstrPreencherLista)
            If Me.ActiveControl.Name = dbcstrInscricaoInicial.Name Then
                LeDaTabelaParaObj "", Me.ActiveControl, strQueryInscricaoInicial
            Else
                PreencherListaDeOpcoes Me.ActiveControl
            End If
            'PreencherListaDeOpcoes Me.ActiveControl
        
        Case UCase(gstrImprimir)
            If blnDadosOk Then
                ImprimeRelatorio rptOcorrenciasDoEconomico, strQueryRelatorio, "Ocorrências do Econômico"
            End If
        Case UCase(gstrNovo)
            Set dbcstrInscricaoInicial.RowSource = Nothing
            Set dbcstrInscricaoFinal.RowSource = Nothing
            dbcstrInscricaoInicial.Text = ""
            dbcstrInscricaoFinal.Text = ""
            dbcstrInscricaoInicial.SetFocus
    End Select
    
End Sub


Private Sub dbcstrInscricaoFinal_Click(Area As Integer)

DropDownDataCombo dbcstrInscricaoInicial, Me, Area

End Sub

Private Sub dbcstrInscricaoFinal_GotFocus()
    MarcaCampo dbcstrInscricaoFinal
End Sub

Private Sub dbcstrInscricaoInicial_Change()

    If dbcstrInscricaoInicial.MatchedWithList Then
        Set dbcstrInscricaoFinal.RowSource = Nothing
        dbcstrInscricaoFinal.BoundText = ""
    End If
    
End Sub

Private Sub dbcstrInscricaoInicial_Click(Area As Integer)
     
    DropDownDataCombo dbcstrInscricaoInicial, Me, Area
     
    If dbcstrInscricaoInicial.MatchedWithList Then
        Set dbcstrInscricaoFinal.RowSource = Nothing
        dbcstrInscricaoFinal.BoundText = ""
    End If
    
End Sub

Private Sub dbcstrInscricaoInicial_GotFocus()
    MarcaCampo dbcstrInscricaoInicial
End Sub

Private Sub dbcstrInscricaoInicial_LostFocus()

    If dbcstrInscricaoInicial.MatchedWithList Then
        PreencherListaDeOpcoes dbcstrInscricaoFinal
        dbcstrInscricaoFinal.Text = dbcstrInscricaoInicial.Text
        dbcstrInscricaoFinal.SelStart = 0
        dbcstrInscricaoFinal.SelLength = Len(dbcstrInscricaoFinal.Text)
    End If

End Sub

Private Sub Form_Load()
    dbcstrInscricaoInicial.Tag = strQueryInscricaoInicial & " ;eco.strInscricaoCadastral"
    dbcstrInscricaoFinal.Tag = strQueryInscricaoFinal & " ;eco.strInscricaoCadastral"
End Sub





