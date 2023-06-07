VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmRelatorioRolLogradouro 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Rol de Logradouros"
   ClientHeight    =   1665
   ClientLeft      =   4185
   ClientTop       =   2925
   ClientWidth     =   5850
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1665
   ScaleWidth      =   5850
   Begin TabDlg.SSTab tab_3dPasta 
      Height          =   1485
      Left            =   90
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   90
      Width           =   5685
      _ExtentX        =   10028
      _ExtentY        =   2619
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Logradouros"
      TabPicture(0)   =   "frmRelatorioRolLogradouro.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblContaBancaria"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "dbcstrBairro"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "chkTodosBairros"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "SSFrame1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      Begin Threed.SSFrame SSFrame1 
         Height          =   960
         Left            =   4275
         TabIndex        =   4
         Top             =   405
         Width           =   1275
         _Version        =   65536
         _ExtentX        =   2249
         _ExtentY        =   1693
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
         Begin VB.OptionButton optNome 
            Caption         =   "Nome"
            Height          =   285
            Left            =   135
            TabIndex        =   6
            Top             =   540
            Width           =   1050
         End
         Begin VB.OptionButton optCodigo 
            Caption         =   "Código"
            Height          =   285
            Left            =   135
            TabIndex        =   5
            Top             =   270
            Value           =   -1  'True
            Width           =   1050
         End
      End
      Begin VB.CheckBox chkTodosBairros 
         Caption         =   "Selecionar todos os Bairros"
         Height          =   195
         Left            =   690
         TabIndex        =   1
         Top             =   840
         Width           =   2865
      End
      Begin MSDataListLib.DataCombo dbcstrBairro 
         Height          =   315
         Left            =   690
         TabIndex        =   2
         Top             =   480
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin VB.Label lblContaBancaria 
         AutoSize        =   -1  'True
         Caption         =   "Bairro"
         Height          =   195
         Left            =   180
         TabIndex        =   3
         Top             =   570
         Width           =   405
      End
   End
End
Attribute VB_Name = "frmRelatorioRolLogradouro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mblnAlterando    As Boolean
Dim mobjAux          As Object
Dim mblnSelecionou   As Boolean
Dim mblnPrimeiraVez  As Boolean

Private Sub chkTodosBairros_Click()
    TrocaCorObjeto dbcstrBairro, chkTodosBairros.Value
End Sub

Private Sub Form_Activate()
    gintCodSeguranca = 1174
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
    dbcstrBairro.Tag = "SELECT BA.Pkid , BA.strDescricao FROM tblBairro BA;strDescricao"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar
End Sub

Public Sub MantemForm(ByVal strModoOperacao As String)
On Error Resume Next
    If UCase(strModoOperacao) = UCase(gstrImprimir) Then
        If blnDadosOk Then
           ImprimeRelatorio rptRolLogradouros, strQueryRelatorio, "Rol de Logradouros"
        End If
    End If
    
    If UCase(strModoOperacao) = UCase(gstrNovo) Then
        Limpa_Controles frmRelatorioRolLogradouro, True, True, False, True, False
        dbcstrBairro.ListField = ""
    End If
    
    If UCase(strModoOperacao) = UCase(gstrFechar) Then
        Unload Me
    End If
    If UCase(strModoOperacao) = UCase(gstrPreencherLista) Then
       PreencherListaDeOpcoes dbcstrBairro
    End If
    
End Sub

Private Function blnDadosOk() As Boolean
    blnDadosOk = False

    If Trim(dbcstrBairro.Text) = "" And chkTodosBairros.Value = 0 Then
       ExibeMensagem "O bairro deve ser selecionado."
       dbcstrBairro.SetFocus
       Exit Function
    End If
    
    blnDadosOk = True
End Function

Private Function strQueryRelatorio() As String
Dim strSql As String
  strSql = ""
  strSql = strSql & "SELECT LO.strCodigo , LO.strDescricao DescLogradouro ,"
  strSql = strSql & " TP.strDescricao DescTipo , TT.strDescricao DescTitulo ,"
  strSql = strSql & " BA.strDescricao DescBairro, LO.intCep, LO.intBairro IntBairro "
  strSql = strSql & " FROM tblLogradouro LO, tblTipoLogradouro TP,"
  strSql = strSql & " tblTituloLogradouro TT, tblBairro BA"
  strSql = strSql & " WHERE LO.intTipoLogradouro " & strOUTJSQLServer & "= TP.Pkid " & strOUTJOracle & " AND"
  strSql = strSql & " LO.intTituloLogradouro " & strOUTJSQLServer & "= TT.Pkid " & strOUTJOracle & " AND"
  strSql = strSql & " LO.intBairro = BA.Pkid"
  If chkTodosBairros.Value = 0 Then
     strSql = strSql & " AND BA.Pkid = " & Val(dbcstrBairro.BoundText)
  End If

  strSql = strSql & " ORDER BY BA.strDescricao, "
  If optCodigo.Value = True Then
     strSql = strSql & gstrCONVERT(CDT_INT, "LO.strCodigo")
  Else
     strSql = strSql & " LO.strDescricao"
  End If
  strQueryRelatorio = strSql
  
  
End Function
