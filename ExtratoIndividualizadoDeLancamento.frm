VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmExtratoIndividualizadoDeLancamento 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Extrato Individualizado de Lançamento"
   ClientHeight    =   3105
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7410
   Icon            =   "ExtratoIndividualizadoDeLancamento.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3105
   ScaleWidth      =   7410
   Begin TabDlg.SSTab tab_3dPasta 
      Height          =   2865
      Left            =   120
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   120
      Width           =   7155
      _ExtentX        =   12621
      _ExtentY        =   5054
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Extrato Individualizado de Lançamento"
      TabPicture(0)   =   "ExtratoIndividualizadoDeLancamento.frx":1042
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin VB.Frame Frame1 
         Height          =   2265
         Left            =   180
         TabIndex        =   5
         Top             =   420
         Width           =   6795
         Begin VB.Frame Frame2 
            Caption         =   "Datas"
            Height          =   675
            Left            =   1590
            TabIndex        =   9
            Top             =   1380
            Width           =   5055
            Begin VB.TextBox txtDtFinal 
               Height          =   285
               Left            =   3840
               MaxLength       =   11
               TabIndex        =   3
               Top             =   240
               Width           =   1065
            End
            Begin VB.TextBox txtDtInicial 
               Height          =   285
               Left            =   900
               MaxLength       =   11
               TabIndex        =   2
               Top             =   240
               Width           =   1065
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "Final"
               Height          =   195
               Left            =   3360
               TabIndex        =   11
               Top             =   330
               Width           =   330
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Inicial"
               Height          =   195
               Left            =   360
               TabIndex        =   10
               Top             =   330
               Width           =   405
            End
         End
         Begin VB.CheckBox chk_Selecionar 
            Caption         =   "Selecionar todos os Contribuintes"
            Height          =   255
            Left            =   1590
            TabIndex        =   6
            Top             =   1020
            Width           =   2835
         End
         Begin MSDataListLib.DataCombo dbcintContribuinteInicial 
            Height          =   315
            Left            =   1590
            TabIndex        =   0
            Top             =   240
            Width           =   5055
            _ExtentX        =   8916
            _ExtentY        =   556
            _Version        =   393216
            IntegralHeight  =   0   'False
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dbcintContribuinteFinal 
            Height          =   315
            Left            =   1590
            TabIndex        =   1
            Top             =   660
            Width           =   5055
            _ExtentX        =   8916
            _ExtentY        =   556
            _Version        =   393216
            IntegralHeight  =   0   'False
            Text            =   ""
         End
         Begin VB.Label lbl_label2 
            AutoSize        =   -1  'True
            Caption         =   "Contribuinte Final"
            Height          =   195
            Left            =   225
            TabIndex        =   8
            Top             =   765
            Width           =   1215
         End
         Begin VB.Label lbl_Label1 
            AutoSize        =   -1  'True
            Caption         =   "Contribuinte Inicial"
            Height          =   195
            Left            =   150
            TabIndex        =   7
            Top             =   345
            Width           =   1290
         End
      End
   End
End
Attribute VB_Name = "frmExtratoIndividualizadoDeLancamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mblnAlterando           As Boolean
Dim mobjAux                 As Object
Dim mblnSelecionou          As Boolean
Dim mblnPrimeiraVez         As Boolean
Dim intCodigoInicial        As Integer
Dim intCodigoFinal          As Integer
Dim CCInicial               As Integer
Dim CCFinal                 As Integer
Dim TipoDeInscricao         As Integer

    'NADA NADA

Private Sub chk_Selecionar_Click()
    If chk_Selecionar.Value = 1 Then
        dbcintContribuinteInicial.BoundText = ""
        dbcintContribuinteFinal.BoundText = ""
        dbcintContribuinteInicial.Enabled = False
        TrocaCorObjeto dbcintContribuinteInicial, True
        dbcintContribuinteFinal.Enabled = False
        TrocaCorObjeto dbcintContribuinteFinal, True
    Else
        dbcintContribuinteInicial.Enabled = True
        TrocaCorObjeto dbcintContribuinteInicial, False
        dbcintContribuinteFinal.Enabled = True
        TrocaCorObjeto dbcintContribuinteFinal, False
    End If
End Sub

Private Sub dbcintContribuinteFinal_Click(Area As Integer)
    DropDownDataCombo dbcintContribuinteFinal, Me, Area
    If Area = 0 Then
        If Trim(dbcintContribuinteFinal.Text) <> "" And Not dbcintContribuinteFinal.MatchedWithList Then
            MantemForm gstrPreencherLista
        End If
    End If
End Sub

Private Sub dbcintContribuinteFinal_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbcintContribuinteFinal, Me, , KeyCode, Shift
End Sub

Private Sub dbcintContribuinteInicial_Click(Area As Integer)
    DropDownDataCombo dbcintContribuinteInicial, Me, Area
    If Area = 0 Then
        If Trim(dbcintContribuinteInicial.Text) <> "" And Not dbcintContribuinteInicial.MatchedWithList Then
            MantemForm gstrPreencherLista
        End If
    End If
End Sub

Private Sub dbcintContribuinteInicial_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbcintContribuinteInicial, Me, , KeyCode, Shift
End Sub

Private Sub Form_Activate()
    gintCodSeguranca = 458
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
    CCInicial = 0
    CCFinal = 0
    
    dbcintContribuinteInicial.Tag = strQuerryContribuinte & ";strNome"
    dbcintContribuinteFinal.Tag = dbcintContribuinteInicial.Tag
    
    txtDtFinal.Text = gstrDataDoSistema
    HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrSalvar, gstrDeletar
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar
    mblnSelecionou = False
    mblnPrimeiraVez = False
End Sub

Public Sub MantemForm(ByVal strModoOperacao As String)

'******************************************************************************************
' Data: 09/05/2003
' Alteração: - Substituição da chamada direta à stored procedure pela função
'            gstrStoredProcedure.
' Responsável: Everton Bianchini
'******************************************************************************************

Dim strSQL As String

Dim strParameters As String

On Error Resume Next
    If UCase(strModoOperacao) = UCase(gstrImprimir) Then
        If blnDadosOk Then
            If chk_Selecionar.Value = 0 Then
            
'                strSQL = " sp_ExtratoIndividual " & dbcintContribuinteInicial.BoundText & ","
                strSQL = "sp_ExtratoIndividual"
                strParameters = dbcintContribuinteInicial.BoundText & ","
                
'                strSQL = strSQL & IIf(Trim(dbcintContribuinteFinal.BoundText) <> "", dbcintContribuinteFinal.BoundText, dbcintContribuinteInicial.BoundText) & ","
                strParameters = strParameters & IIf(Trim(dbcintContribuinteFinal.BoundText) <> "", dbcintContribuinteFinal.BoundText, dbcintContribuinteInicial.BoundText) & ","
                If Val(dbcintContribuinteInicial.BoundText) < Val(dbcintContribuinteFinal.BoundText) Then
                    CCInicial = Val(dbcintContribuinteInicial.BoundText)
                    CCFinal = Val(dbcintContribuinteFinal.BoundText)
                Else
            
                    CCInicial = Val(dbcintContribuinteFinal.BoundText)
                    CCFinal = Val(dbcintContribuinteInicial.BoundText)
                End If
            Else
'                strSQL = " sp_ExtratoIndividual 0,0,"
                strSQL = "sp_ExtratoIndividual"
                strParameters = "0,0,"
            End If
'            strSQL = strSQL & gstrConvDtParaSql(txtDtInicial.Text) & ","
            strParameters = strParameters & gstrConvDtParaSql(txtDtInicial.Text) & ","
'            strSQL = strSQL & gstrConvDtParaSql(txtDtFinal.Text)
            strParameters = strParameters & gstrConvDtParaSql(txtDtFinal.Text)
            
            strSQL = gstrStoredProcedure(strSQL, strParameters, True)
            
            ImprimeRelatorio rptExtratoIndividualizadoDeLancamento, strSQL
        End If
    End If
    
    If UCase(strModoOperacao) = UCase(gstrNovo) Then
        LimpaObjetos
    End If
    
    If UCase(strModoOperacao) = UCase(gstrFechar) Then
        Unload Me
    End If
    
    If UCase(strModoOperacao) = UCase(gstrPreencherLista) And TypeOf ActiveControl Is DataCombo Then
        strSQL = ""
        strSQL = strSQL & " SELECT PKId, strNome FROM " & gstrContribuinte
        If IsNumeric(ActiveControl.Text) Then
            strSQL = strSQL & " WHERE strCodigoAnterior = '" & ActiveControl.Text & "'"
        Else
            strSQL = strSQL & " WHERE strNome LIKE '" & ActiveControl.Text & "%'"
        End If
        LeDaTabelaParaObj "", ActiveControl, strSQL
    End If
    
    
End Sub

Private Sub LimpaObjetos()
    dbcintContribuinteInicial.BoundText = ""
    dbcintContribuinteFinal.BoundText = ""
    txtDtInicial.Text = ""
    txtDtFinal.Text = ""
    dbcintContribuinteInicial.SetFocus
End Sub

Private Function blnDadosOk() As Boolean
blnDadosOk = False
On Error GoTo err_blnDadosOK
    If chk_Selecionar.Value = 0 Then
        If dbcintContribuinteInicial.Text = "" Then
           ExibeMensagem "O Contribuinte Incial tem que ser selecionado."
           dbcintContribuinteInicial.SetFocus
           Exit Function
        End If
    End If
    
    If gblnDataValida(txtDtInicial.Text) = False Then
        ExibeMensagem "A data inicial não é uma data válida."
        txtDtInicial.SetFocus
        Exit Function
    End If
    
    If gblnDataValida(txtDtFinal.Text) = False Then
        ExibeMensagem "A data final não é uma data válida."
        txtDtFinal.SetFocus
        Exit Function
    End If
    
    If CVDate(txtDtInicial.Text) > CVDate(txtDtFinal.Text) Then
        ExibeMensagem "A data inicial não pode ser maior que a data final."
        txtDtInicial.SetFocus
        Exit Function
    End If

blnDadosOk = True
err_blnDadosOK:
End Function

Private Function strQuerryContribuinte() As String
Dim strSQL As String
    strSQL = ""
    strSQL = strSQL & " SELECT "
    strSQL = strSQL & " PKId, strNome "
    strSQL = strSQL & " FROM "
    strSQL = strSQL & gstrContribuinte
    strSQL = strSQL & " ORDER BY strNome "
strQuerryContribuinte = strSQL
End Function

Private Sub dbcintContribuinteFinal_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbcintContribuinteFinal
End Sub

Private Sub dbcintContribuinteInicial_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbcintContribuinteInicial
End Sub

Private Sub txtDtInicial_GotFocus()
    MarcaCampo txtDtInicial
End Sub

Private Sub txtDtInicial_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "D", txtDtInicial
End Sub

Private Sub txtDtFinal_GotFocus()
    MarcaCampo txtDtFinal
End Sub

Private Sub txtDtFinal_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "D", txtDtFinal
End Sub
