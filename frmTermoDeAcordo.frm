VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmTermoDeAcordo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Termo de Acordo"
   ClientHeight    =   1185
   ClientLeft      =   4575
   ClientTop       =   3855
   ClientWidth     =   3525
   Icon            =   "frmTermoDeAcordo.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1185
   ScaleWidth      =   3525
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fra_Parametros 
      Height          =   1065
      Left            =   60
      TabIndex        =   0
      Top             =   30
      Width           =   3390
      Begin MSDataListLib.DataCombo dbc_strInscricao 
         Height          =   315
         Left            =   870
         TabIndex        =   2
         Top             =   450
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
         _Version        =   393216
         IntegralHeight  =   0   'False
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dbc_intExercicio 
         Height          =   315
         Left            =   2220
         TabIndex        =   3
         Top             =   450
         Width           =   990
         _ExtentX        =   1746
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin VB.Label lblNumero 
         AutoSize        =   -1  'True
         Caption         =   "Nº Termo"
         Height          =   195
         Left            =   90
         TabIndex        =   1
         Top             =   510
         Width           =   675
      End
   End
End
Attribute VB_Name = "frmTermoDeAcordo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lngPkid                 As Long
Dim lngPkidAlfaAcordo       As Long
    
Private Sub dbc_intExercicio_Click(Area As Integer)
    DropDownDataCombo dbc_intExercicio, Me, Area
End Sub

Private Sub dbc_intExercicio_GotFocus()
    MarcaCampo dbc_intExercicio
End Sub

Private Sub dbc_intExercicio_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbc_intExercicio, Me, , KeyCode, Shift
End Sub

Private Sub dbc_intExercicio_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", dbc_intExercicio
End Sub

Private Sub dbc_strInscricao_Change()
    If dbc_strInscricao.MatchedWithList And dbc_strInscricao.BoundText <> "" Then
        PreencheExercicio
    Else
        dbc_intExercicio.Text = ""
        dbc_intExercicio.ListField = ""
    End If
End Sub

Private Sub dbc_strInscricao_GotFocus()
    MarcaCampo dbc_strInscricao
End Sub

Private Sub dbc_strInscricao_Click(Area As Integer)
    If Area = 0 Then DropDownDataCombo dbc_strInscricao, Me, Area
End Sub

Private Sub Form_Activate()
    gintCodSeguranca = 1180
    HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrImprimir
End Sub

Public Sub MantemForm(ByVal strModoOperacao As String)
    If UCase(strModoOperacao) = UCase(gstrPreencherLista) Then
        If Me.ActiveControl.Name = "dbc_strInscricao" Then
           LeDaTabelaParaObj "", Me.ActiveControl, strQueryComboAcordo(Trim(Me.ActiveControl.Text))
        End If
    ElseIf UCase(strModoOperacao) = gstrImprimir Then
        If blnDadosOk Then
            ImprimirTermo lngPkid
        End If
    ElseIf UCase(strModoOperacao) = gstrNovo Then
        dbc_strInscricao.Text = ""
        dbc_strInscricao.ListField = ""
        dbc_intExercicio.Text = ""
        dbc_intExercicio.ListField = ""
    End If
End Sub

Private Function blnDadosOk()
    blnDadosOk = False
    If Not dbc_strInscricao.MatchedWithList Then
        ExibeMensagem "O número da inscrição deve ser preenchido."
        dbc_strInscricao.SetFocus
        Exit Function
    ElseIf Not dbc_intExercicio.MatchedWithList Then
        ExibeMensagem "O ano deve ser preenchido."
        dbc_intExercicio.SetFocus
        Exit Function
    ElseIf Not blnInscricao Then
        ExibeMensagem "Essa inscrição não é válida."
        dbc_strInscricao.SetFocus
        Exit Function
    End If
    blnDadosOk = True
End Function

Private Function strQueryComboAcordo(Inscricao As String) As String
Dim strSql As String

  strSql = ""
  strSql = strSql & "SELECT "
  strSql = strSql & "LA.pkID, "
  strSql = strSql & strSUBSTRING & "(LA.strInscricao, " & gintLenInscricao - gintRetornaTamanhoMascara(TYP_ACORDO) + 1 & ", " & gintRetornaTamanhoMascara(TYP_ACORDO) - 4 & ") strInscricao "
  strSql = strSql & "FROM "
  strSql = strSql & gstrLancamentoAlfa & " LA, "
  strSql = strSql & gstrAcordo & " AC "
  strSql = strSql & "WHERE "
  strSql = strSql & "LA.pkID = AC.INTLANCAMENTOALFA "
  
  If Inscricao <> "" Then
     strSql = strSql & "AND LA.strInscricao Like '" & String(16 - Len(Inscricao), "0") & Inscricao & "%' "
  End If
  
  strSql = strSql & "ORDER BY strInscricao"

  strQueryComboAcordo = strSql
  
End Function

'Private Function strQueryInscricao() As String
'Dim strSql As String
'
'    strSql = "SELECT LA.Pkid, " & strSUBSTRING & "(" & gstrRIGHT("LA.strInscricao", gintRetornaTamanhoMascara(TYP_ACORDO)) & ",1," & gintRetornaTamanhoMascara(TYP_ACORDO) - 4 & " ) strInscricao "
'    strSql = strSql & " FROM "
'    strSql = strSql & gstrLancamentoAlfa & " LA "
'    strSql = strSql & " WHERE LA.intUtilizacao = " & TYP_ACORDO
'    strSql = strSql & " ORDER BY strInscricao"
'
'    strQueryInscricao = strSql
'
'End Function

'Private Sub Form_Load()
'    dbc_strInscricao.Tag = strQueryInscricao & ";strInscricao"
'End Sub

Private Function blnInscricao() As Boolean
Dim strSql          As String
Dim adoResultado    As ADODB.Recordset
    
    blnInscricao = False
    strSql = "SELECT Distinct "
    strSql = strSql & "LA.Pkid, "
    strSql = strSql & "LA.strInscricao NumeroInscricao, "
    strSql = strSql & "LA.intExercicio Exercicio "
'    strSql = strSql & "LV.IntlancamentoAlfaAcordo "
    strSql = strSql & "FROM "
    strSql = strSql & gstrLancamentoAlfa & " LA, "
'    strSql = strSql & gstrLancamentoValor & " LV, "
    strSql = strSql & gstrAcordo & " AC "
    strSql = strSql & "WHERE "
    strSql = strSql & "AC.intLancamentoAlfa = LA.Pkid AND "
'    strSql = strSql & "LA.Pkid = LV.Intlancamentoalfa AND "
    'strSql = strSql & "Not LV.IntlancamentoalfaAcordo is Null AND "
    strSql = strSql & "LA.Intutilizacao = " & TYP_ACORDO & " AND "
    strSql = strSql & "LA.strInscricao = '" & String(gintLenInscricao - Len(dbc_strInscricao.Text & dbc_intExercicio.Text), "0") & dbc_strInscricao.Text & dbc_intExercicio.Text & "'"
    strSql = strSql & " ORDER BY LA.strInscricao"
    
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        If adoResultado.EOF Then
            Exit Function
        Else
            lngPkid = gstrENulo(adoResultado!Pkid)
            'lngPkidAlfaAcordo = gstrENulo(adoResultado!IntlancamentoAlfaAcordo)
        End If
    End If
    blnInscricao = True
    Set gobjBanco = Nothing
    
End Function

Private Sub PreencheExercicio()
Dim strSql As String
Dim adoResultado As ADODB.Recordset
    
    Screen.MousePointer = vbHourglass
    
    dbc_intExercicio.Text = ""
    dbc_intExercicio.ListField = ""

    strSql = "SELECT DISTINCT " & gstrRIGHT("LA.strInscricao", 4) & " strInscricao "
    strSql = strSql & " FROM "
    strSql = strSql & gstrLancamentoAlfa & " LA "
    strSql = strSql & " WHERE " & strSUBSTRING & "(" & gstrRIGHT("LA.strInscricao", gintRetornaTamanhoMascara(TYP_ACORDO)) & ",1," & gintRetornaTamanhoMascara(TYP_ACORDO) - 4 & " ) = '" & dbc_strInscricao.Text & "'"
    strSql = strSql & " AND LA.intUtilizacao = " & TYP_ACORDO
    strSql = strSql & " ORDER BY " & gstrRIGHT("LA.strInscricao", 4)

    Set gobjBanco = New clsBanco

    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        If Not adoResultado.EOF Then
            dbc_intExercicio.ListField = adoResultado.Fields(0).Name
            Set dbc_intExercicio.RowSource = adoResultado
        End If
    End If
    
    Screen.MousePointer = vbDefault
    
End Sub

