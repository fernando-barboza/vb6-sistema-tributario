VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmGuiaPrecoPublico 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Preço Público"
   ClientHeight    =   1185
   ClientLeft      =   4605
   ClientTop       =   4275
   ClientWidth     =   3510
   Icon            =   "frmGuiaPrecoPublico.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1185
   ScaleWidth      =   3510
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fra_Parametros 
      Height          =   1065
      Left            =   60
      TabIndex        =   0
      Top             =   30
      Width           =   3405
      Begin MSDataListLib.DataCombo dbc_strInscricao 
         Height          =   315
         Left            =   1065
         TabIndex        =   2
         Top             =   420
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
         _Version        =   393216
         IntegralHeight  =   0   'False
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dbc_intExercicio 
         Height          =   315
         Left            =   2400
         TabIndex        =   3
         Top             =   405
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin VB.Label lblNumero 
         AutoSize        =   -1  'True
         Caption         =   "Nº da Guia"
         Height          =   195
         Left            =   210
         TabIndex        =   1
         Top             =   510
         Width           =   780
      End
   End
End
Attribute VB_Name = "frmGuiaPrecoPublico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim lngPkid                 As Long
    Dim lngPkidPPublico         As Long
    
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
    PreencheExercicio
End Sub

Private Sub dbc_strInscricao_GotFocus()
    MarcaCampo dbc_strInscricao
End Sub

Private Sub dbc_strInscricao_Click(Area As Integer)
    If Area = 0 Then DropDownDataCombo dbc_strInscricao, Me, Area
End Sub

Private Sub Form_Activate()
    gintCodSeguranca = 1210
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrDeletar
End Sub

Private Sub Form_Load()
    dbc_strInscricao.Tag = strQueryInscricao & ";strInscricao"
End Sub

Public Sub MantemForm(ByVal strModoOperacao As String)
    If UCase(strModoOperacao) = UCase(gstrPreencherLista) Then
        PreencherListaDeOpcoes Me.ActiveControl
    ElseIf UCase(strModoOperacao) = gstrImprimir Then
        If blnDadosOk Then
            ImprimirTermo
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

Private Function blnInscricao() As Boolean
    Dim strSql          As String
    Dim adoResultado    As ADODB.Recordset
    
    blnInscricao = False
    
    strSql = strSql & "Select "
    strSql = strSql & "LA.Pkid, "
    strSql = strSql & "LP.Pkid AS intPPublico "
    strSql = strSql & "From "
    strSql = strSql & gstrLancamentoAlfa & " LA, "
    strSql = strSql & gstrComposicaoDaReceita & " CR, "
    strSql = strSql & gstrLancamentoPPublico & " LP "
    strSql = strSql & "Where "
    strSql = strSql & "CR.Pkid = LA.Intcomposicaodareceita AND "
    strSql = strSql & "LA.Pkid = LP.INTLANCAMENTOALFA AND "
    strSql = strSql & "CR.Intutilizacao = " & TYP_PRECO_PUBLICO & " AND "
    strSql = strSql & "LA.strInscricao = '" & String(gintLenInscricao - Len(dbc_strInscricao.Text & dbc_intExercicio.Text), "0") & dbc_strInscricao.Text & dbc_intExercicio.Text & "' AND "
    strSql = strSql & "LA.Intexercicio = " & dbc_intExercicio.Text
    
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        If adoResultado.EOF Then
            Exit Function
        Else
            lngPkid = gstrENulo(adoResultado!Pkid)
            lngPkidPPublico = gstrENulo(adoResultado!intPPublico)
        End If
    End If
    blnInscricao = True
    Set gobjBanco = Nothing
    
End Function

Private Sub ImprimirTermo()
    Dim strSql              As String
    Dim adoResultado        As ADODB.Recordset
    Dim VetWordParcelas()   As String
    
    ReDim VetWordParcelas(12, 0)
    
    'Vamos pegar o Contribuinte requerente do " Preço Público "
    
    strSql = "Select "
    strSql = strSql & strSUBSTRING & "(" & gstrRIGHT("LA.strInscricao", gintRetornaTamanhoMascara(TYP_PRECO_PUBLICO)) & ",1," & gintRetornaTamanhoMascara(TYP_PRECO_PUBLICO) - 4 & " ) " & strCONCAT & "'/'" & strCONCAT & gstrRIGHT("LA.strInscricao", 4) & " AS strInscricao, "
    strSql = strSql & "LA.strnumeroAviso, "
    strSql = strSql & "LA.Strcomposicaodareceita, "
    strSql = strSql & "LA.STRNOMEPROPRIETARIO, "
    strSql = strSql & "LA.Strcnpjcpf, "
    strSql = strSql & "LA.STRIDENTIDADE, "
    strSql = strSql & "Ltrim(Rtrim(LA.STRLOGRADOURO)) " & strCONCAT & " ',' " & strCONCAT & " Ltrim(Rtrim(LA.STRNUMERO)) " & strCONCAT & " ' ' " & strCONCAT & " Ltrim(Rtrim(LA.Strcomplemento)) " & strCONCAT & " ' CEP: ' " & strCONCAT & " Ltrim(Rtrim(LA.INTCEP)) AS strLogradouro, "
    strSql = strSql & "Ltrim(Rtrim(LA.STRBAIRRO)) " & strCONCAT & " ' ' " & strCONCAT & " Ltrim(Rtrim(LA.STRMUNICIPIO)) " & strCONCAT & " ' ' " & strCONCAT & " Ltrim(Rtrim(LA.Struf)) As strBairro, "
    strSql = strSql & "LA.Intexercicio as exercicio, "
    strSql = strSql & "PP.STRIDCCONTABANCARIA, "
    strSql = strSql & "PP.DTMDTVENCIMENTO, "
    strSql = strSql & gstrISNULL("PP.Dblvalor", "0") & ", "
    strSql = strSql & gstrISNULL("PP.DBLMULTA", "0") & ", "
    strSql = strSql & gstrISNULL("PP.Dblcorrecaomonet", "0") & ", "
    strSql = strSql & gstrISNULL("PP.Dbljuros", "0") & " "
    strSql = strSql & "From "
    strSql = strSql & gstrLancamentoAlfa & " LA "
    strSql = strSql & "Where "
    strSql = strSql & "LA.Pkid = " & lngPkid
    
    Set gobjBanco = New clsBanco
    Set adoResultado = New ADODB.Recordset
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        If Not adoResultado.EOF Then
            With adoResultado
                        
            ReDim Preserve VetWordParcelas(12, UBound(VetWordParcelas, 2) + 1)
            
            'VetWordParcelas(0, intContador) = vetParcelas(0, contParcela)                      'Pkid tblLancamentoValor
            'VetWordParcelas(1, intContador) = gstrConvVrDoSql(vetParcelas(4, contParcela), 2)  'Principal
            'VetWordParcelas(2, intContador) = gstrConvVrDoSql(vetParcelas(5, contParcela), 2)  'Multa
            'VetWordParcelas(3, intContador) = gstrConvVrDoSql(vetParcelas(6, contParcela), 2)  'Juros
            'VetWordParcelas(4, intContador) = gstrConvVrDoSql(vetParcelas(7, contParcela), 2)  'Correcao

                
            End With
        Else
            ExibeMensagem "Não foi possivel imprimir a guia, dados não encontrados."
        End If
    End If
    
    'ImprimePrecoPublico strInscricao,strProprietario,strLogradouro,strBairro,"","",strAviso,StrcompReceita,"",
    
End Sub

Private Function strQueryInscricao() As String
Dim strSql As String
    
    strSql = "SELECT LA.Pkid, " & strSUBSTRING & "(" & gstrRIGHT("LA.strInscricao", gintRetornaTamanhoMascara(TYP_PRECO_PUBLICO)) & ",1," & gintRetornaTamanhoMascara(TYP_PRECO_PUBLICO) - 4 & ")  strInscricao "
    strSql = strSql & " FROM "
    strSql = strSql & gstrLancamentoAlfa & " LA, " & gstrComposicaoDaReceita & " CR "
    strSql = strSql & " WHERE LA.intComposicaoDaReceita = CR.Pkid AND CR.intUtilizacao = " & TYP_PRECO_PUBLICO
    strSql = strSql & " ORDER BY strInscricao"
    
    strQueryInscricao = strSql

End Function


Private Sub PreencheExercicio()
Dim strSql As String
Dim adoResultado As ADODB.Recordset
    
    dbc_intExercicio.Text = ""
    dbc_intExercicio.ListField = ""

    strSql = "SELECT DISTINCT " & gstrRIGHT("LA.strInscricao", 4)
    strSql = strSql & " FROM "
    strSql = strSql & gstrLancamentoAlfa & " LA, " & gstrComposicaoDaReceita & " CR "
    strSql = strSql & " WHERE " & strSUBSTRING & "(" & gstrRIGHT("LA.strInscricao", gintRetornaTamanhoMascara(TYP_PRECO_PUBLICO)) & ",1," & gintRetornaTamanhoMascara(TYP_PRECO_PUBLICO) - 4 & " ) = '" & dbc_strInscricao.Text & "'"
    strSql = strSql & " AND LA.intComposicaoDaReceita = CR.Pkid AND CR.intUtilizacao = " & TYP_PRECO_PUBLICO
    strSql = strSql & " ORDER BY " & gstrRIGHT("LA.strInscricao", 4)

    Set gobjBanco = New clsBanco

    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        If Not adoResultado.EOF Then
            dbc_intExercicio.ListField = adoResultado.Fields(0).Name
            Set dbc_intExercicio.RowSource = adoResultado
        End If
    End If

End Sub


