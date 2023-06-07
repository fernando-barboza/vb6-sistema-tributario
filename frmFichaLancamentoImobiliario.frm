VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Begin VB.Form frmFichaLancamentoImobiliario 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ficha de Lançamento Imobiliário"
   ClientHeight    =   3315
   ClientLeft      =   2430
   ClientTop       =   3285
   ClientWidth     =   6390
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3315
   ScaleWidth      =   6390
   Begin TabDlg.SSTab tab_3dPasta 
      Height          =   3075
      Left            =   120
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   5424
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Ficha de Lançamento Imobiliário"
      TabPicture(0)   =   "frmFichaLancamentoImobiliario.frx":0000
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
         Begin VB.ComboBox cbointExercicio 
            Height          =   315
            Left            =   2040
            TabIndex        =   5
            Top             =   960
            Width           =   1455
         End
         Begin VB.CheckBox chk_TodosExercicios 
            Caption         =   "&Todos os exercícios"
            Height          =   255
            Left            =   2040
            TabIndex        =   8
            Top             =   1800
            Width           =   1815
         End
         Begin MSDataListLib.DataCombo dbcstrInscricao 
            Height          =   315
            Left            =   2040
            TabIndex        =   7
            Top             =   1440
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dbcintComposicao 
            Height          =   315
            Left            =   2040
            TabIndex        =   3
            Top             =   480
            Width           =   3375
            _ExtentX        =   5953
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Composição da Receita:"
            Height          =   195
            Left            =   240
            TabIndex        =   2
            Top             =   540
            Width           =   1740
         End
         Begin VB.Label lblDtMovimento 
            AutoSize        =   -1  'True
            Caption         =   "Inscrição:"
            Height          =   195
            Left            =   1290
            TabIndex        =   6
            Top             =   1500
            Width           =   690
         End
         Begin VB.Label lblAgencia 
            AutoSize        =   -1  'True
            Caption         =   "Exercício:"
            Height          =   195
            Left            =   1260
            TabIndex        =   4
            Top             =   1020
            Width           =   720
         End
      End
   End
End
Attribute VB_Name = "frmFichaLancamentoImobiliario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cbointExercicio_Change()
    If cbointExercicio.ListIndex < 0 Then
       Set dbcstrInscricao.RowSource = Nothing
       dbcstrInscricao.Text = ""
       TrocaCorObjeto dbcstrInscricao, True
    End If
End Sub

Private Sub cbointExercicio_Click()
    If cbointExercicio.ListIndex >= 0 Then
       TrocaCorObjeto dbcstrInscricao, False
       dbcstrInscricao.Tag = strQueryInscricao & "; strInscricao"
    End If
End Sub

Private Sub chk_TodosExercicios_Click()
    If dbcintComposicao.MatchedWithList Then
       If chk_TodosExercicios.Value = 1 Then
          TrocaCorObjeto cbointExercicio, True
          TrocaCorObjeto dbcstrInscricao, False
          dbcstrInscricao.SetFocus
       Else
          If cbointExercicio.ListIndex < 0 Then
             TrocaCorObjeto dbcstrInscricao, True
          End If
          TrocaCorObjeto cbointExercicio, False
          cbointExercicio.SetFocus
       End If
    End If
End Sub

Private Sub dbcintComposicao_Change()
    TrocaCorObjeto cbointExercicio, IIf(dbcintComposicao.MatchedWithList, False, True)
    cbointExercicio.Clear
    TrocaCorObjeto dbcstrInscricao, True
    Set dbcstrInscricao.RowSource = Nothing
    dbcstrInscricao.Text = ""
    chk_TodosExercicios.Value = 0
    
    If dbcintComposicao.MatchedWithList Then
       PreencheExercicio
    End If
End Sub

Private Sub dbcintComposicao_Click(Area As Integer)
    If dbcintComposicao.MatchedWithList And Trim(dbcintComposicao.Text) = "" Then
       TrocaCorObjeto cbointExercicio, False
       TrocaCorObjeto dbcstrInscricao, True
    End If
End Sub

Private Function strQueryInscricao() As String
Dim strSql As String

    strSql = "SELECT PKID, " & gstrRIGHT("strInscricao", gintRetornaTamanhoMascara(TYP_IMOBILIARIA)) & " strInscricao "
    strSql = strSql & " FROM " & gstrLancamentoAlfa
    
    strSql = strSql & " WHERE intComposicaoDaReceita = " & dbcintComposicao.BoundText & _
                      " AND intExercicio = " & cbointExercicio.Text & _
                      " AND Dtmdtcancelamento IS NULL "
                      
    If Trim(dbcstrInscricao.Text) <> "" Then
        strSql = strSql & " AND strInscricao = " & String(gintLenInscricao - Len(dbcstrInscricao.Text), "0") & dbcstrInscricao.Text
    End If
    
    strSql = strSql & " ORDER BY strInscricao "
       
    strQueryInscricao = strSql
       
End Function

Private Sub PreencheExercicio()
Dim strSql As String
Dim adoRec As New ADODB.Recordset


    strSql = "SELECT DISTINCT LA.intExercicio"
    strSql = strSql & " FROM " & gstrLancamentoAlfa & " LA, " & _
                                gstrComposicaoDaReceita & " CR "
    strSql = strSql & " Where LA.intComposicaoDaReceita = CR.Pkid" & _
                      " AND CR.Pkid = " & dbcintComposicao.BoundText & _
                      " AND CR.Intutilizacao = " & TYP_IMOBILIARIA
                      
    strSql = strSql & " ORDER BY LA.intExercicio DESC "
    
    Set gobjBanco = New clsBanco
    
    If gobjBanco.CriaADO(strSql, 10, adoRec) Then
        cbointExercicio.Clear
        While Not adoRec.EOF
            cbointExercicio.AddItem adoRec!intExercicio
            adoRec.MoveNext
        Wend
    End If

    adoRec.Close
    Set adoRec = Nothing
    
    Set gobjBanco = Nothing
       
End Sub

Private Function strQueryRelatorio() As String
Dim strSql As String

    strSql = "SELECT LA.PkId IDLancamentoContabil, LF.strDescricao strFator, LF.dblFator,  la.pkid, LA.Dtmdtcancelamento, " & gstrRIGHT("strInscricao", gintRetornaTamanhoMascara(TYP_IMOBILIARIA)) & " strInscricao, LA.strComposicaoDaReceita, " & _
             "La.Intexercicio, " & gstrCONVERT(CDT_numeric, "La.Strnumeroaviso") & " Strnumeroaviso, LA.strnomeproprietario strProprietario, " & _
             "LA.strPromissario, LA.strlogradouro, la.strnumero, la.intcep, la.strComplemento, la.strtipoisencaoimunidade, la.strdefinicaoisencao, la.Strtipoisencaoimunidade, la.Strtipoisencaoimunidade, la.strDefinicaoIsencao, la.strdefinicaoisencao, LI.strlote, LI.strquadra, LI.strloteamento, " & _
             "LA.strBairro, la.strlogradouroc, la.strnumeroc, la.intcepc, la.strcomplementoc, la.strmunicipioc, la.strufc, " & _
             gstrISNULL("li.dblareaterreno", "0") & " dblareaterreno, " & gstrISNULL("li.dblareaexcedente", "0") & " dblareaexcedente," & _
             gstrISNULL("Li.dblvalorterrenoexcedente", "0") & " dblvalorterrenoexcedente, " & _
             gstrISNULL("li.dblimpostoterreno", "0") & " dblimpostoterreno," & gstrISNULL("li.dblimpostoexcedente", "0") & " dblimpostoexcedente," & _
             gstrISNULL("li.dblvalorvenalterreno", "0") & " dblvalorvenalterreno," & gstrISNULL("lpi.ImpostoPredio", "0") & " ImpostoPredio," & _
             gstrISNULL("lpi.VenalPredio", "0") & " VenalPredio," & gstrISNULL("lpi.AreaPredio", "0") & " AreaPredio"

    strSql = strSql & " FROM " & gstrLancamentoAlfa & " LA, " & _
                                 gstrLancamentoIPTU & " LI, " & _
                                 gstrLancamentoFatores & " LF, " & _
                                "( " & _
                                    "SELECT LP.Intlancamentoiptu, " & _
                                            "SUM(LP.Dblimposto) ImpostoPredio, " & _
                                            "SUM(LP.Dblvalorvenalpredio) VenalPredio, " & _
                                            "SUM(LP.Dblmedidadaarea) AreaPredio " & _
                                    "FROM " & gstrLancamentoPredioIPTU & " LP " & _
                                    "GROUP BY LP.Intlancamentoiptu " & _
                                ") LPI "

    strSql = strSql & " WHERE LA.strInscricao = " & String(gintLenInscricao - Len(Trim(dbcstrInscricao)), "0") & UCase(dbcstrInscricao)
    
    If chk_TodosExercicios = vbUnchecked Then
        strSql = strSql & " AND LA.intExercicio = " & cbointExercicio.Text
    End If
                      
    strSql = strSql & " AND LI.Intlancamentoalfa =LA.pkid " & _
                      " AND LPI.Intlancamentoiptu " & strOUTJOracle & " =" & strOUTJSQLServer & " Li.Pkid " & _
                      " AND LA.Dtmdtcancelamento IS NULL " & _
                      " AND LF.intLancamentoIPTU " & strOUTJOracle & " =" & strOUTJSQLServer & " LI.pkid "

    strQueryRelatorio = strSql
                       
End Function

Private Function strQueryComposicao() As String
Dim strSql As String

    strSql = "SELECT DISTINCT CR.Pkid, CR.Strdescricao "
    
    strSql = strSql & "FROM " & gstrLancamentoAlfa & " LA, " & _
                                gstrComposicaoDaReceita & " CR "
    
    strSql = strSql & "Where LA.intComposicaoDaReceita = CR.Pkid " & _
                      " AND CR.Intutilizacao = " & TYP_IMOBILIARIA
                      
    strSql = strSql & " ORDER BY CR.strDescricao "

    strQueryComposicao = strSql
    
End Function

Private Function blnDadosOk() As Boolean
    
    blnDadosOk = False
    
    If Not dbcintComposicao.MatchedWithList Then
        ExibeMensagem "Preencha corretamente o campo Composição da Receita."
        dbcintComposicao.SetFocus
        Exit Function
    End If
    
    If (chk_TodosExercicios.Value = vbUnchecked) And (cbointExercicio.ListIndex = -1) Then
        ExibeMensagem "Preencha corretamente o campo Exercício."
        cbointExercicio.SetFocus
        Exit Function
    End If
    
    If Not dbcstrInscricao.MatchedWithList Then
        ExibeMensagem "Preencha corretamente o campo Inscrição."
        dbcstrInscricao.SetFocus
        Exit Function
    End If
    
    
    
    blnDadosOk = True
    
End Function

Public Sub MantemForm(ByVal strModoOperacao As String)

    Select Case strModoOperacao
        Case UCase(gstrPreencherLista)
            PreencherListaDeOpcoes Me.ActiveControl
        Case UCase(gstrImprimir)
            If blnDadosOk Then
                ImprimeRelatorio rptFichaLancamentoImobiliario, strQueryRelatorio, "Ficha de Lançamento Imobiliário"
            End If
        Case UCase(gstrNovo)
            Set dbcstrInscricao.RowSource = Nothing
            Set dbcintComposicao.RowSource = Nothing
            dbcstrInscricao.Text = ""
            cbointExercicio.Clear
            chk_TodosExercicios.Value = 0
            TrocaCorObjeto cbointExercicio, True
            TrocaCorObjeto dbcstrInscricao, True
            dbcintComposicao.Text = ""
            dbcintComposicao.SetFocus
    End Select
    
End Sub

Private Sub Form_Load()
    dbcintComposicao.Tag = strQueryComposicao & " ;CR.strDescricao"
    TrocaCorObjeto cbointExercicio, True
    TrocaCorObjeto dbcstrInscricao, True
End Sub

