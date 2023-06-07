VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmDocCadastroMobiliario 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Certidão Mobiliaria"
   ClientHeight    =   885
   ClientLeft      =   1560
   ClientTop       =   1455
   ClientWidth     =   2700
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   885
   ScaleWidth      =   2700
   Begin TabDlg.SSTab tab_3DPasta 
      Height          =   855
      Left            =   30
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   30
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   1508
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   529
      TabCaption(0)   =   "Inscrição Cadastral"
      TabPicture(0)   =   "frmRelCadastroMobiliario.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "dbc_Inicio"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin MSDataListLib.DataCombo dbc_Inicio 
         Height          =   315
         Left            =   60
         TabIndex        =   1
         Top             =   390
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   556
         _Version        =   393216
         IntegralHeight  =   0   'False
         Text            =   ""
      End
   End
End
Attribute VB_Name = "frmDocCadastroMobiliario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSql                  As String
Dim lngPkid                 As Long
Dim XArrayAlinhaColunas     As XArrayDB
Dim XValores                As XArrayDB

Private Sub dbc_Inicio_GotFocus()
    MarcaCampo dbc_Inicio
End Sub

Private Sub dbc_Inicio_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbc_Inicio
End Sub

Private Sub Form_Activate()
    gintCodSeguranca = 1186
    HabilitaDesabilitaBotao1 True, gstrNovo, gstrImprimir, gstrPreencherLista
End Sub

Private Sub Form_Load()
    dbc_Inicio.Tag = strQueryInscricaoCadastral & ";strInscricaoCadastral"
End Sub

Private Sub ImprimirTermo()
    Dim strnumero           As String
    Dim adoResultado        As ADODB.Recordset
    Dim adoNumero           As ADODB.Recordset
    
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strQuery, 5, adoResultado) Then
        If Not adoResultado.EOF Then
            With adoResultado
                lngPkid = gstrENulo(!Pkid)
                PreencheCampos lngPkid
                
                'Query utilizada para pegar o Codigo Mobiliario da tblEmpresa
                strSql = ""
                strSql = strSql & "Select Max(intnumerocertidaocadmobiliario) + 1 as Numero From " & gstrEmpresa
                
                Set gobjBanco = New clsBanco
                If gobjBanco.CriaADO(strSql, 5, adoNumero) Then
                    If Not adoResultado.EOF Then
                        If Val(gstrENulo(adoNumero!Numero)) > "0" Then
                            strnumero = gstrENulo(adoNumero!Numero)
                        Else
                            strnumero = 1
                        End If
                        gobjBanco.Execute ("Update " & gstrEmpresa & " Set intnumerocertidaocadmobiliario = " & strnumero)
                    Else
                        Exit Sub
                    End If
                Else
                    Exit Sub
                End If
                
                AlinhaCampos
                OpenWordDocumentCertidaoMobiliario strnumero, gstrENulo(!InsCad), gstrENulo(!RazaoSocial), _
                gstrENulo(!DataCadastro), gstrENulo(!Sigla), gstrENulo(!Logradouro), gstrENulo(!Num), _
                gstrENulo(!Bairro), gstrENulo(!Processo), "", XValores, XArrayAlinhaColunas
            End With
        Else
            ExibeMensagem "Nada foi encontrado com nesse intervalo de Inscrições"
        End If
    End If
    Set gobjBanco = Nothing
End Sub

Private Sub PreencheCampos(intPkid As Long)
    Dim intPosition     As Integer
    Dim varAux          As Variant
    Dim adoResultado    As ADODB.Recordset
    
    Set XValores = New XArrayDB
    XValores.Clear
    
    strSql = ""
    strSql = strSql & "SELECT "
    strSql = strSql & "AEC.Intcodigo CodAtiv, "
    strSql = strSql & "AEM.Blnprincipal Primario, "
    strSql = strSql & "AEC.Strdescricao Atividade, "
    strSql = strSql & "CO.dtmdatacadastro strdatacadastro"
    strSql = strSql & " FROM "
    strSql = strSql & gstrEconomico & " EC, "
    strSql = strSql & gstrContribuinte & " CO, "
    strSql = strSql & gstrAtividadeDaEmpresa & " AEM, "
    strSql = strSql & gstrAtividadeEC & " AEC"
    strSql = strSql & " WHERE "
    strSql = strSql & "EC.PKID = " & intPkid & " and "
    strSql = strSql & "EC.Intcontribuinte = CO.Pkid and "
    strSql = strSql & "AEM.INTECONOMICO = EC.Pkid and "
    strSql = strSql & "AEM.Intatividade = AEC.pkid"
    strSql = strSql & " ORDER BY "
    strSql = strSql & "AEM.blnPrincipal Desc"
    intPosition = 0
    
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        If Not adoResultado.EOF Then
            With adoResultado
                Do While Not .EOF
                    XValores.ReDim 0, intPosition, 0, 3
                    varAux = gstrENulo(!CodAtiv)
                    XValores(intPosition, 0) = varAux
                    varAux = gstrENulo(!Primario)
                    If varAux = 1 Then
                        varAux = "P"
                    Else
                        varAux = "S"
                    End If
                    XValores(intPosition, 1) = varAux
                    varAux = gstrENulo(!Atividade)
                    XValores(intPosition, 2) = varAux
                    varAux = gstrENulo(!strdatacadastro)
                    XValores(intPosition, 3) = varAux
                    intPosition = intPosition + 1
                    .MoveNext
                Loop
            End With
        End If
    End If
End Sub

Private Function strQuery() As String
    
    strSql = ""
    
    strSql = strSql & "SELECT "
    strSql = strSql & "EC.PKID Pkid, "
    strSql = strSql & gstrRIGHT("EC.Strinscricaocadastral", gintRetornaTamanhoMascara(TYP_ECONOMICA)) & " InsCad, "
    strSql = strSql & "CO.Strnome RazaoSocial, "
    strSql = strSql & "CO.Dtmdatacadastro DataCadastro, "
    strSql = strSql & "TL.Strsigla Sigla, "
    strSql = strSql & "LG.Strdescricao Logradouro, "
    strSql = strSql & "EC.Intnumero Num, "
    strSql = strSql & "BA.strDescricao Bairro, "
    strSql = strSql & "PP.strCodigo  " & strCONCAT & "'/'" & strCONCAT & " TO_CHAR(PP.intExercicio )  " & strCONCAT & "'-'" & strCONCAT & "  TO_CHAR(PP.bitDigito) Processo "
    strSql = strSql & " FROM "
    strSql = strSql & gstrEconomico & " EC, "
    strSql = strSql & gstrContribuinte & " CO, "
    strSql = strSql & gstrTipoLogradouro & " TL, "
    strSql = strSql & gstrLogradouro & " LG, "
    strSql = strSql & gstrBairro & " BA, "
    strSql = strSql & gstrProtocolizacaoProcesso & " PP "
    strSql = strSql & " WHERE "
    strSql = strSql & "EC.strInscricaoCadastral = '" & String(gintLenInscricao - Len(dbc_Inicio.Text), "0") & dbc_Inicio.Text & "' And "
    strSql = strSql & "EC.Intcontribuinte" & strOUTJSQLServer & "= CO.pkid" & strOUTJOracle & " and "
    strSql = strSql & "EC.Intlogradouro" & strOUTJSQLServer & "= LG.Pkid" & strOUTJOracle & " and "
    strSql = strSql & "LG.INTTIPOLOGRADOURO" & strOUTJSQLServer & "= TL.Pkid" & strOUTJOracle & " and "
    strSql = strSql & "EC.Intbairro" & strOUTJSQLServer & "= BA.Pkid" & strOUTJOracle & " And "
    strSql = strSql & "PP.intCodContribuinte" & strOUTJOracle & "=" & strOUTJSQLServer & "CO.Pkid "
    
    strQuery = strSql
      
End Function

Private Function strQueryInscricaoCadastral() As String
    strSql = ""
    strSql = strSql & "SELECT "
    strSql = strSql & "Pkid, "
    strSql = strSql & gstrRIGHT("strInscricaoCadastral", gintRetornaTamanhoMascara(TYP_ECONOMICA)) & " strInscricaoCadastral "
    strSql = strSql & " FROM "
    strSql = strSql & gstrEconomico
    strSql = strSql & " WHERE "
    strSql = strSql & "dtmDataEncerramento IS NULL"
    strSql = strSql & " ORDER BY "
    strSql = strSql & "strInscricaoCadastral"
    strQueryInscricaoCadastral = strSql
End Function

Private Function blnDadosOk()
    blnDadosOk = False
    If Not dbc_Inicio.MatchedWithList Then
        ExibeMensagem "O número da inscrição deve ser selecionado."
        dbc_Inicio.SetFocus
        Exit Function
    End If
    blnDadosOk = True
End Function

Public Sub MantemForm(ByVal strModoOperacao As String)
    If UCase(strModoOperacao) = gstrImprimir Then
        If blnDadosOk Then ImprimirTermo
        ElseIf UCase(strModoOperacao) = gstrNovo Then dbc_Inicio.Text = ""
        ElseIf UCase(strModoOperacao) = gstrPreencherLista Then PreencherListaDeOpcoes Me.ActiveControl
    End If
End Sub

Private Sub AlinhaCampos()
    Set XArrayAlinhaColunas = New XArrayDB
   
    With XArrayAlinhaColunas 'Alinhamento
        .Clear
        .ReDim 0, 0, 0, 3
        .Value(0, 0) = WORDALIGNPARAGRAPHCENTER
        .Value(0, 1) = WORDALIGNPARAGRAPHCENTER
        .Value(0, 2) = WORDALIGNPARAGRAPHCENTER
        .Value(0, 3) = WORDALIGNPARAGRAPHCENTER
    End With
End Sub


