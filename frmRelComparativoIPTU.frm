VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmRelComparativoIPTU 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Comparativo de Lançamentos de IPTU"
   ClientHeight    =   3720
   ClientLeft      =   3075
   ClientTop       =   2655
   ClientWidth     =   5985
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   5985
   Begin TabDlg.SSTab tab_3dPasta 
      Height          =   3585
      Left            =   60
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   60
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   6324
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Parâmetros"
      TabPicture(0)   =   "frmRelComparativoIPTU.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fra_Comparativo"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      Begin VB.Frame Frame1 
         Caption         =   "Faixa de Inscrições"
         Height          =   1485
         Left            =   120
         TabIndex        =   8
         Top             =   1920
         Width           =   5655
         Begin VB.CheckBox chk_Ordenacao 
            Caption         =   "Ordenação por Índice"
            Height          =   195
            Left            =   1620
            TabIndex        =   16
            Top             =   1080
            Width           =   2385
         End
         Begin MSMask.MaskEdBox mskstrInscricao 
            Height          =   300
            Left            =   1620
            TabIndex        =   6
            Top             =   240
            Width           =   1830
            _ExtentX        =   3228
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   24
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox mskstrInscricao1 
            Height          =   300
            Left            =   1620
            TabIndex        =   7
            Top             =   690
            Width           =   1830
            _ExtentX        =   3228
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   24
            PromptChar      =   " "
         End
         Begin VB.Label lblInicial 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Inicial:"
            Height          =   195
            Left            =   1110
            TabIndex        =   13
            Top             =   300
            Width           =   450
         End
         Begin VB.Label lblFinal 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Final:"
            Height          =   195
            Left            =   1185
            TabIndex        =   12
            Top             =   735
            Width           =   375
         End
      End
      Begin VB.Frame fra_Comparativo 
         Height          =   1485
         Left            =   120
         TabIndex        =   1
         Top             =   390
         Width           =   5655
         Begin VB.CheckBox chk_Negativo 
            Caption         =   "Incluir índice negativo"
            Height          =   195
            Left            =   1620
            TabIndex        =   5
            Top             =   1200
            Width           =   2085
         End
         Begin VB.TextBox txt_intExercicio1 
            Alignment       =   2  'Center
            Height          =   315
            Left            =   3150
            MaxLength       =   4
            TabIndex        =   3
            Top             =   480
            Width           =   915
         End
         Begin VB.TextBox txt_intExercicio 
            Alignment       =   2  'Center
            Height          =   315
            Left            =   1620
            MaxLength       =   4
            TabIndex        =   2
            Top             =   480
            Width           =   915
         End
         Begin VB.TextBox txt_intIndice 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1620
            MaxLength       =   8
            TabIndex        =   4
            Top             =   855
            Width           =   1515
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Ex.: 1,058294 = 5,8294%"
            Height          =   195
            Left            =   3285
            TabIndex        =   15
            Top             =   945
            Width           =   1785
         End
         Begin VB.Label lbl_divisor 
            Caption         =   "/"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2790
            TabIndex        =   14
            Top             =   510
            Width           =   225
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "2º Exercício"
            Height          =   195
            Left            =   3090
            TabIndex        =   11
            Top             =   240
            Width           =   870
         End
         Begin VB.Label lbl_Indice 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Índice para Relação"
            Height          =   195
            Left            =   105
            TabIndex        =   10
            Top             =   945
            Width           =   1440
         End
         Begin VB.Label lbl_Exercicio 
            AutoSize        =   -1  'True
            Caption         =   "1º Exercício"
            Height          =   195
            Left            =   1590
            TabIndex        =   9
            Top             =   240
            Width           =   870
         End
      End
   End
End
Attribute VB_Name = "frmRelComparativoIPTU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim mobjAux          As Object
    Dim mblnSelecionou   As Boolean

Private Sub Form_Load()
    HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrSalvar, gstrDeletar
    VerificaMascaraInscricao
End Sub
Private Sub Form_Activate()
    gintCodSeguranca = 1284
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

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar
    mblnSelecionou = False
End Sub

Public Sub MantemForm(ByVal strModoOperacao As String)
    If UCase(strModoOperacao) = UCase(gstrImprimir) Then
        If blnDadosOk Then
            ImprimeRelatorio rptComparativoIPTU, strQueryRelatorio
            rptComparativoIPTU.lbl_Exercicio = txt_intExercicio
            rptComparativoIPTU.lbl_exercicio1 = txt_intExercicio1
        End If
    End If
    
    If UCase(strModoOperacao) = UCase(gstrNovo) Then
        Limpa_Controles frmRelComparativoIPTU, True, False, True, True, False
        mskstrInscricao.Text = ""
        mskstrInscricao1.Text = ""
        txt_intExercicio.SetFocus
    End If
    
    If UCase(strModoOperacao) = UCase(gstrFechar) Then
        Unload Me
    End If
    
    If UCase(strModoOperacao) = UCase(gstrPreencherLista) Then
        PreencherListaDeOpcoes Me.ActiveControl
    End If
    
End Sub

Private Function strQueryRelatorio() As String
    Dim strSql As String
    
    strSql = "Select "
    strSql = strSql & "LA.Strinscricao, "
    strSql = strSql & "LA.Dblareaterreno, "
    strSql = strSql & "LA1.Dblareaterreno Dblareaterreno1, "
    strSql = strSql & "LA.DBLVALORVENALTERRENO, "
    strSql = strSql & "LA1.DBLVALORVENALTERRENO DBLVALORVENALTERRENO1, "
    strSql = strSql & "LA.DblAreaPredio, "
    strSql = strSql & "LA1.DblAreaPredio DblAreaPredio1, "
    strSql = strSql & "LA.Dblvalorvenalpredio, "
    strSql = strSql & "LA1.Dblvalorvenalpredio Dblvalorvenalpredio1, "
    strSql = strSql & "LA.Dblvalor, "
    strSql = strSql & "LA1.Dblvalor Dblvalor1, "
    strSql = strSql & "Case when LA.Dblvalor > 0 then LA.Dblvalor Else 1 End / "
    strSql = strSql & "Case when LA1.Dblvalor > 0 then LA1.Dblvalor Else 1 End as Indice "
    strSql = strSql & "From "
    '1º
    strSql = strSql & "(Select la.strdefinicaoisencao , la.strtipoisencaoimunidade, la.dtmdtcancelamento, " & gstrRIGHT("LA.strInscricao", gintRetornaTamanhoMascara(TYP_IMOBILIARIA)) & " Strinscricao, " & gstrISNULL("Li.Dblareaterreno", "0") & " Dblareaterreno, "
    strSql = strSql & gstrISNULL("Li.DBLVALORVENALTERRENO", "0") & " DBLVALORVENALTERRENO, " & gstrISNULL("Pi.Dblmedidadaarea", "0") & " As DblAreaPredio, "
    strSql = strSql & gstrISNULL("PI.Dblvalorvenalpredio", "0") & " As Dblvalorvenalpredio, Sum(" & gstrISNULL("LV.Dblvalor", "0") & ") Dblvalor "
    strSql = strSql & "From " & gstrLancamentoAlfa & " La, " & gstrLancamentoIPTU & " LI, "
    strSql = strSql & gstrLancamentoValor & " LV, "
    strSql = strSql & "(Select Pi.Intlancamentoiptu,"
    strSql = strSql & "Sum(" & gstrISNULL("Pi.Dblmedidadaarea", "0") & ") As Dblmedidadaarea, "
    strSql = strSql & "Sum(" & gstrISNULL("PI.Dblvalorvenalpredio", "0") & ") As Dblvalorvenalpredio "
    strSql = strSql & "From " & gstrLancamentoAlfa & " La, " & gstrLancamentoIPTU & " LI, " & gstrLancamentoPredioIPTU & " PI "
    strSql = strSql & "Where LA.Pkid = LI.Intlancamentoalfa And LI.Pkid = PI.Intlancamentoiptu And "
    strSql = strSql & "La.strinscricao Between '" & String(gintLenInscricao - Len(Trim(mskstrInscricao.Text)), "0") & Trim(mskstrInscricao.Text) & "' And '" & String(gintLenInscricao - Len(Trim(mskstrInscricao1.Text)), "0") & Trim(mskstrInscricao1.Text) & "' and "
    strSql = strSql & "LA.Intexercicio = " & txt_intExercicio & " And LA.Dtmdtcancelamento is null Group By Pi.Intlancamentoiptu) PI "
    strSql = strSql & "Where LA.Pkid = LI.Intlancamentoalfa         And LA.Pkid " & strOUTJSQLServer & "= LV.Intlancamentoalfa" & strOUTJOracle & " And "
    strSql = strSql & "LI.Pkid " & strOUTJSQLServer & "= PI.Intlancamentoiptu" & strOUTJOracle & "       And La.strinscricao Between '" & String(gintLenInscricao - Len(Trim(mskstrInscricao.Text)), "0") & Trim(mskstrInscricao.Text) & "' And '" & String(gintLenInscricao - Len(Trim(mskstrInscricao1.Text)), "0") & Trim(mskstrInscricao1.Text) & "' and "
    strSql = strSql & "LA.Intexercicio = " & txt_intExercicio & "                 And LA.Dtmdtcancelamento is null And "
    strSql = strSql & "LV.Bitparcelavalida = 1 "
    strSql = strSql & "Group By la.strdefinicaoisencao , la.strtipoisencaoimunidade, la.dtmdtcancelamento, " & gstrRIGHT("LA.strInscricao", gintRetornaTamanhoMascara(TYP_IMOBILIARIA)) & ", " & gstrISNULL("Li.Dblareaterreno", "0") & ", " & gstrISNULL("Li.DBLVALORVENALTERRENO", "0") & ", " & gstrISNULL("Pi.Dblmedidadaarea", "0") & ", " & gstrISNULL("PI.Dblvalorvenalpredio", "0")
'    strSql = strSql & "Order By Strinscricao) LA, "
    strSql = strSql & ") LA, "
    '2º
    strSql = strSql & "(Select la.strdefinicaoisencao , la.strtipoisencaoimunidade, la.dtmdtcancelamento, " & gstrRIGHT("LA.strInscricao", gintRetornaTamanhoMascara(TYP_IMOBILIARIA)) & " Strinscricao, " & gstrISNULL("Li.Dblareaterreno", "0") & " Dblareaterreno, "
    strSql = strSql & gstrISNULL("Li.DBLVALORVENALTERRENO", "0") & " DBLVALORVENALTERRENO, " & gstrISNULL("Pi.Dblmedidadaarea", "0") & " As DblAreaPredio, "
    strSql = strSql & gstrISNULL("PI.Dblvalorvenalpredio", "0") & " As Dblvalorvenalpredio, Sum(" & gstrISNULL("LV.Dblvalor", "0") & ") Dblvalor "
    strSql = strSql & "From " & gstrLancamentoAlfa & " La, " & gstrLancamentoIPTU & " LI, "
    strSql = strSql & gstrLancamentoValor & " LV, "
    strSql = strSql & "(Select Pi.Intlancamentoiptu,"
    strSql = strSql & "Sum(" & gstrISNULL("Pi.Dblmedidadaarea", "0") & ") As Dblmedidadaarea, "
    strSql = strSql & "Sum(" & gstrISNULL("PI.Dblvalorvenalpredio", "0") & ") As Dblvalorvenalpredio "
    strSql = strSql & "From " & gstrLancamentoAlfa & " La, " & gstrLancamentoIPTU & " LI, " & gstrLancamentoPredioIPTU & " PI "
    strSql = strSql & "Where LA.Pkid = LI.Intlancamentoalfa And LI.Pkid = PI.Intlancamentoiptu And "
    strSql = strSql & "La.strinscricao Between '" & String(gintLenInscricao - Len(Trim(mskstrInscricao.Text)), "0") & Trim(mskstrInscricao.Text) & "' And '" & String(gintLenInscricao - Len(Trim(mskstrInscricao1.Text)), "0") & Trim(mskstrInscricao1.Text) & "' and "
    strSql = strSql & "LA.Intexercicio = " & txt_intExercicio1 & " And LA.Dtmdtcancelamento is null Group By Pi.Intlancamentoiptu) PI "
    strSql = strSql & "Where LA.Pkid = LI.Intlancamentoalfa         And LA.Pkid " & strOUTJSQLServer & "= LV.Intlancamentoalfa" & strOUTJOracle & " And "
    strSql = strSql & "LI.Pkid " & strOUTJSQLServer & "= PI.Intlancamentoiptu" & strOUTJOracle & "       And La.strinscricao Between '" & String(gintLenInscricao - Len(Trim(mskstrInscricao.Text)), "0") & Trim(mskstrInscricao.Text) & "' And '" & String(gintLenInscricao - Len(Trim(mskstrInscricao1.Text)), "0") & Trim(mskstrInscricao1.Text) & "' and "
    strSql = strSql & "LA.Intexercicio = " & txt_intExercicio1 & "                 And LA.Dtmdtcancelamento is null And "
    strSql = strSql & "LV.Bitparcelavalida = 1 "
    strSql = strSql & "Group By la.strdefinicaoisencao , la.strtipoisencaoimunidade, la.dtmdtcancelamento, " & gstrRIGHT("LA.strInscricao", gintRetornaTamanhoMascara(TYP_IMOBILIARIA)) & ", " & gstrISNULL("Li.Dblareaterreno", "0") & ", " & gstrISNULL("Li.DBLVALORVENALTERRENO", "0") & ", " & gstrISNULL("Pi.Dblmedidadaarea", "0") & ", " & gstrISNULL("PI.Dblvalorvenalpredio", "0") & ") LA1 "
    
    strSql = strSql & "Where LA.Strinscricao = LA1.Strinscricao And "
    strSql = strSql & "la.dtmdtcancelamento is null And "
    strSql = strSql & "((LA.strdefinicaoisencao is null And LA.strtipoisencaoimunidade is null) or ( " & strLen & "(LA.strdefinicaoisencao) = 0 And " & strLen & "(LA.strtipoisencaoimunidade) = 0)) And "
    strSql = strSql & "Case when LA.Dblvalor > 0 then LA.Dblvalor Else 1 End / "
    strSql = strSql & "Case when LA1.Dblvalor > 0 then LA1.Dblvalor Else 1 End " & IIf(chk_Negativo.Value = 0, ">" & gstrConvVrParaSql(txt_intIndice), ">" & gstrConvVrParaSql(txt_intIndice * (-1)))
    
    If chk_Ordenacao Then
        strSql = strSql & " Order By Indice Desc"
    Else
        strSql = strSql & " Order By LA.strInscricao Asc"
    End If
    
    strQueryRelatorio = strSql
End Function

Private Function blnDadosOk() As Boolean
    blnDadosOk = False
    
    If Trim(txt_intExercicio) = "" Or Val(Trim(txt_intExercicio)) < 4 Then
        ExibeMensagem "O campo de 1º exercício  não foi preenchido corretamente."
        txt_intExercicio.SetFocus
        Exit Function
    ElseIf Trim(txt_intExercicio1) = "" Or Val(Trim(txt_intExercicio1)) < 4 Then
        ExibeMensagem "O campo de 2º exercício não foi preenchido corretamente."
        txt_intExercicio1.SetFocus
        Exit Function
    End If
    
    'If chk_Negativo.Value = 0 Then
        If Trim(txt_intIndice) = "" Then
            ExibeMensagem "O campo de índice não foi preenchido corretamente."
            txt_intIndice.SetFocus
            Exit Function
        ElseIf CDbl(txt_intIndice) = 0 Then
            ExibeMensagem "O valor do campo de índice não pode ser 0."
            txt_intIndice.SetFocus
            Exit Function
        End If
    'End If
            
    If Trim(mskstrInscricao.Text) = "" Then
        ExibeMensagem "O campo de inscrição inicial não foi preenchido corretamente."
        mskstrInscricao.SetFocus
        Exit Function
    ElseIf Trim(mskstrInscricao1.Text) = "" Then
        ExibeMensagem "O campo de inscrição final não foi preenchido corretamente."
        mskstrInscricao.SetFocus
        Exit Function
    ElseIf Trim(txt_intExercicio.Text) < Trim(txt_intExercicio1.Text) Then
        ExibeMensagem "O 1º exercício não pode ser menor que o 2º exercício final."
        mskstrInscricao.SetFocus
        Exit Function
    ElseIf Trim(mskstrInscricao.Text) > Trim(mskstrInscricao1.Text) Then
        ExibeMensagem "A inscrição inicial não pode ser maior que a inscrição final."
        mskstrInscricao.SetFocus
        Exit Function
    End If
    blnDadosOk = True
End Function

Private Sub mskstrInscricao_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", mskstrInscricao
End Sub

Private Sub mskstrInscricao_LostFocus()
    MarcaCampo mskstrInscricao
End Sub

Private Sub mskstrInscricao1_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", mskstrInscricao1
End Sub

Private Sub mskstrInscricao1_LostFocus()
    MarcaCampo mskstrInscricao1
End Sub

Private Sub txt_intExercicio_GotFocus()
    MarcaCampo txt_intExercicio
End Sub

Private Sub txt_intExercicio_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txt_intExercicio
End Sub
Private Sub txt_intExercicio1_GotFocus()
    MarcaCampo txt_intExercicio1
End Sub

Private Sub txt_intExercicio1_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txt_intExercicio1
End Sub

Private Sub txt_intIndice_GotFocus()
    MarcaCampo txt_intIndice
End Sub

Private Sub txt_intIndice_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "V", txt_intIndice
End Sub


Sub VerificaMascaraInscricao()
    Dim strSql As String
    Dim adoResultado As ADODB.Recordset
    Dim strMascara   As String
    
    strMascara = ""
    
    strSql = ""
    strSql = strSql & "Select * From " & gstrCampoDeInscricao & " "
    strSql = strSql & "Where intTipoDeInscricao = " & TYP_IMOBILIARIA
    strSql = strSql & "Order By intSequencia"
    
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        With adoResultado
            Do While Not .EOF
                strMascara = strMascara & String(!intTamanho, "#") & gstrVerificaCampoNulo(!strSeparador)
                .MoveNext
            Loop
        End With
    End If
    
    mskstrInscricao.Mask = strMascara
    mskstrInscricao1.Mask = strMascara
End Sub

Private Sub txt_intIndice_LostFocus()
    txt_intIndice = gstrConvVrDoSql(txt_intIndice, 4, , True)
End Sub
