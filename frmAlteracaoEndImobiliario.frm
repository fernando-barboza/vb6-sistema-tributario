VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmAlteracaoEndImobiliario 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Alteração de Endereço de Notificação do Imobiliário"
   ClientHeight    =   3195
   ClientLeft      =   2145
   ClientTop       =   2430
   ClientWidth     =   9270
   Icon            =   "frmAlteracaoEndImobiliario.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   9270
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab tab_3dPasta 
      Height          =   3075
      Left            =   60
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   60
      Width           =   9165
      _ExtentX        =   16166
      _ExtentY        =   5424
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Endereço de Notificação"
      TabPicture(0)   =   "frmAlteracaoEndImobiliario.frx":1042
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblstrInscricaoAnterior"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "mskstrInscricao"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fra_Notificacao"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      Begin VB.Frame fra_Notificacao 
         Height          =   1815
         Left            =   150
         TabIndex        =   17
         Top             =   1050
         Width           =   8865
         Begin VB.TextBox txtintCodigoLogradouro 
            Height          =   315
            Left            =   4095
            MaxLength       =   8
            TabIndex        =   6
            Top             =   210
            Width           =   735
         End
         Begin VB.CommandButton cmd_TipoLogradouro 
            Height          =   300
            Left            =   1815
            Picture         =   "frmAlteracaoEndImobiliario.frx":105E
            Style           =   1  'Graphical
            TabIndex        =   3
            TabStop         =   0   'False
            ToolTipText     =   "Ativa Cadastro de Tipo de Logradouro"
            Top             =   210
            Width           =   330
         End
         Begin VB.CommandButton cmd_TituloLogradouro 
            Height          =   300
            Left            =   3705
            Picture         =   "frmAlteracaoEndImobiliario.frx":1344
            Style           =   1  'Graphical
            TabIndex        =   5
            TabStop         =   0   'False
            ToolTipText     =   "Ativa Cadastro de Título de Logradouro"
            Top             =   210
            Width           =   330
         End
         Begin VB.TextBox txtstrDistritoC 
            Height          =   285
            Left            =   1080
            MaxLength       =   50
            TabIndex        =   15
            Top             =   1215
            Width           =   3525
         End
         Begin VB.CommandButton cmd_MunicipioC 
            Height          =   300
            Left            =   4665
            Picture         =   "frmAlteracaoEndImobiliario.frx":162A
            Style           =   1  'Graphical
            TabIndex        =   12
            TabStop         =   0   'False
            ToolTipText     =   "Ativa Cadastro de Municipio"
            Top             =   870
            Width           =   330
         End
         Begin VB.TextBox txtstrBairroC 
            Height          =   285
            Left            =   5625
            MaxLength       =   50
            TabIndex        =   10
            Top             =   540
            Width           =   3075
         End
         Begin VB.TextBox txtintNumeroC 
            Height          =   285
            Left            =   1080
            MaxLength       =   8
            TabIndex        =   8
            Top             =   555
            Width           =   855
         End
         Begin VB.TextBox txtstrComplementoC 
            Height          =   285
            Left            =   2835
            MaxLength       =   20
            TabIndex        =   9
            Top             =   555
            Width           =   1260
         End
         Begin VB.TextBox txtintCepC 
            Height          =   285
            Left            =   7620
            MaxLength       =   9
            TabIndex        =   14
            Top             =   870
            Width           =   1080
         End
         Begin MSDataListLib.DataCombo dbcintMunicipioC 
            Height          =   315
            Left            =   1080
            TabIndex        =   11
            Top             =   870
            Width           =   3525
            _ExtentX        =   6218
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dbcintUFC 
            Height          =   315
            Left            =   5640
            TabIndex        =   13
            Top             =   870
            Width           =   660
            _ExtentX        =   1164
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dbcintTipoLogradouro 
            Height          =   315
            Left            =   1080
            TabIndex        =   2
            Top             =   210
            Width           =   750
            _ExtentX        =   1323
            _ExtentY        =   556
            _Version        =   393216
            MatchEntry      =   -1  'True
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dbcintTituloLogradouro 
            Height          =   315
            Left            =   2265
            TabIndex        =   4
            Top             =   210
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   556
            _Version        =   393216
            MatchEntry      =   -1  'True
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dbcstrLogradouroC 
            Height          =   315
            Left            =   4875
            TabIndex        =   7
            Top             =   210
            Width           =   3840
            _ExtentX        =   6773
            _ExtentY        =   556
            _Version        =   393216
            MatchEntry      =   -1  'True
            Text            =   ""
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Distrito"
            Height          =   195
            Left            =   540
            TabIndex        =   25
            Top             =   1305
            Width           =   480
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Município"
            Height          =   195
            Left            =   315
            TabIndex        =   24
            Top             =   990
            Width           =   705
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Bairro"
            Height          =   195
            Left            =   5070
            TabIndex        =   23
            Top             =   630
            Width           =   405
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Nº"
            Height          =   195
            Left            =   840
            TabIndex        =   22
            Top             =   630
            Width           =   180
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Compl."
            Height          =   195
            Left            =   2235
            TabIndex        =   21
            Top             =   630
            Width           =   480
         End
         Begin VB.Label Label17 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "UF"
            Height          =   195
            Left            =   5295
            TabIndex        =   20
            Top             =   990
            Width           =   210
         End
         Begin VB.Label Label18 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Cep"
            Height          =   195
            Left            =   7275
            TabIndex        =   19
            Top             =   960
            Width           =   285
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Logradouro"
            Height          =   195
            Left            =   210
            TabIndex        =   18
            Top             =   285
            Width           =   810
         End
      End
      Begin MSMask.MaskEdBox mskstrInscricao 
         Height          =   300
         Left            =   1575
         TabIndex        =   1
         Top             =   600
         Width           =   1830
         _ExtentX        =   3228
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   24
         PromptChar      =   " "
      End
      Begin VB.Label lblstrInscricaoAnterior 
         AutoSize        =   -1  'True
         Caption         =   "Inscrição Cadastral"
         Height          =   195
         Left            =   150
         TabIndex        =   16
         Top             =   705
         Width           =   1350
      End
   End
End
Attribute VB_Name = "frmAlteracaoEndImobiliario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim mblnSelecionou  As Boolean
    Dim mobjAux         As Object
    Dim lngPkid         As Long

Private Function strQuery() As String
    
    Dim strSql  As String
    
    strSql = ""
    strSql = strSql & "Select "
    strSql = strSql & "Pkid, "
    strSql = strSql & gstrISNULL("intTipoLogradouro", "0") & " AS intTipoLogradouro, "
    strSql = strSql & gstrISNULL("intTituloLogradouro", "0") & " AS intTituloLogradouro, "
    strSql = strSql & gstrISNULL("intCodigoLogradouro", "0") & " AS intCodigoLogradouro, "
    strSql = strSql & "strLogradouroC, "
    strSql = strSql & "intNumeroC, "
    strSql = strSql & "strComplementoC, "
    strSql = strSql & "strBairroC, "
    strSql = strSql & gstrISNULL("intMunicipioC", "0") & " AS intMunicipioC, "
    strSql = strSql & gstrISNULL("intUFC", "0") & " AS intUFC, "
    strSql = strSql & gstrISNULL("intCepC", "0") & " AS intCepC, "
    strSql = strSql & "strDistritoC "
    strSql = strSql & "From "
    strSql = strSql & gstrImobiliario & " IM "
    strSql = strSql & "Where "
    strSql = strSql & "IM.Strinscricao = '" & String(gintLenInscricao - Len(Trim(mskstrInscricao.Text)), "0") & UCase(mskstrInscricao.Text) & "'"
    
    strQuery = strSql
    
End Function

Private Sub cmd_MunicipioC_Click()
    ChamaFormCadastro frmCadCidade, dbcintMunicipioC
End Sub

Private Sub cmd_TipoLogradouro_Click()
    CarregaForm frmCadTipoLogradouro, dbcintTipoLogradouro
End Sub

Private Sub cmd_TituloLogradouro_Click()
    CarregaForm frmCadTituloLogradouro, dbcintTituloLogradouro
End Sub

Private Sub dbcintTituloLogradouro_GotFocus()
    MarcaCampo dbcintTipoLogradouro
End Sub

Private Sub dbcstrLogradouroC_GotFocus()
    MarcaCampo dbcstrLogradouroC
End Sub

Private Sub dbcstrLogradouroC_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbcstrLogradouroC
End Sub

Private Sub Form_Activate()
    gintCodSeguranca = 1261
    
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
    dbcintTipoLogradouro.Tag = gstrQueryDataComboTipoLogradouro & ";strSigla"
    dbcintMunicipioC.Tag = gstrQueryDataComboMunicipio & ";strDescricao"
    dbcintUFC.Tag = gstrQueryDataComboUF & ";strSigla"
    dbcintTituloLogradouro.Tag = gstrQueryDataComboTituloLogradouro & ";strDescricao"
    VerificaMascaraInscricao
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar
    mblnSelecionou = False
End Sub

Public Sub MantemForm(ByVal strModoOperacao As String)

    Screen.MousePointer = vbArrow
    
    If UCase(strModoOperacao) = gstrSalvar Then
        If Not blnDadosOk Then Exit Sub
        If gblnExclusaoGravacaoOk("A") Then
            AlteraNotificacao (lngPkid)
        End If
    ElseIf UCase(strModoOperacao) = gstrPreencherLista Then
        dbcstrLogradouroC.Tag = gstrQueryLogradouro & ";L.strDescricao"
        PreencherListaDeOpcoes Me.ActiveControl
        dbcstrLogradouroC.Tag = ""
    ElseIf UCase(strModoOperacao) = gstrNovo Then
        Limpa_Controles Me, True, True, True, True, True
        mskstrInscricao.Text = ""
        mskstrInscricao.SetFocus
        lngPkid = 0
    End If
    
    Screen.MousePointer = vbDefault
    
End Sub

Private Function blnDadosOk()
    
    blnDadosOk = False
    If Not blnPreencheEndereco(False) Then
        Exit Function
    ElseIf Trim(dbcstrLogradouroC.Text) = "" Then
        ExibeMensagem "O campo logradouro deve ser preenchido corretamente."
        dbcstrLogradouroC.SetFocus
        Exit Function
    ElseIf Trim(txtintNumeroC) = "" Then
        ExibeMensagem "O campo número deve ser preenchido corretamente."
        txtintNumeroC.SetFocus
        Exit Function
    ElseIf Trim(txtstrBairroC) = "" Then
        ExibeMensagem "O campo bairro deve ser preenchido corretamente."
        txtstrBairroC.SetFocus
        Exit Function
    ElseIf Not dbcintMunicipioC.MatchedWithList Then
        ExibeMensagem "O campo município deve ser preenchido corretamente."
        dbcintMunicipioC.SetFocus
        Exit Function
    ElseIf Not dbcintUFC.MatchedWithList Then
        ExibeMensagem "O campo UF deve ser preenchido corretamente."
        dbcintUFC.SetFocus
        Exit Function
    ElseIf Trim(txtintCepC) = "" Then
        ExibeMensagem "O campo CEP deve ser preenchido corretamente."
        txtintCepC.SetFocus
        Exit Function
    End If
    blnDadosOk = True
    
End Function

Private Sub mskstrInscricao_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", mskstrInscricao
End Sub

Private Sub mskstrInscricao_LostFocus()
    If mskstrInscricao.Text <> "" Then
        blnPreencheEndereco True
    End If
End Sub

Private Sub txtintCepC_LostFocus()
    txtintCepC = gstrCEPFormatado(txtintCepC)
    dbcintTipoLogradouro.Text = ""
    dbcintTituloLogradouro.Text = ""
    txtintCodigoLogradouro.Text = ""
    txtstrDistritoC.Text = ""
    txtintNumeroC.Text = ""
    txtstrComplementoC.Text = ""
    dbcstrLogradouroC.Tag = gstrQueryLogradouro & ";L.strDescricao"
    CepLogradouro txtintCepC, dbcstrLogradouroC, txtstrBairroC, dbcintMunicipioC, dbcintUFC, dbcintTipoLogradouro, dbcintTituloLogradouro, , True, False, True, True, True, True, False
    dbcstrLogradouroC.Tag = ""
End Sub

Private Sub txtintCepC_GotFocus()
    MarcaCampo txtintCepC
End Sub

Private Sub txtintCepC_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "E", txtintCepC
End Sub

Private Sub dbcintTipoLogradouro_Click(Area As Integer)
    If Area = 0 Then DropDownDataCombo dbcintTipoLogradouro, Me, Area
End Sub

Private Sub dbcintTipoLogradouro_GotFocus()
    MarcaCampo dbcintTipoLogradouro
End Sub

Private Sub dbcintTipoLogradouro_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbcintTipoLogradouro, Me, , KeyCode, Shift
End Sub

Private Sub dbcintTituloLogradouro_Click(Area As Integer)
    If Area = 0 Then DropDownDataCombo dbcintTituloLogradouro, Me, Area
End Sub

Private Sub dbcintTituloLogradouro_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbcintTituloLogradouro, Me, , KeyCode, Shift
End Sub

Private Sub dbcintTituloLogradouro_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbcintTituloLogradouro
End Sub

Private Sub txtintCodigoLogradouro_GotFocus()
    MarcaCampo txtintCodigoLogradouro
End Sub

Private Sub txtintCodigoLogradouro_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub txtintNumeroC_GotFocus()
    MarcaCampo txtintNumeroC
End Sub

Private Sub txtintNumeroC_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtintNumeroC
End Sub

Private Sub txtstrBairroC_GotFocus()
    MarcaCampo txtstrBairroC
End Sub

Private Sub txtstrComplementoC_GotFocus()
    MarcaCampo txtstrComplementoC
End Sub

Private Sub txtstrComplementoC_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrComplementoC
End Sub

Private Sub txtstrBairroC_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrBairroC
End Sub

Private Sub dbcintMunicipioC_Click(Area As Integer)
    If Area = 0 Then DropDownDataCombo dbcintMunicipioC, Me, Area
End Sub

Private Sub dbcintMunicipioC_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbcintMunicipioC, Me, , KeyCode, Shift
End Sub

Private Sub dbcintMunicipioC_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub dbcintUFC_Click(Area As Integer)
    If Area = 0 Then DropDownDataCombo dbcintUFC, Me, Area
End Sub

Private Sub dbcintUFC_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbcintUFC, Me, , KeyCode, Shift
End Sub

Private Sub dbcintUFC_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
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
End Sub

Private Function blnPreencheEndereco(blnLostFocus As Boolean) As Boolean
    Dim adoResultado As ADODB.Recordset
    
    Set gobjBanco = New clsBanco
    blnPreencheEndereco = False
    If gobjBanco.CriaADO(strQuery, 5, adoResultado) Then
        With adoResultado
            If .RecordCount = 1 Then
                If blnLostFocus Then
                    lngPkid = Val(gstrENulo(!Pkid))
                    PreencherListaDeOpcoes dbcintTipoLogradouro, (!intTipoLogradouro)
                    PreencherListaDeOpcoes dbcintTituloLogradouro, (!intTituloLogradouro)
                    txtintCodigoLogradouro = gstrENulo(!intCodigoLogradouro)
                    dbcstrLogradouroC.Text = gstrENulo(!strlogradouroc)
                    txtintNumeroC = gstrENulo(!intNumeroC)
                    txtstrComplementoC = gstrENulo(!strComplementoC)
                    txtstrBairroC = gstrENulo(!strBairroC)
                    PreencherListaDeOpcoes dbcintMunicipioC, (!intMunicipioC)
                    PreencherListaDeOpcoes dbcintUFC, (!intUFC)
                    txtintCepC = gstrCEPFormatado(gstrENulo(!intcepc))
                    txtstrDistritoC = gstrENulo(!strDistritoC)
                Else
                    blnPreencheEndereco = True
                End If
            ElseIf .RecordCount > 1 Then
                ExibeMensagem "Existe mais de um registro com essa inscrição."
            Else
                ExibeMensagem "Inscrição inválida."
            End If
        End With
    End If
    
End Function

Private Function AlteraNotificacao(lngPkid As Long) As Boolean
    Dim strSql As String
    
    strSql = ""
    strSql = strSql & "Update " & gstrImobiliario & " Set "
    strSql = strSql & "intTipoLogradouro = " & gstrENulo(dbcintTipoLogradouro.BoundText, , True) & ", "
    strSql = strSql & "intTituloLogradouro = " & gstrENulo(dbcintTituloLogradouro.BoundText, , True) & ", "
    strSql = strSql & "intCodigoLogradouro = " & gstrENulo(txtintCodigoLogradouro, , True) & ", "
    strSql = strSql & "strlogradouroc = '" & gstrENulo(dbcstrLogradouroC) & "', "
    strSql = strSql & "intNumeroC = " & gstrENulo(txtintNumeroC, , True) & ", "
    strSql = strSql & "strComplementoC = '" & gstrENulo(txtstrComplementoC) & "', "
    strSql = strSql & "strBairroC = '" & gstrENulo(txtstrBairroC) & "', "
    strSql = strSql & "intMunicipioC = " & gstrENulo(dbcintMunicipioC.BoundText, , True) & ", "
    strSql = strSql & "intUFC = " & gstrENulo(dbcintUFC.BoundText, , True) & ", "
    strSql = strSql & "intcepc = " & gstrENulo(Replace(txtintCepC, "-", ""), , True) & ", "
    strSql = strSql & "strDistritoC = '" & gstrENulo(txtstrDistritoC) & "' "
    strSql = strSql & "Where "
    strSql = strSql & "Pkid = " & lngPkid
    
    Set gobjBanco = New clsBanco
    
    gobjBanco.ExecutaBeginTrans
    
    If gobjBanco.Execute(strSql) Then
        gobjBanco.ExecutaCommitTrans
        MantemForm gstrNovo
    Else
        ExibeMensagem "Não foi possível concluir alteração do endereço de notificação."
        gobjBanco.ExecutaRollbackTrans
        mskstrInscricao.SetFocus
    End If
    
End Function
