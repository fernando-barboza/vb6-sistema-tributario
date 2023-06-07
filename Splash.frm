VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmSplash 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4470
   ClientLeft      =   3030
   ClientTop       =   2235
   ClientWidth     =   5820
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "Splash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4470
   ScaleWidth      =   5820
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtUserName 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1200
      MaxLength       =   10
      TabIndex        =   0
      Top             =   3120
      Width           =   1935
   End
   Begin MSComctlLib.ImageList imgBotao 
      Left            =   120
      Top             =   540
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   74
      ImageHeight     =   19
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Splash.frx":1042
            Key             =   "OK1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Splash.frx":17DA
            Key             =   "AV1"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Splash.frx":1FDA
            Key             =   "CA2"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Splash.frx":27E2
            Key             =   "CA1"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Splash.frx":2FE2
            Key             =   "FE2"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Splash.frx":37B6
            Key             =   "FE1"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Splash.frx":3F8A
            Key             =   "OK2"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Splash.frx":4736
            Key             =   "AV2"
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtServidor 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1200
      TabIndex        =   8
      Top             =   3120
      Width           =   1935
   End
   Begin MSComctlLib.ImageList imgListLogo 
      Left            =   690
      Top             =   540
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   388
      ImageHeight     =   299
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   16
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Splash.frx":4F3E
            Key             =   "Tributario"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Splash.frx":8176
            Key             =   "Gerencial"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Splash.frx":CC52
            Key             =   "Seguranca"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Splash.frx":117FE
            Key             =   "Concurso"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Splash.frx":14B9A
            Key             =   "Escolar"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Splash.frx":1850A
            Key             =   "Frota"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Splash.frx":1B536
            Key             =   "Legislacao"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Splash.frx":1E7CE
            Key             =   "Compras"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Splash.frx":2228A
            Key             =   "Material"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Splash.frx":25F86
            Key             =   "Menor"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Splash.frx":2973E
            Key             =   "Orcamentario"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Splash.frx":2CA52
            Key             =   "Ouvidoria"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Splash.frx":30362
            Key             =   "Patrimonio"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Splash.frx":333DA
            Key             =   "Protocolo"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Splash.frx":3667E
            Key             =   "RH"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Splash.frx":3A26E
            Key             =   "SAC"
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtSenha 
      Appearance      =   0  'Flat
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1200
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   3450
      Width           =   1935
   End
   Begin VB.TextBox txtDataBase 
      Appearance      =   0  'Flat
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1200
      TabIndex        =   10
      Top             =   3450
      Width           =   1935
   End
   Begin VB.Label lblVersao 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   5715
      TabIndex        =   12
      Top             =   45
      Width           =   45
   End
   Begin VB.Image pctAtualizacao 
      Height          =   480
      Left            =   5250
      Picture         =   "Splash.frx":3D4FA
      ToolTipText     =   "Não foi possível atualizar o sistema! Localização do servidor de atualização não encontrada!"
      Top             =   195
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label lblProgresso 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   120
      TabIndex        =   11
      Top             =   3840
      Width           =   45
   End
   Begin VB.Image img_Advanced 
      Height          =   285
      Left            =   4560
      MouseIcon       =   "Splash.frx":3D93C
      MousePointer    =   99  'Custom
      Stretch         =   -1  'True
      Top             =   4050
      Width           =   1110
   End
   Begin VB.Image img_Cancel 
      Height          =   285
      Left            =   4560
      MouseIcon       =   "Splash.frx":3DC46
      MousePointer    =   99  'Custom
      Stretch         =   -1  'True
      Top             =   3675
      Width           =   1110
   End
   Begin VB.Image img_OK 
      Height          =   285
      Left            =   4560
      MouseIcon       =   "Splash.frx":3DF50
      MousePointer    =   99  'Custom
      Stretch         =   -1  'True
      Top             =   3330
      Width           =   1110
   End
   Begin VB.Label lblSenha 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Senha"
      Height          =   195
      Left            =   630
      TabIndex        =   7
      Top             =   3525
      Width           =   465
   End
   Begin VB.Label lblLogin 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Login"
      Height          =   195
      Left            =   705
      TabIndex        =   6
      Top             =   3165
      Width           =   390
   End
   Begin VB.Label lblServidor 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Servidor"
      Height          =   195
      Left            =   510
      TabIndex        =   2
      Top             =   3165
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Label lblDataBase 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Database"
      Height          =   195
      Left            =   405
      TabIndex        =   9
      Top             =   3525
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Label lblProtegido 
      BackStyle       =   0  'Transparent
      Caption         =   "Este programa é protegido pela lei de copyright e tratados internacionais, conforme descrito no comando 'Sobre...' da Ajuda."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   120
      TabIndex        =   5
      Top             =   4080
      Width           =   4275
   End
   Begin VB.Label lblLicenciado 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Licenciado"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   210
      Left            =   120
      TabIndex        =   4
      Top             =   255
      Width           =   780
   End
   Begin VB.Label lblLicenciadoPara 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Este produto está licenciado para"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   210
      Left            =   120
      TabIndex        =   1
      Top             =   30
      Width           =   2415
   End
   Begin VB.Image imgLogo 
      Height          =   4485
      Left            =   0
      Top             =   0
      Width           =   5820
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Const flags = SWP_NOMOVE Or SWP_NOSIZE

Dim strCaminhoArqTxt As String

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            img_OK_Click
        Case vbKeyEscape
            img_Cancel_Click
    End Select
End Sub

Private Sub Form_Load()
    
'******************************************************************************************
' Data: 06/03/2003
' Alteração: - Alterada chamada da função gblnParametrosConeccaoOk.
' Responsável: Everton Bianchini
'******************************************************************************************
    
    Dim lnghWind        As Long
    If App.PrevInstance Then
        MsgBox "Já existe uma instância do sistema aberta.", vbSystemModal
        End
    End If
    
    lblVersao.Caption = "Versão: " & App.Major & "." & App.Minor & "." & App.Revision
    lblVersao.ToolTipText = "Versão: " & App.Major & "." & App.Minor & "." & App.Revision
    
    If Not VerificaVersao Then pctAtualizacao.Visible = True
    
    Screen.MousePointer = vbHourglass
    gblnDemonstracao = False
    imgLogo.Picture = imgListLogo.ListImages(App.ProductName).Picture
    img_OK.Picture = imgBotao.ListImages("OK2").Picture
    img_Cancel.Picture = imgBotao.ListImages("CA2").Picture
    img_Advanced.Picture = imgBotao.ListImages("AV2").Picture
    img_Advanced.Tag = "AVANÇADO"
    Me.Refresh
    
    If gstrLeValorRegister(HKEY_CURRENT_USER, "SOFTWARE\CPD\AdGover\Parâmetros", "Empresa") = "" Then
        gGravaValorRegister HKEY_CURRENT_USER, "SOFTWARE\CPD\AdGover\Parâmetros", "Empresa", gstrLeValorRegister(HKEY_LOCAL_MACHINE, "SOFTWARE\CPD\AdGover\Parâmetros", "Empresa")
    End If
    
    If gstrLeValorRegister(HKEY_CURRENT_USER, "SOFTWARE\CPD\AdGover\Parâmetros", "UserName") = "" Then
        gGravaValorRegister HKEY_CURRENT_USER, "SOFTWARE\CPD\AdGover\Parâmetros", "UserName", gstrLeValorRegister(HKEY_LOCAL_MACHINE, "SOFTWARE\CPD\AdGover\Parâmetros", "UserName")
    End If
    
    lblLicenciado = gstrLeValorRegister(HKEY_CURRENT_USER, _
                                        "SOFTWARE\CPD\AdGover\Parâmetros", "Empresa")
                                        
    AlwaysOnTop Me, False
    
    Me.Show
    DoEvents
'    If Not gblnParametrosConeccaoOk(gstrServidor, gstrDatabase) Then
    If Not gblnParametrosConeccaoOk(gstrServidor, gstrDatabase, bytDBType) Then
        img_Advanced_Click
        txtServidor.SetFocus
        Screen.MousePointer = 0
        Exit Sub
    End If
    
    'Vamos obter o caminho do arquivo UserNamePass.txt
    If gstrLeValorRegister(HKEY_CURRENT_USER, "SOFTWARE\CPD\AdGover\Parâmetros", "ArquivoDataBase") = "" Then
        gGravaValorRegister HKEY_CURRENT_USER, "SOFTWARE\CPD\AdGover\Parâmetros", "ArquivoDataBase", "\\Sainf002\Update"
    End If

    strCaminhoArqTxt = gstrLeValorRegister(HKEY_CURRENT_USER, _
                          "SOFTWARE\CPD\AdGover\Parâmetros", "ArquivoDataBase")
    
    GetLoginPass
    
    txtServidor = gstrServidor
    txtDatabase = gstrDatabase
    
    If gblnTrocaUsuario = False Then
        If gblnLinhaCommando(txtUserName, txtSenha) = False Then
            txtUserName = gstrLeValorRegister(HKEY_CURRENT_USER, _
                          "SOFTWARE\CPD\AdGover\Parâmetros", "UserName")
        End If
        txtSenha.SetFocus
    Else
        txtUserName.SetFocus
    End If
    
    Screen.MousePointer = vbDefault
    

    
End Sub

Private Sub AtualizaSenha()
    Dim strSenhaAux     As String
    Dim strSQL          As String
    Dim adoResultado    As ADODB.Recordset
    strSQL = ""
    strSQL = strSQL & "SELECT PKId, strSenha  FROM " & gstrUsuarios
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
        With adoResultado
            Do While .EOF = False
'                strSenhaAux = gstrSenhaCripitografada(!strSenha)
                strSQL = ""
                strSQL = strSQL & "UPDATE " & gstrUsuarios & " SET "
                strSQL = strSQL & "strSenha = "
                strSQL = strSQL & "'" & gstrStringCripitografada(strSenhaAux, True) & "' "
                strSQL = strSQL & "WHERE PKId = " & !Pkid
                Set gobjBanco = New clsBanco
                gobjBanco.Execute strSQL
                .MoveNext
            Loop
        End With
    End If
End Sub

Private Sub img_Advanced_Click()
    If UCase(img_Advanced.Tag) = "AVANÇADO" Then
        img_Advanced.Tag = "Fechar"
        img_Advanced.Picture = imgBotao.ListImages("FE2").Picture
        txtUserName.Top = 2400
        txtSenha.Top = 2745
        lblLogin.Top = 2445
        lblSenha.Top = 2800
        lblServidor.Visible = True
        txtServidor.Visible = True
        lblDataBase.Visible = True
        txtDatabase.Visible = True
        txtServidor.SetFocus
    Else
        img_Advanced.Picture = imgBotao.ListImages("AV2").Picture
        img_Advanced.Tag = "AVANÇADO"
        txtUserName.Top = 3120
        txtSenha.Top = 3465
        lblLogin.Top = 3165
        lblSenha.Top = 3525
        lblServidor.Visible = False
        txtServidor.Visible = False
        lblDataBase.Visible = False
        txtDatabase.Visible = False
        txtUserName.SetFocus
    End If
End Sub

Private Sub img_Cancel_Click()
    End
End Sub

Private Sub img_OK_Click()
    
'******************************************************************************************
' Data: 05/03/2003
' Alteração: - Verificação do preenchimento do txtDataBase alterada de forma que esta só
'            ocorra qdo o Banco de Dados não for Oracle.
'            - Substituição dos comandos nativos do SQL Server pelas variáveis de comandos
'            nativos.
' Responsável: Everton Bianchini
'------------------------------------------------------------------------------------------
' Data: 26/03/2003
' Alteração: - Substituição da variável strISNULL pela função gstrISNULL. Ver definição da
'            função para maiores esclarecimentos.
' Responsável: Everton Bianchini
'******************************************************************************************
    
    Dim blnErroServidor As Boolean
    Dim blnErroDataBase As Boolean
    Dim blnExpirou As Boolean
    Dim strSQL As String
    Dim adoRec As ADODB.Recordset
    Dim adoResultado As ADODB.Recordset
    Dim intNumDias As Integer
    
    Screen.MousePointer = vbHourglass
    
    If Trim(txtUserName) = "" Then
        MsgBox "Usuário tem que ser informado.", vbSystemModal
        txtUserName.SetFocus
        GoTo FimcmdOK_Click
    ElseIf Trim(txtSenha) = "" Then
        MsgBox "Senha tem que ser informada.", vbSystemModal
        txtSenha.SetFocus
        GoTo FimcmdOK_Click
    ElseIf (Trim(txtDatabase) = "") And (bytDBType <> Oracle) Then
        MsgBox "Nome da base de dados tem que ser informado.", vbSystemModal
        verificaFocoAvancado txtDatabase
        GoTo FimcmdOK_Click
    ElseIf (Trim(txtDatabase) = "") And (bytDBType <> Oracle) Then
        MsgBox "Nome da base de dados tem que ser informado.", vbSystemModal
        verificaFocoAvancado txtDatabase
        GoTo FimcmdOK_Click
    End If
    
    gstrServidor = Trim(txtServidor)
    gstrDatabase = Trim(txtDatabase)
    gstrLoginUser = Trim(txtUserName)
    gstrPwdUser = Trim(txtSenha)
    
    lblProgresso.Caption = "Verificando conexão com o banco ..."
    If gblnBancoDadosOK(gstrServidor, gstrDatabase, gstrUsername, gstrPassword, blnErroServidor, blnErroDataBase) Then
        strSQL = ""
'        strSql = strSql & " SELECT dtmExpiracao, GETDATE() AS DataDeHoje, "
        'strSql = strSql & " SELECT dtmExpiracao, " & strGETDATE & " AS DataDeHoje, "
        'strSql = strSql & " dtmInicio "
        'strSql = strSql & " FROM " & gstrEmpresa
        
        'Set gobjBanco = New clsBanco
        
        'If gobjBanco.CriaADO(strSql, 10, adoRec) Then
        '    blnExpirou = False
        '    With adoRec
        '        If Not (.BOF And .EOF) Then
        '            intNumDias = DateDiff("d", !DataDeHoje, CDate(gstrStringCripitografada(!dtmExpiracao)))
        '
        '            If intNumDias > 0 Then
        '
        '                If !DataDeHoje < CDate(gstrStringCripitografada(!dtmInicio)) Then
        '                    strSql = "" 'Atualiza flag auxiliar p/ travamento de acesso aos sistemas
        '                    strSql = strSql & " UPDATE " & gstrItens
        '                    strSql = strSql & " SET blnChild = 1"
        '                    gobjBanco.Execute strSql
        '                    MsgBox "Data do sistema operacional inválida!" & Chr(10) & "Por favor, entre em contato com o administrador do sistema", vbSystemModal
        '                    End
        '                Else
        '                    strSql = ""
        '                    strSql = strSql & " UPDATE " & gstrEmpresa
        '                    strSql = strSql & " SET dtmInicio = '" & gstrStringCripitografada(Date, True) & "'"
        '                    gobjBanco.Execute strSql
        '                End If
        '            Else
        '                strSql = "" 'Atualiza flag auxiliar p/ travamento de acesso aos sistemas
        '                strSql = strSql & " UPDATE " & gstrItens
        '                strSql = strSql & " SET blnChild = 1"
        '                gobjBanco.Execute strSql
        '                MsgBox "Seu período de uso do sistema expirou." & Chr(10) & "Por favor, entre em contato com o administrador do sistema", vbSystemModal
        '                End
        '            End If
        '        End If
        '    End With
        'End If
        
        'strSql = ""
        'By power strSql = strSql & " SELECT TOP 1 (ISNULL(blnChild, 1)) As blnExpirado "
'        strSql = strSql & " SELECT (ISNULL(blnChild, 1)) As blnExpirado "
'        strSql = strSql & " SELECT (" & strISNULL & "(blnChild, 1)) As blnExpirado "
        'strSql = strSql & " SELECT (" & gstrISNULL("blnChild", "1") & ") As blnExpirado "
        'strSql = strSql & " FROM " & gstrItens
        'strSql = strSql & " ORDER BY blnChild DESC"
        'Set gobjBanco = New clsBanco
        'If gobjBanco.CriaADO(strSql, 10, adoRec) Then
        '    With adoRec
        '        If Not (.BOF And .EOF) Then
        '            If Abs(!blnExpirado) = 1 Then
        '                MsgBox "Seu período de uso do sistema expirou." & Chr(10) & "Por favor, entre em contato com o administrador do sistema", vbSystemModal
        '                End
        '            End If
        '        End If
        '    End With
        'Else
        '    End
        'End If
            Set gobjBanco = New clsBanco
    
            strSQL = "SELECT strDocumentos FROM " & gstrEmpresa
    
            If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
                If Not adoResultado.EOF Then
                    If Right(Trim(adoResultado(0)), 1) = "\" Then
                        gstrDirDocumentos = adoResultado(0)
                    Else
                        gstrDirDocumentos = adoResultado(0) & "\"
                    End If
                End If
            End If
    
            Set gobjBanco = Nothing
    ElseIf blnErroServidor Then
         verificaFocoAvancado txtServidor
         GoTo FimcmdOK_Click
    ElseIf blnErroDataBase Then
         verificaFocoAvancado txtDatabase
         GoTo FimcmdOK_Click
    Else
         txtUserName.SetFocus
         GoTo FimcmdOK_Click
    End If
    lblProgresso.Caption = "Verificando conexão avançada ..."

    lblProgresso.Caption = "Estabelecendo conexão ..."
    gblnTrocaNomeBancoDeDados
    
    lblProgresso.Caption = "Verificando permissões do usuário ..."
    If blnUsuarioSenhaOk Then
        CarregaPermissoes
        LeNomeEmpresa
        LeModuloAtual
        gGravaValorRegister HKEY_CURRENT_USER, "SOFTWARE\CPD\AdGover\Parâmetros", "Empresa", gstrNomeEmpresa
        gGravaValorRegister HKEY_CURRENT_USER, "SOFTWARE\CPD\AdGover\Parâmetros", "UserName", Trim(txtUserName)
        gGravaValorRegister HKEY_CURRENT_USER, "SOFTWARE\CPD\AdGover\Parâmetros", "Servidor", gstrServidor
        gGravaValorRegister HKEY_CURRENT_USER, "SOFTWARE\CPD\AdGover\Parâmetros", "DataBase", gstrDatabase
        
        strSQL = "SELECT strAtualizacao, bytatualizaterminais from " & gstrEmpresa
        
        Set gobjBanco = New clsBanco
        
        If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
            If Not adoResultado.EOF Then
                If Val(gstrENulo(adoResultado!bytatualizaterminais)) <> 0 Then
                    gGravaValorRegister HKEY_CURRENT_USER, "SOFTWARE\CPD\AdGover\Parâmetros", "Atualizador", gstrENulo(adoResultado(0))
                End If
            End If
        End If
        
        Set gobjBanco = Nothing
        
        'LeExercicio
        LeMascacaraEspecifica
        Unload Me
        Load MDIMenu
        
        MDIMenu.Caption = MDIMenu.Caption & " - Servidor: " & gstrServidor & " - Base de dados: " & gstrDatabase
        MDIMenu.Show
        If gblnTrocaUsuario = False Then
            Unload frmSplash
        End If
    End If
FimcmdOK_Click:
    Screen.MousePointer = vbDefault
End Sub

Private Sub verificaFocoAvancado(objFoco As Object)
    If UCase(img_Advanced.Tag) = "AVANÇADO" Then
        img_Advanced_Click
        objFoco.SetFocus
    End If
End Sub

Private Sub VerificaReinicializacaoDeSenha(blnNovaSenha As Boolean)
    Dim strSenha As String
    If blnNovaSenha Then
        AlwaysOnTop Me, True
        frmAlteracaoSenha.Show vbModal
    End If
End Sub

Sub LeNomeEmpresa()
    Dim strSQL      As String
    Dim adoEmpresa  As ADODB.Recordset
    'Busca dados do Cadastro de Empresa
    strSQL = ""
    strSQL = strSQL & "SELECT E.strNome, E.intUF, M.strDescricao, "
    strSQL = strSQL & "E.intCidade, U.strSigla, E.intSerie FROM "
    strSQL = strSQL & gstrEmpresa & " E, " & gstrCidade & " M, " & gstrUF & " U "
    strSQL = strSQL & " WHERE M.PKId = E.intCidade"
    strSQL = strSQL & " AND U.PKId = E.intUF"
    
    Set gobjBanco = New clsBanco
    
    If gobjBanco.CriaADO(strSQL, 5, adoEmpresa) Then
        With adoEmpresa
            If .EOF = False Then
                gstrNomeEmpresa = gstrStringCripitografada(Trim(!STRNOME), , True)
                gstrCidadeEmpresa = Trim(!strDescricao)
                gstrUFEmpresa = Trim(!strsigla)
                gintUFEmpresa = IIf(IsNull(!intUf), 0, !intUf)
                gintMunicipioEmpresa = !intCidade
                If Val(gstrENulo(!intSerie)) <> glngNumeroDeSerie(gstrNomeEmpresa) Then
                    MsgBox "Violação do banco de dados. " & _
                    Chr(13) & "Número de série inválido.", vbSystemModal
                    End
                End If
            End If
        End With
        adoEmpresa.Close
        Set adoEmpresa = Nothing
   
    
    End If
    
    Set gobjBanco = Nothing
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Select Case UnloadMode
        Case 0
            End
    End Select
End Sub

Private Sub img_OK_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Shift = 5 Then
        txtSenha = txtUserName
    End If
    img_OK.Picture = imgBotao.ListImages("OK1").Picture
End Sub

Private Sub img_OK_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    img_OK.Picture = imgBotao.ListImages("OK2").Picture
End Sub

Private Sub img_Cancel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    img_Cancel.Picture = imgBotao.ListImages("CA1").Picture
End Sub

Private Sub img_Cancel_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    img_Cancel.Picture = imgBotao.ListImages("CA2").Picture
End Sub

Private Sub img_Advanced_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Shift = 5 Then
        If Button = 1 Then
            txtDatabase = "ADMPublicaNew"
        ElseIf Button = 2 Then
            txtDatabase = "ADMPiedadeNew"
        End If
    End If
    If UCase(img_Advanced.Tag) = "AVANÇADO" Then
        img_Advanced.Picture = imgBotao.ListImages("AV1").Picture
    Else
        img_Advanced.Picture = imgBotao.ListImages("FE1").Picture
    End If
End Sub

Private Sub img_Advanced_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If UCase(img_Advanced.Tag) = "AVANÇADO" Then
        img_Advanced.Picture = imgBotao.ListImages("AV2").Picture
    Else
        img_Advanced.Picture = imgBotao.ListImages("FE2").Picture
    End If
End Sub

Private Sub txtDataBase_GotFocus()
    MarcaCampo txtDatabase
End Sub

Private Sub txtDataBase_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub txtDataBase_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Shift = 1 Then 'Shift
        txtDatabase = "ADMPublicaNew"
    ElseIf Shift = 2 Then 'Control
        txtDatabase = "ADMEdu"
    ElseIf Shift = 4 Then 'Alt
        txtDatabase = "ADM"
    End If
End Sub

Private Sub txtSenha_GotFocus()
    MarcaCampo txtSenha
End Sub

Private Sub txtSenha_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "S", txtSenha
End Sub

Private Sub txtServidor_GotFocus()
    MarcaCampo txtServidor
End Sub

Private Sub txtServidor_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtServidor
End Sub

Private Sub txtUserName_GotFocus()
    MarcaCampo txtUserName
End Sub

Private Sub txtUserName_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtUserName
End Sub

Private Sub txtUserName_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Shift = 1 Then 'Shift
        txtUserName = "cpdmaster"
    ElseIf Shift = 2 Then 'Control
        txtUserName = "selectron"
    End If
End Sub

Private Function blnUsuarioSenhaOk() As Boolean
    '--------------------------------------------------------
    'FUNÇÃO VERIFICAR SE O USUÁRIO INFORMADO ESTÁ CADASTRADO
    'E SE A SENHA INFORMADA ESTÁ CORRETA
    '--------------------------------------------------------
    Dim strSQL              As String
    Dim adoResultado        As ADODB.Recordset
    Static bytQtdTentativa  As Byte

    On Error GoTo ErroblnUsuarioSenhaOk

    bytQtdTentativa = bytQtdTentativa + 1

    'Seleciona o usuário informado da tabela de usuários
    strSQL = ""
    strSQL = strSQL & "SELECT * FROM " & gstrUsuarios & " "
    strSQL = strSQL & "WHERE RTRIM(UPPER(strLogin)) = '" & UCase(txtUserName) & "'"
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
        With adoResultado
            'Verifica se encontrou o usuário
            If .EOF = True Then
                MsgBox "O usuário " & Trim(txtUserName) & " não está cadastrado.", vbSystemModal
                txtUserName.SetFocus
            'Verifica se a senha está correta
            ElseIf gstrStringCripitografada(!strSenha) = UCase(Trim(txtSenha)) Then
                glngCodUsr = !Pkid
                gstrNomeUsuario = gstrENulo(!STRNOME)
                gblnMostraDicas = !blnMostraDicas
                gblnListagemAutomatica = !blnListaAutomatica
                gblnListViewComGrade = !blnObjComGrade
                gblnConfirmaGravacao = !blnConfirmaGravacao
                gIntExercicioUsuario = !intExercicio
                gintExercicio = !intExercicio
                gblnConfirmaExclusao = !blnConfirmaExclusao
                gblnRelatorioZebrado = !blnRelatorioZebrado
                gvntCorZebrado = !strCorZebrado
                gvntFundoObjInacessivel = gstrENulo(!strFundoObjInacessivel)
                gbytchkFundoObjDiferente = !bytchkFundoObjDiferente
                gintEstiloMarcado = !TpLvwRelatorio
                gblnMaster = !blnMaster
                gblnAdmin = !blnAdministrador
                '----- Parâmetros usados no RH
                gblnImprimeTRCTAoGravar = gstrENulo(!ImprimeTRCTAoGravar)
                
                blnUsuarioSenhaOk = True
                bytQtdTentativa = 0
                VerificaReinicializacaoDeSenha !blnReinicializarSenha
            Else
                MsgBox "A senha de acesso informada não está correta.", vbSystemModal
                txtSenha.SetFocus
                MarcaCampo txtSenha
            End If
        End With
        adoResultado.Close
        Set adoResultado = Nothing
    End If

    If bytQtdTentativa > 2 Then
        End
    End If
    Exit Function

ErroblnUsuarioSenhaOk:
End Function

Private Sub GetLoginPass()

    On Error GoTo err_GetLoginPass
    
    If Trim(Command) = "" Then
    
        Open strCaminhoArqTxt & "\UserNamePass.txt" For Input As #1
        
        If Not EOF(1) Then
            Line Input #1, gstrUsername
            Line Input #1, gstrPassword
    
            gstrUsername = gstrStringCripitografada(gstrUsername, False, True)
            gstrPassword = gstrStringCripitografada(gstrPassword, False, True)
            
        End If
        
        Close #1
    Else
        If bytDBType = EDatabases.Oracle Then
            gstrUsername = "cpdmaster"
            gstrPassword = "cpd"
        Else
            gstrUsername = "sa"
            gstrPassword = ""
        End If
        
    End If
    Exit Sub
    
err_GetLoginPass:
    MsgBox "Ocorreu um erro ao encontrar o nome de usuário do banco de Dados. Por Favor, entre em contato com o administrador do sistema. " & vbCrLf & vbCrLf & "Descrição do erro: " & Err.Description, vbCritical + vbOKOnly + vbSystemModal

End Sub


