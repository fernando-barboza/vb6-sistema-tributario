VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.ocx"
Begin VB.Form frmCadSenha 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Alterar senha"
   ClientHeight    =   1890
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3645
   HelpContextID   =   4001
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1890
   ScaleWidth      =   3645
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab SSTab1 
      Height          =   1635
      Left            =   120
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   120
      Width           =   3435
      _ExtentX        =   6059
      _ExtentY        =   2884
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Alterar senha"
      TabPicture(0)   =   "CadSenha.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lbl_Confirmacao"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lbl_Senha"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblstrSenha"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txt_ConfirmaSenha"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txt_Senha"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtNovaSenha"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      Begin VB.TextBox txtNovaSenha 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1230
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   810
         Width           =   1905
      End
      Begin VB.TextBox txt_Senha 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1230
         PasswordChar    =   "*"
         TabIndex        =   0
         Top             =   450
         Width           =   1905
      End
      Begin VB.TextBox txt_ConfirmaSenha 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1230
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   1170
         Width           =   1905
      End
      Begin VB.Label lblstrSenha 
         AutoSize        =   -1  'True
         Caption         =   "Nova"
         Height          =   195
         Left            =   765
         TabIndex        =   6
         Top             =   840
         Width           =   390
      End
      Begin VB.Label lbl_Senha 
         AutoSize        =   -1  'True
         Caption         =   "Atual"
         Height          =   195
         Left            =   795
         TabIndex        =   5
         Top             =   480
         Width           =   360
      End
      Begin VB.Label lbl_Confirmacao 
         AutoSize        =   -1  'True
         Caption         =   "Confirmação"
         Height          =   195
         Left            =   270
         TabIndex        =   4
         Top             =   1200
         Width           =   885
      End
   End
End
Attribute VB_Name = "frmCadSenha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Function blnUsuarioSenhaOk() As Boolean
    '--------------------------------------------------------
    'SE A SENHA INFORMADA ESTÁ CORRETA
    '--------------------------------------------------------
    Dim strSql              As String
    Dim adoResultado        As ADODB.Recordset
    Static bytQtdTentativa  As Byte
    On Error GoTo ErroblnUsuarioSenhaOk
    'Seleciona o usuário informado da tabela de usuários
    strSql = ""
    strSql = strSql & "SELECT strSenha FROM " & gstrUsuarios & " "
    strSql = strSql & "WHERE PKId = " & glngCodUsr
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        With adoResultado
            'Verifica se encontrou o usuário
            If .EOF = True Then
                ExibeMensagem "Erro na leitura da tabela de usuário."
            'Verifica se a senha está correta
            ElseIf gstrStringCripitografada(!strSenha) = UCase(Trim(txt_Senha)) Then
                blnUsuarioSenhaOk = True
                bytQtdTentativa = 0
            Else
                ExibeMensagem "A senha de acesso informada não está correta."
                txt_Senha.SetFocus
                bytQtdTentativa = bytQtdTentativa + 1
            End If
        End With
        adoResultado.Close
        Set adoResultado = Nothing
    End If
    If bytQtdTentativa = 3 Then
        bytQtdTentativa = 0
        Unload Me
    End If
    Exit Function
    
ErroblnUsuarioSenhaOk:
End Function

Public Sub MantemForm(ByVal strModoOperacao As String)
    Select Case UCase(strModoOperacao)
    Case gstrNovo
        LimpaTela
    Case gstrSalvar
        GravaNovaSenha
    Case gstrFechar
        Unload Me
    End Select
End Sub

Private Sub LimpaTela()
    txt_Senha = ""
    txtNovaSenha = ""
    txt_ConfirmaSenha = ""
End Sub

Private Sub GravaNovaSenha()

'******************************************************************************************
' Data: 14/03/2003
' Alteração: - Substituição da chamada explícita da stored procedure sp_password pela
'            chamada à função gstrStoredProcedure.
' Responsável: Everton Bianchini
'******************************************************************************************
    
    Dim strSql As String
    If blnDadosOK Then
        strSql = ""
        strSql = strSql & "UPDATE " & gstrUsuarios & " "
        strSql = strSql & "SET strSenha = "
        strSql = strSql & "'" & gstrStringCripitografada(txtNovaSenha, True) & "' "
        strSql = strSql & "WHERE PKID = " & glngCodUsr
        Set gobjBanco = New clsBanco
        If gobjBanco.Execute(strSql) Then
            strSql = ""
'            strSQL = strSQL & "sp_password '" & Trim(txt_Senha) & "','" & Trim(txtNovaSenha) & "'"
            'strSql = strSql & gstrStoredProcedure("sp_password", "'" & Trim(txt_Senha) & "', '" & Trim(txtNovaSenha) & "', '" & gstrLoginUser & "'")
            'If Not gobjBanco.Execute(strSql, True) Then
            '    ExibeDetalheErro "Não foi possível alterar a senha do login no " & _
            '        IIf((bytDBType = EDatabases.Oracle), "Oracle", "SQLServer") & "."
            'End If

            ExibeMensagem "Senha alterada com sucesso."
            Unload Me
        End If
    End If
End Sub

Private Sub Form_Activate()
    gintCodSeguranca = 382
    VirificaGradeListView Me
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrDeletar, gstrAplicar, gstrImprimir
End Sub

Private Sub Form_Load()

'******************************************************************************************
' Data: 14/03/2003
' Alteração: - Alteração da propriedade maxlength dos objetos txtNovaSenha e
'            txt_ConfirmaSenha para 30 caracteres caso o Banco de Dados corrente seja o
'            Oracle.
' Responsável: Everton Bianchini
'******************************************************************************************
    
    Me.Icon = MDIMenu.Icon
    
    If (bytDBType = EDatabases.Oracle) Then
        txtNovaSenha.MaxLength = 30
        txt_ConfirmaSenha.MaxLength = 30
    End If
    
End Sub

Private Sub txt_ConfirmaSenha_GotFocus()
    MarcaCampo txt_ConfirmaSenha
End Sub

Private Sub txt_ConfirmaSenha_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub txt_Senha_GotFocus()
    MarcaCampo txt_Senha
End Sub

Private Sub txt_Senha_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "S"
End Sub

Private Function blnDadosOK() As Boolean

'******************************************************************************************
' Data: 14/03/2003
' Alteração: - Implementação da verificação da validade do primeiro caracter para o Oracle.
' Responsável: Everton Bianchini
'******************************************************************************************

Dim blnChrValido As Boolean

    blnChrValido = (Asc(UCase(Left(Trim(txtNovaSenha), 1))) >= 65) And _
                (Asc(UCase(Left(Trim(txtNovaSenha), 1))) <= 90)

    If blnUsuarioSenhaOk Then
        If Trim(txtNovaSenha) = "" Then
            ExibeMensagem "Informe sua nova senha."
            txtNovaSenha.SetFocus
        ElseIf (Not blnChrValido) And (bytDBType = EDatabases.Oracle) Then
            ExibeMensagem "O primeiro caracter da senha deve ser uma letra."
            txtNovaSenha.SetFocus
        ElseIf UCase(Trim(txtNovaSenha)) <> UCase(Trim(txt_ConfirmaSenha)) Then
            ExibeMensagem "A senha não foi confirmada."
            txt_ConfirmaSenha.SetFocus
        Else
            blnDadosOK = True
        End If
        
    End If
End Function

Private Sub txtNovaSenha_GotFocus()
    MarcaCampo txtNovaSenha
End Sub

Private Sub txtNovaSenha_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub
