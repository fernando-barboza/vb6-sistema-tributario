VERSION 5.00
Begin VB.Form frmAlteracaoSenha 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Alteração de Senha"
   ClientHeight    =   2055
   ClientLeft      =   4800
   ClientTop       =   4845
   ClientWidth     =   3390
   Icon            =   "AlteracaoSenha.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2055
   ScaleWidth      =   3390
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd_OK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2190
      TabIndex        =   3
      Top             =   1620
      Width           =   1110
   End
   Begin VB.Frame fra_Senha 
      Caption         =   " Digite sua nova senha: "
      Height          =   1425
      Left            =   90
      TabIndex        =   2
      Top             =   90
      Width           =   3195
      Begin VB.TextBox txt_ConfirmaSenha 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1110
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   870
         Width           =   1905
      End
      Begin VB.TextBox txt_Senha 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1110
         PasswordChar    =   "*"
         TabIndex        =   0
         Top             =   450
         Width           =   1905
      End
      Begin VB.Label lbl_Confirmacao 
         AutoSize        =   -1  'True
         Caption         =   "Confirmação"
         Height          =   195
         Left            =   150
         TabIndex        =   5
         Top             =   900
         Width           =   885
      End
      Begin VB.Label lbl_Senha 
         AutoSize        =   -1  'True
         Caption         =   "Senha"
         Height          =   195
         Left            =   570
         TabIndex        =   4
         Top             =   480
         Width           =   465
      End
   End
End
Attribute VB_Name = "frmAlteracaoSenha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_OK_Click()

'******************************************************************************************
' Data: 14/03/2003
' Alteração: - Substituição da chamada explícita da stored procedure sp_password pela
'            chamada à função gstrStoredProcedure.
' Responsável: Everton Bianchini
'******************************************************************************************
    
    Dim strSql     As String
    Dim strNovaPwd As String
    
    If blnDadosOk Then
        strNovaPwd = Trim(txt_Senha)
        strSql = ""
        strSql = strSql & "UPDATE " & gstrUsuarios & " "
        strSql = strSql & "SET strSenha = '" & gstrStringCripitografada(txt_Senha, True) & "', "
        strSql = strSql & "blnReinicializarSenha = 0 "
        strSql = strSql & "WHERE PKID = " & glngCodUsr
        Set gobjBanco = New clsBanco
        gobjBanco.Execute strSql
        
        'strSql = ""
'        strSQL = strSQL & "sp_password '" & gstrPwdUser & "','" & strNovaPwd & "'"
        'strSql = strSql & gstrStoredProcedure("sp_password", "'" & gstrPwdUser & "', '" & strNovaPwd & "', '" & gstrLoginUser & "'")
        'If Not gobjBanco.Execute(strSql, True) Then
        '    ExibeDetalheErro "Não foi possível alterar a senha do login no SQLServer."
        'End If
        
        Unload Me
    End If
End Sub

Private Sub Form_Load()

'******************************************************************************************
' Data: 14/03/2003
' Alteração: - Alteração da propriedade maxlength dos objetos txtNovaSenha e
'            txt_ConfirmaSenha para 30 caracteres caso o Banco de Dados corrente seja o
'            Oracle.
' Responsável: Everton Bianchini
'******************************************************************************************
    
    Screen.MousePointer = 0

    If (bytDBType = EDatabases.Oracle) Then
        txt_Senha.MaxLength = 30
        txt_ConfirmaSenha.MaxLength = 30
    End If

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Select Case UnloadMode
        Case 1
        Case Else
            Cancel = 1
    End Select
End Sub

Private Sub txt_ConfirmaSenha_GotFocus()
    MarcaCampo txt_ConfirmaSenha
End Sub

Private Sub txt_ConfirmaSenha_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txt_ConfirmaSenha
End Sub

Private Sub txt_Senha_GotFocus()
    MarcaCampo txt_Senha
End Sub

Private Sub txt_Senha_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txt_Senha
End Sub

Private Function blnDadosOk() As Boolean

'******************************************************************************************
' Data: 14/03/2003
' Alteração: - Implementação da verificação da validade do primeiro caracter para o Oracle.
' Responsável: Everton Bianchini
'******************************************************************************************

Dim blnChrValido As Boolean

    blnChrValido = (Asc(UCase(Left(Trim(txt_Senha), 1))) >= 65) And _
                (Asc(UCase(Left(Trim(txt_Senha), 1))) <= 90)

    If Trim(txt_Senha) = "" Then
        ExibeMensagem "Informe sua nova senha."
        Exit Function
'    ElseIf (Not blnChrValido) And (bytDBType = EDatabases.Oracle) Then
'        ExibeMensagem "O primeiro caracter da senha deve ser uma letra."
'        txt_Senha.SetFocus
'        Exit Function
    ElseIf UCase(Trim(txt_Senha)) <> UCase(Trim(txt_ConfirmaSenha)) Then
        ExibeMensagem "As senhas não são idênticas."
        Exit Function
    ElseIf Len(Trim(txt_Senha)) > 10 Then
        ExibeMensagem "A senha deve conter no máximo 10 caracteres!"
        Exit Function
    End If
    blnDadosOk = True
End Function
