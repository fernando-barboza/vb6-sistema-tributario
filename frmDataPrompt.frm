VERSION 5.00
Begin VB.Form frmDataPrompt 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "frmDataPrompt"
   ClientHeight    =   1530
   ClientLeft      =   5475
   ClientTop       =   5265
   ClientWidth     =   2955
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1530
   ScaleWidth      =   2955
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtdtmData 
      Height          =   285
      Left            =   270
      MaxLength       =   10
      TabIndex        =   3
      Top             =   600
      Width           =   2505
   End
   Begin VB.CommandButton cmd_cancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   345
      Left            =   1650
      TabIndex        =   2
      Top             =   990
      Width           =   1155
   End
   Begin VB.CommandButton cmd_ok 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   240
      TabIndex        =   1
      Top             =   990
      Width           =   1155
   End
   Begin VB.Label lblPrompt 
      Caption         =   "lblPrompt"
      Height          =   285
      Left            =   330
      TabIndex        =   0
      Top             =   210
      Width           =   2445
   End
End
Attribute VB_Name = "frmDataPrompt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public dtmdataValidacao As Date
Public strEncerramento As String
Public strDescdataValidacao As String
Public strExercicio As String

Public Function DataPrompt(ByVal strPrompt As String, _
    Optional ByVal dtmdataValidacao As Date, _
    Optional ByVal strEncerramento As String, _
    Optional ByVal strDescdataValidacao As String, _
    Optional ByVal strExercicio As String) As String
    
    strDataPrompt = ""
    lblPrompt = strPrompt
    dtmdataValidacao = dtmdataValidacao
    strEncerramento = strEncerramento
    strDescdataValidacao = strDescdataValidacao
    strExercicio = strExercicio
    
    frmDataPrompt.Show vbModal
    DataPrompt = strDataPrompt
End Function

Private Sub cmd_cancelar_Click()
    strDataPrompt = ""
    Unload Me
End Sub

Private Sub cmd_ok_Click()
    If VerificaData = True Then
        strDataPrompt = txtdtmData
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    Me.Caption = App.Title
    txtdtmData.Text = gstrDataDoSistema
End Sub

Private Sub txtdtmData_GotFocus()
    MarcaCampo txtdtmData
End Sub

Private Sub txtdtmData_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "D", txtdtmData
End Sub

Private Sub txtdtmData_LostFocus()
    txtdtmData = gstrDataFormatada(txtdtmData)
End Sub


Private Function VerificaData() As Boolean
    Dim dtmDtEncerramento As Date

    If gblnDataValida(txtdtmData) = False Then
        ExibeMensagem "A data infromada não é valida."
        If txtdtmData.Enabled Then
            If txtdtmData.Enabled Then txtdtmData.SetFocus
        End If
        Exit Function
    End If
    
    If Not dtmdataValidacao = Empty Then
        If CDate(txtdtmData) < dtmdataValidacao Then
            ExibeMensagem "A data informada deve ser maior que a data d" & strDescdataValidacao & "  (" & gstrDataFormatada(dtmdataValidacao) & ")."
            If txtdtmData.Enabled Then txtdtmData.SetFocus
            Exit Function
        End If
    End If
   
    If Not strEncerramento = "" Then
        dtmDtEncerramento = VerificaDataEncerramento(strEncerramento, gintExercicio)
             
        If Not dtmDtEncerramento = Empty Then
             If CDate(txtdtmData) <= dtmDtEncerramento Then
                 ExibeMensagem "A data informada deve ser maior que a data de último encerramento (" & dtmDtEncerramento & ")."
                 If txtdtmData.Enabled Then txtdtmData.SetFocus
                 Exit Function
             End If
         End If
    End If
   
   
    If Not strExercicio = "" Then
        If Year(txtdtmData) <> CInt(strExercicio) Then
            ExibeMensagem "A data informada deve do execício de " & strExercicio & "."
            If txtdtmData.Enabled Then txtdtmData.SetFocus
            Exit Function
        End If
    End If
   
    VerificaData = True

End Function

