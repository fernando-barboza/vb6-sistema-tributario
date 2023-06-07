VERSION 5.00
Begin VB.Form frmFindForm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Find"
   ClientHeight    =   2025
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4440
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2025
   ScaleWidth      =   4440
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3240
      TabIndex        =   9
      Top             =   600
      Width           =   1125
   End
   Begin VB.TextBox txtReplace 
      Enabled         =   0   'False
      Height          =   285
      Left            =   960
      TabIndex        =   5
      Top             =   540
      Width           =   2160
   End
   Begin VB.CommandButton cmdFindNext 
      Caption         =   "Find Next"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3240
      TabIndex        =   4
      Top             =   120
      Width           =   1125
   End
   Begin VB.CommandButton cmdReplace 
      Caption         =   "Replace"
      Height          =   375
      Left            =   3240
      TabIndex        =   3
      Top             =   1080
      Width           =   1125
   End
   Begin VB.CommandButton cmdReplaceAll 
      Caption         =   "Replace All"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3240
      TabIndex        =   2
      Top             =   1560
      Width           =   1125
   End
   Begin VB.ComboBox cboSearch 
      Height          =   315
      ItemData        =   "FindForm.frx":0000
      Left            =   960
      List            =   "FindForm.frx":0007
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   870
      Width           =   2145
   End
   Begin VB.TextBox txtFind 
      Height          =   285
      Left            =   960
      TabIndex        =   0
      Top             =   195
      Width           =   2160
   End
   Begin VB.Label lbl1 
      Caption         =   "Fi&nd:"
      Height          =   255
      Left            =   75
      TabIndex        =   8
      Top             =   210
      Width           =   780
   End
   Begin VB.Label lblReplace 
      Caption         =   "Re&place:"
      Enabled         =   0   'False
      Height          =   255
      Left            =   75
      TabIndex        =   7
      Top             =   555
      Width           =   780
   End
   Begin VB.Label lbl3 
      Caption         =   "&Direction:"
      Height          =   255
      Left            =   75
      TabIndex        =   6
      Top             =   900
      Width           =   780
   End
End
Attribute VB_Name = "frmFindForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Portions of the code are copyright 1995, Microsoft Corporation
Option Explicit

'*************************************************
' Purpose:  Finds a string in the text and replaces
'           it if necessary.
' Effects:  Text on the calling form may change.
' Inputs:   strFind:    The text to find
'           strReplace: The text to replace the found
'                       text with
'           intStart:   Starting position to search
'           intEnd:     Ending extent of search
' Returns:  True if text was found (and replaced)
'           Flase if no text was found.
'*************************************************
Function FindString(strFind As String, strReplace As String, lngStart As Long, lngEnd As Long) As Boolean
    Dim lngPos As Long 'position
    With MDIMenu.ActiveForm.rtf
        ' locate search string
        lngPos = InStr(lngStart, .Text, strFind, 1)
        If lngPos = 0 Then
            ' not found
            FindString = False
        Else
            ' it's found, but is it past the end of
            ' the region that we're concerned with
            If Not ((lngPos + Len(strFind)) > lngEnd) Then
                ' it in bounds
                FindString = True
                .SelStart = lngPos - 1
                .SelLength = Len(strFind)
                ' if the replace string is blank
                ' then this is just a find.
                If strReplace <> "" Then 'Replace It
                    .SelText = strReplace
                End If
            End If
        End If
    End With
End Function


'*************************************************
' Purpose:  Unloads the form
'*************************************************
Private Sub cmdCancel_Click()
    Unload Me
End Sub

'*************************************************
' Purpose:  Sets up the call to Find a string.
'*************************************************
Private Sub cmdFindNext_Click()
    Dim lngLen As Long ' ending position
    Dim lngStart As Long ' starting position

    ' set end to the length of the text in the box.
    lngLen = Len(MDIMenu.ActiveForm.rtf.Text)
    
    Select Case cboSearch.ListIndex
        Case 0 'search all
            lngStart = 1
        Case 1 'search backward
            lngStart = 1
            ' set the end to the end of the current selection.
            lngLen = MDIMenu.ActiveForm.rtf.SelStart
            lngLen = lngLen + MDIMenu.ActiveForm.rtf.SelLength
        Case 2 'search forward
            ' set the start to the end of the current selection
            lngStart = MDIMenu.ActiveForm.rtf.SelStart
            lngStart = lngStart + MDIMenu.ActiveForm.rtf.SelLength
    End Select
    
    'Call the function to do the find.
    If Not FindString(txtFind.Text, "", lngStart, lngLen) Then
        MsgBox "Palavra não encontrada", 0, "Localizar"
    Else
    '    Unload Me
    End If
End Sub

'*************************************************
' Purpose:  Set's up the parameters to call the function
'           to replace a string.
'*************************************************
Private Sub cmdReplace_Click()
    Dim lngLen As Long ' ending pos
    Dim lngStart As Long ' starting pos

    ' set end to the length of the text in the box
    lngLen = Len(MDIMenu.ActiveForm.rtf.Text)

    Select Case cboSearch.ListIndex
        Case 0 'search all
            lngStart = 1
        Case 1 'search backward
            lngStart = 1
            ' set end to the end of the selection
            lngLen = MDIMenu.ActiveForm.rtf.SelStart
            lngLen = lngLen + MDIMenu.ActiveForm.rtf.SelLength
        Case 2 'search forward
            ' set the beginning to the end of the selection
            lngStart = MDIMenu.ActiveForm.rtf.SelStart
            lngStart = lngStart + MDIMenu.ActiveForm.rtf.SelLength
    End Select

    'call the function to do the replace
    If Not FindString(txtFind.Text, txtReplace.Text, lngStart, lngLen) Then
        MsgBox "Palavra não encontrada", 0, "Substituir"
    Else
        Unload Me
    End If
End Sub

'*************************************************
' Purpose:  Set's up the call to the function that
'           will replace the text, and calls it
'           until no more replacements exist.
'*************************************************
Private Sub cmdReplaceAll_Click()
    Dim lngLen As Long 'ending pos
    Dim lngStart As Long ' starting pos
    Dim strMsg As String ' temp string for prompt
    ' build warning prompt
    strMsg = "This process will begin at the top"
    strMsg = strMsg & " of your document"
    strMsg = strMsg & " and replace everything."
    strMsg = strMsg & "  Do you want to continue?"
    ' Ask the user if he/she is sure.
    If MsgBox(strMsg, vbYesNo, "Are You Sure?") = vbYes Then
        'User is sure, replace everything.
        Do
            If lngStart = 0 Then
                lngStart = 1
            Else
                ' Set the new starting position
                ' to the end of the selection
                lngStart = MDIMenu.ActiveForm.rtf.SelStart
                lngStart = lngStart + MDIMenu.ActiveForm.rtf.SelLength
            End If
            ' set the end to the end of the doc
            lngLen = Len(MDIMenu.ActiveForm.rtf.Text)
        ' loop while replacements are made
        Loop While FindString(txtFind.Text, txtReplace.Text, lngStart, lngLen)
    End If
    ' unload form
    Unload Me
End Sub


'*************************************************
' Purpose:  Initialize form
'*************************************************
Private Sub Form_Load()
    ' set default search to 'All'
    cboSearch.ListIndex = 0
End Sub


'*************************************************
' Purpose:  Enable or Disable buttons.
'*************************************************
Private Sub txtFind_Change()
    If Len(txtFind.Text) > 0 Then
        ' search text exists, enable FindNext button
        cmdFindNext.Enabled = True
    Else
        ' no search text exists, disable
        ' FindNext button
        cmdFindNext.Enabled = False
    End If
End Sub


'*************************************************
' Purpose:  Select the entire contents of the box.
'*************************************************
Private Sub txtFind_GotFocus()
    txtFind.SelStart = 0
    txtFind.SelLength = Len(txtFind.Text)
End Sub


'*************************************************
' Purpose:  Enables or disables buttons.
'*************************************************
Private Sub txtReplace_Change()
    If Len(txtReplace.Text) > 0 Then
        ' replace text exists, enable buttons
        cmdReplace.Enabled = True
        cmdReplaceAll.Enabled = True
    Else
        ' no replace text exists, disable buttons
        cmdReplace.Enabled = False
        cmdReplaceAll.Enabled = False
    End If
End Sub

'*************************************************
' Purpose:  Select the entire contents of the box.
'*************************************************
Private Sub txtReplace_GotFocus()
    txtReplace.SelStart = 0
    txtReplace.SelLength = Len(txtReplace.Text)
End Sub

