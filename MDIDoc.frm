VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Begin VB.Form frmMDIDoc 
   Caption         =   "Document"
   ClientHeight    =   7635
   ClientLeft      =   60
   ClientTop       =   495
   ClientWidth     =   11880
   Icon            =   "MDIDoc.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7635
   ScaleWidth      =   11880
   Begin VB.Timer tmr 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   3600
      Top             =   360
   End
   Begin RichTextLib.RichTextBox rtf 
      Height          =   3495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   6165
      _Version        =   393217
      Enabled         =   -1  'True
      HideSelection   =   0   'False
      ScrollBars      =   3
      MaxLength       =   4000
      TextRTF         =   $"MDIDoc.frx":1042
   End
End
Attribute VB_Name = "frmMDIDoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim adoResultado    As ADODB.Recordset
Dim intCodSeguranca As Integer

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Enum eTextMode
    TM_PLAINTEXT = 1
    TM_RICHTEXT = 2                ' /* default behavior */
    TM_SINGLELEVELUNDO = 4
    TM_MULTILEVELUNDO = 8          ' /* default behavior */
    TM_SINGLECODEPAGE = 16
    TM_MULTICODEPAGE = 32          ' /* default behavior */
End Enum

Private Const WM_USER = &H400
Private Const WM_PASTE = &H302
Private Const WM_COPY = &H301
Private Const WM_CUT = &H300

Private Const EM_LINEINDEX = &HBB&
Private Const EM_CANUNDO = &HC6
Private Const EM_UNDO = &HC7
Private Const EM_LINEFROMCHAR = &HC9&
Private Const EM_CANPASTE = (WM_USER + 50)
Private Const EM_HIDESELECTION = (WM_USER + 63)
Private Const EM_REQUESTRESIZE = (WM_USER + 65)
Private Const EM_SETUNDOLIMIT = (WM_USER + 82)
Private Const EM_REDO = (WM_USER + 84)
Private Const EM_CANREDO = (WM_USER + 85)
Private Const EM_GETUNDONAME = (WM_USER + 86)
Private Const EM_GETREDONAME = (WM_USER + 87)
Private Const EM_STOPGROUPTYPING = (WM_USER + 88)
Private Const EM_SETTEXTMODE = (WM_USER + 89)
Private Const EM_GETTEXTMODE = (WM_USER + 90)
Private Const EM_AUTOURLDETECT = (WM_USER + 91)

Private m_ab As ActiveBar2LibraryCtl.ActiveBar2

Implements IMDIDocument

Dim blnAlterou As Boolean
Dim blnJaSalvou As Boolean
Dim blnAbriu As Boolean

Private Sub MostraIconesEditar()
Dim tool As ActiveBar2LibraryCtl.tool
Dim iCat As Integer
Dim keys(0) As New ShortCut
Dim B As ActiveBar2LibraryCtl.Band

If blnAbriu Then
    Exit Sub
End If
    iCat = 303
    Set tool = m_ab.Tools.Add(iCat + 1, "1miESeparador")
    tool.Caption = "": tool.Category = "1Edit"
    tool.Visible = False
    
    Set tool = m_ab.Tools.Add(iCat + 2, "1miFoBold")
    tool.Caption = "&Negrito": tool.Category = "1Edit": tool.ToolTipText = "Negrito"
    tool.SetPicture ddITNormal, MDIMenu.img_ListaIconesGeral.ListImages("NEGRITO").Picture
    keys(0) = "Control+N"
    tool.ShortCuts = keys
    tool.Visible = False
    
    Set tool = m_ab.Tools.Add(iCat + 3, "1miFoItalic")
    tool.Caption = "&Negrito": tool.Category = "1Edit": tool.ToolTipText = "Itálico"
    tool.SetPicture ddITNormal, MDIMenu.img_ListaIconesGeral.ListImages("ITALICO").Picture
    keys(0) = "Control+I"
    tool.ShortCuts = keys
    
    Set tool = m_ab.Tools.Add(iCat + 4, "1miFoUnderline")
    tool.Caption = "S&ublinhado": tool.Category = "1Edit": tool.ToolTipText = "Sublinhado"
    tool.SetPicture ddITNormal, MDIMenu.img_ListaIconesGeral.ListImages("SUBLINHADO").Picture
    keys(0) = "Control+U"
    tool.ShortCuts = keys
    
    Set B = m_ab.Bands("1mnuEdit")
    With B.Tools
        .Insert .Count, m_ab.Tools("1miESeparador")
        B.Tools("1miESeparador").ControlType = ddTTSeparator
        m_ab.Tools("1miESeparador").Visible = True
        
        .Insert .Count, m_ab.Tools("1miFoBold")
        B.Tools("1miFoBold").ControlType = ddTTButton
        m_ab.Tools("1miFoBold").Visible = True
    
        .Insert .Count, m_ab.Tools("1miFoItalic")
        B.Tools("1miFoBold").ControlType = ddTTButton
        m_ab.Tools("1miFoBold").Visible = True
    
        .Insert .Count, m_ab.Tools("1miFoUnderline")
        B.Tools("1miFoUnderline").ControlType = ddTTButton
        m_ab.Tools("1miFoUnderline").Visible = True
    End With
    m_ab.RecalcLayout
    m_ab.Refresh


'        .Tools("miFoBold").Checked = IsNull(rtf.SelBold) Or rtf.SelBold
'        .Tools("miFoItalic").Checked = IsNull(rtf.SelItalic) Or rtf.SelItalic
'        .Tools("miFoUnderline").Checked = IsNull(rtf.SelUnderline) Or rtf.SelUnderline
'
'        .Tools("miECut").Enabled = (rtf.SelLength <> 0)
'        .Tools("miECopy").Enabled = (rtf.SelLength <> 0)
'        .Tools("miEPaste").Enabled = (SendMessage(rtf.hWnd, EM_CANPASTE, 0, 0) = 1)
'        .Tools("miEUndo").Enabled = (SendMessage(rtf.hWnd, EM_CANUNDO, 0, 0) = 1)
'        .Tools("miERedo").Enabled = (SendMessage(rtf.hWnd, EM_CANREDO, 0, 0) = 1)
'
'        .Tools("miFoLeft").Checked = (rtf.SelAlignment = 0)
'        .Tools("miFoCenter").Checked = (rtf.SelAlignment = 2)
'        .Tools("miFoRight").Checked = (rtf.SelAlignment = 1)
'
'        .Tools("miFoBullets").Checked = IIf(IsNull(rtf.SelBullet), False, rtf.SelBullet)
'
'        .Tools("miFoFontName").Text = IIf(IsNull(rtf.SelFontName), "", rtf.SelFontName)
'        .Tools("miFoFontSize").Text = IIf(IsNull(rtf.SelFontSize), "", rtf.SelFontSize)

'        .Insert .Count, ab.Tools("1miESeparador")
'        b.Tools("1miESeparador").ControlType = ddTTSeparator
'        .Insert .Count, ab.Tools("1miFoBold")
End Sub

Private Sub Form_Activate()
    gintCodSeguranca = intCodSeguranca
MDIMenu.blnTexto = True
End Sub

Private Sub Form_Deactivate()
MDIMenu.blnTexto = False
End Sub

Private Sub Form_Load()
Dim strSQL As String

    intCodSeguranca = gintCodSeguranca
    Me.HelpContextID = intCodSeguranca
IMDIDocument_InitDoc MDIMenu.actBarra, "", True

MostraIconesEditar
blnAbriu = True

If MDIMenu.Tag = "ATENDIMENTO" Then
'SELECT NA TABELA DE ATENDIMENTO
    Dim strQueryAtendimento As String
    strSQL = ""
    strSQL = " SELECT * FROM " & gstrTextoAtendimento
    strQueryAtendimento = strSQL
        Set gobjBanco = New clsBanco
        rtf.Text = ""
        If gobjBanco.CriaADO(strQueryAtendimento, 5, adoResultado) Then
            With adoResultado
                Do While Not .EOF
                    rtf.Text = gstrENulo(!strTextoAtendimento)
                    .MoveNext
                Loop
            End With
        End If
ElseIf MDIMenu.Tag = "CARTA" Then
'SELECT NA TABELA DE CARTA
    Dim strQueryCarta As String
    strSQL = ""
    strSQL = " SELECT * FROM " & gstrTextoCarta
    strQueryCarta = strSQL
        Set gobjBanco = New clsBanco
        rtf.Text = ""
        If gobjBanco.CriaADO(strQueryCarta, 5, adoResultado) Then
            With adoResultado
                Do While Not .EOF
                    rtf.Text = gstrENulo(!strTextoCarta)
                    .MoveNext
                Loop
            End With
        End If
ElseIf MDIMenu.Tag = "SOLICITACAO" Then
'SELECT NA TABELA DE SOLICITACAO
    Dim strQuerySolicitacao As String
    strSQL = ""
    strSQL = " SELECT * FROM " & gstrTextoSolicitacao
    strQuerySolicitacao = strSQL
        Set gobjBanco = New clsBanco
        rtf.Text = ""
        If gobjBanco.CriaADO(strQuerySolicitacao, 5, adoResultado) Then
            With adoResultado
                Do While Not .EOF
                    rtf.Text = gstrENulo(!strTextoSolicitacao)
                    .MoveNext
                Loop
            End With
        End If
End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim strMsg As String
Dim intResp As Integer
Dim tolResposta As ActiveBar2LibraryCtl.tool

If blnAlterou Then
    strMsg = ""
    strMsg = strMsg & " Deseja salvar as alterações do Documento? " ' de " & Me.Caption

    intResp = MsgBox(strMsg, vbYesNoCancel, "Editor de Textos")
    
    Select Case intResp
        Case 2 'Cancelar
            Cancel = True
        Case 6 'Salvar
        
            MantemForm ("SALVAR")
            
'            Set tolResposta = New ActiveBar2LibraryCtl.Tool
'            If blnJaSalvou Then
'                tolResposta.Name = "miFSave"
'            Else
'                tolResposta.Name = "miFSaveAs"
'            End If
'
'            MDIMenu.actBarra_ToolClick tolResposta

        Case 7
    End Select
End If

End Sub

Private Sub Form_Resize()
    On Error Resume Next
    rtf.Move 0, 0, ScaleWidth, ScaleHeight
End Sub

Public Function IMDIDocument_CommandHandler(tool As ActiveBar2LibraryCtl.ITool) As Boolean
    
    IMDIDocument_CommandHandler = True
    If tool.Category = "Color" Then
        FormatColor tool
        Exit Function
    End If
    Select Case tool.Name
    ' File
    Case "1miFSave", "miFSave"
        Me.Tag = ""
        FileSaveAs
    Case "1miFSaveAs", "miFSaveAs"
        
        FileSaveAs
    Case "1miFPrint", "miFPrint": FilePrint
    Case "1miFPrintPreview", "miFPrintPreview": FilePrintPreview
    Case "1miFPageSetup", "miFPageSetup": FilePageSetup
    Case "1miFClose", "miFClose"
        Unload Me
        Exit Function
    ' Edit
    Case "1miEUndo", "miEUndo": EditUndo
    Case "1miERedo", "miERedo": EditRedo
    Case "1miECut", "miECut": EditCut
    Case "1miECopy", "miECopy": EditCopy
    Case "1miEPaste", "miEPaste": EditPaste
    Case "1miEClear", "miEClear": EditClear
    Case "1miESelectAll", "miESelectAll": EditSelectAll
    Case "1miEFind", "miEFind": EditFind ""
    Case "1miEFindNext", "miEFindNext": EditFindNext
    Case "1miEReplace", "miEReplace": EditReplace ""
    
    ' Insert
    Case "miIDate": InsertDate
    Case "miITime": InsertTime
    Case "miIPicture": InsertPicture
    
    ' Format
    Case "miFoFont": FormatFont
    Case "miFoFontName": FormatFont
    Case "miFoFontSize": FormatFont
    Case "miFoParagraph": FormatParagraph
    Case "1miFoBold", "miFoBold": FormatBold
    Case "1miFoItalic", "miFoItalic": FormatItalic
    Case "1miFoUnderline", "miFoUnderline": FormatUnderline
    Case "1miFoLeft", "miFoLeft": FormatAlign 0
    Case "1miFoCenter", "miFoCenter": FormatAlign 1
    Case "1miFoRight", "miFoRight": FormatAlign 2
    Case "1miFoBullets", "miFoBullets": FormatBullets
    Case "1miFoTabs", "miFoTabs": FormatTabs
    Case Else
        IMDIDocument_CommandHandler = False
    End Select
    UpdateToolbar
    tmr.Enabled = False
End Function

Private Sub UpdateToolbar()
    With m_ab
'        .Tools("1miFoBold").Checked = IsNull(rtf.SelBold) Or rtf.SelBold
'        .Tools("1miFoItalic").Checked = IsNull(rtf.SelItalic) Or rtf.SelItalic
'        .Tools("1miFoUnderline").Checked = IsNull(rtf.SelUnderline) Or rtf.SelUnderline
'
'        .Tools("miECut").Enabled = (rtf.SelLength <> 0)
'        .Tools("miECopy").Enabled = (rtf.SelLength <> 0)
'        .Tools("miEPaste").Enabled = (SendMessage(rtf.hWnd, EM_CANPASTE, 0, 0) = 1)
'        .Tools("miEUndo").Enabled = (SendMessage(rtf.hWnd, EM_CANUNDO, 0, 0) = 1)
'        .Tools("miERedo").Enabled = (SendMessage(rtf.hWnd, EM_CANREDO, 0, 0) = 1)
'
'        .Tools("miFoLeft").Checked = (rtf.SelAlignment = 0)
'        .Tools("miFoCenter").Checked = (rtf.SelAlignment = 2)
'        .Tools("miFoRight").Checked = (rtf.SelAlignment = 1)
'
'        .Tools("miFoBullets").Checked = IIf(IsNull(rtf.SelBullet), False, rtf.SelBullet)
'
'        .Tools("miFoFontName").Text = IIf(IsNull(rtf.SelFontName), "", rtf.SelFontName)
'        .Tools("miFoFontSize").Text = IIf(IsNull(rtf.SelFontSize), "", rtf.SelFontSize)
'        .Refresh
    End With
    
    tmr.Enabled = False
End Sub

Private Function IMDIDocument_InitDoc(ab As ActiveBar2LibraryCtl.IActiveBar2, sFile As String, bNew As Boolean) As Boolean
Dim bRet As Boolean

blnJaSalvou = False
    
    If Not ab Is Nothing Then
        Set m_ab = ab
        'm_ab.RegisterChildMenu hWnd, "mnuChildDoc"
        'm_ab.RecalcLayout
        bRet = True
    End If
    If bNew Then
        rtf.Text = ""
    Else
        ' open file
        rtf.LoadFile sFile
    End If
    rtf.DataChanged = False
    Caption = sFile
    Me.Top = 0
    Me.Show
    
    IMDIDocument_InitDoc = bRet
End Function

Private Function FileSave(Optional sSaveAsName As String) As Boolean
    On Error GoTo ehFileSave 'set error trap

    blnJaSalvou = True

    If IsMissing(sSaveAsName) Or sSaveAsName = "" Then
        'if no save name specified
        If InStr(Me.Caption, "(untitled)") > 0 Then
            'if no previous name existed
            sSaveAsName = "Documento.rtf"
            If Not MDIMenu.cdlg.VBGetSaveFileName(sSaveAsName, _
                "RichEdit Documento", True, "Rich Text File(*.rtf)|*.rtf", , _
                App.Path, "Salvar Como...", "RTF", Me.hwnd) Then
                FileSave = False
                Exit Function
            End If
        Else
            ' set SaveAsName to the name that
            ' the file was already given
            sSaveAsName = Me.Caption
        End If
    End If
    
    ' save file
    rtf.SaveFile CStr(sSaveAsName)
    
    ' change the caption to reflect name
    Me.Caption = CStr(sSaveAsName)
    
    ' set return value to true
    FileSave = True
    rtf.DataChanged = False
    Exit Function
ehFileSave:
    ' set return value to false
    FileSave = False
    Exit Function
End Function

Private Sub FileSaveAs()
Dim sSaveAsName As String

    On Error GoTo ehFileSaveAs 'set error trap

    blnJaSalvou = True

    sSaveAsName = "Documento.rtf"
    If Not MDIMenu.cdlg.VBGetSaveFileName(sSaveAsName, _
        "RichEdit Documento", True, "Rich Text File(*.rtf)|*.rtf", , _
        App.Path, "Salvar Como...", "RTF", Me.hwnd) Then
                
        Exit Sub
    End If
    Me.Tag = ""
    ' save file
    rtf.SaveFile CStr(sSaveAsName)
    
    ' change the caption to reflect name
    Me.Caption = CStr(sSaveAsName)
    
    ' set return value to true
    rtf.DataChanged = False
    
ehFileSaveAs:
    Exit Sub
End Sub

Private Sub FilePrint()
Dim flags As Long
Dim hdc As Long

    On Error GoTo ehFilePrint ' set error trap
    With MDIMenu.cdlg
        ' show printer dialog
        If .VBPrintDlg(hdc, IIf(rtf.SelLength = 0, eprAll, eprSelection)) = True Then
            ' print selection was selected
            If rtf.SelLength <> 0 Then
                rtf.SelPrint hdc
            Else
                ' print all was selected
                rtf.SelLength = 0
                rtf.SelPrint hdc
            End If
        End If
    End With
ehFilePrint: 'cancel pressed
    Exit Sub
End Sub

Private Sub FilePrintPreview()
    Dim fPreview As New frmPreview
    Dim doc As IMDIDocument
    Set doc = fPreview
    doc.InitDoc m_ab, Me.Caption, False
    'fPreview.PrintPreview rtf, 1440, 1440, 1440, 1440, vbPRORLandscape
End Sub

Private Sub FilePageSetup()
    frmPageSetup.Show vbModal
End Sub

Private Sub EditRedo()
    SendMessage rtf.hwnd, EM_REDO, 0, 0
End Sub

Private Sub EditUndo()
Dim hr As Long
    hr = SendMessage(rtf.hwnd, EM_GETUNDONAME, 0&, 0&)
    ' Debug.Print hr, Choose(hr + 1, "Unknown", "Typing", "Delete", "Drag Drop", "Cut", "Paste")
    SendMessage rtf.hwnd, EM_UNDO, 0, 0
End Sub

Private Sub EditCut()
    SendMessage rtf.hwnd, WM_CUT, 0, 0
'    rtf.SetFocus
End Sub

Private Sub EditCopy()
    SendMessage rtf.hwnd, WM_COPY, 0, 0
End Sub

Private Sub EditPaste()
    SendMessage rtf.hwnd, WM_PASTE, 0, 0
End Sub

Private Sub EditClear()
    rtf.SelText = ""
End Sub

Private Sub EditSelectAll()
    rtf.SelStart = 0
    rtf.SelLength = Len(rtf.Text)
End Sub

Private Sub EditFind(strSearch As String)
    frmFindForm.txtFind = strSearch 'set find text
    frmFindForm.Show
End Sub

Private Sub EditFindNext()
    frmFindForm.cboSearch.ListIndex = 2
    If frmFindForm.txtFind <> "" Then
        frmFindForm.cmdFindNext.Value = True
    End If
    frmFindForm.Show
End Sub

Private Sub EditReplace(strSearch As String)
    With frmFindForm
        .txtFind = strSearch 'set find text
        .txtReplace.Enabled = True 'enable replace
        .lblReplace.Enabled = True 'enable replace
        .Show vbModal 'show modally
    End With
End Sub

Private Sub InsertDate()
    rtf.SelText = Format(Now, "Long Date")
End Sub

Private Sub InsertTime()
    rtf.SelText = Format$(Now, "Hh:Nn:Ss")
End Sub

Private Sub InsertPicture()
' thanks to "Joachim Thiele" www.N-H-P.de
On Error Resume Next
Dim sFile As String

    If MDIMenu.cdlg.VBGetOpenFileName(sFile, "Figura", True, False, False, True, "Picture Files(*.BMP;*.GIF;*.JPG)|*.BMP;*.GIF;*.JPG", , App.Path, "Inserir Figura", , Me.hwnd) Then
        Clipboard.Clear
        DoEvents
        Clipboard.SetData LoadPicture(sFile)
        If Clipboard.GetFormat(vbCFBitmap) = True Then ' Bitmap
            rtf.SetFocus
            EditPaste
        Else
            MsgBox "No Picture selected!"
        End If
    End If
    
End Sub

Private Sub FormatFont()
Dim fnt As New StdFont
Dim clr As Long

    On Error Resume Next
    With rtf
        fnt.Name = .SelFontName
        fnt.Strikethrough = .SelStrikeThru
        fnt.Bold = .SelBold
        fnt.Italic = .SelItalic
        fnt.Underline = .SelUnderline
        fnt.Size = .SelFontSize
        clr = .SelColor
        If MDIMenu.cdlg.VBChooseFont(fnt, , Me.hwnd, clr, 5, 72, CF_ScreenFonts Or CF_EFFECTS) Then
                .SelFontName = fnt.Name
                .SelBold = fnt.Bold
                .SelColor = clr
                .SelItalic = fnt.Italic
                .SelUnderline = fnt.Underline
                .SelFontSize = fnt.Size
                .SelStrikeThru = fnt.Strikethrough
        End If
    End With
    Set fnt = Nothing
End Sub


Private Sub FormatParagraph()
    frmParagraph.Show vbModal
'    rtf.SetFocus
End Sub

Private Sub FormatBold()
    With rtf
        If (IsNull(.SelBold) = True) Or (.SelBold = False) Then
            ' selection is bold or mixed, so set bold
            .SelBold = True
        ElseIf .SelBold = True Then
            'selection is bold, so toggle it
            .SelBold = False
        End If
        .SetFocus
    End With
End Sub

Private Sub FormatItalic()
    With rtf
        If (IsNull(.SelItalic) = True) Or (.SelItalic = False) Then
            ' selection is italic or mixed, so set italic
            .SelItalic = True
        ElseIf .SelItalic = True Then
            'selection is italic, so toggle it
            .SelItalic = False
        End If
'        .SetFocus
    End With
End Sub

Private Sub FormatUnderline()
    With rtf
        If (IsNull(.SelUnderline) = True) Or (.SelUnderline = False) Then
            ' selection is Underline or mixed, so set italic
            .SelUnderline = True
        ElseIf .SelUnderline = True Then
            'selection is Underline, so toggle it
            .SelUnderline = False
        End If
'        .SetFocus
    End With
End Sub

Private Sub FormatColor(tool As ActiveBar2LibraryCtl.tool)
    Dim lClr As Long
    lClr = CLng(tool.TagVariant)
    rtf.SelColor = lClr
End Sub

Private Sub FormatAlign(intIndex As Integer)
    Select Case intIndex
        Case 0 'left
            'set alignment
            rtf.SelAlignment = rtfLeft
        Case 1 'center
            'set alignment
            rtf.SelAlignment = rtfCenter
        Case 2 'right
            'set images
            'set alignment
            rtf.SelAlignment = rtfRight
    End Select
End Sub

Private Sub FormatBullets()
    With rtf
        If (IsNull(.SelBullet) = True) Or (.SelBullet = False) Then
            ' selection is mixed or not bulleted
            ' so set it.
            .SelBullet = True
        ElseIf .SelBullet = True Then
            ' selection is bold, toggle it
            .SelBullet = False
            .SelHangingIndent = False
        End If
    End With
End Sub

Private Sub FormatTabs()
    'frmTabs.Show vbModal
End Sub

Private Sub rtf_SelChange()
   ' tmr.Enabled = False
   ' tmr.Enabled = True
   UpdateToolbar
   blnAlterou = True
End Sub

Private Sub tmr_Timer()
    UpdateToolbar
End Sub

Public Sub MantemForm(strModoOperacao)
Dim strSQL As String
    If UCase(strModoOperacao) = "SALVAR" Then
        If MDIMenu.Tag = "ATENDIMENTO" Then
'            SELECT NA TABELA DE ATENDIMENTO
            If Not rtf.Text <> "" Then
                strSQL = ""
                strSQL = strSQL & "INSERT INTO " & gstrTextoAtendimento & " "
                strSQL = strSQL & "(strTextoAtendimento "
                strSQL = strSQL & ") VALUES ('"
                strSQL = strSQL & rtf.Text
                strSQL = strSQL & "')"
            Else
                strSQL = ""
                strSQL = strSQL & "UPDATE " & gstrTextoAtendimento & " SET "
                strSQL = strSQL & "strTextoAtendimento = " & "('" & rtf.Text & "')"
            End If
        ElseIf MDIMenu.Tag = "SOLICITACAO" Then
            'SELECT NA TABELA DE SOLICITACAO
            If Not rtf.Text <> "" Then
                strSQL = ""
                strSQL = strSQL & "INSERT INTO " & gstrTextoSolicitacao & " "
                strSQL = strSQL & "(strTextoSolicitacao "
                strSQL = strSQL & ") VALUES ('"
                strSQL = strSQL & rtf.Text
                strSQL = strSQL & "')"
            Else
                strSQL = ""
                strSQL = strSQL & "UPDATE " & gstrTextoSolicitacao & " SET "
                strSQL = strSQL & "strTextoSolicitacao = " & "('" & rtf.Text & "')"
            End If
        ElseIf MDIMenu.Tag = "CARTA" Then
            'SELECT NA TABELA DE CARTA
            If Not rtf.Text <> "" Then
                strSQL = ""
                strSQL = strSQL & "INSERT INTO " & gstrTextoCarta & " "
                strSQL = strSQL & "(strTextoCarta "
                strSQL = strSQL & ") VALUES ('"
                strSQL = strSQL & rtf.Text
                strSQL = strSQL & "')"
            Else
                strSQL = ""
                strSQL = strSQL & "UPDATE " & gstrTextoCarta & " SET "
                strSQL = strSQL & "strTextoCarta = " & "('" & rtf.Text & "')"
            End If
        End If
        Set gobjBanco = New clsBanco
        gobjBanco.Execute strSQL
    End If
    If UCase(strModoOperacao) = "IMPRIMIR" Then
'        If MDIMenu.Tag = "ATENDIMENTO" Then
'            'SELECT NA TABELA DE ATENDIMENTO
'            rtf.Text = "RESULTADO DA CONSULTA"
'        ElseIf MDIMenu.Tag = "ATENDIMENTO" Then
'            'SELECT NA TABELA DE ATENDIMENTO
'            rtf.Text = "RESULTADO DA CONSULTA"
'        ElseIf MDIMenu.Tag = "ATENDIMENTO" Then
'            'SELECT NA TABELA DE ATENDIMENTO
'            rtf.Text = "RESULTADO DA CONSULTA"
'        End IF
    End If
End Sub
