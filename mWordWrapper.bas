Attribute VB_Name = "mWordWrapper"
Public Function RecoverError(ByVal stpRotina As String) As Boolean
Dim stpMens As String
             
   Screen.MousePointer = vbDefault
 
   stpMens = "Problema na Aplica��o" & _
             Chr$(13) & Chr$(10) & _
             Chr$(13) & Chr$(10) & _
             "Origem : " & Err.Source & " - " & stpRotina & _
             Chr$(13) & Chr$(10) & _
             "C�digo : " & Err.Number & _
             Chr$(13) & Chr$(10) & _
             "Descri��o : " & Err.Description & _
             Chr$(13) & Chr$(10) & _
             Chr$(13) & Chr$(10) & _
             "Deseja repetir a opera��o ?"
 
   Err.Clear
   
   RecoverError = (MsgBox(stpMens, vbQuestion + vbYesNo, App.Title) = vbYes)
   
End Function
