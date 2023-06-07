Attribute VB_Name = "mWordWrapper"
Public Function RecoverError(ByVal stpRotina As String) As Boolean
Dim stpMens As String
             
   Screen.MousePointer = vbDefault
 
   stpMens = "Problema na Aplicação" & _
             Chr$(13) & Chr$(10) & _
             Chr$(13) & Chr$(10) & _
             "Origem : " & Err.Source & " - " & stpRotina & _
             Chr$(13) & Chr$(10) & _
             "Código : " & Err.Number & _
             Chr$(13) & Chr$(10) & _
             "Descrição : " & Err.Description & _
             Chr$(13) & Chr$(10) & _
             Chr$(13) & Chr$(10) & _
             "Deseja repetir a operação ?"
 
   Err.Clear
   
   RecoverError = (MsgBox(stpMens, vbQuestion + vbYesNo, App.Title) = vbYes)
   
End Function
