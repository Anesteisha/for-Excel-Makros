Attribute VB_Name = "Module2"
Sub sexStringInTwo()
Attribute sexStringInTwo.VB_Description = "Переделывает 6 строк в 2 столбца.ГЛАВНОЕ: в столбце ""А"" не должно быть пустой ячейки!!! Макрос считает строки до первой пустой клетки в столбце ""А"", потом пишет если после пустой ещё будут строки, он их не увидит и удалит!"
Attribute sexStringInTwo.VB_ProcData.VB_Invoke_Func = "w\n14"

' Brunis, Aloha!

Range("A1").Select
n = 0

Dim b As Boolean
b = True

While (b = True)
 ActiveCell.Offset(1, 0).Select
 If (IsEmpty(ActiveCell.Value) = False) Then
    
    n = n + 1
    
    Else
     
     b = False
     
    End If
    Wend
    
Dim a As Boolean
a = True
s = 0 - n
t = n + 10
ActiveCell.Offset(s, 0).Select
n = n + 2 + t

ActiveCell.Offset(1, 0).Activate
 While (a = True)
 ActiveCell.Offset(0, 0).Range("A2").Select
 If (IsEmpty(ActiveCell.Value) = False) Then
 
    Range("A1,B1,C1,D1,E1,F1,A2,B2,C2,D2,E2,F2").Select
    
    Selection.Copy
    ActiveCell.Offset(n, 0).Activate
    
    Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=True
    ActiveWindow.SmallScroll Down:=12
      
      w = 0 - n
    
    ActiveCell.Offset(w, 0).Activate
    
    Range("A2,B2,C2,D2,E2,F2").Select
    
    Selection.Delete Shift:=xlUp
        
    n = n + 6
    
    Else
     
     a = False
     
    End If
    Wend

    ActiveCell.Offset(-2, 0).Range("A1").Select
    Range("A1,B1,C1,D1,E1,F1,A2,B2,C2,D2,E2,F2").Select
    
    Selection.Copy
    ActiveCell.Offset(n, 0).Activate
    
    Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=True
    ActiveWindow.SmallScroll Down:=12
    
    Range("A1,B1,C1,D1,E1,F1,A2,B2,C2,D2,E2,F2,A3,B3,C3,D3,E3,F3").Select
    
    Selection.Delete Shift:=xlUp
    
    While (t > 0)
    Rows("1:1").Select
   
    Selection.Delete Shift:=xlUp
    t = t - 1
    Wend
    
End Sub
