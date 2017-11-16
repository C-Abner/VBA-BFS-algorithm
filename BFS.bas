Attribute VB_Name = "Module1"
Public start_x As Integer
Public start_y As Integer
Public end_x As Integer
Public end_y As Integer
Public visited(20, 20) As Integer
Public parentPos(20, 20) As Variant
Public Q As New Collection
Public T As New Collection

Sub btn1_Click()
    Range("A19:W36").Copy Destination:=Range("A1:W18")
    
    For i = 1 To 20
     For j = 1 To 20
        If Cells(i, j).Interior.Color = 255 Then
                end_x = i
                end_y = j
            GoTo Fin
        End If
     Next j
    Next i
Fin:
    start_x = 9
    start_y = 2
    BFS
    PrintAnswer
End Sub
Function PrintAnswer()
    Dim x, y As Integer
    Dim z As Variant
    x = end_x
    y = end_y
    z = parentPos(x, y)
    While z(0) <> 0 And z(1) <> 0
        Cells(z(0), z(1)).Interior.Color = 0
        z = parentPos(z(0), z(1))
    Wend
End Function
Function BFS()
    Q.Add (Array(start_x, start_y))
    T.Add (Array(start_x, start_y))
    
    visited(start_x, start_y) = 1
    parentPos(start_x, start_y) = Array(0, 0)
    
    While Q.Count <> 0
        curPos = Q.Item(1)
        Q.Remove (1)
        
        If isExit(curPos(0), curPos(1)) Then
            Debug.Print "Goal!!"
            Exit Function
        End If
        
        tmp = Array(Array(curPos(0) + 1, curPos(1)), Array(curPos(0) - 1, curPos(1)), Array(curPos(0), curPos(1) + 1), Array(curPos(0), curPos(1) - 1))
        
        For i = 0 To 3
            nextPos = tmp(i)
            
            If notWall(nextPos(0), nextPos(1)) And visited(nextPos(0), nextPos(1)) = 0 Then
                T.Add Array(nextPos(0), nextPos(1))
                visited(nextPos(0), nextPos(1)) = 1
                parentPos(nextPos(0), nextPos(1)) = Array(curPos(0), curPos(1))
                
                Q.Add nextPos
            End If
        Next i
    Wend
    
End Function

Function notWall(x, y)
    If Cells(x, y).Interior.Color <> 16777215 Then
        notWall = 1
    Else
        notWall = 0
    End If
End Function

Function isExit(x, y)
    If x = end_x And y = end_y Then
        isExit = 1
    Else
        isExit = 0
    End If
End Function
