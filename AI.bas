Attribute VB_Name = "AI"
Dim continue As Boolean
Sub compTurn()
        
    continue = True
    
    checkblock
    If continue Then
        checkWin
    End If
        
    If continue Then
        findempty
    End If
    
    
End Sub
Sub checkWin()
    
    
    For X = 0 To 2
        If board(X, 0) = board(X, 1) And board(X, 0) = 2 And board(X, 2) = 0 Then
            board(X, 2) = 2
            continue = False
            Exit Sub
        End If
        If board(X, 0) = board(X, 2) And board(X, 0) = 2 And board(X, 1) = 0 Then
            board(X, 1) = 2
            continue = False
            Exit Sub
        End If
        If board(X, 1) = board(X, 2) And board(X, 1) = 2 And board(X, 0) = 0 Then
            board(X, 0) = 2
            continue = False
            Exit Sub
        End If
    Next X
    
    For Y = 0 To 2
        If board(0, Y) = board(1, Y) And board(0, Y) = 2 And board(2, Y) = 0 Then
            board(2, Y) = 2
            continue = False
            Exit Sub
        End If
        If board(0, Y) = board(2, Y) And board(0, Y) = 2 And board(1, Y) = 0 Then
            board(1, Y) = 2
            continue = False
            Exit Sub
        End If
        If board(1, Y) = board(2, Y) And board(1, Y) = 2 And board(0, Y) = 0 Then
            board(0, Y) = 2
            continue = False
            Exit Sub
        End If
        
    Next Y
    
    If board(0, 0) = board(1, 1) And board(0, 0) = 2 And board(2, 2) = 0 Then
        board(2, 2) = 2
        continue = False
        Exit Sub
    End If
    
    If board(0, 0) = board(2, 2) And board(0, 0) = 2 And board(1, 1) = 0 Then
        board(1, 1) = 2
        continue = False
        Exit Sub
    End If
    
    If board(1, 1) = board(2, 2) And board(1, 1) = 2 And board(0, 0) = 0 Then
        board(0, 0) = 2
        continue = False
        Exit Sub
    End If
    
    If board(0, 2) = board(1, 1) And board(0, 2) = 2 And board(2, 0) = 0 Then
        board(2, 0) = 2
        continue = False
        Exit Sub
    End If
    
    If board(0, 2) = board(2, 0) And board(2, 0) = 2 And board(1, 1) = 0 Then
        board(1, 1) = 2
        continue = False
        Exit Sub
    End If
    
    If board(2, 0) = board(1, 1) And board(1, 1) = 2 And board(0, 2) = 0 Then
        board(0, 2) = 2
        continue = False
        Exit Sub
    End If
    
End Sub
Sub checkblock()

    For X = 0 To 2
        If board(X, 0) = board(X, 1) And board(X, 0) = 1 And board(X, 2) = 0 Then
            board(X, 2) = 2
            continue = False
            Exit Sub
        End If
        If board(X, 0) = board(X, 2) And board(X, 0) = 1 And board(X, 1) = 0 Then
            board(X, 1) = 2
            continue = False
            Exit Sub
        End If
        If board(X, 1) = board(X, 2) And board(X, 1) = 1 And board(X, 0) = 0 Then
            board(X, 0) = 2
            continue = False
            Exit Sub
        End If
    Next X
    
    For Y = 0 To 2
        If board(0, Y) = board(1, Y) And board(0, Y) = 1 And board(2, Y) = 0 Then
            board(2, Y) = 2
            continue = False
            Exit Sub
        End If
        If board(0, Y) = board(2, Y) And board(0, Y) = 1 And board(1, Y) = 0 Then
            board(1, Y) = 2
            continue = False
            Exit Sub
        End If
        If board(1, Y) = board(2, Y) And board(1, Y) = 1 And board(0, Y) = 0 Then
            board(0, Y) = 2
            continue = False
            Exit Sub
        End If
    Next Y
    
    
    If board(0, 0) = board(1, 1) And board(0, 0) = 1 And board(2, 2) = 0 Then
        board(2, 2) = 2
        continue = False
        Exit Sub
    End If
    
    If board(0, 0) = board(2, 2) And board(0, 0) = 1 And board(1, 1) = 0 Then
        board(1, 1) = 2
        continue = False
        Exit Sub
    End If
    
    If board(1, 1) = board(2, 2) And board(1, 1) = 1 And board(0, 0) = 0 Then
        board(0, 0) = 2
        continue = False
        Exit Sub
    End If
    
    If board(0, 2) = board(1, 1) And board(0, 2) = 1 And board(2, 0) = 0 Then
        board(2, 0) = 2
        continue = False
        Exit Sub
    End If
    
    If board(0, 2) = board(2, 0) And board(2, 0) = 1 And board(1, 1) = 0 Then
        board(1, 1) = 2
        continue = False
        Exit Sub
    End If
    
    If board(2, 0) = board(1, 1) And board(1, 1) = 1 And board(0, 2) = 0 Then
        board(0, 2) = 2
        continue = False
        Exit Sub
    End If
    
End Sub

Sub findempty()
    
    If board(1, 1) = 0 Then
        board(1, 1) = 2
        Exit Sub
    End If
    
    For X = 0 To 2 Step 2
        If board(0, X) = 0 Then
            board(0, X) = 2
            Exit Sub
        End If
    Next X
    
    For Y = 0 To 2 Step 2
        If board(2, Y) = 0 Then
            board(2, Y) = 2
            Exit Sub
        End If
    Next Y
    
    If board(0, 1) = 0 Then
        board(0, 1) = 2
        Exit Sub
    End If
    
    If board(1, 0) = 0 Then
        board(1, 0) = 2
        Exit Sub
    End If
    
    If board(1, 2) = 0 Then
        board(1, 2) = 2
        Exit Sub
    End If
    
    If board(2, 1) = 0 Then
        board(2, 1) = 2
        Exit Sub
    End If
    
End Sub
