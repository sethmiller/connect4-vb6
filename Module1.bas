Attribute VB_Name = "Module1"
Public board(6, 9) As Integer
Public comp As Boolean
Public over As Boolean
Public count As Integer
Public ecks As Integer
Public why As Integer
Public direction As Integer
Public COLOR As String

Function check(tag) As String
    
    'checks horizontal
    For X% = 0 To 6
        For Y% = 0 To 5
            If board(X%, Y%) = board(X%, Y% + 1) And board(X%, Y%) = board(X%, Y% + 2) And board(X%, Y%) = board(X%, Y% + 3) And board(X%, Y%) <> 0 Then
                If ((board(X%, Y%) = 1 And tag = "Blue") Or (board(X%, Y%) = 2 And tag = "Red")) And over = False Then
                    'fill in the board
                    For a = 0 To 6
                        For b = 0 To 9
                            board(a, b) = 1
                        Next b
                    Next a
                    
                    ecks = X%
                    why = Y%
                    direction = 1
                    COLOR = tag
                
                    check = tag + " Wins!!"
                    over = True
                End If
            End If
         Next Y
    Next X
        
    'checks vertical
    For X% = 0 To 9
        For Y% = 0 To 3
            If board(Y%, X%) = board(Y% + 1, X%) And board(Y%, X%) = board(Y% + 2, X%) And board(Y%, X%) = board(Y% + 3, X%) And board(Y%, X%) <> 0 Then
                If ((board(Y%, X%) = 1 And tag = "Blue") Or (board(Y%, X%) = 2 And tag = "Red")) And over = False Then
                    'fill in the board
                    For a = 0 To 6
                        For b = 0 To 9
                            board(a, b) = 1
                        Next b
                    Next a
                    
                    ecks = Y%
                    why = X%
                    direction = 2
                    COLOR = tag
                
                    check = tag + " Wins!!"
                    over = True
                End If
            End If
        Next Y
    Next X
    
    'checks diagonal
    For X = 0 To 3
        For Y = 0 To 5
            If board(X, Y) = board(X + 1, Y + 1) And board(X, Y) = board(X + 2, Y + 2) And board(X, Y) = board(X + 3, Y + 3) And board(X, Y) <> 0 Then
                If ((board(X, Y) = 1 And tag = "Blue") Or (board(X, Y) = 2 And tag = "Red")) And over = False Then
                    'fill in the board
                    For a = 0 To 6
                        For b = 0 To 9
                            board(a, b) = 1
                        Next b
                    Next a
                    
                    ecks = X
                    why = Y
                    direction = 3
                    COLOR = tag
                
                    check = tag + " Wins!!"
                    over = True
                End If
            End If
        Next Y
    Next X
    
    'checks diagonal
    For X = 0 To 3
        For Y = 0 To 5
            If board(X, Y + 3) = board(X + 1, Y + 2) And board(X, Y + 3) = board(X + 2, Y + 1) And board(X, Y + 3) = board(X + 3, Y) And board(X, Y + 3) <> 0 Then
                If ((board(X, Y + 3) = 1 And tag = "Blue") Or (board(X, Y + 3) = 2 And tag = "Red")) And over = False Then
                    'fill in the board
                    For a = 0 To 6
                        For b = 0 To 9
                            board(a, b) = 1
                        Next b
                    Next a
                    
                    ecks = X
                    why = Y
                    direction = 4
                    COLOR = tag
                    
                    check = tag + " Wins!!"
                    over = True
                End If
            End If
        Next Y
    Next X
    
    If over = False Then
        For X = 0 To 6
            For Y = 0 To 8
            If board(X, Y) <> 0 Then
                count = count + 1
            End If
            Next Y
        Next X
    
        If count = 63 Then
            
            check = "Draw!"
            'Score (" ")
        End If
        count = 0
    End If
       
End Function
