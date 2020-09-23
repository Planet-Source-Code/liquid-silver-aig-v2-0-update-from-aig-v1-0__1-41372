Attribute VB_Name = "modMain"

Public Const BrainFile As String = "\brain.data"
Public Star As String

Public Function RequestAnswer(strIn As String) As String
    Dim Data As String
    Dim BFile As Integer
    Dim Part1 As String
    Dim Part2 As String
    Dim strInOld As String
    
    strIn = RemoveInvalid(strIn)
    
    strInOld = strIn
    
    strIn = UCase(Trim(strIn))
    BFile = FreeFile
    RequestAnswer = "I'm sorry, I don't know what to say - no match."
    
    Open App.Path + BrainFile For Input As BFile
    
        Do Until EOF(BFile)
            Star = ""
            Part1 = ""
            Part2 = ""
            
            Line Input #BFile, Data
            
            If Data <> "" Then
                Part1 = UCase(Trim(Split(Data, ">")(0)))
                Part2 = Trim(Split(Data, ">")(1))
                
                If Match(strIn, Part1, strInOld) Then
                    RequestAnswer = Answer(Part2)
                    Exit Do
                End If
            End If
        Loop
        
    Close BFile
End Function

Public Function Match(strQ As String, strP As String, strInOld) As Boolean
    Dim OLDsplitP() As String
    Dim Position As Integer
    Dim LastStr As Integer
    Dim splitP() As String
    
    Match = False
    
    OLDsplitP = Split(strP, "|")
    
    If UBound(OLDsplitP) < 2 Then Exit Function
    
    ReDim splitP(0 To UBound(OLDsplitP) - 1) As String
    
    splitP(0) = ""
    
    Position = 0
    
    For i = 1 To UBound(OLDsplitP) - 1
        If OLDsplitP(i) = "*" Then
            If LastStr = 1 Then
                splitP(Position) = splitP(Position) + OLDsplitP(i)
            Else
                Position = Position + 1
                splitP(Position) = OLDsplitP(i)
            End If
            LastStr = 1
        Else
            If LastStr = 2 Then
                splitP(Position) = splitP(Position) + OLDsplitP(i)
            Else
                Position = Position + 1
                splitP(Position) = OLDsplitP(i)
            End If
            LastStr = 2
        End If
    Next i
    
    ReDim Preserve splitP(0 To Position + 1)
    
    If UBound(splitP) = 2 Then
        If splitP(1) = "*" Then
            Match = True: Star = strInOld: Exit Function
        ElseIf splitP(1) = strQ Then
            Match = True: Exit Function
        Else
            Exit Function
        End If
    End If
    
    If UBound(splitP) = 3 Then
        If splitP(1) = "*" Then
            If InStr(Len(strQ) - Len(splitP(2)), strQ, splitP(2)) = 1 Then
                Match = True: Star = Mid(strInOld, 1, Len(strInOld) - Len(splitP(2))): Exit Function
            Else
                Exit Function
            End If
        ElseIf splitP(2) = "*" Then
            If InStr(strQ, splitP(1)) = 1 Then
                Match = True: Star = Right(strInOld, Len(strInOld) - Len(splitP(1))): Exit Function
            Else
                Exit Function
            End If
        Else
            Exit Function
        End If
    End If
    
    If UBound(splitP) = 4 Then
        If splitP(1) = "*" And splitP(3) = "*" Then
            If InStr(strQ, splitP(2)) > 0 Then
                Match = True: Star = Mid(strInOld, 1, Len(strInOld) - Len(splitP(2))) + Right(strInOld, Len(strInOld) - Len(splitP(1))): Exit Function
            Else
                Exit Function
            End If
        End If
    End If
    
    Position = 0
    
    For i = 1 To UBound(splitP) - 1
        If splitP(i) = Mid(strQ, Position + 1, Len(splitP(i))) Then
            Position = Len(splitP(i)) + Position
        Else
            Exit Function
        End If
    Next i
    Match = True
    
End Function

Public Function Answer(strIn) As String
    Dim splitIn() As String
    
    splitIn = Split(strIn, "|")
    
    For i = 1 To UBound(splitIn) - 1
        If splitIn(i) = "*" Then
            Answer = Answer + Star
        Else
            Answer = Answer + splitIn(i)
        End If
    Next i
    
End Function

Public Function RemoveInvalid(strIn As String) As String
    Dim splitIn() As String
    ReDim splitIn(1 To Len(strIn)) As String
    
    If Right(strIn, 1) = "?" Or Right(strIn, 1) = "." Or Right(strIn, 1) = "!" Then
        strIn = Left(strIn, Len(strIn) - 1)
    End If
    
    For i = 1 To Len(strIn)
        splitIn(i) = Mid(strIn, i, 1)
    Next i
    
    For i = 1 To Len(strIn)
        If splitIn(i) = "," Or splitIn(i) = ":" Or splitIn(i) = ";" Then
            RemoveInvalid = RemoveInvalid + ""
        Else
            RemoveInvalid = RemoveInvalid + splitIn(i)
        End If
    Next i
    
End Function
