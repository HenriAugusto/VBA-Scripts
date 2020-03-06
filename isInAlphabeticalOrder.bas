' Special case: If they're equal = return true
Private Function isInAlphabeticOrder(s1 As String, s2 As String) As Boolean

    Dim charIndex, maxLen, smallerString As Integer
    charIndex = 1
    
    Dim keepChecking As Boolean
    keepChecking = s1 <> s2
    
    If Len(s1) = 0 Then
        isInAlphabeticOrder = True
        Exit Function
    End If

    ' check up to the smaller string
    If Len(s1) >= Len(s2) Then
        maxLen = Len(s2) ' LATER WE SHOULD SWAP THIS'
        smallerString = 1
    Else
        maxLen = Len(s1)
        smallerString = 2
    End If
    
    Dim safetyCount As Integer
    safetyCount = 0
    
    
    While keepChecking And safetyCount < maxLen
        ' get chars
        charA = CStr(Left(s1, charIndex))
            charA = CStr(Right(charA, 1))
            charA = stripAccent(charA)
        charB = CStr(Left(s2, charIndex))
            charB = CStr(Right(charB, 1))
            charB = stripAccent(charB)
        'MsgBox ("comparing " & s1 & " with " & s2 & vbLf & "charA = " & charA & vbLf & "charB = " & charB & vbLf & "charIndex = " & charIndex)
        ' If we needs to Swap
        If UCase(charA) > UCase(charB) Then
            isInAlphabeticOrder = False
            keepChecking = False
        ' If we need to check the next char
        ElseIf UCase(charA) < UCase(charB) Then
            isInAlphabeticOrder = True
            keepChecking = False
        ElseIf UCase(charA) = UCase(charB) Then
            charIndex = charIndex + 1
            'If we reached the end of the smallest string but up to that point strings are still equal
            If charIndex > maxLen Then
                'MsgBox ("[isInAlphabeticOrder] ERROR: charIndex > maxLen" & vbLf & _
                '        "charIndex = " & charIndex & vbLf & _
                '        "charA = " & charA & vbLf & "charB = " & charB & vbLf & _
                '        "s1 = " & s1 & vbLf & "s2 = " & s2 & vbLf)
                keepChecking = False
                If smallerString = 1 Then
                    isInAlphabeticOrder = True
                Else
                    isInAlphabeticOrder = False
                End If
            End If
        End If
        safetyCount = safetyCount + 1
        If safetyCount >= maxLen + 1 Then
            MsgBox ("[isInAlphabeticOrder] ERROR: safetyCount aborting!")
        End If
    Wend

End Function
