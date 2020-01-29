Attribute VB_Name = "nameCompareModule"
'https://www.automateexcel.com/vba/regex/#How_to_Use_Regex_in_VBA

Option Explicit

' Output:
' "exact": names are identical
' "exact (trim)": names are identical only after a trim operation
' "surname deviation": first name matches exactly but sur names are similar
' "first name deviation": first name (and possibly surnames
' "first and surnames deviation": deviation in all names
Private Function nameCompare(ByVal fullName As String, ByVal searchName As String) As String
    Dim regexObj As Object
    Set regexObj = New RegExp
    
    Dim regexPattern As String
    regexPattern = searchName
    
    Dim testStr As String
    testStr = fullName
    
    regexObj.Pattern = regexPattern

    'Test for an exact match
    If regexObj.Test(testStr) Then
        nameCompare = "exact"
        Exit Function
    End If
    
    'Test for an exact match after a trim operation
    regexObj.Pattern = Trim(regexPattern)
    If regexObj.Test(Trim(testStr)) Then
        nameCompare = "exact (trim)"
        Exit Function
    End If
    
    'Test for double whitespaces
    regexObj.Pattern = singleWhitespaces(searchName)
    If regexObj.Test(singleWhitespaces(testStr)) Then
        nameCompare = "exact (white spaces)"
        Exit Function
    End If
    
    'Test for an permissive matching.
    'allows for errors in each name and matches the sur names intelligently
    If nameComparePermissive(fullName, searchName) Then
        nameCompare = "permissive"
        Exit Function
    End If
    
    
    nameCompare = "different"
    
End Function

' insert: \w*
'
' mode: "even"
' in: abcdef   out: a\w*c\w*e\w*
'
' mode: "odd"
' in: abcdef   out: \w*b\w*d\w*f
'
' mode: "odd-first"
' in: abcdef   out: ab\w*d\w*f
Private Function interlace(ByVal str As String, ByVal insert As String, ByVal mode As String) As String
    Dim j As Integer
    Dim replaceChar As Boolean
    interlace = ""
    For j = 1 To Len(str)
        If j Mod 2 = 0 Then
            replaceChar = (mode = "even")
        Else
            replaceChar = (mode = "odd") Or (mode = "odd-first" And j <> 1)
        End If
        
        If replaceChar Then
            interlace = interlace & insert
        Else
            interlace = interlace & Mid(str, j, 1)
        End If
    Next j
End Function

'Input = Gabriela Pedraza Rosa
'(Gabriela|G\w*b\w*i\w*l\w*|Ga\w*r\w*e\w*a)(( Pedraza| P\w*d\w*a\w*a| Pe\w*r\w*z\w*| P(\.)*($| ))|( Rosa| R\w*s\w*| Ro\w*a| R(\.)*($| )))
Private Function nameComparePermissive(ByVal fullName As String, ByVal searchedStr As String) As Boolean
    Dim regex As String
    Dim splited As Variant
    splited = Split(fullName, " ")
    regex = ""
    
    Dim surNamesRegex As String
        surNamesRegex = ""
    Dim firstSurname As Boolean
        firstSurname = True
    
    Dim i As Integer
    For i = LBound(splited) To UBound(splited)
        'MsgBox ("splited(" & i & ") = " & splited(i))

        If i = LBound(splited) Then
            'in: Gabriela
            'out: (Gabriela|G\w*b\w*i\w*l\w*|Ga\w*r\w*e\w*a)
            Dim name As String
            name = splited(i)
            
            'replace every even-indexed char with .*
            Dim interlaced1 As String
            Dim interlaced2 As String

            interlaced1 = interlace(name, "\w*", "even")
            interlaced2 = interlace(name, "\w*", "odd-first")
            regex = regex & "(" & name & "|" & interlaced1 & "|" & interlaced2 & ")"
        Else
            'for every surname
            'in: Pereira
            'out: ( Pereira| P\w*r\w*i\w*a| Pe\w*e\w*r\w*| P(\.)*($| ))
            Dim surNameInterlacedEven As String
            Dim surNameInterlacedOddFirst As String
            Dim firstChar As String
            firstChar = Mid(splited(i), 1, 1)
            Dim firstSurnameChar As String
            If firstSurname Then
                firstSurnameChar = ""
                firstSurname = False
            Else
                firstSurnameChar = "|"
            End If
            surNameInterlacedEven = interlace(splited(i), "\w*", "even")
            surNameInterlacedOddFirst = interlace(splited(i), "\w*", "odd-first")
            surNamesRegex = surNamesRegex & firstSurnameChar & _
                            "( " & splited(i) & _
                            "| " & surNameInterlacedEven & _
                            "| " & surNameInterlacedOddFirst & _
                            "| " & firstChar & "(\.)*($| ))"
        End If
        
    Next i
    regex = regex & "(" & surNamesRegex & ")"

    Dim stringOne As String
    Dim regexOne As Object
    Set regexOne = New RegExp
 
    regexOne.Pattern = regex
 
    stringOne = searchedStr
    
    If regexOne.Test(stringOne) Then
        nameComparePermissive = True
    Else
        nameComparePermissive = False
    End If
End Function


Private Function singleWhitespaces(ByVal s As String) As String
    While InStr(s, "  ")
        s = Replace(s, "  ", " ")
    Wend
    singleWhitespaces = s
End Function
