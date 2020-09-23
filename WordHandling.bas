Attribute VB_Name = "TextHandling"
Option Explicit
'Please keep in mind this program starts at word 0 not 1
Const char_space As String = " "    ' Seperator for words
Const newline = 10                  ' Seperator for lines (also vbnewline)
Public Function findcurrword(phrase As String, char As Long, Optional delimeter As String) As Long
findcurrword = findinstance(phrase, findprevdel(phrase, char, delimeter))
End Function
Public Function findprevdel(phrase As String, char As Long, Optional delimeter As String) As Long
If delimeter = Empty Then delimeter = char_space
Dim spoth As Long
Dim loc As Long
loc = 0
    For spoth = char To 1 Step -1
        If loc = 0 Then
            If LCase(Mid(phrase, spoth, Len(delimeter))) = LCase(delimeter) Then
                loc = spoth
            End If
        End If
    Next
findprevdel = loc
End Function
Public Function findinstance(phrase As String, char As Long) As Long
Dim spoth As Long, temp As Long
temp = 0
If char > Len(phrase) Then char = Len(phrase)
For spoth = 1 To char
    If LCase(Mid(phrase, spoth, 1)) = LCase(Mid(phrase, char, 1)) Then
        temp = temp + 1
    End If
Next
findinstance = temp
End Function
Public Function findchar(ByVal phrase As String, ByVal character As String, ByVal instance As Long) As Long
'locates a specific instance of a character (or string)
Dim count As Long
Dim Found As Boolean
count = 0
Found = False
Dim spoth As Long

If instance > 0 Then
For spoth = 1 To Len(phrase) - (Len(character) - 1)
    If Mid(phrase, spoth, Len(character)) = character Then
        count = count + 1
        If count = instance Then
            'if the instance matches its location is returned
            findchar = spoth
            Found = True
        End If
    End If
Next
Else
    findchar = 1 'makes sure you dont ask for a negative
End If

If Found = False Then
    findchar = Len(phrase) + 1 'there is no space after the last word
End If
End Function
Public Function countwords(phrase As String, Optional delimeter As String) As Byte
If delimeter = Empty Then delimeter = char_space
countwords = countchars(phrase, delimeter)
End Function
Public Function getword(phrase As String, number As Long, Optional delimeter As String) As String
'makes sure you arent asking for a word past the last one
If delimeter = Empty Then delimeter = char_space
If number >= countwords(phrase, delimeter) Then
    number = countwords(phrase, delimeter) - 1
End If
If number = 0 Then
    'for the first word it takes the letters leading to the first char_space
    getword = Mid(phrase, 1, findchar(phrase, delimeter, 1) - 1)
Else
    'for the others it takes the letters starting 1 right of the char_space
    'matching the word number up to the chars to the left of the next char_space
    getword = Mid(phrase, findchar(phrase, delimeter, number) + 1, _
    findchar(phrase, delimeter, number + 1) - findchar(phrase, delimeter, number) - 1)
End If
End Function

Public Function setword(phrase As String, number As Long, word As String, Optional delimeter As String) As String
'replaces the word you want with what you want
If delimeter = Empty Then delimeter = char_space
If number < countwords(phrase, delimeter) Then
Dim newphrase As String
Dim spoth As Long
newphrase = Empty
'adds the words before the one to be set to a buffer
If number > 0 Then
    For spoth = 0 To number - 1
        newphrase = newphrase + getword(phrase, spoth, delimeter) + delimeter
    Next
End If
'adds the word to be set to the buffer
newphrase = newphrase + word + delimeter
'adds the words after the one to be set to the buffer
If number < countwords(phrase, delimeter) - 1 Then
    For spoth = number + 1 To countwords(phrase) - 1
        newphrase = newphrase + getword(phrase, spoth, delimeter) + delimeter
    Next
End If
'removes the final space from the buffer
setword = Mid(newphrase, 1, Len(newphrase) - 1)
Else
'if the word to be set is the after the last one,
'it simply adds it to the phrase
setword = phrase + delimeter + word
End If
End Function

Public Function replaceword(phrase As String, wordin As String, wordout As String, Optional delimeter As String) As String
'searches every word to see if it is wordin
'if it is, its replaced with wordout using setword
If delimeter = Empty Then delimeter = char_space
Dim newphrase As String
Dim count As Long
newphrase = phrase
For count = 0 To countwords(phrase, delimeter) - 1
    If LCase(getword(phrase, count, delimeter)) = LCase(wordin) Then
        newphrase = setword(newphrase, count, wordout, delimeter)
    End If
Next
replaceword = newphrase
End Function

Public Function containsword(phrase As String, word As String) As Boolean
'checks to see if phrase contains word
If Replace(phrase, word, "") <> phrase Then containsword = True Else containsword = False
'Dim count as long
'For count = 0 To countwords(phrase) - 1
'    If LCase(getword(phrase, count, delimeter)) = LCase(word) Then
'        containsword = True
'    End If
'Next
End Function

Public Function countchars(phrase As String, char As String) As Long
'counts words in a phrase by counting the char_space
Dim spoth As Long
Dim count As Long
count = 1   'the first word doesnt have a space before it
For spoth = 1 To Len(phrase) - (Len(char) - 1)
    If Mid(phrase, spoth, Len(char)) = char Then
        count = count + 1
    End If
Next
countchars = count
End Function
Public Function locofword(phrase As String, number As Long, start As Boolean, Optional delimeter As String) As Long
'gives you the start and end location of the word you want
If delimeter = Empty Then delimeter = char_space
If number >= countwords(phrase, delimeter) Then
    number = countwords(phrase, delimeter) - 1
End If
If number = 0 Then
    If start = True Then
        locofword = 1
    Else
        locofword = findchar(phrase, delimeter, 1) - 1
    End If
Else
    If start = True Then
        locofword = findchar(phrase, delimeter, number) + 1
    Else
        locofword = findchar(phrase, delimeter, number + 1) - findchar(phrase, char_space, number) - 1
    End If
End If
End Function

Public Function findnext(phrase As String, start As Long, chars As String) As Long
    Dim temp As Long, hasasked As Boolean
    temp = start + 1
    Do Until containsword(chars, Mid(phrase, temp, 1)) Or temp = Len(phrase)
        temp = temp + 1
        If temp = 100000 And hasasked = False Then
            hasasked = True
            If MsgBox("This task is taking an abnormally long amount of time, would you like to skip it?", vbYesNo + vbQuestion, "Process hanging") = vbYes Then Exit Function
        End If
    Loop
    findnext = temp
End Function
