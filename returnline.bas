Attribute VB_Name = "WilsonMOD"

Public Function ReadLine(Text As String, LineNumber As Integer) As String
Dim TheInt As Integer
Dim TheInt2 As Integer
Dim Lines As Integer
Dim NoNextLine As Boolean
LineNumber = LineNumber - 1

Do
TheInt = TheInt + 1
If Mid(Text, TheInt, 2) = vbCrLf Then Lines = Lines + 1
If Lines = LineNumber Then GoTo 10
If TheInt = Len(Text) Then ReadLine = "": Exit Function
Loop

10 Text = Mid(Text, TheInt + 2)

Do
TheInt2 = TheInt2 + 1
If Mid(Text, TheInt2, 2) = vbCrLf Then GoTo 20
Loop

20
ReadLine = Mid(Text, 1, TheInt2)
End Function

Public Function GetTotalContacts(ListString As String) As Integer
Dim OneLine As String
Dim Spaces As Integer
Dim TheX As Integer
Dim TheX2 As Integer
Dim TheString As String
OneLine = ReadLine(ListString, 1)

Do
TheX = TheX + 1
If Mid(OneLine, TheX, 1) = " " Then Spaces = Spaces + 1
If Spaces = 5 Then GoTo 10
Loop

10 OneLine = Mid(OneLine, TheX + 1)

Do
TheX2 = TheX2 + 1
If Mid(OneLine, TheX2, 1) = " " Then GoTo 20
Loop

20 GetTotalContacts = Mid(OneLine, 1, TheX2)

End Function

Public Function GetContactNumber(ListString As String, TheLine As Integer) As Integer
Dim OneLine As String
Dim Spaces As Integer
Dim TheX As Integer
Dim TheX2 As Integer
Dim TheString As String
OneLine = ReadLine(ListString, TheLine)
If OneLine = "" Then Exit Function

Do
TheX = TheX + 1
If Mid(OneLine, TheX, 1) = " " Then Spaces = Spaces + 1
If Spaces = 4 Then GoTo 10
Loop

10 OneLine = Mid(OneLine, TheX + 1)

Do
TheX2 = TheX2 + 1
If Mid(OneLine, TheX2, 1) = " " Then GoTo 20
Loop

20 GetContactNumber = Mid(OneLine, 1, TheX2)

End Function

'Public Function GetEmail(ListString As String, TheLine As Integer) As Integer
'Dim OneLine As String
'Dim Spaces As Integer
'Dim TheX As Integer
'Dim TheX2 As Integer
'Dim TheString As String
'OneLine = ReadLine(ListString, TheLine)
'If OneLine = "" Then Exit Function
'
'Do
'TheX = TheX + 1
'If Mid(OneLine, TheX, 1) = " " Then Spaces = Spaces + 1
'If Spaces = 6 Then GoTo 10
'Loop
'
'10 OneLine = Mid(OneLine, TheX + 1)
'
'Do
'TheX2 = TheX2 + 1
'If Mid(OneLine, TheX2, 1) = " " Then GoTo 20
'Loop
'
'20 GetContactNumber = Mid(OneLine, 1, TheX2)
'
'End Function


Function RemoveString(Entire As String, Word As String, Replace As String) As String
    Dim I As Integer
    I = 1
    Dim LeftPart
    Do While True
        I = InStr(1, Entire, Word)
        If I = 0 Then
            Exit Do
        Else
            LeftPart = Left(Entire, I - 1)
            Entire = LeftPart & Replace & Right(Entire, Len(Entire) - Len(Word) - Len(LeftPart))
        End If
    Loop
    
   RemoveString = Entire
      
End Function

Public Function GetItem(FullLineStr As String, SpacesIn As Integer, Optional CutOffPoint As String = " ", Optional TillVBCRLF As Boolean = False) As String
Dim FullLine As String
FullLine = FullLineStr

Do
lll = lll + 1
If Mid(FullLine, lll, 1) = " " Then spacesx = spacesx + 1
Loop Until spacesx = SpacesIn

FullLine = Mid(FullLine, lll + 1)

If TillVBCRLF = False Then
Do
kkk = kkk + 1
If Mid(FullLine, kkk, 1) = CutOffPoint Then donex% = 1
Loop Until donex% = 1

GetItem = Mid(FullLine, 1, kkk - 1)

Else

Do
kkk = kkk + 1
If Mid(FullLine, kkk, 2) = vbCrLf Then donex% = 1
Loop Until donex% = 1

End If
GetItem = Mid(FullLine, 1, kkk - 1)




End Function
