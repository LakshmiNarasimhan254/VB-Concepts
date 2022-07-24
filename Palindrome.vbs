'This is a VB Code to check if the given string is palindrom or Not
Dim StdOut : Set StdOut = CreateObject("Scripting.FileSystemObject").GetStandardStream(1)
Dim strOriginal : strOriginal ="Malayalam" 
Dim strReversed : strReversed = StrReverse(strOriginal)

If UCase(strOriginal) = UCase(strReversed) Then
    StdOut.WriteLine strOriginal & " is a Palindrom"
Else
    StdOut.WriteLine strOriginal & " is not a Palindrom"
End If