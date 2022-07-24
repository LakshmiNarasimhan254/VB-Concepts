
'This is a VB Code to print the fibonacci series 
Dim StdOut : Set StdOut = CreateObject("Scripting.FileSystemObject").GetStandardStream(1)
Dim iFirst, iNext, iMax
iFirst =0
iNext = 1
iMax= 200
While(iNext<iMax)
    If iFirst=0 Then
        Stdout.WriteLine(iFirst)
    End If
    Stdout.WriteLine(iNext)
    iTempNum = iNext
    iNext = iFirst+iNext
    iFirst = iTempNum        
Wend


        
        