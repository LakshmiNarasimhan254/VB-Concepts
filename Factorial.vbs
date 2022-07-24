'This is a VB Code to print the factorial of a number
Dim StdOut : Set StdOut = CreateObject("Scripting.FileSystemObject").GetStandardStream(1)
Stdout.WriteLine(factorial(5))
Function factorial (iNum)
    If iNum =0 OR iNum = 1 Then
        factorial =1  
        Exit Function     
    Else
        factorial = (iNum)*factorial(iNum-1)        
    End If    
End Function