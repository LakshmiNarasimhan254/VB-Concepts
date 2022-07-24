'This is a VB Code to find if check if the given number is Armstrong or Not 
Option Explicit
Dim StdOut : Set StdOut = CreateObject("Scripting.FileSystemObject").GetStandardStream(1)
Dim iNum,iTempNum,iArmStrongCheck,digit
iNum=153
iTempNum= iNum
iArmStrongCheck = 0

While(iTempNum<>0)
    digit = iTempNum mod 10
    iArmStrongCheck = iArmStrongCheck + (digit*digit*digit)
    iTempNum = iTempNum\10
Wend  
If iArmStrongCheck =iNum Then
    Stdout.WriteLine(iNum & " is an Armstrong number")
Else
    Stdout.WriteLine(iNum & " is not an Armstrong number")
End If


