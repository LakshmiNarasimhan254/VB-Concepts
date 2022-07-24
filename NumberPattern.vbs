'This is a VB Code to find if print the number sequentially in Pyramid format 
Dim StdOut : Set StdOut = CreateObject("Scripting.FileSystemObject").GetStandardStream(1)
Dim iNum 
iNum = 0
For iRow = 0 To 3 Step 1
    For iSpace = 4-iRow to 1 Step -1
        Stdout.Write " "
    Next
    For iCol = 1 to iRow+1 Step 1
        Stdout.Write iNum & " "
        iNum=iNum+1
    Next

    Stdout.WriteLine()
Next
