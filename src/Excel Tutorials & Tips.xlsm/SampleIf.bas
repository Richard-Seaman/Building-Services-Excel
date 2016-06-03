Attribute VB_Name = "SampleIf"
 
'   SAMPLE IF
'
'       The If statement is probably the most common statement used in code
'
'       Use it if you need to do different things, depending on a given condition
'
'       The macro below asks the user to input two values and then tells the user which of the numbers is biggest
'       It's not very useful but it demonstrates the if statement quite well!


        Sub SampleIf()
         
            ' Grab the first number
            Value1 = InputBox("Enter the first number")
            
            ' Note: The IsNumeric(value) can be used to check if a value is a number, if it is, it will return TRUE
            ' Note: The Not keyword can be used like a NOT logic gate (i.e. if true, return false. if false, return true)
            ' Note: Use Exit Sub to exit the program early
            
            ' Check if the user entered text instead, if they did, exit the macro (finish it early)
            If Not IsNumeric(Value1) Then
                MsgBox ("That wasn't a number, don't try to confuse me!")
                Exit Sub
            End If
            
            
            ' Grab the second number
            Value2 = InputBox("Enter the second number")
            
            ' Check if the user entered text instead, if they did, exit the macro (finish it early)
            If Not IsNumeric(Value2) Then
                MsgBox ("That wasn't a number, don't try to confuse me!")
                Exit Sub
            End If
            
            
            
            ' Check which number is biggest and tell the user
            If Value1 > Value2 Then
                MsgBox ("The first value is the biggest")
            
            ElseIf Value1 < Value2 Then
                MsgBox ("The second value is the biggest")
            
            ElseIf Value1 = Value2 Then
                MsgBox ("The two values are the same")
            
            Else
                MsgBox ("I don't know which number is biggest...")
            
            End If
        
            ' Note: Only the code within the first condition to be true will be executed.
            '       So if Value1 > Value 2 (the first if is true) then the MsgBox("The first value is the biggest") will be executed and then the program will jump to End If
            '       You can see this by using F8 to cycle through each line of code to see what the program is doing.
            
        
        End Sub



'   CONTINUE
'
'       Select the SampleLoop module to continue learning about macros
