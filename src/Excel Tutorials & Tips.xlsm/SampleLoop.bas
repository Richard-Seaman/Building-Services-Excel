Attribute VB_Name = "SampleLoop"

'   SAMPLE LOOP
'
'       If you need to do the same thing a certain number of times, use a Loop.
'       The examples below, demonstrate a "For Loop" and a "Do While Loop"
'       Note: there are other types of loops available which may be more suitable for what you want to do.



'   FOR LOOP
'
'       A For Loop does the same thing for a predefined number of times.
'       The example below will countdown from 5 to 1 and then launch! (in message boxes...)

        Sub sampleFor()
            
            ' define our starting & finishing values
            initialValue = 5
            finalValue = 1
        
            ' This For loop tracks a value "i" and does some code until i = 0
            ' Each time the For loop reaches "Next i", i is changed by -1 (as defined after the Step keyword in the statement)
            For i = initialValue To finalValue Step -1
                
                ' All of the code within this For Loop will be executed each time i changes
                MsgBox ("t-" & i)
            
            Next i
            
            ' This code is only executed after the For Loop has run its course
            MsgBox ("LAUNCH!")
        
        End Sub



'   DO WHILE LOOP
'
'       A Do While Loop will continue to do something as long as a condition is true
'       The example below asks the user to input "Yes master" into a text box
'       If the user enter's the wrong value, they'll be asked again

        Sub sampleDoWhile()
        
            ' Define the text value we want the user to input
            textValueToCheck = "Yes Master"
            
            ' Present an input text box to the user and assign their input to the value "inputtedText"
            inputtedText = InputBox("Type '" & textValueToCheck & "' into the text box below")
            
            ' Keep doing if the value is not equal to (<>) what we want
            Do While inputtedText <> textValueToCheck
            
                inputtedText = InputBox("You must do as I tell you!" & vbNewLine & vbNewLine & "Type '" & textValueToCheck & "' into the text box below")
                
                ' Note: the "vbNewLine" in the code above just skips a line between the text in the inputbox
            
            Loop
            
            ' Congratulate the user for doing what they're told
            MsgBox ("Good human")
            
            ' Notice that if the user inputs the correct text the first time, the do while loop is not executed as the condition is not true
            
            ' Note: You can always use Ctrl + Break on your keyboard to stop a macro that's stuck in a loop (or that's making you call it Master!)
            
        
        End Sub




'   CONTINUE
'
'       Select the UseWorksheet module to continue learning about macros
