Attribute VB_Name = "Introduction"

'   INTRODUCTION TO MACROS (CONTINUED)

'       This is the Visual Basic window
'       More specifically, this is a particular module within the visual basic environment
'       A module is simply a collection of code
'       Modules can have a number of macros (or programs) within them and are useful for keeping your code neat
'       To add a module, use Insert -> Module
'       To rename a module, select it in the project explorer, and change it's name property in the properties window
'       If you don't see the properties window, go to View -> Properties window



'   USING COMMENTS
'
'   <-- note, the symbol at the start of this line, "'" , is used to enter a comment
'       anything after this symbol will be ignored by the program (and turned green)
'       use it to annotate your code, so you or others can follow what's going on



'   MACRO STRUCTURE
'
'       Every macro has the following structure:
'
'           Sub nameWithNoSpaces()
'
'               This code executed first...
'               This code executed second...
'               This code executed next...
'               Etc. etc.
'
'           End Sub
'
'       You must name your macro (don't use spaces - you can use camel case like above) and add a set of brackets after the name ()
'       These brackets tell the macro that no parameters are passed to it (there's no inputs needed). Don't worry about this for the time being, just put them in.




'   RUNNING YOUR MACRO
'
'       When you run your macro, every line between your "Sub nameWithNoSpace()" to your "End Sub" is executed in order, one line at a time.
'
'       You can run your macros in a number of ways
'           - you can press the play button in this window (or F5) when the text cursor is within the macro
'           - you can use the Macros button in the developer tab (from your worksheet, not this visual basic environment) and select your macro
'           - you can assign your macro to a button or shape so that when you click it the macro will run (right click the shape and assign macro)
'
'       You can run your macro one line at a time by using the F8 key while your cursor is in the macro.
'       This is incredibly useful for seeing how the code is executed, what values are being used and for debugging your code



'   YOUR FIRST MACRO
'
'       This macro simply presents a message box to the user, with a predefined message.
'       To change the text in the message box, change the text in the code below (between the two quote marks)
'
'       run the macro below to see what happens

        Sub helloWorld()
            
            MsgBox ("Hello World!")
            
        End Sub


'   YOUR SECOND MACRO
'
'       This macro is similer to the helloWorld macro above except this time it asks the user for some input
'       The value that the user inputs is assigned to a variable called "usersName"
'       This value is then used in the MsgBox

        Sub helloWithInputName()
            
            ' Get the users name
            usersName = InputBox("What is your name?")
            
            ' Output the message
            MsgBox ("Hello " & usersName)
            
        
        End Sub


'   INDENT YOUR CODE
'
'       Notice that I have indented my code and comments in the above examples (using the tab button)
'       This makes it easier to follow, particularly when your code becomes more complicated
'       It is good practice to indent your code



'   DECLARE YOUR VARIABLES
'
'       variable names follow the same rules as macros names (no spaces, case sensitive etc.)
'       there are keywords that you can't use in your names
'       For example, you can't have a variable called "If" because the program will think it's the beginning of an If statement
'       If your name throws an error, just use a different name



'   DECLARE YOUR VARIABLES
'
'       It is also good practice (but not always necessary) to declare your variables before you use them
'       By declare, I mean that you should say that a value is either going to be a String (text), an Integer (whole number) or Double (number with a decimal point).
'       There are other types but these are the most common

        Sub declareVariables()
        
            Dim textValue As String
            Dim wholeNumber As Integer
            Dim decimalNumber As Double
            
            Dim answer As Double
                
                
            textValue = "kg"
            
            wholeNumber = 16
            
            ' If you change below to decimalNumber = "text" , you will get an error when you run this macro as "text" is not a Double, it is a String
            decimalNumber = 5.236
            
            answer = wholeNumber / decimalNumber
            
            MsgBox (wholeNumber & textValue & " / " & decimalNumber & " = " & answer & textValue)
        
        
        End Sub


'       So why bother declaring your variables?
'
'       It becomes more important as your code gets more complicated
'       It prevents you from accidentally assigning the wrong data type to a variable
'
'       For example, if I had a variable that stored how many occupants were in a building, I would want it to always be a whole number (an Integer).
'
'       Consider if I was to use this variable later on to determine the occupantDensity (people/m2) as follows:
'
'           occupantDensity = numberOfOccupants/areaOfBuilding
'
'       If the numberOfOccupants had somehow been set to "5 people" (a string), the program would not be able to divide it by another number (you can't divide text by a number)
'       If I had defined numberOfOccupants as an Integer, I could have avoide this error.



'   CONTINUE
'
'       Select the SampleIf module to continue learning about macros
