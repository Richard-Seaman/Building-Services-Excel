Attribute VB_Name = "UseWorksheet"

'   USE DATA FROM YOUR WORKSHEET
'
'       Most of the time, we're going to want to use some data from our worksheets in our macros.
'       The example below shows you how to grab a value from a worksheet and assign it to a variable.
'
'
'       To assign data from a particular cell, use the following synthax with the cell's sheet name, row and column numbers:
'
'           variableName = Sheets("WorksheetName").Cells(rowNumber, columnNumber)
'
'       So for a value in row 103 and column 2 on a worksheet called "Macros" , we would assign it to a variable called "cellValue" as follows:
'
'           cellValue = Sheets("Macros").Cells(103, 2)
'
'       Alternatively, if "Macros" is the 11th sheet in the workbook, we can use the use the following synthax:
'
'           cellValue = Sheets(11).Cells(103, 2)

        Sub displayData()
        
            ' Grab the data from the worksheet
            cellValue = Sheets("Macros").Cells(103, 2)
            
            ' Show the data to the user
            MsgBox ("Value in yellow box is" & vbNewLine & vbNewLine & cellValue)
        
        End Sub

'       Try to change the above macro so that if references the other yellow cell on the Macros worksheet
'
'
'       Obviously the above macro isn't very useful but it demonstrates how to use data from a worksheet within your macros




'   CONTINUE
'
'       Select the RecordMacro module to continue learning about macros
