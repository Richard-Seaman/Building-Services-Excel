Attribute VB_Name = "Sizer"
Sub Clear()

    Dim sheetName As String
    sheetName = "Pipe Section"
        
    ' Clear any old values
    For i = 5 To 50
        
        Sheets(sheetName).Cells(i, 3) = ""
        Sheets(sheetName).Cells(i, 4) = ""
        Sheets(sheetName).Cells(i, 5) = ""
        Sheets(sheetName).Cells(i, 19) = ""
        Sheets(sheetName).Cells(i, 20) = ""
        Sheets(sheetName).Cells(i, 21) = ""
        Sheets(sheetName).Cells(i, 22) = ""
        Sheets(sheetName).Cells(i, 24) = ""
        Sheets(sheetName).Cells(i, 25) = ""
        Sheets(sheetName).Cells(i, 26) = ""
        Sheets(sheetName).Cells(i, 27) = ""
        Sheets(sheetName).Cells(i, 28) = ""
        Sheets(sheetName).Cells(i, 29) = ""
        Sheets(sheetName).Cells(i, 30) = ""
        
    Next i


End Sub

Sub Go()

    Dim sheetName As String
    sheetName = "Pipe Section"
    setupName = "Settings_ReadMe"
    
    ' Column Numbers
    Dim section, length, size, sectionFeeds(4), sectionServes(2), sectionServesTotal(2) As Integer
    
    section = 1
    length = 2
    additionalPdColumn = 3
    additionalMbarColumn = 4
    size = 5
    
    startColumn = 6
    sectionFeeds(0) = startColumn
    sectionFeeds(1) = startColumn + 1
    sectionFeeds(2) = startColumn + 2
    sectionFeeds(3) = startColumn + 3
    sectionFeeds(4) = startColumn + 4
    
    startColumn = 12
    sectionServes(0) = startColumn
    sectionServes(1) = startColumn + 1
    sectionServes(2) = startColumn + 2
    
    startColumn = 15
    sectionServesTotal(0) = startColumn
    sectionServesTotal(1) = startColumn + 1
    sectionServesTotal(2) = startColumn + 2
    
    flowrateColumn = 18
    maxPdPerMeterColumn = 19
    actualPdPerMeterColumn = 20
    actualPdColumn = 21
    velocityColumn = 22
    
    ' Row Numbers
    Dim startRow, endRow As Integer
    startRow = 5
    
    For i = startRow To 1000
    
        If Sheets(sheetName).Cells(i, length) = "" Or Sheets(sheetName).Cells(i, length) = 0 Then
            endRow = i - 1
            Exit For
        End If
    
    Next i
    
    If startRow = endRow + 1 Then
    
        ' No sections
        MsgBox ("No sections. Make sure the first section is entered at section reference 1 and don't skip rows")
        Exit Sub
    
    End If
    
    ' Check for section flowrates
    For i = startRow To endRow
    
        If Sheets(sheetName).Cells(i, flowrateColumn) = 0 Then
            MsgBox ("Empty flowrate entered for section " & Sheets(sheetName).Cells(i, section) & ". Either enter a load (served this section columns) or delete the section (by deleting the length).")
            Exit Sub
        End If
    
    Next i
    
    ' Load in Values
    maxPd = Sheets(setupName).Cells(22, 2)  'Pa
    pressureAtMeter = Sheets(setupName).Cells(18, 2)   'Pa
    allowanceForFittings = Sheets(setupName).Cells(26, 3)   '%
    
    ' Error messages
    If maxPd = "" Then
    
        MsgBox ("Maximum pressure not entered. See Settings tab")
        Exit Sub
    
    End If
    
    If pressureAtMeter = "" Then
    
        MsgBox ("Available pressure at the meter not entered. See Settings tab")
        Exit Sub
    
    End If
    
    If allowanceForFittings = "" Then
    
        MsgBox ("Allowance for fittings not entered. See Settings tab")
        Exit Sub
    
    End If
    
    ' Create arrays that are the correct size
    numberOfSections = endRow - startRow + 1
    Dim sectionLengths(), sectionFlowrates(), sectionAdditionalPd() As Double
    ReDim sectionLengths(0)
    ReDim sectionFlowrates(0)
    ReDim sectionAdditionalPd(0)
    For i = 1 To numberOfSections - 1
        ReDim Preserve sectionLengths(0 To (UBound(sectionLengths) + 1))
        ReDim Preserve sectionFlowrates(0 To (UBound(sectionFlowrates) + 1))
        ReDim Preserve sectionAdditionalPd(0 To (UBound(sectionAdditionalPd) + 1))
    Next i
        
    For i = startRow To endRow
        
        Index = i - startRow
        
        ' Load in the lengths and flowrates
        sectionLengths(Index) = Sheets(sheetName).Cells(i, length)
        sectionFlowrates(Index) = Sheets(sheetName).Cells(i, flowrateColumn)
        
        'Determine if any additional Pd was enetered
        additionalPd = 0
        
        additionalPdEntered = Sheets(sheetName).Cells(i, additionalPdColumn)
        additionalMbarEntered = Sheets(sheetName).Cells(i, additionalMbarColumn)
        
        ' Check if a valid Pd was entered
        If IsNumeric(additionalPdEntered) Then
            additionalPd = additionalPd + additionalPdEntered
        End If
        
        ' Check if a valid mBar was entered, if so convert and add
        If IsNumeric(additionalMbarEntered) Then
            convertedPd = additionalMbarEntered * 100
            additionalPd = additionalPd + convertedPd
        End If
    
        ' Update the array
        sectionAdditionalPd(Index) = additionalPd
    
    Next i
        
    ' Check that all the sections referenced have a section length (this will prevent errors down the way)
    ' And also that there's no circular referencing
    For r = startRow To endRow
    
        currentSection = Sheets(sheetName).Cells(r, section)
    
        ' Loop through each sections's "Feeds Sections"
        For c = sectionFeeds(0) To sectionFeeds(UBound(sectionFeeds))
        
            ' Get the subsection & it's corresponding row
            currentSubSection = Sheets(sheetName).Cells(r, c)
            subSectionFedRow = currentSubSection + startRow - 1
            
            ' If the sub section has no length, throw an error
            If Sheets(sheetName).Cells(subSectionFedRow, length) = "" Then
            
                MsgBox ("Section " & currentSubSection & " used but no data entered for it.")
                Exit Sub
            
            ' If the section feeds itself, throw an error
            ElseIf currentSection = currentSubSection Then
            
                MsgBox ("Section " & currentSection & " feeds itself. Error @ Row:" & r & " Column:" & c)
                Exit Sub
            
            End If
        
        Next c
    
    Next r
    
    ' Check that all the sections (except section 1, so i = 1) in sectionLengths are used in at least one circuit.
    ' If not, there'll be an error later on
    For i = 1 To UBound(sectionLengths)
        
        ' Get the section number
        currentSection = Sheets(sheetName).Cells(i + startRow, section)
        
        ' See if the section is reference anywhere - loop through each section's row
        For r = startRow To endRow
        
            ' Loop through each sections's "Feeds Sections" columns to see if the section is reference anywhere
            For c = sectionFeeds(0) To sectionFeeds(UBound(sectionFeeds))
            
                ' Check if subsection matches section
                currentSubSection = Sheets(sheetName).Cells(r, c)
                
                If currentSubSection = currentSection Then
                
                    ' The section has a reference, skip to check the next section
                    GoTo Cont
                
                End If
            
            Next c
        
        Next r
        
        ' If the all the rows have been checked and there's no reference to the section, throw an error
        MsgBox ("There is a section length entered for an unreferenced section, please don't confuse me (delete the section length, or reference it in another section)")
        Exit Sub
        
Cont:
    Next i
    
    
    ' Determine circuits ----------------------------------------------------------------------------------------------------
    
    ' Each circuit consists of an array of integers, representing the section refs that the circuit takes
    ' Eg. (1,2,3) takes path 1->2->3
    Dim circuits() As Variant
    ReDim circuits(0)
    
    mainSectionRow = 5
     
        sectionMain = mainSectionRow - startRow + 1
    
        'Look for a circuit
        For j = 0 To 4
        
            sectionFed = Sheets(sheetName).Cells(mainSectionRow, sectionFeeds(j))
            
            If sectionFed = "" Then
            
                ' The section doesn't feed any more sections, set the circuit & exit the for
                
                ' Add an index to the array of arrays
                ReDim Preserve circuits(0 To (UBound(circuits) + 1))
                
                ' Create a new array to add
                Dim circuitArray(0) As Integer
                circuitArray(0) = sectionMain
                
                ' Add the array to the array of arrays
                circuits(UBound(circuits)) = circuitArray
                                
                Exit For
                
            Else
            
                sectionFedRow = sectionFed + startRow - 1
                
                'For each subsection that the main section feeds
                
                'Look for a subcircuit
                For k = 0 To 4
                    
                    subSectionFed = Sheets(sheetName).Cells(sectionFedRow, sectionFeeds(k))
                    If subSectionFed = "" Then
            
                        ' The section doesn't feed any more sections, set the circuit & exit the for
                
                        ' Add an index to the array of arrays
                        ReDim Preserve circuits(0 To (UBound(circuits) + 1))
                
                        ' Create a new array to add
                        Dim circuitArray2(1) As Integer
                        circuitArray2(0) = sectionMain
                        circuitArray2(1) = sectionFed
                        
                        ' Add the array to the array of arrays
                        circuits(UBound(circuits)) = circuitArray2
                                        
                        Exit For
                
                    Else
            
                        'For each subsection that the sub section feeds
                        subSectionFedRow = subSectionFed + startRow - 1
                        
                        'For each subsection that the main section feeds
                        'Look for a subcircuit
                        For l = 0 To 4
                        
                            subSectionFed2 = Sheets(sheetName).Cells(subSectionFedRow, sectionFeeds(l))
                
                            If subSectionFed2 = "" Then
                    
                                ' The section doesn't feed any more sections, set the circuit & exit the for
                        
                                ' Add an index to the array of arrays
                                ReDim Preserve circuits(0 To (UBound(circuits) + 1))
                        
                                ' Create a new array to add
                                Dim circuitArray3(2) As Integer
                                circuitArray3(0) = sectionMain
                                circuitArray3(1) = sectionFed
                                circuitArray3(2) = subSectionFed
                                
                                ' Add the array to the array of arrays
                                circuits(UBound(circuits)) = circuitArray3
                                                
                                Exit For
                    
                            Else
                
                                'For each subsection that the sub section feeds
                                subSectionFedRow2 = subSectionFed2 + startRow - 1
                                
                                'For each subsection that the main section feeds
                                'Look for a subcircuit
                                For m = 0 To 4
                                
                                    subSectionFed3 = Sheets(sheetName).Cells(subSectionFedRow2, sectionFeeds(m))
                        
                                    If subSectionFed3 = "" Then
                            
                                        ' The section doesn't feed any more sections, set the circuit & exit the for
                                
                                        ' Add an index to the array of arrays
                                        ReDim Preserve circuits(0 To (UBound(circuits) + 1))
                                
                                        ' Create a new array to add
                                        Dim circuitArray4(3) As Integer
                                        circuitArray4(0) = sectionMain
                                        circuitArray4(1) = sectionFed
                                        circuitArray4(2) = subSectionFed
                                        circuitArray4(3) = subSectionFed2
                                        
                                        ' Add the array to the array of arrays
                                        circuits(UBound(circuits)) = circuitArray4
                                                        
                                        Exit For
                            
                                    Else
                        
                                        'For each subsection that the sub section feeds
                                        subSectionFedRow3 = subSectionFed3 + startRow - 1
                                        
                                        'For each subsection that the main section feeds
                                        'Look for a subcircuit
                                        For n = 0 To 4
                                        
                                            subSectionFed4 = Sheets(sheetName).Cells(subSectionFedRow3, sectionFeeds(n))
                                
                                            If subSectionFed4 = "" Then
                                    
                                                ' The section doesn't feed any more sections, set the circuit & exit the for
                                        
                                                ' Add an index to the array of arrays
                                                ReDim Preserve circuits(0 To (UBound(circuits) + 1))
                                        
                                                ' Create a new array to add
                                                Dim circuitArray5(4) As Integer
                                                circuitArray5(0) = sectionMain
                                                circuitArray5(1) = sectionFed
                                                circuitArray5(2) = subSectionFed
                                                circuitArray5(3) = subSectionFed2
                                                circuitArray5(4) = subSectionFed3
                                                
                                                ' Add the array to the array of arrays
                                                circuits(UBound(circuits)) = circuitArray5
                                                                
                                                Exit For
                                    
                                            Else
                                
                                                'For each subsection that the sub section feeds
                                                subSectionFedRow4 = subSectionFed4 + startRow - 1
                                                
                                                'For each subsection that the main section feeds
                                                'Look for a subcircuit
                                                For o = 0 To 4
                                                
                                                    subSectionFed5 = Sheets(sheetName).Cells(subSectionFedRow4, sectionFeeds(o))
                                        
                                                    If subSectionFed5 = "" Then
                                            
                                                        ' The section doesn't feed any more sections, set the circuit & exit the for
                                                
                                                        ' Add an index to the array of arrays
                                                        ReDim Preserve circuits(0 To (UBound(circuits) + 1))
                                                
                                                        ' Create a new array to add
                                                        Dim circuitArray6(5) As Integer
                                                        circuitArray6(0) = sectionMain
                                                        circuitArray6(1) = sectionFed
                                                        circuitArray6(2) = subSectionFed
                                                        circuitArray6(3) = subSectionFed2
                                                        circuitArray6(4) = subSectionFed3
                                                        circuitArray6(5) = subSectionFed4
                                                        
                                                        ' Add the array to the array of arrays
                                                        circuits(UBound(circuits)) = circuitArray6
                                                                        
                                                        Exit For
                                            
                                                    Else
                                        
                                                        'For each subsection that the subsection feeds
                                                        subSectionFedRow5 = subSectionFed5 + startRow - 1
                                                    
                                                        'For each subsection that the main section feeds
                                                        'Look for a subcircuit
                                                        For p = 0 To 4
                                                        
                                                            subSectionFed6 = Sheets(sheetName).Cells(subSectionFedRow5, sectionFeeds(p))
                                                
                                                            If subSectionFed6 = "" Then
                                                    
                                                                ' The section doesn't feed any more sections, set the circuit & exit the for
                                                        
                                                                ' Add an index to the array of arrays
                                                                ReDim Preserve circuits(0 To (UBound(circuits) + 1))
                                                        
                                                                ' Create a new array to add
                                                                Dim circuitArray7(6) As Integer
                                                                circuitArray7(0) = sectionMain
                                                                circuitArray7(1) = sectionFed
                                                                circuitArray7(2) = subSectionFed
                                                                circuitArray7(3) = subSectionFed2
                                                                circuitArray7(4) = subSectionFed3
                                                                circuitArray7(5) = subSectionFed4
                                                                circuitArray7(6) = subSectionFed5
                                                                
                                                                ' Add the array to the array of arrays
                                                                circuits(UBound(circuits)) = circuitArray7
                                                                                
                                                                Exit For
                                                    
                                                            Else
                                                
                                                                'For each subsection that the subsection feeds
                                                                subSectionFedRow6 = subSectionFed6 + startRow - 1
                                                            
                                                                'For each subsection that the main section feeds
                                                                'Look for a subcircuit
                                                                For q = 0 To 4
                                                                
                                                                    subSectionFed7 = Sheets(sheetName).Cells(subSectionFedRow6, sectionFeeds(q))
                                                        
                                                                    If subSectionFed7 = "" Then
                                                            
                                                                        ' The section doesn't feed any more sections, set the circuit & exit the for
                                                                
                                                                        ' Add an index to the array of arrays
                                                                        ReDim Preserve circuits(0 To (UBound(circuits) + 1))
                                                                
                                                                        ' Create a new array to add
                                                                        Dim circuitArray8(7) As Integer
                                                                        circuitArray8(0) = sectionMain
                                                                        circuitArray8(1) = sectionFed
                                                                        circuitArray8(2) = subSectionFed
                                                                        circuitArray8(3) = subSectionFed2
                                                                        circuitArray8(4) = subSectionFed3
                                                                        circuitArray8(5) = subSectionFed4
                                                                        circuitArray8(6) = subSectionFed5
                                                                        circuitArray8(7) = subSectionFed6
                                                                        
                                                                        ' Add the array to the array of arrays
                                                                        circuits(UBound(circuits)) = circuitArray8
                                                                                        
                                                                        Exit For
                                                            
                                                                    Else
                                                        
                                                                        'For each subsection that the subsection feeds
                                                                        subSectionFedRow7 = subSectionFed7 + startRow - 1
                                                                    
                                                                        'For each subsection that the main section feeds
                                                                        'Look for a subcircuit
                                                                        For r = 0 To 4
                                                                        
                                                                            subSectionFed8 = Sheets(sheetName).Cells(subSectionFedRow7, sectionFeeds(r))
                                                                
                                                                            If subSectionFed8 = "" Then
                                                                    
                                                                                ' The section doesn't feed any more sections, set the circuit & exit the for
                                                                        
                                                                                ' Add an index to the array of arrays
                                                                                ReDim Preserve circuits(0 To (UBound(circuits) + 1))
                                                                        
                                                                                ' Create a new array to add
                                                                                Dim circuitArray9(8) As Integer
                                                                                circuitArray9(0) = sectionMain
                                                                                circuitArray9(1) = sectionFed
                                                                                circuitArray9(2) = subSectionFed
                                                                                circuitArray9(3) = subSectionFed2
                                                                                circuitArray9(4) = subSectionFed3
                                                                                circuitArray9(5) = subSectionFed4
                                                                                circuitArray9(6) = subSectionFed5
                                                                                circuitArray9(7) = subSectionFed6
                                                                                circuitArray9(8) = subSectionFed7
                                                                                
                                                                                ' Add the array to the array of arrays
                                                                                circuits(UBound(circuits)) = circuitArray9
                                                                                                
                                                                                Exit For
                                                                    
                                                                            Else
                                                                
                                                                                'For each subsection that the subsection feeds
                                                                                subSectionFedRow8 = subSectionFed8 + startRow - 1
                                                                            
                                                                                'For each subsection that the main section feeds
                                                                                'Look for a subcircuit
                                                                                For s = 0 To 4
                                                                                
                                                                                    subSectionFed9 = Sheets(sheetName).Cells(subSectionFedRow8, sectionFeeds(s))
                                                                        
                                                                                    If subSectionFed9 = "" Then
                                                                            
                                                                                        ' The section doesn't feed any more sections, set the circuit & exit the for
                                                                                
                                                                                        ' Add an index to the array of arrays
                                                                                        ReDim Preserve circuits(0 To (UBound(circuits) + 1))
                                                                                
                                                                                        ' Create a new array to add
                                                                                        Dim circuitArray10(9) As Integer
                                                                                        circuitArray10(0) = sectionMain
                                                                                        circuitArray10(1) = sectionFed
                                                                                        circuitArray10(2) = subSectionFed
                                                                                        circuitArray10(3) = subSectionFed2
                                                                                        circuitArray10(4) = subSectionFed3
                                                                                        circuitArray10(5) = subSectionFed4
                                                                                        circuitArray10(6) = subSectionFed5
                                                                                        circuitArray10(7) = subSectionFed6
                                                                                        circuitArray10(8) = subSectionFed7
                                                                                        circuitArray10(9) = subSectionFed8
                                                                                        
                                                                                        ' Add the array to the array of arrays
                                                                                        circuits(UBound(circuits)) = circuitArray10
                                                                                                        
                                                                                        Exit For
                                                                            
                                                                                    Else
                                                                        
                                                                                        'For each subsection that the subsection feeds
                                                                                        MsgBox ("Maximum 10 sections per circuit. If you need more sections talk to Richard Seaman and he'll fix the code. If not, delete some sections and try again.")
                                                                                        End
                                                                            
                                                                            
                                                                                    End If
                                                                        
                                                                                Next s
                                                                                
                                                                            End If
                                                                
                                                                        Next r
                                                                        
                                                                    End If
                                                        
                                                                Next q
                                                    
                                                            End If
                                                
                                                        Next p
                                        
                                                    End If
                                    
                                                Next o
                                                
                                            End If
                            
                                        Next n
                                        
                                    End If
                    
                                Next m
                                
                            End If
            
                        Next l
                
                    End If
        
                Next k
               
            End If
        
        Next j
        
    ' Circuits Have been determined at this point ----------------------------------------------------------------------------
    
    ' Determine the lengths of each circuit (index run = longest length) -----------------------------------------------------
    
    ' circuitsLengths[0] correspond to the length of circuits[0]
    Dim circuitLengths(), circuitLengthsEquiv(), circuitAdditionalPd() As Double
    ReDim circuitLengths(0)
    ReDim circuitLengthsEquiv(0)
    ReDim circuitAdditionalPd(0)
    
    Dim maxCircuitPdPerMeter() As Double
    ReDim maxCircuitPdPerMeter(0)
    
    ' Sum up the lengths and additional Pd's
    For i = 1 To UBound(circuits)
    
        numberOfSectionsInCircuit = UBound(circuits(i)) + 1
        
        ' Reset length & additionalPd
        length = 0
        additionalPd = 0
        
        ' Sum all the lengths in the circuit
        For j = 0 To numberOfSectionsInCircuit - 1
        
            sectionRef = circuits(i)(j)
            length = length + sectionLengths(sectionRef - 1)
            additionalPd = additionalPd + sectionAdditionalPd(sectionRef - 1)
        
        Next j
        
        ' Update the arrays
        ' Add an index to the array of arrays
        
        ReDim Preserve circuitLengths(0 To (UBound(circuitLengths) + 1))
        circuitLengths(UBound(circuitLengths)) = length
        
        ReDim Preserve circuitLengthsEquiv(0 To (UBound(circuitLengthsEquiv) + 1))
        circuitLengthsEquiv(UBound(circuitLengthsEquiv)) = length * (1 + allowanceForFittings)
    
        ReDim Preserve circuitAdditionalPd(0 To (UBound(circuitAdditionalPd) + 1))
        circuitAdditionalPd(UBound(circuitAdditionalPd)) = additionalPd
        
    Next i
    
    ' Determine the max Pd/m for each circuit
    For i = 1 To UBound(circuits)
           
        ReDim Preserve maxCircuitPdPerMeter(0 To (UBound(maxCircuitPdPerMeter) + 1))
        
        If maxPd - circuitAdditionalPd(i) < 0 Then
        
                MsgBox ("There is too much additional pressure drop. Pd exceeds 10% of that at meter. Can't continue...")
                Exit Sub
        
        End If
        
        maxCircuitPdPerMeter(UBound(maxCircuitPdPerMeter)) = (maxPd - circuitAdditionalPd(i)) / circuitLengthsEquiv(i)
    
    Next i
    
    
    
    ' Determine the longest run
    indexLength = 0
    For i = 0 To UBound(circuitLengths)
    
        If circuitLengths(i) > indexLength Then
            indexLength = circuitLengths(i)
        End If
    
    Next i
    
    ' Add a section max Pd/m array
    ' pipeSizes[0] correspond to the sectionFlowrates[0]
    Dim maxSectionPdPerMeter() As Double
    ReDim maxSectionPdPerMeter(0)
    For i = 1 To numberOfSections - 1
        ReDim Preserve maxSectionPdPerMeter(0 To (UBound(maxSectionPdPerMeter) + 1))
    Next i
    
    ' Determine the max pressure drop accross each section (some sections are in multiple circuits)
    ' for each section
    For i = 0 To UBound(sectionLengths)
    
        maxPdPerMeterForSection = 1000
        
        ' Loop through each circuit (circuit(0) is empty)
        For j = 1 To UBound(circuits)
        
            ' Loop through each section in circuit
            For k = 0 To UBound(circuits(j))
                
                ' If the section is in the circuit
                currentSectionInCircuit = circuits(j)(k)
                If currentSectionInCircuit = i + 1 Then
                
                    ' Get the max pd/m for the circuit
                    maxPdPerMeterForThisCircuit = maxCircuitPdPerMeter(j)
                    
                    ' If it's lower than the section's current value, replace it
                    If maxPdPerMeterForThisCircuit < maxPdPerMeterForSection Then
                    
                        maxPdPerMeterForSection = maxPdPerMeterForThisCircuit
                    
                    End If
                
                End If
                
            Next k
        
        Next j
        
        'Update the array
        maxSectionPdPerMeter(i) = maxPdPerMeterForSection
    
    Next i
    
    ' Circuit lengths and length of index run knonw at this point ----------------------------------------------------------
    
    ' Determine pipe sizes & pressure drops --------------------------------------------------------------------------------
    'maxPdPerMeter = maxPd / indexLength     'Pa/m
    
    ' Add a pipe size array
    ' pipeSizes[0] correspond to the sectionFlowrates[0]
    Dim pipeSizes() As Integer
    ReDim pipeSizes(0)
    For i = 1 To numberOfSections - 1
        ReDim Preserve pipeSizes(0 To (UBound(pipeSizes) + 1))
    Next i
    
    Dim pressureDropsPerMeter() As Double
    ReDim pressureDropsPerMeter(0)
    For i = 1 To numberOfSections - 1
        ReDim Preserve pressureDropsPerMeter(0 To (UBound(pressureDropsPerMeter) + 1))
    Next i
    
    Dim pressureDrops() As Double
    ReDim pressureDrops(0)
    For i = 1 To numberOfSections - 1
        ReDim Preserve pressureDrops(0 To (UBound(pressureDrops) + 1))
    Next i

    Dim sectionVelocities() As Double
    ReDim sectionVelocities(0)
    For i = 1 To numberOfSections - 1
        ReDim Preserve sectionVelocities(0 To (UBound(sectionVelocities) + 1))
    Next i
    
    ' Set up the table data
    tableSheetName = "CIBSE Table"
    firstRow = 5
    lastRow = 34
    firstColumn = 2
    lastColumn = 14
    
    ' Pa/m array
    Dim dpPerMeter() As Double
    ReDim dpPerMeter(0)
    For i = firstRow To lastRow - 1
        ReDim Preserve dpPerMeter(0 To (UBound(dpPerMeter) + 1))
    Next i
    
    ' Table pipe sizes array
    Dim availablePipeSizes() As Integer
    ReDim availablePipeSizes(0)
    For i = firstColumn To lastColumn - 1
        ReDim Preserve availablePipeSizes(0 To (UBound(availablePipeSizes) + 1))
    Next i
    
    For i = 0 To UBound(availablePipeSizes)
        availablePipeSizes(i) = Sheets(tableSheetName).Cells(2, firstColumn + i)
    Next i
    
    
    For section = 0 To UBound(sectionLengths)
    
        maxPdPerMeter = maxSectionPdPerMeter(section)
            
        targetRow = 0
                
        For i = 0 To UBound(dpPerMeter)
        
            nextDp = Sheets(tableSheetName).Cells(firstRow + i + 1, 1)
            thisDp = Sheets(tableSheetName).Cells(firstRow + i, 1)
            
            dpPerMeter(i) = thisDp
            
            currentRow = i + firstRow
            
            If (maxPdPerMeter < nextDp And maxPdPerMeter >= thisDp) Then
                targetRow = currentRow
                Exit For
            End If
            
        Next i
        
        If targetRow = 0 Then
        MsgBox ("Maximum allowable pressure drop can't be achieved. Either reduce the circuit length/load or increase the available pressure.")
        End
        End If
        
        i = section
           
        ' Pipe Size
        flowrate = sectionFlowrates(i)
        targetColumn = 0
                
        For j = firstColumn To lastColumn
            
            currentFlowrate = Sheets(tableSheetName).Cells(targetRow, j)
            If flowrate < currentFlowrate Then
            
                targetColumn = j
                Exit For
            
            End If
    
        Next j
        
        If targetColumn = 0 Then
        MsgBox ("Pipe size is over 150 and I can currently only handle up to 150 dia.  pleaes give out to Richard S")
        End
        End If
        
        pipeSize = Sheets(tableSheetName).Cells(2, targetColumn)
        pipeSizes(i) = pipeSize
        
        targetRowForPd = 0
        
        ' Pressure drop
        For x = firstRow To lastRow
            
            Dim thisFlowrate, nextFlowrate As Double
            
            thisFlowrate = Sheets(tableSheetName).Cells(x, targetColumn)
            nextFlowrate = Sheets(tableSheetName).Cells(x + 1, targetColumn)
            
            If targetRow = startRow And thisFlowrate >= flowrate Then
                
                ' If the first row is the right row to use
                targetRowForPd = targetRow
                Exit For
                
            ElseIf x = firstRow And thisFlowrate > flowrate Then
                
                ' If the flowrate is less than the minimum flowrate for this size
                targetRowForPd = x
                Exit For
            
            ElseIf thisFlowrate <= flowrate And nextFlowrate > flowrate Then
            
                targetRowForPd = x
                Exit For
                
            End If
            
        Next x
                
        higherFlow = Sheets(tableSheetName).Cells(targetRowForPd + 1, targetColumn)
        lowerFlow = Sheets(tableSheetName).Cells(targetRowForPd, targetColumn)
        
        higherDp = Sheets(tableSheetName).Cells(targetRowForPd + 1, 1)
        lowerDp = Sheets(tableSheetName).Cells(targetRowForPd, 1)
        
        actualDpPerMeter = ((flowrate - lowerFlow) / (higherFlow - lowerFlow)) * (higherDp - lowerDp) + lowerDp
        
        ' The additional Pd must be included. This is okay as the pipe has already been sized.
        ' The actual pd/m does not change as this is based on the flowrate, not the pd from fittings etc.
        pressureDropsPerMeter(i) = actualDpPerMeter
        pressureDrops(i) = pressureDropsPerMeter(i) * sectionLengths(i) + sectionAdditionalPd(i)
        
        ' Determine section velocity
        sectionVelocities(i) = sectionFlowrates(i) * 4 / (3.141592654 * (pipeSizes(i) / 1000) * (pipeSizes(i) / 1000))
        
        ' Update
        Sheets(sheetName).Cells(startRow + i, size) = pipeSizes(i)
        Sheets(sheetName).Cells(startRow + i, maxPdPerMeterColumn) = maxSectionPdPerMeter(i)
        Sheets(sheetName).Cells(startRow + i, actualPdPerMeterColumn) = pressureDropsPerMeter(i)
        Sheets(sheetName).Cells(startRow + i, actualPdColumn) = pressureDrops(i)
        Sheets(sheetName).Cells(startRow + i, velocityColumn) = sectionVelocities(i)
        
    Next section
    
    ' Determine actual pressure drop accross circuits & create circuit String array
    ' create the array
    Dim actualCircuitPds(), actualCircuitPdPermeter() As Double
    Dim circuitStrings() As String
    ReDim actualCircuitPds(0)
    ReDim actualCircuitPdPermeter(0)
    ReDim circuitStrings(0)
    For i = 1 To UBound(circuits)
        ReDim Preserve actualCircuitPds(0 To (UBound(actualCircuitPds) + 1))
        ReDim Preserve actualCircuitPdPermeter(0 To (UBound(actualCircuitPdPermeter) + 1))
        ReDim Preserve circuitStrings(0 To (UBound(circuitStrings) + 1))
    Next i
        
    Dim residualCircuitPressures(), residualCircuitPressuresMbar(), percentagePressureDropForCircuit() As Double
    ReDim residualCircuitPressures(0)
    ReDim residualCircuitPressuresMbar(0)
    ReDim percentagePressureDropForCircuit(0)
    For i = 1 To UBound(circuits)
        ReDim Preserve residualCircuitPressures(0 To (UBound(residualCircuitPressures) + 1))
        ReDim Preserve residualCircuitPressuresMbar(0 To (UBound(residualCircuitPressuresMbar) + 1))
        ReDim Preserve percentagePressureDropForCircuit(0 To (UBound(percentagePressureDropForCircuit) + 1))
    Next i
    
    ' loop through each circuit
    For circuit = 1 To UBound(circuits)
        
        ' Reset the variable
        actualPd = 0
        
        ' Create the string
        circuitString = "Circuit: "
        
        ' Loop through each section in circuit
        For section = 0 To UBound(circuits(circuit))
            
            ' If the section is in the circuit
            currentSectionInCircuit = circuits(circuit)(section)
            
            For i = 0 To UBound(pressureDrops)
            
                If currentSectionInCircuit = i + 1 Then
                
                    ' Sum up the actual Pds accross the sections
                    actualPd = actualPd + pressureDrops(i)
                
                    ' Add the section to the string
                    circuitString = circuitString & circuits(circuit)(section) & " "
                    
                End If
                
            Next i
            
        Next section
        
        ' Update the arrays
        actualCircuitPds(circuit) = actualPd
        circuitStrings(circuit) = circuitString
        actualCircuitPdPermeter(circuit) = actualCircuitPds(circuit) / circuitLengths(circuit)
        
        ' Determine residual pressure available
        residualCircuitPressures(circuit) = (pressureAtMeter - actualCircuitPds(circuit)) / 1000                            'kPa
        residualCircuitPressuresMbar(circuit) = residualCircuitPressures(circuit) * 10                                      'mBar
        percentagePressureDropForCircuit(circuit) = 1 - (residualCircuitPressuresMbar(circuit) / (pressureAtMeter / 100))   '%
        
    
    Next circuit
    
    'Print the circuits & pressure drops
    For i = 0 To UBound(circuitStrings)
        
        Sheets(sheetName).Cells(i + startRow, 24) = circuitStrings(i)
        Sheets(sheetName).Cells(i + startRow, 25) = circuitLengths(i)
        Sheets(sheetName).Cells(i + startRow, 26) = actualCircuitPdPermeter(i)
        Sheets(sheetName).Cells(i + startRow, 27) = actualCircuitPds(i)
        Sheets(sheetName).Cells(i + startRow, 28) = residualCircuitPressures(i)
        Sheets(sheetName).Cells(i + startRow, 29) = residualCircuitPressuresMbar(i)
        Sheets(sheetName).Cells(i + startRow, 30) = percentagePressureDropForCircuit(i)
        
        
    Next i
    

End Sub
