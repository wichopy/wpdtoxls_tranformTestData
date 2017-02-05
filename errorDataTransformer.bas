'Script for automated error position input to access database.
'created by William Chou Mar 5 2012

'The script works be converting the test report table into paragraphs.
'Each paragraph will be an entry into an array where all the test values will be held.
'The array will be sorted and formatted, then output into excel.

'Debugging will be in the section of code where the temp array transfers data to the final array.
'I had 2 cases where either the data was 1 spot before or after the standard start position.
'IF you run into error you would want to modify the code highlighted below.

'ENJOY :D

Sub GetFromWord()

Dim appWD As Word.Application
Dim appDoc As Word.Document
Dim trange As Word.range
Dim temp As Variant
Dim tempstring1 As String
Dim tempstring2 As String
Dim final As Variant

Dim Table As Word.Table
Dim storage(0 To 23) As Variant

Dim p As Long, r As Long, c As Long, size As Long, xlrow As Long, trow As Long 'long as there may be a lot of entries.
Dim x As Integer, y As Integer
'________________________________________________________________________

    Sheets.Add ' create a new spreadsheet
    'Put in Headings.
    With range("A1")
        .Formula = "Copy the below table values and 'Paste Append' to the database"
        .Font.Bold = True
        .Font.size = 14
    End With
    
    With range("A2")
        .Formula = "Error ID"
    End With
    
    With range("B2")
        .Formula = "Standard"
    End With
        
    With range("C2")
        .Formula = "Response"
    End With
    
    With range("D2")
        .Formula = "Measurement"
    End With
    
    With range("E2")
        .Formula = "Voltage Setpoint"
    End With
    
    With range("F2")
        .Formula = "Current Setpoint"
    End With
    
    With range("G2")
        .Formula = "Phase Setpoint"
    End With
    
    With range("H2")
        .Formula = "Load"
    End With

    With range("I2")
        .Formula = "Element"
    End With
    
    With range("J2")
        .Formula = "Isolation"
    End With
    With range("K2")
        .Formula = "VoltageSource"
    End With
    
    With range("L2")
        .Formula = "Reference Error"
    End With
    
    With range("M2")
        .Formula = "Station 1 Correction"
    End With
    
    With range("N2")
        .Formula = "Station 2 Correction"
    End With
    
    With range("O2")
        .Formula = "Station 3 Correction"
    End With
    
    With range("P2")
        .Formula = "Station 4 Correction"
    End With
    
    With range("Q2")
        .Formula = "Station 5 Correction"
    End With
    
    With range("R2")
        .Formula = "Station 6 Correction"
    End With
    
    With range("S2")
        .Formula = "Station 7 Correction"
    End With
    
    With range("T2")
        .Formula = "Station 8 Correction"
    End With
    
    With range("U2")
        .Formula = "Station 9 Correction"
    End With
    
    With range("V2")
        .Formula = "Station 10 Correction"
    End With
    
    With range("W2")
        .Formula = "Station 11 Correction"
    End With
    
    With range("X2")
        .Formula = "Station 12 Correction"
    End With
    
'_________________________________________________________________________________________________
    xlrow = 3    'start row for pasting data in excel
    trow = 0    'table position in document
    
    Set appWD = CreateObject("Word.Application")
    appWD.Visible = False
    Set appDoc = appWD.Documents.Open(Sheet3.TextBox1.Value) 'check your filename and location
    ActiveSheet.Name = Sheet3.TextBox2
    Set Table = appDoc.Tables.Item(1)
    Set trange = Table.ConvertToText(Separator:=vbTab) 'convert table to text
    temp = Split(trange.Text, vbCr) 'each paragraph marker will indicate a new element in array
    appDoc.Undo    'undo convert to text
    size = UBound(temp)
    ReDim final(size) 'define the size of the array
    
' DEBUG THIS SECTION IF ERRORS ARISE************************************
  '*********************************************************************
  'Use watches and examine temp, final, and storage arrays.
    x = 0 ' flags
    y = 0
    For r = 0 To size 'create the array which holds all table values.

        If r = size Then 'used when values are not aligned by 1 position to the right.
            x = 0
        End If

        final(r) = Split(temp(r + x - y), vbTab) 'when values are not aligned by 1 position to the left.
        If r = 49 And temp(49) = "" Then
            y = 1
        End If
        
        If r = 50 And temp(51) = "" Then 'when values are  not aligned by 1 position to the right.
            x = 1
        End If
    Next
'******************************************************************************
'****************************************************************************

'_______________________________________________________________________________________________________
    'Constants in every row
    storage(0) = ""
    storage(2) = ""
    storage(9) = "Off"
    storage(10) = "Parallel"
    storage(11) = "'0.00"
    storage(22) = "'0.00"
    storage(23) = "'0.00"

 '_____________________________________________________________________________________________
    'define row values from string array

For p = 1 To (size - 50) / 10
    
    storage(4) = final(51 + trow) 'voltage
    storage(4)(0) = Format(storage(4)(0), "###.0") 'formatting for  decimal places
    storage(5) = final(52 + trow) 'current
    storage(5)(0) = Format(storage(5)(0), "0.00")
    storage(6) = final(53 + trow) 'phase angle
    storage(6)(0) = Format(storage(6)(0), "0.0")
    storage(7) = final(54 + trow) 'load
    storage(8) = ":)" ' to be changed every row
    
    For x = 1 To 10
        final(59 + trow)(0) = Replace(final(59 + trow)(0), "", "-") 'correct for negative signs.
        storage(11 + x) = final(59 + trow) 'duplicate errors for 10 stations
    Next x
    
    y = 1
    'output array
    For x = xlrow To xlrow + 3
        range("A" & x).Formula = storage(0)
        range("B" & x).Formula = storage(1)
        range("C" & x).Formula = storage(2)
        range("D" & x).Formula = storage(3)
        range("E" & x).Formula = storage(4)
        range("F" & x).Formula = storage(5)
        range("G" & x).Formula = storage(6)
        range("H" & x).Formula = storage(7)
       
        If y = 1 Then 'output all 4 elements for each test ( Single, Left, Middle, or Right)
            range("I" & x).Formula = "S"
        End If
        If y = 2 Then
            range("I" & x).Formula = "L"
        End If
        If y = 3 Then
            range("I" & x).Formula = "M"
        End If
        If y = 4 Then
            range("I" & x).Formula = "R"
        End If
        range("J" & x).Formula = storage(9)
        range("K" & x).Formula = storage(10)
        range("L" & x).Formula = storage(11)
        storage(12)(0) = Format(storage(12)(0), "0.00")
        range("M" & x).Formula = storage(12)
        storage(13)(0) = Format(storage(13)(0), "0.00")
        range("N" & x).Formula = storage(13)
        storage(14)(0) = Format(storage(14)(0), "0.00")
        range("O" & x).Formula = storage(14)
        storage(15)(0) = Format(storage(15)(0), "0.00")
        range("P" & x).Formula = storage(15)
        storage(16)(0) = Format(storage(16)(0), "0.00")
        range("Q" & x).Formula = storage(16)
        storage(17)(0) = Format(storage(17)(0), "0.00")
        range("R" & x).Formula = storage(17)
        storage(18)(0) = Format(storage(18)(0), "0.00")
        range("S" & x).Formula = storage(18)
        storage(19)(0) = Format(storage(19)(0), "0.00")
        range("T" & x).Formula = storage(19)
        storage(20)(0) = Format(storage(20)(0), "0.00")
        range("U" & x).Formula = storage(20)
        storage(21)(0) = Format(storage(21)(0), "0.00")
        range("V" & x).Formula = storage(21)
        range("W" & x).Formula = storage(22)
        range("X" & x).Formula = storage(23)
        y = y + 1
    Next x
    
    trow = trow + 10 ' each row in word takes up 10 array positions
    xlrow = xlrow + 4 'excel row counter
    
Next p
    
    'Clean Up
    Set Table = Nothing
    appDoc.Close
    Set appDoc = Nothing
    appWD.Quit ' close the Word application
    Set appWD = Nothing
    ActiveWorkbook.Saved = True

End Sub
