# Excel_VBA_Macro---practicing-code

here is my practicing VBA code
Sub example()

'Inserting value
Sheets("sheet1").Activate
Range("D6").Value = 52
Range("D8").Value = "Ganesh"
Sheets("sheet2").Activate
Range("e10").Value = 566656
Range("e12").Value = "Ganesh naik"
Sheets("sheet1").Activate

'Formating the cells
Range("D2:E11").Font.Size = 25
Range("d2:E11").Font.Size = 12
Range("d2:E11").Font.Bold = True
Range("d2:E11").Font.Bold = False
Range("d2:e11").Font.Italic = True
Range("d2:e11").Font.Italic = False
Range("d2:e11").Font.Underline = True
Range("d2:e11").Font.Underline = False
Range("d2:e11").Font.Name = "Ariel"
Range("d2:e11").Border.Value = 1
End Sub

Sub test()

' modifying the sheet
Sheets("sheet2").Visible = 2
Sheets("sheet2").Visible = 1
Sheets("sheet2").Activate
Range("b2") = Range("c2")
Range("b3").Value = Range("c3").Value
Range("b3").Font.Bold = True
Range("c3").Font.Bold = True
Range("c3").Font.Italic = True
Range("b3").Font.Italic = Range("c3").Font.Italic

'modifying the cell value
Range("a1") = Range("a1") + 2
Range("b3").Font.Size = 18
Range("b3").Font.Bold = True
Range("b3").Font.Italic = True
Range("b3").Font.Name = "Arial"

Sheets("Sheet2").Range("a1") = Range("a1") + 2
Sheets("Sheet2").Range("b3").Font.Size = 18
Sheets("Sheet2").Range("b3").Font.Bold = True
Sheets("Sheet2").Range("b3").Font.Italic = True
Sheets("Sheet2").Range("b3").Font.Name = "Arial"

'To avoding the repeating sheet names
'Start of the instruction with; with
With Sheets("sheet2").Range("c3")
  .Font.Bold = True
  .Font.Italic = True
  .Font.Size = 12
  .Font.Name = "Arial"
End With

'Same as we can also avoid the repeating of font text

With Sheets("sheet2").Range("c5")
     With .Font
         .Bold = True
         .Italic = True
         .Size = 25
         .Name = "Arial"
         End With
        End With

' Creating a message Box MsgBOX
Range("c5").ClearContents
MsgBox "The content of c5 has been cleared"

' creating box
If MsgBox(" Are you sure you want to delete the contents of c3?", vbYesNo, "Confirmation") = vbYes Then
Range("c3").ClearContents
MsgBox "The content of c5 has been cleared!"
End If

' then now
If MsgBox("Text", vbYesNo, "Title") = vbYes Then 'if the yes button is clicked
End If



End Sub

