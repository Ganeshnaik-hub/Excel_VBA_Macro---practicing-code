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
    
