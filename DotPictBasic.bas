Attribute VB_Name = "Module1"
Sub MakeGradation()

Dim firstColor As Long
Dim firstRGB(3) As Long

Dim endColor As Long
Dim endRGB(3) As Long

Dim maxoffset As Long

For i = 0 To Selection.Rows.Count - 1
    maxoffset = Selection.Columns.Count - 1
    firstColor = Selection(1).offset(i, 0).Interior.Color
    endColor = Selection(1).offset(i, maxoffset).Interior.Color
     
    firstRGB(1) = firstColor And &HFF
    firstRGB(2) = firstColor And &HFF00&
    firstRGB(3) = firstColor And &HFF0000
    firstRGB(2) = firstRGB(2) \ (2 ^ 8)
    firstRGB(3) = firstRGB(3) \ (2 ^ 16)
    
    endRGB(1) = endColor And &HFF
    endRGB(2) = endColor And &HFF00&
    endRGB(3) = endColor And &HFF0000
    endRGB(2) = endRGB(2) \ (2 ^ 8)
    endRGB(3) = endRGB(3) \ (2 ^ 16)
    
    For j = 0 To maxoffset
        Dim r As Long
        Dim g As Long
        Dim b As Long
        
        Dim rate As Double
        
        rate = (CDbl(j) / CDbl(Selection.Columns.Count - 1))
    
        r = CLng(firstRGB(1) * (1 - rate) + endRGB(1) * (rate))
        g = CLng(firstRGB(2) * (1 - rate) + endRGB(2) * (rate))
        b = CLng(firstRGB(3) * (1 - rate) + endRGB(3) * (rate))
    
        Selection(1).offset(i, j).Interior.Color = RGB(r, g, b)
    Next j
Next i

End Sub
