Attribute VB_Name = "ModuleChartFormat"
Sub formatChart()
Attribute formatChart.VB_ProcData.VB_Invoke_Func = " \n14"
'
' formatChart Macro
'

'
    Dim ch As Chart
    Dim n As Integer
    Set ch = Chart3
    n = ch.SeriesCollection.Count
    
    'set color 1 and color 2, with series fading between the two
    r1 = 255
    g1 = 0
    b1 = 0 'color 1 is red
    r2 = 0
    g2 = 0
    b2 = 255 'color 2 is blue
    
    hasLine = msoTrue 'msoFalse
    
    For i = 1 To n
        'Set the color scheme
        Dim colori As Long
        ri = CInt(((i - 1) / (n - 1)) * (r2 - r1) + r1)
        gi = CInt(((i - 1) / (n - 1)) * (g2 - g1) + g1)
        bi = CInt(((i - 1) / (n - 1)) * (b2 - b1) + b1)
        colori = RGB(ri, gi, bi)
        
        Dim S As Series
        Dim P As Point
        Set S = ch.SeriesCollection(i)
        S.ClearFormats
        
        'Set the type
        'S.ChartType = xlXYScatter 'xlXYScatterLines 'xlXYScatterLinesNoMarkers
        'S.Type = xlXYScatter
        
        'Set the line style
        With S.Format.Line
            .Visible = hasLine
            .Weight = xlHairline
            .Style = msoLineSingle
            .DashStyle = msoLineSingle
            .ForeColor.RGB = colori
        End With
        
        'Set marker fill
        S.MarkerBackgroundColorIndex = xlColorIndexNone
        '.MarkerBackgroundColor = colori
        
        'Set marker line color
        '.MarkerForegroundColorIndex = xlColorIndexNone
        S.MarkerForegroundColor = colori
        
        'Set marker size and style
        S.MarkerSize = 7
        S.MarkerStyle = xlMarkerStyleCircle 'xlMarkerStyleNone
        
        'Set the marker line style
        For Each P In S.Points
            With P
        '        .MarkerStyle = xlMarkerStyleCircle
        '        .Format.Line.DashStyle = msoLineSolid
        '        .Format.Line.Weight = xlThin
        '        .Format.Line.ForeColor.RGB = RGB(0, 0, 0)
            End With
        Next P
        
    Next i
End Sub
