Attribute VB_Name = "ModuleChartFormat"
Sub formatChart()
Attribute formatChart.VB_ProcData.VB_Invoke_Func = " \n14"
'
' formatChart Macro
'
' Takes the active charts and reformats it as specified
'
    'change these options as needed
    Dim fontSize As Variant: fontSize = 14 'font size
    Dim visibleLine As Boolean: visibleLine = False 'line connecting points
    Dim markerStyle As Integer: markerStyle = xlMarkerStyleCircle 'style of markers, can also be xlMarkerStyleNone
    Dim markerFill As Boolean: markerFill = True
    Dim markerSize As Integer: markerSize = 7 'how large the markers are
    Dim chartType As Integer: chartType = xlXYScatter 'which type of chart, can also be xlXYScatterLines or xlXYScatterLinesNoMarkers

    Dim redcolor, pinkcolor, purplecolor, indigocolor, bluecolor, greencolor, orangecolor, colori As Long
    redcolor = RGB(244, 67, 54)
    pinkcolor = RGB(233, 30, 99)
    purplecolor = RGB(156, 39, 176)
    indigocolor = RGB(63, 81, 181)
    bluecolor = RGB(33, 150, 243)
    greencolor = RGB(76, 175, 80)
    orangecolor = RGB(255, 152, 0)
    
    Dim colorSet, nColors As Variant
    Dim ch As Chart
    Dim n As Integer
    Dim invert As Boolean
    Dim S As Series
    Dim P As Point
    
    colorSet = Array(bluecolor, redcolor, greencolor, purplecolor, orangecolor, pinkcolor, indigocolor) 'first series is red, then blue, etc.
    nColors = UBound(colorSet, 1) - LBound(colorSet, 1) + 1
    
    Set ch = ActiveChart
    n = ch.SeriesCollection.Count
    
    ch.ChartArea.Format.TextFrame2.TextRange.Font.Size = fontSize
    
    For i = 1 To n
        'Set the color scheme
        colori = colorSet((i - 1) Mod (nColors))
        
        'determine if cycling back on the first color, if so invert
        If Int((i - 1) / nColors) > 0 Then
            invert = True
        End If
            
        Set S = ch.SeriesCollection(i)
        With S
            .ClearFormats
            .Type = chartType
            
            'Set connecting line properties
            .Format.Line.Visible = visibleLine
            If visibleLine Then
                With .Border
                    .Color = colori
                    .Weight = xlHairline
                    '.LineStyle = xlDash
                End With
            End If
            
            'Set marker properties
            If ((Not invert) = markerFill) Then
                .MarkerBackgroundColor = colori
                .MarkerForegroundColorIndex = xlColorIndexNone
            Else
                .MarkerBackgroundColorIndex = xlColorIndexNone
                .MarkerForegroundColor = colori
            End If
            .markerSize = markerSize
            .markerStyle = markerStyle
        
            'Set the marker line style point by point
            'For Each P In S.Points
            '    With P
            '        .MarkerStyle = xlMarkerStyleCircle
            '        .Format.Line.DashStyle = msoLineSolid
            '        .Format.Line.Weight = xlThin
            '        .Format.Line.ForeColor.RGB = RGB(0, 0, 0)
            '    End With
            'Next P
        End With
    Next i
End Sub
