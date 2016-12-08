Attribute VB_Name = "ModuleChartFormat"
Sub formatChart()
Attribute formatChart.VB_ProcData.VB_Invoke_Func = " \n14"
'
' formatChart Macro
'
' Takes the active charts and reformats it as specified
'
    'change these options as needed
    Dim visibleLine As Boolean: visibleLine = True 'line connecting points
    Dim markerFill As Boolean: markerFill = True 'either fill with no border, or border with no fill
    Dim markerTransparency As Double: markerTransparency = 0.5
    Dim markerSize As Integer: markerSize = 7
    Dim markerStyle As Integer: markerStyle = xlMarkerStyleCircle
    Dim fontSize As Variant: fontSize = 14

    Dim redcolor, pinkcolor, purplecolor, indigocolor, bluecolor, greencolor, orangecolor, colori As Long
    redcolor = RGB(244, 67, 54)
    pinkcolor = RGB(233, 30, 99)
    purplecolor = RGB(156, 39, 176)
    indigocolor = RGB(63, 81, 181)
    bluecolor = RGB(33, 150, 243)
    greencolor = RGB(76, 175, 80)
    orangecolor = RGB(255, 152, 0)
    
    Dim colorSet As Variant
    Dim ch As Chart
    Dim n, chartType, nColors As Integer
    Dim invert As Boolean
    Dim S As Series
    'Dim P As Point
    
    colorSet = Array(redcolor, bluecolor, greencolor, purplecolor, orangecolor, pinkcolor, indigocolor) 'first series is red, then blue, etc.
    nColors = UBound(colorSet, 1) - LBound(colorSet, 1) + 1
    
    Set ch = ActiveChart
    
    If visibleLine Then
        chartType = xlXYScatterLines
    Else
        chartType = xlXYScatter
    End If
    ch.chartType = chartType
    ch.ClearToMatchStyle
    ch.ChartArea.Format.TextFrame2.TextRange.Font.Size = fontSize
    n = ch.SeriesCollection.Count
    
    For i = 1 To n
        
        'set color and determine if cycling back on the first color, if so invert
        colori = colorSet((i - 1) Mod (nColors))
        If Int((i - 1) / nColors) > 0 Then
            invert = True
        End If
            
        Set S = ch.SeriesCollection(i)

        'clear formatting
        S.Shadow = False
        S.Smooth = False
        
        'Set marker properties
        If ((Not invert) = markerFill) Then
            'set marker border invisible
            S.MarkerForegroundColorIndex = xlColorIndexNone

            'set marker fill color
            S.Format.Fill.Visible = msoTrue
            S.Format.Fill.ForeColor.RGB = colori
            S.Format.Fill.BackColor.RGB = colori
            
            'apply transparency
            S.Format.Fill.Transparency = markerTransparency
            
        Else
            'set marker border color
            S.MarkerForegroundColor = colori
            
            'set marker fill invisible
            S.MarkerBackgroundColorIndex = xlColorIndexNone
        End If

        S.markerSize = markerSize
        S.markerStyle = markerStyle
        
        'Set connecting line properties
        If visibleLine Then
            With S.Border
                .Color = colori
                .Weight = xlHairline
                '.LineStyle = xlDash
            End With
        End If
        
        'Set the marker line style point by point
        'For Each P In S.Points
        '    With P
        '        .MarkerStyle = xlMarkerStyleCircle
        '        .Format.Line.DashStyle = msoLineSolid
        '        .Format.Line.Weight = xlThin
        '        .Format.Line.ForeColor.RGB = RGB(0, 0, 0)
        '    End With
        'Next P

    Next i
End Sub
