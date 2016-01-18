Attribute VB_Name = "ExcelCharts"
Option Explicit

Public Enum Corner
    TopLeft = 1
    TopRight = 2
    BottomLeft = 3
    BottomRight = 4
End Enum

'---------------------------------------------------------------------------------------
' Procedure : MakeActiveWorkbookChartsIdentical
' Author    : bradley_handziuk
' Date      : 3/26/2014
' Purpose   : This is intended to be used to uniformly format a bunch of charts in a file.
'               Change as appropriate. Probably not going to suite the current file's needs OTB but is a good start
'---------------------------------------------------------------------------------------
Sub MakeActiveWorkbookChartsIdentical()
    Dim cht As Chart
    Dim wkbk As Workbook
    Set wkbk = ActiveWorkbook
    Dim ax As Axis
    For Each cht In wkbk.Charts
        Debug.Print cht.Name
        If cht.HasTitle Then
            cht.ChartTitle.Left = cht.ChartArea.Width
            cht.ChartTitle.Left = cht.ChartTitle.Left / 2
            cht.ChartTitle.Top = 5
        End If
        
        With cht.PageSetup
            .BottomMargin = 18
            .TopMargin = 18
            .LeftMargin = 18
            .RightMargin = 18 ' 18 = 0.25 inches (72 dpi)
        End With
        
        If cht.HasTitle Then
            cht.plotArea.Top = cht.ChartTitle.Top + cht.ChartTitle.Height + 10
        End If
        cht.plotArea.Left = 40
        cht.plotArea.Width = 800
        
        Set ax = cht.Axes(1, xlPrimary)
        
        cht.plotArea.Height = ax.Height + 390
        If ax.HasTitle Then
            ax.AxisTitle.Left = cht.plotArea.Width / 2 + cht.plotArea.Left - ax.AxisTitle.Width / 2
            ax.AxisTitle.Top = cht.plotArea.Height + ax.Height + cht.ChartTitle.Height + 30
        End If
        ax.MaximumScale = 41577
        ax.majorUnit = 20
        
        Set ax = cht.Axes(2, xlPrimary)
        If ax.HasTitle Then
            ax.AxisTitle.Top = cht.plotArea.Height / 2 + cht.plotArea.Top - ax.AxisTitle.Height / 2
            ax.AxisTitle.Left = 5
        End If
        
        
        If cht.HasAxis(2, 2) Then
            Set ax = cht.Axes(2, xlSecondary)
            If ax.HasTitle Then
                ax.AxisTitle.Top = cht.plotArea.Height / 2 + cht.plotArea.Top - ax.AxisTitle.Height / 2
            
                cht.plotArea.Width = cht.plotArea.Width - ax.Width - ax.AxisTitle.Width + 10
                ax.AxisTitle.Left = cht.plotArea.Left + cht.plotArea.Width
            End If
        
        End If
    Next cht

End Sub

Sub FormatCharts()
    Dim cht As Chart
    
    Dim DoLogChart As Boolean
    DoLogChart = False
    
    For Each cht In ActiveWorkbook.Charts
        
        FormatLegend cht
        FormatPrimaryHorizontalTitle cht
        
        SnapShapeToClosestCorner cht, "Interim"
        
        Dim vertAxis As Axis
        Set vertAxis = cht.Axes(XlAxisType.xlValue, xlPrimary)
        vertAxis.HasMajorGridlines = True
        vertAxis.HasMinorGridlines = False
        
        cht.plotArea.Border.color = RGB(0, 0, 0)
                
        Dim horxAxis As Axis
        Set horxAxis = cht.Axes(XlAxisType.xlCategory, xlPrimary)
        horxAxis.HasMajorGridlines = False
        horxAxis.HasMinorGridlines = False
        
        AddInjectionTestTimeline cht
       
        If DoLogChart Then
            'FormatLogAxis cht, vertAxis, 1000, 0.001 ' CVOCs
            'FormatLogAxis cht, vertAxis, 100, 0.0001 ' dissolved gasses
        Else
           ' FormatLinearAxis cht, vertAxis, 300, -550, 50
        End If
        
    Next cht
End Sub

'---------------------------------------------------------------------------------------
' Procedure : ShadeBelow
' Author    : bradley_handziuk
' Date      : 12/19/2014
' Purpose   : http://peltiertech.com/Excel/Charts/VBAdraw.html#ShadeBelowCode
'           Shades below the specified series.
'---------------------------------------------------------------------------------------
Sub ShadeBelow(myCht As Chart, mySrs As Series)
      
    Dim Npts As Integer, Ipts As Integer
    Dim myBuilder As FreeformBuilder
    Dim myShape As Shape
    Dim Xnode As Double, Ynode As Double
    Dim Xmin As Double, Xmax As Double
    Dim Ymin As Double, Ymax As Double
    Dim Xleft As Double, Ytop As Double
    Dim Xwidth As Double, Yheight As Double
    Dim ybottom As Double, xRight As Double
      
    Xleft = myCht.plotArea.InsideLeft
    Xwidth = myCht.plotArea.InsideWidth
    Ytop = myCht.plotArea.InsideTop
    Yheight = myCht.plotArea.InsideHeight
    ybottom = Ytop + Yheight
    xRight = Xleft + Xwidth
    
    Dim shapeName As String
    shapeName = mySrs.Name & "Fill Under"
    
    Dim shp As Shape
    For Each shp In myCht.Shapes
        If shp.Name = shapeName Then
            shp.Delete
        End If
    Next shp
    
    Xmin = myCht.Axes(1).MinimumScale
    Xmax = myCht.Axes(1).MaximumScale
    Ymin = myCht.Axes(2).MinimumScale
    Ymax = myCht.Axes(2).MaximumScale
    
    Debug.Print "abc",
    Npts = mySrs.Points.Count
    
    ' first point
    Xnode = Xleft + (mySrs.XValues(1) - Xmin) * Xwidth / (Xmax - Xmin)
    Ynode = Ytop + Yheight
    Set myBuilder = myCht.Shapes.BuildFreeform(msoEditingAuto, Xnode, Ynode)
    
    ' remaining points
    For Ipts = 1 To Npts
        
        Xnode = mySrs.Points(Ipts).Left
        Ynode = mySrs.Points(Ipts).Top
     'Xnode = Xleft + (mySrs.XValues(Ipts) - Xmin) * Xwidth / (Xmax - Xmin)
     ' Ynode = Ytop + (Ymax - mySrs.Values(Ipts)) * Yheight / (Ymax - Ymin)
      
      If Ynode < Ytop Then Ynode = Ytop
      If Ynode > ybottom Then Ynode = ybottom
      If Xnode < Xleft Then Xnode = Xleft
      If Xnode > xRight Then Xnode = xRight
      
      Debug.Print Xnode & ", " & Ynode
      myBuilder.AddNodes msoSegmentLine, msoEditingAuto, Xnode, Ynode
    Next
    
    'add bottom to last point
    Xnode = Xleft + (mySrs.XValues(Npts) - Xmin) * Xwidth / (Xmax - Xmin)
    Ynode = Ytop + Yheight
    myBuilder.AddNodes msoSegmentLine, msoEditingAuto, Xnode, Ynode
    
    'add bottom to first point
    Xnode = Xleft + (mySrs.XValues(1) - Xmin) * Xwidth / (Xmax - Xmin)
    Ynode = Ytop + Yheight
    myBuilder.AddNodes msoSegmentLine, msoEditingAuto, Xnode, Ynode
    
    Set myShape = myBuilder.ConvertToShape
    
    With myShape
      .Name = shapeName
      .Fill.ForeColor.RGB = RGB(155, 155, 155) ' YELLOW  ' mySrs.Format.Line.ForeColor.RGB '
      .Fill.Transparency = 0.7
      .Line.Visible = False
      .ZOrder msoSendToBack
    End With
    
End Sub

'---------------------------------------------------------------------------------------
' Procedure : AddInjectionTestTimeline
' Author    : bradley_handziuk
' Date      : 12/22/2014
' Purpose   :
'---------------------------------------------------------------------------------------
Private Sub AddInjectionTestTimeline(cht As Chart)
    Dim serColl As SeriesCollection
    Dim ser As Series
    Set serColl = cht.SeriesCollection()
    Dim newSeriesName As String
    newSeriesName = "Injection Test"
    
    On Error Resume Next
    Set ser = cht.SeriesCollection(newSeriesName)
    If Not ser Is Nothing Then
        ser.Delete
    End If
    On Error GoTo 0
    
    Set ser = cht.SeriesCollection.NewSeries
    Dim wk As Workbook
    Set wk = ActiveWorkbook

    With ser
        .Name = newSeriesName
        .XValues = wk.Names("InjectionTestDates").RefersTo '"=Data!InjectionTestDates"
        .Values = wk.Names("InjectionTestValues").RefersTo ' "=Data!InjectionTestValues"
        .PlotOrder = 1
        With .Format.Line
            .Visible = msoTrue
            .ForeColor.ObjectThemeColor = msoThemeColorText1
            .ForeColor.TintAndShade = 0
            .ForeColor.Brightness = 0
        End With
    End With
    
    If cht.HasLegend Then
        Dim n As Long
        With cht
          For n = .SeriesCollection.Count To 1 Step -1
             With .SeriesCollection(n)
                 If .Name = newSeriesName And cht.SeriesCollection.Count = cht.Legend.LegendEntries.Count Then
                    cht.Legend.LegendEntries(n).Delete
                 End If
             End With
          Next n
        End With
    End If
    
    ShadeBelow cht, ser
    
    Dim shp As Shape
    On Error Resume Next
    For Each shp In cht.Shapes
        If shp.OLEFormat.Object.text Like "*Pre Injection*" Then
            shp.Delete
        End If
        If shp.OLEFormat.Object.text Like "*Post Injection*" Then
            shp.Delete
        End If
    Next shp
    On Error GoTo 0
    
    Dim verticalCenterPre As Double, verticalCenterPost As Double
    Dim pts As Points
    Set pts = ser.Points
    verticalCenterPre = pts(1).Left - 35
    verticalCenterPost = pts(3).Left + 35
    AddTextBox cht, "Pre Injection", verticalCenterPre
    AddTextBox cht, "Post Injection", verticalCenterPost
    
End Sub

'---------------------------------------------------------------------------------------
' Procedure : AddTextBox
' Author    : bradley_handziuk
' Date      : 12/22/2014
' Purpose   : adds a text box with the specified text to the chart, cht.
'---------------------------------------------------------------------------------------
Private Sub AddTextBox(cht As Chart, text As String, Optional verticalCenter As Double = 0)

    Dim shp
    Set shp = cht.Shapes.AddTextBox(msoTextOrientationHorizontal, 163, 58, 179, 19) ' add it anywhere then move it around later
    
    With shp
        .TextFrame2.TextRange.Characters.text = text
        With .TextFrame2
            .VerticalAnchor = msoAnchorMiddle
            .HorizontalAnchor = msoAnchorCenter
            .WordWrap = msoFalse
            .MarginRight = 0
            .MarginLeft = 0
        End With
        .Width = 20
        .Left = verticalCenter - .Width / 2
        .Top = cht.plotArea.Top + 5
    End With
End Sub

'---------------------------------------------------------------------------------------
' Procedure : FormatLegend
' Author    : bradley_handziuk
' Date      : 12/22/2014
' Purpose   : formats the legend position
'---------------------------------------------------------------------------------------
Private Sub FormatLegend(cht As Chart)
    If cht.HasLegend Then
        cht.Legend.Left = 0
        cht.Legend.Width = cht.ChartArea.Width
        cht.Legend.Top = 455.855748031496
        cht.Legend.Border.color = RGB(255, 255, 255)
    End If
    
End Sub

'---------------------------------------------------------------------------------------
' Procedure : FormatPrimaryHorizontalTitle
' Author    : bradley_handziuk
' Date      : 12/22/2014
' Purpose   : Formats teh primary axis title position.
'---------------------------------------------------------------------------------------
Sub FormatPrimaryHorizontalTitle(cht As Chart)
    Dim horxAxis As Axis
    Set horxAxis = cht.Axes(XlAxisType.xlCategory, xlPrimary)
    
    Dim centerOfTitle As Double
    centerOfTitle = cht.plotArea.Left + cht.plotArea.Width / 2
    horxAxis.AxisTitle.Left = centerOfTitle - horxAxis.AxisTitle.Width
End Sub

'---------------------------------------------------------------------------------------
' Procedure : SnapShapeToCorner
' Author    : bradley_handziuk
' Date      : 12/22/2014
' Purpose   : if a corner is not specified then the shape is snapped to the closest corner
'---------------------------------------------------------------------------------------
Sub SnapShapeToClosestCorner(cht As Chart, withTextLike As String, Optional snapToCorner As Corner)

    Dim shp As Shape

    For Each shp In cht.Shapes
        If shp.OLEFormat.Object.text Like "*" & withTextLike & "*" Then
            If snapToCorner = 0 Then snapToCorner = ClosestCorner(cht.plotArea, shp)
            
            Select Case snapToCorner
                Case Corner.BottomLeft
                    shp.Left = cht.plotArea.InsideLeft
                    shp.Top = cht.plotArea.InsideTop + cht.plotArea.InsideHeight - shp.Height
                    
                Case Corner.BottomRight
                    shp.Left = cht.plotArea.InsideLeft + cht.plotArea.InsideWidth - shp.Width
                    shp.Top = cht.plotArea.InsideTop + cht.plotArea.InsideHeight - shp.Height
                    
                Case Corner.TopLeft
                    shp.Left = cht.plotArea.InsideLeft
                    shp.Top = cht.plotArea.InsideTop
                    
                Case Corner.TopRight
                    shp.Left = cht.plotArea.InsideLeft + cht.plotArea.InsideWidth - shp.Width
                    shp.Top = cht.plotArea.InsideTop
                    
            End Select
          
            shp.OLEFormat.Object.Border.Weight = cht.plotArea.Border.Weight
        End If
    Next shp
End Sub

'---------------------------------------------------------------------------------------
' Procedure : ClosestCorner
' Author    : bradley_handziuk
' Date      : 12/22/2014
' Purpose   : Returns the corner of the plot area closest to the shape passed in.
'---------------------------------------------------------------------------------------
Public Function ClosestCorner(area As plotArea, shapeToSnap As Shape) As Corner
    Dim distanceToTop As Double, distanceToBottom As Double, distanceToLeft As Double, distanceToRight As Double
    
    distanceToLeft = shapeToSnap.Left - area.Left
    distanceToRight = (area.Left + area.Width) - (shapeToSnap.Left + shapeToSnap.Width)
    
    distanceToTop = shapeToSnap.Top - area.Top
    distanceToBottom = (area.Top + area.Height) - (shapeToSnap.Top + shapeToSnap.Height)
    
    If distanceToLeft > distanceToRight And distanceToTop > distanceToBottom Then
        ClosestCorner = BottomRight
    ElseIf distanceToLeft < distanceToRight And distanceToTop > distanceToBottom Then
        ClosestCorner = BottomLeft
    ElseIf distanceToLeft < distanceToRight And distanceToTop < distanceToBottom Then
        ClosestCorner = TopLeft
    ElseIf distanceToLeft > distanceToRight And distanceToTop < distanceToBottom Then
        ClosestCorner = TopRight
    End If
    
End Function

'---------------------------------------------------------------------------------------
' Procedure : FormatLogAxis
' Author    : bradley_handziuk
' Date      : 12/22/2014
' Purpose   : formats the specified axis as a log axis
'---------------------------------------------------------------------------------------
Sub FormatLogAxis(ax As Axis, Optional axisMaxValue, Optional axisMinValue)
 
    axisMaxValue = IIf(IsMissing(axisMaxValue), 100, axisMaxValue)
    axisMinValue = IIf(IsMissing(axisMinValue), 0.0001, axisMinValue)
    
    ax.MaximumScale = 100
    ax.MinimumScale = 0.0001
    ax.ScaleType = xlScaleLogarithmic
    ax.LogBase = 10
    ax.CrossesAt = axisMinValue
    ax.majorUnit = 10
    
    ax.MaximumScale = axisMaxValue
    ax.MinimumScale = axisMinValue

    ax.TickLabels.NumberFormat = "General"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : FormatLinearAxis
' Author    : bradley_handziuk
' Date      : 12/22/2014
' Purpose   : Formats the specified axis as liear and sets the max/min limits
'---------------------------------------------------------------------------------------
Sub FormatLinearAxis(cht As Chart, ax As Axis, Optional axisMaxValue, Optional axisMinValue, Optional majorUnit)

    Dim sers As SeriesCollection, ser As Series, val
    Dim maxValue As Double, minValue As Double
    maxValue = 0
        For Each ser In cht.SeriesCollection
            For Each val In ser.Values
                If val > maxValue Then maxValue = val
                If val < minValue Then minValue = val
            Next val
        Next ser

    Dim exp As Double
    Dim maxNextPowerOfTen As Double
    Dim minPrevValueOfOneHundred As Double, maxNextValueofOneHundred As Double


    exp = WorksheetFunction.Ceiling(WorksheetFunction.Log10(maxValue), 1)
    maxNextPowerOfTen = 10 ^ exp

    maxNextValueofOneHundred = WorksheetFunction.Ceiling(maxValue / 100, 1) * 100
    minPrevValueOfOneHundred = WorksheetFunction.Floor(minValue / 100, 1) * 100

    axisMaxValue = IIf(IsMissing(axisMaxValue), maxNextValueofOneHundred, axisMaxValue)
    axisMinValue = IIf(IsMissing(axisMinValue), minPrevValueOfOneHundred, axisMinValue)
    ''''''''''''''
    'dynamic
    ax.ScaleType = xlScaleLinear
    ax.majorUnit = IIf(IsMissing(majorUnit), CInt(Abs(axisMaxValue) / 10), majorUnit)


    ax.CrossesAt = axisMinValue
    
    ax.MaximumScale = axisMaxValue
    ax.MinimumScale = axisMinValue

    ax.TickLabels.NumberFormat = "General"
End Sub


