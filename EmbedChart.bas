Attribute VB_Name = "EmbedChart"
Sub ChartMakeButton()

Dim ChartCount As Integer
ChartCount = ActiveSheet.ChartObjects.Count

If ChartCount > 0 Then
ActiveSheet.ChartObjects.Delete
End If

Dim NLoops As Integer, LoopNumber As Integer

NLoops = Worksheets("Data Input").Range("Q2").Value - 1

LoopNumber = 1

Dim cht As ChartObject, rngChart As Range, destinationSheet As String, Results As Worksheet, StartRow As Integer, EndRow As Integer, StartChart As Integer, EndChart As Integer

Set Results = Worksheets("Results")

destinationSheet = ActiveSheet.Name

'Average Disease Risk

Do While NLoops >= 0

StartRow = NLoops * 27 + 1
EndRow = NLoops * 27 + 6
StartChart = NLoops * 13 + 13
EndChart = NLoops * 13 + 24

Set co = Sheets("Results").ChartObjects.Add(6, StartChart, 5, 8)
ActiveSheet.ChartObjects(LoopNumber).Activate
co.Chart.SetSourceData Source:=Results.Range(Results.Cells(StartRow, "A"), Results.Cells(EndRow, "B"))
co.Chart.ChartType = xlColumnClustered
co.Chart.ChartTitle.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(72, 72, 72)
co.Chart.Axes(xlCategory).TickLabels.Font.Color = RGB(72, 72, 72)
co.Chart.Axes(xlValue).TickLabels.Font.Color = RGB(72, 72, 72)
ActiveChart.SeriesCollection(1).Points(1).Interior.Color = RGB(59, 110, 172)
ActiveChart.SeriesCollection(1).Points(2).Interior.Color = RGB(59, 110, 172)
ActiveChart.SeriesCollection(1).Points(3).Interior.Color = RGB(59, 110, 172)
ActiveChart.SeriesCollection(1).Points(4).Interior.Color = RGB(59, 110, 172)
ActiveChart.SeriesCollection(1).Points(5).Interior.Color = RGB(59, 110, 172)
ActiveChart.SeriesCollection(1).Format.Shadow.Visible = msoTrue
Set cht = co.Chart.Parent
With cht
.Left = Results.Cells(StartChart, "F").Left
.Top = Results.Cells(StartChart, "F").Top
.Height = Results.Range(Results.Cells(StartChart, "F"), Cells(EndChart, "J")).Height
.Width = Results.Range(Results.Cells(StartChart, "F"), Cells(EndChart, "J")).Width
End With
co.Chart.Legend.Select
Selection.Delete

NLoops = NLoops - 1

LoopNumber = LoopNumber + 1

Loop

'Average Disease Cost

NLoops = Worksheets("Data Input").Range("Q2").Value - 1

Do While NLoops >= 0

StartRow = NLoops * 27 + 8
EndRow = NLoops * 27 + 13
StartChart = NLoops * 13 + 13
EndChart = NLoops * 13 + 24

Set co = Sheets("Results").ChartObjects.Add(6, StartChart, 5, 8)
ActiveSheet.ChartObjects(LoopNumber).Activate
co.Chart.SetSourceData Source:=Results.Range(Results.Cells(StartRow, "A"), Results.Cells(EndRow, "B"))
co.Chart.ChartType = xlColumnClustered
co.Chart.ChartTitle.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(72, 72, 72)
co.Chart.Axes(xlCategory).TickLabels.Font.Color = RGB(72, 72, 72)
co.Chart.Axes(xlValue).TickLabels.Font.Color = RGB(72, 72, 72)
ActiveChart.SeriesCollection(1).Points(1).Interior.Color = RGB(59, 110, 172)
ActiveChart.SeriesCollection(1).Points(2).Interior.Color = RGB(59, 110, 172)
ActiveChart.SeriesCollection(1).Points(3).Interior.Color = RGB(59, 110, 172)
ActiveChart.SeriesCollection(1).Points(4).Interior.Color = RGB(59, 110, 172)
ActiveChart.SeriesCollection(1).Points(5).Interior.Color = RGB(59, 110, 172)
ActiveChart.SeriesCollection(1).Format.Shadow.Visible = msoTrue
Set cht = co.Chart.Parent
With cht
.Left = Results.Cells(StartChart, "L").Left
.Top = Results.Cells(StartChart, "L").Top
.Height = Results.Range(Results.Cells(StartChart, "L"), Cells(EndChart, "P")).Height
.Width = Results.Range(Results.Cells(StartChart, "L"), Cells(EndChart, "P")).Width
End With
co.Chart.Legend.Select
Selection.Delete

NLoops = NLoops - 1

LoopNumber = LoopNumber + 1

Loop

'Average Biometrics

NLoops = Worksheets("Data Input").Range("Q2").Value - 1

Do While NLoops >= 0

StartRow = NLoops * 27 + 15
EndRow = NLoops * 27 + 26
StartChart = NLoops * 13 + 13
EndChart = NLoops * 13 + 24

Set co = Sheets("Results").ChartObjects.Add(6, StartChart, 5, 8)
ActiveSheet.ChartObjects(LoopNumber).Activate
co.Chart.SetSourceData Source:=Results.Range(Results.Cells(StartRow, "A"), Results.Cells(EndRow, "B"))
co.Chart.ChartType = xlColumnClustered
co.Chart.ChartTitle.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(72, 72, 72)
co.Chart.Axes(xlCategory).TickLabels.Font.Color = RGB(72, 72, 72)
co.Chart.Axes(xlValue).TickLabels.Font.Color = RGB(72, 72, 72)
ActiveChart.SeriesCollection(1).Points(1).Interior.Color = RGB(59, 110, 172)
ActiveChart.SeriesCollection(1).Points(2).Interior.Color = RGB(59, 110, 172)
ActiveChart.SeriesCollection(1).Points(3).Interior.Color = RGB(59, 110, 172)
ActiveChart.SeriesCollection(1).Points(4).Interior.Color = RGB(59, 110, 172)
ActiveChart.SeriesCollection(1).Points(5).Interior.Color = RGB(59, 110, 172)
ActiveChart.SeriesCollection(1).Points(6).Interior.Color = RGB(59, 110, 172)
ActiveChart.SeriesCollection(1).Points(7).Interior.Color = RGB(59, 110, 172)
ActiveChart.SeriesCollection(1).Points(8).Interior.Color = RGB(59, 110, 172)
ActiveChart.SeriesCollection(1).Points(9).Interior.Color = RGB(59, 110, 172)
ActiveChart.SeriesCollection(1).Points(10).Interior.Color = RGB(59, 110, 172)
ActiveChart.SeriesCollection(1).Points(11).Interior.Color = RGB(59, 110, 172)
ActiveChart.SeriesCollection(1).Format.Shadow.Visible = msoTrue
Set cht = co.Chart.Parent
With cht
.Left = Results.Cells(StartChart, "R").Left
.Top = Results.Cells(StartChart, "R").Top
.Height = Results.Range(Results.Cells(StartChart, "R"), Cells(EndChart, "V")).Height
.Width = Results.Range(Results.Cells(StartChart, "R"), Cells(EndChart, "V")).Width
End With
co.Chart.Legend.Select
Selection.Delete

NLoops = NLoops - 1

LoopNumber = LoopNumber + 1

Loop

'Metrics

Dim NYears As Integer

Dim NBar As Integer

Dim MetricCount As Integer

Dim LeftChart As Integer

Dim RightChart As Integer

MetricCount = 1

LeftChart = 24
RightChart = 28

NYears = Worksheets("Data Input").Range("Q2").Value

NBar = 1

NLoops = 21

FactorNumber = 5

StartChart = 13
EndChart = 24

Do While NLoops > 0

Set co = Sheets("Results").ChartObjects.Add(6, StartChart, 5, 8)
ActiveSheet.ChartObjects(LoopNumber).Activate
co.Chart.SetSourceData Source:=Results.Range(Results.Cells(1, FactorNumber), Results.Cells(NYears + 1, FactorNumber + 1))
co.Chart.ChartType = xlColumnClustered
co.Chart.ChartTitle.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(72, 72, 72)
co.Chart.Axes(xlCategory).TickLabels.Font.Color = RGB(72, 72, 72)
co.Chart.Axes(xlValue).TickLabels.Font.Color = RGB(72, 72, 72)

Do While NBar <= NYears
ActiveChart.SeriesCollection(1).Points(NBar).Interior.Color = RGB(59, 110, 172)
NBar = NBar + 1
Loop

ActiveChart.SeriesCollection(1).Format.Shadow.Visible = msoTrue
Set cht = co.Chart.Parent
With cht
.Left = Results.Cells(StartChart, LeftChart).Left
.Top = Results.Cells(StartChart, LeftChart).Top
.Height = Results.Range(Results.Cells(StartChart, LeftChart), Cells(EndChart, RightChart)).Height
.Width = Results.Range(Results.Cells(StartChart, LeftChart), Cells(EndChart, RightChart)).Width
End With
co.Chart.Legend.Select
Selection.Delete

NLoops = NLoops - 1

LoopNumber = LoopNumber + 1

FactorNumber = FactorNumber + 2

MetricCount = MetricCount + 1

StartChart = StartChart + 13
EndChart = EndChart + 13

If MetricCount = 4 Then
MetricCount = 1
StartChart = 13
EndChart = 24
LeftChart = LeftChart + 6
RightChart = RightChart + 6
End If

Loop

End Sub
