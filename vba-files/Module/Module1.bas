Attribute VB_Name = "Module1"
Private Declare PtrSafe Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Private Const VK_RETURN As Long = &HD

Private Declare PtrSafe Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)


Sub Main()
    Dim ChartObj As ChartObject
    Dim ChartData As Chart
    Dim SeriesDataX As Range
    Dim SeriesDataY As Range
    
    Set SeriesDataX = GetRange
    Sleep 500
    Set SeriesDataY = GetRange

    Set ChartObj = ActiveSheet.ChartObjects.Add(Left:=100, Width:=375, Top:=75, Height:=225)
    Set ChartData = ChartObj.Chart

    With ChartData
        .SeriesCollection.NewSeries
        .SeriesCollection(1).XValues = SeriesDataX
        
        .SeriesCollection.NewSeries
        .SeriesCollection(1).Values = SeriesDataY
        
        .ChartType = xlLine
        
        .HasLegend = False
        .HasTitle = True
        .ChartTitle.Text = "bbbb"
    End With

End Sub


Function GetRange() As Range
    UserForm1.Show
    Do
        DoEvents
    Loop Until GetAsyncKeyState(VK_RETURN) < 0
    UserForm1.Hide
    Set GetRange = Selection
End Function
