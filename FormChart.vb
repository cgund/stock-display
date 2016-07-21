
'Application allows user to graphically display the stock performance of an individaul company over
'a selected time period


Imports System.Windows.Forms.DataVisualization.Charting

'Class to display stock prices in line chart
Public Class FormChart
    'Display method
    Public Sub display(symbol As String, company As String, startDate As Date, endDate As Date)
        Dim stock = New StockPrice(symbol, startDate, endDate) 'Creates new StockPrice

        Dim lineChart As Chart = New Chart() 'Line chart to display stock prices
        Dim chartArea As ChartArea = New ChartArea("ChartArea") 'Area to hold line chart
        Dim series As Series = New Series("Series") 'Series to hold data

        series.ChartArea = "ChartArea"
        series.ChartType = SeriesChartType.Line

        lineChart.ChartAreas.Add(chartArea)
        lineChart.Series.Add(series)

        lineChart.Series("Series").XValueMember = "Date" 'Binds X values to values in Date field of table
        lineChart.Series("Series").YValueMembers = "Close" 'Binds Y values to values in Close field of table

        lineChart.Size = pnlChart.Size 'Sets size of LineChart to Panel size

        Dim title = New Title(company & " (" & symbol & ")") 'Sets title
        title.Font = New Font("Tahoma", 12, FontStyle.Bold)
        lineChart.Titles.Add(title)

        chartArea.AxisY.Title = "Share Price ($)"

        pnlChart.Controls.Add(lineChart)

        'Gets and displays summary statistics
        Dim table = stock.getClosePrices
        Dim avgPrice = stock.getAvgPrice
        Dim maxPrice = stock.getMaxPrice
        Dim minPrice = stock.getMinPrice
        Dim percentChange = stock.getPercentChange
        lblAvg.Text = avgPrice.ToString("c")
        lblMax.Text = maxPrice.ToString("c")
        lblMin.Text = minPrice.ToString("c")
        If percentChange >= 0 Then
            lblChange.ForeColor = Color.Black
        Else
            lblChange.ForeColor = Color.Red
        End If
        lblChange.Text = percentChange.ToString("p2")
        lineChart.DataSource = table 'Binds line chart to table of stock prices

        lineChart.Show()

        Me.ShowDialog()
    End Sub

    Private Sub btnClose_Click(sender As System.Object, e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub
End Class