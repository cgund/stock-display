
'Application allows user to graphically display the stock performance of an individaul company over
'a selected time period


'Class for representing daily stock prices of individual company
Public Class StockPrice
    Private adapter As New StockDataSetTableAdapters.NYSE_2014TableAdapter 'Table adapter reference

    Private symbol As String
    Private startDate As Date
    Private endDate As Date

    'Three agrugment constructor.  Initializes fields
    Public Sub New(sym As String, sDate As Date, eDate As Date)
        symbol = sym
        startDate = sDate
        endDate = eDate
    End Sub

    'Returns table of closing prices for individual company over selected period
    Public Function getClosePrices() As DataTable
        Dim dailyTable = adapter.GetPriceByDateRange(symbol, CStr(startDate), CStr(endDate))
        Return (CType(dailyTable, DataTable))
    End Function

    'Gets average price of stock during time period
    Public Function getAvgPrice() As Decimal
        Dim avg As Decimal
        Try
            avg = CDec(adapter.GetAvgPrice(symbol, CStr(startDate), CStr(endDate)))
        Catch ex As Exception
            Console.WriteLine(ex.Message)
        End Try
        Return CDec(avg)
    End Function

    'Gets max price of stock during time period.  Adjusts dates to avoid null values
    Public Function getMaxPrice() As Decimal
        Dim max As Nullable(Of Decimal) = adapter.GetMaxPrice(symbol, CStr(startDate), CStr(endDate))
        If max.HasValue = False Then
            Dim maxIsNull = True
            While maxIsNull
                startDate = startDate.AddDays(1)
                max = adapter.GetMaxPrice(symbol, CStr(startDate), CStr(endDate))
                If max.HasValue Then
                    maxIsNull = False
                Else
                    endDate = endDate.AddDays(1)
                    max = adapter.GetMaxPrice(symbol, CStr(startDate), CStr(endDate))
                    If max.HasValue Then
                        maxIsNull = False
                    End If
                End If
            End While
        End If
        Return CDec(max)
    End Function

    'Gets min price of stock during time period.  Adjusts dates to avoid null values
    Public Function getMinPrice() As Decimal
        Dim min As Nullable(Of Decimal) = adapter.GetMinPrice(symbol, CStr(startDate), CStr(endDate))
        If Not min.HasValue Then
            Dim minIsNull As Boolean
            While minIsNull
                startDate = startDate.AddDays(1)
                min = adapter.GetMinPrice(symbol, CStr(startDate), CStr(endDate))
                If min.HasValue Then
                    minIsNull = False
                Else
                    endDate = endDate.AddDays(1)
                    min = adapter.GetMinPrice(symbol, CStr(startDate), CStr(endDate))
                    If min.HasValue Then
                        minIsNull = False
                    End If
                End If
            End While
        End If
        Return CDec(min)
    End Function

    'Gets percentage change of stock during time period.  Adjusts dates to avoid null values
    Public Function getPercentChange() As Decimal
        Dim startDatePrice As Nullable(Of Decimal) = adapter.GetStartPrice(symbol, CStr(startDate))

        If Not startDatePrice.HasValue Then
            Dim dateExists = False
            While dateExists = False
                startDate = startDate.AddDays(1)
                startDatePrice = adapter.GetStartPrice(symbol, CStr(startDate))
                If startDatePrice.HasValue Then
                    dateExists = True
                End If
            End While
        End If

        Dim endDatePrice As Nullable(Of Decimal) = adapter.GetEndPrice(symbol, CStr(endDate))

        If Not endDatePrice.HasValue Then
            Dim dateExists = False
            While dateExists = False
                endDate = endDate.AddDays(1)
                endDatePrice = adapter.GetEndPrice(symbol, CStr(endDate))
                If endDatePrice.HasValue Then
                    dateExists = True
                End If
            End While
        End If
        Return CDec((endDatePrice / startDatePrice) - 1)
    End Function
End Class
