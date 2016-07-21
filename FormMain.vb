
'Application allows user to graphically display the stock performance of an individaul company over
'a selected time period

Imports System.Data.SqlClient


'Main Form.  User manually enters stock symbol(if known), selects date range (dates prior to or after 2014 not 
'permitted), or selects option to lookup stock symbol
Public Class FormMain
    'Handles button click
    Private Sub btnDisplay_Click(sender As System.Object, e As System.EventArgs) Handles btnDisplay.Click
        errProvider.SetError(dtpEnd, String.Empty)
        errProvider.SetError(txtSymbol, String.Empty)

        Dim symbol As String = txtSymbol.Text().Trim()
        Dim startDate As Date = dtpStart.Value
        Dim endDate As Date = dtpEnd.Value

        Dim sSymbol As New StockSymbol(symbol) 'New instance of StockSymbol

        'Queries dataset to determine if symbol exists, checks proper date inputs, creates new instance 
        'of FormChart, gets company name matched to stock symbol, calls FormChart display method
        Try
            If (sSymbol.exists()) Then
                If startDate < endDate Then
                    Dim formChart As New FormChart
                    Dim company = sSymbol.getCompanyName()
                    formChart.display(symbol, company, startDate, endDate)
                Else
                    errProvider.SetError(dtpEnd, "End date must be greater than start date")
                End If
            Else
                errProvider.SetError(txtSymbol, "Stock symbol does not exist")
            End If
        Catch ex As SqlException
            MessageBox.Show(ex.Message, "DataBase Connection Error")
        End Try
    End Sub

    'Removes error messages when form is activated
    Private Sub FormMain_Activated(sender As Object, e As System.EventArgs) Handles Me.Activated
        errProvider.SetError(dtpEnd, String.Empty)
        errProvider.SetError(txtSymbol, String.Empty)
    End Sub

    'Opens search form
    Private Sub linkSearch_LinkClicked(sender As System.Object, e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles linkSearch.LinkClicked
        FormSearch.errProvider.Clear()
        FormSearch.ShowDialog()
    End Sub
End Class
