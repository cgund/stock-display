
'Application allows user to graphically display the stock performance of an individaul company over
'a selected time period


'Form for searching for stock ticker symbols
Public Class FormSearch
    Private symbol As New StockSymbol 'StockSymbol reference variable

    'Returns table containing search results.  If table is Null, then
    'displays error message
    Private Sub btnGo_Click(sender As System.Object, e As System.EventArgs) Handles btnGo.Click
        errProvider.Clear()
        Try
            Dim table = symbol.searchByCompany(txtCompany.Text()) 'Returns a table of search results
            dgvSymbol.DataSource = table
            Dim colSymbol = dgvSymbol.Columns.Item(0) 'Gets stock symbol from table
            colSymbol.HeaderText = "Symbol"
            colSymbol.Width = CInt(dgvSymbol.Width * 0.33) 'Sets width

            Dim colName = dgvSymbol.Columns.Item(1) 'Gets company name from table
            colName.HeaderText = "Company"
            colName.Width = CInt(dgvSymbol.Width * 0.67) 'Sets width
        Catch ex As SqlClient.SqlException
            MessageBox.Show(ex.Message, "Connection Error")
        Catch ex As ArgumentOutOfRangeException
            errProvider.SetError(dgvSymbol, "No Results Found")
        End Try
    End Sub

    'Clears controls, closes form
    Private Sub btnClose_Click(sender As System.Object, e As System.EventArgs) Handles btnClose.Click
        txtCompany.Text = String.Empty
        dgvSymbol.DataSource = Nothing
        errProvider.Clear()
        Me.Close()
    End Sub

    'Clears controls
    Private Sub btnClear_Click(sender As System.Object, e As System.EventArgs) Handles btnClear.Click
        errProvider.Clear()
        dgvSymbol.DataSource = Nothing
    End Sub

    'Gets collections of rows, gets selected row in collection, symbol from appropriate cell, and sets
    'the TextBox Text property in main form to contain symbol
    Private Sub btnSelect_Click(sender As System.Object, e As System.EventArgs) Handles btnSelect.Click
        If Not dgvSymbol.DataSource Is Nothing Then
            Dim selectedRows = dgvSymbol.SelectedRows
            Dim selectedRow = selectedRows.Item(0)
            Dim symbol = selectedRow.Cells.Item(0).Value
            FormMain.txtSymbol.Text = CStr(symbol)

            txtCompany.Text = String.Empty
            dgvSymbol.DataSource = Nothing
            Me.Close()
        Else
            errProvider.Clear()
            errProvider.SetError(btnSelect, "Enter Company Name")
        End If
    End Sub

    Private Sub FormSearch_FormClosing(sender As Object, e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        txtCompany.Text = String.Empty
        dgvSymbol.DataSource = Nothing
    End Sub
End Class