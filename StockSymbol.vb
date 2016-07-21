
'Application allows user to graphically display the stock performance of an individaul company over
'a selected time period


'Contains methods for determining if symbol exists in dataset, searching for symbol, and retrieving 
'company name
Public Class StockSymbol
    Private Shared adapter As New StockDataSetTableAdapters.symbolsTableAdapter 'Reference to table adapter

    'Class member variables
    Private symbol As String
    Private companyName As String = String.Empty

    'Default no-arg constructor
    Public Sub New()

    End Sub

    'Overloaded constructor
    Public Sub New(symbol As String)
        Me.symbol = symbol
    End Sub

    'Queries dataset.  If parameter sym is contained in the dataset, then returns true.  Also
    'intializes companyName variable
    Public Function exists() As Boolean
        Dim table = adapter.GetBySymbol(symbol)
        Dim rowCount As Integer = table.Count
        If rowCount > 0 Then
            Dim row As StockDataSet.symbolsRow = table.Item(0)
            Try
                companyName = CStr(row.Item(1))
                Return True
            Catch ex As Exception
                Console.WriteLine(ex.Message)
                Return False
            End Try
        Else
            Return False
        End If
    End Function

    'Returns company name
    Public Function getCompanyName() As String
        Return companyName
    End Function

    'Searches dataset for company names that match user input.  Allows for partial matches, but 
    'not spelling errors
    Public Function searchByCompany(company As String) As DataTable
        Try
            company = company.Trim
            Dim table As DataTable = adapter.GetByCompany(company)
            If table.Rows.Count > 0 Then 'Exact match
                Return table
            Else
                Dim companyMinusEndPeriod = company.Substring(0, company.Length - 1)
                table = adapter.GetByCompany(companyMinusEndPeriod)
                If table.Rows.Count > 0 Then 'Match minus end period
                    Return table
                Else
                    Dim fullTable = adapter.GetData 'Returns full table
                    Dim alDataTable = New List(Of DataTable) 'List for matching tables

                    Dim charArray() As Char = {CChar(" ")}

                    Dim userInput = company.ToLower
                    For Each row As DataRow In fullTable.Rows 'Iterates through table
                        Dim matchTable = New DataTable
                        Dim storedCompanyName = CStr(row.Item(1))
                        Dim storedNameToLower = storedCompanyName.ToLower
                        Dim storedNameArray() = storedNameToLower.Split(charArray) 'Splits with delimiter

                        Dim isMatch = False
                        Dim counter = 0
                        Dim partialCompanyName = String.Empty
                        Dim recombinedName = String.Empty

                        While (counter < storedNameArray.Length And isMatch = False)
                            partialCompanyName = storedNameArray(counter) 'Each element of array
                            partialCompanyName = partialCompanyName.Trim
                            recombinedName = recombinedName & " " & storedNameArray(counter) 'Rebuilds name
                            recombinedName = recombinedName.Trim
                            If (userInput.Equals(partialCompanyName) Or userInput.Equals(recombinedName)) Then
                                matchTable = adapter.GetByCompany(storedCompanyName)
                                alDataTable.Add(matchTable) 'Adds Table to List
                                isMatch = True 'Terminates loop
                            End If
                            counter = counter + 1
                        End While
                    Next

                    Dim mergedTable = New DataTable
                    For Each table In alDataTable 'Merges all tables in List
                        mergedTable.Merge(table)
                    Next

                    Return mergedTable
                End If
            End If
        Catch ex As Exception

        End Try
        Return Nothing
    End Function
End Class
