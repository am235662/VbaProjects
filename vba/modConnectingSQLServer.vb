Option Explicit

Sub DboConnect()

        Dim connection As ADODB.connection
        Set connection = New ADODB.connection
        
        Dim serverName As String, databaseName As String
        Let serverName = "localhost"
        Let databaseName = "ATGTester"
        
        With connection
            .ConnectionString = "provider=SQLOLEDB;Server=" & serverName & _
                ";database=" & databaseName & "; Integrated Security=SSPI;"
            .CommandTimeout = 10
            .Open
        End With
        
        If connection.State = 1 Then
            Debug.Print "Connected"
        End If
            
        Dim sqlQuery As String
        Let sqlQuery = "select * from [ATGTester].[TestTable].[students]"
        
        Dim rsSQL As New ADODB.Recordset
        rsSQL.CursorLocation = adUseClient
        rsSQL.Open sqlQuery, connection, adOpenStatic
        
        Sheet1.Range("A1").CopyFromRecordset rsSQL
        

End Sub