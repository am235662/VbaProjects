    Option Explicit
Sub writeDataToWorksheet(resultset As ADODB.Recordset)
    Dim ws As Worksheet
    Dim f As ADODB.Field
    Dim i As Integer                                 ' it is default set to 0. hence no need to explicitly set it to 0
    
    Set ws = Worksheets.Add
    ws.Select
    
    ' it will loop through each header from record set and copy it as well
    For Each f In resultset.Fields
        i = i + 1
        ws.Cells(1, i).Value = f.Name
            
    Next f
    
    Range("A2").CopyFromRecordset resultset
    
    'tidying up the data
    Range("A1").CurrentRegion.WrapText = False
    Range("A1").CurrentRegion.EntireColumn.AutoFit
    
    
End Sub
Sub ConnectToMySQL()
    Dim conn As ADODB.Connection
    Set conn = New ADODB.Connection
    
    Dim SQL As String
    SQL = "Select ID, first_name from sql_project.employees Where city  = 'New York' Order by salary desc"
    
    Dim rs As ADODB.Recordset
    
    ' Set connection string using the DSN you created and provide the correct username and password
    conn.ConnectionString = "DSN=MYSQLConnector;UID=root;PWD=Abh@1113;"
    
    ' Open the connection
    On Error GoTo ErrorHandler
    conn.Open
    
    Set rs = New ADODB.Recordset
    rs.ActiveConnection = conn
    
    'setting table as source for recordset
    rs.Source = "sql_project.employees"
    rs.Open SQL, conn
    
    'copy data from recordset or table from database to range(A1) in current excel
    'ShRecordSet.Range("A1").CopyFromRecordset rs
    
    writeDataToWorksheet rs
    
    'close recordset
    rs.Close
    
    ' Close the connection
    conn.Close
    Set conn = Nothing
    Exit Sub
    
ErrorHandler:
    MsgBox "Error connecting to MySQL: " & Err.Description
    conn.Close
    Set conn = Nothing
End Sub


