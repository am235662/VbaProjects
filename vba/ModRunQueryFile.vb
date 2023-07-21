Option Explicit

Sub OpenQueryFile()

        Dim fso As Scripting.FileSystemObject
        
        ' to open the SQL file
        Dim ts As Scripting.TextStream
        
        'to read content of file as string
        Dim queryString As String
        ' to hold the path of selected file
        Dim queryPath As String
        
        'creates a file crowser dialog
        Dim SourceFile As FileDialog
        Set SourceFile = Application.FileDialog(msoFileDialogFilePicker)
        
        'it makes the chosen file path as default
            SourceFile.InitialFileName = "C:\Users\am235\OneDrive\Desktop\Excel VBA"
            SourceFile.ButtonName = "Run Query File"
            SourceFile.Title = "Choose SQL Script"
        ' Force user to select only one file at a time
            SourceFile.AllowMultiSelect = False
        'make user select only SQL file
            SourceFile.Filters.Clear
            SourceFile.Filters.Add "SQL Script", "*.sql"
        
        'checks if the file has been selected or not
        If SourceFile.Show = 0 Then Exit Sub
        
        queryPath = SourceFile.SelectedItems(1)
        
        Set fso = New Scripting.FileSystemObject
        'opening sql file
        Set ts = fso.OpenTextFile(queryPath)
        
        queryString = ts.ReadAll
        
        ' close sql file
        ts.Close
        
        getQueryResult queryString

End Sub

Sub getQueryResult(SQLString As String)
        
        Dim conn As ADODB.Connection
        Dim rs As ADODB.Recordset
        
        Set conn = New ADODB.Connection
        conn.ConnectionString = "DSN=MYSQLConnector;UID=root;PWD=Abh@1113;"
        
        ' Open the connection
        On Error GoTo ErrorHandler
        conn.Open
        
        'creating new instance of recordset
        Set rs = New ADODB.Recordset
        
        rs.ActiveConnection = conn
        rs.Source = SQLString
        rs.CursorType = adOpenForwardOnly
        rs.LockType = adLockReadOnly
        rs.Open
       
       'copies data from recordset to sheet "Raw Data"
        ShRwData.Select
        ShRwData.Range("A2").CopyFromRecordset rs
        
        
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
