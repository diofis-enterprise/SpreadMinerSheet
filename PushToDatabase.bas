Attribute VB_Name = "PushToDatabase"
Sub PushDatabase()

    ' Create a new ADODB connection
    Dim con As ADODB.Connection
    Set con = New ADODB.Connection
    
    ' Open the connection to the SQL Server database using the ODBC driver "PushEng"
    con.Open "SpreadMinerSheet;Server=sql.freedb.tech;Database=freedb_minersheet;Uid=freedb_minersheet;Pwd=9nub5MN&Yg5Dynn;"
    
    ' Mendefinisikan range yang berisi data valid
    Dim PushRange As String
    ' Menggunakan referensi Name Manager sebagai nilai range data valid
    PushRange = Range("ValidatedData").Value
    
    ' Define the range of data (PushRange)
    Dim rng As Range: Set rng = Application.Range(PushRange)
    Dim row As Range
    
    ' Delete existing data with the same numb
    For Each row In rng.Rows
    
        ' Retrieve numb from the first cell in the row
        numb = CLng(row.Cells(1).Value)
        
        ' Construct the SQL delete statement to remove existing data with the same numb
        Dim DeleteSql As String
        DeleteSql = "DELETE FROM tasks WHERE numb = '" & numb & "'"
        
        ' Execute the SQL statement to delete existing data
        con.Execute DeleteSql
    Next row
    
    ' Insert new data without deleting existing data with the same tanggal
    For Each row In rng.Rows
        ' Retrieve values from each cell in the row
        numb = row.Cells(1).Value
        items = row.Cells(2).Value
        pic = row.Cells(3).Value
                
        ' Get the current timestamp
        timestampValue = Format(Now(), "yyyy-mm-dd hh:mm:ss")
        
        ' Construct the SQL insert statement without specifying the 'id' column
        Dim InsertSql As String
        InsertSql = "INSERT INTO tasks (numb, items, pic, timestamp) VALUES ('" & numb & "', '" & items & "', '" & pic & "', '" & timestampValue & "')"
        
        ' Execute the SQL statement to insert the data into the table
        con.Execute InsertSql
    Next row
    
    ' Close the database connection
    con.Close
    
    ' Display a message box indicating completion
    MsgBox "Complete"
End Sub
