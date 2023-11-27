Option Explicit

Sub ADO_Connection()

'Creating objects of Connection and Recordset

Dim conn As New Connection, rec As New Recordset

Dim DBPATH, PRVD, connString, query As String

'Declaring fully qualified name of database. Change it with your database's location and name.

DBPATH = "C:\Users\ExcelTip\Desktop\Test Database.accdb"

'This is the connection provider. Remember this for your interview.

PRVD = "Microsoft.ace.OLEDB.12.0;"

'This is the connection string that you will require when opening the the connection.

connString = "Provider=" & PRVD & "Data Source=" & DBPATH

'opening the connection

conn.Open connString

'the query I want to run on the database.

query = "SELECT * from customerT;"

'running the query on the open connection. It will get all the data in the rec object.

rec.Open query, conn

'clearing the content of the cells

Cells.ClearContents

'getting data from the recordset if any and printing it in column A of excel sheet.

If (rec.RecordCount <> 0) Then

Do While Not rec.EOF

Range("A" & Cells(Rows.Count, 1).End(xlUp).Row).Offset(1, 0).Value2 = _

rec.Fields(1).Value

rec.MoveNext

Loop

End If

'closing the connections

rec.Close

conn.Close

End Sub
