Attribute VB_Name = "main_module"
Function Col_Letter(lngCol As Long) As String
    Dim vArr
    vArr = Split(Cells(1, lngCol).Address(True, False), "$")
    Col_Letter = vArr(0)
End Function

Sub get_data_sub()

Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
Application.AskToUpdateLinks = False

Dim cn As Object, rs As Object
Dim col_index As Integer
Dim db_path As String
Dim output_range As Range

Set data_out = ThisWorkbook.Worksheets("data_out")
Set main = ThisWorkbook.Worksheets("main")

database_name = main.Range("F1").Value
database = "[" & database_name & "]"
schema = "[" & main.Range("F2").Value & "]"
table_name = "[" & main.Range("F3").Value & "]"

table_full = database & "." & schema & "." & table_name

Dim cnn As ADODB.connection
Set cnn = New ADODB.connection

cnn.Provider = "SQLOLEDB.1" '<~~ mssql
user = "user"
pw = "pw"
database = "db"

' Open a connection using an ODBC DSN.
cnn.ConnectionString = "driver={SQL Server};" & "server=some_server_name;uid=" & user & ";pwd=" & pw & ";database=" & database & ";Trusted_Connection=True;"
cnn.Open

If cnn.State = adStateOpen Then
Else
    MsgBox "Connection to SQL server could not be established. Exiting..."
    Exit Sub
End If

Set rs = CreateObject("ADODB.Recordset")
rs.CursorLocation = adUseClient ' to je tukaj zato, da lahko preštejem število vrstic; drugaèe ne deluje

'build the query string
select_str = "SELECT "
from_str = " FROM " & table_full
where_str = " WHERE "


Dim last_col As Long
Dim curr_col As Long
last_col = main.Cells(4, main.Columns.Count).End(xlToLeft).Column

For Each col In main.Range("F4:" & Col_Letter(last_col) & "4")
    col_name = col.Value
    curr_col = col.Column
    curr_col_name = Col_Letter(curr_col)
    
    If col_name = "" Then GoTo next_col 'skip
    
    '---adds field name to string
    If curr_col <> last_col Then
        select_str = select_str & "[" & col_name & "], "
    Else
        select_str = select_str & "[" & col_name & "]"
    End If
    
    last_val = main.Cells(main.Rows.Count, curr_col).End(xlUp).Row
    
    
    If last_val <> 4 Then 'v primeru, da so vpisane vrednosti po katerih naj filtrira query, se tu zgradi IN() string
    
    'generates and adds inclusion/exclusion columns
    If InStr(1, where_str, "IN(") > 0 Then
        in_str = " AND [" & col_name & "] IN("
    Else
        in_str = "[" & col_name & "] IN("
    End If
    
        For Each filt In main.Range(curr_col_name & "5:" & curr_col_name & last_val)
            If filt.Row <> last_val Then
                in_str = in_str & "'" & filt.Value & "',"
            Else
                in_str = in_str & "'" & filt.Value & "')"
            End If
        Next filt
    where_str = where_str & in_str
    in_str = ""
    End If

next_col:
Next col

Debug.Print select_str & from_str & where_str
If InStr(1, where_str, "IN(") = 0 Then where_str = "" 'èe ni filtrov, resetira where string

rs.Open select_str & from_str & where_str, cnn, , , adCmdText

'excel has internal row number limit -> the following is to prevent looping errors
stevilo_vrstic = rs.RecordCount
If stevilo_vrstic > 999999 Then
    MsgBox ("Too much data in query. Reduce number by applying filters in fields.")
    GoTo skip_copy
End If

data_out.Cells.Clear

'column names
For intColIndex = 0 To rs.Fields.Count - 1
    data_out.Range("A1").Offset(0, intColIndex).Value = rs.Fields(intColIndex).Name
Next

data_out.Range("A2").CopyFromRecordset rs

'close and destroy objects; for freeing memory
skip_copy:
rs.Close
Set rs = Nothing
cnn.Close
Set cnn = Nothing

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
Application.AskToUpdateLinks = True

End Sub


Sub get_cols_sub()

Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
Application.AskToUpdateLinks = False

Dim cn As Object, rs As Object
Dim col_index As Integer
Dim db_path As String
Dim output_range As Range

Set main = ThisWorkbook.Worksheets("main")
database_name = main.Range("F1").Value
database = "[" & database_name & "]"
schema = "[" & main.Range("F2").Value & "]"
table_name = "[" & main.Range("F3").Value & "]"

table_full = database & "." & schema & "." & table_name

Dim cnn As ADODB.connection
Set cnn = New ADODB.connection

cnn.Provider = "SQLOLEDB.1" '<~~ mssql
user = "user"
pw = user & "0"

' Open a connection using an ODBC DSN.
cnn.ConnectionString = "driver={SQL Server};" & "server=some_server_name;uid=" & user & ";pwd=" & pw & ";database=" & database_name & ";Trusted_Connection=True;"
cnn.Open

If cnn.State = adStateOpen Then
Else
    MsgBox "Connection to SQL server could not be established. Exiting..."
    Exit Sub
End If

Set rs = CreateObject("ADODB.Recordset")
rs.Open "SELECT TOP 1 * FROM " & table_full, cnn, , , adCmdText

For intColIndex = 0 To rs.Fields.Count - 1
    main.Range("F4").Offset(0, intColIndex).Value = rs.Fields(intColIndex).Name
Next

'close and destroy objects; for freeing memory
Set rs = Nothing
cnn.Close
Set cnn = Nothing

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
Application.AskToUpdateLinks = True

End Sub


Sub get_tables_sub()

Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
Application.AskToUpdateLinks = False

Dim cn As Object, rs As Object
Dim col_index As Integer
Dim db_path As String
Dim output_range As Range

Set seznam = ThisWorkbook.Worksheets("list_of_tables")
database_name = "db"
database = "[" & database_name & "]"

Dim cnn As ADODB.connection
Set cnn = New ADODB.connection

cnn.Provider = "SQLOLEDB.1" '<~~ mssql
user = "user"
pw = user & "0"

' Open a connection using an ODBC DSN.
cnn.ConnectionString = "driver={SQL Server};" & "server=some_server_name;uid=" & user & ";pwd=" & pw & ";database=" & database_name & ";Trusted_Connection=True;"
cnn.Open

If cnn.State = adStateOpen Then
Else
    MsgBox "Connection to SQL server could not be established. Exiting..."
    Exit Sub
End If

Set rs = CreateObject("ADODB.Recordset")
rs.Open "SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE='BASE TABLE'", cnn, , , adCmdText

seznam.Cells.Clear

For intColIndex = 0 To rs.Fields.Count - 1
    seznam.Range("A1").Offset(0, intColIndex).Value = rs.Fields(intColIndex).Name
Next

seznam.Range("A2").CopyFromRecordset rs

'close and destroy objects; for freeing memory
Set rs = Nothing
cnn.Close
Set cnn = Nothing

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
Application.AskToUpdateLinks = True

End Sub




