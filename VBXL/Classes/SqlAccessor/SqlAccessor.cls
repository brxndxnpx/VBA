Option Explicit

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' Module Summary
'''   An object made to execute database queries and return the dataset as an array.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Const AppPID_Command As String = "ADODB.Command"
Private Const AppPID_Recordset As String = "ADODB.Recordset"
Private Const AppPID_Connection As String = "ADODB.Connection"

Private Type TSQL
    ConnectionString As String
End Type

Private this As TSQL

''' Summary:
'''     The database connection string.
Public Property Get ConnectionString() As String
    ConnectionString = this.ConnectionString
End Property

Public Property Let ConnectionString(ByVal Value As String)
    this.ConnectionString = Trim(Value)
End Property


''' Summary:
'''     Executes a query to the database.
''' Parameters:
'''     ByVal QueryString: The query to execute. Could be a stored procedure or SQL text.
'''     Optional ByVal TimeOut: The query timeout in seconds. This will default to 300 seconds (5 minutes).
''' Returns:
'''     A Variant; The query results as an array.
Public Function Query(ByVal QueryString As String, Optional ByVal TimeOut As Long = 300)
    ' Check if the connection string is set
    If ConnectionString = vbNullString Then Query = Null: Exit Function

    Dim x              As Long
    Dim y              As Long
    Dim QueryResults   As Variant
    Dim OutputData     As Variant
    Dim Connection     As Object
    Dim Recordset      As Object
    
    Set Connection = CreateObject(AppPID_Connection)
    
    '   Set the command timeout and open the connection
    Connection.CommandTimeout = TimeOut
    Connection.Open ConnectionString
    
    '   Execute the query and store the results in a record set
    Set Recordset = Connection.Execute(QueryString)
    
    '   Check to see if there aren't any records; Return the header fields only
    If Recordset.BOF = True Then
        y = Recordset.Fields.Count: x = 1
        ReDim OutputData(1 To x, 1 To y)
        
        For y = 1 To Recordset.Fields.Count
            OutputData(x, y) = Recordset.Fields(y - 1).Name
        Next
    
        GoTo QueryCleanup
    End If
    
    '   Get the data from the record set
    QueryResults = Recordset.GetRows
    
    '   Store the results into an array
    y = Recordset.Fields.Count: x = UBound(QueryResults, 2) + 2
    ReDim OutputData(1 To x, 1 To y)
    
    For y = 1 To Recordset.Fields.Count
        x = 1
        OutputData(x, y) = Recordset.Fields(y - 1).Name
        For x = 2 To UBound(QueryResults, 2) + 2
            OutputData(x, y) = QueryResults(y - 1, x - 2)
        Next
    Next
    
'   Clean up the connection/objects
QueryCleanup:
    Recordset.Close: Set Recordset = Nothing
    Connection.Close: Set Connection = Nothing
    
    Query = OutputData
End Function

