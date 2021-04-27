# SqlAccessor

A late bound object made to execute simple database queries and return the dataset as an array.
- Uses the `ADODB.Command`, `ADODB.Recordset`, and `ADODB.Connection` objects.
    - See MSDN Documentation [here](https://docs.microsoft.com/en-us/sql/ado/guide/data/creating-and-executing-a-simple-command?view=sql-server-ver15).

---

## Properties

| Property           | Description                     |
|--------------------|---------------------------------|
| `ConnectionString` | The database connection string. |


## Methods & Functions

|                   | Description                       |
|-------------------|-----------------------------------|
| [`Query`](#query) | Executes a query to the database. |

---

### [`Query`](SqlAccessor.cls#L36)

Executes a query to the database.
- The `ConnectionString` must be set prior to executing a query.
- The first record in the resulting array are the headers.

**Parameters**
- `Source` `ByVal`
    - The query to execute. Could be a stored procedure or SQL text.
- `TimeOut` `ByVal` [`Optional`]
    - The query timeout in seconds. This will default to 600 seconds (10 minutes).

**Returns**
- `Variant`: The query results as an array. The first record in the resulting array are the headers.

---

## Usage

The example below demonstrates how to execute a query with this object.

```vb
Private Sub Demo()
    Dim sql As New SqlAccessor
    Dim results As Variant

    ' A dummy connection string
    sql.ConnectionString = "DRIVER=SQL Server; UID=ExampleUsername; " & _
        "PWD=ExamplePassword; SERVER=ExampleServer; DATABASE=ExampleDatabase"

    ' Execute the query
    results = sql.Query("SELECT TOP 10 * FROM dbo.ExampleTable", 60)
End Sub
```
