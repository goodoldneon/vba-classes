# VBA Classes


## RecordsetToRange

*Prints recordsets as Excel ranges, with optional styling.*

### Methods

#### `create(rs, rng)`
Prints the recordset (`rs`) into the range (`rng`).
* `rs`
  * Type: ADODB.Recordset
* `rng`
  * Type: `Range`
  * Top left cell of the desired output range. Must be a single cell (e.g. `Range("A1")`).

#### `resetToDefaults()`
Reset settings back to defaults.

#### `styleHeader(rng)`
Styles the headers.
* `rng`
  * Type: `Range`
  * One or more cells.

#### `styleTitle(rng)`
Styles the title bar. The title bar is above the headers and spans the entire width of the table.
* `rng`
  * Type: `Range`
  * One or more cells.

#### `styleBorder(rng)`
Styles the borders.
* `rng`
  * Type: `Range`
  * One or more cells.

### Settings

#### `border`
Whether borders (outer and inner) should be rendered.
* Type: `Boolean`

#### `headerLeft`
Whether the first column should be styled as headers.
* Type: `Boolean`

#### `headerLeftAlign`
If first column is rendered as headers, then this controls how its text should be aligned.
* type: `String`
* Accepts:
  * `"left"`
  * `"center"`
  * `"right"`

#### `headerTop`
Whether the first row should be styled as headers.
* Type: `Boolean`

#### `headerToptAlign`
If first row is rendered as headers, then this controls how its text should be aligned.
* type: `String`
* Accepts:
  * `"left"`
  * `"center"`
  * `"right"`

#### `title`
Creates a title bar that spans the width of the recordset.
* type: `String`

### Example
The following example requires the "Sql" class in this repo.
```vba
Sub sqlTest()
    Dim sqlObj As Sql
    Set sqlObj = New Sql
    Dim recordsetToRangeObj As RecordsetToRange
    Set recordsetToRangeObj = New RecordsetToRange
    Dim rs As ADODB.Recordset
    Dim query As String

    'Query that will be run.
    query = "select count(*) as Count from USSGLW.dbo.bop_ht"

    'Set connection string.
    sqlObj.connStr = "Steel Server"

    'If a connection can be made...
    If sqlObj.isConnValid Then
        'Then run the query and store the results in the rs (recordset) variable.
        Set rs = sqlObj.runQuery(query)
    End If
    
    With recordsetToRangeObj
        .headerTop = True
        Call .create(rs, Range("A1"))
    End With
    
    'Close the recordset and connection.
    rs.Close
End Sub
```


## Sql

### Methods

### Settings

### Example