# Master Data Mapping

## Database Columns Name Mapping Java Entity Attribute Name(Lower Camel Case)

### API Response: DB -> entity
- Can get in Mybatis.xml
- **Note: If used @JsonProperty(value="") to modify the property name, can only manually process the correction**

### API Request: entity -> DB

API /internal/psg/search-results:
- requestBody
  - psgSubmitDateFrom -> no mapping , using for logic search

## DB Table Mapping API End Point

- Manual in Mater Data Excel **[Database Operation Table]** column
- Mapping in DB Schema Excel Sheet Name

## (Deprecated) Master Data [Database Operation Table] column multi-options setting

1. select cell where need to apply
2. Menu bar -> Data tab -> Data Tools group -> Data Validation button
3. source field input `='DB-TABLE'!$A$1:$A$1000`, will get data from sheet 'DB-TABLE'
4. above only support select single option
5. implement mutil-options can code with VBA in active sheet

```shell
Private Sub Worksheet_Change(ByVal Target As Range)

'code runs on protected sheet
Dim oldVal As String
Dim newVal As String
Dim strSep As String
Dim strType As Long

'add comma and space between items
strSep = ", "

If Target.Count > 1 Then GoTo exitHandler

'checks validation type of target cell
'type 3 is a drop down list
On Error Resume Next
strType = Target.Validation.Type

'19 is Database Operation Table column count
'
If Target.Column = 19 And strType = 3 Then
  Application.EnableEvents = False
  newVal = Target.Value
  Application.Undo
  oldVal = Target.Value
  If oldVal = "" Or newVal = "" Then
    Target.Value = newVal
  Else
    Target.Value = oldVal _
      & strSep & newVal
  End If
End If

exitHandler:
  Application.EnableEvents = True
End Sub


```