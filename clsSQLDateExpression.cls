VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSQLDateExpression"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Useful functions for abstracting date formats for different database engines

Private Function StringField(theString As String, separator As String, fieldNumber As Integer) As String
'Auxiliar function for parsing the strings
'Returns the field number "fieldNumber" from the string "theString"
'We number fields from 1 and on
'Example:
'       theString : "20/02/06"
'       separator : "/"
'       fields : 1) "20"
'                2) "02"
'                3) "06"
'                4) ""

Dim pos1 As Integer, pos2 As Integer
Dim i As Integer
Dim strField As String

'we try to skip "fieldNumber - 1" separators
pos1 = 1
pos2 = InStr(pos1, theString, separator)
i = 1
While (i < fieldNumber And pos2 > 0)
    pos1 = pos2 + 1
    pos2 = InStr(pos1, theString, separator)
    i = i + 1
Wend

'let's see what happened
If i < fieldNumber Then
    'there weren't so many fields
    strField = ""
Else
    If pos2 = 0 Then
        'our field is the last one
        strField = Mid(theString, pos1)
    Else
        'our field is between pos1 and pos2
        strField = Mid(theString, pos1, pos2 - pos1)
    End If
End If

StringField = Trim(strField)

End Function


Public Function ValidateDate(theString As String, separador As String, strFormat As String, ByRef myDate As Date) As Boolean
'It validates a date string according to the format
'If theString contains a valid date, it returns true, otherwise, it returns false
'if it returns true, the byref parameter "myDate" returns a date datatype variable
'  with the converted date
Dim strISODate As String
Dim strYear As String, strMonth As String, strDay As String
Dim strFieldFormat As String
Dim i As Integer

'let's look for each date component according to the input format
For i = 1 To 3
    strFieldFormat = StringField(strFormat, "/", i)
    Select Case strFieldFormat
        Case "yy"
            strYear = StringField(theString, separador, i)
        Case "mm"
            strMonth = StringField(theString, separador, i)
        Case "dd"
            strDay = StringField(theString, separador, i)
    End Select
Next

If strYear = "" Then
    'if we haven't got a year, asume current
    strYear = format("1/1", "yyyy")
ElseIf Len(strYear) < 4 Then
    'complete 4 digits for year
    strYear = format("1/1/" & strYear, "yyyy")
End If

'we now build a "yyyy/mm/dd" string, the only string that is not ambiguate
strISODate = strYear & "/" & strMonth & "/" & strDay

If (IsDate(strISODate)) Then
    'it's a date
    myDate = CDate(strISODate)
    ValidateDate = True
Else
    'no valid date
    ValidateDate = False
End If

End Function

Public Function DateToSQLExpression(myDate As Date, DBType As String) As String
'It receives a date variable and returns a string that you can use in an SQL
'  sentence (like INSERT ... values() or SELECT ... where
'it receives the database engine type in DBType
'    ("Access","SQLServer","Oracle","Informix" or "MySQL")
    
    Dim cadExpr As String
    If (DBType = "SQLServer") Then
        cadExpr = "CONVERT(datetime,'" + format(myDate, "dd/mm/yyyy") + "', 103)"
    ElseIf (DBType = "Oracle") Then
        cadExpr = "to_date('" + format(myDate, "dd/mm/yyyy") + "','DD/MM/YYYY')"
    ElseIf (DBType = "MySQL") Then
        cadExpr = "'" + format(myDate, "yyyy/mm/dd") + "'"
    ElseIf (DBType = "Access") Then
        cadExpr = "#" + format(myDate, "yyyy/mm/dd") + "#"
    Else
        cadExpr = "'" + format(myDate, "yyyy/mm/dd") + "'"
    End If
    
    DateToSQLExpression = cadExpr
End Function

