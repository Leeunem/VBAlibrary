# VBAlibrary
VBA Function library
### Find column by cell value
'Works only for exact match of the cell value, not if string is a part of the value

Public PV_col As Integer
Public traderef_col As Integer

Dim PV_str As String
Dim traderef_str As String

PV_str = " Par Value"
traderef_str = " Trade Ref #"

PV_col = WorksheetFunction.Match(PV_str, ActiveWorkbook.ActiveSheet.Range("1:1"), 0)
traderef_col = WorksheetFunction.Match(traderef_str, ActiveWorkbook.ActiveSheet.Range("1:1"), 0)
