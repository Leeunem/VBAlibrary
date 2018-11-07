Private Sub Worksheet_SelectionChange(ByVal Target As Excel.range)
'Updateby Extendoffice 20161123
 
  Application.EnableEvents = False
  With Target
  If .Address = range("L" & ActiveCell.Row).Address And ActiveCell.Row > 1 And ActiveCell.Row < 24 Then
    Select Case .Value
      Case "r"
        .Value = "D"
      Case Else
        .Value = "r"
    End Select
  End If
  End With
  
   With Target
  If .Address = range("M" & ActiveCell.Row).Address And ActiveCell.Row > 1 And ActiveCell.Row < 24 Then
    Select Case .Value
      Case "r"
        .Value = "rS"
          Case "rS"
        .Value = "rSr"
          Case "rSr"
        .Value = "rSrS"
          Case "rSrS"
        .Value = "rSrSr"
          Case "rSrSr"
        .Value = "rSrSrS"
      Case Else
        .Value = "r"
    End Select
  End If
  End With
  
   With Target
  If .Address = range("N" & ActiveCell.Row).Address And ActiveCell.Row > 1 And ActiveCell.Row < 24 Then
    Select Case .Value
      Case Else
        .Value = "L"
    End Select
  End If
  End With
  
     With Target
  If .Address = range("O" & ActiveCell.Row).Address And ActiveCell.Row > 1 And ActiveCell.Row < 24 Then
    Select Case .Value
      Case Else
        .Value = "S"
    End Select
  End If
  End With
  
  With Target
  If .Address = range("P" & ActiveCell.Row).Address And ActiveCell.Row > 1 And ActiveCell.Row < 24 Then
    Select Case .Value
      Case Else
        .Value = "Rp"
    End Select
  End If
  End With
  
  With Target
  If .Address = range("Q" & ActiveCell.Row).Address And ActiveCell.Row > 1 And ActiveCell.Row < 24 Then
    Select Case .Value
      Case Else
        .Value = "O"
    End Select
  End If
  End With
  
   With Target
  If .Address = range("S" & ActiveCell.Row).Address And ActiveCell.Row > 1 And ActiveCell.Row < 24 Then
    Select Case .Value
      Case "S"
        .Value = "D"
      Case Else
        .Value = "r"
    End Select
  End If
  End With
  
  Application.EnableEvents = True
End Sub
