Attribute VB_Name = "Module1"
Sub ticker()

  ' Set an initial variable for holding the brand name
  Dim ticker_name As String

  ' Loop through all credit card purchases
  For i = 2 To 797711

    ' Check if we are still within the same credit card brand, if we are not...
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

      ' Message Box the unique Bank Name
      MsgBox (Cells(i, 1).Value)

    End If

  Next i

End Sub

