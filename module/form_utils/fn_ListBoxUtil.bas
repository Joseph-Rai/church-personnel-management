Attribute VB_Name = "fn_ListBoxUtil"
Option Explicit

Function CountSelectedItems(listBox As MSForms.listBox)

    Dim result As Integer
    Dim i As Integer
    For i = 0 To listBox.ListCount - 1
        If listBox.Selected(i) = True Then
            result = result + 1
        End If
    Next
    
    CountSelectedItems = result
End Function

Function SelectAllItems(listBox As MSForms.listBox, blnSelect As Boolean)

    Dim i As Integer
    For i = 0 To listBox.ListCount - 1
        listBox.Selected(i) = blnSelect
    Next
    
End Function
