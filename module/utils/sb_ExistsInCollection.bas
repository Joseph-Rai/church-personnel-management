Attribute VB_Name = "sb_ExistsInCollection"
Option Explicit

Public Function ExistsInCollection(col As Collection, key As Variant) As Boolean
    On Error GoTo err
    ExistsInCollection = True
    IsObject (col.Item(key))
    Exit Function
err:
    ExistsInCollection = False
End Function
