Attribute VB_Name = "mod_Array"
Option Explicit

Public Function CollectionToArray(Collection As Collection) As Variant
   
   Dim tmp
   ReDim tmp(Collection.Count - 1)
   
   Dim i
   i = LBound(tmp)
   
   Dim item
   For Each item In Collection
      tmp(i) = item
      Inc i
   Next
   
   CollectionToArray = tmp
   
End Function

'Public Sub ArrayEnsureBounds(Arr)
'
''   Dim tmp_ptr&
''   MemCopy tmp_ptr, VarPtr(Arr) + 8, 4 ' resolve Variant
''   MemCopy tmp_ptr, tmp_ptr, 4               ' get arraypointer
''
''   Dim bIsNullArray As Boolean
''   bIsNullArray = (tmp_ptr = 0)
'' On Error Resume Next
'
'   Dim bIsNullArray As Boolean
'   bIsNullArray = (Not Not Arr) = 0 'use vbBug to get pointer to Arr
'
''   Rnd 1 ' catch Expression too complex error that is cause by the bug
''On Error GoTo 0
'
''   Exit Function
'
'   If bIsNullArray Then
'
'   ElseIf (UBound(Arr) - LBound(Arr)) < 0 Then
'   Else
'      Exit Function
'   End If
'
'   ReDim Arr(0)
'   ArrayEnsureBounds = True
'   Exit Function

Public Sub ArrayEnsureBounds(Arr)

On Error GoTo Array_err
  ' IsArray(Arr)=False        ->  13 - Type Mismatch
  ' [Arr has no Elements]     ->  9 - Subscript out of range
  ' ZombieArray[arr=Array()]  -> GoTo Array_new
   If UBound(Arr) - LBound(Arr) < 0 Then GoTo Array_new
Exit Sub
Array_err:
Select Case Err
    Case 9, 13
Array_new:
      ArrayDelete Arr

'   Case Else
'      Err.Raise Err.Number, "", "Error in ArrayEnsureBounds: " & Err.Description

End Select

End Sub



Public Sub ArrayAdd(Arr, Optional Element = "")
   ArrayEnsureBounds Arr
   ReDim Preserve Arr(LBound(Arr) To UBound(Arr) + 1)
   Arr(UBound(Arr)) = Element

End Sub


'Public Sub ArrayAdd(Arr As Variant, Optional element = "")
'' Is that already a Array?
'   If IsArray(Arr) Then
'      ReDim Preserve Arr(LBound(Arr) To UBound(Arr) + 1)
'
' ' VarType(Arr) = vbVariant must be
'   Else 'If VarType(Arr) = vbVariant Then
'      ReDim Arr(0)
'   End If
'
'   Arr(UBound(Arr)) = element
'
'End Sub

Public Sub ArrayRemoveLast(Arr)
   ReDim Preserve Arr(LBound(Arr) To UBound(Arr) - 1)
End Sub

Public Sub ArrayDelete(Arr)
   ReDim Arr(0)
   'Arr = Array()
   'Set Arr = Nothing
End Sub


Public Function ArrayGetLast(Arr)
ArrayEnsureBounds Arr
   ArrayGetLast = Arr(UBound(Arr))
End Function
Public Sub ArraySetLast(Arr, Element)
ArrayEnsureBounds Arr
    Arr(UBound(Arr)) = Element
End Sub
Public Sub ArrayAppendLast(Arr(), Element)
ArrayEnsureBounds Arr
    Arr(UBound(Arr)) = Arr(UBound(Arr)) & Element
End Sub


Public Function ArrayGetFirst(Arr)
ArrayEnsureBounds Arr
   ArrayGetFirst = Arr(LBound(Arr))
End Function
Public Sub ArraySetFirst(Arr, Element)
ArrayEnsureBounds Arr
    Arr(LBound(Arr)) = Element
End Sub
Public Sub ArrayAppendFirst(Arr, Element)
ArrayEnsureBounds Arr
    Arr(LBound(Arr)) = Arr(LBound(Arr)) & Element
End Sub


