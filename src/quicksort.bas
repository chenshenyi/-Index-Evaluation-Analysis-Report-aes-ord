Attribute VB_Name = "quick_sort"

Sub quickSort(list, ByVal left As Integer, ByVal right As Integer, Optional sortBy As String = "���W")
    Dim pivotIndex As Integer
    
    If left < right Then
        pivotIndex = (left + right) / 2
        pivotIndex = partition(list, left, right, pivotIndex, sortBy)
        quickSort list, left, pivotIndex-1, sortBy
        quickSort list, pivotIndex+1, right, sortBy
    End If
End Sub

Function partition(list, ByVal left As Integer, ByVal right As Integer, ByVal pivotIndex As Integer, sortBy As String) As Integer
    Dim pivotValue As Integer
    Dim storeIndex As Integer
    pivotValue = list(pivotIndex)
    swap list, pivotIndex, right ' ��pivot���쵲��
    storeIndex = left
    For i = left To right-1
        If list(i) < pivotValue And sortBy = "���W" Then
            swap list, storeIndex, i
            storeIndex = storeIndex + 1
        ElseIf list(i) > pivotValue And sortBy = "����" Then
            swap list, storeIndex, i
            storeIndex = storeIndex + 1
        End If
    Next
    swap list, right, storeIndex ' ��pivot���쥦�̫᪺�a��
    partition = storeIndex
End Function

Sub swap(list, ByVal a As Integer, ByVal b As Integer)
    ' Swap the items at indices a and b in the collection
    Dim temp As Variant
    temp = list(a)
    list(a) = list(b)
    list(b) = temp
End Sub

' Passed test
Sub test_quickSort()

    Dim list() As Variant

    list = Array(5, 3, 8, 7, 6, 2, 9, 1, 4)

    quickSort list, LBound(list), UBound(list), "���W"
    For Each i In list
        Debug.Print i
    Next

    quickSort list, LBound(list), UBound(list), "����"
    For Each i In list
        Debug.Print i
    Next

End Sub
