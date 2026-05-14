Attribute VB_Name = "PolymorphicSort"
'@IgnoreModule AssignmentNotUsed, ProcedureNotUsed, UnassignedVariableUsage, UseMeaningfulName, MultipleDeclarations, ParameterCanBeByVal, HungarianNotation, VariableNotAssigned
'@Folder("Module")
'@ModuleDescription "Polymorphic sorting and searching."

'------------------------------------------------------------------------------
' MIT License
'
' Copyright (c) 2025 Vincent van Geerestein
'
' Permission is hereby granted, free of charge, to any person obtaining a copy
' of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights
' to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is
' furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all
' copies or substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
' SOFTWARE.
'------------------------------------------------------------------------------

Option Explicit

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Comments
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' Author: Vincent van Geerestein
' E-mail: vincent@vangeerestein.com
' Description: Polymorphic Sorting and Searching Module
' Dependency: ICompare (implements CompareDefault, CompareBinary, CompareText)
' Add-in: RubberDuck (https://rubberduckvba.com/)
' Version: 2025.10.12

' This module implements polymorphic sorting and searching using the ICompare
' interface which implements the actual value comparison. The sorting algorithm
' may not represent the fastest method under all circumstances and merely is a
' show case for using interfaces for coding polymorphic sorting.
'
' Methods
' Sort arr [, idx, method, asc]             Sorts an array
' Search(arr, val [, idx, method, start])   Searches for a value in a sorted array
' IsSorted(arr [, idx, method])             Returns True if an array is sorted
'
' Property (Get/Set)
' CompareCustom                     Sets or returns the custom ICompare interface
'
' The actual sort order can be reversed by providing an optional parameter.
' The optional idx parameter determines whether a sort or search is done in
' place or by index. The index array is automatically created by the sort
' routine and is returned sorted to the caller.
'
' ICompare provides polymorphic support by defining an interface for methods
' required for sorting and searching: Assign, Compare, Equal and Swap. The
' interface also defines three constants: ecLess, ecEqual, ecGreater. ICompare
' is implemented for default comparison by operator, as well as for Binary and
' Text comparison. When the custom ICompare interface is not set, the choice
' between various available ICompare implementations is made automatically.
' For complicated cases like sorting a specific object array, the interface
' functions need to be coded and the interface should subsequently be set with
' Set CompareCustom. Polymorphic sorting can't be directly performed on an
' array of UDT's.
'
' The DualPivotQuickSort algorithm is used for sorting. Sub-arrays shorter
' than a threshold length are sorted by an insertion sort algorithm, see:
' http://hg.openjdk.java.net/jdk8/jdk8/jdk/file/tip/src/share/classes/java/util/DualPivotQuicksort.java

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Private declarations
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' Threshold for applying InsertionSort (taken from Java).
Private Const INSERTION_SORT_THRESHOLD As Long = 47

' Selected VB errors.
Private Enum VBERROR
    vbErrorInvalidProcedureCall = 5
    vbErrorSubscriptOutOfRange = 9
    vbErrorTypeMismatch = 13
    vbErrorCantPerformRequestedOperation = 17
End Enum

' Wrapper for private data.
Private Type TPRIVATE
    CompareCustom As ICompare
End Type
Private this As TPRIVATE


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Public methods
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'@Description "Gets or sets the compare interface."
Public Property Get CompareCustom() As ICompare

    Set CompareCustom = this.CompareCustom

End Property
Public Property Set CompareCustom(ByVal RHS As ICompare)

    Set this.CompareCustom = RHS

End Property


'@Description "Sorts an array either in place or by index."
Public Sub Sort( _
    ByRef arr As Variant, _
    Optional ByRef idx As Variant, _
    Optional ByVal method As VbCompareMethod, _
    Optional ByVal asc As Boolean = True _
)

    ' Seed the random-number generator.
    Static Seed As Boolean
    If Seed = False Then Randomize: Seed = True

    ' The array must not be empty and must be one dimensional.
    If IsVector(arr) = False Then Err.Raise vbErrorTypeMismatch

    ' Set the ICompare interface.
    Dim Compare As ICompare: Set Compare = CompareInterface(arr, method)

    ' Set the sort order.
    Dim Order As ECompare: Order = VBA.IIf(asc, ecLess, ecGreater)

    ' Perform the sort either in place or by index.
    If VBA.IsMissing(idx) Then
        DualQuickSortInPlace arr, LBound(arr), UBound(arr), Order, Compare
    Else
        idx = CreateIndexArray(arr)
        DualQuickSortByIndex arr, idx, LBound(arr), UBound(arr), Order, Compare
    End If

End Sub


'@Description "Searches for a value in a sorted array."
Public Function Search( _
    ByRef arr As Variant, _
    ByVal value As Variant, _
    Optional ByRef idx As Variant, _
    Optional ByVal method As VbCompareMethod, _
    Optional ByVal Start As Variant _
) As Variant
' Returns Null if the search value is not found.

' Looks for next value at start+1 if start is provided. If the other parameters
' have changed since the original search unpredictable results will be obtained.

    ' Save helper for when start is provided.
    Static helper As ICompare

    ' The array must not be empty and must be one dimensional.
    If IsVector(arr) = False Then Err.Raise vbErrorTypeMismatch

    ' Perform the binary search either in place or by index.
    If VBA.IsMissing(idx) Then
        If VBA.IsMissing(Start) Then
            Set helper = CompareInterface(arr, method)
            Search = BinarySearchInPlace(arr, value, helper)
        ElseIf helper Is Nothing = False Then
            Search = NextValueInPlace(arr, value, helper, Start)
        Else
            Err.Raise vbErrorInvalidProcedureCall, , "Start is invalid"
        End If
    ElseIf IsIndexArray(idx, arr) Then
        If VBA.IsMissing(Start) Then
            Set helper = CompareInterface(arr, method)
            Search = BinarySearchByIndex(arr, value, idx, helper)
        ElseIf helper Is Nothing = False Then
            Search = NextValueByIndex(arr, value, idx, helper, Start)
        Else
            Err.Raise vbErrorInvalidProcedureCall, , "Start is invalid"
        End If
    Else
        Err.Raise vbErrorInvalidProcedureCall, , "Index is invalid"
    End If

    If VBA.IsNull(Search) Then Set helper = Nothing

End Function


'@Description "Returns True if an array is sorted or False otherwise."
Public Function IsSorted( _
    ByRef arr As Variant, _
    Optional ByRef idx As Variant, _
    Optional ByVal method As VbCompareMethod _
) As Boolean

    ' The array must not be empty and must be one dimensional.
    If IsVector(arr) = False Then Err.Raise vbErrorTypeMismatch

    ' Perform the check either in place or by index.
    If VBA.IsMissing(idx) Then
        IsSorted = IsSortedInPlace(arr, CompareInterface(arr, method))
    ElseIf IsIndexArray(idx, arr) Then
        IsSorted = IsSortedbyIndex(arr, idx, CompareInterface(arr, method))
    Else
        Err.Raise vbErrorInvalidProcedureCall , , "Index is invalid"
    End If

End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Private methods
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'@Description "Searches for a value in an sorted array."
Private Function BinarySearchInPlace( _
    ByRef arr As Variant, _
    ByVal value As Variant, _
    ByVal helper As ICompare _
) As Variant
With helper

    Dim Lower As Long: Lower = LBound(arr)
    Dim upper As Long: upper = UBound(arr)

    ' Determine the order in the array (ascending or descending).
    Dim Order As ECompare
    Order = .Compare(arr(Lower), arr(upper))
    Do
        Dim middle As Long: middle = (Lower + upper) \ 2
        Select Case .Compare(value, arr(middle))
        Case ecEqual
            ' Return the lowest index for which arr(i) = value
            Dim i As Long
            For i = middle - 1 To Lower Step -1
                If .Compare(value, arr(i)) <> ecEqual Then
                    Exit For
                End If
            Next
            BinarySearchInPlace = i + 1
            Exit Function
        Case Order
            upper = middle - 1
        Case Else
            Lower = middle + 1
        End Select
    Loop Until Lower > upper

    BinarySearchInPlace = Null

End With
End Function


'@Description "Searches for a value in an array sorted by index."
Private Function BinarySearchByIndex( _
    ByRef arr As Variant, _
    ByVal value As Variant, _
    ByRef idx As Variant, _
    ByVal helper As ICompare _
) As Variant
With helper

    Dim Lower As Long: Lower = LBound(idx)
    Dim upper As Long: upper = UBound(idx)

    ' Determine the order in the array (ascending or descending).
    Dim Order As ECompare
    Order = .Compare(arr(idx(Lower)), arr(idx(upper)))
    Do
        Dim middle As Long: middle = (Lower + upper) \ 2
        Select Case .Compare(value, arr(idx(middle)))
        Case ecEqual
            ' Return the lowest index for which arr(i) = value
            Dim i As Long
            For i = middle - 1 To Lower Step -1
                If .Compare(value, arr(idx(i))) <> ecEqual Then
                    Exit For
                End If
            Next
            BinarySearchByIndex = i + 1
            Exit Function
        Case Order
            upper = middle - 1
        Case Else
            Lower = middle + 1
        End Select
    Loop Until Lower > upper

    BinarySearchByIndex = Null

End With
End Function


'@Description "Gives the next matching value in an sorted array."
Private Function NextValueInPlace( _
    ByRef arr As Variant, _
    ByVal value As Variant, _
    ByVal helper As ICompare, _
    ByVal Start As Long _
) As Variant
With helper

    If Start + 1 <= UBound(arr) Then
        If .Compare(value, arr(Start + 1)) = ecEqual Then
            NextValueInPlace = Start + 1
            Exit Function
        End If
    End If

    NextValueInPlace = Null

End With
End Function


'@Description "Gives the next matching value in an sorted array by index."
Private Function NextValueByIndex( _
    ByRef arr As Variant, _
    ByVal value As Variant, _
    ByRef idx As Variant, _
    ByVal helper As ICompare, _
    ByVal Start As Long _
) As Variant
With helper

    If Start + 1 <= UBound(idx) Then
        If .Compare(value, arr(idx(Start + 1))) = ecEqual Then
            NextValueByIndex = Start + 1
            Exit Function
        End If
    End If

    NextValueByIndex = Null

End With
End Function


'@Description "Sorts a (sub)array."
Private Sub DualQuickSortInPlace( _
    ByRef arr As Variant, _
    ByVal Lower As Long, _
    ByVal upper As Long, _
    ByVal Order As ECompare, _
    ByVal helper As ICompare _
)
With helper

    ' Part 0: stop the recursion and sort the remaining subarray by an insertion sort.
    Dim length As Long: length = upper - Lower + 1

    If length <= 2 Then
        If length = 2 Then
            If helper.Compare(arr(Lower + 1), arr(Lower)) = Order Then
                helper.Swap arr(Lower), arr(Lower + 1)
            End If
        End If
        Exit Sub
    End If

    If length < INSERTION_SORT_THRESHOLD Then
        InsertionSortInPlace arr, Lower, upper, Order, helper
        Exit Sub
    End If

    ' Part 1: randomly select the left and right pivots.
    Dim LowerPivot As Long, UpperPivot As Long
    LowerPivot = Lower + 1 + Int(Rnd * (length \ 3))
    UpperPivot = upper - 1 - Int(Rnd * (length \ 3))

    If .Compare(arr(LowerPivot), arr(UpperPivot)) = Order Then
        .Swap arr(LowerPivot), arr(Lower)
        .Swap arr(UpperPivot), arr(upper)
    Else
        .Swap arr(LowerPivot), arr(upper)
        .Swap arr(UpperPivot), arr(Lower)
    End If

    Dim LowerPivotValue As Variant, UpperPivotValue As Variant
    .Assign LowerPivotValue, arr(Lower)
    .Assign UpperPivotValue, arr(upper)

    ' Part 2: partition the array.
    Dim LessThan As Long, GreaterThan As Long
    LessThan = Lower + 1
    GreaterThan = upper - 1

    Dim i As Long: i = LessThan
    Do While i <= GreaterThan
        If .Compare(arr(i), LowerPivotValue) = Order Then
            ' Elements < left pivot.
            .Swap arr(i), arr(LessThan)
            LessThan = LessThan + 1
        ElseIf .Compare(UpperPivotValue, arr(i)) = Order Then
            ' Elements > right pivot.
            For GreaterThan = GreaterThan To i + 1 Step -1
                If .Compare(UpperPivotValue, arr(GreaterThan)) <> Order Then
                    Exit For
                End If
            Next
            .Swap arr(i), arr(GreaterThan)
            GreaterThan = GreaterThan - 1
            If .Compare(arr(i), LowerPivotValue) = Order Then
                .Swap arr(i), arr(LessThan)
                LessThan = LessThan + 1
            End If
        End If
        i = i + 1
    Loop

    .Swap arr(Lower), arr(LessThan - 1)
    .Swap arr(upper), arr(GreaterThan + 1)

    ' Part 3: sort the three partitions.
    ' Left partition.
    DualQuickSortInPlace arr, Lower, LessThan - 2, Order, helper
    ' Right partition.
    DualQuickSortInPlace arr, GreaterThan + 2, upper, Order, helper

    ' If center part is too large (> 4/7 of the length) swap internal pivot values to ends.
    If (7 * (GreaterThan - LessThan)) > (4 * length) Then
        ' Process possibly equal elements.
        If Not .Equal(LowerPivotValue, UpperPivotValue) Then
            i = LessThan
            Do While i <= GreaterThan
                If .Equal(arr(i), LowerPivotValue) Then
                    .Swap arr(i), arr(LessThan)
                    LessThan = LessThan + 1
                ElseIf .Equal(UpperPivotValue, arr(i)) Then
                    .Swap arr(i), arr(GreaterThan)
                    GreaterThan = GreaterThan - 1
                    If .Equal(arr(i), LowerPivotValue) Then
                        .Swap arr(i), arr(LessThan)
                        LessThan = LessThan + 1
                    End If
                End If
                i = i + 1
            Loop
        End If
    End If

    If .Compare(LowerPivotValue, UpperPivotValue) = Order Then
    ' Center partition.
        DualQuickSortInPlace arr, LessThan, GreaterThan, Order, helper
    End If

End With
End Sub


'@Description "Sorts a (sub)array by index."
Private Sub DualQuickSortByIndex( _
    ByRef arr As Variant, _
    ByRef idx As Variant, _
    ByVal Lower As Long, _
    ByVal upper As Long, _
    ByVal Order As ECompare, _
    ByVal helper As ICompare _
)
With helper

    ' Part 0: stop the recursion and sort the remaining subarray by an insertion sort.
    Dim length As Long: length = upper - Lower + 1

    If length <= 2 Then
        If length = 2 Then
            Dim x As Long
            If helper.Compare(arr(idx(Lower + 1)), arr(idx(Lower))) = Order Then
                x = idx(Lower): idx(Lower) = idx(Lower + 1): idx(Lower + 1) = x
            End If
        End If
        Exit Sub
    End If

    If length < INSERTION_SORT_THRESHOLD Then
        InsertionSortByIndex arr, idx, Lower, upper, Order, helper
        Exit Sub
    End If

    ' Part 1: randomly select the left and right pivots.
    Dim LowerPivot As Long, UpperPivot As Long
    LowerPivot = Lower + 1 + Int(Rnd * (length \ 3))
    UpperPivot = upper - 1 - Int(Rnd * (length \ 3))

    If .Compare(arr(idx(LowerPivot)), arr(idx(UpperPivot))) = Order Then
        x = idx(LowerPivot): idx(LowerPivot) = idx(Lower): idx(Lower) = x
        x = idx(UpperPivot): idx(UpperPivot) = idx(upper): idx(upper) = x
    Else
        x = idx(LowerPivot): idx(LowerPivot) = idx(upper): idx(upper) = x
        x = idx(UpperPivot): idx(UpperPivot) = idx(Lower): idx(Lower) = x
    End If

    Dim LowerPivotValue As Variant, UpperPivotValue As Variant
    .Assign LowerPivotValue, arr(idx(Lower))
    .Assign UpperPivotValue, arr(idx(upper))

    ' Part 2: partition the array.
    Dim LessThan As Long, GreaterThan As Long
    LessThan = Lower + 1
    GreaterThan = upper - 1

    Dim i As Long: i = LessThan
    Do While i <= GreaterThan
        If .Compare(arr(idx(i)), LowerPivotValue) = Order Then
            ' Elements < left pivot.
            x = idx(i): idx(i) = idx(LessThan): idx(LessThan) = x
            LessThan = LessThan + 1
        ElseIf .Compare(UpperPivotValue, arr(idx(i))) = Order Then
            ' Elements > right pivot.
            For GreaterThan = GreaterThan To i + 1 Step -1
                If .Compare(UpperPivotValue, arr(idx(GreaterThan))) <> Order Then
                    Exit For
                End If
            Next
            x = idx(i): idx(i) = idx(GreaterThan): idx(GreaterThan) = x
            GreaterThan = GreaterThan - 1
            If .Compare(arr(idx(i)), LowerPivotValue) = Order Then
                x = idx(i): idx(i) = idx(LessThan): idx(LessThan) = x
                LessThan = LessThan + 1
            End If
        End If
        i = i + 1
    Loop

    x = idx(Lower): idx(Lower) = idx(LessThan - 1): idx(LessThan - 1) = x
    x = idx(upper): idx(upper) = idx(GreaterThan + 1): idx(GreaterThan + 1) = x

    ' Part 3: sort the three partitions.
    ' Left partition.
    DualQuickSortByIndex arr, idx, Lower, LessThan - 2, Order, helper
     ' Right partition.
    DualQuickSortByIndex arr, idx, GreaterThan + 2, upper, Order, helper

    ' If center part is too large (> 4/7 of the length) swap internal pivot values to ends.
   If (7 * (GreaterThan - LessThan)) > (4 * length) Then
        ' Process possibly equal elements.
        If Not .Equal(LowerPivotValue, UpperPivotValue) Then
            i = LessThan
            Do While i <= GreaterThan
                If .Equal(arr(idx(i)), LowerPivotValue) Then
                    x = idx(i): idx(i) = idx(LessThan): idx(LessThan) = x
                    LessThan = LessThan + 1
                ElseIf .Equal(UpperPivotValue, arr(idx(i))) Then
                    x = idx(i): idx(i) = idx(GreaterThan): idx(GreaterThan) = x
                    GreaterThan = GreaterThan - 1
                    If .Equal(arr(idx(i)), LowerPivotValue) Then
                        x = idx(i): idx(i) = idx(LessThan): idx(LessThan) = x
                        LessThan = LessThan + 1
                    End If
                End If
                i = i + 1
            Loop
        End If
    End If

    ' Center partition.
    If .Compare(LowerPivotValue, UpperPivotValue) = Order Then
        DualQuickSortByIndex arr, idx, LessThan, GreaterThan, Order, helper
    End If

End With
End Sub


'@Description "Sorts a (sub)array."
Private Sub InsertionSortInPlace( _
    ByRef arr As Variant, _
    ByVal Lower As Long, _
    ByVal upper As Long, _
    ByVal Order As ECompare, _
    ByVal helper As ICompare _
)
With helper
    Dim i As Long, j As Long, value As Variant
    For i = Lower + 1 To upper
        .Assign value, arr(i)
        For j = i To Lower + 1 Step -1
            If .Compare(value, arr(j - 1)) <> Order Then Exit For
            .Assign arr(j), arr(j - 1)
        Next
        .Assign arr(j), value
    Next

End With
End Sub


'@Description "Sorts a (sub)array by index."
Private Sub InsertionSortByIndex( _
    ByRef arr As Variant, _
    ByRef idx As Variant, _
    ByVal Lower As Long, _
    ByVal upper As Long, _
    ByVal Order As ECompare, _
    ByVal helper As ICompare _
)
With helper
    Dim i As Long, j As Long, k As Long, value As Variant
    For i = Lower + 1 To upper
        k = idx(i)
        .Assign value, arr(k)
        For j = i To Lower + 1 Step -1
            If .Compare(value, arr(idx(j - 1))) <> Order Then Exit For
            idx(j) = idx(j - 1)
        Next
        idx(j) = k
    Next

End With
End Sub


'@Description "Returns True if an array is sorted or False otherwise."
Private Function IsSortedInPlace( _
    ByRef arr As Variant, _
    ByVal helper As ICompare _
) As Boolean
With helper

    Dim Lower As Long: Lower = LBound(arr)
    Dim upper As Long: upper = UBound(arr)

    ' Determine the order in the array (ascending or descending).
    Dim Order As ECompare: Order = .Compare(arr(Lower), arr(upper))
    Dim i As Long, Current As Variant, Previous As Variant
    If Order = ecEqual Then
        ' The array is only ordered when all elements have the same value.
        .Assign Current, arr(Lower)
        For i = Lower + 1 To upper
            If .Equal(arr(i), Current) = False Then Exit Function
        Next
    Else
        ' Check the order for all array elements.
        .Assign Previous, arr(Lower)
        For i = Lower + 1 To upper
            .Assign Current, arr(i)
            If .Compare(Current, Previous) = Order Then Exit Function
            .Assign Previous, Current
        Next
    End If
    IsSortedInPlace = True

End With
End Function


'@Description "Returns True if an array is sorted or False otherwise."
Private Function IsSortedbyIndex( _
    ByRef arr As Variant, _
    ByRef idx As Variant, _
    ByVal helper As ICompare _
) As Boolean
With helper

    Dim Lower As Long: Lower = LBound(idx)
    Dim upper As Long: upper = UBound(idx)

    ' Determine the order in the array (ascending or descending).
    Dim Order As ECompare: Order = .Compare(arr(idx(Lower)), arr(idx(upper)))
    Dim i As Long, Current As Variant, Previous As Variant
    If Order = ecEqual Then
        ' The array is only ordered when all elements have the same value.
        .Assign Current, arr(idx(Lower))
        For i = Lower + 1 To upper
            If .Equal(arr(idx(i)), Current) = False Then Exit Function
        Next
    Else
        ' Check the order for all array elements via the index array.
        .Assign Previous, arr(idx(Lower))
        For i = Lower + 1 To upper
            .Assign Current, arr(idx(i))
            If .Compare(Current, Previous) = Order Then Exit Function
            .Assign Previous, Current
        Next
    End If
    IsSortedbyIndex = True

End With
End Function


'@Description "Sets the (automatic) ICompare interface."
Private Function CompareInterface( _
    ByRef arr As Variant, _
    Optional ByVal method As VbCompareMethod _
) As ICompare

    If this.CompareCustom Is Nothing = False Then
        ' Explicit setting using the custom ICompare implementation.
        Set CompareInterface = this.CompareCustom
    ElseIf IsNumericArray(arr) Then
        Set CompareInterface = New CompareDefault
    ElseIf IsStringArray(arr) Then
        Select Case method
        Case vbBinaryCompare
            Set CompareInterface = New CompareBinary
        Case vbTextCompare
            Set CompareInterface = New CompareText
        Case Else
            Set CompareInterface = New CompareDefault
        End Select
    Else
        ' Provide custom ICompare implementation or expand the code.
        Err.Raise vbErrorCantPerformRequestedOperation, , "Custom compare not implemented"
    End If

End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Private utility methods
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'@Description "Returns True if an array is empty or False otherwise."
Private Function IsEmptyArray(ByRef arr As Variant) As Boolean

    On Error GoTo ForcedError
    IsEmptyArray = LBound(arr) > UBound(arr)
    Exit Function

ForcedError:
    If Err.Number <> vbErrorSubscriptOutOfRange Then Err.Raise Err.Number
    IsEmptyArray = True
    Err.Clear

End Function


'@Description "Returns True if a variable is an empty array or False otherwise."
Private Function IsArrayEmpty(ByRef var As Variant) As Boolean

    If VBA.IsArray(var) = False Then Exit Function

    On Error Resume Next
    Dim Lo As Long: Lo = LBound(var)
    Dim Hi As Long: Hi = UBound(var)
    If Err.Number = 0 Then
        IsArrayEmpty = (Hi < Lo)
    Else
        IsArrayEmpty = True
    End If
    On Error GoTo 0

End Function


'@Description "Returns the number of array dimensions of a variable."
Private Function ArrayNDims(ByRef var As Variant) As Long
' An allocated but empty first dimension returns 0.

    If VBA.IsArray(var) = False Then Exit Function

    Dim NDims As Long
    Dim Lo As Long, Hi As Long
    On Error GoTo ForcedError
    Do
        NDims = NDims + 1
        Lo = LBound(var, NDims)
        Hi = UBound(var, NDims)
    Loop While Lo <= Hi

ForcedError:
    Err.Clear
    On Error GoTo 0

    ArrayNDims = NDims - 1

End Function


'@Description "Returns True if a variable is a vector or False otherwise."
Private Function IsVector(ByRef var As Variant) As Boolean
' A vector is an one-dimensional non-empty array of any type.

    If VBA.IsArray(var) = False Then Exit Function

    Dim NDims As Long
    Dim Lo As Long, Hi As Long
    On Error GoTo ForcedError
    Do
        Lo = LBound(var, NDims + 1)
        Hi = UBound(var, NDims + 1)
        NDims = NDims + 1
    Loop While Lo <= Hi And NDims <= 1

ForcedError:
    If Err.Number <> 0 Then
        Err.Clear
        IsVector = (NDims = 1)
    End If
    On Error GoTo 0

End Function


'@Description "Creates an index array."
Private Function CreateIndexArray( _
    ByRef arr As Variant, _
    Optional ByVal base As Long _
) As Long()

    Dim idx() As Long: ReDim idx(base To base + UBound(arr) - LBound(arr))
    Dim offset As Long: offset = LBound(arr) - base
    Dim i As Long
    For i = LBound(idx) To UBound(idx)
        idx(i) = offset + i
    Next

    CreateIndexArray = idx

End Function


'@Description "Returns True if an array is a valid index array or False otherwise."
Private Function IsIndexArray( _
    ByRef idx As Variant, _
    ByRef arr As Variant _
) As Boolean
' Check the integrity of an index array and check for consistancy with the indexed array.

    If IsVector(idx) = False Then Exit Function

    Dim Lower As Long: Lower = LBound(arr)
    Dim upper As Long: upper = UBound(arr)

    ' Check whether the array lengths match.
    If UBound(idx) - LBound(idx) <> upper - Lower Then Exit Function

    ' Check whether all indices are in range and used only once.
    Dim Used() As Boolean: ReDim Used(Lower To upper)
    Dim i As Long, index As Long
    For i = LBound(idx) To UBound(idx)
        index = idx(i)
        If index < Lower Or index > upper Then Exit Function
        If Used(index) Then Exit Function
        Used(index) = True
    Next

    IsIndexArray = True

End Function


'@Description "Returns True if an array contains numeric values only or False otherwise."
Private Function IsNumericArray(ByRef arr As Variant) As Boolean
' This function also works on multidimensional arrays.

    Select Case VBA.VarType(arr)
    Case VBA.vbVariant Or VBA.vbArray, VBA.vbString Or VBA.vbArray
        Dim Item As Variant
        For Each Item In arr
            If VBA.IsNumeric(Item) = False Then Exit Function
        Next
        IsNumericArray = True
    Case VBA.vbObject Or VBA.vbArray, VBA.vbUserDefinedType Or VBA.vbArray
        IsNumericArray = False
    Case Else
        IsNumericArray = True
    End Select

End Function


'@Description "Returns True if an array contains strings only or False otherwise."
Private Function IsStringArray(ByRef arr As Variant) As Boolean
' This function also works for multidimensional arrays.

    Select Case VBA.VarType(arr)
    Case VBA.vbVariant Or VBA.vbArray
        Dim Item As Variant
        For Each Item In arr
            If VBA.VarType(Item) <> VBA.vbString Then Exit Function
        Next
        IsStringArray = True
    Case VBA.vbString Or VBA.vbArray
        IsStringArray = True
    End Select

End Function
