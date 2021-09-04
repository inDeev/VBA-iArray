VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "iArray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'MIT License
'
'Copyright (c) 2019-2021 Petr Katerinak
'
'Permission is hereby granted, free of charge, to any person obtaining a copy
'of this software and associated documentation files (the "Software"), to deal
'in the Software without restriction, including without limitation the rights
'to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
'copies of the Software, and to permit persons to whom the Software is
'furnished to do so, subject to the following conditions:
'
'The above copyright notice and this permission notice shall be included in all
'copies or substantial portions of the Software.
'
'THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
'IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
'FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
'AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
'LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
'OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
'SOFTWARE.
'
' Version of iArray: 1.0

Private iArray As Collection
Private i As Long ' Counter

' #####################
' #    CONSTRUCTOR    #
' #####################
Private Sub Class_Initialize()
  Set iArray = New Collection
End Sub

' #####################
' #  Getter / Setter  #
' #####################
'Attribute Element.VB_UserMemId = 0 saved in both procedures

'@DefaultMember
Public Property Get Element(ByVal index As Long) As Variant
Attribute Element.VB_UserMemId = 0
  Element = iArray(index)
End Property

'@DefaultMember
Public Property Let Element(ByVal index As Long, ByVal val As Variant)
Attribute Element.VB_UserMemId = 0
  If index > iArray.Count Then
    Dim eMsg As String
    eMsg = IIf(iArray.Count > 0, _
               "Index (" & index & ") out of iArray bounds (1-" & iArray.Count & ")", _
               "iArray doesn't have any indexes")
    Err.Raise Number:=9, Description:=eMsg
    Exit Property
  End If
  iArray.Remove index
  iArray.Add val, before:=index
End Property

' #############################
' # COMMON COLLECTION METHODS #
' #############################
Public Function Count() As Long
  Count = iArray.Count
End Function
Public Sub Add(ByVal val As Variant, Optional ByVal bef As Variant = Empty, Optional ByVal aft As Variant = Empty)
  If IsEmpty(bef) And IsEmpty(aft) Then
    iArray.Add val
  ElseIf IsEmpty(bef) Then
    iArray.Add val, , , aft
  ElseIf IsEmpty(aft) Then
    iArray.Add val, , bef
  Else
    iArray.Add val, , bef, aft
  End If
End Sub
Public Sub Remove(index As Long)
  iArray.Remove index
End Sub

' ####################
' #  PUBLIC METHODS  #
' ####################

'AddAfter
' Adds element after given index.
' {1,2,3} -> index >= 3 -> Push val
' {1,2,3} -> index < 1 -> Unshift val
' {} -> index = any -> Push val
' @param Long index | Index after which will be added an element
' @param Variant val | One element (String, number, ...) to add into iArray
' @return Long | Count of elements inside iArray
Public Function AddAfter(ByVal index As Long, ByVal val As Variant) As Long
  If Me.Count < 1 Or index >= Me.Count Then
    Me.Push val
  ElseIf index < 1 Then
    Me.Unshift val
  Else
    Me.Add val, aft:=index
  End If
  AddAfter = Me.Count
End Function

'AddArrayAfter
' Adds elements after given index.
' {1,2,3} -> index >= 3 -> PushArray val
' {1,2,3} -> index < 1 -> UnshiftArray val
' {} -> index = any -> PushArray val
' @param Long index | Index after which will be added an element
' @param Variant val | Array() or iArray of elements to add into iArray
' @return Long | Count of elements inside iArray
Public Function AddArrayAfter(ByVal index As Long, ByVal val As Variant) As Long
  If Me.Count < 1 Or index >= Me.Count Then
    Me.PushArray val
  ElseIf index < 1 Then
    Me.UnshiftArray val
  Else
    If TypeOf val Is iArray Then
      For i = 1 To val.Count
        Me.Add val(i), aft:=index
        Inc index
      Next i
    Else
      For i = LBound(val) To UBound(val)
        Me.Add val(i), aft:=index
        Inc index
      Next i
    End If
  End If
  AddArrayAfter = Me.Count
End Function

'AddArrayBefore
' Adds elements before given index.
' {1,2,3} -> index > 3 -> PushArray val
' {1,2,3} -> index <= 1 -> UnshiftArray val
' {} -> index = any -> PushArray val
' @param Long index | Index after which will be added an element
' @param Variant val | Array() or iArray of elements to add into iArray
' @return Long | Count of elements inside iArray
Public Function AddArrayBefore(ByVal index As Long, ByVal val As Variant) As Long
  If Me.Count < 1 Or index > Me.Count Then
    Me.PushArray val
  ElseIf index <= 1 Then
    Me.UnshiftArray val
  Else
    If TypeOf val Is iArray Then
      For i = val.Count To 1 Step -1
        Me.Add val(i), bef:=index
      Next i
    Else
      For i = UBound(val) To LBound(val) Step -1
        Me.Add val(i), bef:=index
      Next i
    End If
  End If
  AddArrayBefore = Me.Count
End Function

'AddBefore
' Adds element before given index.
' {1,2,3} -> index > 3 -> Push val
' {1,2,3} -> index <= 1 -> Unshift val
' {} -> index = any -> Push val
' @param Long index | Index before which will be added an element
' @param Variant val | One element (String, number, ...) to add into iArray
' @return Long | Count of elements inside iArray
Public Function AddBefore(ByVal index As Long, ByVal val As Variant) As Long
  If Me.Count < 1 Or index > Me.Count Then
    Me.Push val
  ElseIf index <= 1 Then
    Me.Unshift val
  Else
    Me.Add val, bef:=index
  End If
  AddBefore = Me.Count
End Function

'Avg
' Calculates the average of the numeric iArray.
' @return Variant | Average value / "NaN" if iArray contains non-numeric value(s)
Public Function Avg() As Variant
  If Not Me.ContainsOnlyNumeric Then Avg = "NaN": Exit Function
  Avg = Me.Sum / Me.Count
End Function

'Clear
' Empties iArray
Public Function Clear() As Boolean
  For i = iArray.Count To 1 Step -1
    iArray.Remove i
  Next i
  Clear = IIf(iArray.Count = 0, True, False)
End Function

'Clone
' Makes a hard copy of the iArray
' @return iArray | Return exact copy of itself
Public Function Clone() As iArray
  Set Clone = New iArray
  For i = 1 To Me.Count
    Clone.Push Me(i)
  Next i
End Function

'Contains
' Checks if given value is used inside iArray.
' @param Variant val | One element (String, number, ...) to by checked if it exists in iArray
' $return Boolean | True = exists, False = doesn't exists
Public Function Contains(ByVal val As Variant) As Boolean
  Contains = False
  Dim el As Variant
  For Each el In iArray
    If el = val Then Contains = True: Exit Function
  Next el
End Function

'ContainsAll
' Checks if all given values are used inside iArray.
' @param Variant val | values to by checked if it exists in iArray
' $return Boolean | True = all exists, False = one ore more values doesn't exists
Public Function ContainsAll(ByVal val As Variant) As Boolean
  ContainsAll = False
  Dim foundOccurrences As Integer: foundOccurrences = 0
  Dim values As New iArray
  If Not TypeOf val Is iArray Then
    If IsArray(val) Then
      values.PushArray (val)
    Else
      values.Push (val)
    End If
  Else
    Set values = val
  End If
  For i = 1 To values.Count
    If Me.Contains(values(i)) Then Inc foundOccurrences
  Next i
  If foundOccurrences = values.Count Then ContainsAll = True
End Function

'CountOccurrences
' Checks how many times is given value used inside iArray.
' @param Variant val | One element (String, number, ...) to by checked
' @return Long | Count of matched occurrences
Public Function CountOccurrences(val As Variant) As Long
  CountOccurrences = 0
  Dim el As Variant
  For Each el In iArray
    If el = val Then Inc CountOccurrences
  Next el
End Function

'Dequeue
' Removes an element from the begining of the iArray.
' @return Variant | Removed element or Empty, if iArray is empty
Public Function Dequeue() As Variant
  Dequeue = Shift()
End Function

'Difference
' Checks for number of differences between two arrays, what was added/deleted or combination.
' @param iArray oldArray | Old iArray to be compared with current iArray
' @param String retType | “d” = deleted from old iArray, “a” = added in current iArray, “c” = combination of both (default)
' @return iArray | iArray with differences found
Public Function Difference(ByRef oldArray As iArray, Optional ByVal retType As String = "c") As iArray
  Set Difference = New iArray
  Select Case LCase(retType)
    Case "d"
      For i = 1 To oldArray.Count
        If Me.Contains(oldArray(i)) = False Then Difference.Push oldArray(i)
      Next i
    Case "a"
      For i = 1 To Count
        If oldArray.Contains(Me(i)) = False Then Difference.Push Me(i)
      Next i
    Case "c"
      For i = 1 To oldArray.Count
        If Me.Contains(oldArray(i)) = False Then Difference.Push oldArray(i)
      Next i
      For i = 1 To Count
        If oldArray.Contains(Me(i)) = False Then Difference.Push Me(i)
      Next i
    Case Else
      Err.Raise Number:=5, Description:="Unknown Difference return type attribute:" & vbCrLf & _
                                        "Allowed are 'A'dded/'D'eleted/'C'ombination of both"
  End Select
End Function

'DropLeft
' Remove n elements from the beginning of the iArray. If n > count of iArray elements, all elements are removed.
' @param Long n | Number of elements to be removed
' @return iArray | Array of the removed elements
Public Function DropLeft(ByRef n As Long) As iArray
  Set DropLeft = New iArray
  If n > Me.Count Then n = Me.Count
  For i = 1 To n
    DropLeft.Push Me(i)
  Next i
  For i = n To 1 Step -1
    Me.Remove (i)
  Next i
End Function

'DropRight
' Remove n elements from the end of the iArray. If n > count of iArray elements, all elements are removed.
' @param Long n | Number of elements to be removed
' @return iArray | Array of the removed elements
Public Function DropRight(ByRef n As Long) As iArray
  Set DropRight = New iArray
  If n > Me.Count Then n = Me.Count
  For i = Me.Count - n + 1 To Me.Count
    DropRight.Push Me(i)
  Next i
  For i = Me.Count To Me.Count - n + 1 Step -1
    Me.Remove (i)
  Next i
End Function

'Enqueue
' Adds an element at the end of the iArray.
' @param Variant val | One element (String, number, ...) to add into iArray
' @return Long | Count of elements inside iArray
Public Function Enqueue(ByVal val As Variant) As Long
  Enqueue = Push(val)
End Function

'EnqueueArray
' Adds elements at the end of the iArray
' @param Variant val | Array() or iArray of elements to add into iArray
' @return Long | Count of elements inside array
Public Function EnqueueArray(ByVal val As Variant) As Long
  EnqueueArray = PushArray(val)
End Function

'First
' Returns value of the first element of the iArray.
' @return Variant | Value of the first element or Empty if iArray is Empty
Public Function First() As Variant
  First = IIf(Me.Count > 0, Me(1), Empty)
End Function

'Head
' Returns all elements of iArray, except last one.
' @return iArray | Copy of original array, without last element
Public Function Head() As iArray
  Set Head = Me.Clone
  Head.Pop
End Function

'Intersect
' Search for element that exists in both iArrays.
' @param iArray intArray | iArray to be intersected with current iArray
' @return iArray | Intersected values
Public Function Intersect(ByRef intersectedArr As iArray) As iArray
  Set Intersect = New iArray
  Dim el As Variant
  For Each el In iArray
    If intersectedArr.Contains(el) Then
      If Not Intersect.Contains(el) Then Intersect.Push (el)
    End If
  Next el
End Function

'ContainsOnlyNumeric
' Verify that all elements are numbers or convertable to numbers
' $return Boolean | True = contains only numeric values, False = contains inconvertible values
Public Function ContainsOnlyNumeric() As Boolean
  ContainsOnlyNumeric = True
  Dim el As Variant
  For Each el In iArray
    If Not IsNumeric(el) Then ContainsOnlyNumeric = False: Exit Function
  Next el
End Function

'Join
' Joins two iArrays.
' @param iArray jArray | iArray to be joined with current iArray
' @return iArray | Joined iArray
Public Function Join(ByRef joinArr As iArray) As iArray
  Set Join = New iArray
  Join.PushArray Me
  Join.PushArray joinArr
End Function

'Last
' Returns value of the last element of the iArray.
' @return Variant | Value of the last element or Empty if iArray is Empty
Public Function Last() As Variant
  Last = IIf(Me.Count > 0, Me(Me.Count), Empty)
End Function

'OccurrenceIndexes
' Returns all indexes of a value in iArray.
' @return iArray | iArray containing indexes of occurrence
Public Function OccurenceIndexes(val As Variant) As iArray
  Set OccurenceIndexes = New iArray
  For i = 1 To Me.Count
    If iArray(i) = val Then OccurenceIndexes.Push (i)
  Next i
End Function

'Pop
' Removes an element from the end of the iArray.
' @return Variant | Removed element or Empty, if iArray is empty
Public Function Pop() As Variant
  Pop = Empty
  If iArray.Count > 0 Then
    Pop = Me(iArray.Count)
    iArray.Remove iArray.Count
  End If
End Function

'Push
' Adds an element at the end of the iArray.
' @param Variant val | One element (String, number, ...) to add into iArray
' @return Long | Count of elements inside iArray
Public Function Push(ByVal val As Variant) As Long
  iArray.Add val
  Push = iArray.Count
End Function

'PushArray
' Adds elements at the end of the iArray.
' @param Variant val | Array() or iArray of elements to add into iArray
' @return Long | Count of elements inside array
Public Function PushArray(ByVal val As Variant) As Long
  If TypeOf val Is iArray Then
    For i = 1 To val.Count
      Me.Push val(i)
    Next i
  Else
    For i = LBound(val) To UBound(val)
      Me.Push val(i)
    Next i
  End If
  PushArray = Me.Count
End Function

'RemoveDuplicates
' Keeps only the first occurrence of the value. Affects original array
' @return Long | Count of the removed elements
Public Function RemoveDuplicates() As Long
  RemoveDuplicates = 0
  For i = Me.Count To 1 Step -1
    If Me.CountOccurrences(Me(i)) > 1 Then
      Me.Remove i
      Inc RemoveDuplicates
    End If
  Next i
End Function

'Reverse
' Reverses the content of the iArray
' @return iArray | Reversed iArray
Public Function Reverse() As iArray
  Set Reverse = New iArray
  For i = 1 To Me.Count
    Reverse.Unshift Me(i)
  Next i
End Function

'Shift
' Removes an element from the begining of the iArray.
' @return Variant | Removed element or Empty, if iArray is empty
Public Function Shift() As Variant
  Shift = Empty
  If iArray.Count > 0 Then
    Shift = Me(1)
    iArray.Remove 1
  End If
End Function

'Shuffle
' Randomly mixes content of the iArray
' @return iArray | Shuffled iArray
Public Function Shuffle() As iArray
  Set Shuffle = New iArray
  Dim keys As New iArray
  For i = 1 To Me.Count
    keys.Push (i)
  Next i
  Dim rndNo As Long
  Do While keys.Count > 0
    rndNo = randomFromTo(1, keys.Count)
    Shuffle.Push Me(keys(rndNo))
    keys.Remove (rndNo)
  Loop
End Function

'Sum
' Calculates the sum of the numeric iArray.
' @return Variant | Sum value / "NaN" if iArray contains non-numeric value(s)
Public Function Sum() As Variant
  If Not Me.ContainsOnlyNumeric Then Sum = "NaN": Exit Function
  Sum = 0
  Dim el As Variant
  For Each el In iArray
    Sum = Sum + el
  Next el
End Function

'Tail
' Returns all elements of iArray, except first one.
' @return iArray | Copy of original array, without first element
Public Function Tail() As iArray
  Set Tail = Me.Clone
  Tail.Dequeue
End Function

'ToString
' Creates string representation of the iArray.
' @param String delimiter | Optional character to separate the iArray's elements (default = ";")
' @return String | Formated representation of tha iArray
Public Function ToString(Optional ByVal delimiter As String = ";") As String
  If Me.Count = 0 Then ToString = "{}": Exit Function
  ToString = vbNullString
  For i = 1 To Me.Count
    If i = 1 Then ToString = ToString + "{"
    If i > 1 Then ToString = ToString + delimiter
    If VarType(Me(i)) = vbString Then ToString = ToString + Chr$(34)
    ToString = ToString + CStr(Me(i))
    If VarType(Me(i)) = vbString Then ToString = ToString + Chr$(34)
    If i = iArray.Count Then ToString = ToString + "}"
  Next i
End Function

'Union
' Combines both iArrays and removes duplicities
' @param iArray uArray | iArray to be unioned with current iArray
' @return iArray | Combined values without duplicities
Public Function Union(ByRef uArr As iArray) As iArray
  Set Union = Me.Clone
  Union.PushArray uArr
  Union.RemoveDuplicates
End Function

'Unique
' Returns copy of iArray without duplicated values. Original array is not affected
' @return iArray | iArray with unique values
Public Function Unique() As iArray
  Set Unique = New iArray
  Dim el As Variant
  For Each el In iArray
    If Not Unique.Contains(el) Then Unique.Push (el)
  Next el
End Function

'Unshift
' Add an element at the begining of the iArray.
' @param Variant val | One element (String, number, ...) to add into iArray
' @return Long | Count of elements inside iArray
Public Function Unshift(ByVal val As Variant) As Long
  If iArray.Count = 0 Then
    Unshift = Push(val)
  Else
    iArray.Add val, before:=1
    Unshift = iArray.Count
  End If
End Function

'UnshiftArray
' Add elements at the begining of the iArray.
' @param Variant val | Array() or iArray of elements to add into iArray
' @return Long | Count of elements inside iArray
Public Function UnshiftArray(ByVal val As Variant) As Long
  If TypeOf val Is iArray Then
    For i = val.Count To 1 Step -1
      Unshift val(i)
    Next i
  Else
    For i = UBound(val) To LBound(val) Step -1
      Unshift val(i)
    Next i
  End If
  UnshiftArray = iArray.Count
End Function


' #####################
' # PRIVATE FUNCTIONS #
' #####################
Sub Inc(ByRef i)
   i = i + 1
End Sub

Private Function rangeToArray(ByVal rng As Range) As Variant
  rangeToArray = rng.Value
End Function

Private Function randomFromTo(ByVal from As Long, ByVal upTo As Long) As Long
  Randomize
  randomFromTo = Int(from + Rnd() * (upTo - from + 1))
End Function

