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
'Copyright (c) 2019 Petr Katerinak
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
' Version of iArray: 0.1

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
'Attribute Item.VB_UserMemId = 0 saved in both procedures

'@DefaultMember
Public Property Get Item(ByVal index As Long) As Variant
Attribute Item.VB_UserMemId = 0
  Item = iArray(index)
End Property

'@DefaultMember
Public Property Let Item(ByVal index As Long, ByVal val As Variant)
Attribute Item.VB_UserMemId = 0
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
Public Sub Add(ByVal val As Variant, Optional ByVal before As Variant = Empty, Optional ByVal after As Variant = Empty)
  iArray.Add val, before:=before, after:=after
End Sub
Public Sub Remove(index As Long)
  iArray.Remove index
End Sub

' ####################
' #  PUBLIC METHODS  #
' ####################

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
' Checks how many times is given value used inside iArray.
' @param Variant val | One item (String, number, ...) to by checked if it exists in iArray
' $return Boolean | True = exists, False = doesn't exists
Public Function Contains(ByVal val As Variant) As Boolean
  Contains = False
  Dim el As Variant
  For Each el In iArray
    If el = val Then Contains = True: Exit Function
  Next el
End Function

'CountOccurences
' Checks how many times is given value used inside iArray.
' @param Variant val | One item (String, number, ...) to by checked
' @return Long | Count of matched occurences
Public Function CountOccurences(val As Variant) As Long
  CountOccurences = 0
  Dim el As Variant
  For Each el In iArray
    If el = val Then CountOccurences = CountOccurences + 1
  Next el
End Function

'Dequeue
' Removes an item from the begining of the iArray.
' @return Variant | Removed item or Null, if iArray is empty
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
' Remove n items from the beginning of the iArray. If n > count of iArray items, all items are removed.
' @param Long n | Number of items to be removed
' @return iArray | Array of the removed items
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
' Remove n items from the end of the iArray. If n > count of iArray items, all items are removed.
' @param Long n | Number of items to be removed
' @return iArray | Array of the removed items
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
' Adds an item at the end of the iArray.
' @param Variant val | One item (String, number, ...) to add into iArray
' @return Long | Count of items inside iArray
Public Function Enqueue(ByVal val As Variant) As Long
  Enqueue = Push(val)
End Function

'EnqueueArray
' Adds items at the end of the iArray
' @param Variant val | Array() of items to add into iArray
' @return Long | Count of items inside array
Public Function EnqueueArray(ByVal val As Variant) As Long
  EnqueueArray = PushArray(val)
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

'Pop
' Removes an item from the end of the iArray.
' @return Variant | Removed item or Null, if iArray is empty
Public Function Pop() As Variant
  Pop = Null
  If iArray.Count > 0 Then
    Pop = Me(iArray.Count)
    iArray.Remove iArray.Count
  End If
End Function

'Push
' Adds an item at the end of the iArray.
' @param Variant val | One item (String, number, ...) to add into iArray
' @return Long | Count of items inside iArray
Public Function Push(ByVal val As Variant) As Long
  iArray.Add val
  Push = iArray.Count
End Function

'PushArray
' Adds items at the end of the iArray.
' @param Variant val | Array() of items to add into iArray
' @return Long | Count of items inside array
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
' Keeps only first occurences of the values
' @return Long | Count of the removed items
Public Function RemoveDuplicates() As Long
  RemoveDuplicates = 0
  For i = Me.Count To 1 Step -1
    If Me.CountOccurences(Me(i)) > 1 Then
      Me.Remove i
      RemoveDuplicates = RemoveDuplicates + 1
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
' Removes an item from the begining of the iArray.
' @return Variant | Removed item or Null, if iArray is empty
Public Function Shift() As Variant
  Shift = Null
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

'ToString
' Creates string representation of the iArray.
' @param String delimiter | Optional character to separate the iArray's items (default = ",")
' @return String | Formated representation of tha iArray
Public Function ToString(Optional ByVal delimiter As String = ",") As String
  If iArray.Count = 0 Then ToString = "{}": Exit Function
  ToString = vbNullString
  For i = 1 To iArray.Count
    If i = 1 Then ToString = ToString + "{"
    If i > 1 Then ToString = ToString + delimiter
    If VarType(Me(i)) = vbString Then ToString = ToString + Chr$(34)
    ToString = ToString + CStr(Me(i))
    If VarType(Me(i)) = vbString Then ToString = ToString + Chr$(34)
    If i = iArray.Count Then ToString = ToString + "}"
  Next i
End Function

'Unshift
' Add an item at the begining of the iArray.
' @param Variant val | One item (String, number, ...) to add into iArray
' @return Long | Count of items inside iArray
Public Function Unshift(ByVal val As Variant) As Long
  If iArray.Count = 0 Then
    Unshift = Push(val)
    Exit Function
  Else
    iArray.Add val, before:=1
    Unshift = iArray.Count
  End If
End Function

'UnshiftArray
' Add items at the begining of the iArray.
' @param Variant val | Array() of items to add into iArray
' @return Long | Count of items inside iArray
Public Function UnshiftArray(ByVal val As Variant) As Long
  For i = UBound(val) To LBound(val) Step -1
    Unshift val(i)
  Next i
  UnshiftArray = iArray.Count
End Function


' #####################
' # PRIVATE FUNCTIONS #
' #####################
Private Function rangeToArray(ByVal rng As Range) As Variant
  rangeToArray = rng.Value
End Function

Private Function randomFromTo(ByVal from As Long, ByVal upTo As Long) As Long
  Randomize
  randomFromTo = Int(from + Rnd() * (upTo - from + 1))
End Function

