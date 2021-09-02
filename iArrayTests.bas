Attribute VB_Name = "iArrayTests"
Option Explicit

Private iArr As iArray
Private iArr1 As iArray
Private iArr2 As iArray

Public Sub iArrayTest()
  ' ##### PUSH / POP TEST
  Debug.Print vbCrLf & " #### Push/Pop test"
  Set iArr = New iArray
  iArr.PushArray Array("a", True, 1)
  Call validate("Push", 4, iArr.Push("Hello world"))
  
  iArr.Pop
  iArr.Pop
  Call validate("Pop", True, iArr.Pop)
    
  ' ##### SHIFT / UNSHIFT TEST
  Debug.Print vbCrLf & " #### Shift/Unshift test"
  Set iArr = New iArray
  iArr.Unshift "..."
  Call validate("Unshift", "{""...""}", iArr.ToString)
  iArr.Unshift 123456
  iArr.UnshiftArray Array(3.1415, Empty, vbNullString, "a")
  Call validate("UnshiftArray", Array("{3.1415;;"""";""a"";123456;""...""}", "{3,1415;;"""";""a"";123456;""...""}"), iArr.ToString)
  iArr.Shift
  iArr.Shift
  iArr.Shift
  Call validate("Shift", "{""a"";123456;""...""}", iArr.ToString)

  ' ##### ENQUEUE / DEQUEUE TEST
  Debug.Print vbCrLf & " #### Enqueue/Dequeue test"
  Set iArr = New iArray
  iArr.Enqueue ("Queued element")
  Call validate("Enqueue", "{""Queued element""}", iArr.ToString)
  iArr.EnqueueArray Array(1, "2", 3.14, False, "Last")
  Call validate("EnqueueArray", Array("{""Queued element"";1;""2"";3.14;False;""Last""}", "{""Queued element"";1;""2"";3,14;False;""Last""}"), iArr.ToString)
  iArr.Dequeue
  Call validate("Dequeue", 1, iArr.Dequeue)

  ' ##### DEFAULT MEMBERS TEST
  Debug.Print vbCrLf & " #### Default members test"
  Set iArr = New iArray
  iArr.PushArray Array("1", 2, "3", 4)
  Call validate("Default Members set", 2, iArr(2))
  iArr(2) = "Two"
  Call validate("Default Members edit", "Two", iArr(2))

  ' ##### CLEAR ARRAY TEST
  Debug.Print vbCrLf & " #### Clear array test"
  Set iArr = New iArray
  iArr.PushArray Array(1, 2, 3, 4, 5, "a", "b", "c", True, Empty)
  iArr.Clear
  Call validate("Clear", "{}", iArr.ToString)

  ' ##### COUNT OCCURRENCES TEST
  Debug.Print vbCrLf & " #### Count occurrences test"
  Set iArr = New iArray
  iArr.PushArray Array(1, 2, 2, 1, 3, 1, 2)
  Call validate("Count occurrences (yes)", 3, iArr.CountOccurrences(2))
  Call validate("Count occurrences (not)", 0, iArr.CountOccurrences(4))

  ' ##### CONTAINS TEST
  Debug.Print vbCrLf & " #### Contains test"
  Set iArr = New iArray
  iArr.PushArray Array(1, 2, 2, 1, 3, 1, 2)
  Call validate("Contains (yes)", True, iArr.Contains(1))
  Call validate("Contains (not)", False, iArr.Contains(5))

  ' ##### CONTAINS ALL TEST
  Debug.Print vbCrLf & " #### Contains All test"
  Set iArr = New iArray
  iArr.PushArray Array(1, 2, 2, 1, 3, 1, 2)
  Call validate("ContainsAll (yes)", True, iArr.ContainsAll(Array(1, 3)))
  Call validate("ContainsAll (not)", False, iArr.ContainsAll(Array(1, 4, 5)))

  ' ##### FIND DIFFERENCES TEST
  Debug.Print vbCrLf & " #### Difference test"
  Set iArr1 = New iArray
  iArr1.PushArray Array(1, 2, 3)
  Set iArr2 = New iArray
  iArr2.PushArray Array(2, 3, 4)
  Set iArr = iArr2.Difference(iArr1)
  Call validate("Difference", "{1;4}", iArr.ToString)
  Set iArr = iArr2.Difference(iArr1, "d")
  Call validate("Difference (with ""d"" param)", "{1}", iArr.ToString)
  Set iArr = iArr2.Difference(iArr1, "a")
  Call validate("Difference (with ""a"" param)", "{4}", iArr.ToString)

  ' ##### JOINING ARRAYS TEST
  Debug.Print vbCrLf & " #### Joining arrays test"
  Set iArr1 = New iArray
  iArr1.PushArray Array(1, 2, 3, "a", "b", "c")
  Set iArr2 = New iArray
  iArr2.PushArray Array(4, 5, 6, "d", "e", "f")
  Set iArr = iArr1.Join(iArr2)
  Call validate("Join", "{1;2;3;""a"";""b"";""c"";4;5;6;""d"";""e"";""f""}", iArr.ToString)

  ' ##### DROP LEFT/RIGHT TEST
  Debug.Print vbCrLf & " #### Drop left/right test"
  Set iArr = New iArray
  iArr.PushArray Array("1", "Two", "3", 4, 1, 2, 3, 4, 5, "a", "b", "c", True)
  Call validate("DropLeft return", "{""1"";""Two""}", iArr.DropLeft(2).ToString)
  Call validate("DropLeft", "{""3"";4;1;2;3;4;5;""a"";""b"";""c"";True}", iArr.ToString)
  Call validate("DropRight return", "{""b"";""c"";True}", iArr.DropRight(3).ToString)
  Call validate("DropRight", "{""3"";4;1;2;3;4;5;""a""}", iArr.ToString)

  ' ##### UNIQUE / REMOVE DUPLICATES TEST
  Debug.Print vbCrLf & " #### Unique / Remove duplicates test"
  Set iArr = New iArray
  iArr.PushArray Array("3", 4, 1, 2, 3, 4, 5, "a", "b", "3", "c", "a", True)
  Dim uniqueArr As New iArray
  Set uniqueArr = iArr.Unique
  Call validate("Unique", "{""3"";4;1;2;3;5;""a"";""b"";""c"";True}", uniqueArr.ToString)
  Call validate("RemoveDuplicates (removed count)", 3, iArr.RemoveDuplicates)
  Call validate("RemoveDuplicates", "{""3"";4;1;2;3;5;""a"";""b"";""c"";True}", iArr.ToString)

  ' ##### CLONE ARRAY TEST
  Debug.Print vbCrLf & " #### Clone array test"
  Set iArr = New iArray
  iArr.PushArray Array("3", 4, 1, 2, 3, 4, 5, "a", "b", "c", True)
  Dim arrCloned As New iArray
  Set arrCloned = iArr.Clone
  iArr.Clear
  Call validate("Clone", "{""3"";4;1;2;3;4;5;""a"";""b"";""c"";True}", arrCloned.ToString)

  ' ##### SHUFFLE ARRAY TEST
  Debug.Print vbCrLf & " #### Shuffle array test"
  Set iArr = New iArray
  iArr.PushArray Array("3", 4, 1, 2, 3, 4, 5, "a", "b", "c", True)
  Dim arrShufl As New iArray
  Set arrShufl = iArr.Shuffle
  Call validate("Shuffle", True, arrShufl.ToString <> iArr.ToString And arrShufl.ContainsAll(iArr))

  ' ##### REVERSE ARRAY TEST
  Debug.Print vbCrLf & " #### Reverse array test"
  Set iArr = New iArray
  iArr.PushArray Array("3", 4, 1, 2, 3, 4, 5, "a", "b", "c", True)
  Dim arrRev As New iArray
  Set arrRev = iArr.Reverse
  Call validate("Reverse", "{True;""c"";""b"";""a"";5;4;3;2;1;4;""3""}", arrRev.ToString)

  ' ##### FIRST/LAST TEST
  Debug.Print vbCrLf & " #### First/Last test"
  Set iArr = New iArray
  iArr.PushArray Array(1, 2, 3, 4, 5)
  Call validate("First", 1, iArr.First)
  Call validate("Last", 5, iArr.Last)

  ' ##### ADD AFTER/BEFORE TEST
  Debug.Print vbCrLf & " #### Add after/before test"
  Set iArr = New iArray
  iArr.AddBefore 2, "Something"
  iArr.PushArray Array(1, 2, 3, 4, 5)
  iArr.AddBefore 1, "New First"
  Call validate("AddBefore", "{""New First"";""Something"";1;2;3;4;5}", iArr.ToString)
  iArr.AddAfter 4, "Hello"
  Call validate("AddAfter", "{""New First"";""Something"";1;2;""Hello"";3;4;5}", iArr.ToString)

  ' ##### ADD AFTER/BEFORE ARRAY TEST
  Debug.Print vbCrLf & " #### Add after/before array test"
  Set iArr = New iArray
  iArr.PushArray Array(1, 2, 3, 4, 5)
  iArr.AddArrayAfter 2, Array("a", "b", "c")
  Call validate("AddArrayBefore", "{1;2;""a"";""b"";""c"";3;4;5}", iArr.ToString)
  iArr.AddArrayBefore 7, Array(True, False)
  Call validate("AddArrayBefore", "{1;2;""a"";""b"";""c"";3;True;False;4;5}", iArr.ToString)
  
  ' ##### TAIL/HEAD TEST
  Debug.Print vbCrLf & " #### Tail / Head test"
  Set iArr = New iArray
  iArr.PushArray Array("3", 4, 1, 2, 3, 4, 5, "a", "b", "c", True)
  Dim tailArr As New iArray
  Set tailArr = iArr.Tail
  Call validate("Tail", "{4;1;2;3;4;5;""a"";""b"";""c"";True}", tailArr.ToString)
  Dim headArr As New iArray
  Set headArr = tailArr.Head
  Call validate("Head", "{4;1;2;3;4;5;""a"";""b"";""c""}", headArr.ToString)
End Sub

Private Sub validate(name As String, expected, actual As String)
  If Not IsArray(expected) Then expected = Array(expected)
  
  Dim found As Boolean: found = False
  Dim possibleResult As String
  
  Dim i As Integer
  For i = LBound(expected) To UBound(expected)
    If expected(i) = actual Then found = True: Exit For
  Next i
  
  If found Then
    Debug.Print name + " - OK"
  Else
    Debug.Print name + " - NOK"
    Debug.Print " - Actual value: " + actual
    
    Dim expectedString As Variant
    For i = LBound(expected) To UBound(expected)
      If i > LBound(expected) Then expectedString = expectedString + " or "
      expectedString = expectedString + expected(i)
    Next i
    Debug.Print " - Expected value: " + CStr(expectedString)
  End If
End Sub
