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
  Debug.Print iArr.Push("Hello world")           ' 4
  Debug.Print iArr.ToString                      ' {"a",True,1,"Hello world"}
  iArr.Pop
  iArr.Pop
  Debug.Print iArr.Pop                           ' True
  iArr.Pop
  Debug.Print iArr.Pop                           ' Empty value
  Debug.Print iArr.ToString                      ' {}
  
  ' ##### SHIFT / UNSHIFT TEST
  Debug.Print vbCrLf & " #### Shift/Unshift test"
  Set iArr = New iArray
  iArr.Unshift "..."
  Debug.Print iArr.ToString                      ' {"..."}
  iArr.Unshift 123456
  iArr.UnshiftArray Array(3.1415, Empty, vbNullString, "a")
  Debug.Print iArr.ToString                      ' {3.1415,,"","a",123456,"..."}
  Debug.Print iArr.Shift                         ' 3.1415
  Debug.Print iArr.Shift                         ' Empty value
  Debug.Print iArr.Shift                         '
  Debug.Print iArr.ToString                      ' {"a",123456,"..."}

  ' ##### ENQUEUE / DEQUEUE TEST
  Debug.Print vbCrLf & " #### Enqueue/Dequeue test"
  Set iArr = New iArray
  iArr.Enqueue ("Queued Item")
  iArr.EnqueueArray Array(1, "2", 3.14, False, "Last")
  Debug.Print iArr.Dequeue                       ' "Queued Item"
  Debug.Print iArr.Dequeue                       ' 1
  Debug.Print iArr.ToString                      ' {"2",3.14,False,"Last"}
  
  ' ##### DEFAULT MEMBERS TEST
  Debug.Print vbCrLf & " #### Default members test"
  Set iArr = New iArray
  iArr.PushArray Array("1", 2, "3", 4)
  Debug.Print iArr(2)                            ' 2
  iArr(2) = "Two"
  Debug.Print iArr.ToString                      ' {"1","Two","3", 4}

  ' ##### CLEAR ARRAY TEST
  Debug.Print vbCrLf & " #### Clear array test"
  Set iArr = New iArray
  iArr.PushArray Array(1, 2, 3, 4, 5, "a", "b", "c", True, Empty)
  Debug.Print iArr.ToString                      ' {1,2,3,4,5,"a","b","c",True,}
  iArr.Clear
  Debug.Print iArr.ToString                      ' {}
  
  ' ##### COUNT OCCURENCES TEST
  Debug.Print vbCrLf & " #### Count occurences test"
  Set iArr = New iArray
  iArr.PushArray Array(1, 2, 2, 1, 3, 1, 2)
  Debug.Print iArr.CountOccurences(2)            ' 3
  Debug.Print iArr.CountOccurences(4)            ' 0
  
  
  ' ##### CONTAINS TEST
  Debug.Print vbCrLf & " #### Contains test"
  Set iArr = New iArray
  iArr.PushArray Array(1, 2, 2, 1, 3, 1, 2)
  Debug.Print iArr.Contains(1)                   ' True
  Debug.Print iArr.Contains(5)                   ' False

  ' ##### FIND DIFFERENCES TEST
  Debug.Print vbCrLf & " #### Difference test"
  Set iArr1 = New iArray
  iArr1.PushArray Array(1, 2, 3)
  Set iArr2 = New iArray
  iArr2.PushArray Array(2, 3, 4)
  Set iArr = iArr2.Difference(iArr1)
  Debug.Print iArr.ToString                      ' {1,4}
  Set iArr = iArr2.Difference(iArr1, "d")
  Debug.Print iArr.ToString                      ' {1}
  Set iArr = iArr2.Difference(iArr1, "a")
  Debug.Print iArr.ToString                      ' {4}
  
  ' ##### JOINING ARRAYS TEST
  Debug.Print vbCrLf & " #### Joining arrays test"
  Set iArr1 = New iArray
  iArr1.PushArray Array(1, 2, 3, "a", "b", "c")
  Set iArr2 = New iArray
  iArr2.PushArray Array(4, 5, 6, "d", "e", "f")
  Set iArr = iArr1.Join(iArr2)
  Debug.Print iArr.ToString                      ' {1,2,3,"a","b","c",4,5,6,"d","e","f"}
  
  ' ##### DROP LEFT/RIGHT TEST
  Debug.Print vbCrLf & " #### Drop left/right test"
  Set iArr = New iArray
  iArr.PushArray Array("1", "Two", "3", 4, 1, 2, 3, 4, 5, "a", "b", "c", True)
  Debug.Print iArr.DropLeft(2).ToString          ' {"1", "Two"}
  Debug.Print iArr.ToString                      ' {"3",4,1,2,3,4,5,"a","b","c",True}
  Debug.Print iArr.DropRight(3).ToString         ' {"b","c",True}
  Debug.Print iArr.ToString                      ' {"3",4,1,2,3,4,5,"a"}
  
  ' ##### REMOVE DUPLICATES TEST
  Debug.Print vbCrLf & " #### Remove duplicates test"
  Set iArr = New iArray
  iArr.PushArray Array(1, 2, "a", 2, 3, 2, 3.14, "b", True, 4, "a")
  Debug.Print iArr.RemoveDuplicates              ' 3
  Debug.Print iArr.ToString                      ' {1,2,"a",3,3.14,"b",True,4}
  
  ' ##### CLONE ARRAY TEST
  Debug.Print vbCrLf & " #### Clone array test"
  Set iArr = New iArray
  iArr.PushArray Array("3", 4, 1, 2, 3, 4, 5, "a", "b", "c", True)
  Dim arrCloned As New iArray
  Set arrCloned = iArr.Clone
  iArr.Clear
  Debug.Print arrCloned.ToString                 ' {"3",4,1,2,3,4,5,"a","b","c",True}
  
  ' ##### SHUFFLE ARRAY TEST
  Debug.Print vbCrLf & " #### Shuffle array test"
  Set iArr = New iArray
  iArr.PushArray Array("3", 4, 1, 2, 3, 4, 5, "a", "b", "c", True)
  Dim arrShufl As New iArray
  Set arrShufl = iArr.Shuffle
  Debug.Print arrShufl.ToString                  ' e.g. {"a",4,4,2,3,1,5,"3",True,"b","c"}
  
  ' ##### REVERSE ARRAY TEST
  Debug.Print vbCrLf & " #### Reverse array test"
  Set iArr = New iArray
  iArr.PushArray Array("3", 4, 1, 2, 3, 4, 5, "a", "b", "c", True)
  Dim arrRev As New iArray
  Set arrRev = iArr.Reverse
  Debug.Print arrRev.ToString                    ' {True,"c","b","a",5,4,3,2,1,4,"3"}
  
  ' ##### FIRST/LAST TEST
  Debug.Print vbCrLf & " #### First/Last test"
  Set iArr = New iArray
  iArr.PushArray Array(1, 2, 3, 4, 5)
  Debug.Print iArr.First                         ' 1
  Debug.Print iArr.Last                          ' 5
  
  ' ##### ADD AFTER/BEFORE TEST
  Debug.Print vbCrLf & " #### Add after/before test"
  Set iArr = New iArray
  iArr.AddBefore 2, "Something"
  iArr.PushArray Array(1, 2, 3, 4, 5)
  Debug.Print iArr.ToString                      ' {"Something",1,2,3,4,5}
  iArr.AddBefore 1, "New First"
  Debug.Print iArr.ToString                      ' {"New First","Something",1,2,3,4,5}
  iArr.AddAfter 4, "Hello"
  Debug.Print iArr.ToString                      ' {"New First","Something",1,2,"Hello",3,4,5}
  
  ' ##### ADD AFTER/BEFORE ARRAY TEST
  Debug.Print vbCrLf & " #### Add after/before array test"
  Set iArr = New iArray
  iArr.PushArray Array(1, 2, 3, 4, 5)
  iArr.AddArrayAfter 2, Array("a", "b", "c")
  Debug.Print iArr.ToString                      ' {1,2,"a","b","c",3,4,5}
  iArr.AddArrayBefore 7, Array(True, False)
  Debug.Print iArr.ToString                      ' {1,2,"a","b","c",3,True,False,4,5}
  
End Sub

