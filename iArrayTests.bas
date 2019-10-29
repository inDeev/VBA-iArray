Attribute VB_Name = "iArrayTests"
Option Explicit

Public Sub iArrayTest()
  ' ##### PUSH / POP TEST
  Debug.Print vbCrLf & " #### Push/Pop test"
  Dim arrPP As New iArray
  arrPP.PushArray Array("a", True, 1)
  Debug.Print arrPP.Push("Hello world")          ' 4
  Debug.Print arrPP.ToString                     ' {"a",True,1,"Hello world"}
  arrPP.Pop
  arrPP.Pop
  Debug.Print arrPP.Pop                          ' True
  arrPP.Pop
  Debug.Print arrPP.Pop                          ' Null
  Debug.Print arrPP.ToString                     ' {}
  
  ' ##### SHIFT / UNSHIFT TEST
  Debug.Print vbCrLf & " #### Shift/Unshift test"
  Dim arrSU As New iArray
  arrSU.Unshift "..."
  Debug.Print arrSU.ToString                     ' {"..."}
  arrSU.Unshift 123456
  arrSU.UnshiftArray Array(3.1415, Empty, vbNullString, "a")
  Debug.Print arrSU.ToString                     ' {3.1415,,"","a",123456,"..."}
  Debug.Print arrSU.Shift                        ' 3.1415
  Debug.Print arrSU.Shift                        ' Null
  Debug.Print arrSU.Shift                        '
  Debug.Print arrSU.ToString                     ' {"a",123456,"..."}

  ' ##### ENQUEUE / DEQUEUE TEST
  Debug.Print vbCrLf & " #### Enqueue/Dequeue test"
  Dim arrED As New iArray
  arrED.Enqueue ("Queued Item")
  arrED.EnqueueArray Array(1, "2", 3.14, False, "Last")
  Debug.Print arrED.Dequeue                      ' "Queued Item"
  Debug.Print arrED.Dequeue                      ' 1
  Debug.Print arrED.ToString                     ' {"2",3.14,False,"Last"}
  
  ' ##### DEFAULT MEMBERS TEST
  Debug.Print vbCrLf & " #### Default members test"
  Dim arrDM As New iArray
  arrDM.PushArray Array("1", 2, "3", 4)
  Debug.Print arrDM(2)                           ' 2
  arrDM(2) = "Two"
  Debug.Print arrDM.ToString                     ' {"1","Two","3", 4}

  ' ##### CLEAR ARRAY TEST
  Debug.Print vbCrLf & " #### Clear array test"
  Dim arrCA As New iArray
  arrCA.PushArray Array(1, 2, 3, 4, 5, "a", "b", "c", True, Empty)
  Debug.Print arrCA.ToString                     ' {1,2,3,4,5,"a","b","c",True,}
  arrCA.Clear
  Debug.Print arrCA.ToString                     ' {}
  
  ' ##### COUNT OCCURENCES TEST
  Debug.Print vbCrLf & " #### Count occurences test"
  Dim arrCO As New iArray
  arrCO.PushArray Array(1, 2, 2, 1, 3, 1, 2)
  Debug.Print arrCO.CountOccurences(2)           ' 3
  Debug.Print arrCO.CountOccurences(4)           ' 0
  
  
  ' ##### CONTAINS TEST
  Debug.Print vbCrLf & " #### Contains test"
  Dim arrCon As New iArray
  arrCon.PushArray Array(1, 2, 2, 1, 3, 1, 2)
  Debug.Print arrCon.Contains(1)                 ' True
  Debug.Print arrCon.Contains(5)                 ' False

  ' ##### FIND DIFFERENCES TEST
  Debug.Print vbCrLf & " #### Difference test"
  Dim arrDiff1 As New iArray
  arrDiff1.PushArray Array(1, 2, 3)
  Dim arrDiff2 As New iArray
  arrDiff2.PushArray Array(2, 3, 4)
  Dim arrDiff3 As New iArray
  Set arrDiff3 = arrDiff2.Difference(arrDiff1)
  Debug.Print arrDiff3.ToString                  ' {1,4}
  Set arrDiff3 = arrDiff2.Difference(arrDiff1, "d")
  Debug.Print arrDiff3.ToString                  ' {1}
  Set arrDiff3 = arrDiff2.Difference(arrDiff1, "a")
  Debug.Print arrDiff3.ToString                  ' {4}
  
  ' ##### JOINING ARRAYS TEST
  Debug.Print vbCrLf & " #### Joining arrays test"
  Dim arrJA1 As New iArray
  arrJA1.PushArray Array(1, 2, 3, "a", "b", "c")
  Dim arrJA2 As New iArray
  arrJA2.PushArray Array(4, 5, 6, "d", "e", "f")
  Dim arrJoined As iArray
  Set arrJoined = arrJA1.Join(arrJA2)
  Debug.Print arrJoined.ToString                 ' {1,2,3,"a","b","c",4,5,6,"d","e","f"}
  
  ' ##### DROP LEFT/RIGHT TEST
  Debug.Print vbCrLf & " #### Drop left/right test"
  Dim arrDrops As New iArray
  arrDrops.PushArray Array("1", "Two", "3", 4, 1, 2, 3, 4, 5, "a", "b", "c", True)
  Debug.Print arrDrops.DropLeft(2).ToString      ' {"1", "Two"}
  Debug.Print arrDrops.ToString                  ' {"3",4,1,2,3,4,5,"a","b","c",True}
  Debug.Print arrDrops.DropRight(3).ToString     ' {"b","c",True}
  Debug.Print arrDrops.ToString                  ' {"3",4,1,2,3,4,5,"a"}
  
  ' ##### REMOVE DUPLICATES TEST
  Debug.Print vbCrLf & " #### Remove duplicates test"
  Dim arrDupl As New iArray
  arrDupl.PushArray Array(1, 2, "a", 2, 3, 2, 3.14, "b", True, 4, "a")
  Debug.Print arrDupl.RemoveDuplicates           ' 3
  Debug.Print arrDupl.ToString                   ' {1,2,"a",3,3.14,"b",True,4}
  
  ' ##### CLONE ARRAY TEST
  Debug.Print vbCrLf & " #### Clone array test"
  Dim arrToClone As New iArray
  arrToClone.PushArray Array("3", 4, 1, 2, 3, 4, 5, "a", "b", "c", True)
  Dim arrCloned As New iArray
  Set arrCloned = arrToClone.Clone
  arrToClone.Clear
  Debug.Print arrCloned.ToString                 ' {"3",4,1,2,3,4,5,"a","b","c",True}
  
  ' ##### SHUFFLE ARRAY TEST
  Debug.Print vbCrLf & " #### Shuffle array test"
  Dim arrToShuffle As New iArray
  arrToShuffle.PushArray Array("3", 4, 1, 2, 3, 4, 5, "a", "b", "c", True)
  Dim arrShufl As New iArray
  Set arrShufl = arrToShuffle.Shuffle
  Debug.Print arrShufl.ToString                  ' {MIXED}
  
  ' ##### REVERSE ARRAY TEST
  Debug.Print vbCrLf & " #### Reverse array test"
  Dim arrToReverse As New iArray
  arrToReverse.PushArray Array("3", 4, 1, 2, 3, 4, 5, "a", "b", "c", True)
  Dim arrRev As New iArray
  Set arrRev = arrToReverse.Reverse
  Debug.Print arrRev.ToString                    ' {True,"c","b","a",5,4,3,2,1,4,"3"}
   
End Sub
